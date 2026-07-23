#!/usr/bin/env bash
set -euo pipefail

usage() {
    cat <<'EOF'
Validate a freshly generated deck with Microsoft's Open XML SDK.

Usage:
  scripts/validate-openxml.sh              Generate a representative deck, then validate it.
  scripts/validate-openxml.sh FILE.pptx    Validate an existing PPTX instead of generating one.

The command requires Docker, but does not require host Rust or .NET. Rust runs in
rust:1.97.1 with one CPU and one Cargo build job; the validator runs in
mcr.microsoft.com/dotnet/sdk:8.0 with DocumentFormat.OpenXml 3.3.0.
Validation errors include the package part, XML path, and SDK description.
EOF
}

if [[ "${1:-}" == "--help" || "${1:-}" == "-h" ]]; then
    usage
    exit 0
fi

if [[ "$#" -gt 1 ]]; then
    usage >&2
    exit 2
fi

script_dir="$(dirname "${BASH_SOURCE[0]}")"
repo_root="$(cd "$script_dir/.." && pwd)"
work_dir="$(mktemp -d "${TMPDIR:-/tmp}/church-powerpoint-openxml.XXXXXX")"
trap 'rm -rf "$work_dir"' EXIT

if [[ "$#" -eq 0 ]]; then
    input_path="$work_dir/generated.pptx"
    mkdir -p "$work_dir/target" "$work_dir/cargo-home"
    docker run --rm --cpus=1 \
        --mount "type=bind,source=$repo_root,target=/workspace,readonly" \
        --mount "type=bind,source=$work_dir,target=/output" \
        --user "$(id -u):$(id -g)" \
        -e CARGO_BUILD_JOBS=1 \
        -e CARGO_TARGET_DIR=/output/target \
        -e CARGO_HOME=/output/cargo-home \
        -e OPENXML_VALIDATOR_OUTPUT=/output/generated.pptx \
        -w /workspace \
        rust:1.97.1 \
        cargo test -p deck-builder --test build_deck builds_valid_pptx_from_service_record -- --exact --nocapture
    container_input="/output/generated.pptx"
else
    input_path="$1"
    if [[ ! -f "$input_path" ]]; then
        printf 'PPTX file does not exist: %s\n' "$input_path" >&2
        exit 2
    fi
    input_path="$(cd "$(dirname "$input_path")" && pwd)/$(basename "$input_path")"
    container_input="/input/$(basename "$input_path")"
fi

if [[ "$#" -eq 0 && ! -s "$input_path" ]]; then
    printf 'Deck generation did not produce a non-empty PPTX: %s\n' "$input_path" >&2
    exit 2
fi

if [[ "$#" -eq 1 ]]; then
    docker run --rm --cpus=1 \
        --mount "type=bind,source=$input_path,target=$container_input,readonly" \
        --mount "type=bind,source=$repo_root,target=/src,readonly" \
        -e "OPENXML_VALIDATOR_INPUT=$container_input" \
        mcr.microsoft.com/dotnet/sdk:8.0 \
        sh -eu -c 'cp -R /src/tools/openxml-validator /tmp/openxml-validator && dotnet run --project /tmp/openxml-validator/OpenXmlValidator.csproj -- "$OPENXML_VALIDATOR_INPUT"'
else
    docker run --rm --cpus=1 \
        --mount "type=bind,source=$work_dir,target=/output,readonly" \
        --mount "type=bind,source=$repo_root,target=/src,readonly" \
        -e OPENXML_VALIDATOR_INPUT=/output/generated.pptx \
        mcr.microsoft.com/dotnet/sdk:8.0 \
        sh -eu -c 'cp -R /src/tools/openxml-validator /tmp/openxml-validator && dotnet run --project /tmp/openxml-validator/OpenXmlValidator.csproj -- "$OPENXML_VALIDATOR_INPUT"'
fi
