# syntax=docker/dockerfile:1
#
# Dependency compilation is cached separately from application code.
#
# The previous version was:
#   COPY . .
#   RUN cargo build --release -p server
# which makes one layer depend on every file in the repository, so editing a
# single line of Rust -- or the README -- recompiled all 301 crates in
# Cargo.lock. That was roughly 5 minutes of an 8 minute build, every time.
#
# cargo-chef splits it: `prepare` reduces the workspace to a recipe describing
# only its dependencies, and `cook` builds those. The cook layer is keyed on the
# recipe, i.e. on Cargo.toml and Cargo.lock, so it is reused for any change that
# does not alter dependencies. Change a dependency and it correctly rebuilds.

FROM rust:1.97.1 AS chef
# Pinned so this layer is stable; it is the slowest thing to rebuild if it ever
# does get invalidated.
RUN cargo install cargo-chef --locked --version 0.1.71
WORKDIR /app

FROM chef AS planner
COPY . .
RUN cargo chef prepare --recipe-path recipe.json

FROM chef AS builder
COPY --from=planner /app/recipe.json recipe.json
# Dependencies only. Deliberately not restricted with `-p server`: cooking the
# whole workspace builds a superset, avoids relying on cargo-chef's argument
# passthrough, and costs nothing because the result is cached either way.
RUN cargo chef cook --release --recipe-path recipe.json
# Only now does application source enter the image, so everything above this
# line survives an ordinary code change.
COPY . .
RUN cargo build --release -p server

FROM debian:stable-slim AS runtime
RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates \
    && rm -rf /var/lib/apt/lists/*
COPY --from=builder /app/target/release/server /usr/local/bin/church-deck-server
ENV PORT=8080
EXPOSE 8080
CMD ["/usr/local/bin/church-deck-server"]
