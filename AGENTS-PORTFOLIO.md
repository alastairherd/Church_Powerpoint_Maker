# OpenCode subagent portfolio

This project configuration uses:

- OpenAI through ChatGPT Plus/Pro OAuth for GPT-5.6 Sol, Terra and Luna.
- DeepSeek's official API for DeepSeek V4 Flash.
- CrofAI's OpenAI-compatible API for GLM-5.2.

## Files

```text
opencode.json
.opencode/agents/
  orchestrator.md
  explore.md
  scout.md
  diagnose.md
  implement.md
  overflow-implement.md
  review.md
  critical-review.md
  architect.md
  hard-fix.md
  final-escalation.md
compat/
  final-escalation-xhigh.md
```

## Authentication

### OpenAI via ChatGPT Plus/Pro

Run OpenCode, enter `/connect`, choose **OpenAI**, then choose **ChatGPT Plus/Pro** and complete browser authentication.

### DeepSeek

Enter `/connect`, choose **DeepSeek**, and enter your official DeepSeek API key.

### CrofAI

Set the CrofAI key in your shell before launching OpenCode:

```bash
export CROF_API_KEY='nahcrof_...'
```

For persistence, add the export to your shell profile. The supplied custom provider points to `https://crof.ai/v1` and uses model ID `glm-5.2`.

## Installation

Copy `opencode.json` and the `.opencode` directory into the root of the project in which these agents should apply.

Then verify the models:

```bash
opencode models
```

Expected model identifiers:

```text
openai/gpt-5.6-sol
openai/gpt-5.6-terra
openai/gpt-5.6-luna
deepseek/deepseek-v4-flash
crofai/glm-5.2
```

The default primary agent is `orchestrator`. Subagent depth is set to `1`, so specialists cannot recursively launch further specialists.

## Routing summary

| Agent | Model | Purpose |
|---|---|---|
| orchestrator | GPT-5.6 Sol Low | Interpret, decompose, delegate and integrate |
| explore | DeepSeek V4 Flash non-thinking | Local repository discovery |
| scout | DeepSeek V4 Flash non-thinking | External docs and upstream research |
| diagnose | GLM-5.2 via CrofAI | Root-cause analysis without edits |
| implement | GPT-5.6 Luna High | Default bounded implementation |
| overflow-implement | GLM-5.2 via CrofAI | Quota-preserving or independent implementation |
| review | GLM-5.2 via CrofAI | Independent read-only review |
| architect | GPT-5.6 Terra xhigh | Cross-cutting design and planning |
| critical-review | GPT-5.6 Terra xhigh | High-risk audit |
| hard-fix | GPT-5.6 Sol High | Difficult implementation after failure |
| final-escalation | GPT-5.6 Sol Max | Manual last resort |

## Compatibility notes

### GPT-5.6 Max reasoning

GPT-5.6 supports `max`, but some OpenCode builds currently expose reasoning variants only through `xhigh`. The main `final-escalation.md` requests `max`. If your OpenCode build rejects it, replace that file with `compat/final-escalation-xhigh.md`.

### GLM-5.2 reasoning effort

CrofAI documents `glm-5.2` as a reasoning model but does not currently document a portable High/Max effort parameter. The GLM roles therefore use different prompts and step budgets rather than sending an unsupported effort field.

### DeepSeek V4 Flash

The exploration agents explicitly disable thinking mode. This keeps them fast and avoids spending reasoning tokens on bounded search tasks. The provider entry also declares `reasoning_content` interleaving for compatibility if thinking is enabled later.

### Repository-specific commands

The Bash allow-lists cover common Python, JavaScript, Rust, Go and .NET test commands. Remove commands your project does not use and add exact project commands where appropriate. Broad or destructive commands remain `ask` or `deny`.
