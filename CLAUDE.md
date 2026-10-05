# Consulting workspace

## Context management
- A project skill named `compact` lives at `.claude/skills/compact/SKILL.md`.
- When the conversation gets long, or before switching workstreams, invoke it to save a
  structured checkpoint to `.claude/context/checkpoint.md`.
- At the start of a session, if `.claude/context/checkpoint.md` exists, read it and briefly
  restate the objective and next steps before doing new work.
