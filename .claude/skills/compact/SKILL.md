---
name: compact
description: Compact the working context into a structured checkpoint so a long session can continue (or restart after /clear) without losing the thread. Use when the user says "compact", "checkpoint", "summarize where we are", "context is getting long", or before switching workstreams. Writes the checkpoint to .claude/context/checkpoint.md and reloads from it on request.
---

# Compact — context checkpoint

Goal: shrink everything that matters in this conversation into a short, decision-ready
checkpoint, so the next turn (or a fresh session) can pick up exactly where we left off.

This skill does not delete the conversation itself. Claude Code's built-in `/compact`
command does that. Use this skill to write a durable checkpoint first, and optionally run
`/compact` or `/clear` afterwards.

## When invoked with no argument (or "save")

1. Review the full conversation and identify, consultant-style (MECE, answer-first):
   - **Objective**: the client or personal question being answered, in one sentence.
   - **Governing thought**: the current best answer or hypothesis.
   - **Issue tree status**: each branch → done / in progress / open.
   - **Key facts & numbers**: only figures we will reuse, each with its source URL or
     file path. Drop anything unsourced or superseded.
   - **Decisions made**: what was decided and why (one line each).
   - **Assumptions to challenge**: open assumptions that still need validation.
   - **Files touched**: paths created or edited, with a 3–8 word purpose each.
   - **Next steps**: ordered, concrete actions, with who acts (user or Claude).
   - **Open questions for the user**: anything blocking.
2. Drop: dead ends (keep a one-line "tried X, rejected because Y"), raw tool output,
   repeated drafts, pleasantries.
3. Write the result to `.claude/context/checkpoint.md` using the template below,
   overwriting the previous checkpoint. Keep it under ~400 lines; aim for under 150.
4. Reply to the user with the 5-line summary (objective, governing thought, top 3
   next steps) and tell them they can now run `/compact` or `/clear` safely, then
   say "load checkpoint" (or `/compact load`) to resume.

## When invoked with "load" (or "resume")

1. Read `.claude/context/checkpoint.md`. If missing, say so and stop.
2. Restate the objective, governing thought and next steps in under 10 lines.
3. Ask whether to continue with next step #1 or redirect.

## Checkpoint template

```markdown
# Checkpoint — <topic>
_Updated: <YYYY-MM-DD>_

## Objective
<one sentence>

## Governing thought
<current answer / hypothesis>

## Issue tree status
- <Branch 1> — done | in progress | open: <one line>
- <Branch 2> — ...

## Key facts & numbers
| Fact | Value | Source |
|---|---|---|

## Decisions made
- <decision> — <why>

## Assumptions to challenge
- <assumption> — <how to test>

## Rejected paths
- Tried <X>; rejected because <Y>

## Files touched
- `<path>` — <purpose>

## Next steps
1. <action> — <owner>

## Open questions for the user
- <question>
```

## Rules

- Never invent facts or sources to fill a section; write "none yet" instead.
- Keep source URLs exactly as found in the conversation.
- Prefer numbers with units and dates (e.g. "EV/EBITDA 11.2x, LTM Jun-26").
