---
name: internal-mathparser-protect-workspace
description: >-
  Internal MathParser workspace backup rules before editing source files and
  scripts (documentation needs no backup).
  Not a user entry point — load only when referenced by mathparser-add-or-fix
  or mathparser-refactor.
disable-model-invocation: true
---

# Internal: Protect Uncommitted Work Before Editing

Load **before the first edit** to any source file or script you might change.

## Check

Use `git status` for the **source files and scripts** you will touch: `MathParser*.bas`, `SmokeTest_MathParser.bas`, `cpp/MathParser*.cpp` / `.hpp`, `cpp/Tests_MathParser.cpp`, generated test data (`tools/*_generated*.bas`, `tools/*_generated*.cpp.inc`), other SmartMath/Basic/C++ sources, and scripts (`*.bat`, `tools/*.py`).

Documentation needs no backup: `USAGE_AND_SYNTAX.md`, `README.md`, `docs/`, `changelog.txt`, and `.cursor/` skills, flows, and rules.

An untracked file that existed before the task counts as dirty (git has no copy of it). Files the agent created during the task do not need a backup.

Unsaved editor changes are not on disk, so neither git nor a file copy captures them. If the open-files list shows a file you will touch as modified/unsaved, ask the user to save it before editing.

If those paths are clean, no backup is required.

## If dirty — backup first

Create a backup **before** any edit, revert, reset, `git checkout` / `git restore` / `git stash push` that could lose work:

1. **Required:** copy each affected file to `tools/backup/<filename>.<YYYYMMDD-HHMMSS>.bak` (create `tools/backup/` if needed). Use a filesystem copy — do not rely on git alone.
2. **Optional extra:** `git stash store -m "agent-backup-before-<task-summary>" $(git stash create)` records tracked changes without touching the working tree; only in addition to the file copy, and only when tracked files have changes (otherwise `git stash create` prints nothing and `store` fails). Never use `git stash push` for a backup: it reverts the files it stashes and skips untracked files.
3. In the first progress note to the user, list the **exact** backup path(s).
4. **Forbidden without explicit user request:** `git checkout HEAD -- <file>`, `git restore --source=HEAD <file>`, or any overwrite of the user’s uncommitted version. Recover from the timestamped copy or ask the user.
5. If you already overwrote a file by mistake: stop editing, say so immediately, help recover from backup / IDE local history / `git stash list` — do not repeat destructive git commands.

Applies to feature work, fixes, refactors, reorder passes, and test-only edits that touch source files or scripts. To inspect `HEAD` without touching the working tree, use `git show HEAD:<path>` or a temporary `git worktree` outside the repo (remove it with `git worktree remove` when done).
