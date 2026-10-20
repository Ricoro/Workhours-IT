# Repository workflow

- The user requires all development, fixes, testing, and related file edits on the `test` branch.
- Check `git branch --show-current` before making edits. If it is not `test`, preserve local changes and switch to the existing `test` branch before working. Do not silently create a different development branch.
- `main` is the stable release branch. Do not commit, merge, reset, rebase, or push changes to `main` unless the user explicitly instructs promotion to main in the current conversation.
- Push development updates only to `origin test`, with an explicit refspec. Never use `git push --all`, mirror pushes, force pushes, or an implicit push that could update main.
- Work in this existing isolated cloud checkout. Do not create additional worktrees unless the user requests them.
- Before saving development work, run `npm test`, `npm run build`, and production browser checks with `PLAYWRIGHT_SERVER_COMMAND='npm run preview -- --port 5174 --strictPort' npm run test:e2e`.
- Keep company records, workbooks, and JSON backups out of Git. The repository owner has confirmed it is private.
