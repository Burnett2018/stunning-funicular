# stunning-funicular

Project instructions for Claude Code. Claude loads this file automatically at
the start of every session, local or cloud.

## Status

New repository with no source code yet. Update this file as the project takes
shape: fill in the sections below once a language, framework and tooling are
chosen.

## Project overview

<!-- What this project does and who it is for, in a few sentences. -->

## Commands

<!-- Replace with the real commands once they exist. -->
- Install: `TODO`
- Build: `TODO`
- Test: `TODO`
- Lint / format: `TODO`

## Layout

<!-- Key directories and what lives in each. -->

## Conventions

- Keep changes small and focused; one concern per commit.
- Write clear commit messages: a short imperative summary line, then details.
- Add or update tests alongside behavior changes.
- Don't commit secrets, credentials or `.env` files.

## Cloud sessions

- Work happens on `claude/*` branches; open a draft PR for review.
- Once there are dependencies, add a SessionStart hook in
  `.claude/settings.json` so cloud sessions install them and can run tests
  and linters.
