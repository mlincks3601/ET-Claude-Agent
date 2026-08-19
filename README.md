# AI Agent Workspace

This repo drives an AI agent (via Claude Code) that runs the company's
recurring research, reporting, and CRM/list-hygiene work. It's a portable
version of the same setup originally built as a local folder — the logic
hasn't changed, just where it lives.

## How it's organized

- **CLAUDE.md** — the agent's standing persona and rules (audience, output
  formats, when to ask clarifying questions, where files get saved). This is
  read automatically by Claude Code, both locally and via GitHub Actions.
- **workflows/** — plain-English task recipes. Start at
  `workflows/README.md` for the current list; tell the agent which one to
  run, or just describe the job and it'll match one.
- **resources/** — reference docs and reusable tools, including a tested
  Python script (`resources/HubSpot_Matching_Model/`) for reconciling
  marketing lists against a HubSpot export.
- **output/** — where finished deliverables land, organized into one
  subfolder per topic/run. Generated files are gitignored by default (see
  `output/.gitignore`) since they can contain company data — remove that if
  you want specific reports kept in git history.

## Running it locally

1. Clone this repo and open it in VS Code with the Claude Code extension
   installed.
2. `pip install -r requirements.txt`
3. Open a Claude Code session in this folder — it reads CLAUDE.md
   automatically. Tell it which workflow to run, or describe the task.

## Running it via GitHub

This repo is set up to also run through Claude Code's GitHub integration, so
teammates can trigger it without installing anything locally, and it can
eventually run on its own schedule. See the setup notes below or ask the
agent to walk you through `/install-github-app`.

## For whoever's maintaining this

Adding a new recurring job means adding one new file to `workflows/` (copy
the shape of an existing one) — no code changes needed unless the job needs
a new script, in which case that script goes in `resources/`.
