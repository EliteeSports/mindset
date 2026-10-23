# Repository Guidelines

## Project Structure & Module Organization
This repository is a small Python service centered on `main.py`, which contains the HTTP-triggered Cloud Function entry point and the full Mindset assessment workflow. Deployment and runtime helpers live at the top level:

- `gcp_build.sh` builds the local Docker image, deploys the Gen 2 Cloud Function, routes traffic to the latest revision, and rebuilds the Cloud Scheduler job.
- `local_build_run.sh` builds and runs the Docker image locally.
- `requirements.txt` lists Python dependencies.

There is no dedicated `src/` or `tests/` directory yet; keep new modules small and place them at the repo root unless you are doing a larger refactor.

## Build, Test, and Development Commands
- `python3.11 -m venv .venv && source .venv/bin/activate` creates the local environment.
- `pip install -r requirements.txt` installs runtime dependencies.
- `python main.py` runs the workflow directly against live services.
- `./local_build_run.sh` builds the Docker image and runs it with the repo mounted into `/app`.
- `./gcp_build.sh` deploys the function to `europe-west2`, shifts traffic to the latest revision, and recreates the 2-hour Cloud Scheduler trigger.

## Coding Style & Naming Conventions
Follow existing Python conventions in `main.py`: snake_case for functions and variables, UPPER_CASE for true constants, and concise docstrings on non-trivial functions. Keep imports grouped and prefer short, focused helper functions over adding more logic to the entry point. Shell scripts should stay POSIX-friendly and use clear variable names.

## Testing Guidelines
There is no automated test suite configured today. Validate changes with the smallest safe scope possible:

- run `python main.py` only when live GCP, Salesforce, and Microsoft Graph access is available
- use `./local_build_run.sh` for container verification
- document manual checks in the PR, especially BigQuery, GCS, SharePoint, Salesforce, and scheduler side effects

If you add tests, use `tests/test_<feature>.py` and start with isolated pure-Python helpers rather than live integration coverage.

## Commit & Pull Request Guidelines
Recent commits use short imperative subjects, sometimes with a prefix such as `docs:`. Keep commit messages concise and specific, for example `docs: fold scheduler setup into deploy script`.

PRs should include a short summary, operational impact, manual verification steps, and links to any relevant ticket. Include screenshots only when a generated report or external UI output materially changed.
