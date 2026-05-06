# Scrape-GitHub

This repository analyzes paper repositories and documentation quality using GitHub API + LLM-based checks. Some checks have rule-based (no-LLM) companions for fast, deterministic evaluation.

Two requirement areas under `requirement_checks/`, each with several questions:

- **`5.1.code_availability/`** — Is the code available and well-documented?
  - **5.1.3** — does the repo contain real preprocessing / pipeline code?
  - **5.1.4** — documentation quality: inline comments, installation instructions, usage / example commands, API documentation (LLM + rule-based variants)
  - **5.1.5** — is the code released under an explicit license?
- **`5.2.practitioner_usability_and_popularity/`** — Is the code maintained and adopted?
  - **5.2.2** — ongoing-maintenance indicators (recent commits, contributors, versioned releases, staleness, archived status, …)
  - **5.2.3** — adoption metrics: GitHub stars, forks, PyPI monthly downloads
  - **5.2.4** — post-publication maintenance: date of last commit + total commit count

## Project Structure

At a glance:

- `requirement_checks/common/`
  Shared helpers used by every script: GitHub API, LLM calls + token tracking, Excel output, pipeline runner.
- `requirement_checks/data/`
  Centralized paper list (`papers_from_database.py`) consumed by every checker.
- `requirement_checks/5.1.code_availability/5.1.3.pre-processing_&_pipeline_code/`
  Q 5.1.3 — preprocessing / pipeline code classifier (LLM).
- `requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/`
  Q 5.1.4 — documentation quality (inline comments, installation, usage examples, API docs).
- `requirement_checks/5.1.code_availability/5.1.5.code_license/`
  Q 5.1.5 — automatic license extraction (no LLM).
- `requirement_checks/5.2.practitioner_usability_and_popularity/5.2.2.maintenance_activity_indicators/`
  Q 5.2.2 — maintenance indicators (no LLM).
- `requirement_checks/5.2.practitioner_usability_and_popularity/5.2.3.adoption_metrics/`
  Q 5.2.3 — adoption metrics: stars, forks, PyPI monthly downloads (no LLM).
- `requirement_checks/5.2.practitioner_usability_and_popularity/5.2.4.post_publication_maintenance/`
  Q 5.2.4 — last commit date + total commit count (no LLM).
- `requirement_checks/openai_client.py`
  LiteLLM-backed client supporting multiple providers (Azure OpenAI, OpenAI, Anthropic, Mistral, Ollama, Gemini).

Project map (Mermaid):

```mermaid
flowchart TD
  root["Repo Root"]
  root --> common["requirement_checks/common<br/>github_helpers, llm_helpers,<br/>checker_pipeline, excel_output"]
  root --> data["requirement_checks/data<br/>papers_from_database.py"]
  root --> q51["requirement_checks/5.1.code_availability"]
  q51 --> q513["5.1.3 preprocessing pipeline"]
  q51 --> q514["5.1.4 documentation quality"]
  q51 --> q515["5.1.5 code license"]
  q514 --> apidoc["api documentation<br/>(LLM + rule-based)"]
  q514 --> inline["inline comments"]
  q514 --> install["installation instructions"]
  q514 --> usage["usage examples"]
  root --> q52["requirement_checks/5.2.practitioner_usability_and_popularity"]
  q52 --> q522["5.2.2 maintenance indicators"]
  q52 --> q523["5.2.3 adoption metrics"]
  q52 --> q524["5.2.4 post-publication maintenance"]
  root --> client["openai_client.py"]
```

If your editor shows Mermaid as code, open Markdown Preview (`Ctrl+Shift+V`) or view the file on GitHub.

## Main Scripts And Outputs

All scripts read papers from the central [`data/papers_from_database.py`](requirement_checks/data/papers_from_database.py).

| Question | Script | Approach | Output |
|---|---|---|---|
| Q 5.1.3: Does the repo contain preprocessing/pipeline code? | [check_paper_appendix_for_data_preprocessing_code.py](requirement_checks/5.1.code_availability/5.1.3.pre-processing_&_pipeline_code/check_paper_appendix_for_data_preprocessing_code.py) | LLM | [results/preprocessing_code_results.xlsx](requirement_checks/5.1.code_availability/5.1.3.pre-processing_&_pipeline_code/results/preprocessing_code_results.xlsx) |
| Q 5.1.4: Inline comments? | [check_github_repo_inline_comments.py](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_inline_comments/check_github_repo_inline_comments.py) | LLM | [results/inline_comments_results.xlsx](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_inline_comments/results/inline_comments_results.xlsx) |
| Q 5.1.4: Installation instructions? | [check_github_repo_installation_instructions.py](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_installation_instructions/check_github_repo_installation_instructions.py) | LLM | [results/installation_instructions_results.xlsx](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_installation_instructions/results/installation_instructions_results.xlsx) |
| Q 5.1.4: Usage / example commands? | [check_github_repo_example_commands.py](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_usage_examples/check_github_repo_example_commands.py) | LLM | [results/usage_examples_results.xlsx](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_usage_examples/results/usage_examples_results.xlsx) |
| Q 5.1.4: API documentation? | [check_github_repo_api_documentation.py](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_api_documentation/check_github_repo_api_documentation.py) | LLM | [results/api_documentation_results.xlsx](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_api_documentation/results/api_documentation_results.xlsx) |
| Q 5.1.4: API documentation (no LLM)? | [api_documentation_check_no_llm_used.py](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_api_documentation/api_documentation_check_no_llm_used.py) | Rule-based (regex + Python AST) | [results/api_documentation_no_llm.xlsx](requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_api_documentation/results/api_documentation_no_llm.xlsx) |
| Q 5.1.5: Is the code released under an explicit license? | [5.1.5.code_license.py](requirement_checks/5.1.code_availability/5.1.5.code_license/5.1.5.code_license.py) | GitHub licensee API + LICENSE-file scan (no LLM) | [code_license.xlsx](requirement_checks/5.1.code_availability/5.1.5.code_license/code_license.xlsx) |
| Q 5.2.2: Ongoing-maintenance indicators? | [5.2.2.maintenance_activity_indicators.py](requirement_checks/5.2.practitioner_usability_and_popularity/5.2.2.maintenance_activity_indicators/5.2.2.maintenance_activity_indicators.py) | GitHub API only (no LLM) | [maintenance_indicators.xlsx](requirement_checks/5.2.practitioner_usability_and_popularity/5.2.2.maintenance_activity_indicators/maintenance_indicators.xlsx) |
| Q 5.2.3: Adoption metrics (stars, forks, PyPI monthly downloads)? | [5.2.3.adoption_metrics.py](requirement_checks/5.2.practitioner_usability_and_popularity/5.2.3.adoption_metrics/5.2.3.adoption_metrics.py) | GitHub API + pypistats.org (no LLM) | [adoption_metrics.xlsx](requirement_checks/5.2.practitioner_usability_and_popularity/5.2.3.adoption_metrics/adoption_metrics.xlsx) |
| Q 5.2.4: Post-publication maintenance (last commit date + total commits)? | [5.2.4.post_publication_maintenance.py](requirement_checks/5.2.practitioner_usability_and_popularity/5.2.4.post_publication_maintenance/5.2.4.post_publication_maintenance.py) | GitHub API only (no LLM) | [post_publication_maintenance.xlsx](requirement_checks/5.2.practitioner_usability_and_popularity/5.2.4.post_publication_maintenance/post_publication_maintenance.xlsx) |

### Multi-model runs

For LLM-based checkers, set `AZURE_OPENAI_DEPLOYMENT` to a JSON array (e.g. `["gpt-5", "gpt-5-mini"]`) to run every model in one pass. Each model produces its own `Results` and `Summary` sheet, plus a `Model Comparison` sheet showing per-paper agreement.

### Single-repo CLI for the no-LLM API doc checker

The rule-based API doc checker also works as a single-repo CLI:

```bash
python "requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_api_documentation/api_documentation_check_no_llm_used.py" pallets/flask
```

Without arguments it runs batch mode over the whole paper list.

## Shared Utilities

Files under `requirement_checks/common/`:

- [`github_helpers.py`](requirement_checks/common/github_helpers.py) — GitHub URL parsing, repo file listing, file-content fetching, pagination-aware GET, plus higher-level helpers for assembling LLM-friendly content with character budgets and README prioritisation.
- [`llm_helpers.py`](requirement_checks/common/llm_helpers.py) — `llm_call_parse_retry`, JSON response parsing, and `TokenUsageTracker` for per-model token accounting.
- [`checker_pipeline.py`](requirement_checks/common/checker_pipeline.py) — the `run_pipeline` orchestrator that loops papers across one or more LLM deployments and writes per-model + comparison sheets.
- [`excel_output.py`](requirement_checks/common/excel_output.py) — Excel writing helpers: borders, header rows, status-coloured data rows, summary sheets, model-comparison sheets.

## Environment Variables

Copy [.env.example](.env.example) to `.env` and fill in your keys.

Required for the LLM checkers (Azure OpenAI by default):

- `OPENAI_KEY`
- `AZURE_OPENAI_ENDPOINT`
- `AZURE_OPENAI_API_VERSION`
- `AZURE_OPENAI_DEPLOYMENT` — single name (e.g. `gpt-4o`) or JSON array (e.g. `["gpt-5","gpt-5-mini"]`)

Recommended:

- `GITHUB_TOKEN` — lifts the GitHub API rate limit from 60 → 5,000 req/hour. Without it, large runs WILL hit the limit.

Optional (provider switching via LiteLLM):

- `LLM_PROVIDER` — one of `azure` (default), `openai`, `anthropic`, `mistral`, `ollama`, `gemini`.
- `OPENAI_API_KEY`, `OPENAI_MODEL` — fallback for the OpenAI provider.

## First-Time Setup

Setup files:

- [requirements.txt](requirements.txt)
- [.env.example](.env.example)
- [setup.ps1](setup.ps1)
- [setup.sh](setup.sh)

Use one command to prepare a local environment and install dependencies:

- Windows PowerShell:

```powershell
./setup.ps1
```

- macOS/Linux:

```bash
bash ./setup.sh
```

This setup will:

- create `.venv`
- install dependencies from [requirements.txt](requirements.txt)
- create `.env` from [.env.example](.env.example) if missing

## Quick Run Commands (from repo root)

LLM-based checkers (run against every paper in `data/papers_from_database.py`, every model in `AZURE_OPENAI_DEPLOYMENT`):

```bash
python "requirement_checks/5.1.code_availability/5.1.3.pre-processing_&_pipeline_code/check_paper_appendix_for_data_preprocessing_code.py"
python "requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_inline_comments/check_github_repo_inline_comments.py"
python "requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_installation_instructions/check_github_repo_installation_instructions.py"
python "requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_usage_examples/check_github_repo_example_commands.py"
python "requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_api_documentation/check_github_repo_api_documentation.py"
```

Rule-based / no-LLM checkers (no API keys needed beyond `GITHUB_TOKEN`):

```bash
# API documentation — batch over all papers
python "requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_api_documentation/api_documentation_check_no_llm_used.py"

# API documentation — single repo, prints JSON
python "requirement_checks/5.1.code_availability/5.1.4.code_documentation_quality/check_github_repo_api_documentation/api_documentation_check_no_llm_used.py" pallets/flask

# License (Q 5.1.5)
python "requirement_checks/5.1.code_availability/5.1.5.code_license/5.1.5.code_license.py"

# Maintenance indicators
python "requirement_checks/5.2.practitioner_usability_and_popularity/5.2.2.maintenance_activity_indicators/5.2.2.maintenance_activity_indicators.py"

# Adoption metrics (GitHub stars/forks + PyPI monthly downloads)
python "requirement_checks/5.2.practitioner_usability_and_popularity/5.2.3.adoption_metrics/5.2.3.adoption_metrics.py"

# Post-publication maintenance (last commit date + total commits)
python "requirement_checks/5.2.practitioner_usability_and_popularity/5.2.4.post_publication_maintenance/5.2.4.post_publication_maintenance.py"
```
