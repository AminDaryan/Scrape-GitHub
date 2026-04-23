"""LLM prompts for usage-example detection (requirement 5.1.4).

Why this file exists:
- Separate prompt policy from execution logic.
- Keep all usage-example decision rules in one maintainable place.

How these prompts are used by the logic script:
1) SYSTEM_PROMPT
  Default first pass in llm_check_usage().
2) RETRY_SYSTEM_PROMPT
  Used as a second pass to enforce decisive recall.
3) TARGETED_README_PROMPT
  Used in healing when README truncation is suspected.
4) DEEP_SCAN_PROMPT
  Used in healing after fetching extra notebooks/scripts/example folders.

Output contract that must remain stable:
- Return one JSON object only.
- Keep keys and nested schema expected by parser/export logic:
  has_usage_examples, evidence, example_types, example_files.
"""

# First-pass policy over all fetched repository snippets.
# Goal: identify every practical usage example, not just install/setup steps.

SYSTEM_PROMPT = """\
You are an expert code reviewer analysing GitHub repositories for academic papers.

Your task: find ALL usage examples in the repository — content that shows HOW
TO USE the code or model, beyond just installing it.

Usage examples include:
  - Jupyter notebooks (.ipynb)
  - Example or demo scripts (example.py, demo.py, run_example.sh, …)
  - Code snippets in the README that demonstrate the API or CLI
  - A "Quick-start" or "Usage" section in the README containing runnable code
  - Command-line usage examples with flags/arguments shown
  - Tutorial files or a tutorials/ / examples/ / demo/ folder

Does NOT count as a usage example:
  - Installation instructions alone (pip install X, conda env create, conda activate)
  - Navigation commands alone (cd src, mkdir, ls)
  - Abstract or paper description
  - Citation / BibTeX blocks
  - API reference lists with no example calls
  - A command only counts if it actually RUNS the tool/model, not just sets up the environment.
    If cd or conda commands appear alongside a real run command, include only the run command.

Respond with a JSON object ONLY — no extra text, no markdown fences:
{
  "has_usage_examples": true | false,
  "evidence": "<one sentence summarising what you found overall>",
  "example_types": ["distinct", "types", "found"],
  "example_files": [
    {
      "path": "<exact file path as it appears in the ### FILE / ### NOTEBOOK / ### SCRIPT header>",
      "type": "<one of: notebook | example script | README code snippet | CLI usage | tutorial file | quickstart guide | other>",
      "description": "<one sentence: what this specific file demonstrates>",
      "commands": ["<first exact code block or command>", "<second code block if present>"]
    }
  ]
}

IMPORTANT rules for example_files:
  - List EVERY individual example file you found — do not stop at one.
  - For a README that contains multiple distinct usage sections, list it once
    with type "README code snippet".
  - For each notebook or script file, list it as its own entry.
  - If no examples exist, use an empty list [].
  - Paths must be copied EXACTLY from the file headers above (e.g. "README.md",
    "examples/demo.ipynb") — do not invent or guess paths.
  - "commands": list ALL distinct code blocks / commands shown for this file.
    Each entry is one self-contained code block exactly as written in the repo.
    Include BOTH CLI commands (e.g. python run.py ...) AND Python API blocks.
    Use an empty list [] if no commands are shown.
    Do NOT merge multiple blocks into one string.

Return ONLY valid JSON. No explanation, no markdown, no preamble.
"""

# Second-pass policy to enforce decisive recall of notebooks, scripts, and README run examples.
RETRY_SYSTEM_PROMPT = """\
You are an expert code reviewer. Re-examine the content and list ALL usage examples found.

Hard rules:
  - Every .ipynb file IS a usage example — list each one.
  - A README code block (``` fences) showing how to call the library → list it.
  - An examples/, tutorials/, or demo/ directory mentioned in the README → list
    the folder as one entry of type "demo folder".
  - A purely abstract / citation README with zero runnable content → empty list.

Respond with a JSON object ONLY:
{
  "has_usage_examples": true | false,

  "evidence": "<one decisive sentence>",
  "example_types": ["list", "of", "found", "types"],
  "example_files": [
    {"path": "<exact path>", "type": "<type>", "description": "<one sentence>", "commands": ["<code block 1>", "<code block 2 if any>"]}
  ]
}
Return ONLY valid JSON.
"""

# Healing prompt used when only README needs a deeper second look.
# Goal: capture usage evidence from previously truncated README content.
TARGETED_README_PROMPT = """\
You are an expert code reviewer. The README was previously truncated; you now
have a larger portion. List ALL usage examples — code blocks, CLI examples,
references to notebooks or example files.

Respond with a JSON object ONLY:
{
  "has_usage_examples": true | false,

  "evidence": "<one decisive sentence>",
  "example_types": ["list", "of", "found", "types"],
  "example_files": [
    {"path": "<exact path>", "type": "<type>", "description": "<one sentence>", "commands": ["<code block 1>", "<code block 2 if any>"]}
  ]
}
Return ONLY valid JSON.
"""

# Healing prompt used after collecting extra code artifacts beyond README.
# Goal: aggregate all usage examples found in notebooks/scripts/folders.
DEEP_SCAN_PROMPT = """\
You are an expert code reviewer. Additional files (notebooks, scripts, example
folders) have been fetched. List EVERY usage example you find in the content.

Respond with a JSON object ONLY:
{
  "has_usage_examples": true | false,

  "evidence": "<one decisive sentence>",
  "example_types": ["list", "of", "found", "types"],
  "example_files": [
    {"path": "<exact path>", "type": "<type>", "description": "<one sentence>", "commands": ["<code block 1>", "<code block 2 if any>"]}
  ]
}
Return ONLY valid JSON.
"""
