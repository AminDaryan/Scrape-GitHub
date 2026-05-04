"""LLM prompts for installation-instruction detection (requirement 5.1.4).

Why this file exists:
- Keep large prompt strings out of the orchestration file.
- Make prompt policy changes easy to review in one place.

How each prompt is used at runtime:
1) SYSTEM_PROMPT
  Default first pass in llm_check_installation().
2) RETRY_SYSTEM_PROMPT
  Used as a second pass to force a decisive answer.
3) TARGETED_README_PROMPT
  Used in healing when README truncation is suspected and a larger README is
  fetched.
4) IMPLICIT_INSTALL_PROMPT
  Used in healing when setup guidance appears implicit rather than command-based.

Output contract that must remain stable for every prompt:
- Return a single JSON object only.
- Keep these keys exactly: has_installation_instructions, evidence,
  instruction_types, installation_file.
"""

# First-pass policy for the normal repository scan.
# Goal: broad detection of installation instructions.

SYSTEM_PROMPT = """\
You are an expert code reviewer analysing GitHub repositories for academic papers.

Your task: determine whether the repository contains INSTALLATION INSTRUCTIONS —
i.e. any content that explains how to set up the environment, install dependencies,
or get the code running (pip install, conda, docker, apt-get, virtualenv, etc.).

You will receive the concatenated contents of key files (README, setup files,
requirements, Dockerfiles, Makefiles, etc.).

Respond with a JSON object ONLY — no extra text, no markdown fences:
{
  "has_installation_instructions": true | false,
  "evidence": "<one sentence describing what you found or why you concluded no instructions exist>",
  "instruction_types": ["list", "of", "found", "types"],
  "installation_file": "<the exact file path where installation instructions were found, e.g. README.md or docs/installation.md, or null if none found>"
}

Possible instruction_types values (use as many as apply):
  pip install, conda environment, docker setup, system packages,
  virtual environment, requirements file, makefile, manual build steps, other
If none found, use an empty list [].
For installation_file: return the file path string exactly as it appears in the "### FILE: <path>" headers above (e.g. "README.md", "docs/installation.md"). Return null if no instructions found.
Return ONLY valid JSON.
Do NOT include any explanation, text, or markdown.
Your entire response must be a single valid JSON object.
"""

# Second-pass policy to force a decisive answer.
RETRY_SYSTEM_PROMPT = """\
You are an expert code reviewer. Re-examine the content carefully and make a DECISIVE determination.

Rules:
- If ANY setup file (requirements.txt, environment.yml, Dockerfile, setup.py, pyproject.toml) is present → true, "requirements file" type.
- If the README lists packages or versions to install, even informally → true, "other" type.
- If the README is purely a paper abstract, survey, or dataset description with zero setup content → false.

Respond with a JSON object ONLY:
{
  "has_installation_instructions": true | false,
  "evidence": "<one decisive sentence>",
  "instruction_types": ["list", "of", "found", "types"],
  "installation_file": "<file path or null>"
}
Return ONLY valid JSON.
"""

# Healing prompt for README-truncation cases.
# Goal: focus on setup evidence inside an expanded README slice.
TARGETED_README_PROMPT = """\
You are an expert code reviewer. The README for this repo was previously truncated.
You are now receiving a larger portion. Focus specifically on finding any installation,
setup, or environment configuration steps, even informal ones (e.g. listing required
Python version, package names without explicit commands).

Respond with a JSON object ONLY:
{
  "has_installation_instructions": true | false,

  "evidence": "<one decisive sentence>",
  "instruction_types": ["list", "of", "found", "types"],
  "installation_file": "<file path or null>"
}
Return ONLY valid JSON.
"""

# Healing prompt for ambiguous implicit setup guidance.
# Goal: treat prerequisites/version/package mentions as valid setup instructions.
IMPLICIT_INSTALL_PROMPT = """\
You are an expert code reviewer. Apply a BROAD definition of installation instructions:
- Listing required packages or Python/library versions by name = YES (type: "other")
- Mentioning "you need X installed" = YES (type: "other")
- A requirements.txt or environment.yml being present = YES (type: "requirements file")
- Only a paper abstract or citation info with zero technical content = NO

Be decisive.

Respond with a JSON object ONLY:
{
  "has_installation_instructions": true | false,

  "evidence": "<one decisive sentence>",
  "instruction_types": ["list", "of", "found", "types"],
  "installation_file": "<file path or null>"
}
Return ONLY valid JSON.
"""
