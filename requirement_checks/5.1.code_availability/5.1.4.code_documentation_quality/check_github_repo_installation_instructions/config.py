"""Configuration constants for the installation-instructions checker."""

# Files to fetch from each repo and send to the LLM
TARGET_FILENAMES = {
    "readme.md", "readme.rst", "readme.txt", "readme",
    "install.md", "install.rst", "install.txt",
    "setup.py", "setup.cfg", "pyproject.toml",
    "requirements.txt", "requirements-dev.txt",
    "environment.yml", "environment.yaml",
    "dockerfile", "docker-compose.yml", "docker-compose.yaml",
    "makefile",
    "getting_started.md", "getting-started.md",
    "docs/install.md", "docs/installation.md", "docs/setup.md",
}

# Max chars sent to LLM per repo
MAX_CONTENT_CHARS = 120_000
