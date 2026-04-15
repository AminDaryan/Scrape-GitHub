"""Configuration constants for the usage-examples checker."""

# Named documentation files we always want to fetch (compared in lowercase).
TARGET_FILENAMES = {
    "readme.md", "readme.rst", "readme.txt", "readme",
    "usage.md", "usage.rst", "usage.txt",
    "quickstart.md", "quick_start.md", "quick-start.md",
    "tutorial.md", "tutorials.md",
    "examples.md", "example.md",
    "demo.md",
    "getting_started.md", "getting-started.md",
    "docs/usage.md", "docs/quickstart.md", "docs/tutorial.md",
    "docs/examples.md", "docs/demo.md", "docs/getting_started.md",
    "pyproject.toml", "setup.cfg",
}

# File extensions that are always usage examples by definition — we fetch ALL
# of them regardless of where they live in the repo (not capped like before).
# NOTE: .rmd and .qmd are excluded — they are analysis documents, not tutorials.
USAGE_EXTENSIONS = {".ipynb"}

# Extensions of plain example/demo scripts we also want to collect
SCRIPT_EXTENSIONS = {".py", ".r", ".sh", ".bash", ".m", ".jl"}

# Folder names whose entire contents are example/usage material
EXAMPLE_FOLDER_PREFIXES = ("examples/", "tutorials/", "demo/", "demos/",
                            "notebooks/", "notebook/", "scripts/", "sample/", "samples/")

# Maximum characters sent to the LLM per request
MAX_CONTENT_CHARS = 400_000

# Maximum number of notebooks to fully fetch (summarised) in one pass.
MAX_NOTEBOOKS = 20

# Maximum number of example scripts to fetch per pass
MAX_SCRIPTS = 15
