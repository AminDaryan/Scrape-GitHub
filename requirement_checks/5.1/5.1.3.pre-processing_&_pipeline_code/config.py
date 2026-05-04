"""Configuration constants for the preprocessing / pipeline code checker."""

import re

# A file's basename matching this pattern is treated as preprocessing-relevant.
PREPROC_NAME_RE = re.compile(
    r"(preprocess|pre[_-]?proc|tokeniz|"
    r"clean(?:up|ing|er)?|filter|dedup(?:lic)?|"
    r"dataset|data[_-]?loader|dataloader|"
    r"prepare[_-]?data|prep[_-]?data|build[_-]?data|process[_-]?data|make[_-]?data|"
    r"prompt[_-]?(?:build|wrap|template)|pipeline|wrapper)",
    re.IGNORECASE,
)

# Any file under one of these top-level folders is treated as preprocessing-relevant.
PREPROC_FOLDER_PREFIXES = (
    "data/", "data_processing/", "data-processing/",
    "preprocessing/", "preprocess/",
    "pipeline/", "pipelines/",
    "tokenizer/", "tokenization/",
    "scripts/",
    "src/data/", "src/preprocessing/", "src/pipeline/",
)

# Extensions of source code files we want to analyse.
SOURCE_EXTS = (
    ".py", ".sh", ".ipynb", ".js", ".ts",
    ".rb", ".go", ".jl", ".r", ".scala", ".rs",
)

# Folder prefixes to skip (test / CI / vendored / build artefacts).
SKIP_FOLDER_PREFIXES = (
    "test/", "tests/", "__tests__/", ".github/",
    "node_modules/", "vendor/", "third_party/", "third-party/",
    "examples/", "docs/", "doc/",
    "build/", "dist/", ".cache/",
)

# Maximum characters sent to LLM per repo.
MAX_TOTAL_CHARS = 80_000

# Per-file character cap so that one huge file doesn't eat the whole budget.
PER_FILE_CHAR_CAP = 8_000

# Maximum number of files to fetch per bucket.
MAX_NAMED_FILES = 12
MAX_FOLDER_FILES = 12
MAX_GENERIC_FILES = 8
