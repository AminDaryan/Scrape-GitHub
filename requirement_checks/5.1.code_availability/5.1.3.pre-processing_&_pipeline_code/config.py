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

# Maximum characters sent to LLM per repo (sum of per-bucket budgets below).
MAX_TOTAL_CHARS = 80_000

# Per-bucket character budgets. NAMED gets the largest slice so that high-signal
# files like preprocess.py are not crowded out by unrelated scripts/*.js files
# that happen to live in a preprocessing-folder prefix (cf. VULCA-EMNLP2025
# where scripts/ contained 70 codebase-maintenance scripts).
BUCKET_BUDGET_NAMED   = 50_000
BUCKET_BUDGET_FOLDER  = 20_000
BUCKET_BUDGET_GENERIC = 10_000

# Per-file character cap. NAMED files get a larger cap so long preprocessing
# modules can be sent in (near-)full.
PER_FILE_CHAR_CAP_NAMED = 16_000
PER_FILE_CHAR_CAP       = 8_000

# Maximum number of files to fetch per bucket.
MAX_NAMED_FILES   = 12
MAX_FOLDER_FILES  = 12
MAX_GENERIC_FILES = 8
