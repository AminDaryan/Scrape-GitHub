"""Configuration constants for the inline-comments checker."""

# Extensions of source code files we want to analyse for inline comments.
SOURCE_CODE_EXTENSIONS = {
    ".py", ".r", ".rmd", ".jl", ".ipynb",           # scripting / data-science
    ".cpp", ".c", ".h", ".hpp", ".cc", ".cxx",     # C / C++
    ".java", ".scala", ".kt",                       # JVM
    ".js", ".ts", ".jsx", ".tsx",                   # JavaScript / TypeScript
    ".go", ".rs", ".swift",                         # systems
    ".m", ".matlab",                                # MATLAB
    ".sh", ".bash",                                 # shell
    ".lua", ".rb", ".php", ".pl",                   # other
    ".cu", ".cuh",                                  # CUDA
}

# Folder prefixes to skip (test / CI / vendored / build artefacts).
SKIP_FOLDER_PREFIXES = (
    "test/", "tests/", "testing/", "__pycache__/",
    ".github/", ".circleci/", ".gitlab/",
    "node_modules/", "vendor/", "third_party/", "thirdparty/",
    "build/", "dist/", "egg-info/",
)

# File basenames to skip (config, not real source code).
SKIP_BASENAMES = {
    "setup.py", "setup.cfg", "conftest.py",
    "__init__.py", "manage.py",
}

# Maximum number of source files to fetch
MAX_SOURCE_FILES = 25

# Maximum characters sent to LLM per repo
MAX_CONTENT_CHARS = 200_000

# Per-file character cap so that one huge file doesn't eat the whole budget.
PER_FILE_CHAR_CAP = 10_000
