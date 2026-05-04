"""Configuration constants for the API documentation checker."""

# Source code file extensions to scan for docstrings.
SOURCE_CODE_EXTENSIONS = {
    ".py", ".r", ".jl",                                # scripting / data-science
    ".cpp", ".c", ".h", ".hpp", ".cc", ".cxx",         # C / C++
    ".java", ".scala", ".kt",                          # JVM
    ".js", ".ts", ".jsx", ".tsx",                      # JavaScript / TypeScript
    ".go", ".rs", ".swift",                            # systems
    ".cs",                                             # C#
    ".m",                                              # MATLAB / Objective-C
    ".rb", ".php",                                     # scripting
    ".cu", ".cuh",                                     # CUDA
}

# Folder prefixes to skip when scanning for source code.
SKIP_FOLDER_PREFIXES = (
    "test/", "tests/", "testing/", "__pycache__/",
    ".github/", ".circleci/", ".gitlab/",
    "node_modules/", "vendor/", "third_party/", "thirdparty/",
    "build/", "dist/", "egg-info/",
)

# File basenames to skip.
SKIP_BASENAMES = {
    "setup.py", "setup.cfg", "conftest.py",
    "__init__.py", "manage.py",
}

# Named documentation / API-spec files we always want to fetch (lowercased).
TARGET_DOC_FILENAMES = {
    # Root-level readme
    "readme.md", "readme.rst", "readme.txt", "readme",
    # Sphinx
    "conf.py", "index.rst",
    "docs/conf.py", "docs/index.rst", "docs/index.md",
    # MkDocs
    "mkdocs.yml", "mkdocs.yaml",
    # Doxygen
    "doxyfile", "doxyfile.in",
    # ReadTheDocs
    ".readthedocs.yml", ".readthedocs.yaml",
    "readthedocs.yml", "readthedocs.yaml",
    # OpenAPI / Swagger
    "openapi.yaml", "openapi.yml", "openapi.json",
    "swagger.yaml", "swagger.yml", "swagger.json",
    "api.yaml", "api.yml", "api.json",
    # JSDoc / TypeDoc
    "jsdoc.json", ".jsdoc.json", "typedoc.json",
    # pyproject with doc tool config
    "pyproject.toml",
}

# Folder prefixes whose markdown/RST/HTML we treat as API documentation pages.
# More specific api/reference prefixes are checked first.
DOC_FOLDER_PREFIXES = (
    "docs/api/", "docs/reference/", "docs/api-reference/",
    "apidoc/", "api_docs/", "api-docs/",
    "javadoc/",
    "docs/", "doc/", "documentation/",
)

# Extensions of documentation page files we want to collect from doc folders.
DOC_PAGE_EXTENSIONS = {".md", ".rst", ".txt"}

# Maximum number of source code files to scan for docstrings.
MAX_SOURCE_FILES = 20

# Maximum number of documentation page files to fetch.
MAX_DOC_FILES = 15

# Maximum characters sent to the LLM per request.
MAX_CONTENT_CHARS = 300_000

# Per-file character cap for source code files.
SOURCE_FILE_CHAR_CAP = 10_000

# Per-file character cap for documentation page files.
DOC_FILE_CHAR_CAP = 15_000
