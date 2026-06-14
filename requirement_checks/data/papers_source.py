"""Single source of truth for the paper list every 5.1/5.2 checker iterates over.

By default the list comes from ``data/papers_from_database.py`` (the vendored,
git-ignored corpus).  When the environment variable ``PAPERS_JSON`` points at a
JSON file, that file is used instead.

This is the hook the Streamlit UI uses to feed an *uploaded* papers list (or a
single paper) into the existing checker scripts without modifying them: the UI
writes the upload to a temp file, sets ``PAPERS_JSON``, and runs the checker as
a subprocess.  Because every checker now calls :func:`load_papers`, they all
pick up the override automatically.

The JSON may be a list of paper objects or a single paper object; a single
object is wrapped in a one-element list.  Each object uses the same schema as
``papers_from_database.PAPERS``::

    {"title": ..., "repo": "https://github.com/owner/repo", "semanticscholarid": ...}
"""

import json
import os


def load_papers():
    """Return the papers to check: the ``PAPERS_JSON`` override, or the vendored list."""
    path = os.environ.get("PAPERS_JSON", "").strip()
    if path:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list):
            raise ValueError(
                f"PAPERS_JSON ({path}) must contain a JSON list of papers or a "
                f"single paper object, got {type(data).__name__}."
            )
        return data

    # Fall back to the hand-curated vendored list (may be absent on a fresh
    # clone — same behaviour as before this hook existed).
    from data.papers_from_database import PAPERS
    return PAPERS
