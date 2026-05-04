"""Internal helpers for the no-LLM API-documentation checker.

This package collects the pieces that don't need to live next to the
orchestrator in ``api_documentation_check_no_llm_used.py``:

  * :mod:`.patterns`         — regex patterns and priority lists.
  * :mod:`.source_analysis`  — language-specific docstring/comment
    detectors (Python via AST, others via regex).

The entry point in the parent folder imports from these modules.  Nothing
here depends on any LLM SDK or on a network library other than what is
needed to parse already-fetched file content.
"""
