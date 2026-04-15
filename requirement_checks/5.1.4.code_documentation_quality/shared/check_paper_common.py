"""Shared check_paper skeleton for all 5.1.4 checker scripts.

Extracts the common URL-validation → GitHub-parsing → fetch → LLM → status
pattern so each checker only supplies its checker-specific config/callbacks.
"""

from common.fetch_and_parse_github_repo import parse_github_repo, is_github


def check_paper_generic(
    paper,
    *,
    extra_defaults,
    collect_content_fn,
    llm_check_fn,
    map_llm_result_fn,
    boolean_key,
    require_files=False,
    no_files_message="No relevant files found in the repository",
):
    """Run the standard validate → fetch → LLM → status pipeline for one paper.

    Parameters
    ----------
    paper : dict
        Must contain "title" and "repo".
    extra_defaults : dict
        Checker-specific result keys with their default values.  Merged into
        the base result dict (which already has title, repo, status, evidence,
        files_checked, note).
    collect_content_fn : callable(owner, repo, result) -> (content_str, files_list)
        Fetches repo content.  Receives the mutable *result* dict so it can
        stash extra metadata (e.g. ``result["all_example_meta"]``).
        Must return ``(content_str, files_checked_list)``.
    llm_check_fn : callable(content_str, paper_title) -> dict
        Sends content to the LLM and returns the parsed JSON response.
    map_llm_result_fn : callable(result, llm_result, owner, repo) -> None
        Copies checker-specific fields from *llm_result* into *result*.
        Must NOT set ``result["status"]`` — that is handled here.
    boolean_key : str
        Key inside the LLM result dict that holds the True/False answer
        (e.g. ``"has_inline_comments"``).
    require_files : bool
        If True, return an error when ``collect_content_fn`` yields no files.
    no_files_message : str
        Error message used when *require_files* triggers.
    """
    url = paper.get("repo", "")

    result = {
        "title":        paper["title"],
        "repo":         url,
        "status":       None,
        "evidence":     "",
        "files_checked": [],
        "note":         "",
        **extra_defaults,
    }

    if not is_github(url):
        result["status"] = "skipped"
        result["note"] = (
            f"Not a GitHub repo — manual review needed. URL: {url}"
            if url else "No repo URL provided"
        )
        return result

    owner, repo = parse_github_repo(url)
    if not owner:
        result["status"] = "error"
        result["note"] = "Could not parse GitHub URL"
        return result

    try:
        repo_content, files_checked = collect_content_fn(owner, repo, result)
        result["files_checked"] = files_checked

        if require_files and not files_checked:
            result["status"] = "error"
            result["note"] = no_files_message
            return result

        llm_result = llm_check_fn(repo_content, paper["title"])

        result["evidence"] = llm_result.get("evidence", "")
        map_llm_result_fn(result, llm_result, owner, repo)

        if llm_result.get(boolean_key) is None:
            result["status"] = "error"
            result["note"] = "Empty LLM response — content may be too large for this model"
        else:
            result["status"] = "yes" if llm_result[boolean_key] else "no"

    except RuntimeError as e:
        result["status"] = "error"
        result["note"] = str(e)
        return result
    except Exception as e:
        result["status"] = "error"
        result["note"] = f"Unexpected error: {e}"
        return result

    return result
