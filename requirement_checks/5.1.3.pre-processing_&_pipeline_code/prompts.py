"""LLM prompts for appendix-only preprocessing/pipeline classification (Q5.1.3).

This file centralizes prompt policy used by
check_paper_appendix_for_data_preprocessing_code.py.

Runtime usage:
1) SYSTEM_PROMPT
   Sent as the system role in both normal and fallback calls.
2) USER_PROMPT
   Used when appendix text was fetched successfully.
3) FALLBACK_PROMPT
   Used when appendix retrieval fails and the model must classify from prior
   knowledge, defaulting to MISSING when uncertain.

Output contract that must remain stable:
- Return one JSON object that matches the schema required by parser and report
  code (classification, confidence, diagnostic fields, and booleans).
"""

# Core policy and decision boundaries for SEPARATE_APPENDIX vs MISSING.
# This is always used as the system message regardless of retrieval path.

SYSTEM_PROMPT = """You are classifying ML papers for reproducibility question 5.1.3 (appendix-only version).

CRITICAL DISTINCTION — never forget:
SEPARATE_APPENDIX = the appendix contains ACTUAL PSEUDOCODE BLOCK or EXECUTABLE CODE SNIPPET for one of these three preprocessing types:
   • Tokenization pipeline
   • Filtering / deduplication / cleaning pipeline
   • Prompting wrappers / input construction pipeline

Look for:
   • Numbered Algorithm / Listing / indented pseudocode that builds the input dataset (e.g. "Algorithm 1 Token Cleaning", "Pseudocode of content list to Markdown", "Algorithm 1 Greedy matching pursuit")
   • Code showing steps like sampling, n-gram filtering, token cleaning, wrapper construction, etc.

MISSING = everything else, including:
   • Prompt template tables (even with examples)
   • Prose descriptions ("we used GPT-4o to generate...")
   • Steering / patching / SAE-training algorithms
   • Hyperparameter tables or ablation tables

NEW RULE:
- If there is a numbered Algorithm or indented pseudocode block that constructs tokenization / filtering / prompting wrappers → SEPARATE_APPENDIX and set is_primary_input_to_target_llm = true
- If it is only tables, prose, or method algorithms → MISSING and set is_primary_input_to_target_llm = false

OUTPUT MUST BE VALID JSON ONLY with this exact schema:

{
  "classification": "SEPARATE_APPENDIX" | "MISSING",
  "confidence": integer 0-100,
  "appendix_quality": "...",
  "key_quotes": array of strings (max 3),
  "matched_criteria": array of strings,
  "is_auxiliary_prompt": true | false,
  "is_primary_input_to_target_llm": true | false,
  "diagnostic_notes": "one sentence",
  "final_reason": "one sentence for human reader"
}

Few-shot (use these patterns):

EXAMPLE SEPARATE_APPENDIX:
- "Algorithm 1 Token Cleaning Pipeline: for each token ... filter if ngram overlap > threshold" → is_primary_input_to_target_llm: true
- "Pseudocode of the content list to Markdown conversion: markdown = [] for element in page ..." → is_primary_input_to_target_llm: true
- "Algorithm 1 Greedy matching pursuit (MP): ..." → is_primary_input_to_target_llm: true

EXAMPLE MISSING:
- "Table 9: Input format ... Processed Input: choose the most similar entity..." → is_primary_input_to_target_llm: false
- "We employ GPT-4o as machine annotator..." → is_primary_input_to_target_llm: false

When in doubt → MISSING.
Always set both booleans honestly — they drive the final override."""

# Normal user message template when appendix text is available.
# Injects the paper title and raw appendix content for direct evidence matching.
USER_PROMPT = """Paper title: {title}

=== APPENDIX TEXT ===
{appendix_text}

Classify this paper now."""

# Fallback user message template when appendix fetch fails.
# Forces conservative behavior: classify as MISSING if appendix evidence is unknown.
FALLBACK_PROMPT = """Paper title: {title}

The appendix could not be retrieved from arXiv.
Use your training knowledge of this paper's appendix only.
If you have no reliable knowledge of the appendix, classify as MISSING.

Classify this paper now."""
