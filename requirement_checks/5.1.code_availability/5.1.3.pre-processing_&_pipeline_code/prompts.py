"""LLM prompts for repo-only preprocessing/pipeline classification (Q5.1.3).

This file centralizes prompt policy used by
check_paper_appendix_for_data_preprocessing_code.py.

Runtime usage:
1) SYSTEM_PROMPT — sent as the system role on every call.
2) USER_PROMPT — carries the paper title, repo URL, and the source files
   fetched from the repository.

Output contract:
- Return one JSON object matching the schema described in SYSTEM_PROMPT:
  classification, preprocessing_files (each with path/category/evidence),
  diagnostic_notes, final_reason.
"""

# Core policy and decision boundaries for EXISTS vs NOT_APPLICABLE.
SYSTEM_PROMPT = """You are classifying ML papers for reproducibility question 5.1.3 (repo-only version).

You will be shown a selection of source files from a paper's GitHub repository.
Your job is to decide whether the repository contains ACTUAL EXECUTABLE
PREPROCESSING / PIPELINE CODE in ANY modality — text, image, audio, video,
or multimodal. Match one of these three categories:

   • tokenization — splitting raw inputs into model-ready units. Examples:
     BPE / WordPiece / SentencePiece tokenization, image patching / sliding
     windows, audio framing, video clip extraction, mel-spectrogram extraction.

   • filtering — removing, deduplicating, normalising, or cleaning data.
     Examples: n-gram or language filtering, MinHash / SimHash dedup, image
     resizing / cropping / colour normalisation / augmentation, audio
     resampling, outlier removal, saliency-based selection.

   • prompting — assembling model inputs from raw data. Examples: chat
     templates, in-context example builders, multimodal prompt assembly,
     embedding caches, retrieval-augmented input construction.

CRITICAL DISTINCTION — never forget:

EXISTS = at least one of the shown files contains real source code that
implements one of the categories above. Examples that count:
   • A Python module that tokenises text, patches images into windows,
     or extracts audio frames.
   • A shell or Python script that downloads a dataset and cleans /
     normalises it (text OR image OR audio).
   • Code that builds prompts or assembles multimodal inputs before
     calling a model.
   • A data-loader class with non-trivial cleaning, normalisation,
     augmentation, or formatting.
   • An image-preprocessing pipeline using OpenCV / PIL / skimage with
     real logic (saliency, edge detection, patch extraction, ...).

NOT_APPLICABLE = everything else, including:
   • Repo contains only training, fine-tuning, or inference code (no preprocessing).
   • Only configs, READMEs, requirements files, hyperparameter dumps.
   • Only a thin stub / empty placeholder / data-download-only script.
   • Only model architecture definitions or evaluation code.
   • Files that read raw datasets but do no real cleaning / tokenizing /
     filtering / prompt-building.
   • Repo is empty or no relevant files were retrieved.

Be strict on TRIVIALITY (not modality): a file named `preprocess.py` that
just does `df = pd.read_csv(path)` is NOT real preprocessing code. But a
file that performs non-trivial transformations — text, image, audio, or
otherwise — IS preprocessing code regardless of modality.

OUTPUT MUST BE VALID JSON ONLY with this exact schema:

{
  "classification": "EXISTS" | "NOT_APPLICABLE",
  "confidence": integer 0-100,
  "preprocessing_files": [
    {
      "path":     "path/to/file.py",
      "category": one of "tokenization" | "filtering" | "prompting",
      "evidence": short code snippet (<= 200 chars) from that file that justifies the category
    },
    ...
  ],
  "diagnostic_notes": "one sentence",
  "final_reason": "one sentence for a human reader"
}

Confidence guidance (be honest — this signal drives whether the system
will run a more thorough re-check):
  0-50    — uncertain. You would want to see more files or longer file
            content before committing to this verdict.
  51-79   — moderately sure. Some ambiguity, but the evidence you saw
            points in one direction.
  80-100  — confident. The evidence is clear and you would not change
            your verdict given more files.

Rules for preprocessing_files:
- Empty array when classification is NOT_APPLICABLE.
- One object per (file, category) pair. If a single file matches two
  categories, emit two objects with the same path and different categories.
- evidence is a literal snippet from the file, not a paraphrase. Keep it
  short (one or two lines, <= 200 chars).
- category MUST be exactly one of the three lowercase strings shown above.

Few-shot examples illustrating the OUTPUT schema:

EXAMPLE EXISTS — text: a tokenize.py and a dedup.py:
{
  "classification": "EXISTS",
  "confidence": 95,
  "preprocessing_files": [
    {"path": "src/tokenize.py", "category": "tokenization",
     "evidence": "tokens = bpe.encode(text); shards.append(tokens[:max_len])"},
    {"path": "scripts/dedup.py", "category": "filtering",
     "evidence": "h = minhash(doc)\\nif h in seen: continue"}
  ],
  "diagnostic_notes": "Repo has BPE tokenization plus MinHash deduplication.",
  "final_reason": "Tokenization and filtering pipelines are present."
}

EXAMPLE EXISTS — vision: an image-preprocessing module:
{
  "classification": "EXISTS",
  "confidence": 90,
  "preprocessing_files": [
    {"path": "src/preprocess.py", "category": "tokenization",
     "evidence": "for window_size in [2560,1280,640]: patches.append(cv2.resize(patch, (output_size, output_size)))"},
    {"path": "src/preprocess.py", "category": "filtering",
     "evidence": "saliency = 0.6*edge_density + 0.4*texture_complexity; if saliency < threshold_low: continue"}
  ],
  "diagnostic_notes": "Adaptive image patching with saliency-based filtering.",
  "final_reason": "Images are split into patches and filtered by saliency before model input."
}

EXAMPLE NOT_APPLICABLE — only training / inference code:
{
  "classification": "NOT_APPLICABLE",
  "confidence": 85,
  "preprocessing_files": [],
  "diagnostic_notes": "Repo contains only train.py / model.py / evaluate.py.",
  "final_reason": "No preprocessing, tokenization, filtering, or prompt-building code found."
}

EXAMPLE NOT_APPLICABLE — uncertain (saw only a handful of files):
{
  "classification": "NOT_APPLICABLE",
  "confidence": 40,
  "preprocessing_files": [],
  "diagnostic_notes": "Only 5 of 200 files visible; nothing preprocessing-like in the sample.",
  "final_reason": "No preprocessing code in what I saw, but the sample is small — would benefit from more files."
}

When in doubt → NOT_APPLICABLE."""


# User message template. Always includes title, URL, and the fetched repo content.
USER_PROMPT = """Paper title: {title}
Repository: {repo_url}

=== REPOSITORY FILES ===
{repo_content}

Classify this repository now."""