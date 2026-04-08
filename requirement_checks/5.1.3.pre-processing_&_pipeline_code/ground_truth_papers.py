PAPERS = [{
    "title": "Tokenization as Finite-State Transduction",
    "semanticscholarid": "2410.15696",
    "repo": "",
    "ground_truth": "SEPARATE_APPENDIX",
    "gt_reason": "Algorithm 1 (MaxMatch tokenization inference) and Algorithm 5 (generic transducer composition) appear in the main body and appendix — explicit pseudocode for the tokenization pipeline including greedy left-to-right token selection steps.",
},
{
    "title": "AICC: Parse HTML Finer, Make Models Better",
    "semanticscholarid": "2511.16397",
    "repo": "https://github.com/AICC-Project/AICC",
    "ground_truth": "SEPARATE_APPENDIX",
    "gt_reason": "Appendix 6.2 contains pseudocode for the content list to Markdown conversion pipeline — explicit step-by-step logic iterating over content list elements and applying type-specific rendering rules.",
},
{
    "title": "Achieving Tokenizer Flexibility in Language Models through Heuristic Adaptation and Supertoken Learning",
    "semanticscholarid": "2505.09738",
    "repo": "",
    "ground_truth": "SEPARATE_APPENDIX",
    "gt_reason": "Appendix A contains Algorithm 2 (Supertoken Tokenizer Training): pseudocode showing pre-tokenization, stochastic chunking, and BPE training steps as a concrete tokenization pipeline.",
},
{
    "title": "Boundless Byte Pair Encoding: Breaking the Pre-tokenization Barrier",
    "semanticscholarid": "2504.00178",
    "repo": "https://github.com/kensho-technologies/boundless-bpe",
    "ground_truth": "SEPARATE_APPENDIX",
    "gt_reason": (
        "Appendix A contains the BOUNDLESS_PATTERN regex definition and "
        "Appendix B contains Listing 2 — code branches B-2 to B-6 showing "
        "how variable/function names are split into pretokens based on "
        "capitalization. This is executable pre-tokenization pipeline code."
    ),
},

    # ══════════════════════════════════════════════════════════════════════════
    # MISSING
    # ══════════════════════════════════════════════════════════════════════════
 
    {
        "title": "Adaptive Activation Steering: A Tuning-Free LLM Truthfulness Improvement Method for Diverse Hallucinations Categories",
        "semanticscholarid": "db37dd57fbf61ea0377011acbc2bf53f7134d330",
        "repo": "https://github.com/tianlwang/ACT",
        "ground_truth": "MISSING",
        "gt_reason": "Steering vector construction and clustering pipeline described in prose and figures; no structured appendix tables or pseudocode.",
    },
    {
        "title": "Towards Automated Circuit Discovery for Mechanistic Interpretability",
        "semanticscholarid": "eefbd8b384a58f464827b19e30a6920ba976def9",
        "repo": "https://github.com/ArthurConmy/Automatic-Circuit-Discovery",
        "ground_truth": "MISSING",
        "gt_reason": "Corrupted input construction and patching setup described in prose; no structured appendix artefacts.",
    },
    {
        "title": "Transformer Interpretability Beyond Attention Visualization",
        "semanticscholarid": "0acd7ff5817d29839b40197f7a4b600b7fba24e4",
        "repo": "https://github.com/hila-chefer/Transformer-Explainability",
        "ground_truth": "MISSING",
        "gt_reason": "NLP benchmark use and tokenization described briefly in prose; no structured pipeline tables.",
    },
    {
        "title": "Adversarial Representation Engineering: A General Model Editing Framework for Large Language Models",
        "semanticscholarid": "e9a7d5e9c5a635c3947a8ac3d471c43b5714370c",
        "repo": "https://github.com/Zhang-Yihao/Adversarial-Representation-Engineering",
        "ground_truth": "MISSING",
        "gt_reason": "Contrastive input construction and discriminator training pipeline described in prose only.",
    },
    {
        "title": "Probing Multimodal Large Language Models for Global and Local Semantic Representations",
        "semanticscholarid": "9ebc3c4ac71a73ea7652afd9e5230575de783068",
        "repo": "https://github.com/kobayashikanna01/probing_MLLM_rep",
        "ground_truth": "MISSING",
        "gt_reason": "Probe dataset construction from MS COCO and prompt template described in prose; no structured appendix artefacts.",
    },
    {
        "title": "RAVEL: Evaluating Interpretability Methods on Disentangling Language Model Representations",
        "semanticscholarid": "d9a449e1123ca37375c9977f51b7ea6129905803",
        "repo": "https://github.com/explanare/ravel",
        "ground_truth": "MISSING",
        "gt_reason": "Entity/attribute dataset construction and causal intervention setup described in prose; no structured tables.",
    },
    {
        "title": "Tracking the Feature Dynamics in LLM Training: A Mechanistic Study",
        "semanticscholarid": "d96c679b3daa68e8c73b07d55a96bc197911f121",
        "repo": "https://github.com/Superposition09m/SAE-Track",
        "ground_truth": "MISSING",
        "gt_reason": "SAE training pipeline and checkpoint schedule described in prose; the schedule formula in the appendix is a hyperparameter spec, not a data preprocessing artefact.",
    },
    {
        "title": "Navigating the Ocean of Biases: Political Bias Attribution in Language Models via Causal Structures",
        "semanticscholarid": "ca483f42d70047fe4fb5c1abe9a2a6cd734329ae",
        "repo": "https://github.com/david-jenny/LLM-Political-Study",
        "ground_truth": "MISSING",
        "gt_reason": "Political prompt construction and causal probing pipeline described in prose; no structured appendix tables.",
    },
    {
        "title": "Seeing It or Not? Interpretable Vision-aware Latent Steering to Mitigate Object Hallucinations",
        "semanticscholarid": "d6c0f8d3056e951873449fb2f6d9c7ae52a4a31a",
        "repo": "https://github.com/Ziwei-Zheng/VaLSe",
        "ground_truth": "MISSING",
        "gt_reason": "Sampling, masking pipeline, and threshold settings described in appendix prose; no tables or pseudocode.",
    },
    {
        "title": "Transcoders Find Interpretable LLM Feature Circuits",
        "semanticscholarid": "3c6da6f1601aee99b8e5b8dcf2d21c42d9252b04",
        "repo": "https://github.com/jacobdunefsky/transcoder_circuits/",
        "ground_truth": "MISSING",
        "gt_reason": "Transcoder training data and circuit analysis pipeline described in prose; no structured artefacts.",
    },
]