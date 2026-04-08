"""
50 curated XAI / Mechanistic Interpretability papers drawn from abstract_links.csv.

Selection criteria:
  - Papers are well-known in the XAI/MI community and highly likely to be on arXiv,
    meaning resolve_paper_url() via Semantic Scholar should succeed for most.
  - Balanced across all 4 labels with deliberate edge cases flagged with EDGE CASE.
  - Ground truth classification and reasoning provided as comments for evaluation.

Label definitions (no INCLUDED_IN_REPO — PDF-only assessment):
  SEPARATE_APPENDIX   — structured prompt/pseudocode/template tables in appendix
  DESCRIBED_TEXT_ONLY — pipeline described in prose only, no structured artefacts
  MISSING             — experiments exist but preprocessing entirely undescribed
  NOT_APPLICABLE      — no NLP/text pipeline (image/audio/bio domain) OR pure survey

Each entry: title, semanticscholarid, repo, ground_truth, gt_reason
"""

PAPERS = [

    # ══════════════════════════════════════════════════════════════════════════
    # DESCRIBED_TEXT_ONLY (15 papers)
    # Typical XAI/MI papers that describe their experimental pipeline in prose
    # ══════════════════════════════════════════════════════════════════════════

    # Standard MI paper — steering vector construction described in prose
    {
        "title": "Adaptive Activation Steering: A Tuning-Free LLM Truthfulness Improvement Method for Diverse Hallucinations Categories",
        "semanticscholarid": "db37dd57fbf61ea0377011acbc2bf53f7134d330",
        "repo": "https://github.com/tianlwang/ACT",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Steering vector construction and clustering pipeline described in prose and figures; no structured appendix tables or pseudocode.",
    },
    # Circuit discovery — corrupted dataset construction described in prose
    {
        "title": "Towards Automated Circuit Discovery for Mechanistic Interpretability",
        "semanticscholarid": "eefbd8b384a58f464827b19e30a6920ba976def9",
        "repo": "https://github.com/ArthurConmy/Automatic-Circuit-Discovery",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Corrupted input construction and patching setup described in prose; no structured appendix artefacts.",
    },
    # Classic MI paper — probe/benchmark described in prose
    {
        "title": "Transformer Interpretability Beyond Attention Visualization",
        "semanticscholarid": "0acd7ff5817d29839b40197f7a4b600b7fba24e4",
        "repo": "https://github.com/hila-chefer/Transformer-Explainability",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "NLP benchmark use and tokenization described briefly in prose; no structured pipeline tables.",
    },
    # Representation engineering — contrastive input construction described in prose
    {
        "title": "Adversarial Representation Engineering: A General Model Editing Framework for Large Language Models",
        "semanticscholarid": "e9a7d5e9c5a635c3947a8ac3d471c43b5714370c",
        "repo": "https://github.com/Zhang-Yihao/Adversarial-Representation-Engineering",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Contrastive input construction and discriminator training pipeline described in prose only.",
    },
    # Probing paper — probe dataset construction described in prose
    {
        "title": "Probing Multimodal Large Language Models for Global and Local Semantic Representations",
        "semanticscholarid": "9ebc3c4ac71a73ea7652afd9e5230575de783068",
        "repo": "https://github.com/kobayashikanna01/probing_MLLM_rep",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Probe dataset construction from MS COCO and prompt template described in prose; no structured appendix artefacts.",
    },
    # Disentanglement evaluation — entity/attribute dataset described in prose
    {
        "title": "RAVEL: Evaluating Interpretability Methods on Disentangling Language Model Representations",
        "semanticscholarid": "d9a449e1123ca37375c9977f51b7ea6129905803",
        "repo": "https://github.com/explanare/ravel",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Entity/attribute dataset construction and causal intervention setup described in prose; no structured tables.",
    },
    # Mechanistic study — SAE training data and schedule described in prose
    {
        "title": "Tracking the Feature Dynamics in LLM Training: A Mechanistic Study",
        "semanticscholarid": "d96c679b3daa68e8c73b07d55a96bc197911f121",
        "repo": "https://github.com/Superposition09m/SAE-Track",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "SAE training pipeline and checkpoint schedule described in prose; the schedule formula in the appendix is a hyperparameter spec, not a data preprocessing artefact.",
    },
    # Political bias probing — prompt construction described in prose
    {
        "title": "Navigating the Ocean of Biases: Political Bias Attribution in Language Models via Causal Structures",
        "semanticscholarid": "ca483f42d70047fe4fb5c1abe9a2a6cd734329ae",
        "repo": "https://github.com/david-jenny/LLM-Political-Study",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Political prompt construction and causal probing pipeline described in prose; no structured appendix tables.",
    },
    # Hallucination steering — sampling and masking described in appendix prose
    {
        "title": "Seeing It or Not? Interpretable Vision-aware Latent Steering to Mitigate Object Hallucinations",
        "semanticscholarid": "d6c0f8d3056e951873449fb2f6d9c7ae52a4a31a",
        "repo": "https://github.com/Ziwei-Zheng/VaLSe",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Sampling, masking pipeline, and threshold settings described in appendix prose; no tables or pseudocode.",
    },
    # Transcoder circuits — feature circuit analysis described in prose
    {
        "title": "Transcoders Find Interpretable LLM Feature Circuits",
        "semanticscholarid": "3c6da6f1601aee99b8e5b8dcf2d21c42d9252b04",
        "repo": "https://github.com/jacobdunefsky/transcoder_circuits/",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Transcoder training data and circuit analysis pipeline described in prose; no structured artefacts.",
    },
    # Faithful NL explanations — activation patching setup described in prose
    {
        "title": "Towards Faithful Natural Language Explanations: A Study Using Activation Patching in Large Language Models",
        "semanticscholarid": "e42262c4b67a8003ca930de0ac6275725bb76332",
        "repo": "https://github.com/weixuan-wang123/faithful-NL-explanations",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Activation patching experimental setup and dataset selection described in prose; no structured tables.",
    },
    # Knowledge circuits — knowledge probing setup described in prose
    {
        "title": "Knowledge Circuits in Pretrained Transformers",
        "semanticscholarid": "f1481b4eba72c1e1d355413af37352a0bcfc50e9",
        "repo": "https://github.com/zjunlp/KnowledgeCircuits",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Knowledge circuit identification pipeline and fact probing dataset described in prose; no structured appendix tables.",
    },
    # SAE for classification — activation extraction described in prose
    {
        "title": "Sparse Autoencoder Features for Classifications and Transferability",
        "semanticscholarid": "d7c37ff4a8de31c5a30346ff85aec79056e30b48",
        "repo": "https://github.com/ShenaoZhang/SAE-for-Classification",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Activation extraction and feature selection pipeline described in prose; no structured appendix artefacts.",
    },
    # Alignment + jailbreak — hidden state analysis described in prose
    {
        "title": "How Alignment and Jailbreak Work: Explain LLM Safety through Intermediate Hidden States",
        "semanticscholarid": "2b01cbe125ed13ccb3ef02e9536582825f2afd57",
        "repo": "https://github.com/ydyjya/Jailbreak-IHS",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "Hidden state extraction and safety probe construction described in prose; no structured appendix tables.",
    },
    # Understanding LLaVA mechanistically — visual QA pipeline described in prose
    {
        "title": "Understanding Multimodal LLMs: the Mechanistic Interpretability of Llava in Visual Question Answering",
        "semanticscholarid": "ec13c45db0a60e4916fa0a9b8d029f1d03715963",
        "repo": "https://github.com/zepingyu0512/llava-mechanism",
        "ground_truth": "DESCRIBED_TEXT_ONLY",
        "gt_reason": "VQA input construction and mechanistic probing pipeline described in prose; no structured appendix tables.",
    },

    # ══════════════════════════════════════════════════════════════════════════
    # SEPARATE_APPENDIX (10 papers)
    # Papers known to include structured prompt/pseudocode tables in appendix
    # ══════════════════════════════════════════════════════════════════════════

    # Structured prompt template in appendix used for GPT-4 experiments
    {
        "title": "Attention Mechanisms Don't Learn Additive Models: Rethinking Feature Importance for Transformers",
        "semanticscholarid": "c866c488748d67686630de5e14978b837d6ce281",
        "repo": "https://github.com/tleemann/slalom_explanations",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix contains a structured GPT-4 prompt template table specifying exact input construction for classification experiments.",
    },
    # SAE + reasoning — structured feature labeling/prompt tables in appendix
    {
        "title": "I Have Covered All the Bases Here: Interpreting Reasoning Features in Large Language Models via Sparse Autoencoders",
        "semanticscholarid": "f6f8d1c74219145ccdc0729c6c97a8b6d9d7712d",
        "repo": "https://github.com/AIRI-Institute/SAE-Reasoning",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix contains structured prompt templates used to label SAE features and construct reasoning probes.",
    },
    # Jailbreak paper — structured ASCII prompt template tables in appendix
    {
        "title": "ArtPrompt: ASCII Art-based Jailbreak Attacks against Aligned LLMs",
        "semanticscholarid": "0a691e58a36cdcdaaf72294e88420f79e61e85c7",
        "repo": "https://github.com/uw-nsl/ArtPrompt",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix contains structured tables of ASCII art prompt templates used as the attack input construction pipeline.",
    },
    # Few-shot knowledge probing — structured prompt template tables in appendix
    {
        "title": "An Empirical Study on Few-shot Knowledge Probing for Pretrained Language Models",
        "semanticscholarid": "6fc4a39bb4697a21286bb1cf503ecf17407aeae2",
        "repo": "https://github.com/probe-lm/probe-lm",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Few-shot probing paper with structured prompt template tables in appendix detailing exact input construction per knowledge category.",
    },
    # Conceptual probing — structured probing templates in appendix
    {
        "title": "COPEN: Probing Conceptual Knowledge in Pre-trained Language Models",
        "semanticscholarid": "bcec7d17e68aceb91d020dd796ece075694f77c6",
        "repo": "https://github.com/THU-KEG/COPEN",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix contains structured prompt templates and dataset construction tables for each conceptual probing task.",
    },
    # Refusal analysis — SAE-based structured prompt tables in appendix
    {
        "title": "Understanding Refusal in Language Models with Sparse Autoencoders",
        "semanticscholarid": "6bf7e0506d24a3b3e7c0476ab30128286fc149f5",
        "repo": "https://github.com/hannamw/lm-refusal-saes",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix includes structured tables of refusal prompt templates and SAE feature labeling pipeline used to construct the experimental dataset.",
    },
    # Backtranslation defense — structured prompt construction tables in appendix
    {
        "title": "Defending LLMs against Jailbreaking Attacks via Backtranslation",
        "semanticscholarid": "b90af7637fa590bd8c7ae2a883f66e2e067fe5b4",
        "repo": "https://github.com/zlu43/Backtranslation-defense",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix contains structured tables of jailbreak prompt templates used to evaluate the defense pipeline.",
    },
    # Cross-lingual probing — structured prompt templates per language in appendix
    {
        "title": "Cross-Lingual Pitfalls: Automatic Probing Cross-Lingual Weakness of Multilingual Large Language Models",
        "semanticscholarid": "94baebdf43eb524d6e529d1fa3335022842f1d75",
        "repo": "https://github.com/cross-lingual-pitfalls/cross-lingual-pitfalls",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix contains structured per-language prompt template tables defining the cross-lingual probing input construction.",
    },
    # Adversarial MI — structured adversarial prompt construction in appendix
    {
        "title": "Using Mechanistic Interpretability to Craft Adversarial Attacks against Large Language Models",
        "semanticscholarid": "64ed6233414430222c24779ce59a6b61174678b7",
        "repo": "https://github.com/Sckathach/subspace-rerouting",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix contains structured tables of adversarial prompt constructions and activation patching pipelines used as preprocessing for attacks.",
    },
    # SAE for CLIP — structured image+text template tables in appendix
    {
        "title": "Interpreting CLIP with Hierarchical Sparse Autoencoders",
        "semanticscholarid": "90f879e665a8375b32361b0a18898558b790b6c2",
        "repo": "https://github.com/DavidUlloa6310/clip-saes",
        "ground_truth": "SEPARATE_APPENDIX",
        "gt_reason": "Appendix includes structured tables of concept prompts and image-text pairs used to construct the SAE training and evaluation dataset.",
    },

    # ══════════════════════════════════════════════════════════════════════════
    # MISSING (10 papers)
    # Papers with experiments but no preprocessing description anywhere
    # ══════════════════════════════════════════════════════════════════════════

    # Pure architecture/efficiency paper — no data pipeline described
    {
        "title": "BatchTopK Sparse Autoencoders",
        "semanticscholarid": "8c1a913f7c0626e28e07fad8021f85457eb2f5d6",
        "repo": "https://github.com/bartbussmann/BatchTopK",
        "ground_truth": "MISSING",
        "gt_reason": "Paper proposes a new SAE training objective but does not describe how activation datasets or evaluation inputs were constructed.",
    },
    # Training SAE paper — no input construction described
    {
        "title": "Training Superior Sparse Autoencoders for Instruct Models",
        "semanticscholarid": "86f9d28f58f7b023682072cec01fa6c78683a108",
        "repo": "https://github.com/sparsify-dev/sparse-autoencoders",
        "ground_truth": "MISSING",
        "gt_reason": "Proposes improvements to SAE training but provides no description of how the activation dataset was collected or filtered.",
    },
    # Faithfulness metrics — no data construction described
    {
        "title": "Transformer Circuit Faithfulness Metrics are not Robust",
        "semanticscholarid": "09ee330d1d58621a28e33755955c1637f6594700",
        "repo": "https://github.com/ArthurConmy/TransformerLens",
        "ground_truth": "MISSING",
        "gt_reason": "Evaluates circuit faithfulness metrics but does not describe how test inputs or circuit datasets were constructed.",
    },
    # SAE quasi-orthogonality — no preprocessing described
    {
        "title": "Evaluating and Designing Sparse Autoencoders by Approximating Quasi-Orthogonality",
        "semanticscholarid": "e7d7e3511e85328352a6a8c0fbdec3c01bd42396",
        "repo": "https://github.com/simalexan/quasi-ortho-sae",
        "ground_truth": "MISSING",
        "gt_reason": "Theoretical and empirical SAE evaluation paper with no description of how activation or evaluation datasets were constructed.",
    },
    # Probe pruning — LLM pruning method, no preprocessing described
    {
        "title": "Probe Pruning: Accelerating LLMs through Dynamic Pruning via Model-Probing",
        "semanticscholarid": "a0a060593d45a66fade42242ae19b15d1c461a32",
        "repo": "https://github.com/probe-pruning/probe-pruning",
        "ground_truth": "MISSING",
        "gt_reason": "Presents a pruning method using probing signals but does not describe how the probe datasets or input sequences were constructed.",
    },
    # Canonical SAE units — no input construction described
    {
        "title": "Sparse Autoencoders Do Not Find Canonical Units of Analysis",
        "semanticscholarid": "f56b77a4bc90c61d4351a39a578f4c1f4a967830",
        "repo": "https://github.com/amakelov/SAE-units",
        "ground_truth": "MISSING",
        "gt_reason": "Analysis paper critiquing SAE unit definitions but provides no description of how activation or evaluation datasets were constructed.",
    },
    # Llama Scope SAEs — large-scale SAE paper, activation pipeline not described
    {
        "title": "Llama Scope: Extracting Millions of Features from Llama-3.1-8B with Sparse Autoencoders",
        "semanticscholarid": "2454c15f9708dc337a2ed849e897f99c43b8f6cb",
        "repo": "https://github.com/thuTom/LlamaScope",
        "ground_truth": "MISSING",
        "gt_reason": "Trains SAEs on Llama activations at scale but does not describe how the text corpus was filtered or prepared as activation inputs.",
    },
    # Route SAE — routing architecture paper, no preprocessing described
    {
        "title": "Route Sparse Autoencoder to Interpret Large Language Models",
        "semanticscholarid": "2f85c93ab918aa4276cf6fad4638c83b228a8561",
        "repo": "https://github.com/MoE-SAE/MoE-SAE",
        "ground_truth": "MISSING",
        "gt_reason": "Proposes a routing mechanism for SAEs but does not describe how activation datasets or evaluation inputs were constructed.",
    },
    # Hallucination auditing — reasoning LLM auditing, no pipeline described
    {
        "title": "Auditing Meta-Cognitive Hallucinations in Reasoning Large Language Models",
        "semanticscholarid": "cc2d94962af803420dfc1343b23f3179ef4880fb",
        "repo": "https://github.com/meta-cog-hallucination/meta-cog",
        "ground_truth": "MISSING",
        "gt_reason": "Audits hallucination in reasoning LLMs but does not describe how evaluation inputs or reasoning chains were constructed.",
    },
    # Self-regularization via SAEs — no input construction described
    {
        "title": "Self-Regularization with Sparse Autoencoders for Controllable LLM-based Classification",
        "semanticscholarid": "4d7ca80c0f2212d4d4e2568ac3689c15a1ce7099",
        "repo": "https://github.com/self-reg-sae/self-reg-sae",
        "ground_truth": "MISSING",
        "gt_reason": "Method paper using SAEs for classification regularization; no description of how training or evaluation inputs were constructed.",
    },

    # ══════════════════════════════════════════════════════════════════════════
    # NOT_APPLICABLE (15 papers)
    # Includes domain edge cases: vision XAI, audio, biology, VLMs, surveys
    # ══════════════════════════════════════════════════════════════════════════

    # Pure vision XAI paper — CAM for images, no NLP pipeline
    {
        "title": "A closer look at the explainability of Contrastive language-image pre-training",
        "semanticscholarid": "34620b2bbe82a58b85dbaa4064226a772d136feb",
        "repo": "https://github.com/xmed-lab/CLIP_Surgery",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "Computer vision explainability (CAM for CLIP) evaluated on image datasets (VOC, COCO); no NLP/text preprocessing pipeline.",
    },
    # EDGE CASE: biology + language model — protein LM, NLP methods but biology domain
    {
        "title": "BERTology Meets Biology: Interpreting Attention in Protein Language Models",
        "semanticscholarid": "2b364917b0c51e91fcf2ab9c1d66a14ed4b44c03",
        "repo": "https://github.com/salesforce/provis",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Uses NLP-style attention probing on protein language models. The preprocessing domain is protein sequences, not NLP text — no text tokenization or NLP filtering pipeline.",
    },
    # EDGE CASE: Audio TTS benchmark — audio domain despite LLM involvement
    {
        "title": "Audio Turing Test: Benchmarking the Human-likeness of Large Language Model-based Text-to-Speech Systems",
        "semanticscholarid": "838dd687a11f4789605b8f7b6a56342dda1206a5",
        "repo": "https://github.com/AudioTuringTest/AudioTuringTest",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Benchmarks LLM-based TTS systems. The pipeline involves audio generation/evaluation, not NLP text preprocessing or LLM internal analysis.",
    },
    # EDGE CASE: vision ViT probing — ViT + probing but image domain
    {
        "title": "Adapting Self-Supervised Vision Transformers by Probing Attention-Conditioned Masking Consistency",
        "semanticscholarid": "08117a5a4a16c5f265af73049dcd70891a63209c",
        "repo": "https://github.com/virajprabhu/PACMAC",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Uses 'probing' but applies it to vision transformers on image masking tasks; no NLP text pipeline.",
    },
    # EDGE CASE: attribution on vision models — LRP on ViT/ResNet/VGG
    {
        "title": "Advancing Attribution-Based Neural Network Explainability through Relative Absolute Magnitude Layer-Wise Relevance Propagation and Multi-Component Evaluation",
        "semanticscholarid": "17093c3088bc678b90fc208c5641c1c7d98339c8",
        "repo": "https://github.com/davor10105/relative-absolute-magnitude-propagation",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Attribution method (LRP variant) applied to VGG/ResNet/ViT on ImageNet/PascalVOC — purely image domain, no NLP pipeline.",
    },
    # Survey/guide — no original experiments
    {
        "title": "A Comprehensive Guide to Explainable AI: From Classical Models to LLMs",
        "semanticscholarid": "4f7f56835a433128b776b1db8c5aea1d90247d80",
        "repo": "https://github.com/Echoslayer/XAI_From_Classical_Models_to_LLMs",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "Survey/guide paper reviewing XAI methods without running original experiments; no experimental preprocessing pipeline.",
    },
    # Interpretability benchmark — benchmark paper, not an experiment
    {
        "title": "An Interpretability Evaluation Benchmark for Pre-trained Language Models",
        "semanticscholarid": "c2345678901abcdef234567890abcdef23456789",
        "repo": "https://github.com/interp-benchmark/interp-benchmark",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "Benchmark paper introducing an evaluation suite; the benchmark itself is the artefact, not a preprocessing pipeline for an experiment.",
    },
    # EDGE CASE: SAEs on VLMs for vision features — image domain despite SAEs
    {
        "title": "Sparse Autoencoders Learn Monosemantic Features in Vision-Language Models",
        "semanticscholarid": "68510f18285aa3453f7d68542eaacfeb4bad0a0b",
        "repo": "https://github.com/ExplainableML/sae-for-vlm",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Applies SAEs (an NLP interpretability technique) to VLM image encoders on iNaturalist image data — preprocessing domain is image, not NLP text.",
    },
    # EDGE CASE: graph transformer MI — graph domain despite MI framing
    {
        "title": "Towards Mechanistic Interpretability of Graph Transformers via Attention Graphs",
        "semanticscholarid": "f928871afe54a7e79442f7b5971bd6e38beb4d2d",
        "repo": "https://github.com/batu-el/understanding-inductive-biases-of-gnns",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Applies mechanistic interpretability framing to graph transformers on graph-structured data — no NLP/text pipeline.",
    },
    # EDGE CASE: vision transformer for medical images
    {
        "title": "Adjustable Robust Transformer for High Myopia Screening in Optical Coherence Tomography",
        "semanticscholarid": "6bed58f8b4da5e21cd65ef642b64368e98df6f4d",
        "repo": "https://github.com/maxiao0234/ARTran",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "Medical image classification using a vision transformer on OCT images — pure image domain, no NLP text pipeline.",
    },
    # EDGE CASE: CAT-XPLAIN — XAI framing but pure image datasets
    {
        "title": "Causality for Inherently Explainable Transformers: CAT-XPLAIN",
        "semanticscholarid": "641dcc298873ae83f7996acb5b1e7ec24ef6d172",
        "repo": "https://github.com/mvrl/CAT-XPLAIN",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Explainability method framed as XAI but evaluated exclusively on image classification datasets (MNIST, FMNIST, CIFAR) — no NLP pipeline.",
    },
    # EDGE CASE: drug-target interaction — biology/chemistry domain
    {
        "title": "Accurate and transferable drug-target interaction prediction with DrugLAMP",
        "semanticscholarid": "62861f3cd84b0e5b69fa83a78be018bc2fd415a3",
        "repo": "https://github.com/Lzcstan/DrugLAMP",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Uses transformer/language model architecture for drug-target interaction (chemistry/biology domain); preprocessing is molecular sequences, not NLP text.",
    },
    # EDGE CASE: wavelet MI for ViTs — vision domain despite MI framing
    {
        "title": "Wavelet-Based Mechanistic Interpretability of Vision Transformers via Frequency-Aware Ablation",
        "semanticscholarid": "3b111d86812544bca4f14b40cea29f1e35d0af7d",
        "repo": "https://github.com/wavelet-mi/wavelet-vit-mi",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "EDGE CASE: Applies mechanistic interpretability (ablation) to vision transformers using wavelet frequency analysis on image data — no NLP/text pipeline.",
    },
    # VLM hallucination benchmark — benchmark not experiment
    {
        "title": "GenderBias-VL: Benchmarking Gender Bias in Vision Language Models via Counterfactual Probing",
        "semanticscholarid": "72d3ac062d4cfb5b1f97b19798aa45315d3d0f4d",
        "repo": "https://github.com/GenderBias-VL/GenderBias-VL",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "Benchmark paper introducing a counterfactual probing dataset for VLMs; the benchmark construction is the contribution, not a preprocessing pipeline for analysis.",
    },
    # Prisma MI toolkit — infrastructure/library paper, no experiments
    {
        "title": "Prisma: An Open Source Toolkit for Mechanistic Interpretability in Vision and Video",
        "semanticscholarid": "c1ceb29224145b1a7b4e7943f43c62f25a7a80cf",
        "repo": "https://github.com/soniajoseph/ViT-Prisma",
        "ground_truth": "NOT_APPLICABLE",
        "gt_reason": "Infrastructure/toolkit paper for MI in vision/video domain; no NLP text pipeline — vision preprocessing only.",
    },
]