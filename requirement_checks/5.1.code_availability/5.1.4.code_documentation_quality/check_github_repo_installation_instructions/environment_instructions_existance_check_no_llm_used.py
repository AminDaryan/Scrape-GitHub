"""
Environment / installation instructions checker (rule-based, no LLM).

Answers: does this GitHub repo's README contain installation instructions?

Logic — four priority tiers, checked in order; the first match wins:
  1. Regex scan for definitive install commands.                       → classify_text_for_setup()
  2. Strong heading + keyword confirmation.                            → classify_text_for_setup(), split_into_sections()
  3. Partial heading + keyword + NLP confirmation.                     → classify_text_for_setup()
  4. NLP zero-shot classification on every section as last resort.     → classify_text_for_setup()

Entry point: check_setup_with_nlp() fetches the README and calls classify_text_for_setup().

Returns "Yes" / "No" along with the matched evidence tier.
"""

import re
import base64
import requests
from transformers import pipeline

# =============================================================================
# NLP MODEL — only loaded and used as last resort (Priority 4)
# =============================================================================
print("Loading NLP model...")
classifier = pipeline("zero-shot-classification", model="cross-encoder/nli-deberta-v3-base")
print("Model loaded.")

# =============================================================================
# PRIORITY 1: DEFINITIVE INSTALL COMMANDS
# These are unambiguous shell/package-manager commands.
# If ANY of these appear anywhere in the README → instant "Yes", no NLP needed.
# Covers: Python, Node, Ruby, PHP, Rust, Go, Java, .NET, Docker, system packages
# =============================================================================
DEFINITIVE_INSTALL_COMMANDS = [
    # Python
    r"pip install\s+\S+",           # pip install flask
    r"pip3 install\s+\S+",          # pip3 install flask
    r"pipenv install",               # pipenv install records[pandas]
    r"python -m pip install",        # python -m pip install X
    r"python setup\.py install",     # python setup.py install
    r"conda install",                # conda install numpy
    r"conda env create",             # conda env create -f environment.yml
    r"poetry install",               # poetry install
    r"easy_install\s+\S+",          # easy_install package (older Python)

    # JavaScript / Node
    r"npm install",                  # npm install
    r"npm i ",                       # npm i package
    r"yarn add\s+\S+",              # yarn add react
    r"yarn install",                 # yarn install
    r"pnpm install",                 # pnpm install

    # System package managers
    r"apt-get install",              # apt-get install something
    r"apt install",                  # apt install something
    r"brew install\s+\S+",          # brew install thefuck
    r"brew tap\s+\S+",              # brew tap repo/name
    r"yum install",                  # yum install package
    r"dnf install",                  # dnf install package
    r"pacman -s",                    # pacman -S package (Arch)
    r"snap install",                 # snap install package

    # Other languages
    r"gem install\s+\S+",           # gem install rails (Ruby)
    r"cargo install",               # cargo install (Rust)
    r"cargo add",                   # cargo add package (Rust)
    r"go install",                  # go install (Go)
    r"go get",                      # go get package (Go)
    r"composer require",            # composer require vendor/package (PHP)
    r"dotnet add package",          # dotnet add package X (.NET)
    r"nuget install",               # nuget install package (.NET)
    r"mvn install",                 # mvn install (Java Maven)
    r"gradle install",              # gradle install (Java)

    # Docker
    r"docker-compose up",           # docker-compose up
    r"docker compose up",           # docker compose up
    r"docker pull\s+\S+",          # docker pull image
    r"docker build",                # docker build -t name .
    r"docker run",                  # docker run image

    # Build tools
    r"make install",                # make install
    r"make setup",                  # make setup
    r"cmake",                       # cmake . (C/C++)
    r"./configure",                 # ./configure (autotools)
    r"./install",                   # ./install.sh
    r"bash install",                # bash install.sh
    r"sh install",                  # sh install.sh
]

# =============================================================================
# PRIORITY 2: STRONG HEADING PATTERNS
# Section headings that almost certainly contain installation instructions.
# Note: NO ^ anchor — allows symbols/emoji before the word (e.g. ☤ Installation)
# If heading matches AND content has any keyword → instant "Yes", no NLP needed.
# =============================================================================
STRONG_HEADING_PATTERNS = [
    r"install",           # Installation, Installing, Install
    r"setup",             # Setup, Set Up
    r"set up",            # Set Up
    r"getting started",   # Getting Started
    r"quick start",       # Quick Start, Quickstart
    r"prerequisites",     # Prerequisites
    r"requirements",      # Requirements
]

# =============================================================================
# PRIORITY 3: PARTIAL HEADING PATTERNS
# Headings that might contain setup info but need keyword confirmation in content.
# =============================================================================
PARTIAL_HEADING_PATTERNS = [
    r"usage",             # Usage (sometimes has install steps)
    r"how to",            # How to run, How to use
    r"build",             # Build
    r"deploy",            # Deployment
    r"environment",       # Environment
    r"configuration",     # Configuration
    r"develop",           # Development
    r"local",             # Local setup, Running locally
    r"bootstrap",         # Bootstrapping
    r"first steps",       # First Steps
    r"begin",             # Beginning, Get going
    r"start",             # Start, Starting
    r"workflow",          # Workflow
    r"onboard",           # Onboarding
    r"init",              # Initialization
    r"prepare",           # Preparation
    r"launch",            # Launching
]

# =============================================================================
# KEYWORD CONFIRMATION LIST
# Used in Priorities 2 and 3 to confirm a section has real setup content.
# A section heading match alone is not enough — content must also have one of these.
# =============================================================================
CONFIRMATION_KEYWORDS = [
    "install", "setup", "set up",
    "pip", "npm", "yarn", "brew", "conda", "docker",
    "clone", "git clone",
    "requirements", "dependencies",
    "environment", "virtualenv", "venv",
    "run ", "running", "execute",
    "configure", "configuration",
    "make", "build",
    ".env", "prerequisites",
]

# =============================================================================
# NLP SETUP (Priority 4 only)
# =============================================================================
CANDIDATE_LABELS = [
    "environment setup instructions",
    "installation guide",
    "getting started guide",
    "general information",
    "unrelated content",
]
SETUP_LABELS = {
    "environment setup instructions",
    "installation guide",
    "getting started guide",
}


# =============================================================================
# HELPER: SPLIT README INTO SECTIONS BY MARKDOWN HEADINGS
# Returns list of {"heading": str, "content": str}
# If no headings found, returns whole text as one section.
# =============================================================================
def split_into_sections(text):
    """Split a Markdown README into sections based on heading lines (# to ####).

    Returns a list of {"heading": str, "content": str} dicts. If the README has
    no headings at all (plain-prose style), the entire text is returned as one
    unnamed section so the rest of the detection logic still works.
    """
    heading_pattern = re.compile(r"^(#{1,4})\s+(.+)$", re.MULTILINE)
    matches = list(heading_pattern.finditer(text))

    if not matches:
        # No headings at all — return whole text as one unnamed section
        return [{"heading": "", "content": text}]

    sections = []
    for i, match in enumerate(matches):
        heading = match.group(2).strip()
        start = match.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start:end].strip()
        sections.append({"heading": heading, "content": content})

    return sections


# =============================================================================
# MAIN DETECTION FUNCTION
# Tries 4 strategies in order — stops as soon as one succeeds.
# =============================================================================
def classify_text_for_setup(text, threshold=0.7):
    """Run the four-tier detection pipeline on README text.

    Returns (label, confidence, snippet) where label is one of the CANDIDATE_LABELS
    on a hit, or None on a miss. Tiers are tried in order and the function returns
    immediately on the first match so cheap regex checks always run before the NLP model.
    The NLP tier uses a higher confidence threshold (threshold + 0.15) because there
    is no corroborating signal from headings or keywords to back it up.
    """
    if not text.strip():
        return None, 0.0, ""

    # ------------------------------------------------------------------
    # PRIORITY 1: Scan entire raw text for definitive install commands
    # Runs on the FULL text before any section splitting.
    # This catches install commands regardless of where they appear
    # or what heading they're under.
    # e.g. "pipenv install records", "brew install thefuck", "docker run X"
    # ------------------------------------------------------------------
    text_lower = text.lower()
    for pattern in DEFINITIVE_INSTALL_COMMANDS:
        match = re.search(pattern, text_lower)
        if match:
            # Extract surrounding context as a readable snippet
            start = max(0, match.start() - 20)
            end = min(len(text), match.end() + 60)
            snippet = text[start:end].replace("\n", " ").strip()
            return "installation guide", 1.0, f"[command] ...{snippet}..."

    # ------------------------------------------------------------------
    # PRIORITY 2 & 3: Section-based detection
    # Split README into sections and check each one.
    # ------------------------------------------------------------------
    sections = split_into_sections(text)

    for section in sections:
        # Strip symbols/emoji from heading for reliable pattern matching
        # e.g. "☤ Installation" → "Installation"
        raw_heading = section["heading"]
        clean_heading = re.sub(r"[^\w\s]", "", raw_heading).strip().lower()
        content = section["content"]
        content_lower = content.lower()

        if len(content) < 10:
            continue  # skip empty sections

        # Check if content has any confirmation keyword
        keyword_in_content = any(kw in content_lower for kw in CONFIRMATION_KEYWORDS)

        # --------------------------------------------------------------
        # PRIORITY 2: Strong heading match
        # Heading clearly signals installation (e.g. "Installation", "Setup")
        # + at least one keyword in content → instant Yes, skip NLP
        # --------------------------------------------------------------
        strong_match = any(re.search(p, clean_heading) for p in STRONG_HEADING_PATTERNS)
        if strong_match and keyword_in_content:
            snippet = content[:100].replace("\n", " ")
            return "installation guide", 1.0, f"[heading: {raw_heading}] {snippet}"

        # --------------------------------------------------------------
        # PRIORITY 3: Partial heading match + keyword in content → NLP confirms
        # Heading is ambiguous (e.g. "Usage", "Build") but content looks promising
        # NLP makes the final call here
        # --------------------------------------------------------------
        partial_match = any(re.search(p, clean_heading) for p in PARTIAL_HEADING_PATTERNS)
        if partial_match and keyword_in_content:
            clean_content = re.sub(r"[#*`>\[\]!]", "", content).strip()
            model_input = f"{raw_heading}: {clean_content[:1500]}"
            result = classifier(model_input, CANDIDATE_LABELS)
            top_label = result["labels"][0]
            top_score = result["scores"][0]
            if top_label in SETUP_LABELS and top_score >= threshold:
                snippet = content[:100].replace("\n", " ")
                return top_label, top_score, f"[heading: {raw_heading}] {snippet}"

    # ------------------------------------------------------------------
    # PRIORITY 4: NLP as last resort — no heading or command matched
    # Runs NLP on ALL sections. Catches:
    # - Prose-style READMEs with no headings
    # - Completely unconventional section names
    # Uses a higher threshold (threshold + 0.15) since there's no
    # corroborating signal from headings or keywords.
    # ------------------------------------------------------------------
    for section in sections:
        content = section["content"]
        if len(content) < 50:
            continue  # need enough text for NLP to work reliably

        clean_content = re.sub(r"[#*`>\[\]!]", "", content).strip()
        model_input = f"{section['heading']}: {clean_content[:1500]}" if section["heading"] else clean_content[:1500]

        result = classifier(model_input, CANDIDATE_LABELS)
        top_label = result["labels"][0]
        top_score = result["scores"][0]

        # Higher threshold here — no heading/keyword to back it up
        if top_label in SETUP_LABELS and top_score >= (threshold + 0.15):
            snippet = content[:100].replace("\n", " ")
            return top_label, top_score, f"[NLP only: {section['heading'] or 'no heading'}] {snippet}"

    return None, 0.0, ""


# =============================================================================
# FETCH README AND RUN DETECTION
# =============================================================================
def check_setup_with_nlp(owner, repo, headers=None, threshold=0.7, check_all_readmes=False):
    """Fetch the root README from GitHub and run install-instruction detection on it.

    When check_all_readmes=True, also scans every README found in subdirectories.
    This catches monorepos where each sub-package has its own setup instructions,
    but it triggers many extra API calls so it is off by default.
    Returns a human-readable string: "Yes: <evidence>" or "No".
    """
    readme_url = f"https://api.github.com/repos/{owner}/{repo}/readme"
    readme_response = requests.get(readme_url, headers=headers)

    if readme_response.status_code != 200:
        return "No README found"

    readme_content = base64.b64decode(
        readme_response.json()["content"]
    ).decode("utf-8", errors="ignore")

    found_locations = []
    label, score, snippet = classify_text_for_setup(readme_content, threshold)
    if label:
        found_locations.append(f"README ({label}, confidence: {score:.0%}) — {snippet}")

    # Optional: scan all READMEs in subdirectories too
    if check_all_readmes:
        tree_url = f"https://api.github.com/repos/{owner}/{repo}/git/trees/HEAD?recursive=1"
        tree_response = requests.get(tree_url, headers=headers)
        if tree_response.status_code == 200:
            all_files = [
                item["path"] for item in tree_response.json().get("tree", [])
                if item["type"] == "blob"
            ]
            for filepath in all_files:
                filename = filepath.split("/")[-1].lower()
                if filename.startswith("readme") and "/" in filepath:
                    content = fetch_file_content(owner, repo, filepath, headers)
                    if content:
                        label, score, snippet = classify_text_for_setup(content, threshold)
                        if label:
                            found_locations.append(
                                f"{filepath} ({label}, confidence: {score:.0%}) — {snippet}"
                            )

    return "Yes: " + " | ".join(found_locations) if found_locations else "No"