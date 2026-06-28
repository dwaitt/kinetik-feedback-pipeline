# BDR Feedback Digest Pipeline

An automated weekly intelligence pipeline built for a B2B signal intelligence pilot program. The system ingests raw field feedback from sales reps, synthesizes it using Claude AI, and automatically routes structured outputs to Microsoft Teams and Planner, eliminating manual triage in sprint planning.

**Built by:** Asna Muzafar & David Waitt at [Kinetik](https://kinetik.solutions)

---

## What it does

Each week, BDRs submit feedback through Microsoft Forms, session notes, and meeting transcripts. This pipeline:

1. **Ingests** three input types: Microsoft Forms responses (Excel), partner session notes (Markdown/DOCX), and meeting transcripts (VTT/DOCX/TXT)
2. **Synthesizes** everything using the Claude API into a structured weekly digest, ranked by severity, with deduplication across sources
3. **Outputs** three files: a human-readable Markdown digest, a backlog JSON for Planner ingestion, and an account-requests JSON
4. **Automates delivery** via Power Automate: the digest posts to a Teams channel and backlog items become Planner cards with zero manual handoff

---

## Architecture

```
BDR Inputs                    Python Pipeline              Power Automate
─────────────────             ─────────────────            ──────────────────────
Microsoft Forms (Excel)  ──►
Partner Notes (MD/DOCX)  ──►  feedback_digest.py  ──►  SharePoint drop folder
Meeting Transcripts(VTT) ──►  + Claude API                    |
                                                               ├──► Teams: weekly digest post
                                                               └──► Planner: backlog cards
```

---

## Repo structure

```
kinetik-feedback-pipeline/
├── feedback_digest.py              # Core Python pipeline
├── .env.example                    # Environment variable template
├── .gitignore
├── /power-automate
│   ├── Backlog-JSON-to-Planner.zip     # Flow: reads backlog JSON to Planner cards
│   └── Weekly-Digest-Teams-Post.zip   # Flow: reads digest MD to Teams post
└── README.md
```

---

## Setup

**Requirements:** Python 3.10+, an Anthropic API key, Microsoft 365 (Teams, Planner, SharePoint)

```bash
# 1. Clone the repo
git clone https://github.com/dwaitt/kinetik-feedback-pipeline.git
cd kinetik-feedback-pipeline

# 2. Install dependencies
pip install anthropic openpyxl python-dotenv python-docx

# 3. Configure environment
cp .env.example .env
# Edit .env with your ANTHROPIC_API_KEY and BASE_DIR path

# 4. Run
python feedback_digest.py
```

**Expected folder structure under `BASE_DIR`:**

```
BASE_DIR/
├── Forms/              <- Microsoft Forms Excel export
├── Partner Notes/      <- Session notes (.md, .txt, .docx)
├── Transcripts/        <- Meeting recordings (.vtt, .txt, .md, .docx)
└── Digest Outputs/     <- Auto-created; digest and JSON outputs land here
```

---

## Power Automate flows

Import the ZIPs from `/power-automate/` into your Power Automate environment:

**Backlog-JSON-to-Planner** - Reads `backlog_<week>.json` from the SharePoint output folder. For each item in the array, creates a Microsoft Planner card with title, priority, module, and effort populated from the JSON schema.

**Weekly-Digest-Teams-Post** - Reads `digest_<week>.md` from the SharePoint output folder and posts the formatted content to a designated Microsoft Teams channel.

To import: Power Automate -> My Flows -> Import -> Upload package (.zip) -> configure connections.

---

## Key features

- **Multi-format ingestion** - handles Excel, Markdown, plain text, DOCX, and VTT transcript files
- **SHA-256 state tracking** - skips unchanged files across runs so nothing gets reprocessed
- **Fingerprint deduplication** - normalizes item titles and account names so the same issue never creates two Planner cards across weeks
- **Severity ranking** - Claude outputs items ranked Critical, High, Medium, Low with source attribution
- **New account detection** - strict logic separates genuine "please add this account" requests from accounts merely mentioned in conversation

---

## Tech stack

Python · Anthropic Claude API · Microsoft Power Automate · SharePoint · Microsoft Teams · Microsoft Planner · OpenPyXL · python-docx
