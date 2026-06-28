#!/usr/bin/env python3
"""
BDR Feedback Digest Pipeline
=============================

Synthesises BDR Microsoft Forms responses + partner session notes
+ meeting transcripts (vtt and docx) into a prioritised weekly digest,
a structured backlog payload ready for Power Automate to push into Planner,
and a separate account-request list for new accounts BDRs want to see in
the pipeline.

Setup:
    pip install anthropic openpyxl python-dotenv python-docx
    Copy .env.example to .env and fill in your values.
    python feedback_digest.py

Required environment variables (see .env.example):
    ANTHROPIC_API_KEY   Your Anthropic API key
    BASE_DIR            Absolute path to the root feedback folder

Expected folder structure under BASE_DIR:
    Forms/              Microsoft Forms Excel export (.xlsx)
    Partner Notes/      Session notes (.md, .txt, .docx)
    Transcripts/        Meeting transcripts (.vtt, .txt, .md, .docx)

Outputs (written to BASE_DIR/Digest Outputs/):
    digest_<week>.md              human-readable weekly digest
    backlog_<week>.json           structured cards for Planner ingestion
    account_requests_<week>.json  new account asks (one card per account)

State (so we don't reprocess unchanged files or duplicate Planner cards):
    BASE_DIR/Digest Outputs/.state.json
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from anthropic import Anthropic
from docx import Document
from dotenv import load_dotenv
from openpyxl.utils.exceptions import InvalidFileException

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

script_dir = Path(__file__).parent
load_dotenv(script_dir / ".env")
load_dotenv(script_dir / "key.env")  # fallback for local dev

MODEL = "claude-opus-4-8"

# Root folder containing Forms/, Partner Notes/, and Transcripts/ subfolders.
# Set BASE_DIR in your .env file — see .env.example.
_base_dir_env = os.environ.get("BASE_DIR")
if not _base_dir_env:
    print("error: BASE_DIR is not set. Add it to your .env file.", file=sys.stderr)
    print("  Example: BASE_DIR=/path/to/your/feedback/folder", file=sys.stderr)
    sys.exit(1)
BASE_DIR = Path(_base_dir_env)

FORMS_XLSX_PATH = BASE_DIR / "Forms" / "Signals Intelligence AI Pilot - Quick Feedback.xlsx"
PARTNER_NOTES_DIR = BASE_DIR / "Partner Notes"
TRANSCRIPTS_DIR = BASE_DIR / "Transcripts"
OUTPUT_DIR = BASE_DIR / "Digest Outputs"
STATE_PATH = OUTPUT_DIR / ".state.json"

WEEK_LABEL = datetime.now().strftime("%Y-W%V")

# Form columns we care about (partial-match on header text, case-insensitive)
FORM_COL_MODULE = "module is your feedback"
FORM_COL_TYPE = "type of feedback"
FORM_COL_PRIORITY = "priority of this feedback"
FORM_COL_DESCRIBE = "describe your feedback"
FORM_COL_ACCOUNT_REL = "account name"
FORM_COL_ACCOUNT_NEW = "new account you would like"
FORM_COL_NAME = "your name"
FORM_COL_COMPLETED = "completion time"

# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """You are the synthesis layer of a B2B signal intelligence pilot feedback loop. You receive raw feedback from three sources every week:

1. BDR-submitted Microsoft Forms responses (structured 60-second submissions, plus optional new-account asks).
2. Partner session notes (rich qualitative observations from 1:1 walkthroughs; some are Word docs).
3. Meeting transcripts (vtt or docx; raw text, may contain timestamp noise that has been pre-stripped).

Your job is to produce a weekly digest that the delivery lead can read in 3 minutes and act on in sprint planning the same day.

# Output structure

Produce a Markdown document with exactly these sections in this order:

## Week summary
One paragraph. Volume of feedback (forms vs partner notes vs transcripts), modules covered, overall sentiment, the single biggest signal of the week.

## Top themes
3 to 5 themes that cut across multiple feedback items. For each theme:
- A bold title (5 to 8 words)
- Affected modules
- Underlying item count
- One sentence describing the pattern
- One verbatim quote from the source material that captures it best

## Severity-ranked items
Every distinct item, deduplicated and merged across sources. Order by severity (Critical, High, Medium, Low). For each item:
- Severity tag
- Module tag
- Type tag (Bug / UX / DataQuality / Feature / Positive)
- Title (5 to 8 words, action-oriented)
- Description in your own words (1 to 2 sentences)
- Originating quote where it adds context
- Source breakdown (how many forms, how many partner notes, how many transcripts mention it)

## New account requests
Accounts a BDR has explicitly asked to have ADDED to the platform that are not already there. Be strict about what qualifies:

INCLUDE an account only when:
- It comes from the Forms field "Is there a new account you would like to see the outputs for?" (this is an explicit, unambiguous request), OR
- Someone in a transcript or note directly asks for an account to be built, loaded, or run that is not already in the platform (e.g. "can you add X", "we should run Y", "I'd like to see Z in here").

EXCLUDE:
- Accounts merely mentioned, discussed, or named as ones a BDR works or is targeting (for example a BDR introduction listing the accounts on their patch). Naming an account is not requesting it.
- Accounts the source says are already loaded, already run, already in the system, or where the ask is only to "confirm visibility" of an existing account. Those are confirmation tasks, not new-account requests, and belong in the backlog if anywhere, not here.
- Accounts where the context is about sales strategy for that account rather than a request to add it to the platform.

If in doubt, leave it out. A false new-account card creates a spurious "load this account" task for an account that may already exist. For each genuine request:
- Account name
- Who asked (BDR first names)
- Source context (what they said, and why it is a new ask rather than an existing account)

## Backlog cards (comprehensive)
EVERY distinct, deduplicated action item across all sources. This is not a top-10 list. Each item becomes a Planner card, so completeness matters more than brevity here.

Important: if a source contains an explicit action-item matrix or table with its own IDs (for example an integrated feedback report with rows numbered 1.1, 1.2, 4.3, and so on, each carrying Owner, Priority, and Phase), preserve those IDs verbatim and extract every single row as its own card. Do not collapse or summarise a structured matrix down to highlights. Merge an item from the matrix with the same item appearing in a transcript or form (same underlying ask) into one card, keeping the matrix ID and noting the reinforcement in the source counts.

For items that come from unstructured sources (transcripts, partner prose, forms) and have no pre-assigned ID, generate one in the form NEW-1, NEW-2, and so on.

For each card list: ID, Title, Module, Owner, Priority, Phase, Effort (S/M/L), Why it matters.

## Next sprint recommendation
3 to 5 cards from the backlog above that should be prioritised for the next sprint, with brief reasoning per card. Be opinionated. The delivery lead needs a crisp call, not a balanced essay. This is the human-reading distillation; the backlog itself stays comprehensive.

## Close-the-loop note
A short paragraph the team could paste into a monthly "what we heard, what we shipped" email to the BDRs. Warm, specific, names the BDRs whose feedback shaped the sprint.

# After the Markdown

After the Markdown digest, emit TWO separate JSON blocks.

First, a backlog block delimited by `<backlog_json>` and `</backlog_json>` containing a JSON array of ALL backlog cards (comprehensive, every distinct item) in this schema:

{
  "id": "string (matrix ID like 4.3 if the source provided one, else NEW-1, NEW-2, ...)",
  "title": "string (action-oriented, ready for Planner)",
  "module": "BCA | BCI | EP | RBP | APPM | EngagementScoring | BuyingStage | BuyingGroup | Cross",
  "owner": "string (person named as owner in the source, else null)",
  "priority": "Critical | High | Medium | Low | Decision",
  "phase": "string (the Phase value from the source matrix if present, e.g. Today, Now, This week, Sprint 1, Phase 2; else infer from priority)",
  "type": "Bug | UX | DataQuality | Feature | Positive",
  "effort": "S | M | L",
  "why_it_matters": "string",
  "originating_quote": "string or null",
  "originating_bdrs": ["array of first names"],
  "source_counts": {"forms": int, "partner_notes": int, "transcripts": int},
  "fingerprint": "the matrix ID if present (e.g. 4.3), else lowercase-hyphenated-title"
}

For the fingerprint: when an item has a stable matrix ID, use that ID as the fingerprint so the same item never duplicates across runs even if its wording shifts slightly. Only fall back to the hyphenated title for items with no ID.

Then an account-requests block delimited by `<account_requests_json>` and `</account_requests_json>` containing a JSON array of new-account asks in this schema:

{
  "account": "string",
  "requested_by": ["array of first names"],
  "context": "one-sentence summary of what they want from this account or why",
  "fingerprint": "lowercase-hyphenated-account-name"
}

The fingerprint fields exist so the orchestrator can dedupe across runs. Use the same fingerprint format consistently so the same item in two consecutive weeks resolves to the same string.

# Tone and style

Be direct. Use UK English. No em dashes. Bold only for headings and item titles. Prioritise specificity over diplomacy. If a piece of feedback contradicts another, surface that explicitly rather than averaging it out. If something is a non-issue or a one-off, say so and exclude it.
"""

# ---------------------------------------------------------------------------
# Fingerprint normalisation (deduplication robustness)
# ---------------------------------------------------------------------------

# Common English stop words to strip from title-based fingerprints so that
# "fix account search filter" and "fix the account search filter" resolve to
# the same key.
_STOP_WORDS = frozenset({
    "a", "an", "the", "to", "in", "of", "for", "with", "and", "or",
    "is", "are", "be", "by", "on", "at", "from", "as", "into", "it",
    "its", "this", "that", "not", "no", "so", "we", "i", "do",
})

# Corporate suffixes to strip from account-name fingerprints so that
# "Salesforce", "Salesforce Inc", and "Salesforce Corporation" all
# resolve to the same key.
_CORPORATE_SUFFIXES = frozenset({
    "inc", "corp", "corporation", "ltd", "llc", "llp", "plc", "co",
    "company", "group", "holdings", "international", "global",
    "solutions", "services", "technologies", "technology", "tech",
    "limited", "partners", "partnership",
})


def normalize_fp(fp: str) -> str:
    """Return a canonical form of a fingerprint for deduplication.

    Steps:
    1. Lowercase and replace every non-alphanumeric run with a space.
    2. Drop tokens that are stop words or corporate suffixes.
    3. Rejoin with hyphens.

    This makes the deduplication tolerant of minor title rewording
    ("fix-account-filter" vs "fix-the-account-filter") and account-name
    variants ("salesforce" vs "salesforce-inc").
    """
    if not fp:
        return fp or ""
    tokens = re.sub(r"[^a-z0-9]+", " ", fp.lower()).split()
    tokens = [t for t in tokens if t not in _STOP_WORDS and t not in _CORPORATE_SUFFIXES]
    return "-".join(tokens)


# ---------------------------------------------------------------------------
# State (idempotency)
# ---------------------------------------------------------------------------


def _migrate_registry(registry: dict) -> dict:
    """Re-key an existing fingerprint registry using normalized keys.

    Needed once after deploying the normalization fix so that cards already
    recorded in .state.json are found under their new canonical keys.
    """
    return {normalize_fp(k): v for k, v in registry.items()}


def load_state() -> dict:
    if STATE_PATH.exists():
        try:
            state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
            # Migrate any pre-existing fingerprint registries to normalized form
            # so that old entries still match against new normalized lookups.
            for key in ("backlog_fingerprints", "account_fingerprints"):
                if key in state:
                    state[key] = _migrate_registry(state[key])
            return state
        except json.JSONDecodeError:
            print(f"warning: {STATE_PATH} corrupt, starting fresh", file=sys.stderr)
    return {"files": {}, "backlog_fingerprints": {}, "account_fingerprints": {}}


def save_state(state: dict) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def is_unchanged(path: Path, state: dict) -> bool:
    key = str(path.resolve())
    entry = state["files"].get(key)
    if not entry:
        return False
    return entry.get("hash") == file_hash(path)


def mark_processed(path: Path, state: dict) -> None:
    key = str(path.resolve())
    state["files"][key] = {
        "hash": file_hash(path),
        "processed_at": datetime.now().isoformat(timespec="seconds"),
    }


# ---------------------------------------------------------------------------
# Format-specific readers
# ---------------------------------------------------------------------------


def read_docx(path: Path) -> str:
    """Extract paragraphs and tables from a .docx, preserving order and basic structure."""
    doc = Document(path)
    parts = []
    for el in doc.element.body.iter():
        tag = el.tag.split("}")[-1]
        if tag == "p":
            # Each paragraph: join all runs' text
            runs = el.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
            text = "".join(r.text or "" for r in runs).strip()
            if text:
                parts.append(text)
        elif tag == "tbl":
            # Render tables as tab-separated rows
            for row in el.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr"):
                cells = []
                for cell in row.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc"):
                    runs = cell.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
                    cells.append("".join(r.text or "" for r in runs).strip())
                if any(cells):
                    parts.append("\t".join(cells))
    return "\n".join(parts)


VTT_TIMESTAMP_LINE = re.compile(r"^\d{2}:\d{2}:\d{2}\.\d{3}\s+-->\s+\d{2}:\d{2}:\d{2}\.\d{3}.*$")
VTT_SPEAKER_OPEN = re.compile(r"<v\s+([^>]+)>")
VTT_SPEAKER_CLOSE = re.compile(r"</v>")
VTT_CUE_ID = re.compile(r"^[a-f0-9-]{8,}/.*$|^\d+$")


def read_vtt(path: Path) -> str:
    """Strip WEBVTT cue IDs and timestamps; collapse speaker tags to 'Name: text'."""
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    out = []
    current_speaker = None
    for line in lines:
        line = line.strip()
        if not line or line == "WEBVTT" or VTT_TIMESTAMP_LINE.match(line) or VTT_CUE_ID.match(line):
            continue
        # Detect speaker tag
        m = VTT_SPEAKER_OPEN.search(line)
        if m:
            current_speaker = m.group(1).strip()
            line = VTT_SPEAKER_OPEN.sub("", line)
        line = VTT_SPEAKER_CLOSE.sub("", line).strip()
        if not line:
            continue
        if current_speaker and (not out or not out[-1].startswith(f"{current_speaker}:")):
            out.append(f"{current_speaker}: {line}")
        elif out and out[-1].split(":", 1)[0] == current_speaker:
            out[-1] += " " + line
        else:
            out.append(line)
    return "\n".join(out)


def read_any(path: Path) -> str | None:
    """Dispatch by extension. Returns None if unsupported."""
    ext = path.suffix.lower()
    try:
        if ext in {".md", ".txt"}:
            return path.read_text(encoding="utf-8", errors="ignore")
        if ext == ".vtt":
            return read_vtt(path)
        if ext == ".docx":
            return read_docx(path)
    except Exception as e:
        print(f"warning: failed to read {path.name}: {e}", file=sys.stderr)
        return None
    return None


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------


def get_col(row: dict, partial: str) -> str:
    needle = partial.lower()
    for key, val in row.items():
        if needle in str(key).lower():
            return str(val) if val is not None else ""
    return ""


def load_forms_responses(path: Path) -> list[dict]:
    if not path.exists():
        print(f"warning: Forms Excel not found at\n  {path}", file=sys.stderr)
        return []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except (InvalidFileException, PermissionError) as e:
        print(f"warning: could not open Excel: {e}", file=sys.stderr)
        return []
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if len(rows) < 2:
        return []
    headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[0])]
    return [
        {headers[i]: row[i] for i in range(len(headers))}
        for row in rows[1:]
        if any(row)
    ]


def load_directory(dir_path: Path, state: dict) -> tuple[list[dict], list[Path]]:
    """Return (new_or_changed_files, all_files). Skip files unchanged since last run."""
    if not dir_path.exists():
        return [], []
    new_items: list[dict] = []
    all_paths: list[Path] = []
    for p in sorted(dir_path.iterdir()):
        if not p.is_file() or p.name.startswith("."):
            continue
        all_paths.append(p)
        if is_unchanged(p, state):
            continue
        content = read_any(p)
        if content is None:
            continue
        new_items.append({"filename": p.name, "content": content, "path": p})
    return new_items, all_paths


# ---------------------------------------------------------------------------
# Bundle builder
# ---------------------------------------------------------------------------


def format_form_response(r: dict, idx: int) -> str:
    parts = [
        f"### Forms response #{idx}",
        f"- Module: {get_col(r, FORM_COL_MODULE)}",
        f"- Type: {get_col(r, FORM_COL_TYPE)}",
        f"- Priority: {get_col(r, FORM_COL_PRIORITY)}",
        f"- Account (feedback is about): {get_col(r, FORM_COL_ACCOUNT_REL) or '(not provided)'}",
    ]
    new_acct = get_col(r, FORM_COL_ACCOUNT_NEW)
    if new_acct and new_acct.lower() not in {"no", "n/a", "none", "-"}:
        parts.append(f"- New account requested: {new_acct}")
    parts.extend([
        f"- BDR: {get_col(r, FORM_COL_NAME) or '(anonymous)'}",
        f"- Submitted: {get_col(r, FORM_COL_COMPLETED)}",
        "",
        get_col(r, FORM_COL_DESCRIBE),
    ])
    return "\n".join(parts)


def build_input_bundle(forms: list[dict], notes: list[dict], transcripts: list[dict]) -> str:
    parts = [f"# Feedback for week {WEEK_LABEL}\n"]

    parts.append(f"\n## Microsoft Forms submissions ({len(forms)})\n")
    if forms:
        for i, r in enumerate(forms, 1):
            parts.append(format_form_response(r, i))
    else:
        parts.append("_(none this week)_\n")

    parts.append(f"\n## Partner session notes ({len(notes)})\n")
    if notes:
        for n in notes:
            parts.append(f"### {n['filename']}\n\n{n['content']}\n")
    else:
        parts.append("_(none new this week)_\n")

    parts.append(f"\n## Meeting transcripts ({len(transcripts)})\n")
    if transcripts:
        for t in transcripts:
            parts.append(f"### {t['filename']}\n\n{t['content']}\n")
    else:
        parts.append("_(none new this week)_\n")

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Claude call
# ---------------------------------------------------------------------------


def synthesise(bundle: str) -> str:
    client = Anthropic()
    chunks: list[str] = []
    with client.messages.stream(
        model=MODEL,
        max_tokens=32000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": bundle}],
    ) as stream:
        for text in stream.text_stream:
            chunks.append(text)
    return "".join(chunks)


# ---------------------------------------------------------------------------
# Output handling
# ---------------------------------------------------------------------------

BACKLOG_BLOCK = re.compile(r"<backlog_json>(.*?)</backlog_json>", re.DOTALL)
ACCOUNTS_BLOCK = re.compile(r"<account_requests_json>(.*?)</account_requests_json>", re.DOTALL)


def split_output(output: str) -> tuple[str, list[dict], list[dict]]:
    backlog = []
    accounts = []
    m_backlog = BACKLOG_BLOCK.search(output)
    if m_backlog:
        try:
            backlog = json.loads(m_backlog.group(1).strip())
        except json.JSONDecodeError as e:
            print(f"warning: could not parse backlog JSON: {e}", file=sys.stderr)
    m_accounts = ACCOUNTS_BLOCK.search(output)
    if m_accounts:
        try:
            accounts = json.loads(m_accounts.group(1).strip())
        except json.JSONDecodeError as e:
            print(f"warning: could not parse account requests JSON: {e}", file=sys.stderr)
    digest = BACKLOG_BLOCK.sub("", output)
    digest = ACCOUNTS_BLOCK.sub("", digest).strip()
    return digest, backlog, accounts


def dedupe_against_state(
    items: list[dict], registry_key: str, state: dict
) -> tuple[list[dict], list[dict]]:
    """Split items into (new_for_planner, already_seen). Update state in place for new ones.

    Fingerprints are normalised before comparison so that minor wording
    changes in titles ("fix-account-filter" vs "fix-the-account-filter")
    and account-name variants ("salesforce" vs "salesforce-inc") do not
    produce duplicate Planner cards.
    """
    registry = state.setdefault(registry_key, {})
    new_items, already_seen = [], []
    for it in items:
        fp = it.get("fingerprint")
        if not fp:
            new_items.append(it)
            continue
        norm_fp = normalize_fp(fp)
        if norm_fp in registry:
            already_seen.append(it)
            continue
        new_items.append(it)
        registry[norm_fp] = {
            "first_seen_week": WEEK_LABEL,
            "title": it.get("title") or it.get("account"),
        }
    return new_items, already_seen


def write_outputs(
    digest_md: str,
    backlog_new: list[dict],
    accounts_new: list[dict],
) -> tuple[Path, Path, Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    digest_path = OUTPUT_DIR / f"digest_{WEEK_LABEL}.md"
    backlog_path = OUTPUT_DIR / f"backlog_{WEEK_LABEL}.json"
    accounts_path = OUTPUT_DIR / f"account_requests_{WEEK_LABEL}.json"
    digest_path.write_text(digest_md, encoding="utf-8")
    backlog_path.write_text(json.dumps(backlog_new, indent=2), encoding="utf-8")
    accounts_path.write_text(json.dumps(accounts_new, indent=2), encoding="utf-8")
    return digest_path, backlog_path, accounts_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("error: ANTHROPIC_API_KEY is not set.", file=sys.stderr)
        print("  Copy .env.example to .env and add your key.", file=sys.stderr)
        sys.exit(1)

    state = load_state()

    forms = load_forms_responses(FORMS_XLSX_PATH)
    notes_new, notes_all = load_directory(PARTNER_NOTES_DIR, state)
    transcripts_new, transcripts_all = load_directory(TRANSCRIPTS_DIR, state)

    if not (forms or notes_new or transcripts_new):
        print("nothing new to process. Forms responses present, but all partner notes and transcripts are unchanged since the last run.")
        if forms:
            print("(rerunning anyway because there may be a new form response)")
        else:
            sys.exit(0)

    print(
        f"loaded {len(forms)} form responses, "
        f"{len(notes_new)} new/changed partner notes ({len(notes_all)} total), "
        f"{len(transcripts_new)} new/changed transcripts ({len(transcripts_all)} total). "
        "calling Claude..."
    )

    bundle = build_input_bundle(forms, notes_new, transcripts_new)
    output = synthesise(bundle)
    digest, backlog, accounts = split_output(output)

    backlog_new, backlog_seen = dedupe_against_state(backlog, "backlog_fingerprints", state)
    accounts_new, accounts_seen = dedupe_against_state(accounts, "account_fingerprints", state)

    # Mark all newly read files as processed
    for n in notes_new:
        mark_processed(n["path"], state)
    for t in transcripts_new:
        mark_processed(t["path"], state)

    digest_path, backlog_path, accounts_path = write_outputs(digest, backlog_new, accounts_new)
    save_state(state)

    print(f"digest:           {digest_path}")
    print(f"backlog cards:    {backlog_path}  ({len(backlog_new)} new, {len(backlog_seen)} skipped as duplicate)")
    print(f"account requests: {accounts_path}  ({len(accounts_new)} new, {len(accounts_seen)} skipped as duplicate)")


if __name__ == "__main__":
    main()
