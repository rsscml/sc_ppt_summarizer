"""
Email History Module
====================
Persists accepted/finalized email summaries to disk as JSON files.
Supports retrieval of the latest accepted email for delta comparison.
"""

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


# Default storage directory (alongside the app)
_DEFAULT_DIR = Path(__file__).parent / "email_history"


def _ensure_dir(directory: Path):
    directory.mkdir(parents=True, exist_ok=True)


def _history_dir() -> Path:
    d = Path(os.getenv("EMAIL_HISTORY_DIR", str(_DEFAULT_DIR)))
    _ensure_dir(d)
    return d


def save_accepted_email(
    email_content: str,
    source_filename: str,
    session_id: str,
    section_summaries_text: str = "",
    metadata: dict[str, Any] | None = None,
) -> dict:
    """
    Save a finalized email summary to disk.
    Returns the stored record (including its ID and timestamp).
    """
    now = datetime.now(timezone.utc)
    record_id = now.strftime("%Y%m%dT%H%M%SZ") + f"_{session_id[:8]}"
    record = {
        "id": record_id,
        "accepted_at": now.isoformat(),
        "source_filename": source_filename,
        "session_id": session_id,
        "email_content": email_content,
        "section_summaries_text": section_summaries_text,
        "metadata": metadata or {},
    }

    filepath = _history_dir() / f"{record_id}.json"
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=2)

    return record


def get_latest_accepted_email() -> dict | None:
    """
    Return the most recently accepted email record, or None if no history exists.
    Files are named with ISO timestamps so lexicographic sort gives chronological order.
    """
    d = _history_dir()
    files = sorted(d.glob("*.json"), reverse=True)
    if not files:
        return None
    with open(files[0], "r", encoding="utf-8") as f:
        return json.load(f)


def list_accepted_emails(limit: int = 20) -> list[dict]:
    """
    Return a list of accepted email records (most recent first), without the
    full email_content/section_summaries_text to keep responses lightweight.
    """
    d = _history_dir()
    files = sorted(d.glob("*.json"), reverse=True)[:limit]
    results = []
    for fp in files:
        try:
            with open(fp, "r", encoding="utf-8") as f:
                rec = json.load(f)
            results.append({
                "id": rec["id"],
                "accepted_at": rec["accepted_at"],
                "source_filename": rec.get("source_filename", ""),
                "session_id": rec.get("session_id", ""),
                "preview": rec.get("email_content", "")[:200],
            })
        except Exception:
            continue
    return results


def get_accepted_email_by_id(record_id: str) -> dict | None:
    """Retrieve a specific accepted email record by its ID."""
    filepath = _history_dir() / f"{record_id}.json"
    if not filepath.exists():
        return None
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)
