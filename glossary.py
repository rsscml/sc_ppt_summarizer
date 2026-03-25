"""
Glossary Module
===============
Loads company-specific glossary JSON files and formats them for LLM context injection.

Supported JSON formats:
  1. Flat key-value:     {"BHV": "Bogen plant", "KAM": "Key Account Manager", ...}
  2. Categorized:        {"locations": {"BHV": "Bogen plant"}, "terms": {"KAM": "..."}}
  3. Array of objects:   [{"abbr": "BHV", "meaning": "Bogen plant", "category": "location"}, ...]

All formats are normalised into a unified internal structure:
  { "ABBR": {"meaning": "...", "category": "..."}, ... }
"""

import json
import os
from pathlib import Path
from typing import Any


# ─── Normalisation ────────────────────────────────────────────────────

def _normalise_flat(data: dict[str, str], category: str = "general") -> dict:
    """Flat key-value dict → unified format."""
    result = {}
    for abbr, meaning in data.items():
        if isinstance(meaning, str):
            result[abbr.strip()] = {"meaning": meaning.strip(), "category": category}
    return result


def _normalise_categorised(data: dict[str, dict]) -> dict:
    """Nested category → entries dict → unified format."""
    result = {}
    for category, entries in data.items():
        if isinstance(entries, dict):
            for abbr, meaning in entries.items():
                if isinstance(meaning, str):
                    result[abbr.strip()] = {"meaning": meaning.strip(), "category": category}
    return result


def _normalise_array(data: list[dict]) -> dict:
    """Array of objects → unified format. Detects common field names."""
    result = {}
    abbr_keys = ("abbr", "abbreviation", "key", "code", "short", "acronym", "symbol", "id", "name")
    meaning_keys = ("meaning", "full_name", "full", "description", "definition", "expansion",
                    "long", "value", "text", "explanation", "label")
    cat_keys = ("category", "cat", "type", "group", "domain", "section", "area")

    for item in data:
        if not isinstance(item, dict):
            continue
        item_lower = {k.lower().strip(): v for k, v in item.items()}

        abbr = None
        for k in abbr_keys:
            if k in item_lower and isinstance(item_lower[k], str):
                abbr = item_lower[k].strip()
                break

        meaning = None
        for k in meaning_keys:
            if k in item_lower and isinstance(item_lower[k], str):
                meaning = item_lower[k].strip()
                break

        category = "general"
        for k in cat_keys:
            if k in item_lower and isinstance(item_lower[k], str):
                category = item_lower[k].strip()
                break

        if abbr and meaning:
            result[abbr] = {"meaning": meaning, "category": category}

    return result


def normalise_json(data: Any, filename: str = "") -> dict:
    """Auto-detect JSON format and normalise to unified structure."""
    if isinstance(data, list):
        return _normalise_array(data)

    if isinstance(data, dict):
        # Check if it's categorised (all values are dicts) or flat (all values are strings)
        values = list(data.values())
        if not values:
            return {}

        all_str = all(isinstance(v, str) for v in values)
        all_dict = all(isinstance(v, dict) for v in values)

        if all_str:
            # Derive a category from filename if useful
            cat = Path(filename).stem if filename else "general"
            return _normalise_flat(data, category=cat)

        if all_dict:
            # Could be categorised OR array-of-objects-in-dict
            # Check if inner dicts look like entries (string values) or records (mixed)
            sample = values[0]
            if all(isinstance(v, str) for v in sample.values()):
                return _normalise_categorised(data)
            else:
                # Treat each top-level key as a category label with a nested dict
                return _normalise_categorised(data)

        # Mixed: try flat first (skip non-string values)
        return _normalise_flat({k: v for k, v in data.items() if isinstance(v, str)})

    return {}


# ─── Loading ──────────────────────────────────────────────────────────

def load_glossary_file(filepath: str) -> dict:
    """Load and normalise a single JSON glossary file."""
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Glossary file not found: {filepath}")

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    return normalise_json(data, filename=path.name)


def load_glossary_dir(directory: str) -> dict:
    """Load all .json files from a directory, merge into one unified glossary."""
    dir_path = Path(directory)
    if not dir_path.is_dir():
        return {}

    merged: dict = {}
    files_loaded = []

    for json_file in sorted(dir_path.glob("*.json")):
        try:
            entries = load_glossary_file(str(json_file))
            merged.update(entries)
            files_loaded.append({"file": json_file.name, "entries": len(entries)})
        except Exception as e:
            files_loaded.append({"file": json_file.name, "entries": 0, "error": str(e)})

    return {
        "entries": merged,
        "files_loaded": files_loaded,
        "total_entries": len(merged),
    }


# ─── Prompt Rendering ────────────────────────────────────────────────

def render_glossary_for_prompt(glossary_entries: dict, max_chars: int = 8000) -> str:
    """
    Render the glossary entries as a compact reference block for LLM system prompts.
    Groups by category for readability. Truncates if it would exceed max_chars.
    """
    if not glossary_entries:
        return ""

    # Group by category
    by_category: dict[str, list[tuple[str, str]]] = {}
    for abbr, info in sorted(glossary_entries.items()):
        cat = info.get("category", "general")
        by_category.setdefault(cat, []).append((abbr, info["meaning"]))

    lines = []
    lines.append("COMPANY GLOSSARY — Use this reference to correctly interpret abbreviations, "
                 "location codes, business entities, and domain-specific terms found in the presentation. "
                 "When you encounter any of these abbreviations, always use or mention the full meaning "
                 "in your output alongside the abbreviation on first use.")
    lines.append("")

    for cat in sorted(by_category.keys()):
        items = by_category[cat]
        lines.append(f"[{cat.upper()}]")
        for abbr, meaning in items:
            lines.append(f"  {abbr} = {meaning}")
        lines.append("")

    text = "\n".join(lines)

    # Truncate if too long (keep the instruction header)
    if len(text) > max_chars:
        header_end = text.index("\n\n") + 2
        text = text[:max_chars - 50] + "\n  ... [glossary truncated for token budget]"

    return text
