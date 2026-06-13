from __future__ import annotations

from dataclasses import dataclass, field
import json
import re
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd


DEFAULT_TAG_WHITELIST = {
    "AU",
    "NZ",
    "JB",
    "HN",
    "HNNZ",
    "NLG",
    "JBNZ",
    "HC",
    "TGG",
    "BL",
    "DJS",
    "BHLT",
    "BSR",
    "BR",
    "RETRAVISION",
}

DISPLAY_STATUS_WORDS = re.compile(r"^\s*(?:on display|not on display)\b", re.IGNORECASE)


@dataclass(frozen=True)
class NormalizedIdentifier:
    raw: str
    key: str
    stripped_tags: tuple[str, ...] = ()
    unknown_bracket_tags: tuple[str, ...] = ()


@dataclass(frozen=True)
class CanonicalEntry:
    domain: str
    canonical: str
    aliases: tuple[str, ...] = ()
    metadata: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class MatchItem:
    domain: str
    raw_value: str
    normalized_key: str
    status: str
    canonical: str = ""
    reason: str = ""
    candidates: tuple[str, ...] = ()
    confusable_with: tuple[str, ...] = ()
    differing_token: str = ""
    similarity: int = 0
    stripped_tags: tuple[str, ...] = ()
    unknown_bracket_tags: tuple[str, ...] = ()

    def as_dict(self) -> dict[str, object]:
        return {
            "domain": self.domain,
            "raw_value": self.raw_value,
            "normalized_key": self.normalized_key,
            "status": self.status,
            "canonical": self.canonical,
            "reason": self.reason,
            "candidates": list(self.candidates),
            "confusable_with": list(self.confusable_with),
            "differing_token": self.differing_token,
            "similarity": self.similarity,
            "stripped_tags": list(self.stripped_tags),
            "unknown_bracket_tags": list(self.unknown_bracket_tags),
        }


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def split_tag_text(tag_text: str) -> list[str]:
    return [part.strip().upper() for part in re.split(r"[,;/|]+", tag_text) if part.strip()]


def strip_whitelisted_trailing_tags(
    text: str,
    tag_whitelist: set[str],
) -> tuple[str, list[str], list[str]]:
    stripped_tags: list[str] = []
    unknown_tags: list[str] = []
    current = text.strip()
    bracket_pattern = re.compile(r"\s*[\(\[\{]([^()\[\]{}]+)[\)\]\}]\s*$")

    while True:
        match = bracket_pattern.search(current)
        if not match:
            break
        tags = split_tag_text(match.group(1))
        if tags and all(tag in tag_whitelist for tag in tags):
            stripped_tags.extend(tags)
            current = current[: match.start()].strip()
            continue
        if tags:
            unknown_tags.extend(tags)
        break

    return current, stripped_tags, unknown_tags


def normalize_identifier(
    value: object,
    tag_whitelist: Optional[Iterable[str]] = None,
) -> NormalizedIdentifier:
    whitelist = {tag.upper() for tag in (tag_whitelist or DEFAULT_TAG_WHITELIST)}
    raw = clean_text(value)
    stripped_text, stripped_tags, unknown_tags = strip_whitelisted_trailing_tags(raw, whitelist)
    normalized = stripped_text.casefold()
    normalized = re.sub(r"[-_./]+", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return NormalizedIdentifier(
        raw=raw,
        key=normalized,
        stripped_tags=tuple(stripped_tags),
        unknown_bracket_tags=tuple(unknown_tags),
    )


def edit_distance(left: str, right: str, limit: int = 3) -> int:
    if abs(len(left) - len(right)) > limit:
        return limit + 1
    previous = list(range(len(right) + 1))
    for i, left_char in enumerate(left, start=1):
        current = [i]
        row_min = i
        for j, right_char in enumerate(right, start=1):
            cost = 0 if left_char == right_char else 1
            value = min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + cost)
            current.append(value)
            row_min = min(row_min, value)
        if row_min > limit:
            return limit + 1
        previous = current
    return previous[-1]


def is_confusable(left_key: str, right_key: str) -> tuple[bool, int, str]:
    if not left_key or not right_key or left_key == right_key:
        return False, 0, ""
    distance = edit_distance(left_key, right_key, limit=2)
    prefix_or_substring = left_key in right_key or right_key in left_key
    if distance <= 2 or prefix_or_substring:
        return True, distance if distance <= 2 else 99, differing_token(left_key, right_key)
    return False, distance, ""


def differing_token(left_key: str, right_key: str) -> str:
    left_tokens = left_key.split()
    right_tokens = right_key.split()
    for left, right in zip(left_tokens, right_tokens):
        if left != right:
            return f"{left} <> {right}"
    if len(left_tokens) != len(right_tokens):
        return f"{' '.join(left_tokens)} <> {' '.join(right_tokens)}"
    return left_key


def normalize_decision(decision: dict[str, object]) -> dict[str, object]:
    domain = clean_text(decision.get("domain", "")).lower()
    raw_value = clean_text(decision.get("raw_value", ""))
    canonical = clean_text(decision.get("canonical", ""))
    action = clean_text(decision.get("action", "")).lower()
    return {
        **decision,
        "domain": domain,
        "raw_value": raw_value,
        "canonical": canonical,
        "action": action,
    }


def load_alias_decisions(path: Optional[Path]) -> list[dict[str, object]]:
    if not path or not path.exists():
        return []
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if isinstance(data, dict):
        data = data.get("decisions", [])
    if not isinstance(data, list):
        raise ValueError("Identifier alias map must be a list or {'decisions': [...]} JSON.")
    return [normalize_decision(item) for item in data if isinstance(item, dict)]


def save_alias_decisions(path: Path, decisions: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 1,
        "description": "Per-session identifier alias decisions for Calculator Online.",
        "decisions": [normalize_decision(item) for item in decisions],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def decision_lookup(
    decisions: list[dict[str, object]],
    domain: str,
) -> tuple[dict[str, str], set[str]]:
    mappings: dict[str, str] = {}
    skipped: set[str] = set()
    seen_raw_to_canonical: dict[str, str] = {}
    for decision in decisions:
        item = normalize_decision(decision)
        if item["domain"] != domain:
            continue
        raw_key = normalize_identifier(item["raw_value"]).key
        action = str(item["action"])
        if action in {"skip_always", "skip_this_run"}:
            skipped.add(raw_key)
            continue
        if action not in {"map", "add_new"}:
            continue
        canonical = str(item["canonical"])
        existing = seen_raw_to_canonical.get(raw_key)
        if existing and existing != canonical:
            raise ValueError(
                f"Collision: raw {item['raw_value']} maps to both {existing} and {canonical}."
            )
        seen_raw_to_canonical[raw_key] = canonical
        mappings[raw_key] = canonical
    return mappings, skipped


def build_resolution_items(
    domain: str,
    raw_values: Iterable[object],
    canonicals: Iterable[CanonicalEntry],
    decisions: Optional[list[dict[str, object]]] = None,
    tag_whitelist: Optional[Iterable[str]] = None,
) -> list[MatchItem]:
    decisions = decisions or []
    raw_mapping, skipped = decision_lookup(decisions, domain)

    canonical_by_norm: dict[str, list[str]] = {}
    normalized_canonicals: list[tuple[str, str]] = []
    for entry in canonicals:
        values = [entry.canonical, *entry.aliases]
        for value in values:
            normalized = normalize_identifier(value, tag_whitelist)
            if not normalized.key:
                continue
            canonical_by_norm.setdefault(normalized.key, [])
            if entry.canonical not in canonical_by_norm[normalized.key]:
                canonical_by_norm[normalized.key].append(entry.canonical)
        canonical_key = normalize_identifier(entry.canonical, tag_whitelist).key
        normalized_canonicals.append((entry.canonical, canonical_key))

    items: list[MatchItem] = []
    seen_raw_values: set[str] = set()
    for raw_value in raw_values:
        normalized = normalize_identifier(raw_value, tag_whitelist)
        if not normalized.key or normalized.raw in seen_raw_values:
            continue
        seen_raw_values.add(normalized.raw)

        if normalized.key in skipped:
            items.append(
                MatchItem(
                    domain=domain,
                    raw_value=normalized.raw,
                    normalized_key=normalized.key,
                    status="skipped",
                    reason="previous skip decision",
                    stripped_tags=normalized.stripped_tags,
                    unknown_bracket_tags=normalized.unknown_bracket_tags,
                )
            )
            continue

        if normalized.key in raw_mapping:
            items.append(
                MatchItem(
                    domain=domain,
                    raw_value=normalized.raw,
                    normalized_key=normalized.key,
                    status="auto_normalized",
                    canonical=raw_mapping[normalized.key],
                    reason="confirmed alias map",
                    stripped_tags=normalized.stripped_tags,
                    unknown_bracket_tags=normalized.unknown_bracket_tags,
                )
            )
            continue

        matches = canonical_by_norm.get(normalized.key, [])
        if len(matches) == 1:
            items.append(
                MatchItem(
                    domain=domain,
                    raw_value=normalized.raw,
                    normalized_key=normalized.key,
                    status="auto_normalized",
                    canonical=matches[0],
                    reason="exact equality after safe normalization",
                    stripped_tags=normalized.stripped_tags,
                    unknown_bracket_tags=normalized.unknown_bracket_tags,
                )
            )
            continue
        if len(matches) > 1:
            items.append(
                MatchItem(
                    domain=domain,
                    raw_value=normalized.raw,
                    normalized_key=normalized.key,
                    status="ambiguous",
                    candidates=tuple(matches),
                    reason="multiple canonicals share this normalized key",
                    stripped_tags=normalized.stripped_tags,
                    unknown_bracket_tags=normalized.unknown_bracket_tags,
                )
            )
            continue

        confusables = []
        best_similarity = 0
        best_diff = ""
        for canonical, canonical_key in normalized_canonicals:
            confusable, similarity, diff = is_confusable(normalized.key, canonical_key)
            if confusable:
                confusables.append(canonical)
                if not best_diff or similarity < best_similarity:
                    best_similarity = similarity
                    best_diff = diff
        if confusables:
            items.append(
                MatchItem(
                    domain=domain,
                    raw_value=normalized.raw,
                    normalized_key=normalized.key,
                    status="confusable",
                    confusable_with=tuple(confusables),
                    reason="near but not equal; likely different entity",
                    differing_token=best_diff,
                    similarity=best_similarity,
                    stripped_tags=normalized.stripped_tags,
                    unknown_bracket_tags=normalized.unknown_bracket_tags,
                )
            )
            continue

        items.append(
            MatchItem(
                domain=domain,
                raw_value=normalized.raw,
                normalized_key=normalized.key,
                status="unmatched",
                reason="no normalization match and no confirmed alias",
                stripped_tags=normalized.stripped_tags,
                unknown_bracket_tags=normalized.unknown_bracket_tags,
            )
        )

    return items


def summarize_items(items: list[MatchItem]) -> dict[str, int]:
    return {
        "auto_normalized": sum(item.status == "auto_normalized" for item in items),
        "needs_review": sum(item.status == "confusable" for item in items),
        "ambiguous": sum(item.status == "ambiguous" for item in items),
        "unmatched": sum(item.status == "unmatched" for item in items),
        "skipped": sum(item.status == "skipped" for item in items),
    }


def items_to_frame(items: list[MatchItem]) -> pd.DataFrame:
    return pd.DataFrame([item.as_dict() for item in items])


def display_like_columns(raw_df: pd.DataFrame) -> list[str]:
    rows: list[str] = []
    for column in raw_df.columns:
        values = raw_df[column].dropna().astype(str)
        if values.empty:
            continue
        if values.str.contains(DISPLAY_STATUS_WORDS, na=False).any():
            rows.append(str(column))
    return rows


def raw_values_from_frame(raw_df: pd.DataFrame, column: str) -> list[str]:
    if column not in raw_df.columns:
        return []
    return sorted({clean_text(value) for value in raw_df[column].dropna().tolist() if clean_text(value)})


def apply_sku_decisions_to_specs(
    sku_specs: list[dict[str, object]],
    decisions: list[dict[str, object]],
) -> list[dict[str, object]]:
    specs = [
        {
            **spec,
            "raw_columns": list(spec.get("raw_columns", [])),
        }
        for spec in sku_specs
    ]
    by_canonical = {
        normalize_identifier(spec.get("sku", "")).key: spec
        for spec in specs
        if clean_text(spec.get("sku", ""))
    }
    for decision in decisions:
        item = normalize_decision(decision)
        if item["domain"] != "sku" or item["action"] not in {"map", "add_new"}:
            continue
        canonical_key = normalize_identifier(item["canonical"]).key
        target = by_canonical.get(canonical_key)
        if target is None and item["action"] == "add_new":
            target = {
                "category": clean_text(item.get("category", "")),
                "sku": item["canonical"],
                "raw_columns": [],
            }
            specs.append(target)
            by_canonical[canonical_key] = target
        if target is None:
            continue
        raw_columns = target.setdefault("raw_columns", [])
        raw_value = item["raw_value"]
        if raw_value and raw_value not in raw_columns:
            raw_columns.append(raw_value)
    return specs


def account_decision_map(decisions: list[dict[str, object]]) -> dict[str, str]:
    result: dict[str, str] = {}
    for decision in decisions:
        item = normalize_decision(decision)
        if item["domain"] == "account" and item["action"] in {"map", "add_new"}:
            result[item["raw_value"]] = item["canonical"]
    return result


def country_decision_map(decisions: list[dict[str, object]]) -> dict[str, str]:
    result: dict[str, str] = {}
    for decision in decisions:
        item = normalize_decision(decision)
        if item["domain"] == "country" and item["action"] in {"map", "add_new"}:
            result[item["raw_value"]] = item["canonical"]
    return result
