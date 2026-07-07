from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def clean_text(value: Any, limit: int = 1000) -> str:
    text = "" if value is None else str(value)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:limit]


def safe_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace("%", "").replace(",", "").strip())
    except ValueError:
        return None


def normalize_token_text(value: Any) -> str:
    text = clean_text(value, 1000).lower()
    text = re.sub(r"[\s_\-:/\\|（）()\[\]【】,，;；.。]+", "", text)
    return text


def tokenize(value: Any) -> set[str]:
    text = clean_text(value, 4000).lower()
    tokens = re.findall(r"[a-z0-9]+|[\u4e00-\u9fff]+", text)
    stop_words = {"total", "table", "start", "end", "sample", "size", "series", "系列", "列"}
    return {token for token in tokens if len(token) > 1 and token not in stop_words}


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    ensure_dir(path.parent)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({key: row.get(key, "") for key in fieldnames})


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def unique_name(value: str, used: dict[str, int]) -> str:
    base = value
    if base not in used:
        used[base] = 1
        return base
    used[base] += 1
    return f"{base}_{used[base]}"


def slug(value: Any, limit: int = 90) -> str:
    text = clean_text(value, 500)
    text = re.sub(r'[\\/:*?"<>|\[\]()]', "_", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return (text or "table")[:limit]

