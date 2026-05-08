from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def load_text(path: str | Path) -> str:
    return Path(path).read_text(encoding="utf-8")


def split_url_and_body(raw_text: str) -> tuple[str | None, str]:
    lines = raw_text.splitlines()
    if not lines:
        return None, ""
    first_line = lines[0].strip()
    if first_line.startswith("http://") or first_line.startswith("https://"):
        return first_line, "\n".join(lines[1:]).strip()
    return None, raw_text.strip()


def parse_json_document(raw_text: str) -> dict[str, Any]:
    _, body = split_url_and_body(raw_text)
    return json.loads(body)


def parse_multi_json_documents(raw_text: str) -> list[dict[str, Any]]:
    _, body = split_url_and_body(raw_text)
    decoder = json.JSONDecoder()
    items: list[dict[str, Any]] = []
    index = 0
    length = len(body)
    while index < length:
        while index < length and body[index].isspace():
            index += 1
        if index >= length:
            break
        item, next_index = decoder.raw_decode(body, index)
        items.append(item)
        index = next_index
    return items


def load_mitra_list_from_file(path: str | Path) -> list[dict[str, Any]]:
    payload = parse_json_document(load_text(path))
    return list(payload.get("mitras", []))


def load_mitra_detail_from_file(path: str | Path) -> dict[str, Any]:
    payload = parse_json_document(load_text(path))
    return dict(payload.get("mitra", {}))


def load_history_documents_from_file(path: str | Path) -> list[dict[str, Any]]:
    return parse_multi_json_documents(load_text(path))
