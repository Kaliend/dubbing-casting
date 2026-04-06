from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping

from .core import MAX_EPISODES, normalize_text
from .i18n import t
from .importers import (
    ImportableWorkbookSheet,
    ImportedEpisodeSource,
    import_episode_source,
    list_importable_xlsx_sheets,
)

PROJECT_FORMAT_VERSION = 2
DEFAULT_PROJECT_TITLE = "Obsazení projektu"
DEFAULT_EPISODE_COUNT = 1
EPISODE_LABELS = tuple(f"{index + 1:02d}" for index in range(MAX_EPISODES))
DEFAULT_EXPORT_OPTIONS = {"herciByEpisode": False}
SUPPORTED_IMPORT_SUFFIXES = {".txt", ".tsv", ".csv", ".xlsx", ".doc", ".docx"}


@dataclass(frozen=True)
class BulkImportSource:
    source_name: str
    label: str
    content: str
    assignments: dict[str, dict[str, str]] = field(default_factory=dict)


def clamp_episode_count(value: Any, fallback: int = DEFAULT_EPISODE_COUNT) -> int:
    try:
        count = int(value)
    except (TypeError, ValueError):
        count = fallback
    return max(1, min(MAX_EPISODES, count))


def default_episode_label(index: int) -> str:
    safe_index = max(0, min(index, MAX_EPISODES - 1))
    return EPISODE_LABELS[safe_index]


def normalize_episode_label(value: Any, index: int) -> str:
    label = normalize_text(str(value or ""))
    return label or default_episode_label(index)


def make_episode_payload(index: int, label: Any = None, content: Any = "") -> dict[str, str]:
    return {
        "label": normalize_episode_label(label, index),
        "content": str(content or ""),
    }


def next_episode_label(episodes: list[Mapping[str, Any]]) -> str:
    used_labels = {
        normalize_text(str(episode.get("label", "")))
        for episode in episodes
        if isinstance(episode, Mapping)
    }
    for label in EPISODE_LABELS:
        if label not in used_labels:
            return label
    return default_episode_label(len(episodes))


def deduplicate_episode_labels(
    labels: list[str],
    reserved_labels: list[str] | None = None,
    start_index: int = 0,
) -> list[str]:
    seen = {
        normalize_text(label).casefold()
        for label in (reserved_labels or [])
        if normalize_text(label)
    }
    resolved: list[str] = []

    for offset, raw_label in enumerate(labels):
        base_label = normalize_text(raw_label) or default_episode_label(start_index + offset)
        candidate = base_label
        attempt = 2
        while candidate.casefold() in seen:
            candidate = f"{base_label} ({attempt})"
            attempt += 1
        seen.add(candidate.casefold())
        resolved.append(candidate)

    return resolved


def _label_from_file_name(path: Path) -> str:
    normalized = normalize_text(re.sub(r"_+", " ", path.stem))
    return normalized or path.stem or path.name


def _bulk_import_source_from_imported(path: Path, imported: ImportedEpisodeSource, label: str | None = None) -> BulkImportSource:
    return BulkImportSource(
        source_name=path.name,
        label=normalize_text(label or _label_from_file_name(path)),
        content=imported.content,
        assignments=dict(imported.assignments),
    )


def read_bulk_import_sources_from_workbook(path: Path) -> list[BulkImportSource]:
    if path.suffix.lower() != ".xlsx":
        raise ValueError(t("store.error.workbook_not_xlsx"))

    options = list_importable_xlsx_sheets(path)
    return [
        BulkImportSource(
            source_name=f"{path.name} · {option.sheet_name}",
            label=normalize_text(option.sheet_name) or _label_from_file_name(path),
            content=option.content,
            assignments=dict(option.assignments),
        )
        for option in options
    ]


def read_bulk_import_sources_from_files(paths: list[Path]) -> list[BulkImportSource]:
    if not paths:
        raise ValueError(t("store.error.no_files"))

    sources: list[BulkImportSource] = []
    for path in sorted(paths, key=lambda item: item.name.casefold()):
        suffix = path.suffix.lower()
        if suffix not in SUPPORTED_IMPORT_SUFFIXES:
            continue

        if suffix == ".xlsx":
            options = list_importable_xlsx_sheets(path)
            if len(options) != 1:
                raise ValueError(t("store.error.multi_sheet", name=path.name))
            option = options[0]
            sources.append(
                BulkImportSource(
                    source_name=path.name,
                    label=_label_from_file_name(path),
                    content=option.content,
                    assignments=dict(option.assignments),
                )
            )
            continue

        imported = import_episode_source(path)
        sources.append(_bulk_import_source_from_imported(path, imported))

    if not sources:
        raise ValueError(t("store.error.no_inputs"))
    return sources


def read_bulk_import_sources_from_directory(path: Path) -> list[BulkImportSource]:
    if not path.is_dir():
        raise ValueError(t("store.error.not_dir", name=path.name))

    files = [
        child
        for child in sorted(path.iterdir(), key=lambda item: item.name.casefold())
        if child.is_file() and not child.name.startswith(".") and child.suffix.lower() in SUPPORTED_IMPORT_SUFFIXES
    ]
    if not files:
        raise ValueError(t("store.error.empty_dir"))
    return read_bulk_import_sources_from_files(files)


def empty_project_payload(count: int = DEFAULT_EPISODE_COUNT) -> dict[str, Any]:
    episode_count = clamp_episode_count(count)
    return {
        "title": t("project.default_title"),
        "episodes": [make_episode_payload(index) for index in range(episode_count)],
        "assignments": {},
        "exportOptions": dict(DEFAULT_EXPORT_OPTIONS),
    }


def normalize_project_payload(payload: Mapping[str, Any] | None) -> dict[str, Any]:
    base = empty_project_payload()
    if not payload:
        return base

    title = payload.get("title")
    if isinstance(title, str) and title.strip():
        base["title"] = title

    episodes = payload.get("episodes")
    if isinstance(episodes, list):
        normalized_episodes = []
        for index, episode_payload in enumerate(episodes[:MAX_EPISODES]):
            episode_data = episode_payload if isinstance(episode_payload, Mapping) else {}
            normalized_episodes.append(
                make_episode_payload(
                    index,
                    episode_data.get("label"),
                    episode_data.get("content", ""),
                )
            )
        if normalized_episodes:
            base["episodes"] = normalized_episodes

    assignments = payload.get("assignments")
    if isinstance(assignments, Mapping):
        cleaned: dict[str, dict[str, str]] = {}
        for character, assignment in assignments.items():
            if not isinstance(character, str) or not character.strip():
                continue
            normalized_character = character.strip()
            if not isinstance(assignment, Mapping):
                assignment = {}
            cleaned[normalized_character] = {
                "actor": str(assignment.get("actor", "") or ""),
                "note": str(assignment.get("note", "") or ""),
            }
        base["assignments"] = cleaned

    export_options = payload.get("exportOptions")
    if isinstance(export_options, Mapping):
        base["exportOptions"] = {
            "herciByEpisode": bool(export_options.get("herciByEpisode", False)),
        }

    return base


def load_project_file(path: Path) -> dict[str, Any]:
    try:
        raw_payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(t("store.error.invalid_json", name=path.name)) from exc

    if not isinstance(raw_payload, Mapping):
        raise ValueError(t("store.error.invalid_json_structure"))

    return normalize_project_payload(raw_payload)


def save_project_file(path: Path, payload: Mapping[str, Any]) -> None:
    normalized = normalize_project_payload(payload)
    document = {
        "version": PROJECT_FORMAT_VERSION,
        "title": normalized["title"],
        "episodes": normalized["episodes"],
        "assignments": normalized["assignments"],
        "exportOptions": normalized["exportOptions"],
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(document, ensure_ascii=False, indent=2), encoding="utf-8")


def list_episode_source_options(path: Path) -> list[ImportableWorkbookSheet]:
    if path.suffix.lower() != ".xlsx":
        return []
    return list_importable_xlsx_sheets(path)


def read_episode_source(path: Path, sheet_name: str | None = None) -> ImportedEpisodeSource:
    return import_episode_source(path, sheet_name=sheet_name)
