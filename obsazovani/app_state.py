from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .core import MAX_EPISODES, build_project, loose_match_key, normalize_actor, normalize_character, normalize_text
from .i18n import t
from .exporter import export_project_workbook
from .project_store import (
    BulkImportSource,
    deduplicate_episode_labels,
    empty_project_payload,
    load_project_file,
    make_episode_payload,
    next_episode_label,
    read_episode_source,
    save_project_file,
)


@dataclass(frozen=True)
class BulkImportPlanEntry:
    source_name: str
    source_label: str
    target_index: int
    target_label: str
    overwrites_content: bool
    creates_episode: bool
    content: str
    assignments: dict[str, dict[str, str]]


class AppState:
    def __init__(self) -> None:
        self._payload: dict[str, Any] = empty_project_payload()
        self._analysis: dict[str, Any] = build_project(self._payload)
        self.project_path: Path | None = None
        self.dirty = False

    @property
    def payload(self) -> dict[str, Any]:
        return deepcopy(self._payload)

    @property
    def analysis(self) -> dict[str, Any]:
        return deepcopy(self._analysis)

    @property
    def title(self) -> str:
        return str(self._payload.get("title", ""))

    @property
    def episode_count(self) -> int:
        return len(self._payload.get("episodes", []))

    @property
    def herci_by_episode_export(self) -> bool:
        export_options = self._payload.get("exportOptions", {})
        if not isinstance(export_options, dict):
            return False
        return bool(export_options.get("herciByEpisode", False))

    def reset(self) -> dict[str, Any]:
        self._payload = empty_project_payload()
        self._analysis = build_project(self._payload)
        self.project_path = None
        self.dirty = False
        return self.analysis

    def load_project(self, path: Path) -> dict[str, Any]:
        payload = load_project_file(path)
        self._payload = payload
        self._analysis = build_project(payload)
        self.project_path = path
        self.dirty = False
        return self.analysis

    def save_project(self, path: Path | None = None) -> Path:
        target = Path(path) if path is not None else self.project_path
        if target is None:
            raise ValueError(t("state.error.no_path"))
        save_project_file(target, self._payload)
        self.project_path = target
        self.dirty = False
        return target

    def set_title(self, title: str) -> None:
        self._payload["title"] = title
        self.dirty = True

    def set_herci_by_episode_export(self, enabled: bool) -> None:
        export_options = self._payload.setdefault("exportOptions", {})
        normalized = bool(enabled)
        if bool(export_options.get("herciByEpisode", False)) == normalized:
            return
        export_options["herciByEpisode"] = normalized
        self.dirty = True

    def episode_label(self, episode_index: int) -> str:
        return str(self._payload["episodes"][episode_index].get("label", ""))

    def set_episode_content(self, episode_index: int, content: str) -> None:
        self._payload["episodes"][episode_index]["content"] = content
        self.dirty = True

    def import_episode_file(self, episode_index: int, path: Path, sheet_name: str | None = None) -> str:
        source = read_episode_source(path, sheet_name=sheet_name)
        self.set_episode_content(episode_index, source.content)
        self._merge_imported_assignments(source.assignments)
        return source.content

    def clear_episode(self, episode_index: int) -> None:
        self.set_episode_content(episode_index, "")

    def preview_bulk_import(
        self,
        start_episode_index: int,
        sources: list[BulkImportSource],
    ) -> list[BulkImportPlanEntry]:
        if not sources:
            raise ValueError(t("state.error.no_sources"))
        if start_episode_index < 0:
            raise ValueError(t("state.error.bad_position"))

        end_index = start_episode_index + len(sources)
        if end_index > MAX_EPISODES:
            available = MAX_EPISODES - start_episode_index
            raise ValueError(t("state.error.no_slots", available=available, count=len(sources)))

        episodes = list(self._payload.get("episodes", []))
        reserved_labels = [
            str(episode.get("label", ""))
            for index, episode in enumerate(episodes)
            if index < start_episode_index or index >= end_index
        ]
        target_labels = deduplicate_episode_labels(
            [source.label for source in sources],
            reserved_labels=reserved_labels,
            start_index=start_episode_index,
        )

        plan: list[BulkImportPlanEntry] = []
        for offset, source in enumerate(sources):
            target_index = start_episode_index + offset
            existing = episodes[target_index] if target_index < len(episodes) else None
            plan.append(
                BulkImportPlanEntry(
                    source_name=source.source_name,
                    source_label=source.label,
                    target_index=target_index,
                    target_label=target_labels[offset],
                    overwrites_content=bool(existing and str(existing.get("content", "")).strip()),
                    creates_episode=existing is None,
                    content=source.content,
                    assignments=dict(source.assignments),
                )
            )
        return plan

    def apply_bulk_import(self, plan: list[BulkImportPlanEntry]) -> list[int]:
        if not plan:
            raise ValueError(t("state.error.no_plan"))

        episodes = self._payload.setdefault("episodes", [])
        highest_target = max(entry.target_index for entry in plan)
        while len(episodes) <= highest_target:
            episodes.append(make_episode_payload(len(episodes), next_episode_label(episodes)))

        for entry in plan:
            episodes[entry.target_index]["label"] = entry.target_label
            episodes[entry.target_index]["content"] = entry.content
            self._merge_imported_assignments(entry.assignments)

        self.dirty = True
        return [entry.target_index for entry in plan]

    def add_episode(self) -> int:
        episodes = self._payload.setdefault("episodes", [])
        if len(episodes) >= MAX_EPISODES:
            raise ValueError(t("state.error.max_episodes", max=MAX_EPISODES))

        episodes.append(make_episode_payload(len(episodes), next_episode_label(episodes)))
        self.dirty = True
        return len(episodes) - 1

    def rename_episode(self, episode_index: int, label: str) -> None:
        episodes = self._payload.get("episodes", [])
        if not 0 <= episode_index < len(episodes):
            raise IndexError(t("state.error.episode_not_found"))

        normalized_label = normalize_text(label)
        if not normalized_label:
            raise ValueError(t("state.error.empty_label"))

        for index, episode in enumerate(episodes):
            if index == episode_index:
                continue
            if normalize_text(str(episode.get("label", ""))) == normalized_label:
                raise ValueError(t("state.error.duplicate_label", label=normalized_label))

        if normalize_text(str(episodes[episode_index].get("label", ""))) == normalized_label:
            return
        episodes[episode_index]["label"] = normalized_label
        self.dirty = True

    def remove_episode(self, episode_index: int) -> None:
        episodes = self._payload.get("episodes", [])
        if len(episodes) <= 1:
            raise ValueError(t("state.error.min_episodes"))
        if not 0 <= episode_index < len(episodes):
            raise IndexError(t("state.error.episode_not_found"))

        episodes.pop(episode_index)
        self.dirty = True

    def set_assignment(self, character: str, field: str, value: str) -> None:
        assignments = self._payload.setdefault("assignments", {})
        assignment = assignments.setdefault(character, {"actor": "", "note": ""})
        assignment[field] = value
        if not assignment.get("actor") and not assignment.get("note"):
            assignments.pop(character, None)
        self.dirty = True

    def unify_actor_variants(self, variants: list[str], target_name: str) -> int:
        normalized_target = normalize_actor(target_name)
        if not normalized_target:
            raise ValueError(t("state.error.empty_actor"))

        variant_keys = {
            loose_match_key(normalize_actor(variant))
            for variant in variants
            if normalize_actor(variant)
        }
        variant_keys.discard("")
        if not variant_keys:
            raise ValueError(t("state.error.no_variants"))

        assignments = self._payload.setdefault("assignments", {})
        changed = 0
        for assignment in assignments.values():
            actor_value = normalize_actor(str(assignment.get("actor", "") or ""))
            if not actor_value:
                continue
            if loose_match_key(actor_value) not in variant_keys:
                continue
            if actor_value == normalized_target:
                continue
            assignment["actor"] = normalized_target
            changed += 1

        if changed:
            self.dirty = True
        return changed

    def _merge_imported_assignments(self, imported_assignments: dict[str, dict[str, str]]) -> None:
        if not imported_assignments:
            return

        assignments = self._payload.setdefault("assignments", {})
        for character, imported in imported_assignments.items():
            normalized_character = normalize_character(character)
            if not normalized_character:
                continue

            assignment = assignments.setdefault(normalized_character, {"actor": "", "note": ""})
            imported_actor = normalize_actor(str(imported.get("actor", "") or ""))
            imported_note = normalize_text(str(imported.get("note", "") or ""))

            if imported_actor and not normalize_actor(str(assignment.get("actor", "") or "")):
                assignment["actor"] = imported_actor
            if imported_note and not normalize_text(str(assignment.get("note", "") or "")):
                assignment["note"] = imported_note

            if not assignment.get("actor") and not assignment.get("note"):
                assignments.pop(normalized_character, None)

    def recompute(self) -> dict[str, Any]:
        analysis = build_project(self._payload)
        self._analysis = analysis
        return self.analysis

    def export_workbook(self) -> bytes:
        analysis = build_project(self._payload)
        self._analysis = analysis
        return export_project_workbook(analysis, herci_by_episode=self.herci_by_episode_export)
