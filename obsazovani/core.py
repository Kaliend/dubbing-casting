from __future__ import annotations

import csv
import io
import math
import re
import unicodedata
from typing import Dict, Iterable, List, Tuple

WORD_RE = re.compile(r"\S+")
SPACE_RE = re.compile(r"\s+")
LOOSE_KEY_RE = re.compile(r"[^a-z0-9]+")
MAX_EPISODES = 6
IMPORTANT_UNASSIGNED_REPLICAS_WARNING = 50
GENERIC_BUCKET_INFO_REPLICAS = 20
HIGH_ACTOR_TOTAL_REPLICAS_INFO = 120
HIGH_ACTOR_TOTAL_INPUTS_INFO = 60
HIGH_ACTOR_TOTAL_REPLICAS_WARNING = 180
HIGH_ACTOR_TOTAL_INPUTS_WARNING = 90
HIGH_ACTOR_EPISODE_REPLICAS_INFO = 80
HIGH_ACTOR_EPISODE_INPUTS_INFO = 40
HIGH_ACTOR_EPISODE_REPLICAS_WARNING = 120
HIGH_ACTOR_EPISODE_INPUTS_WARNING = 60
MANY_ACTOR_CHARACTER_COUNT_INFO = 5
MANY_ACTOR_CHARACTER_MIN_REPLICAS_INFO = 40
MANY_ACTOR_EPISODE_CHARACTER_COUNT_INFO = 3
MANY_ACTOR_EPISODE_MIN_REPLICAS_INFO = 25
GENERIC_CHARACTER_KEYS = {
    "sbor",
    "hlasy",
    "hlas",
    "dav",
    "nikdo",
    "crowd",
    "voice",
    "voices",
}
VALIDATION_SEVERITY_ORDER = {"warning": 0, "info": 1}

HEADER_ALIASES = {
    "postava": "character",
    "postavy": "character",
    "character": "character",
    "char": "character",
    "tc": "timecode",
    "timecode": "timecode",
    "cas": "timecode",
    "čas": "timecode",
    "text": "text",
    "dialog": "text",
    "dialogue": "text",
    "replika": "text",
    "repliky": "replicas",
    "replicas": "replicas",
    "replica": "replicas",
    "vstupy": "inputs",
    "vstupy": "inputs",
    "inputs": "inputs",
    "input": "inputs",
    "poznamky": "note",
    "poznámky": "note",
    "poznamka": "note",
    "poznámka": "note",
    "daber": "actor",
    "dabér": "actor",
    "actor": "actor",
}


def strip_diacritics(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    return "".join(char for char in normalized if not unicodedata.combining(char))


def normalize_header(value: str) -> str:
    return SPACE_RE.sub(" ", strip_diacritics((value or "").replace("\ufeff", "")).strip().lower())


def normalize_text(value: str) -> str:
    return SPACE_RE.sub(" ", (value or "").strip())


def normalize_character(value: str) -> str:
    cleaned = normalize_text(value)
    return cleaned.rstrip(":")


def normalize_actor(value: str) -> str:
    return normalize_text(value)


def loose_match_key(value: str) -> str:
    normalized = normalize_text(strip_diacritics(value).lower())
    return LOOSE_KEY_RE.sub("", normalized)


def count_words(text: str) -> int:
    if not text:
        return 0
    return len(WORD_RE.findall(text.strip()))


def count_replicas(word_count: int) -> int:
    if word_count <= 0:
        return 0
    return int(math.ceil(word_count / 8))


def choose_delimiter(sample_lines: Iterable[str]) -> str:
    sample = "\n".join(sample_lines)
    candidates = ["\t", ";", ",", "|"]
    best_delimiter = "\t"
    best_score = (-1, -1)

    for delimiter in candidates:
        reader = csv.reader(io.StringIO(sample), delimiter=delimiter)
        rows = [row for row in reader if row]
        if not rows:
            continue
        header_hits = sum(1 for cell in rows[0] if HEADER_ALIASES.get(normalize_header(cell)))
        width = max((len(row) for row in rows), default=0)
        score = (header_hits, width)
        if score > best_score:
            best_score = score
            best_delimiter = delimiter

    return best_delimiter


def detect_mapping(first_row: List[str]) -> Tuple[Dict[int, str], bool]:
    mapping: Dict[int, str] = {}
    normalized = [normalize_header(cell) for cell in first_row]

    for index, header in enumerate(normalized):
        canonical = HEADER_ALIASES.get(header)
        if canonical and canonical not in mapping.values():
            mapping[index] = canonical

    has_header = "character" in mapping.values() and (
        "text" in mapping.values()
        or ("inputs" in mapping.values() and "replicas" in mapping.values())
    )
    if has_header:
        return mapping, True

    if len(first_row) >= 3 and first_row[1].strip().isdigit() and first_row[2].strip().isdigit():
        return {0: "character", 1: "inputs", 2: "replicas"}, False
    if len(first_row) >= 3:
        return {0: "character", 1: "timecode", 2: "text"}, False
    if len(first_row) == 2:
        return {0: "character", 1: "text"}, False
    if len(first_row) == 1:
        raise ValueError("Vstup musí obsahovat aspoň sloupec POSTAVA a TEXT nebo VSTUPY/REPLIKY.")
    return {}, False


def parse_rows(raw_text: str) -> Tuple[List[Dict[str, str]], str]:
    lines = [line for line in raw_text.splitlines() if line.strip()]
    if not lines:
        return [], "empty"

    delimiter = choose_delimiter(lines[:8])
    reader = csv.reader(io.StringIO(raw_text), delimiter=delimiter)
    table = [[cell.strip() for cell in row] for row in reader if any(cell.strip() for cell in row)]
    if not table:
        return [], "empty"

    mapping, has_header = detect_mapping(table[0])
    rows = table[1:] if has_header else table
    parsed_rows: List[Dict[str, str]] = []

    for raw_row in rows:
        if not raw_row:
            continue
        row: Dict[str, str] = {}
        for index, target in mapping.items():
            if target == "text" and not has_header and len(raw_row) > 3 and index == 2:
                row[target] = delimiter.join(raw_row[index:]).strip()
                break
            row[target] = raw_row[index].strip() if index < len(raw_row) else ""
        if any(value for value in row.values()):
            parsed_rows.append(row)

    if "text" in mapping.values():
        return parsed_rows, "dialogue"
    if "inputs" in mapping.values() and "replicas" in mapping.values():
        return parsed_rows, "summary"
    raise ValueError("Nepodařilo se rozpoznat vstup. Očekávám hlavičky POSTAVA/TC/TEXT nebo POSTAVA/VSTUPY/REPLIKY.")


def aggregate_episode(rows: List[Dict[str, str]], source_mode: str, label: str) -> Dict[str, object]:
    by_character: Dict[str, Dict[str, object]] = {}

    if source_mode == "summary":
        for row in rows:
            character = normalize_character(row.get("character", ""))
            if not character:
                continue
            inputs = int(float(row.get("inputs", "0") or 0))
            replicas = int(float(row.get("replicas", "0") or 0))
            current = by_character.setdefault(
                character,
                {
                    "character": character,
                    "inputs": 0,
                    "replicas": 0,
                    "timecodes": [],
                },
            )
            current["inputs"] += inputs
            current["replicas"] += replicas
    else:
        for row in rows:
            character = normalize_character(row.get("character", ""))
            text = row.get("text", "").strip()
            if not character and not text:
                continue
            if not character:
                continue

            current = by_character.setdefault(
                character,
                {
                    "character": character,
                    "inputs": 0,
                    "replicas": 0,
                    "timecodes": [],
                },
            )
            word_count = count_words(text)
            current["inputs"] += 1
            current["replicas"] += count_replicas(word_count)
            timecode = normalize_text(row.get("timecode", ""))
            if timecode:
                current["timecodes"].append(timecode)

    characters = sorted(
        by_character.values(),
        key=lambda item: (-int(item["replicas"]), -int(item["inputs"]), str(item["character"]).lower()),
    )

    return {
        "label": label,
        "sourceMode": source_mode,
        "characters": characters,
        "totals": {
            "characters": len(characters),
            "inputs": sum(int(item["inputs"]) for item in characters),
            "replicas": sum(int(item["replicas"]) for item in characters),
        },
    }


def sanitize_episode_payload(episodes: List[Dict[str, str]]) -> List[Dict[str, str]]:
    sanitized = []
    for index, raw_payload in enumerate(episodes[:MAX_EPISODES]):
        payload = raw_payload if isinstance(raw_payload, dict) else {}
        sanitized.append(
            {
                "label": normalize_text(payload.get("label") or f"{index + 1:02d}") or f"{index + 1:02d}",
                "content": str(payload.get("content", "") or ""),
            }
        )
    if not sanitized:
        sanitized.append({"label": "01", "content": ""})
    return sanitized


def sanitize_assignments(assignments: Dict[str, Dict[str, str]]) -> Dict[str, Dict[str, str]]:
    cleaned: Dict[str, Dict[str, str]] = {}
    for character, payload in (assignments or {}).items():
        key = normalize_character(character)
        if not key:
            continue
        cleaned[key] = {
            "actor": normalize_actor(payload.get("actor", "")),
            "note": normalize_text(payload.get("note", "")),
        }
    return cleaned


def append_validation(
    validations: List[Dict[str, object]],
    severity: str,
    category: str,
    message: str,
    weight: int = 0,
    **metadata: object,
) -> None:
    item: Dict[str, object] = {
        "severity": severity,
        "category": category,
        "message": message,
        "weight": int(weight),
    }
    item.update(metadata)
    validations.append(item)


def classify_load_severity(
    inputs: int,
    replicas: int,
    info_inputs: int,
    info_replicas: int,
    warning_inputs: int,
    warning_replicas: int,
) -> str | None:
    if replicas >= warning_replicas or inputs >= warning_inputs:
        return "warning"
    if replicas >= info_replicas or inputs >= info_inputs:
        return "info"
    return None


def build_validations(complete_rows: List[Dict[str, object]], episodes: List[Dict[str, object]]) -> List[Dict[str, object]]:
    validations: List[Dict[str, object]] = []
    actor_variants: Dict[str, set[str]] = {}
    character_variants: Dict[str, set[str]] = {}
    actor_stats: Dict[str, Dict[str, object]] = {}

    for row in complete_rows:
        character = str(row.get("character", ""))
        actor = normalize_actor(str(row.get("actor", "")))
        total_inputs = int(row.get("totalInputs", 0))
        total_replicas = int(row.get("totalReplicas", 0))

        character_key = loose_match_key(character)
        if character_key:
            character_variants.setdefault(character_key, set()).add(character)

        if not actor and total_replicas >= IMPORTANT_UNASSIGNED_REPLICAS_WARNING:
            works = [str(cell.get("label", "")) for cell in row.get("episodes", []) if int(cell.get("inputs", 0)) > 0]
            works_suffix = f"; díla: {', '.join(works)}" if works else ""
            append_validation(
                validations,
                "warning",
                "Neobsazeno",
                (
                    f"Postava {character} zatím nemá dabéra "
                    f"({total_inputs} vstupů, {total_replicas} replik{works_suffix})."
                ),
                weight=total_replicas,
            )

        if character_key in GENERIC_CHARACTER_KEYS and total_replicas >= GENERIC_BUCKET_INFO_REPLICAS:
            append_validation(
                validations,
                "info",
                "Postavy",
                (
                    f"Postava {character} má {total_replicas} replik. "
                    "Zkontroluj, jestli pod ní není víc konkrétních rolí."
                ),
                weight=total_replicas,
            )

        if not actor:
            continue

        actor_key = loose_match_key(actor)
        if actor_key:
            actor_variants.setdefault(actor_key, set()).add(actor)

        actor_entry = actor_stats.setdefault(
            actor,
            {
                "actor": actor,
                "totalInputs": 0,
                "totalReplicas": 0,
                "characters": set(),
                "episodes": {},
            },
        )
        actor_entry["totalInputs"] += total_inputs
        actor_entry["totalReplicas"] += total_replicas
        actor_entry["characters"].add(character)

        for cell in row.get("episodes", []):
            inputs = int(cell.get("inputs", 0))
            replicas = int(cell.get("replicas", 0))
            if inputs <= 0 and replicas <= 0:
                continue
            label = str(cell.get("label", ""))
            episode_entry = actor_entry["episodes"].setdefault(
                label,
                {
                    "label": label,
                    "inputs": 0,
                    "replicas": 0,
                    "characters": set(),
                },
            )
            episode_entry["inputs"] += inputs
            episode_entry["replicas"] += replicas
            episode_entry["characters"].add(character)

    for variants in character_variants.values():
        if len(variants) < 2:
            continue
        ordered = sorted(variants, key=lambda value: (value.casefold(), value))
        append_validation(
            validations,
            "warning",
            "Jména",
            f"Možná jde o stejnou postavu: {', '.join(ordered)}.",
            weight=len(ordered),
            kind="character_variants",
            variants=ordered,
            actionable=False,
        )

    for variants in actor_variants.values():
        if len(variants) < 2:
            continue
        ordered = sorted(variants, key=lambda value: (value.casefold(), value))
        append_validation(
            validations,
            "info",
            "Jména",
            f"Možná jde o stejného dabéra: {', '.join(ordered)}.",
            weight=len(ordered),
            kind="actor_variants",
            variants=ordered,
            actionable=True,
        )

    for actor, stats in actor_stats.items():
        total_inputs = int(stats["totalInputs"])
        total_replicas = int(stats["totalReplicas"])
        character_count = len(stats["characters"])

        total_load_severity = classify_load_severity(
            total_inputs,
            total_replicas,
            info_inputs=HIGH_ACTOR_TOTAL_INPUTS_INFO,
            info_replicas=HIGH_ACTOR_TOTAL_REPLICAS_INFO,
            warning_inputs=HIGH_ACTOR_TOTAL_INPUTS_WARNING,
            warning_replicas=HIGH_ACTOR_TOTAL_REPLICAS_WARNING,
        )
        if total_load_severity and character_count >= 2:
            append_validation(
                validations,
                total_load_severity,
                "Zátěž",
                (
                    f"Dabér {actor} má vysokou celkovou zátěž "
                    f"({total_inputs} vstupů, {total_replicas} replik)."
                ),
                weight=total_replicas,
            )
        elif (
            character_count >= MANY_ACTOR_CHARACTER_COUNT_INFO
            and total_replicas >= MANY_ACTOR_CHARACTER_MIN_REPLICAS_INFO
        ):
            append_validation(
                validations,
                "info",
                "Obsazení",
                (
                    f"Dabér {actor} má přiřazeno {character_count} postav "
                    f"({total_inputs} vstupů, {total_replicas} replik)."
                ),
                weight=total_replicas,
            )

        for episode in episodes:
            label = str(episode.get("label", ""))
            episode_stats = stats["episodes"].get(label)
            if not episode_stats:
                continue

            episode_inputs = int(episode_stats["inputs"])
            episode_replicas = int(episode_stats["replicas"])
            episode_character_count = len(episode_stats["characters"])

            episode_load_severity = classify_load_severity(
                episode_inputs,
                episode_replicas,
                info_inputs=HIGH_ACTOR_EPISODE_INPUTS_INFO,
                info_replicas=HIGH_ACTOR_EPISODE_REPLICAS_INFO,
                warning_inputs=HIGH_ACTOR_EPISODE_INPUTS_WARNING,
                warning_replicas=HIGH_ACTOR_EPISODE_REPLICAS_WARNING,
            )
            if episode_load_severity and episode_character_count >= 2:
                append_validation(
                    validations,
                    episode_load_severity,
                    "Zátěž",
                    (
                        f"Dabér {actor} má v díle {label} vysokou zátěž "
                        f"({episode_inputs} vstupů, {episode_replicas} replik)."
                    ),
                    weight=episode_replicas,
                )
            elif (
                episode_character_count >= MANY_ACTOR_EPISODE_CHARACTER_COUNT_INFO
                and episode_replicas >= MANY_ACTOR_EPISODE_MIN_REPLICAS_INFO
            ):
                append_validation(
                    validations,
                    "info",
                    "Obsazení",
                    (
                        f"Dabér {actor} má v díle {label} {episode_character_count} postavy "
                        f"({episode_inputs} vstupů, {episode_replicas} replik)."
                    ),
                    weight=episode_replicas,
                )

    validations.sort(
        key=lambda item: (
            VALIDATION_SEVERITY_ORDER.get(str(item.get("severity", "info")), 99),
            -int(item.get("weight", 0)),
            str(item.get("category", "")),
            str(item.get("message", "")).casefold(),
        )
    )
    for item in validations:
        item.pop("weight", None)
    return validations


def build_project(payload: Dict[str, object]) -> Dict[str, object]:
    title = normalize_text(str(payload.get("title", "Obsazení projektu"))) or "Obsazení projektu"
    episodes_payload = sanitize_episode_payload(list(payload.get("episodes", [])))
    assignments = sanitize_assignments(dict(payload.get("assignments", {})))

    episodes: List[Dict[str, object]] = []
    all_characters: Dict[str, Dict[str, object]] = {}

    for episode_index, episode_payload in enumerate(episodes_payload):
        rows, source_mode = parse_rows(episode_payload["content"])
        episode = aggregate_episode(rows, source_mode, episode_payload["label"])
        episode["index"] = episode_index + 1
        episodes.append(episode)

        for character_row in episode["characters"]:
            character = str(character_row["character"])
            aggregate = all_characters.setdefault(
                character,
                {
                    "character": character,
                    "actor": assignments.get(character, {}).get("actor", ""),
                    "note": assignments.get(character, {}).get("note", ""),
                    "totalInputs": 0,
                    "totalReplicas": 0,
                    "episodes": {},
                },
            )
            aggregate["totalInputs"] += int(character_row["inputs"])
            aggregate["totalReplicas"] += int(character_row["replicas"])
            aggregate["episodes"][episode_index] = {
                "inputs": int(character_row["inputs"]),
                "replicas": int(character_row["replicas"]),
            }

    complete_rows = []
    for character, aggregate in all_characters.items():
        episode_cells = []
        for episode_index, episode in enumerate(episodes):
            stats = aggregate["episodes"].get(episode_index, {"inputs": 0, "replicas": 0})
            episode_cells.append(
                {
                    "label": episode["label"],
                    "inputs": int(stats["inputs"]),
                    "replicas": int(stats["replicas"]),
                    "display": f"{int(stats['inputs'])} / {int(stats['replicas'])}",
                }
            )

        complete_rows.append(
            {
                "character": character,
                "actor": str(aggregate["actor"]),
                "note": str(aggregate["note"]),
                "totalInputs": int(aggregate["totalInputs"]),
                "totalReplicas": int(aggregate["totalReplicas"]),
                "episodes": episode_cells,
            }
        )

    complete_rows.sort(
        key=lambda item: (-int(item["totalReplicas"]), -int(item["totalInputs"]), item["character"].lower())
    )

    actor_totals: Dict[str, Dict[str, object]] = {}
    missing = {"characters": 0, "inputs": 0, "replicas": 0}
    for row in complete_rows:
        actor = normalize_actor(str(row["actor"]))
        if not actor:
            missing["characters"] += 1
            missing["inputs"] += int(row["totalInputs"])
            missing["replicas"] += int(row["totalReplicas"])
            continue

        stats = actor_totals.setdefault(
            actor,
            {"actor": actor, "totalInputs": 0, "totalReplicas": 0, "note": ""},
        )
        stats["totalInputs"] += int(row["totalInputs"])
        stats["totalReplicas"] += int(row["totalReplicas"])

    actors = sorted(
        actor_totals.values(),
        key=lambda item: (-int(item["totalReplicas"]), -int(item["totalInputs"]), str(item["actor"]).lower()),
    )
    validations = build_validations(complete_rows, episodes)
    validation_summary = {
        "warningCount": sum(1 for item in validations if item.get("severity") == "warning"),
        "infoCount": sum(1 for item in validations if item.get("severity") == "info"),
    }

    return {
        "title": title,
        "episodes": episodes,
        "complete": complete_rows,
        "actors": actors,
        "validations": validations,
        "validationSummary": validation_summary,
        "missing": missing,
        "stats": {
            "episodeCount": len(episodes),
            "characterCount": len(complete_rows),
            "inputs": sum(int(row["totalInputs"]) for row in complete_rows),
            "replicas": sum(int(row["totalReplicas"]) for row in complete_rows),
        },
    }
