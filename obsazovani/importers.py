from __future__ import annotations

import csv
import html
import io
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZipFile

from .core import (
    HEADER_ALIASES,
    normalize_actor,
    normalize_character,
    normalize_header,
    normalize_text,
    parse_rows,
)

XLSX_NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

CODE_TOKEN_RE = re.compile(r"^[A-ZÁČĎÉĚÍĹĽŇÓÔŘŠŤÚŮÝŽ]\d+[.,-]?$")
TR_RE = re.compile(r"<tr\b[^>]*>(.*?)</tr>", re.IGNORECASE | re.DOTALL)
TD_RE = re.compile(r"<td\b[^>]*>(.*?)</td>", re.IGNORECASE | re.DOTALL)
TAG_RE = re.compile(r"<[^>]+>")


@dataclass(frozen=True)
class ImportedEpisodeSource:
    content: str
    source_format: str
    assignments: dict[str, dict[str, str]] = field(default_factory=dict)


@dataclass(frozen=True)
class ImportableWorkbookSheet:
    sheet_name: str
    source_format: str
    row_count: int
    display_name: str
    content: str
    assignments: dict[str, dict[str, str]] = field(default_factory=dict)


def import_episode_source(path: Path, sheet_name: str | None = None) -> ImportedEpisodeSource:
    suffix = path.suffix.lower()
    if suffix in {".txt", ".tsv", ".csv"}:
        content = _read_text_file(path)
        return ImportedEpisodeSource(content, "plain-text", _extract_content_assignments(content))
    if suffix == ".xlsx":
        return _import_xlsx_source(path, sheet_name)
    if suffix == ".docx":
        if _is_iyuno_docx(path):
            dialogue_rows = _parse_iyuno_docx(path)
            return ImportedEpisodeSource(
                _serialize_dialogue_rows(dialogue_rows),
                "iyuno-docx",
                _collect_assignments(dialogue_rows),
            )
        dialogue_rows = _parse_classic_docx(path)
        return ImportedEpisodeSource(
            _serialize_dialogue_rows(dialogue_rows),
            "classic-docx",
            _collect_assignments(dialogue_rows),
        )
    if suffix == ".doc":
        dialogue_rows = _parse_iyuno_doc(path)
        return ImportedEpisodeSource(
            _serialize_dialogue_rows(dialogue_rows),
            "iyuno-doc",
            _collect_assignments(dialogue_rows),
        )
    raise ValueError(
        "Nepodporovaný formát souboru. Použij .txt, .tsv, .csv, .xlsx, .doc nebo .docx."
    )


def list_importable_xlsx_sheets(path: Path) -> list[ImportableWorkbookSheet]:
    return _collect_xlsx_sheet_candidates(path)


def _read_text_file(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "cp1250", "cp1252"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")


def _serialize_dialogue_rows(rows: list[dict[str, str]]) -> str:
    output = io.StringIO()
    writer = csv.writer(output, delimiter="\t", lineterminator="\n")
    include_actor = any(normalize_actor(row.get("actor", "")) for row in rows)
    include_note = any(normalize_text(row.get("note", "")) for row in rows)

    headers = ["POSTAVA", "TC", "TEXT"]
    if include_actor:
        headers.append("DABÉR")
    if include_note:
        headers.append("POZNÁMKA")
    writer.writerow(headers)

    for row in rows:
        values = [row["character"], row["timecode"], row["text"]]
        if include_actor:
            values.append(normalize_actor(row.get("actor", "")))
        if include_note:
            values.append(normalize_text(row.get("note", "")))
        writer.writerow(values)
    return output.getvalue().rstrip("\n")


def _serialize_summary_rows(rows: list[dict[str, str]]) -> str:
    output = io.StringIO()
    writer = csv.writer(output, delimiter="\t", lineterminator="\n")
    include_actor = any(normalize_actor(row.get("actor", "")) for row in rows)
    include_note = any(normalize_text(row.get("note", "")) for row in rows)

    headers = ["POSTAVA", "VSTUPY", "REPLIKY"]
    if include_actor:
        headers.append("DABÉR")
    if include_note:
        headers.append("POZNÁMKA")
    writer.writerow(headers)

    for row in rows:
        values = [row["character"], row["inputs"], row["replicas"]]
        if include_actor:
            values.append(normalize_actor(row.get("actor", "")))
        if include_note:
            values.append(normalize_text(row.get("note", "")))
        writer.writerow(values)
    return output.getvalue().rstrip("\n")


def _collect_assignments(rows: list[dict[str, str]]) -> dict[str, dict[str, str]]:
    assignments: dict[str, dict[str, str]] = {}
    for row in rows:
        character = normalize_character(row.get("character", ""))
        if not character:
            continue

        actor = normalize_actor(row.get("actor", ""))
        note = normalize_text(row.get("note", ""))
        if not actor and not note:
            continue

        assignment = assignments.setdefault(character, {"actor": "", "note": ""})
        if actor and not assignment["actor"]:
            assignment["actor"] = actor
        if note and not assignment["note"]:
            assignment["note"] = note
    return assignments


def _extract_content_assignments(content: str) -> dict[str, dict[str, str]]:
    try:
        rows, _source_mode = parse_rows(content)
    except ValueError:
        return {}
    return _collect_assignments(rows)


def _normalize_timecode(value: str) -> str:
    cleaned = normalize_text(value)
    if not cleaned:
        return ""
    if cleaned.startswith(("O", "o")) and len(cleaned) > 1 and cleaned[1].isdigit():
        cleaned = f"0{cleaned[1:]}"
    return cleaned


def _normalize_import_character(raw_value: str) -> str:
    speaker = normalize_text(raw_value).rstrip(":").strip()
    if not speaker:
        return ""

    if " : " in speaker:
        trailing = speaker.rsplit(":", 1)[-1].strip()
        if trailing:
            speaker = trailing

    code_prefix, separator, remainder = speaker.partition("-")
    if separator and CODE_TOKEN_RE.fullmatch(code_prefix.strip()):
        speaker = remainder.strip()

    tokens = speaker.split()
    while tokens and (CODE_TOKEN_RE.fullmatch(tokens[0]) or tokens[0] in {":", ",", "."}):
        tokens.pop(0)
    normalized = " ".join(tokens).strip()
    if normalized:
        return normalize_character(normalized)
    return normalize_character(speaker)


def _cell_reference_to_index(cell_reference: str) -> int:
    letters = "".join(character for character in cell_reference if character.isalpha())
    index = 0
    for character in letters:
        index = index * 26 + (ord(character.upper()) - ord("A") + 1)
    return index - 1


def _extract_shared_strings(archive: ZipFile) -> list[str]:
    shared_strings_path = "xl/sharedStrings.xml"
    if shared_strings_path not in archive.namelist():
        return []

    shared_strings_root = ET.fromstring(archive.read(shared_strings_path))
    values: list[str] = []
    for item in shared_strings_root.findall("a:si", XLSX_NS):
        values.append("".join(node.text or "" for node in item.iterfind(".//a:t", XLSX_NS)))
    return values


def _xlsx_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        index = int(cell.findtext("a:v", default="0", namespaces=XLSX_NS))
        return shared_strings[index]
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iterfind(".//a:t", XLSX_NS))
    return cell.findtext("a:v", default="", namespaces=XLSX_NS)


def _dense_row_values(row_values: dict[int, str]) -> list[str]:
    if not row_values:
        return []
    max_index = max(row_values)
    return [str(row_values.get(index, "") or "").strip() for index in range(max_index + 1)]


def _parse_workbook_sheets(path: Path) -> list[tuple[str, list[list[str]]]]:
    with ZipFile(path) as archive:
        shared_strings = _extract_shared_strings(archive)
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationship_map = {relationship.attrib["Id"]: relationship.attrib["Target"] for relationship in rels_root}
        sheet_elements = workbook_root.findall("a:sheets/a:sheet", XLSX_NS)
        if not sheet_elements:
            raise ValueError("Excel soubor neobsahuje žádný list.")

        workbook_sheets: list[tuple[str, list[list[str]]]] = []
        for sheet in sheet_elements:
            relationship_id = sheet.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            if relationship_id not in relationship_map:
                continue

            target = relationship_map[relationship_id]
            if not target.startswith("xl/"):
                target = f"xl/{target.lstrip('/')}"
            if not target.startswith("xl/worksheets/"):
                continue

            sheet_root = ET.fromstring(archive.read(target))
            rows: list[list[str]] = []
            for row in sheet_root.findall("a:sheetData/a:row", XLSX_NS):
                row_values: dict[int, str] = {}
                for cell in row.findall("a:c", XLSX_NS):
                    reference = cell.attrib.get("r", "")
                    row_values[_cell_reference_to_index(reference)] = _xlsx_cell_value(cell, shared_strings)
                if row_values:
                    rows.append(_dense_row_values(row_values))
            workbook_sheets.append((sheet.attrib.get("name", ""), rows))
        return workbook_sheets


def _header_positions(row: list[str]) -> dict[str, int]:
    positions: dict[str, int] = {}
    for index, cell in enumerate(row):
        canonical = HEADER_ALIASES.get(normalize_header(cell))
        if canonical and canonical not in positions:
            positions[canonical] = index
    return positions


def _summary_character_is_ignored(value: str) -> bool:
    normalized = normalize_header(value)
    return normalized in {"(blank)", "grand total", "row labels"}


def _sheet_is_known_aggregate(sheet_name: str) -> bool:
    normalized = normalize_header(sheet_name)
    return normalized in {"herci", "komplet", "prehled"}


def _xlsx_candidate_display_name(sheet_name: str, source_format: str, row_count: int) -> str:
    labels = {
        "netflix-xlsx": "Netflix dialogy",
        "xlsx-dialogue": "POSTAVA / TC / TEXT",
        "xlsx-summary": "POSTAVA / VSTUPY / REPLIKY",
    }
    row_text = f"{row_count} řádků" if row_count else "bez dat"
    return f"{sheet_name} · {labels.get(source_format, source_format)} · {row_text}"


def _build_importable_sheet(
    sheet_name: str,
    source_format: str,
    content: str,
    row_count: int,
    assignments: dict[str, dict[str, str]] | None = None,
) -> ImportableWorkbookSheet:
    return ImportableWorkbookSheet(
        sheet_name=sheet_name,
        source_format=source_format,
        row_count=row_count,
        display_name=_xlsx_candidate_display_name(sheet_name, source_format, row_count),
        content=content,
        assignments=assignments or {},
    )


def _parse_netflix_sheet(sheet_name: str, rows: list[list[str]]) -> ImportableWorkbookSheet | None:
    if not rows:
        return None

    header_lookup = {normalize_text(value).upper(): index for index, value in enumerate(rows[0])}
    required_headers = {"SOURCE", "DIALOGUE", "IN-TIMECODE"}
    if not required_headers.issubset(header_lookup):
        return None

    dialogue_rows: list[dict[str, str]] = []
    for row in rows[1:]:
        source_index = header_lookup["SOURCE"]
        dialogue_index = header_lookup["DIALOGUE"]
        timecode_index = header_lookup["IN-TIMECODE"]
        character = _normalize_import_character(row[source_index] if source_index < len(row) else "")
        text = normalize_text(row[dialogue_index] if dialogue_index < len(row) else "")
        if not character or not text:
            continue
        dialogue_rows.append(
            {
                "character": character,
                "timecode": _normalize_timecode(row[timecode_index] if timecode_index < len(row) else ""),
                "text": text,
            }
        )
    return _build_importable_sheet(
        sheet_name,
        "netflix-xlsx",
        _serialize_dialogue_rows(dialogue_rows),
        len(dialogue_rows),
        _collect_assignments(dialogue_rows),
    )


def _parse_dialogue_sheet(sheet_name: str, rows: list[list[str]]) -> ImportableWorkbookSheet | None:
    if not rows or _sheet_is_known_aggregate(sheet_name):
        return None

    positions = _header_positions(rows[0])
    character_index = positions.get("character")
    text_index = positions.get("text")
    timecode_index = positions.get("timecode")
    actor_index = positions.get("actor")
    note_index = positions.get("note")
    if character_index is None or text_index is None:
        return None
    if timecode_index is not None and max(character_index, text_index, timecode_index) > 3:
        return None
    if timecode_index is None and max(character_index, text_index) > 2:
        return None

    dialogue_rows: list[dict[str, str]] = []
    for row in rows[1:]:
        character = _normalize_import_character(row[character_index] if character_index < len(row) else "")
        text = normalize_text(row[text_index] if text_index < len(row) else "")
        if not character or not text:
            continue
        timecode = row[timecode_index] if timecode_index is not None and timecode_index < len(row) else ""
        dialogue_rows.append(
            {
                "character": character,
                "timecode": _normalize_timecode(timecode),
                "text": text,
                "actor": row[actor_index] if actor_index is not None and actor_index < len(row) else "",
                "note": row[note_index] if note_index is not None and note_index < len(row) else "",
            }
        )
    return _build_importable_sheet(
        sheet_name,
        "xlsx-dialogue",
        _serialize_dialogue_rows(dialogue_rows),
        len(dialogue_rows),
        _collect_assignments(dialogue_rows),
    )


def _parse_summary_sheet(sheet_name: str, rows: list[list[str]]) -> ImportableWorkbookSheet | None:
    if not rows or _sheet_is_known_aggregate(sheet_name):
        return None

    positions = _header_positions(rows[0])
    character_index = positions.get("character")
    inputs_index = positions.get("inputs")
    replicas_index = positions.get("replicas")
    actor_index = positions.get("actor")
    note_index = positions.get("note")
    if character_index is None or inputs_index is None or replicas_index is None:
        return None
    if max(character_index, inputs_index, replicas_index) > 3:
        return None

    summary_rows: list[dict[str, str]] = []
    for row in rows[1:]:
        raw_character = row[character_index] if character_index < len(row) else ""
        if _summary_character_is_ignored(raw_character):
            continue

        character = _normalize_import_character(raw_character)
        inputs = normalize_text(row[inputs_index] if inputs_index < len(row) else "")
        replicas = normalize_text(row[replicas_index] if replicas_index < len(row) else "")
        if not character:
            continue
        if inputs in {"", "0", "0.0"} and replicas in {"", "0", "0.0"}:
            continue

        summary_rows.append(
            {
                "character": character,
                "inputs": inputs or "0",
                "replicas": replicas or "0",
                "actor": row[actor_index] if actor_index is not None and actor_index < len(row) else "",
                "note": row[note_index] if note_index is not None and note_index < len(row) else "",
            }
        )
    return _build_importable_sheet(
        sheet_name,
        "xlsx-summary",
        _serialize_summary_rows(summary_rows),
        len(summary_rows),
        _collect_assignments(summary_rows),
    )


def _collect_xlsx_sheet_candidates(path: Path) -> list[ImportableWorkbookSheet]:
    candidates: list[ImportableWorkbookSheet] = []
    for sheet_name, rows in _parse_workbook_sheets(path):
        candidate = (
            _parse_netflix_sheet(sheet_name, rows)
            or _parse_dialogue_sheet(sheet_name, rows)
            or _parse_summary_sheet(sheet_name, rows)
        )
        if candidate is not None:
            candidates.append(candidate)
    if not candidates:
        raise ValueError(
            "Excel soubor neobsahuje podporovaný list. Očekávám hlavičky POSTAVA/TC/TEXT nebo POSTAVA/VSTUPY/REPLIKY."
        )
    return candidates


def _import_xlsx_source(path: Path, sheet_name: str | None) -> ImportedEpisodeSource:
    candidates = _collect_xlsx_sheet_candidates(path)
    selected = candidates[0] if sheet_name is None else next(
        (candidate for candidate in candidates if candidate.sheet_name == sheet_name),
        None,
    )
    if selected is None:
        raise ValueError(f"V Excel souboru nebyl nalezen list „{sheet_name}“.")
    return ImportedEpisodeSource(selected.content, selected.source_format, selected.assignments)


def _iter_docx_paragraph_lines(path: Path) -> list[str]:
    document_root = _read_docx_document_root(path)

    lines: list[str] = []
    for paragraph in document_root.findall(".//w:body/w:p", WORD_NS):
        parts: list[str] = []
        for element in paragraph.iter():
            tag = element.tag.split("}")[-1]
            if tag == "t":
                parts.append(element.text or "")
            elif tag == "tab":
                parts.append("\t")
            elif tag in {"br", "cr"}:
                parts.append("\n")
        text = "".join(parts).strip("\n")
        if text.strip():
            lines.append(text)
    return lines


def _parse_classic_docx(path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    current_character = ""

    for line in _iter_docx_paragraph_lines(path):
        parts = line.split("\t")
        if len(parts) < 3:
            continue

        raw_character, raw_timecode = parts[0], parts[1]
        raw_text = normalize_text("\t".join(parts[2:]))
        if not raw_text:
            continue

        character = _normalize_import_character(raw_character)
        if character:
            current_character = character
        else:
            character = current_character

        if not character:
            continue

        rows.append(
            {
                "character": character,
                "timecode": _normalize_timecode(raw_timecode),
                "text": raw_text,
            }
        )

    if not rows:
        raise ValueError("V .docx souboru se nepodařilo najít dialogové řádky.")
    return rows


def _read_docx_document_root(path: Path) -> ET.Element:
    with ZipFile(path) as archive:
        return ET.fromstring(archive.read("word/document.xml"))


def _word_cell_text(cell: ET.Element) -> str:
    parts: list[str] = []
    for element in cell.iter():
        tag = element.tag.split("}")[-1]
        if tag == "t":
            parts.append(element.text or "")
        elif tag == "tab":
            parts.append("\t")
        elif tag in {"br", "cr"}:
            parts.append("\n")
    return "".join(parts).strip()


def _iter_word_tables(path: Path) -> list[list[list[str]]]:
    document_root = _read_docx_document_root(path)
    body = document_root.find("w:body", WORD_NS)
    if body is None:
        return []

    tables: list[list[list[str]]] = []
    for table in body.findall("w:tbl", WORD_NS):
        rows: list[list[str]] = []
        for row in table.findall("w:tr", WORD_NS):
            cells = [_word_cell_text(cell) for cell in row.findall("w:tc", WORD_NS)]
            if cells:
                rows.append(cells)
        if rows:
            tables.append(rows)
    return tables


def _is_iyuno_docx(path: Path) -> bool:
    expected_header = ["Character", "TC", "Note", "TEXT"]
    for table in _iter_word_tables(path):
        if table and table[0][:4] == expected_header:
            return True
    return False


def _parse_iyuno_docx(path: Path) -> list[dict[str, str]]:
    current_character = ""
    dialogue_rows: list[dict[str, str]] = []

    for table in _iter_word_tables(path):
        if not table or table[0][:4] != ["Character", "TC", "Note", "TEXT"]:
            continue

        for cells in table[1:]:
            if len(cells) < 4:
                continue

            character = _normalize_import_character(cells[0])
            timecode = _normalize_timecode(cells[1])
            text = normalize_text(cells[3].replace("\n", " "))
            if not text:
                continue

            if character:
                current_character = character
            else:
                character = current_character

            if not character:
                continue

            dialogue_rows.append(
                {
                    "character": character,
                    "timecode": timecode,
                    "text": text,
                    "note": normalize_text(cells[2].replace("\n", " ")),
                }
            )

    if not dialogue_rows:
        raise ValueError("V IYUNO .docx souboru se nepodařilo najít žádné dialogové řádky.")
    return dialogue_rows


def _convert_doc_to_html(path: Path) -> str:
    if sys.platform == "darwin":
        try:
            completed = subprocess.run(
                ["textutil", "-convert", "html", "-stdout", str(path)],
                check=True,
                capture_output=True,
            )
        except (OSError, subprocess.CalledProcessError) as exc:
            raise ValueError(f"Nepodařilo se převést Word .doc soubor: {exc}") from exc
        return completed.stdout.decode("utf-8", errors="replace")

    raise ValueError(
        "Import .doc je zatím implementovaný přes textutil na macOS. Pro Windows bude potřeba přidat druhý backend."
    )


def _clean_html_cell(cell_html: str) -> str:
    text = re.sub(r"<br\s*/?>", "\n", cell_html, flags=re.IGNORECASE)
    text = TAG_RE.sub("", text)
    text = html.unescape(text).replace("\xa0", " ")
    lines = [normalize_text(line) for line in text.splitlines()]
    return "\n".join(line for line in lines if line)


def _extract_html_rows(fragment: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for row_html in TR_RE.findall(fragment):
        cells = [_clean_html_cell(cell_html) for cell_html in TD_RE.findall(row_html)]
        if cells:
            rows.append(cells)
    return rows


def _parse_iyuno_doc(path: Path) -> list[dict[str, str]]:
    html_document = _convert_doc_to_html(path)
    marker = html_document.find("Statistics timestamp:")
    if marker == -1:
        raise ValueError("IYUNO .doc neobsahuje očekávanou část s dialogovým listingem.")

    script_fragment = html_document[marker:]
    html_rows = _extract_html_rows(script_fragment)

    dialogue_rows: list[dict[str, str]] = []
    current_character = ""
    for cells in html_rows:
        if len(cells) < 4:
            continue

        character = _normalize_import_character(cells[0])
        timecode = _normalize_timecode(cells[1])
        text = normalize_text(cells[3].replace("\n", " "))
        if not text:
            continue

        if character:
            current_character = character
        else:
            character = current_character

        if not character:
            continue

        dialogue_rows.append(
            {
                "character": character,
                "timecode": timecode,
                "text": text,
                "note": normalize_text(cells[2].replace("\n", " ")),
            }
        )

    if not dialogue_rows:
        raise ValueError("V IYUNO .doc souboru se nepodařilo najít žádné dialogové řádky.")
    return dialogue_rows
