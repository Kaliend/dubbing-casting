from __future__ import annotations

import copy
import io
from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

from .core import MAX_EPISODES, normalize_text

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
X14AC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
XML_NS = "http://www.w3.org/XML/1998/namespace"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

COMPATIBILITY_PREFIX_URIS = {
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x15ac": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
    "x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
    "xcalcf": "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures",
    "xlrd2": "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2",
}

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("x14ac", X14AC_NS)
for prefix, uri in COMPATIBILITY_PREFIX_URIS.items():
    if prefix not in {"", "r", "x14ac"}:
        ET.register_namespace(prefix, uri)
ET.register_namespace("vt", VT_NS)

TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "templates" / "SABLONA_OBSAZENI(01-06).xlsx"

SHEET_PATHS = {
    "HERCI": "xl/worksheets/sheet1.xml",
    "KOMPLET": "xl/worksheets/sheet2.xml",
    "PŘEHLED": "xl/worksheets/sheet3.xml",
    "01": "xl/worksheets/sheet4.xml",
    "02": "xl/worksheets/sheet5.xml",
    "03": "xl/worksheets/sheet6.xml",
    "04": "xl/worksheets/sheet7.xml",
    "05": "xl/worksheets/sheet8.xml",
    "06": "xl/worksheets/sheet9.xml",
}

TABLE_PATHS = {
    "HERCI": "xl/tables/table1.xml",
    "KOMPLET": "xl/tables/table2.xml",
    "01": "xl/tables/table3.xml",
    "02": "xl/tables/table4.xml",
    "03": "xl/tables/table5.xml",
    "04": "xl/tables/table6.xml",
    "05": "xl/tables/table7.xml",
    "06": "xl/tables/table8.xml",
}

DEFAULT_EXPORT_SLOT_LABELS = tuple(f"{index + 1:02d}" for index in range(MAX_EPISODES))
EPISODE_SHEET_KEYS = DEFAULT_EXPORT_SLOT_LABELS
STATIC_SHEET_NAMES = ("HERCI", "KOMPLET")
STATIC_KOMPLET_COLUMNS = ("POSTAVA", "DABÉR", "VSTUPY", "REPLIKY", "POZNÁMKY")
INVALID_SHEET_NAME_RE = re.compile(r"[\[\]:*?/\\\\]")
HEADER_FILL_RGB = "FF3F332A"
HEADER_FONT_RGB = "FFF9F6F1"
KOMPLET_HEADER_PRIMARY_FILL_RGB = "FF2F241D"
KOMPLET_HEADER_EPISODE_FILL_RGB = "FF625040"
KOMPLET_HEADER_SUMMARY_FILL_RGB = "FF435B6D"
KOMPLET_HEADER_NOTE_FILL_RGB = "FF544E48"
MISSING_FILL_RGB = "FFFFF2E0"
HERCI_HEADER_ROW_HEIGHT = 34.0
HERCI_BODY_ROW_HEIGHT = 20.0


def q(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def aq(name: str) -> str:
    return f"{{{APP_NS}}}{name}"


def vq(name: str) -> str:
    return f"{{{VT_NS}}}{name}"


def ensure_compatibility_prefixes(xml_bytes: bytes) -> bytes:
    text = xml_bytes.decode("utf-8")
    referenced_prefixes: set[str] = set()
    declared_prefixes = set(re.findall(r"\bxmlns:([A-Za-z0-9_]+)=", text))

    for attribute in ("Ignorable", "Requires"):
        for _, raw_value in re.findall(rf'{attribute}=(["\'])(.*?)\1', text):
            referenced_prefixes.update(token for token in raw_value.split() if ":" not in token)

    missing = [
        prefix
        for prefix in referenced_prefixes
        if prefix in COMPATIBILITY_PREFIX_URIS and prefix not in declared_prefixes
    ]
    if not missing:
        return xml_bytes

    declarations = "".join(
        f' xmlns:{prefix}="{COMPATIBILITY_PREFIX_URIS[prefix]}"'
        for prefix in missing
    )
    patched = re.sub(r"(<[A-Za-z0-9:_-]+)([^>]*?)>", rf"\1\2{declarations}>", text, count=1)
    return patched.encode("utf-8")


def serialize_xml(root: ET.Element) -> bytes:
    return ensure_compatibility_prefixes(ET.tostring(root, encoding="utf-8", xml_declaration=True))


def serialize_plain_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def inline_cell(ref: str, value: str = "", style: int | None = None) -> ET.Element:
    cell = ET.Element(q("c"), {"r": ref})
    if style is not None:
        cell.set("s", str(style))
    if value != "":
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, q("is"))
        text = ET.SubElement(inline, q("t"))
        if value.strip() != value or "  " in value:
            text.set(f"{{{XML_NS}}}space", "preserve")
        text.text = value
    return cell


def number_cell(ref: str, value: int | float | None = None, style: int | None = None) -> ET.Element:
    cell = ET.Element(q("c"), {"r": ref})
    if style is not None:
        cell.set("s", str(style))
    if value is not None:
        ET.SubElement(cell, q("v")).text = str(value)
    return cell


def blank_cell(ref: str, style: int | None = None) -> ET.Element:
    return inline_cell(ref, "", style)


def column_name(index: int) -> str:
    if index <= 0:
        raise ValueError("Column index must be positive.")

    label = ""
    current = index
    while current:
        current, remainder = divmod(current - 1, 26)
        label = chr(65 + remainder) + label
    return label


def cell_ref(column_index: int, row_index: int) -> str:
    return f"{column_name(column_index)}{row_index}"


def column_letters(ref: str) -> str:
    match = re.match(r"([A-Z]+)\d+$", ref)
    return match.group(1) if match else ""


def cell_text(cell: ET.Element | None) -> str:
    if cell is None:
        return ""

    inline_text = cell.find(q("is"))
    if inline_text is not None:
        text_node = inline_text.find(q("t"))
        return text_node.text or "" if text_node is not None else ""

    value = cell.find(q("v"))
    if value is not None and value.text is not None:
        return value.text
    return ""


def find_row(root: ET.Element, row_index: int) -> ET.Element | None:
    sheet_data = root.find(q("sheetData"))
    if sheet_data is None:
        return None
    return sheet_data.find(f"{q('row')}[@r='{row_index}']")


def find_cell(row: ET.Element | None, ref: str) -> ET.Element | None:
    if row is None:
        return None
    return row.find(f"{q('c')}[@r='{ref}']")


def suggested_width(values: list[str], minimum: float, maximum: float, padding: float = 2.0) -> float:
    visible_lengths = [len(normalize_text(value)) for value in values if normalize_text(value)]
    if not visible_lengths:
        return minimum
    return max(minimum, min(maximum, max(visible_lengths) + padding))


def make_row(index: int, max_col: int, cells: list[ET.Element], template_row: ET.Element | None = None) -> ET.Element:
    if template_row is not None:
        row = copy.deepcopy(template_row)
        row.clear()
        row.tag = q("row")
        for key in list(template_row.attrib):
            row.attrib[key] = template_row.attrib[key]
    else:
        row = ET.Element(q("row"))
    row.set("r", str(index))
    row.set("spans", f"1:{max_col}")
    row.set(f"{{{X14AC_NS}}}dyDescent", "0.2")
    for cell in cells:
        row.append(cell)
    return row


def reset_sheet_data(root: ET.Element, rows: list[ET.Element], dimension_ref: str) -> None:
    sheet_data = root.find(q("sheetData"))
    if sheet_data is None:
        raise ValueError("Worksheet is missing sheetData.")
    sheet_data.clear()
    for row in rows:
        sheet_data.append(row)
    dimension = root.find(q("dimension"))
    if dimension is not None:
        dimension.set("ref", dimension_ref)


def update_table(root: ET.Element, ref: str) -> None:
    root.set("ref", ref)
    root.set("totalsRowCount", "0")
    auto_filter = root.find(q("autoFilter"))
    if auto_filter is not None:
        auto_filter.set("ref", ref)
    sort_state = root.find(q("sortState"))
    if sort_state is not None:
        root.remove(sort_state)
    strip_table_formula_metadata(root)


def strip_table_formula_metadata(root: ET.Element) -> None:
    table_columns = root.find(q("tableColumns"))
    if table_columns is None:
        return

    for column in table_columns.findall(q("tableColumn")):
        for tag_name in ("calculatedColumnFormula", "totalsRowFormula"):
            element = column.find(q(tag_name))
            if element is not None:
                column.remove(element)


def ensure_solid_fill(styles_root: ET.Element, rgb: str) -> int:
    fills = styles_root.find(q("fills"))
    if fills is None:
        raise ValueError("Workbook styles are missing fills.")

    for index, fill in enumerate(fills.findall(q("fill"))):
        pattern = fill.find(q("patternFill"))
        if pattern is None or pattern.attrib.get("patternType") != "solid":
            continue
        fg = pattern.find(q("fgColor"))
        if fg is not None and fg.attrib.get("rgb") == rgb:
            return index

    fill = ET.Element(q("fill"))
    pattern = ET.SubElement(fill, q("patternFill"), {"patternType": "solid"})
    ET.SubElement(pattern, q("fgColor"), {"rgb": rgb})
    ET.SubElement(pattern, q("bgColor"), {"indexed": "64"})
    fills.append(fill)
    fills.set("count", str(len(fills)))
    return len(fills) - 1


def ensure_colored_font(
    styles_root: ET.Element,
    cache: dict[tuple[int, str], int],
    base_font_id: int,
    rgb: str,
) -> int:
    key = (base_font_id, rgb)
    if key in cache:
        return cache[key]

    fonts = styles_root.find(q("fonts"))
    if fonts is None:
        raise ValueError("Workbook styles are missing fonts.")

    font_list = fonts.findall(q("font"))
    if base_font_id >= len(font_list):
        raise ValueError(f"Base font {base_font_id} does not exist.")

    font = copy.deepcopy(font_list[base_font_id])
    color = font.find(q("color"))
    if color is None:
        color = ET.SubElement(font, q("color"))
    color.attrib.clear()
    color.set("rgb", rgb)
    fonts.append(font)
    fonts.set("count", str(len(fonts)))
    font_id = len(fonts) - 1
    cache[key] = font_id
    return font_id


def ensure_filled_style(
    styles_root: ET.Element,
    cache: dict[tuple[int, int, int | None], int],
    base_style: int,
    fill_id: int,
    font_id: int | None = None,
) -> int:
    key = (base_style, fill_id, font_id)
    if key in cache:
        return cache[key]

    cell_xfs = styles_root.find(q("cellXfs"))
    if cell_xfs is None:
        raise ValueError("Workbook styles are missing cellXfs.")

    xfs = cell_xfs.findall(q("xf"))
    if base_style >= len(xfs):
        raise ValueError(f"Base style {base_style} does not exist.")

    xf = copy.deepcopy(xfs[base_style])
    xf.set("fillId", str(fill_id))
    xf.set("applyFill", "1")
    if font_id is not None:
        xf.set("fontId", str(font_id))
        xf.set("applyFont", "1")
    cell_xfs.append(xf)
    cell_xfs.set("count", str(len(cell_xfs)))
    style_id = len(cell_xfs) - 1
    cache[key] = style_id
    return style_id


def set_row_height(row: ET.Element | None, height: float) -> None:
    if row is None:
        return
    row.set("ht", str(height))
    row.set("customHeight", "1")


def apply_header_fill(
    sheet_root: ET.Element,
    styles_root: ET.Element,
    fill_style_cache: dict[tuple[int, int, int | None], int],
    font_cache: dict[tuple[int, str], int],
    fill_id: int,
    font_rgb: str,
    column_fill_ids: dict[str, int] | None = None,
    row_height: float = 24.0,
) -> None:
    header_row = find_row(sheet_root, 1)
    if header_row is None:
        return

    set_row_height(header_row, row_height)
    cell_xfs = styles_root.find(q("cellXfs"))
    if cell_xfs is None:
        raise ValueError("Workbook styles are missing cellXfs.")
    xfs = cell_xfs.findall(q("xf"))
    for cell in header_row.findall(q("c")):
        base_style = int(cell.attrib.get("s", "0"))
        if base_style >= len(xfs):
            raise ValueError(f"Base style {base_style} does not exist.")
        base_font_id = int(xfs[base_style].attrib.get("fontId", "0"))
        header_font_id = ensure_colored_font(styles_root, font_cache, base_font_id, font_rgb)
        cell_fill_id = column_fill_ids.get(column_letters(cell.attrib.get("r", "")), fill_id) if column_fill_ids else fill_id
        cell.set("s", str(ensure_filled_style(styles_root, fill_style_cache, base_style, cell_fill_id, header_font_id)))


def highlight_cells(
    sheet_root: ET.Element,
    styles_root: ET.Element,
    fill_style_cache: dict[tuple[int, int, int | None], int],
    fill_id: int,
    refs: list[str],
) -> None:
    for ref in refs:
        match = re.match(r"([A-Z]+)(\d+)$", ref)
        if not match:
            continue
        row = find_row(sheet_root, int(match.group(2)))
        cell = find_cell(row, ref)
        if cell is None:
            continue
        base_style = int(cell.attrib.get("s", "0"))
        cell.set("s", str(ensure_filled_style(styles_root, fill_style_cache, base_style, fill_id)))


def set_column_widths(root: ET.Element, widths: list[float]) -> None:
    existing = root.find(q("cols"))
    if existing is not None:
        root.remove(existing)

    cols = ET.Element(q("cols"))
    for index, width in enumerate(widths, start=1):
        col = ET.Element(q("col"))
        col.set("min", str(index))
        col.set("max", str(index))
        col.set("width", str(width))
        col.set("customWidth", "1")
        cols.append(col)

    sheet_data = root.find(q("sheetData"))
    if sheet_data is None:
        raise ValueError("Worksheet is missing sheetData.")
    insert_at = list(root).index(sheet_data)
    root.insert(insert_at, cols)


def set_freeze_panes(root: ET.Element, frozen_columns: int = 0, frozen_rows: int = 1) -> None:
    sheet_views = root.find(q("sheetViews"))
    if sheet_views is None:
        raise ValueError("Worksheet is missing sheetViews.")

    sheet_view = sheet_views.find(q("sheetView"))
    if sheet_view is None:
        raise ValueError("Worksheet is missing sheetView.")

    for pane in list(sheet_view.findall(q("pane"))):
        sheet_view.remove(pane)
    for selection in list(sheet_view.findall(q("selection"))):
        sheet_view.remove(selection)
    sheet_view.attrib.pop("topLeftCell", None)

    if frozen_columns <= 0 and frozen_rows <= 0:
        ET.SubElement(sheet_view, q("selection"), {"activeCell": "A1", "sqref": "A1"})
        return

    top_left = cell_ref(max(1, frozen_columns + 1), max(1, frozen_rows + 1))
    pane_attrib = {"topLeftCell": top_left, "state": "frozen"}
    if frozen_columns > 0:
        pane_attrib["xSplit"] = str(frozen_columns)
    if frozen_rows > 0:
        pane_attrib["ySplit"] = str(frozen_rows)

    if frozen_columns > 0 and frozen_rows > 0:
        pane_attrib["activePane"] = "bottomRight"
        ET.SubElement(sheet_view, q("pane"), pane_attrib)
        ET.SubElement(sheet_view, q("selection"), {"pane": "bottomRight", "activeCell": top_left, "sqref": top_left})
        return

    active_pane = "bottomLeft" if frozen_rows > 0 else "topRight"
    pane_attrib["activePane"] = active_pane
    ET.SubElement(sheet_view, q("pane"), pane_attrib)
    ET.SubElement(sheet_view, q("selection"), {"pane": active_pane, "activeCell": top_left, "sqref": top_left})


def sanitize_export_sheet_name(value: str, fallback: str) -> str:
    sanitized = INVALID_SHEET_NAME_RE.sub(" ", normalize_text(value) or fallback)
    sanitized = normalize_text(sanitized).strip().strip("'")
    sanitized = sanitized[:31].rstrip().strip("'")
    return sanitized or fallback


def unique_export_label(label: str, index: int, seen: set[str]) -> str:
    fallback = DEFAULT_EXPORT_SLOT_LABELS[index]
    base = sanitize_export_sheet_name(label, fallback)
    candidate = base
    candidate_key = candidate.casefold()
    if candidate and candidate_key not in seen:
        seen.add(candidate_key)
        return candidate

    base_for_suffix = base
    initial_suffix = " (2)"
    trimmed_initial = base_for_suffix[: max(1, 31 - len(initial_suffix))].rstrip().strip("'") or fallback
    candidate = f"{trimmed_initial}{initial_suffix}"
    attempt = 3
    while candidate.casefold() in seen:
        suffix = f" ({attempt})"
        trimmed = base_for_suffix[: max(1, 31 - len(suffix))].rstrip().strip("'") or fallback
        candidate = f"{trimmed}{suffix}"
        attempt += 1
    seen.add(candidate.casefold())
    return candidate


def resolve_export_slot_labels(episodes: list[dict]) -> list[str]:
    labels: list[str] = []
    seen: set[str] = {label.casefold() for label in STATIC_KOMPLET_COLUMNS}
    for index, episode in enumerate(episodes[:MAX_EPISODES]):
        fallback = DEFAULT_EXPORT_SLOT_LABELS[index]
        raw_label = episode.get("label", fallback)
        normalized = normalize_text(str(raw_label or "")) or fallback
        labels.append(unique_export_label(normalized, index, seen))
    return labels


def update_komplet_table_columns(table_root: ET.Element, episode_labels: list[str]) -> None:
    table_columns = table_root.find(q("tableColumns"))
    if table_columns is None:
        raise ValueError("KOMPLET table is missing tableColumns.")

    columns = list(table_columns.findall(q("tableColumn")))
    if len(columns) < MAX_EPISODES + len(STATIC_KOMPLET_COLUMNS):
        raise ValueError("KOMPLET table does not contain enough columns for díla.")

    selected = [copy.deepcopy(columns[0])]
    for index, _ in enumerate(episode_labels, start=1):
        column = copy.deepcopy(columns[index])
        selected.append(column)
    selected.extend(copy.deepcopy(column) for column in columns[-4:])

    expected_names = [
        "POSTAVA",
        *episode_labels,
        "DABÉR",
        "VSTUPY",
        "REPLIKY",
        "POZNÁMKY",
    ]
    if len(selected) != len(expected_names):
        raise ValueError("KOMPLET table columns do not match expected export layout.")

    table_columns.clear()
    table_columns.set("count", str(len(selected)))
    for index, (column, name) in enumerate(zip(selected, expected_names), start=1):
        column.set("id", str(index))
        column.set("name", name)
        table_columns.append(column)


def make_komplet_header_row(template_row: ET.Element, episode_labels: list[str]) -> ET.Element:
    cells = [inline_cell("A1", "POSTAVA", 2)]
    for index, label in enumerate(episode_labels, start=2):
        cells.append(inline_cell(cell_ref(index, 1), label, 1))

    actor_column = len(episode_labels) + 2
    inputs_column = len(episode_labels) + 3
    replicas_column = len(episode_labels) + 4
    note_column = len(episode_labels) + 5
    cells.extend(
        [
            inline_cell(cell_ref(actor_column, 1), "DABÉR", 18),
            inline_cell(cell_ref(inputs_column, 1), "VSTUPY", 10),
            inline_cell(cell_ref(replicas_column, 1), "REPLIKY", 10),
            inline_cell(cell_ref(note_column, 1), "POZNÁMKY", 10),
        ]
    )
    return make_row(1, note_column, cells, template_row)


def make_episode_header_row(template_row: ET.Element, title: str | None) -> ET.Element:
    cells = [
        inline_cell("A1", "POSTAVA", 17),
        inline_cell("B1", "VSTUPY", 3),
        inline_cell("C1", "REPLIKY", 25),
        inline_cell("D1", "DABÉR", 25),
        inline_cell("E1", "KONTROLA DABÉRA", 3),
        inline_cell("F1", "TEST", 25),
    ]
    if title:
        cells.append(inline_cell("G1", f"DÍLO: {title}", 17))
    return make_row(1, 13, cells, template_row)


def build_actor_export_rows(project: dict) -> list[dict]:
    episodes = list(project.get("episodes", []))
    actor_rows: dict[str, dict[str, object]] = {}
    order: list[str] = []

    for actor in project.get("actors", []):
        actor_name = normalize_text(str(actor.get("actor", "")))
        if not actor_name:
            continue
        actor_rows[actor_name] = {
            "actor": actor_name,
            "totalInputs": int(actor.get("totalInputs", 0)),
            "totalReplicas": int(actor.get("totalReplicas", 0)),
            "note": normalize_text(str(actor.get("note", ""))),
            "episodes": [{"inputs": 0, "replicas": 0} for _ in episodes],
        }
        order.append(actor_name)

    for complete_row in project.get("complete", []):
        actor_name = normalize_text(str(complete_row.get("actor", "")))
        if not actor_name:
            continue

        current = actor_rows.setdefault(
            actor_name,
            {
                "actor": actor_name,
                "totalInputs": int(complete_row.get("totalInputs", 0)),
                "totalReplicas": int(complete_row.get("totalReplicas", 0)),
                "note": "",
                "episodes": [{"inputs": 0, "replicas": 0} for _ in episodes],
            },
        )
        if actor_name not in order:
            order.append(actor_name)

        note = normalize_text(str(complete_row.get("note", "")))
        if note and not current["note"]:
            current["note"] = note

        episode_cells = list(complete_row.get("episodes", []))
        for index, episode in enumerate(episode_cells[: len(episodes)]):
            current["episodes"][index]["inputs"] += int(episode.get("inputs", 0))
            current["episodes"][index]["replicas"] += int(episode.get("replicas", 0))

    return [actor_rows[name] for name in order]


def update_herci_table_columns(table_root: ET.Element, headers: list[str]) -> None:
    table_columns = table_root.find(q("tableColumns"))
    if table_columns is None:
        raise ValueError("HERCI table is missing tableColumns.")

    table_columns.clear()
    table_columns.set("count", str(len(headers)))
    for index, header in enumerate(headers, start=1):
        column = ET.Element(q("tableColumn"))
        column.set("id", str(index))
        column.set("name", header)
        table_columns.append(column)


def make_herci_header_row(template_row: ET.Element, episode_labels: list[str], detailed: bool) -> tuple[ET.Element, list[str]]:
    headers = ["DABÉR"]
    cells = [inline_cell("A1", "DABÉR", 2)]
    column_index = 2

    if detailed:
        for label in episode_labels:
            headers.extend((f"VSTUPY {label}", f"REPLIKY {label}"))
            cells.append(inline_cell(cell_ref(column_index, 1), f"VSTUPY {label}", 22))
            cells.append(inline_cell(cell_ref(column_index + 1, 1), f"REPLIKY {label}", 19))
            column_index += 2

    headers.extend(("VSTUPY CELKEM", "REPLIKY CELKEM", "POZNÁMKY"))
    cells.extend(
        [
            inline_cell(cell_ref(column_index, 1), "VSTUPY CELKEM", 22),
            inline_cell(cell_ref(column_index + 1, 1), "REPLIKY CELKEM", 19),
            inline_cell(cell_ref(column_index + 2, 1), "POZNÁMKY", 21),
        ]
    )
    return make_row(1, column_index + 2, cells, template_row), headers


def build_herci_sheet(
    sheet_root: ET.Element,
    table_root: ET.Element,
    actor_rows: list[dict],
    episode_labels: list[str],
    missing: dict,
    detailed: bool,
) -> None:
    template_rows = sheet_root.find(q("sheetData"))
    if template_rows is None or len(template_rows) < 2:
        raise ValueError("HERCI sheet is missing template rows.")

    header_row, headers = make_herci_header_row(template_rows[0], episode_labels, detailed)
    body_template = template_rows[1]
    rows = [header_row]

    for row_index, actor in enumerate(actor_rows, start=2):
        cells = [inline_cell(f"A{row_index}", str(actor["actor"]), 20)]
        next_column = 2

        if detailed:
            for episode in actor["episodes"]:
                cells.append(number_cell(cell_ref(next_column, row_index), int(episode["inputs"]), 3))
                cells.append(number_cell(cell_ref(next_column + 1, row_index), int(episode["replicas"]), 3))
                next_column += 2

        cells.extend(
            [
                number_cell(cell_ref(next_column, row_index), int(actor["totalInputs"]), 3),
                number_cell(cell_ref(next_column + 1, row_index), int(actor["totalReplicas"]), 3),
                inline_cell(cell_ref(next_column + 2, row_index), str(actor.get("note", "")), 2),
            ]
        )
        row = make_row(row_index, len(headers), cells, body_template)
        set_row_height(row, HERCI_BODY_ROW_HEIGHT)
        rows.append(row)

    status_row_index = len(rows) + 1
    status_text = (
        f"JEŠTĚ CHYBÍ DOPLNIT {missing['characters']} POSTAV, "
        f"{missing['inputs']} VSTUPŮ A {missing['replicas']} REPLIK."
    )
    status_cells = [inline_cell(f"A{status_row_index}", status_text, 20)]
    for column_index in range(2, len(headers)):
        status_cells.append(number_cell(cell_ref(column_index, status_row_index), 0, 3))
    status_cells.append(blank_cell(cell_ref(len(headers), status_row_index), 2))
    status_row = make_row(status_row_index, len(headers), status_cells, body_template)
    set_row_height(status_row, HERCI_BODY_ROW_HEIGHT)
    rows.append(status_row)

    end_column = column_name(len(headers))
    reset_sheet_data(sheet_root, rows, f"A1:{end_column}{status_row_index}")
    update_herci_table_columns(table_root, headers)
    update_table(table_root, f"A1:{end_column}{status_row_index}")


def build_komplet_sheet(sheet_root: ET.Element, table_root: ET.Element, complete_rows: list[dict], episode_labels: list[str]) -> None:
    template_rows = sheet_root.find(q("sheetData"))
    if template_rows is None or len(template_rows) < 2:
        raise ValueError("KOMPLET sheet is missing template rows.")

    episode_count = max(1, len(episode_labels))
    header_row = make_komplet_header_row(template_rows[0], episode_labels)
    body_template = template_rows[1]
    rows = [header_row]

    if not complete_rows:
        complete_rows = [
            {
                "character": "",
                "actor": "",
                "note": "",
                "totalInputs": 0,
                "totalReplicas": 0,
                "episodes": [{"display": "0 / 0"} for _ in range(episode_count)],
            }
        ]

    for row_index, row_data in enumerate(complete_rows, start=2):
        episodes = list(row_data["episodes"][:episode_count]) + [{"display": "0 / 0"}] * max(
            0, episode_count - len(row_data["episodes"])
        )
        cells = [inline_cell(f"A{row_index}", row_data["character"], 25)]
        for column_offset, episode in enumerate(episodes, start=2):
            cells.append(inline_cell(cell_ref(column_offset, row_index), episode["display"], 9))

        actor_column = episode_count + 2
        inputs_column = episode_count + 3
        replicas_column = episode_count + 4
        note_column = episode_count + 5
        cells.extend(
            [
                inline_cell(cell_ref(actor_column, row_index), row_data.get("actor", ""), 27),
                number_cell(cell_ref(inputs_column, row_index), int(row_data["totalInputs"]), 10),
                number_cell(cell_ref(replicas_column, row_index), int(row_data["totalReplicas"]), 10),
                inline_cell(cell_ref(note_column, row_index), row_data.get("note", ""), 1),
            ]
        )
        rows.append(make_row(row_index, note_column, cells, body_template))

    end_column = column_name(episode_count + 5)
    reset_sheet_data(sheet_root, rows, f"A1:{end_column}{len(rows)}")
    update_komplet_table_columns(table_root, episode_labels)
    update_table(table_root, f"A1:{end_column}{len(rows)}")


def build_episode_sheet(sheet_root: ET.Element, table_root: ET.Element, episode: dict | None, title: str | None = None) -> None:
    template_rows = sheet_root.find(q("sheetData"))
    if template_rows is None or len(template_rows) < 2:
        raise ValueError("Episode sheet is missing template rows.")

    header_row = make_episode_header_row(template_rows[0], title)
    body_template = template_rows[1]
    rows = [header_row]
    characters = list((episode or {}).get("characters", []))

    if not characters:
        blank_row = make_row(
            2,
            13,
            [
                blank_cell("A2", 25),
                blank_cell("B2", 25),
                blank_cell("C2", 25),
                blank_cell("D2", 14),
                blank_cell("E2", 14),
                blank_cell("F2", 14),
                blank_cell("G2", 14),
                blank_cell("H2", 14),
                blank_cell("I2", 14),
                blank_cell("J2", 14),
                blank_cell("K2", 14),
                blank_cell("L2", 14),
                blank_cell("M2", 14),
            ],
            body_template,
        )
        rows.append(blank_row)
        reset_sheet_data(sheet_root, rows, "A1:M2")
        update_table(table_root, "A1:F2")
        return

    for row_index, row_data in enumerate(characters, start=2):
        actor = row_data.get("actor", "")
        cells = [
            inline_cell(f"A{row_index}", row_data["character"], 25),
            number_cell(f"B{row_index}", int(row_data["inputs"]), 25),
            number_cell(f"C{row_index}", int(row_data["replicas"]), 25),
            inline_cell(f"D{row_index}", actor, 14),
            blank_cell(f"E{row_index}", 14),
            blank_cell(f"F{row_index}", 14),
            blank_cell(f"G{row_index}", 14),
            blank_cell(f"H{row_index}", 14),
            blank_cell(f"I{row_index}", 14),
            blank_cell(f"J{row_index}", 14),
            blank_cell(f"K{row_index}", 14),
            blank_cell(f"L{row_index}", 14),
            blank_cell(f"M{row_index}", 14),
        ]
        rows.append(make_row(row_index, 13, cells, body_template))

    reset_sheet_data(sheet_root, rows, f"A1:M{len(rows)}")
    update_table(table_root, f"A1:F{len(rows)}")


def build_herci_column_widths(actor_rows: list[dict], episode_labels: list[str], detailed: bool) -> list[float]:
    widths = [suggested_width(["DABÉR", *(str(row.get("actor", "")) for row in actor_rows)], 18, 28)]

    if detailed:
        for label in episode_labels:
            widths.append(suggested_width([f"VSTUPY {label}"], 12, 18))
            widths.append(suggested_width([f"REPLIKY {label}"], 12, 18))

    widths.append(14.0)
    widths.append(14.0)
    widths.append(suggested_width(["POZNÁMKY", *(str(row.get("note", "")) for row in actor_rows)], 18, 30))
    return widths


def build_komplet_column_widths(complete_rows: list[dict], episode_labels: list[str]) -> list[float]:
    widths = [
        suggested_width(["POSTAVA", *(str(row.get("character", "")) for row in complete_rows)], 18, 34),
    ]
    for index, label in enumerate(episode_labels):
        display_values = [str(row.get("episodes", [])[index].get("display", "")) for row in complete_rows if index < len(row.get("episodes", []))]
        widths.append(suggested_width([label, *display_values], 12, 18))
    widths.extend(
        [
            suggested_width(["DABÉR", *(str(row.get("actor", "")) for row in complete_rows)], 16, 24),
            12.0,
            12.0,
            suggested_width(["POZNÁMKY", *(str(row.get("note", "")) for row in complete_rows)], 18, 34),
        ]
    )
    return widths


def build_episode_column_widths(episode: dict | None, title: str | None) -> list[float]:
    characters = list((episode or {}).get("characters", []))
    return [
        suggested_width(["POSTAVA", *(str(item.get("character", "")) for item in characters)], 18, 30),
        12.0,
        12.0,
        suggested_width(["DABÉR", *(str(item.get("actor", "")) for item in characters)], 16, 24),
        18.0,
        12.0,
        suggested_width([f"DÍLO: {title}" if title else ""], 16, 28),
    ]


def build_komplet_header_fill_map(
    episode_labels: list[str],
    primary_fill_id: int,
    episode_fill_id: int,
    summary_fill_id: int,
    note_fill_id: int,
) -> dict[str, int]:
    fill_map = {"A": primary_fill_id}
    for column_index in range(2, len(episode_labels) + 2):
        fill_map[column_name(column_index)] = episode_fill_id

    actor_column = len(episode_labels) + 2
    inputs_column = len(episode_labels) + 3
    replicas_column = len(episode_labels) + 4
    note_column = len(episode_labels) + 5
    fill_map[column_name(actor_column)] = primary_fill_id
    fill_map[column_name(inputs_column)] = summary_fill_id
    fill_map[column_name(replicas_column)] = summary_fill_id
    fill_map[column_name(note_column)] = note_fill_id
    return fill_map


def apply_herci_visual_polish(
    sheet_root: ET.Element,
    styles_root: ET.Element,
    fill_style_cache: dict[tuple[int, int, int | None], int],
    font_cache: dict[tuple[int, str], int],
    header_fill_id: int,
    actor_rows: list[dict],
    episode_labels: list[str],
    detailed: bool,
) -> None:
    apply_header_fill(
        sheet_root,
        styles_root,
        fill_style_cache,
        font_cache,
        header_fill_id,
        HEADER_FONT_RGB,
        row_height=HERCI_HEADER_ROW_HEIGHT,
    )
    set_column_widths(sheet_root, build_herci_column_widths(actor_rows, episode_labels, detailed))
    set_freeze_panes(sheet_root, frozen_rows=1)


def apply_komplet_visual_polish(
    sheet_root: ET.Element,
    styles_root: ET.Element,
    fill_style_cache: dict[tuple[int, int, int | None], int],
    font_cache: dict[tuple[int, str], int],
    header_fill_ids: dict[str, int],
    missing_fill_id: int,
    complete_rows: list[dict],
    episode_labels: list[str],
) -> None:
    apply_header_fill(
        sheet_root,
        styles_root,
        fill_style_cache,
        font_cache,
        header_fill_ids["default"],
        HEADER_FONT_RGB,
        build_komplet_header_fill_map(
            episode_labels,
            header_fill_ids["primary"],
            header_fill_ids["episode"],
            header_fill_ids["summary"],
            header_fill_ids["note"],
        ),
    )
    set_column_widths(sheet_root, build_komplet_column_widths(complete_rows, episode_labels))
    set_freeze_panes(sheet_root, frozen_columns=1, frozen_rows=1)

    actor_column_index = len(episode_labels) + 2
    for row_offset, row_data in enumerate(complete_rows, start=2):
        if normalize_text(str(row_data.get("actor", ""))):
            continue
        if not normalize_text(str(row_data.get("character", ""))):
            continue
        highlight_cells(
            sheet_root,
            styles_root,
            fill_style_cache,
            missing_fill_id,
            [cell_ref(1, row_offset), cell_ref(actor_column_index, row_offset)],
        )


def apply_episode_visual_polish(
    sheet_root: ET.Element,
    styles_root: ET.Element,
    fill_style_cache: dict[tuple[int, int, int | None], int],
    font_cache: dict[tuple[int, str], int],
    header_fill_id: int,
    missing_fill_id: int,
    episode: dict | None,
    title: str | None,
) -> None:
    apply_header_fill(sheet_root, styles_root, fill_style_cache, font_cache, header_fill_id, HEADER_FONT_RGB)
    set_column_widths(sheet_root, build_episode_column_widths(episode, title))
    set_freeze_panes(sheet_root, frozen_rows=1)

    characters = list((episode or {}).get("characters", []))
    for row_offset, character in enumerate(characters, start=2):
        if normalize_text(str(character.get("actor", ""))):
            continue
        if not normalize_text(str(character.get("character", ""))):
            continue
        highlight_cells(
            sheet_root,
            styles_root,
            fill_style_cache,
            missing_fill_id,
            [f"A{row_offset}", f"D{row_offset}"],
        )


def update_workbook_sheets(files: dict[str, bytes], episode_labels: list[str]) -> None:
    workbook_root = ET.fromstring(files["xl/workbook.xml"])
    workbook_rels_root = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    relationship_map = {
        relationship.attrib["Id"]: relationship
        for relationship in workbook_rels_root.findall(f"{{{PKG_NS}}}Relationship")
    }

    keep_sheet_paths = {
        SHEET_PATHS["HERCI"],
        SHEET_PATHS["KOMPLET"],
        *(SHEET_PATHS[key] for key in EPISODE_SHEET_KEYS[: len(episode_labels)]),
    }
    episode_name_map = {
        SHEET_PATHS[key]: episode_labels[index]
        for index, key in enumerate(EPISODE_SHEET_KEYS[: len(episode_labels)])
    }

    sheets_element = workbook_root.find(q("sheets"))
    if sheets_element is None:
        raise ValueError("Workbook is missing sheets.")

    for sheet in list(sheets_element.findall(q("sheet"))):
        relationship_id = sheet.attrib.get(f"{{{R_NS}}}id")
        relationship = relationship_map.get(relationship_id or "")
        if relationship is None:
            sheets_element.remove(sheet)
            continue

        target = relationship.attrib.get("Target", "")
        if not target.startswith("xl/"):
            target = f"xl/{target.lstrip('/')}"

        if target not in keep_sheet_paths:
            sheets_element.remove(sheet)
            workbook_rels_root.remove(relationship)
            continue

        if target in episode_name_map:
            sheet.set("name", episode_name_map[target])

    files["xl/workbook.xml"] = serialize_xml(workbook_root)
    files["xl/_rels/workbook.xml.rels"] = serialize_plain_xml(workbook_rels_root)


def update_app_properties(files: dict[str, bytes], episode_labels: list[str]) -> None:
    app_root = ET.fromstring(files["docProps/app.xml"])
    total_sheet_names = [*STATIC_SHEET_NAMES, *episode_labels]
    sheet_count = len(total_sheet_names)

    heading_pairs = app_root.find(aq("HeadingPairs"))
    if heading_pairs is not None:
        vector = heading_pairs.find(vq("vector"))
        if vector is not None:
            vector.set("size", "2")
            variants = vector.findall(vq("variant"))
            if len(variants) >= 2:
                count_value = variants[1].find(vq("i4"))
                if count_value is not None:
                    count_value.text = str(sheet_count)

    titles_of_parts = app_root.find(aq("TitlesOfParts"))
    if titles_of_parts is not None:
        vector = titles_of_parts.find(vq("vector"))
        if vector is not None:
            vector.clear()
            vector.set("size", str(sheet_count))
            vector.set("baseType", "lpstr")
            for sheet_name in total_sheet_names:
                part = ET.SubElement(vector, vq("lpstr"))
                part.text = sheet_name

    files["docProps/app.xml"] = serialize_plain_xml(app_root)


def prune_unused_export_parts(files: dict[str, bytes], episode_labels: list[str]) -> None:
    removable_static_sheet_paths = {SHEET_PATHS["PŘEHLED"]}
    removable_sheet_paths = {
        SHEET_PATHS[key]
        for key in EPISODE_SHEET_KEYS[len(episode_labels) :]
    }
    removable_sheet_paths.update(removable_static_sheet_paths)
    removable_table_paths = {
        TABLE_PATHS[key]
        for key in EPISODE_SHEET_KEYS[len(episode_labels) :]
    }

    for path in removable_sheet_paths | removable_table_paths:
        files.pop(path, None)

    removable_sheet_rels = {
        f"xl/worksheets/_rels/{Path(path).name}.rels"
        for path in removable_sheet_paths
    }
    for path in removable_sheet_rels:
        files.pop(path, None)

    types_root = ET.fromstring(files["[Content_Types].xml"])
    for override in list(types_root):
        part_name = override.attrib.get("PartName", "")
        normalized = part_name.lstrip("/")
        if normalized in removable_sheet_paths or normalized in removable_table_paths:
            types_root.remove(override)
    files["[Content_Types].xml"] = serialize_plain_xml(types_root)


def remove_calc_chain(files: dict[str, bytes]) -> None:
    files.pop("xl/calcChain.xml", None)

    rels_path = "xl/_rels/workbook.xml.rels"
    rels_root = ET.fromstring(files[rels_path])
    for relationship in list(rels_root):
        if relationship.attrib.get("Type", "").endswith("/calcChain"):
            rels_root.remove(relationship)
    files[rels_path] = serialize_xml(rels_root)

    types_path = "[Content_Types].xml"
    types_root = ET.fromstring(files[types_path])
    for override in list(types_root):
        if override.attrib.get("PartName") == "/xl/calcChain.xml":
            types_root.remove(override)
    files[types_path] = serialize_xml(types_root)


def remove_pivot_artifacts(files: dict[str, bytes]) -> None:
    pivot_prefixes = ("xl/pivotCache/", "xl/pivotTables/")
    for name in [name for name in files if name.startswith(pivot_prefixes)]:
        files.pop(name, None)

    workbook_path = "xl/workbook.xml"
    workbook_root = ET.fromstring(files[workbook_path])
    pivot_caches = workbook_root.find(q("pivotCaches"))
    if pivot_caches is not None:
        workbook_root.remove(pivot_caches)
    files[workbook_path] = serialize_xml(workbook_root)

    rels_path = "xl/_rels/workbook.xml.rels"
    rels_root = ET.fromstring(files[rels_path])
    for relationship in list(rels_root):
        if relationship.attrib.get("Type", "").endswith("/pivotCacheDefinition"):
            rels_root.remove(relationship)
    files[rels_path] = serialize_xml(rels_root)

    sheet3_rels_path = "xl/worksheets/_rels/sheet3.xml.rels"
    if sheet3_rels_path in files:
        sheet3_rels_root = ET.fromstring(files[sheet3_rels_path])
        for relationship in list(sheet3_rels_root):
            if relationship.attrib.get("Type", "").endswith("/pivotTable"):
                sheet3_rels_root.remove(relationship)
        if len(sheet3_rels_root) == 0:
            files.pop(sheet3_rels_path, None)
        else:
            files[sheet3_rels_path] = serialize_xml(sheet3_rels_root)

    types_path = "[Content_Types].xml"
    types_root = ET.fromstring(files[types_path])
    for override in list(types_root):
        part_name = override.attrib.get("PartName", "")
        if part_name.startswith("/xl/pivotCache/") or part_name.startswith("/xl/pivotTables/"):
            types_root.remove(override)
    files[types_path] = serialize_xml(types_root)


def export_project_workbook(
    project: dict,
    template_path: Path | None = None,
    herci_by_episode: bool = False,
) -> bytes:
    source_path = Path(template_path or TEMPLATE_PATH)
    with ZipFile(source_path) as template:
        files = {name: template.read(name) for name in template.namelist()}

    styles_root = ET.fromstring(files["xl/styles.xml"])
    header_fill_id = ensure_solid_fill(styles_root, HEADER_FILL_RGB)
    komplet_header_fill_ids = {
        "default": header_fill_id,
        "primary": ensure_solid_fill(styles_root, KOMPLET_HEADER_PRIMARY_FILL_RGB),
        "episode": ensure_solid_fill(styles_root, KOMPLET_HEADER_EPISODE_FILL_RGB),
        "summary": ensure_solid_fill(styles_root, KOMPLET_HEADER_SUMMARY_FILL_RGB),
        "note": ensure_solid_fill(styles_root, KOMPLET_HEADER_NOTE_FILL_RGB),
    }
    missing_fill_id = ensure_solid_fill(styles_root, MISSING_FILL_RGB)
    fill_style_cache: dict[tuple[int, int, int | None], int] = {}
    font_cache: dict[tuple[int, str], int] = {}

    sheet_roots = {name: ET.fromstring(files[path]) for name, path in SHEET_PATHS.items()}
    table_roots = {name: ET.fromstring(files[path]) for name, path in TABLE_PATHS.items()}

    episodes = list(project.get("episodes", []))
    slot_labels = resolve_export_slot_labels(episodes)
    actor_rows = build_actor_export_rows(project)
    for episode in episodes:
        assignment_lookup = {row["character"]: row.get("actor", "") for row in project.get("complete", [])}
        for character in episode.get("characters", []):
            character["actor"] = assignment_lookup.get(character["character"], "")

    build_herci_sheet(
        sheet_roots["HERCI"],
        table_roots["HERCI"],
        actor_rows,
        slot_labels,
        dict(project.get("missing", {})),
        herci_by_episode,
    )
    apply_herci_visual_polish(
        sheet_roots["HERCI"],
        styles_root,
        fill_style_cache,
        font_cache,
        header_fill_id,
        actor_rows,
        slot_labels,
        herci_by_episode,
    )
    build_komplet_sheet(
        sheet_roots["KOMPLET"],
        table_roots["KOMPLET"],
        list(project.get("complete", [])),
        slot_labels,
    )
    apply_komplet_visual_polish(
        sheet_roots["KOMPLET"],
        styles_root,
        fill_style_cache,
        font_cache,
        komplet_header_fill_ids,
        missing_fill_id,
        list(project.get("complete", [])),
        slot_labels,
    )
    for slot, sheet_name in enumerate(["01", "02", "03", "04", "05", "06"]):
        episode = episodes[slot] if slot < len(episodes) else None
        title = slot_labels[slot] if slot < len(episodes) else None
        build_episode_sheet(sheet_roots[sheet_name], table_roots[sheet_name], episode, title)
        if slot < len(episodes):
            apply_episode_visual_polish(
                sheet_roots[sheet_name],
                styles_root,
                fill_style_cache,
                font_cache,
                header_fill_id,
                missing_fill_id,
                episode,
                title,
            )

    active_sheet_keys = ["HERCI", "KOMPLET", *EPISODE_SHEET_KEYS[: len(episodes)]]
    active_table_keys = ["HERCI", "KOMPLET", *EPISODE_SHEET_KEYS[: len(episodes)]]

    for name in active_sheet_keys:
        root = sheet_roots[name]
        files[SHEET_PATHS[name]] = serialize_xml(root)
    for name in active_table_keys:
        root = table_roots[name]
        files[TABLE_PATHS[name]] = serialize_xml(root)

    update_workbook_sheets(files, slot_labels)
    update_app_properties(files, slot_labels)
    prune_unused_export_parts(files, slot_labels)
    remove_pivot_artifacts(files)
    remove_calc_chain(files)
    files["xl/styles.xml"] = serialize_xml(styles_root)

    output = io.BytesIO()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as workbook:
        for name, content in files.items():
            workbook.writestr(name, content)
    return output.getvalue()
