from __future__ import annotations

import io
import posixpath
import re
import unittest
import xml.etree.ElementTree as ET
from pathlib import PurePosixPath
from zipfile import ZipFile

from obsazovani.core import build_project
from obsazovani.exporter import (
    HEADER_FILL_RGB,
    HEADER_FONT_RGB,
    KOMPLET_HEADER_EPISODE_FILL_RGB,
    KOMPLET_HEADER_NOTE_FILL_RGB,
    KOMPLET_HEADER_PRIMARY_FILL_RGB,
    KOMPLET_HEADER_SUMMARY_FILL_RGB,
    export_project_workbook,
)

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


def resolve_relationship_target(rels_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")

    rels_parts = PurePosixPath(rels_path).parts
    source_parts = rels_parts[:-1]
    if source_parts and source_parts[-1] == "_rels":
        source_parts = source_parts[:-1]
    base_dir = posixpath.join(*source_parts) if source_parts else ""
    return posixpath.normpath(posixpath.join(base_dir, target))


def find_cell_text(sheet_root: ET.Element, cell_ref: str) -> str | None:
    cell = sheet_root.find(f".//{{{MAIN_NS}}}c[@r='{cell_ref}']")
    if cell is None:
        return None

    inline_text = cell.find(f"{{{MAIN_NS}}}is/{{{MAIN_NS}}}t")
    if inline_text is not None:
        return inline_text.text or ""

    value = cell.find(f"{{{MAIN_NS}}}v")
    if value is not None:
        return value.text or ""
    return None


def table_column_names(table_root: ET.Element) -> list[str]:
    return [
        str(column.attrib.get("name", ""))
        for column in table_root.findall(f"{{{MAIN_NS}}}tableColumns/{{{MAIN_NS}}}tableColumn")
    ]


def table_column_ids(table_root: ET.Element) -> list[int]:
    return [
        int(column.attrib.get("id", "0"))
        for column in table_root.findall(f"{{{MAIN_NS}}}tableColumns/{{{MAIN_NS}}}tableColumn")
    ]


def workbook_sheet_names(archive: ZipFile) -> list[str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    return [
        str(sheet.attrib.get("name", ""))
        for sheet in workbook_root.findall(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet")
    ]


def workbook_sheet_targets(archive: ZipFile) -> list[str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
    }
    targets: list[str] = []
    for sheet in workbook_root.findall(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet"):
        relationship_id = sheet.attrib.get(f"{{{R_NS}}}id", "")
        targets.append(rel_map[relationship_id])
    return targets


def app_titles_of_parts(archive: ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("docProps/app.xml"))
    vector = root.find(f"{{{APP_NS}}}TitlesOfParts/{{{VT_NS}}}vector")
    if vector is None:
        return []
    return [str(item.text or "") for item in vector.findall(f"{{{VT_NS}}}lpstr")]


def table_formula_counts(table_root: ET.Element) -> tuple[int, int]:
    calculated = len(table_root.findall(f".//{{{MAIN_NS}}}calculatedColumnFormula"))
    totals = len(table_root.findall(f".//{{{MAIN_NS}}}totalsRowFormula"))
    return calculated, totals


def find_cell_style(sheet_root: ET.Element, cell_ref: str) -> int | None:
    cell = sheet_root.find(f".//{{{MAIN_NS}}}c[@r='{cell_ref}']")
    if cell is None or "s" not in cell.attrib:
        return None
    return int(cell.attrib["s"])


def find_fill_rgb(styles_root: ET.Element, style_id: int | None) -> str | None:
    if style_id is None:
        return None

    cell_xfs = styles_root.find(f"{{{MAIN_NS}}}cellXfs")
    fills = styles_root.find(f"{{{MAIN_NS}}}fills")
    if cell_xfs is None or fills is None:
        return None

    xfs = cell_xfs.findall(f"{{{MAIN_NS}}}xf")
    if style_id >= len(xfs):
        return None
    fill_id = int(xfs[style_id].attrib.get("fillId", "0"))
    fill_list = fills.findall(f"{{{MAIN_NS}}}fill")
    if fill_id >= len(fill_list):
        return None
    pattern = fill_list[fill_id].find(f"{{{MAIN_NS}}}patternFill")
    if pattern is None:
        return None
    fg_color = pattern.find(f"{{{MAIN_NS}}}fgColor")
    if fg_color is None:
        return None
    return fg_color.attrib.get("rgb")


def find_font_rgb(styles_root: ET.Element, style_id: int | None) -> str | None:
    if style_id is None:
        return None

    cell_xfs = styles_root.find(f"{{{MAIN_NS}}}cellXfs")
    fonts = styles_root.find(f"{{{MAIN_NS}}}fonts")
    if cell_xfs is None or fonts is None:
        return None

    xfs = cell_xfs.findall(f"{{{MAIN_NS}}}xf")
    if style_id >= len(xfs):
        return None
    font_id = int(xfs[style_id].attrib.get("fontId", "0"))
    font_list = fonts.findall(f"{{{MAIN_NS}}}font")
    if font_id >= len(font_list):
        return None
    color = font_list[font_id].find(f"{{{MAIN_NS}}}color")
    if color is None:
        return None
    return color.attrib.get("rgb")


def pane_attributes(sheet_root: ET.Element) -> dict[str, str]:
    pane = sheet_root.find(f"{{{MAIN_NS}}}sheetViews/{{{MAIN_NS}}}sheetView/{{{MAIN_NS}}}pane")
    return dict(pane.attrib) if pane is not None else {}


def sheet_view_attributes(sheet_root: ET.Element) -> dict[str, str]:
    sheet_view = sheet_root.find(f"{{{MAIN_NS}}}sheetViews/{{{MAIN_NS}}}sheetView")
    return dict(sheet_view.attrib) if sheet_view is not None else {}


def row_attributes(sheet_root: ET.Element, row_index: int) -> dict[str, str]:
    row = sheet_root.find(f"{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row[@r='{row_index}']")
    return dict(row.attrib) if row is not None else {}


def column_width(sheet_root: ET.Element, column_index: int) -> float | None:
    cols = sheet_root.find(f"{{{MAIN_NS}}}cols")
    if cols is None:
        return None
    for col in cols.findall(f"{{{MAIN_NS}}}col"):
        minimum = int(col.attrib.get("min", "0"))
        maximum = int(col.attrib.get("max", "0"))
        if minimum <= column_index <= maximum:
            return float(col.attrib.get("width", "0"))
    return None


def declared_prefixes(xml_bytes: bytes) -> set[str]:
    return set(re.findall(rb"\bxmlns:([A-Za-z0-9_]+)=", xml_bytes))


def referenced_compatibility_prefixes(xml_bytes: bytes) -> set[str]:
    text = xml_bytes.decode("utf-8")
    prefixes: set[str] = set()
    for attribute in ("Ignorable", "Requires"):
        for _, raw_value in re.findall(rf'{attribute}=(["\'])(.*?)\1', text):
            prefixes.update(token for token in raw_value.split() if ":" not in token)
    return prefixes


class ExporterWorkbookTests(unittest.TestCase):
    def assert_workbook_structurally_consistent(self, workbook: bytes) -> None:
        with ZipFile(io.BytesIO(workbook)) as archive:
            names = set(archive.namelist())

            self.assertNotIn("xl/calcChain.xml", names)
            self.assertFalse(any(name.startswith("xl/pivotCache/") for name in names))
            self.assertFalse(any(name.startswith("xl/pivotTables/") for name in names))

            for name in names:
                if name.endswith(".xml") or name.endswith(".rels"):
                    ET.fromstring(archive.read(name))

            workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
            self.assertIsNone(
                workbook_root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}pivotCaches")
            )

            for name in names:
                if not name.endswith(".xml"):
                    continue
                raw_xml = archive.read(name)
                required_prefixes = referenced_compatibility_prefixes(raw_xml)
                if not required_prefixes:
                    continue
                self.assertTrue(
                    required_prefixes.issubset({prefix.decode("utf-8") for prefix in declared_prefixes(raw_xml)}),
                    msg=f"Undeclared compatibility prefixes in {name}",
                )

            for table_name in [name for name in names if name.startswith("xl/tables/") and name.endswith(".xml")]:
                table_root = ET.fromstring(archive.read(table_name))
                calculated_count, totals_count = table_formula_counts(table_root)
                self.assertEqual(calculated_count, 0, msg=f"Static export still contains calculated columns in {table_name}")
                self.assertEqual(totals_count, 0, msg=f"Static export still contains totals formulas in {table_name}")

            for rels_name in [name for name in names if name.endswith(".rels")]:
                rels_root = ET.fromstring(archive.read(rels_name))
                for relationship in rels_root:
                    target_mode = relationship.attrib.get("TargetMode")
                    if target_mode == "External":
                        continue
                    target = resolve_relationship_target(rels_name, relationship.attrib["Target"])
                    self.assertIn(target, names, msg=f"Missing OOXML target: {rels_name} -> {target}")

    def test_exported_workbook_is_structurally_consistent(self) -> None:
        project = build_project(
            {
                "title": "Export Validation",
                "episodes": [
                    {
                        "label": "01",
                        "content": (
                            "POSTAVA\tTC\tTEXT\n"
                            "ALFA\t00:00:01\tAhoj světe\n"
                            "BETA\t00:00:02\tJedna dvě tři čtyři pět šest sedm osm devět\n"
                        ),
                    }
                ],
                "assignments": {
                    "ALFA": {"actor": "Herec A", "note": "Poznámka"},
                    "BETA": {"actor": "Herec B", "note": ""},
                },
            }
        )

        workbook = export_project_workbook(project)
        self.assert_workbook_structurally_consistent(workbook)

    def test_exported_workbook_only_contains_active_work_sheets_for_two_works(self) -> None:
        project = build_project(
            {
                "title": "Two Works",
                "episodes": [
                    {"label": "Pilot", "content": "POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tPrvní replika\n"},
                    {"label": "Finále", "content": "POSTAVA\tTC\tTEXT\nBETA\t00:00:02\tDruhá replika\n"},
                ],
                "assignments": {},
            }
        )

        workbook = export_project_workbook(project)
        with ZipFile(io.BytesIO(workbook)) as archive:
            names = set(archive.namelist())
            komplet_root = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
            komplet_table_root = ET.fromstring(archive.read("xl/tables/table2.xml"))
            first_work_root = ET.fromstring(archive.read("xl/worksheets/sheet4.xml"))
            second_work_root = ET.fromstring(archive.read("xl/worksheets/sheet5.xml"))

            self.assertEqual(workbook_sheet_names(archive), ["HERCI", "KOMPLET", "Pilot", "Finále"])
            self.assertEqual(
                workbook_sheet_targets(archive),
                [
                    "worksheets/sheet1.xml",
                    "worksheets/sheet2.xml",
                    "worksheets/sheet4.xml",
                    "worksheets/sheet5.xml",
                ],
            )
            self.assertEqual(app_titles_of_parts(archive), ["HERCI", "KOMPLET", "Pilot", "Finále"])

        self.assertEqual(find_cell_text(komplet_root, "B1"), "Pilot")
        self.assertEqual(find_cell_text(komplet_root, "C1"), "Finále")
        self.assertEqual(find_cell_text(komplet_root, "D1"), "DABÉR")
        self.assertEqual(find_cell_text(komplet_root, "A2"), "ALFA")
        self.assertEqual(find_cell_text(komplet_root, "A3"), "BETA")
        self.assertEqual(
            table_column_names(komplet_table_root),
            ["POSTAVA", "Pilot", "Finále", "DABÉR", "VSTUPY", "REPLIKY", "POZNÁMKY"],
        )
        self.assertEqual(table_column_ids(komplet_table_root), list(range(1, 8)))
        self.assertEqual(komplet_table_root.attrib.get("ref"), "A1:G3")
        self.assertEqual(komplet_root.find(f"{{{MAIN_NS}}}dimension").attrib.get("ref"), "A1:G3")

        self.assertEqual(find_cell_text(first_work_root, "G1"), "DÍLO: Pilot")
        self.assertEqual(find_cell_text(second_work_root, "G1"), "DÍLO: Finále")
        self.assertNotIn("xl/worksheets/sheet3.xml", names)
        self.assertNotIn("xl/worksheets/_rels/sheet3.xml.rels", names)
        self.assertNotIn("xl/worksheets/sheet6.xml", names)
        self.assertNotIn("xl/worksheets/sheet7.xml", names)
        self.assertNotIn("xl/worksheets/sheet8.xml", names)
        self.assertNotIn("xl/worksheets/sheet9.xml", names)
        self.assertNotIn("xl/tables/table5.xml", names)
        self.assertNotIn("xl/tables/table6.xml", names)
        self.assertNotIn("xl/tables/table7.xml", names)
        self.assertNotIn("xl/tables/table8.xml", names)

    def test_exported_workbook_sanitizes_and_deduplicates_sheet_names(self) -> None:
        project = build_project(
            {
                "title": "Sanitized Labels",
                "episodes": [
                    {"label": "Pilot/Finale", "content": "POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tPrvní replika\n"},
                    {"label": "Pilot?Finale", "content": "POSTAVA\tTC\tTEXT\nBETA\t00:00:02\tDruhá replika\n"},
                    {"label": "DABÉR", "content": ""},
                ],
                "assignments": {},
            }
        )

        workbook = export_project_workbook(project)
        with ZipFile(io.BytesIO(workbook)) as archive:
            komplet_table_root = ET.fromstring(archive.read("xl/tables/table2.xml"))
            sheet_names = workbook_sheet_names(archive)

        self.assertEqual(sheet_names, ["HERCI", "KOMPLET", "Pilot Finale", "Pilot Finale (2)", "DABÉR (2)"])
        self.assertEqual(
            table_column_names(komplet_table_root),
            ["POSTAVA", "Pilot Finale", "Pilot Finale (2)", "DABÉR (2)", "DABÉR", "VSTUPY", "REPLIKY", "POZNÁMKY"],
        )
        self.assertEqual(table_column_ids(komplet_table_root), list(range(1, 9)))

    def test_exported_workbook_handles_one_work(self) -> None:
        project = build_project(
            {
                "title": "One Work",
                "episodes": [
                    {"label": "Pilot", "content": "POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tPrvní replika\n"},
                ],
                "assignments": {},
            }
        )

        workbook = export_project_workbook(project)
        with ZipFile(io.BytesIO(workbook)) as archive:
            komplet_root = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
            komplet_table_root = ET.fromstring(archive.read("xl/tables/table2.xml"))

            self.assertEqual(workbook_sheet_names(archive), ["HERCI", "KOMPLET", "Pilot"])

        self.assertEqual(find_cell_text(komplet_root, "B1"), "Pilot")
        self.assertEqual(find_cell_text(komplet_root, "C1"), "DABÉR")
        self.assertEqual(find_cell_text(komplet_root, "A2"), "ALFA")
        self.assertEqual(table_column_names(komplet_table_root), ["POSTAVA", "Pilot", "DABÉR", "VSTUPY", "REPLIKY", "POZNÁMKY"])
        self.assertEqual(table_column_ids(komplet_table_root), list(range(1, 7)))
        self.assertEqual(komplet_table_root.attrib.get("ref"), "A1:F2")

    def test_exported_workbook_handles_six_works(self) -> None:
        project = build_project(
            {
                "title": "Six Works",
                "episodes": [
                    {"label": f"Dílo {index + 1}", "content": f"POSTAVA\tTC\tTEXT\nPOSTAVA{index + 1}\t00:00:01\tText\n"}
                    for index in range(6)
                ],
                "assignments": {},
            }
        )

        workbook = export_project_workbook(project)
        with ZipFile(io.BytesIO(workbook)) as archive:
            komplet_table_root = ET.fromstring(archive.read("xl/tables/table2.xml"))
            self.assertEqual(
                workbook_sheet_names(archive),
                ["HERCI", "KOMPLET", "Dílo 1", "Dílo 2", "Dílo 3", "Dílo 4", "Dílo 5", "Dílo 6"],
            )
            self.assertIn("xl/worksheets/sheet9.xml", archive.namelist())

        self.assertEqual(
            table_column_names(komplet_table_root),
            ["POSTAVA", "Dílo 1", "Dílo 2", "Dílo 3", "Dílo 4", "Dílo 5", "Dílo 6", "DABÉR", "VSTUPY", "REPLIKY", "POZNÁMKY"],
        )
        self.assertEqual(table_column_ids(komplet_table_root), list(range(1, 12)))

    def test_exported_workbook_does_not_include_prehled_sheet(self) -> None:
        project = build_project(
            {
                "title": "No Prehled",
                "episodes": [
                    {"label": "Pilot", "content": "POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tPrvní replika\n"},
                    {"label": "Finále", "content": "POSTAVA\tTC\tTEXT\nBETA\t00:00:02\tDruhá replika\n"},
                ],
                "assignments": {},
            }
        )

        workbook = export_project_workbook(project)
        with ZipFile(io.BytesIO(workbook)) as archive:
            names = set(archive.namelist())

            self.assertNotIn("PŘEHLED", workbook_sheet_names(archive))
            self.assertNotIn("xl/worksheets/sheet3.xml", names)
            self.assertNotIn("xl/worksheets/_rels/sheet3.xml.rels", names)
            self.assertNotIn("worksheets/sheet3.xml", workbook_sheet_targets(archive))
            self.assertNotIn("PŘEHLED", app_titles_of_parts(archive))

    def test_exported_workbook_keeps_simple_herci_layout_by_default(self) -> None:
        project = build_project(
            {
                "title": "Simple Herci",
                "episodes": [
                    {"label": "Empire", "content": "POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tJedna dvě tři čtyři pět šest sedm osm devět\n"},
                    {"label": "Hero Road", "content": "POSTAVA\tTC\tTEXT\nBETA\t00:00:02\tJedna dvě tři čtyři\n"},
                ],
                "assignments": {
                    "ALFA": {"actor": "Herec A", "note": ""},
                    "BETA": {"actor": "Herec B", "note": ""},
                },
            }
        )

        workbook = export_project_workbook(project)
        self.assert_workbook_structurally_consistent(workbook)
        with ZipFile(io.BytesIO(workbook)) as archive:
            herci_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            herci_table_root = ET.fromstring(archive.read("xl/tables/table1.xml"))

        self.assertEqual(
            table_column_names(herci_table_root),
            ["DABÉR", "VSTUPY CELKEM", "REPLIKY CELKEM", "POZNÁMKY"],
        )
        self.assertEqual(find_cell_text(herci_root, "A1"), "DABÉR")
        self.assertEqual(find_cell_text(herci_root, "B1"), "VSTUPY CELKEM")
        self.assertEqual(find_cell_text(herci_root, "C1"), "REPLIKY CELKEM")
        self.assertEqual(find_cell_text(herci_root, "D1"), "POZNÁMKY")

    def test_exported_workbook_can_split_herci_by_works_for_two_works(self) -> None:
        project = build_project(
            {
                "title": "Detailed Herci",
                "episodes": [
                    {"label": "Empire", "content": "POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tJedna dvě tři čtyři pět šest sedm osm devět\n"},
                    {"label": "Hero Road", "content": "POSTAVA\tTC\tTEXT\nALFA\t00:00:02\tAhoj světe\nBETA\t00:00:03\tJedna dvě tři čtyři\n"},
                ],
                "assignments": {
                    "ALFA": {"actor": "Herec A", "note": ""},
                    "BETA": {"actor": "Herec B", "note": ""},
                },
            }
        )

        workbook = export_project_workbook(project, herci_by_episode=True)
        self.assert_workbook_structurally_consistent(workbook)
        with ZipFile(io.BytesIO(workbook)) as archive:
            herci_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            herci_table_root = ET.fromstring(archive.read("xl/tables/table1.xml"))

        self.assertEqual(
            table_column_names(herci_table_root),
            [
                "DABÉR",
                "VSTUPY Empire",
                "REPLIKY Empire",
                "VSTUPY Hero Road",
                "REPLIKY Hero Road",
                "VSTUPY CELKEM",
                "REPLIKY CELKEM",
                "POZNÁMKY",
            ],
        )
        self.assertEqual(find_cell_text(herci_root, "A1"), "DABÉR")
        self.assertEqual(find_cell_text(herci_root, "B1"), "VSTUPY Empire")
        self.assertEqual(find_cell_text(herci_root, "C1"), "REPLIKY Empire")
        self.assertEqual(find_cell_text(herci_root, "D1"), "VSTUPY Hero Road")
        self.assertEqual(find_cell_text(herci_root, "E1"), "REPLIKY Hero Road")
        self.assertEqual(find_cell_text(herci_root, "F1"), "VSTUPY CELKEM")
        self.assertEqual(find_cell_text(herci_root, "G1"), "REPLIKY CELKEM")
        self.assertEqual(find_cell_text(herci_root, "H1"), "POZNÁMKY")

        self.assertEqual(find_cell_text(herci_root, "A2"), "Herec A")
        self.assertEqual(find_cell_text(herci_root, "B2"), "1")
        self.assertEqual(find_cell_text(herci_root, "C2"), "2")
        self.assertEqual(find_cell_text(herci_root, "D2"), "1")
        self.assertEqual(find_cell_text(herci_root, "E2"), "1")
        self.assertEqual(find_cell_text(herci_root, "F2"), "2")
        self.assertEqual(find_cell_text(herci_root, "G2"), "3")
        self.assertEqual(herci_table_root.attrib.get("ref"), "A1:H4")

    def test_exported_workbook_can_split_herci_by_works_for_six_works(self) -> None:
        project = build_project(
            {
                "title": "Detailed Herci 6",
                "episodes": [
                    {"label": f"Dílo {index + 1}", "content": f"POSTAVA\tTC\tTEXT\nPOSTAVA{index + 1}\t00:00:01\tText\n"}
                    for index in range(6)
                ],
                "assignments": {
                    f"POSTAVA{index + 1}": {"actor": f"Herec {index + 1}", "note": ""}
                    for index in range(6)
                },
            }
        )

        workbook = export_project_workbook(project, herci_by_episode=True)
        self.assert_workbook_structurally_consistent(workbook)
        with ZipFile(io.BytesIO(workbook)) as archive:
            herci_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            herci_table_root = ET.fromstring(archive.read("xl/tables/table1.xml"))

        self.assertEqual(len(table_column_names(herci_table_root)), 16)
        self.assertEqual(find_cell_text(herci_root, "A1"), "DABÉR")
        self.assertEqual(find_cell_text(herci_root, "B1"), "VSTUPY Dílo 1")
        self.assertEqual(find_cell_text(herci_root, "C1"), "REPLIKY Dílo 1")
        self.assertEqual(find_cell_text(herci_root, "L1"), "VSTUPY Dílo 6")
        self.assertEqual(find_cell_text(herci_root, "M1"), "REPLIKY Dílo 6")
        self.assertEqual(find_cell_text(herci_root, "N1"), "VSTUPY CELKEM")
        self.assertEqual(find_cell_text(herci_root, "O1"), "REPLIKY CELKEM")
        self.assertEqual(find_cell_text(herci_root, "P1"), "POZNÁMKY")

    def test_exported_workbook_applies_visual_polish_metadata(self) -> None:
        project = build_project(
            {
                "title": "Visual Polish",
                "episodes": [
                    {
                        "label": "Pilot",
                        "content": "POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tPrvní replika s více slovy\n",
                    },
                    {
                        "label": "Finále",
                        "content": "POSTAVA\tTC\tTEXT\nBETA\t00:00:02\tKrátká replika\n",
                    },
                ],
                "assignments": {
                    "BETA": {"actor": "Herec B", "note": "Hotovo"},
                },
            }
        )

        workbook = export_project_workbook(project)
        self.assert_workbook_structurally_consistent(workbook)

        with ZipFile(io.BytesIO(workbook)) as archive:
            styles_root = ET.fromstring(archive.read("xl/styles.xml"))
            herci_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            komplet_root = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
            pilot_root = ET.fromstring(archive.read("xl/worksheets/sheet4.xml"))

        self.assertEqual(pane_attributes(herci_root).get("ySplit"), "1")
        self.assertEqual(pane_attributes(herci_root).get("topLeftCell"), "A2")
        self.assertNotIn("topLeftCell", sheet_view_attributes(herci_root))
        self.assertEqual(row_attributes(herci_root, 1).get("ht"), "34.0")
        self.assertEqual(row_attributes(herci_root, 1).get("customHeight"), "1")
        self.assertEqual(row_attributes(herci_root, 2).get("ht"), "20.0")
        self.assertEqual(row_attributes(herci_root, 2).get("customHeight"), "1")
        self.assertEqual(pane_attributes(komplet_root).get("xSplit"), "1")
        self.assertEqual(pane_attributes(komplet_root).get("ySplit"), "1")
        self.assertEqual(pane_attributes(komplet_root).get("topLeftCell"), "B2")
        self.assertNotIn("topLeftCell", sheet_view_attributes(komplet_root))
        self.assertEqual(pane_attributes(pilot_root).get("ySplit"), "1")
        self.assertEqual(pane_attributes(pilot_root).get("topLeftCell"), "A2")
        self.assertNotIn("topLeftCell", sheet_view_attributes(pilot_root))

        self.assertGreater(column_width(herci_root, 1) or 0, 17.0)
        self.assertGreater(column_width(komplet_root, 1) or 0, 17.0)
        self.assertGreater(column_width(komplet_root, 4) or 0, 15.0)
        self.assertGreater(column_width(pilot_root, 4) or 0, 15.0)

        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(herci_root, "A1")), HEADER_FILL_RGB)
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(pilot_root, "A1")), HEADER_FILL_RGB)
        self.assertEqual(find_font_rgb(styles_root, find_cell_style(herci_root, "A1")), HEADER_FONT_RGB)
        self.assertEqual(find_font_rgb(styles_root, find_cell_style(pilot_root, "A1")), HEADER_FONT_RGB)

        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "A1")), KOMPLET_HEADER_PRIMARY_FILL_RGB)
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "B1")), KOMPLET_HEADER_EPISODE_FILL_RGB)
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "C1")), KOMPLET_HEADER_EPISODE_FILL_RGB)
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "D1")), KOMPLET_HEADER_PRIMARY_FILL_RGB)
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "E1")), KOMPLET_HEADER_SUMMARY_FILL_RGB)
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "F1")), KOMPLET_HEADER_SUMMARY_FILL_RGB)
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "G1")), KOMPLET_HEADER_NOTE_FILL_RGB)
        for ref in ("A1", "B1", "C1", "D1", "E1", "F1", "G1"):
            self.assertEqual(find_font_rgb(styles_root, find_cell_style(komplet_root, ref)), HEADER_FONT_RGB)

        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "A2")), "FFFFF2E0")
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(komplet_root, "D2")), "FFFFF2E0")
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(pilot_root, "A2")), "FFFFF2E0")
        self.assertEqual(find_fill_rgb(styles_root, find_cell_style(pilot_root, "D2")), "FFFFF2E0")


if __name__ == "__main__":
    unittest.main()
