from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

from obsazovani.app_state import AppState
from obsazovani.importers import import_episode_source, list_importable_xlsx_sheets
from obsazovani.project_store import (
    read_bulk_import_sources_from_directory,
    read_bulk_import_sources_from_files,
    read_bulk_import_sources_from_workbook,
)


def column_name(index: int) -> str:
    name = ""
    current = index + 1
    while current:
        current, remainder = divmod(current - 1, 26)
        name = chr(65 + remainder) + name
    return name


def inline_cell(reference: str, value: str) -> str:
    return (
        f'<c r="{reference}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    )


def make_sheet_xml(rows: list[list[str]]) -> str:
    row_xml: list[str] = []
    for row_index, values in enumerate(rows, start=1):
        cells = [
            inline_cell(f"{column_name(column_index)}{row_index}", value)
            for column_index, value in enumerate(values)
            if value != ""
        ]
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        "</worksheet>"
    )


def make_workbook_xml(sheet_names: list[str]) -> str:
    sheets_xml = "".join(
        f'<sheet name="{escape(sheet_name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, sheet_name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets_xml}</sheets>"
        "</workbook>"
    )


def make_workbook_rels(sheet_names: list[str]) -> str:
    relationships_xml = "".join(
        (
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
        for index, _ in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{relationships_xml}"
        "</Relationships>"
    )


def write_test_xlsx(path: Path, sheets: list[tuple[str, list[list[str]]]]) -> None:
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("xl/workbook.xml", make_workbook_xml([sheet_name for sheet_name, _ in sheets]))
        archive.writestr("xl/_rels/workbook.xml.rels", make_workbook_rels([sheet_name for sheet_name, _ in sheets]))
        for index, (_, rows) in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", make_sheet_xml(rows))


class ImportersXlsxTests(unittest.TestCase):
    def test_dialogue_sheet_is_imported_from_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "dialogue.xlsx"
            write_test_xlsx(
                path,
                [
                    (
                        "Dialog",
                        [
                            ["POSTAVA", "TC", "TEXT", "POČET SLOV"],
                            ["ALFA", "00:00:01", "Ahoj světe", "2"],
                        ],
                    )
                ],
            )

            options = list_importable_xlsx_sheets(path)
            imported = import_episode_source(path, sheet_name="Dialog")

        self.assertEqual([option.sheet_name for option in options], ["Dialog"])
        self.assertEqual(imported.source_format, "xlsx-dialogue")
        self.assertIn("POSTAVA\tTC\tTEXT", imported.content)
        self.assertIn("ALFA\t00:00:01\tAhoj světe", imported.content)

    def test_summary_sheets_are_listed_and_aggregate_sheet_is_skipped(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "summary.xlsx"
            write_test_xlsx(
                path,
                [
                    (
                        "KOMPLET",
                        [
                            ["POSTAVA", "1", "2", "DABÉR", "VSTUPY", "REPLIKY"],
                            ["ABBY", "12 / 15", "10 / 11", "Herec", "22", "26"],
                        ],
                    ),
                    (
                        "01",
                        [
                            ["POSTAVA", "VSTUPY", "REPLIKY", "DABÉR"],
                            ["ABBY", "12", "15", "Herec"],
                        ],
                    ),
                    (
                        "02",
                        [
                            ["POSTAVA", "VSTUPY", "REPLIKY", "DABÉR"],
                            ["BETA", "4", "6", "Jiný herec"],
                        ],
                    ),
                ],
            )

            options = list_importable_xlsx_sheets(path)
            imported = import_episode_source(path, sheet_name="02")

        self.assertEqual([option.sheet_name for option in options], ["01", "02"])
        self.assertEqual(imported.source_format, "xlsx-summary")
        self.assertIn("POSTAVA\tVSTUPY\tREPLIKY\tDABÉR", imported.content)
        self.assertIn("BETA\t4\t6\tJiný herec", imported.content)
        self.assertEqual(imported.assignments, {"BETA": {"actor": "Jiný herec", "note": ""}})

    def test_dialogue_sheet_imports_actor_and_note_assignments(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "dialogue-with-cast.xlsx"
            write_test_xlsx(
                path,
                [
                    (
                        "Dialog",
                        [
                            ["POSTAVA", "TC", "TEXT", "DABÉR", "POZNÁMKA"],
                            ["ALFA", "00:00:01", "Ahoj světe", "Herec A", "Lead"],
                            ["ALFA", "00:00:02", "Další replika", "", ""],
                            ["BETA", "00:00:03", "Třetí replika", "Herec B", ""],
                        ],
                    )
                ],
            )

            imported = import_episode_source(path, sheet_name="Dialog")

        self.assertIn("POSTAVA\tTC\tTEXT\tDABÉR\tPOZNÁMKA", imported.content)
        self.assertEqual(
            imported.assignments,
            {
                "ALFA": {"actor": "Herec A", "note": "Lead"},
                "BETA": {"actor": "Herec B", "note": ""},
            },
        )

    def test_app_state_import_merges_assignments_without_overwriting_manual_actor(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "summary-with-cast.xlsx"
            write_test_xlsx(
                path,
                [
                    (
                        "01",
                        [
                            ["POSTAVA", "VSTUPY", "REPLIKY", "DABÉR", "POZNÁMKA"],
                            ["ALFA", "4", "5", "Import Herec", "Import Poznámka"],
                            ["BETA", "2", "3", "Import Beta", ""],
                        ],
                    )
                ],
            )

            state = AppState()
            state.set_assignment("ALFA", "actor", "Ruční herec")
            state.import_episode_file(0, path, sheet_name="01")
            analysis = state.recompute()

        rows = {row["character"]: row for row in analysis["complete"]}
        self.assertEqual(rows["ALFA"]["actor"], "Ruční herec")
        self.assertEqual(rows["ALFA"]["note"], "Import Poznámka")
        self.assertEqual(rows["BETA"]["actor"], "Import Beta")

        state.set_assignment("BETA", "actor", "")
        analysis_after_clear = state.recompute()
        rows_after_clear = {row["character"]: row for row in analysis_after_clear["complete"]}
        self.assertEqual(rows_after_clear["BETA"]["actor"], "")

    def test_bulk_import_sources_from_workbook_use_sheet_names(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "project.xlsx"
            write_test_xlsx(
                path,
                [
                    ("Empire", [["POSTAVA", "VSTUPY", "REPLIKY"], ["ALFA", "4", "5"]]),
                    ("Hero Road", [["POSTAVA", "VSTUPY", "REPLIKY"], ["BETA", "2", "3"]]),
                ],
            )

            sources = read_bulk_import_sources_from_workbook(path)

        self.assertEqual([source.label for source in sources], ["Empire", "Hero Road"])
        self.assertEqual([source.source_name for source in sources], ["project.xlsx · Empire", "project.xlsx · Hero Road"])

    def test_bulk_import_sources_from_files_use_file_names(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            first_path = Path(temp_dir) / "Empire_script.tsv"
            second_path = Path(temp_dir) / "Hero Road.csv"
            first_path.write_text("POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tAhoj světe\n", encoding="utf-8")
            second_path.write_text("POSTAVA;VSTUPY;REPLIKY\nBETA;2;3\n", encoding="utf-8")

            sources = read_bulk_import_sources_from_files([second_path, first_path])

        self.assertEqual([source.label for source in sources], ["Empire script", "Hero Road"])
        self.assertEqual([source.source_name for source in sources], ["Empire_script.tsv", "Hero Road.csv"])

    def test_bulk_import_sources_from_files_reject_ambiguous_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "ambiguous.xlsx"
            write_test_xlsx(
                path,
                [
                    ("01", [["POSTAVA", "VSTUPY", "REPLIKY"], ["ALFA", "4", "5"]]),
                    ("02", [["POSTAVA", "VSTUPY", "REPLIKY"], ["BETA", "2", "3"]]),
                ],
            )

            with self.assertRaisesRegex(ValueError, "více použitelných listů"):
                read_bulk_import_sources_from_files([path])

    def test_bulk_import_sources_from_directory_collect_supported_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            (directory / "Empire.tsv").write_text("POSTAVA\tTC\tTEXT\nALFA\t00:00:01\tText\n", encoding="utf-8")
            (directory / "Hero.csv").write_text("POSTAVA;VSTUPY;REPLIKY\nBETA;2;3\n", encoding="utf-8")
            (directory / ".DS_Store").write_text("ignore", encoding="utf-8")

            sources = read_bulk_import_sources_from_directory(directory)

        self.assertEqual([source.label for source in sources], ["Empire", "Hero"])

    def test_netflix_sheet_remains_supported(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "netflix.xlsx"
            write_test_xlsx(
                path,
                [
                    (
                        "Netflix",
                        [
                            ["SOURCE", "DIALOGUE", "IN-TIMECODE"],
                            ["ALFA", "Hello there", "00:00:01"],
                        ],
                    )
                ],
            )

            options = list_importable_xlsx_sheets(path)
            imported = import_episode_source(path)

        self.assertEqual([option.source_format for option in options], ["netflix-xlsx"])
        self.assertEqual(imported.source_format, "netflix-xlsx")
        self.assertIn("ALFA\t00:00:01\tHello there", imported.content)


if __name__ == "__main__":
    unittest.main()
