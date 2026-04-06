from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from obsazovani.app_state import AppState
from obsazovani.core import build_project
from obsazovani.project_store import BulkImportSource


def make_dialogue(character: str, text: str) -> str:
    return f"POSTAVA\tTC\tTEXT\n{character}\t00:00:01\t{text}\n"


def make_summary(*rows: tuple[str, int, int]) -> str:
    header = "POSTAVA\tVSTUPY\tREPLIKY\n"
    body = "\n".join(f"{character}\t{inputs}\t{replicas}" for character, inputs, replicas in rows)
    return header + body + ("\n" if body else "")


class DynamicEpisodeModelTests(unittest.TestCase):
    def test_build_project_preserves_active_episode_count(self) -> None:
        project = build_project(
            {
                "title": "Dynamic Works",
                "episodes": [
                    {"label": "Pilot", "content": make_dialogue("ALFA", "Ahoj světe")},
                    {"label": "Finále", "content": ""},
                    {"label": "Trailer", "content": make_dialogue("BETA", "Jedna dvě tři čtyři")},
                ],
                "assignments": {},
            }
        )

        self.assertEqual(project["stats"]["episodeCount"], 3)
        self.assertEqual([episode["label"] for episode in project["episodes"]], ["Pilot", "Finále", "Trailer"])

    def test_build_project_keeps_duplicate_labels_separate(self) -> None:
        project = build_project(
            {
                "title": "Duplicate Labels",
                "episodes": [
                    {"label": "Pilot", "content": make_dialogue("ALFA", "První replika")},
                    {"label": "Pilot", "content": make_dialogue("ALFA", "Druhá replika")},
                ],
                "assignments": {},
            }
        )

        row = project["complete"][0]
        self.assertEqual(row["character"], "ALFA")
        self.assertEqual(len(row["episodes"]), 2)
        self.assertEqual([cell["inputs"] for cell in row["episodes"]], [1, 1])
        self.assertEqual([cell["replicas"] for cell in row["episodes"]], [1, 1])

    def test_app_state_save_load_supports_add_rename_remove(self) -> None:
        state = AppState()
        self.assertEqual(state.episode_count, 1)

        state.rename_episode(0, "Pilot")
        state.set_episode_content(0, make_dialogue("ALFA", "Ahoj světe"))

        second_index = state.add_episode()
        state.rename_episode(second_index, "Finále")
        state.set_episode_content(second_index, make_dialogue("BETA", "Druhá replika"))

        third_index = state.add_episode()
        state.rename_episode(third_index, "Trailer")
        state.set_episode_content(third_index, make_dialogue("GAMA", "Třetí replika"))

        state.remove_episode(1)

        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "dynamic-project.json"
            state.save_project(target)

            loaded = AppState()
            loaded.load_project(target)

        self.assertEqual(loaded.episode_count, 2)
        self.assertEqual(
            [episode["label"] for episode in loaded.payload["episodes"]],
            ["Pilot", "Trailer"],
        )
        self.assertIn("ALFA", loaded.payload["episodes"][0]["content"])
        self.assertIn("GAMA", loaded.payload["episodes"][1]["content"])

    def test_app_state_loads_legacy_six_episode_project(self) -> None:
        legacy_payload = {
            "version": 1,
            "title": "Legacy",
            "episodes": [
                {"label": f"{index + 1:02d}", "content": make_dialogue(f"POSTAVA{index + 1}", "Text")}
                for index in range(6)
            ],
            "assignments": {},
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "legacy-project.json"
            target.write_text(json.dumps(legacy_payload, ensure_ascii=False), encoding="utf-8")

            state = AppState()
            state.load_project(target)

        self.assertEqual(state.episode_count, 6)
        self.assertEqual([episode["label"] for episode in state.payload["episodes"]], ["01", "02", "03", "04", "05", "06"])

    def test_app_state_persists_herci_export_mode(self) -> None:
        state = AppState()
        state.set_herci_by_episode_export(True)

        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "export-mode.json"
            state.save_project(target)

            loaded = AppState()
            loaded.load_project(target)

        self.assertTrue(loaded.herci_by_episode_export)

    def test_app_state_bulk_import_adds_episodes_and_deduplicates_labels(self) -> None:
        state = AppState()
        state.rename_episode(0, "Pilot")
        state.set_episode_content(0, make_dialogue("ALFA", "Původní replika"))

        sources = [
            BulkImportSource("Empire.tsv", "Pilot", make_dialogue("BETA", "Nová replika")),
            BulkImportSource("Hero.tsv", "Pilot", make_dialogue("GAMA", "Druhá replika")),
            BulkImportSource("Bonus.tsv", "Bonus", make_dialogue("DELTA", "Třetí replika")),
        ]

        plan = state.preview_bulk_import(0, sources)
        imported_indexes = state.apply_bulk_import(plan)
        analysis = state.recompute()

        self.assertEqual(imported_indexes, [0, 1, 2])
        self.assertEqual(
            [episode["label"] for episode in state.payload["episodes"]],
            ["Pilot", "Pilot (2)", "Bonus"],
        )
        self.assertEqual(state.episode_count, 3)
        self.assertEqual(plan[0].overwrites_content, True)
        self.assertEqual(plan[1].creates_episode, True)
        self.assertEqual(plan[2].creates_episode, True)
        self.assertEqual(analysis["stats"]["episodeCount"], 3)

    def test_app_state_bulk_import_respects_maximum_episode_count(self) -> None:
        state = AppState()
        state.rename_episode(0, "01")
        for _ in range(5):
            state.add_episode()

        sources = [
            BulkImportSource(f"Source {index}", f"Label {index}", make_dialogue("ALFA", "Text"))
            for index in range(2)
        ]

        with self.assertRaisesRegex(ValueError, "k dispozici jen 1 slot"):
            state.preview_bulk_import(5, sources)

    def test_build_project_warns_only_for_unassigned_roles_from_fifty_replicas(self) -> None:
        project = build_project(
            {
                "title": "Unassigned Threshold",
                "episodes": [
                    {
                        "label": "Pilot",
                        "content": make_summary(
                            ("LOW", 10, 49),
                            ("HIGH", 10, 50),
                        ),
                    },
                ],
                "assignments": {},
            }
        )

        warning_messages = [
            str(item["message"]) for item in project["validations"] if item.get("severity") == "warning"
        ]

        self.assertFalse(any("LOW" in message for message in warning_messages))
        self.assertTrue(
            any(
                message == "Postava HIGH zatím nemá dabéra (10 vstupů, 50 replik; díla: Pilot)."
                for message in warning_messages
            )
        )
        self.assertEqual(project["validationSummary"]["warningCount"], 1)

    def test_build_project_detects_character_and_actor_name_variants(self) -> None:
        project = build_project(
            {
                "title": "Variant Names",
                "episodes": [
                    {
                        "label": "Pilot",
                        "content": make_summary(
                            ("SBOR", 2, 14),
                            ("Sbor", 1, 8),
                            ("ALFA", 4, 6),
                            ("BETA", 3, 5),
                        ),
                    },
                ],
                "assignments": {
                    "ALFA": {"actor": "Jan Battěk", "note": ""},
                    "BETA": {"actor": "JAN BATTEK:", "note": ""},
                    "SBOR": {"actor": "Kompars", "note": ""},
                    "Sbor": {"actor": "Kompars", "note": ""},
                },
            }
        )

        messages = [str(item["message"]) for item in project["validations"]]

        self.assertTrue(
            any(
                message.startswith("Možná jde o stejnou postavu:")
                and "SBOR" in message
                and "Sbor" in message
                for message in messages
            )
        )
        self.assertTrue(
            any(
                message.startswith("Možná jde o stejného dabéra:")
                and "Jan Battěk" in message
                and "JAN BATTEK:" in message
                for message in messages
            )
        )
        actor_variant_validation = next(
            item for item in project["validations"] if item.get("kind") == "actor_variants"
        )
        self.assertEqual(actor_variant_validation["variants"], ["JAN BATTEK:", "Jan Battěk"])
        self.assertTrue(actor_variant_validation["actionable"])
        self.assertEqual(project["validationSummary"]["warningCount"], 1)
        self.assertEqual(project["validationSummary"]["infoCount"], 1)

    def test_build_project_keeps_bucket_and_many_roles_as_info(self) -> None:
        project = build_project(
            {
                "title": "Info Validations",
                "episodes": [
                    {
                        "label": "Pilot",
                        "content": make_summary(
                            ("SBOR", 3, 22),
                            ("R1", 5, 10),
                            ("R2", 4, 12),
                            ("R3", 3, 8),
                        ),
                    },
                ],
                "assignments": {
                    "SBOR": {"actor": "Kompars", "note": ""},
                    "R1": {"actor": "Eva", "note": ""},
                    "R2": {"actor": "Eva", "note": ""},
                    "R3": {"actor": "Eva", "note": ""},
                },
            }
        )

        info_messages = [str(item["message"]) for item in project["validations"] if item.get("severity") == "info"]

        self.assertIn(
            "Postava SBOR má 22 replik. Zkontroluj, jestli pod ní není víc konkrétních rolí.",
            info_messages,
        )
        self.assertIn("Dabér Eva má v díle Pilot 3 postavy (12 vstupů, 30 replik).", info_messages)
        self.assertEqual(project["validationSummary"]["warningCount"], 0)
        self.assertEqual(project["validationSummary"]["infoCount"], 2)

    def test_build_project_skips_high_actor_load_for_single_character(self) -> None:
        project = build_project(
            {
                "title": "Single Character Load",
                "episodes": [
                    {
                        "label": "Pilot",
                        "content": make_summary(("LEAD", 45, 90)),
                    },
                    {
                        "label": "Finále",
                        "content": make_summary(("LEAD", 20, 35)),
                    },
                ],
                "assignments": {
                    "LEAD": {"actor": "Hlavní Dabér", "note": ""},
                },
            }
        )

        messages = [str(item["message"]) for item in project["validations"]]
        self.assertNotIn("Dabér Hlavní Dabér má vysokou celkovou zátěž (65 vstupů, 125 replik).", messages)
        self.assertNotIn("Dabér Hlavní Dabér má v díle Pilot vysokou zátěž (45 vstupů, 90 replik).", messages)

    def test_build_project_reports_high_actor_load_for_multiple_characters(self) -> None:
        project = build_project(
            {
                "title": "Multi Character Load",
                "episodes": [
                    {
                        "label": "Pilot",
                        "content": make_summary(
                            ("LEAD_A", 45, 90),
                            ("LEAD_B", 20, 35),
                        ),
                    },
                ],
                "assignments": {
                    "LEAD_A": {"actor": "Hlavní Dabér", "note": ""},
                    "LEAD_B": {"actor": "Hlavní Dabér", "note": ""},
                },
            }
        )

        messages = [str(item["message"]) for item in project["validations"]]
        self.assertIn(
            "Dabér Hlavní Dabér má vysokou celkovou zátěž (65 vstupů, 125 replik).",
            messages,
        )
        self.assertIn(
            "Dabér Hlavní Dabér má v díle Pilot vysokou zátěž (65 vstupů, 125 replik).",
            messages,
        )

    def test_build_project_applies_high_actor_load_per_scope(self) -> None:
        project = build_project(
            {
                "title": "Scoped Load",
                "episodes": [
                    {
                        "label": "Pilot",
                        "content": make_summary(("LEAD_A", 45, 90)),
                    },
                    {
                        "label": "Finále",
                        "content": make_summary(("LEAD_B", 20, 35)),
                    },
                ],
                "assignments": {
                    "LEAD_A": {"actor": "Hlavní Dabér", "note": ""},
                    "LEAD_B": {"actor": "Hlavní Dabér", "note": ""},
                },
            }
        )

        messages = [str(item["message"]) for item in project["validations"]]
        self.assertIn(
            "Dabér Hlavní Dabér má vysokou celkovou zátěž (65 vstupů, 125 replik).",
            messages,
        )
        self.assertNotIn(
            "Dabér Hlavní Dabér má v díle Pilot vysokou zátěž (45 vstupů, 90 replik).",
            messages,
        )
        self.assertNotIn(
            "Dabér Hlavní Dabér má v díle Finále vysokou zátěž (20 vstupů, 35 replik).",
            messages,
        )

    def test_app_state_can_unify_actor_variants_in_assignments(self) -> None:
        state = AppState()
        state.set_episode_content(
            0,
            make_summary(
                ("ALFA", 5, 10),
                ("BETA", 6, 12),
                ("GAMA", 4, 9),
            ),
        )
        state.set_assignment("ALFA", "actor", "Jan Battěk")
        state.set_assignment("BETA", "actor", "JAN BATTĚK")
        state.set_assignment("GAMA", "actor", "Jan  Battek")
        state.dirty = False

        changed = state.unify_actor_variants(["Jan Battěk", "JAN BATTĚK", "Jan  Battek"], "Jan Battěk")
        analysis = state.recompute()

        self.assertEqual(changed, 2)
        self.assertTrue(state.dirty)
        self.assertEqual(state.payload["assignments"]["ALFA"]["actor"], "Jan Battěk")
        self.assertEqual(state.payload["assignments"]["BETA"]["actor"], "Jan Battěk")
        self.assertEqual(state.payload["assignments"]["GAMA"]["actor"], "Jan Battěk")
        self.assertFalse(any(item.get("kind") == "actor_variants" for item in analysis["validations"]))


if __name__ == "__main__":
    unittest.main()
