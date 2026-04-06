from __future__ import annotations

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from obsazovani.desktop.main_window import MainWindow
from obsazovani.desktop.models import CastingFilterProxyModel, CastingTableModel


def get_app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def make_analysis() -> dict:
    return {
        "episodes": [
            {"label": "Pilot"},
            {"label": "Finále"},
        ],
        "complete": [
            {
                "character": "ALFA",
                "episodes": [
                    {"inputs": 1, "replicas": 8, "display": "1 / 8"},
                    {"inputs": 0, "replicas": 0, "display": "0 / 0"},
                ],
                "totalInputs": 1,
                "totalReplicas": 8,
                "actor": "",
                "note": "hlavní role",
            },
            {
                "character": "BETA",
                "episodes": [
                    {"inputs": 3, "replicas": 1, "display": "3 / 1"},
                    {"inputs": 2, "replicas": 1, "display": "2 / 1"},
                ],
                "totalInputs": 5,
                "totalReplicas": 2,
                "actor": "Boris",
                "note": "vedlejší role",
            },
            {
                "character": "GAMA",
                "episodes": [
                    {"inputs": 3, "replicas": 4, "display": "3 / 4"},
                    {"inputs": 1, "replicas": 2, "display": "1 / 2"},
                ],
                "totalInputs": 4,
                "totalReplicas": 6,
                "actor": "Adam",
                "note": "pilotní scéna",
            },
        ],
    }


class CastingProxyModelTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = get_app()

    def setUp(self) -> None:
        self.model = CastingTableModel()
        self.model.set_analysis(make_analysis())
        self.proxy = CastingFilterProxyModel()
        self.proxy.setSourceModel(self.model)

    def test_search_matches_character_actor_and_note(self) -> None:
        self.proxy.set_search_text("alfa")
        self.assertEqual(self.proxy.rowCount(), 1)
        self.assertEqual(self.proxy.index(0, 0).data(), "ALFA")

        self.proxy.set_search_text("boris")
        self.assertEqual(self.proxy.rowCount(), 1)
        self.assertEqual(self.proxy.index(0, 0).data(), "BETA")

        self.proxy.set_search_text("pilotní")
        self.assertEqual(self.proxy.rowCount(), 1)
        self.assertEqual(self.proxy.index(0, 0).data(), "GAMA")

    def test_assignment_filter_supports_unassigned_and_assigned(self) -> None:
        self.proxy.set_assignment_filter(CastingFilterProxyModel.FILTER_UNASSIGNED)
        self.assertEqual(self.proxy.rowCount(), 1)
        self.assertEqual(self.proxy.index(0, 0).data(), "ALFA")

        self.proxy.set_assignment_filter(CastingFilterProxyModel.FILTER_ASSIGNED)
        self.assertEqual(self.proxy.rowCount(), 2)
        self.assertEqual({self.proxy.index(row, 0).data() for row in range(self.proxy.rowCount())}, {"BETA", "GAMA"})

    def test_sorting_is_numeric_for_replicas_and_episode_columns(self) -> None:
        self.proxy.sort(self.model.replicas_column, Qt.DescendingOrder)
        self.assertEqual(
            [self.proxy.index(row, 0).data() for row in range(self.proxy.rowCount())],
            ["ALFA", "GAMA", "BETA"],
        )

        self.proxy.sort(1, Qt.DescendingOrder)
        self.assertEqual(
            [self.proxy.index(row, 0).data() for row in range(self.proxy.rowCount())],
            ["GAMA", "BETA", "ALFA"],
        )

        self.model.setData(self.model.index(0, self.model.actor_column), "Cyril")
        self.proxy.sort(self.model.actor_column, Qt.AscendingOrder)
        self.assertEqual(
            [self.proxy.index(row, 0).data() for row in range(self.proxy.rowCount())],
            ["GAMA", "BETA", "ALFA"],
        )

    def test_editing_through_proxy_updates_source_model_under_filter(self) -> None:
        self.proxy.set_assignment_filter(CastingFilterProxyModel.FILTER_UNASSIGNED)
        self.proxy.set_search_text("alfa")
        actor_index = self.proxy.index(0, self.model.actor_column)

        self.assertTrue(self.proxy.setData(actor_index, "Eva"))
        self.assertEqual(self.model.index(0, self.model.actor_column).data(), "Eva")
        self.assertEqual(self.proxy.rowCount(), 0)

        self.proxy.set_assignment_filter(CastingFilterProxyModel.FILTER_ASSIGNED)
        self.proxy.set_search_text("adam")
        self.proxy.sort(self.model.actor_column, Qt.AscendingOrder)
        note_index = self.proxy.index(0, self.model.note_column)

        self.assertTrue(self.proxy.setData(note_index, "nová poznámka"))
        self.assertEqual(self.model.index(2, self.model.note_column).data(), "nová poznámka")


class MainWindowCastingToolsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = get_app()

    def test_main_window_exposes_search_and_filter_controls(self) -> None:
        window = MainWindow()
        try:
            self.assertEqual(
                window._casting_search_edit.placeholderText(),
                "Hledat v postavách, dabérech a poznámkách",
            )
            self.assertEqual(window._casting_filter_combo.count(), 3)
            self.assertTrue(window._casting_table.isSortingEnabled())
        finally:
            window.deleteLater()
            self.app.processEvents()

    def test_validation_action_button_is_enabled_only_for_actor_variants(self) -> None:
        window = MainWindow()
        try:
            window._state.set_episode_content(
                0,
                "POSTAVA\tVSTUPY\tREPLIKY\nALFA\t4\t6\nBETA\t3\t5\n",
            )
            window._state.set_assignment("ALFA", "actor", "Jan Battěk")
            window._state.set_assignment("BETA", "actor", "JAN BATTĚK")
            window.refresh_analysis()

            self.assertFalse(window._validation_action_button.isEnabled())
            window._validation_table.selectRow(0)
            self.app.processEvents()
            self.assertTrue(window._validation_action_button.isEnabled())
        finally:
            window.deleteLater()
            self.app.processEvents()

    def test_komplet_focus_mode_hides_side_panels_and_restores_layout(self) -> None:
        window = MainWindow()
        try:
            window._state.set_episode_content(
                0,
                "POSTAVA\tVSTUPY\tREPLIKY\nALFA\t4\t6\nBETA\t3\t5\n",
            )
            window._state.set_assignment("BETA", "actor", "Boris")
            window.refresh_analysis()
            window.show()
            self.app.processEvents()

            root_before = window._root_splitter.sizes()
            right_before = window._right_splitter.sizes()

            window._komplet_focus_action.trigger()
            self.app.processEvents()

            self.assertTrue(window._komplet_focus_action.isChecked())
            self.assertTrue(window._editor_group.isHidden())
            self.assertFalse(window._komplet_group.isHidden())
            self.assertTrue(window._summary_group.isHidden())
            self.assertTrue(window._validation_group.isHidden())
            self.assertTrue(window._casting_table.isSortingEnabled())

            window._casting_search_edit.setText("alfa")
            window._casting_filter_combo.setCurrentIndex(1)
            self.app.processEvents()
            self.assertEqual(window._casting_proxy_model.rowCount(), 1)

            actor_index = window._casting_proxy_model.index(0, window._casting_model.actor_column)
            self.assertTrue(window._casting_proxy_model.setData(actor_index, "Eva"))
            self.assertEqual(window._casting_model.index(0, window._casting_model.actor_column).data(), "Eva")
            self.assertEqual(window._casting_proxy_model.rowCount(), 0)

            window._komplet_focus_action.trigger()
            self.app.processEvents()

            self.assertFalse(window._komplet_focus_action.isChecked())
            self.assertFalse(window._editor_group.isHidden())
            self.assertFalse(window._summary_group.isHidden())
            self.assertFalse(window._validation_group.isHidden())
            self.assertEqual(window._root_splitter.sizes(), root_before)
            self.assertEqual(window._right_splitter.sizes(), right_before)
        finally:
            window.deleteLater()
            self.app.processEvents()


if __name__ == "__main__":
    unittest.main()
