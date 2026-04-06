from __future__ import annotations

import re
from pathlib import Path

from PySide6.QtCore import QSignalBlocker, QTimer, Qt
from PySide6.QtGui import QAction, QCloseEvent, QColor, QKeySequence, QPalette
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStatusBar,
    QStyledItemDelegate,
    QTabWidget,
    QTableView,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from obsazovani.app_state import AppState, BulkImportPlanEntry
from obsazovani.core import MAX_EPISODES, normalize_text
from obsazovani.i18n import get_language, set_language, t
from obsazovani.desktop.models import (
    ActorSummaryTableModel,
    CastingFilterProxyModel,
    CastingTableModel,
    ValidationTableModel,
)
from obsazovani.desktop.widgets import EpisodeEditorWidget
from obsazovani.project_store import (
    BulkImportSource,
    list_episode_source_options,
    read_bulk_import_sources_from_directory,
    read_bulk_import_sources_from_files,
    read_bulk_import_sources_from_workbook,
)


def slugify_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip())
    return cleaned.strip("-") or "casting"


class CastingEditorDelegate(QStyledItemDelegate):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._base_color = QColor("#fffaf4")
        self._text_color = QColor("#20170f")
        self._selection_color = QColor("#8f4b2b")
        self._selection_text_color = QColor("#fff9f4")

    def createEditor(self, parent: QWidget, option, index):  # type: ignore[override]
        editor = super().createEditor(parent, option, index)
        palette = editor.palette()
        palette.setColor(QPalette.Base, self._base_color)
        palette.setColor(QPalette.Text, self._text_color)
        palette.setColor(QPalette.Highlight, self._selection_color)
        palette.setColor(QPalette.HighlightedText, self._selection_text_color)
        editor.setPalette(palette)
        return editor


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self._state = AppState()
        self._analysis_timer = QTimer(self)
        self._analysis_timer.setSingleShot(True)
        self._analysis_timer.setInterval(300)
        self._analysis_timer.timeout.connect(self.refresh_analysis)

        self._casting_model = CastingTableModel(self)
        self._casting_proxy_model = CastingFilterProxyModel(self)
        self._casting_proxy_model.setSourceModel(self._casting_model)
        self._actor_model = ActorSummaryTableModel(self)
        self._validation_model = ValidationTableModel(self)
        self._episode_editors: list[EpisodeEditorWidget] = []
        self._metric_values: dict[str, QLabel] = {}
        self._editor_group: QGroupBox | None = None
        self._komplet_group: QGroupBox | None = None
        self._summary_group: QGroupBox | None = None
        self._validation_group: QGroupBox | None = None
        self._root_splitter: QSplitter | None = None
        self._right_splitter: QSplitter | None = None
        self._saved_root_splitter_sizes: list[int] | None = None
        self._saved_right_splitter_sizes: list[int] | None = None

        self._create_actions()
        self._build_ui()
        self._connect_signals()
        self._populate_from_state()
        self.refresh_analysis()

    def _create_actions(self) -> None:
        self._new_action = QAction(t("action.new_project"), self)
        self._new_action.setShortcut(QKeySequence.New)
        self._new_action.triggered.connect(self.new_project)

        self._open_action = QAction(t("action.open_project"), self)
        self._open_action.setShortcut(QKeySequence.Open)
        self._open_action.triggered.connect(self.open_project)

        self._bulk_import_action = QAction(t("action.bulk_import"), self)
        self._bulk_import_action.triggered.connect(self.bulk_import_episodes)

        self._save_action = QAction(t("action.save_project"), self)
        self._save_action.setShortcut(QKeySequence.Save)
        self._save_action.triggered.connect(self.save_project)

        self._save_as_action = QAction(t("action.save_as"), self)
        self._save_as_action.setShortcut(QKeySequence.SaveAs)
        self._save_as_action.triggered.connect(self.save_project_as)

        self._recalculate_action = QAction(t("action.recalculate"), self)
        self._recalculate_action.setShortcut("Ctrl+R")
        self._recalculate_action.triggered.connect(self.refresh_analysis)

        self._export_action = QAction(t("action.export_xlsx"), self)
        self._export_action.setShortcut("Ctrl+E")
        self._export_action.triggered.connect(self.export_workbook)

        self._komplet_focus_action = QAction(t("action.komplet_focus"), self)
        self._komplet_focus_action.setCheckable(True)
        self._komplet_focus_action.setShortcut("Ctrl+Shift+K")
        self._komplet_focus_action.setToolTip(t("action.komplet_focus.tip"))
        self._komplet_focus_action.toggled.connect(self._set_komplet_focus_mode)

        self._quit_action = QAction(t("action.quit"), self)
        self._quit_action.setShortcut(QKeySequence.Quit)
        self._quit_action.triggered.connect(self.close)

        self._lang_cs_action = QAction(t("action.lang_cs"), self)
        self._lang_cs_action.setCheckable(True)
        self._lang_cs_action.triggered.connect(lambda: self._switch_language("cs"))

        self._lang_en_action = QAction(t("action.lang_en"), self)
        self._lang_en_action.setCheckable(True)
        self._lang_en_action.triggered.connect(lambda: self._switch_language("en"))

        self._update_lang_action_checks()

    def _build_ui(self) -> None:
        self.setWindowTitle(t("app.name"))
        self.resize(1500, 920)
        self.setMinimumSize(1180, 760)
        self._apply_style()

        self._file_menu = self.menuBar().addMenu(t("menu.file"))
        self._file_menu.addAction(self._new_action)
        self._file_menu.addAction(self._open_action)
        self._file_menu.addAction(self._bulk_import_action)
        self._file_menu.addSeparator()
        self._file_menu.addAction(self._save_action)
        self._file_menu.addAction(self._save_as_action)
        self._file_menu.addSeparator()
        self._file_menu.addAction(self._export_action)
        self._file_menu.addSeparator()
        self._file_menu.addAction(self._quit_action)

        self._tools_menu = self.menuBar().addMenu(t("menu.tools"))
        self._tools_menu.addAction(self._recalculate_action)
        self._tools_menu.addAction(self._komplet_focus_action)

        self._lang_menu = self.menuBar().addMenu(t("menu.language"))
        self._lang_menu.addAction(self._lang_cs_action)
        self._lang_menu.addAction(self._lang_en_action)

        toolbar = QToolBar(t("toolbar.main"), self)
        toolbar.setMovable(False)
        toolbar.addAction(self._new_action)
        toolbar.addAction(self._open_action)
        toolbar.addAction(self._bulk_import_action)
        toolbar.addAction(self._save_action)
        toolbar.addAction(self._save_as_action)
        toolbar.addSeparator()
        toolbar.addAction(self._recalculate_action)
        toolbar.addAction(self._komplet_focus_action)
        toolbar.addAction(self._export_action)
        self.addToolBar(toolbar)

        status_bar = QStatusBar(self)
        self.setStatusBar(status_bar)

        self._title_edit = QLineEdit()
        self._title_edit.setPlaceholderText(t("title.placeholder"))
        self._herci_by_episode_checkbox = QCheckBox(t("checkbox.herci_by_episode"))
        self._herci_by_episode_checkbox.setToolTip(t("checkbox.herci_by_episode.tip"))

        title_row = QHBoxLayout()
        self._title_label = QLabel(t("label.project_name"))
        self._title_label.setMinimumWidth(110)
        title_row.addWidget(self._title_label)
        title_row.addWidget(self._title_edit, 1)
        self._title_row = title_row

        export_options_row = QHBoxLayout()
        export_options_row.addWidget(self._herci_by_episode_checkbox)
        export_options_row.addStretch(1)

        self._casting_search_edit = QLineEdit()
        self._casting_search_edit.setClearButtonEnabled(True)
        self._casting_search_edit.setPlaceholderText(t("search.placeholder"))
        self._casting_filter_combo = QComboBox()
        self._casting_filter_combo.addItem(t("filter.all"), CastingFilterProxyModel.FILTER_ALL)
        self._casting_filter_combo.addItem(t("filter.unassigned"), CastingFilterProxyModel.FILTER_UNASSIGNED)
        self._casting_filter_combo.addItem(t("filter.assigned"), CastingFilterProxyModel.FILTER_ASSIGNED)

        self._search_label = QLabel(t("label.search"))
        self._show_label = QLabel(t("label.show"))
        casting_tools_row = QHBoxLayout()
        casting_tools_row.addWidget(self._search_label)
        casting_tools_row.addWidget(self._casting_search_edit, 1)
        casting_tools_row.addWidget(self._show_label)
        casting_tools_row.addWidget(self._casting_filter_combo)

        metrics_row = QHBoxLayout()
        metrics_row.setSpacing(10)
        for key, label in (
            ("characterCount", t("metric.characters")),
            ("inputs", t("metric.inputs")),
            ("replicas", t("metric.replicas")),
            ("missing", t("metric.missing")),
        ):
            card, value_label = self._create_metric_card(label)
            metrics_row.addWidget(card)
            self._metric_values[key] = value_label
        metrics_row.addStretch(1)

        self._episode_tabs = QTabWidget()
        self._episode_tabs.setDocumentMode(True)

        self._add_episode_button = QPushButton(t("btn.add_episode"))
        self._rename_episode_button = QPushButton(t("btn.rename_episode"))
        self._remove_episode_button = QPushButton(t("btn.remove_episode"))

        episode_actions = QHBoxLayout()
        episode_actions.addWidget(self._add_episode_button)
        episode_actions.addWidget(self._rename_episode_button)
        episode_actions.addWidget(self._remove_episode_button)
        episode_actions.addStretch(1)

        editor_group = QGroupBox(t("group.episodes"))
        editor_layout = QVBoxLayout(editor_group)
        editor_layout.addLayout(episode_actions)
        editor_layout.addWidget(self._episode_tabs)
        self._editor_group = editor_group

        self._casting_table = QTableView()
        self._casting_table.setModel(self._casting_proxy_model)
        self._casting_table.setAlternatingRowColors(True)
        self._casting_table.setSelectionBehavior(QTableView.SelectRows)
        self._casting_table.setSelectionMode(QTableView.SingleSelection)
        self._casting_table.setItemDelegate(CastingEditorDelegate(self._casting_table))
        self._casting_table.setSortingEnabled(True)
        self._casting_table.verticalHeader().setVisible(False)
        self._casting_table.horizontalHeader().setStretchLastSection(False)
        self._configure_casting_table_columns()
        self._configure_casting_table_appearance()

        komplet_group = QGroupBox(t("group.komplet"))
        komplet_layout = QVBoxLayout(komplet_group)
        komplet_layout.addLayout(metrics_row)
        komplet_layout.addLayout(casting_tools_row)
        komplet_layout.addWidget(self._casting_table, 1)
        self._komplet_group = komplet_group

        self._actor_table = QTableView()
        self._actor_table.setModel(self._actor_model)
        self._actor_table.setAlternatingRowColors(True)
        self._actor_table.setSelectionBehavior(QTableView.SelectRows)
        self._actor_table.setSelectionMode(QTableView.SingleSelection)
        self._actor_table.setEditTriggers(QTableView.NoEditTriggers)
        self._actor_table.verticalHeader().setVisible(False)
        self._actor_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._actor_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._actor_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)

        self._status_label = QLabel()
        self._status_label.setWordWrap(True)
        self._status_label.setProperty("statusLevel", "info")

        summary_group = QGroupBox(t("group.actor_summary"))
        summary_layout = QVBoxLayout(summary_group)
        summary_layout.addWidget(self._actor_table, 1)
        summary_layout.addWidget(self._status_label)
        self._summary_group = summary_group

        self._validation_table = QTableView()
        self._validation_table.setModel(self._validation_model)
        self._validation_table.setAlternatingRowColors(False)
        self._validation_table.setSelectionBehavior(QTableView.SelectRows)
        self._validation_table.setSelectionMode(QTableView.SingleSelection)
        self._validation_table.setEditTriggers(QTableView.NoEditTriggers)
        self._validation_table.setWordWrap(True)
        self._validation_table.verticalHeader().setVisible(False)
        self._validation_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self._validation_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._validation_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)

        self._validation_action_button = QPushButton(t("btn.unify_name"))
        self._validation_action_button.setEnabled(False)
        self._validation_action_button.setToolTip(t("btn.unify_name.tip"))

        validation_group = QGroupBox(t("group.validation"))
        validation_layout = QVBoxLayout(validation_group)
        validation_layout.addWidget(self._validation_table, 1)
        validation_layout.addWidget(self._validation_action_button)
        self._validation_group = validation_group

        right_splitter = QSplitter(Qt.Vertical)
        right_splitter.addWidget(komplet_group)
        right_splitter.addWidget(summary_group)
        right_splitter.addWidget(validation_group)
        right_splitter.setStretchFactor(0, 5)
        right_splitter.setStretchFactor(1, 2)
        right_splitter.setStretchFactor(2, 2)
        self._right_splitter = right_splitter

        root_splitter = QSplitter(Qt.Horizontal)
        root_splitter.addWidget(editor_group)
        root_splitter.addWidget(right_splitter)
        root_splitter.setStretchFactor(0, 3)
        root_splitter.setStretchFactor(1, 6)
        self._root_splitter = root_splitter

        central = QWidget()
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(12, 12, 12, 12)
        central_layout.setSpacing(12)
        central_layout.addLayout(title_row)
        central_layout.addLayout(export_options_row)
        central_layout.addWidget(root_splitter, 1)
        self.setCentralWidget(central)

    def _connect_signals(self) -> None:
        self._title_edit.textChanged.connect(self._handle_title_changed)
        self._herci_by_episode_checkbox.toggled.connect(self._handle_herci_export_mode_changed)
        self._casting_search_edit.textChanged.connect(self._handle_casting_search_changed)
        self._casting_filter_combo.currentIndexChanged.connect(self._handle_casting_filter_changed)
        self._casting_model.assignmentEdited.connect(self._handle_assignment_edited)
        self._validation_action_button.clicked.connect(self.unify_selected_validation_actor_name)
        self._validation_table.selectionModel().selectionChanged.connect(self._update_validation_actions)
        self._add_episode_button.clicked.connect(self.add_episode)
        self._rename_episode_button.clicked.connect(self.rename_current_episode)
        self._remove_episode_button.clicked.connect(self.remove_current_episode)
        self._episode_tabs.currentChanged.connect(self._update_episode_controls)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QFrame[metricCard="true"] {
                border: 1px solid #3b3028;
                border-radius: 10px;
                background: #171311;
                padding: 6px 10px;
            }
            QLabel[metricLabel="true"] {
                color: #cdb9a8;
                font-size: 12px;
                font-weight: 600;
            }
            QLabel#metricValue {
                font-size: 22px;
                font-weight: 600;
                color: #fff9f4;
            }
            QLabel[statusLevel="ok"] {
                border-radius: 10px;
                padding: 10px;
                background: #e6f4ea;
                color: #1f5134;
            }
            QLabel[statusLevel="warning"] {
                border-radius: 10px;
                padding: 10px;
                background: #fff4dd;
                color: #7a4d00;
            }
            QLabel[statusLevel="info"] {
                border-radius: 10px;
                padding: 10px;
                background: #eef4ff;
                color: #1f3d68;
            }
            QLabel[statusLevel="error"] {
                border-radius: 10px;
                padding: 10px;
                background: #fde8e8;
                color: #8a1f1f;
            }
            """
        )

    def _configure_casting_table_appearance(self) -> None:
        palette = self._casting_table.palette()
        palette.setColor(QPalette.Base, QColor("#fffdf9"))
        palette.setColor(QPalette.AlternateBase, QColor("#f6efe6"))
        palette.setColor(QPalette.Text, QColor("#20170f"))
        palette.setColor(QPalette.Highlight, QColor("#8f4b2b"))
        palette.setColor(QPalette.HighlightedText, QColor("#fff9f4"))
        palette.setColor(QPalette.Button, QColor("#efe4d8"))
        palette.setColor(QPalette.ButtonText, QColor("#20170f"))
        palette.setColor(QPalette.WindowText, QColor("#20170f"))
        self._casting_table.setPalette(palette)
        self._casting_table.setStyleSheet(
            """
            QTableView {
                background: #fffdf9;
                alternate-background-color: #f6efe6;
                color: #20170f;
                gridline-color: #d7cabc;
                border: 1px solid #d7cabc;
                selection-background-color: #8f4b2b;
                selection-color: #fff9f4;
            }
            QTableView::item {
                padding: 4px 6px;
            }
            QTableView::item:selected {
                background: #8f4b2b;
                color: #fff9f4;
            }
            QTableView QLineEdit {
                background: #fffaf4;
                color: #20170f;
                border: 1px solid #c9b7a6;
                selection-background-color: #8f4b2b;
                selection-color: #fff9f4;
            }
            QHeaderView::section {
                background: #efe4d8;
                color: #20170f;
                border: 0;
                border-right: 1px solid #d7cabc;
                border-bottom: 1px solid #d7cabc;
                padding: 6px 8px;
                font-weight: 600;
            }
            QTableCornerButton::section {
                background: #efe4d8;
                border: 0;
                border-right: 1px solid #d7cabc;
                border-bottom: 1px solid #d7cabc;
            }
            """
        )

    def _configure_casting_table_columns(self) -> None:
        header = self._casting_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for column in range(1, self._casting_model.episode_count + 1):
            header.setSectionResizeMode(column, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(self._casting_model.inputs_column, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(self._casting_model.replicas_column, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(self._casting_model.actor_column, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(self._casting_model.note_column, QHeaderView.Stretch)

    def _create_metric_card(self, label_text: str) -> tuple[QFrame, QLabel]:
        frame = QFrame()
        frame.setProperty("metricCard", True)

        label = QLabel(label_text)
        label.setProperty("metricLabel", True)

        value = QLabel("0")
        value.setObjectName("metricValue")

        layout = QVBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(label)
        layout.addWidget(value)
        return frame, value

    def _connect_episode_editor(self, editor: EpisodeEditorWidget) -> None:
        editor.contentChanged.connect(self._handle_episode_changed)
        editor.importRequested.connect(self.import_episode_file)
        editor.clearRequested.connect(self.clear_episode)

    def _rebuild_episode_tabs(self, episodes: list[dict], selected_episode_index: int | None = None) -> None:
        target_index = self._episode_tabs.currentIndex() if selected_episode_index is None else selected_episode_index

        blocker = QSignalBlocker(self._episode_tabs)
        while self._episode_tabs.count():
            widget = self._episode_tabs.widget(0)
            self._episode_tabs.removeTab(0)
            if widget is not None:
                widget.deleteLater()
        self._episode_editors = []

        for index, episode in enumerate(episodes):
            label = str(episode.get("label", f"{index + 1:02d}")) or f"{index + 1:02d}"
            editor = EpisodeEditorWidget(index, label, self)
            editor.set_content(str(episode.get("content", "")))
            self._connect_episode_editor(editor)
            self._episode_editors.append(editor)
            self._episode_tabs.addTab(editor, label)

        del blocker

        if self._episode_editors:
            safe_index = max(0, min(target_index, len(self._episode_editors) - 1))
            self._episode_tabs.setCurrentIndex(safe_index)
        self._update_episode_controls()

    def _populate_from_state(self, selected_episode_index: int | None = None) -> None:
        payload = self._state.payload
        title_blocker = QSignalBlocker(self._title_edit)
        self._title_edit.setText(str(payload.get("title", "")))
        del title_blocker
        export_mode_blocker = QSignalBlocker(self._herci_by_episode_checkbox)
        self._herci_by_episode_checkbox.setChecked(self._state.herci_by_episode_export)
        del export_mode_blocker

        episodes = list(payload.get("episodes", []))
        self._rebuild_episode_tabs(episodes, selected_episode_index)
        self._update_window_title()

    def _update_episode_controls(self, current_index: int | None = None) -> None:
        has_episode = self._episode_tabs.count() > 0
        self._rename_episode_button.setEnabled(has_episode)
        self._add_episode_button.setEnabled(self._state.episode_count < MAX_EPISODES)
        self._remove_episode_button.setEnabled(self._state.episode_count > 1)

    def _handle_title_changed(self, text: str) -> None:
        self._state.set_title(text)
        self._update_window_title()

    def _handle_herci_export_mode_changed(self, checked: bool) -> None:
        self._state.set_herci_by_episode_export(checked)
        self._update_window_title()

    def _handle_casting_search_changed(self, text: str) -> None:
        self._casting_proxy_model.set_search_text(text)

    def _handle_casting_filter_changed(self, current_index: int) -> None:
        mode = str(self._casting_filter_combo.itemData(current_index) or CastingFilterProxyModel.FILTER_ALL)
        self._casting_proxy_model.set_assignment_filter(mode)

    def _handle_episode_changed(self, episode_index: int, content: str) -> None:
        self._state.set_episode_content(episode_index, content)
        self._update_window_title()
        self._analysis_timer.start()

    def _handle_assignment_edited(self, character: str, field: str, value: str) -> None:
        self._state.set_assignment(character, field, value)
        self._update_window_title()
        self.refresh_analysis()

    def _default_root_splitter_sizes(self) -> list[int]:
        total = max(self.width(), self.minimumWidth())
        left = max(360, total // 3)
        return [left, max(640, total - left)]

    def _default_right_splitter_sizes(self) -> list[int]:
        total = max(self.height(), self.minimumHeight())
        komplet = max(360, (total * 5) // 9)
        summary = max(160, (total * 2) // 9)
        validation = max(160, total - komplet - summary)
        return [komplet, summary, validation]

    def _has_visible_splitter_sizes(self, sizes: list[int] | None) -> bool:
        return bool(sizes and any(size > 0 for size in sizes))

    def _restore_splitter_sizes(self) -> None:
        if self._root_splitter is not None:
            root_sizes = (
                self._saved_root_splitter_sizes
                if self._has_visible_splitter_sizes(self._saved_root_splitter_sizes)
                else self._default_root_splitter_sizes()
            )
            self._root_splitter.setSizes(root_sizes)
        if self._right_splitter is not None:
            right_sizes = (
                self._saved_right_splitter_sizes
                if self._has_visible_splitter_sizes(self._saved_right_splitter_sizes)
                else self._default_right_splitter_sizes()
            )
            self._right_splitter.setSizes(right_sizes)

    def _set_komplet_focus_mode(self, enabled: bool) -> None:
        if not all(
            (
                self._editor_group,
                self._summary_group,
                self._validation_group,
                self._root_splitter,
                self._right_splitter,
            )
        ):
            return

        if enabled:
            self._saved_root_splitter_sizes = self._root_splitter.sizes()
            self._saved_right_splitter_sizes = self._right_splitter.sizes()
            self._editor_group.hide()
            self._summary_group.hide()
            self._validation_group.hide()
            self._root_splitter.setSizes([0, max(1, self._default_root_splitter_sizes()[1])])
            self._right_splitter.setSizes([max(1, self._default_right_splitter_sizes()[0]), 0, 0])
            self.statusBar().showMessage(t("status.komplet_on"), 3000)
            return

        self._editor_group.show()
        self._summary_group.show()
        self._validation_group.show()
        self._restore_splitter_sizes()
        self.statusBar().showMessage(t("status.komplet_off"), 3000)

    def _update_window_title(self) -> None:
        project_name = self._state.project_path.name if self._state.project_path else t("project.unsaved")
        dirty_marker = "*" if self._state.dirty else ""
        self.setWindowTitle(f"{dirty_marker}{self._state.title or t('project.default_title')} · {project_name}")

    def _set_summary_status(self, text: str, level: str) -> None:
        self._status_label.setProperty("statusLevel", level)
        self._status_label.setText(text)
        self._status_label.style().unpolish(self._status_label)
        self._status_label.style().polish(self._status_label)

    def _selected_validation(self) -> dict | None:
        current_index = self._validation_table.currentIndex()
        if not current_index.isValid():
            return None
        return self._validation_model.validation_at(current_index.row())

    def _update_validation_actions(self, *_args) -> None:
        validation = self._selected_validation()
        actionable = bool(validation and validation.get("kind") == "actor_variants" and validation.get("actionable"))
        self._validation_action_button.setEnabled(actionable)

    def _update_metrics(self, analysis: dict) -> None:
        stats = dict(analysis.get("stats", {}))
        missing = dict(analysis.get("missing", {}))
        self._metric_values["characterCount"].setText(str(stats.get("characterCount", 0)))
        self._metric_values["inputs"].setText(str(stats.get("inputs", 0)))
        self._metric_values["replicas"].setText(str(stats.get("replicas", 0)))
        self._metric_values["missing"].setText(str(missing.get("characters", 0)))

    def _apply_analysis(self, analysis: dict) -> None:
        self._casting_model.set_analysis(analysis)
        self._configure_casting_table_columns()
        self._actor_model.set_analysis(analysis)
        self._validation_model.set_analysis(analysis)
        self._validation_table.clearSelection()
        self._update_validation_actions()
        self._update_metrics(analysis)

        missing = dict(analysis.get("missing", {}))
        validation_summary = dict(analysis.get("validationSummary", {}))
        warning_count = int(validation_summary.get("warningCount", 0))
        info_count = int(validation_summary.get("infoCount", 0))

        if int(missing.get("characters", 0)) > 0:
            validation_suffix = ""
            if warning_count or info_count:
                validation_suffix = t("summary.missing.also", warnings=warning_count)
                if info_count:
                    validation_suffix += t("summary.missing.also.info", info=info_count)
                else:
                    validation_suffix += t("summary.missing.also.no_info")
            self._set_summary_status(
                t("summary.missing",
                  chars=missing.get("characters", 0),
                  inputs=missing.get("inputs", 0),
                  replicas=missing.get("replicas", 0)) + validation_suffix,
                "warning",
            )
        elif warning_count > 0:
            suffix = t("summary.validation.warnings.suffix", info=info_count) if info_count else ""
            self._set_summary_status(t("summary.validation.warnings", warnings=warning_count, suffix=suffix), "warning")
        elif info_count > 0:
            self._set_summary_status(t("summary.validation.info_only", info=info_count), "info")
        else:
            self._set_summary_status(t("summary.ok"), "ok")

        self.statusBar().showMessage(t("status.analysis_current"), 3000)

    def unify_selected_validation_actor_name(self) -> None:
        validation = self._selected_validation()
        if not validation or validation.get("kind") != "actor_variants" or not validation.get("actionable"):
            return

        variants = [
            normalize_text(str(variant))
            for variant in validation.get("variants", [])
            if normalize_text(str(variant))
        ]
        if len(variants) < 2:
            QMessageBox.information(self, t("unify.info.title"), t("unify.info.not_enough"))
            return

        choice, accepted = QInputDialog.getItem(
            self,
            t("unify.title"),
            t("unify.label"),
            [*variants, t("unify.custom")],
            editable=False,
        )
        if not accepted:
            return

        target_name = choice
        if choice == t("unify.custom"):
            custom_name, custom_accepted = QInputDialog.getText(
                self,
                t("unify.custom.title"),
                t("unify.custom.label"),
                text=variants[0],
            )
            if not custom_accepted:
                return
            target_name = custom_name

        normalized_target = normalize_text(target_name)
        if not normalized_target:
            QMessageBox.warning(self, t("unify.fail.title"), t("unify.fail.empty"))
            return

        confirmation = QMessageBox.question(
            self,
            t("unify.confirm.title"),
            t("unify.confirm.msg", variants=", ".join(variants), name=normalized_target),
            QMessageBox.Yes | QMessageBox.Cancel,
            QMessageBox.Yes,
        )
        if confirmation != QMessageBox.Yes:
            return

        try:
            changed = self._state.unify_actor_variants(variants, normalized_target)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, t("unify.fail.title"), str(exc))
            return

        if changed <= 0:
            QMessageBox.information(self, t("unify.info.title"), t("unify.info.none_found"))
            return

        self.refresh_analysis()
        self._update_window_title()
        self.statusBar().showMessage(t("status.actor_unified", name=normalized_target), 4000)

    def refresh_analysis(self) -> None:
        try:
            analysis = self._state.recompute()
        except Exception as exc:  # noqa: BLE001
            self._set_summary_status(str(exc), "error")
            self.statusBar().showMessage(str(exc), 7000)
            return

        self._apply_analysis(analysis)
        self._update_window_title()

    def new_project(self) -> None:
        if not self._maybe_save():
            return
        self._state.reset()
        self._populate_from_state()
        self.refresh_analysis()
        self.statusBar().showMessage(t("status.new_project"), 3000)

    def open_project(self) -> None:
        if not self._maybe_save():
            return

        selected, _ = QFileDialog.getOpenFileName(
            self,
            t("dialog.open_project"),
            str(self._state.project_path.parent if self._state.project_path else Path.cwd()),
            t("filter.project_file"),
        )
        if not selected:
            return

        try:
            self._state.load_project(Path(selected))
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, t("error.open_project"), str(exc))
            return

        self._populate_from_state()
        self.refresh_analysis()
        self.statusBar().showMessage(t("status.project_loaded"), 3000)

    def save_project(self) -> bool:
        if self._state.project_path is None:
            return self.save_project_as()

        try:
            self._state.save_project()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, t("error.save_project"), str(exc))
            return False

        self._update_window_title()
        self.statusBar().showMessage(t("status.project_saved"), 3000)
        return True

    def save_project_as(self) -> bool:
        default_name = slugify_filename(self._state.title or "casting-project")
        base_dir = self._state.project_path.parent if self._state.project_path else Path.cwd()
        default_path = base_dir / f"{default_name}.json"
        selected, _ = QFileDialog.getSaveFileName(
            self,
            t("dialog.save_as"),
            str(default_path),
            t("filter.project_file"),
        )
        if not selected:
            return False

        target = Path(selected)
        if target.suffix.lower() != ".json":
            target = target.with_suffix(".json")

        try:
            self._state.save_project(target)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, t("error.save_project"), str(exc))
            return False

        self._update_window_title()
        self.statusBar().showMessage(t("status.project_saved_to", name=target.name), 3000)
        return True

    def add_episode(self) -> None:
        try:
            episode_index = self._state.add_episode()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, t("error.add_episode"), str(exc))
            return

        label = self._state.episode_label(episode_index)
        self._populate_from_state(episode_index)
        self.refresh_analysis()
        self._update_window_title()
        self.statusBar().showMessage(t("status.episode_added", label=label), 3000)

    def rename_current_episode(self) -> None:
        episode_index = self._episode_tabs.currentIndex()
        if episode_index < 0:
            return

        current_label = self._state.episode_label(episode_index)
        label, accepted = QInputDialog.getText(
            self,
            t("rename.title"),
            t("rename.label"),
            text=current_label,
        )
        if not accepted:
            return

        try:
            self._state.rename_episode(episode_index, label)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, t("error.rename_episode"), str(exc))
            return

        updated_label = self._state.episode_label(episode_index)
        self._populate_from_state(episode_index)
        self.refresh_analysis()
        self._update_window_title()
        self.statusBar().showMessage(t("status.episode_renamed", label=updated_label), 3000)

    def remove_current_episode(self) -> None:
        episode_index = self._episode_tabs.currentIndex()
        if episode_index < 0:
            return

        label = self._state.episode_label(episode_index)
        result = QMessageBox.question(
            self,
            t("confirm.remove_episode.title"),
            t("confirm.remove_episode", label=label),
            QMessageBox.Yes | QMessageBox.Cancel,
            QMessageBox.Cancel,
        )
        if result != QMessageBox.Yes:
            return

        try:
            self._state.remove_episode(episode_index)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, t("error.remove_episode"), str(exc))
            return

        next_index = max(0, min(episode_index, self._state.episode_count - 1))
        self._populate_from_state(next_index)
        self.refresh_analysis()
        self._update_window_title()
        self.statusBar().showMessage(t("status.episode_removed", label=label), 3000)

    def import_episode_file(self, episode_index: int) -> None:
        label = self._state.episode_label(episode_index)
        start_dir = self._state.project_path.parent if self._state.project_path else Path.cwd()
        selected, _ = QFileDialog.getOpenFileName(
            self,
            t("dialog.load_input", label=label),
            str(start_dir),
            t("filter.input_files"),
        )
        if not selected:
            return

        path = Path(selected)
        selected_sheet_name: str | None = None
        selected_sheet_label = ""
        if path.suffix.lower() == ".xlsx":
            try:
                options = list_episode_source_options(path)
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, t("error.load_input"), str(exc))
                return

            if len(options) > 1:
                items = [option.display_name for option in options]
                selection, accepted = QInputDialog.getItem(
                    self,
                    t("dialog.select_sheet"),
                    t("dialog.select_sheet.msg"),
                    items,
                    editable=False,
                )
                if not accepted:
                    return
                selected_option = next(option for option in options if option.display_name == selection)
            else:
                selected_option = options[0]

            selected_sheet_name = selected_option.sheet_name
            selected_sheet_label = selected_option.sheet_name

        try:
            content = self._state.import_episode_file(episode_index, path, sheet_name=selected_sheet_name)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, t("error.load_input"), str(exc))
            return

        self._episode_tabs.setCurrentIndex(episode_index)
        self._episode_editors[episode_index].set_content(content)
        self.refresh_analysis()
        self._update_window_title()
        if selected_sheet_label:
            self.statusBar().showMessage(t("status.episode_loaded_sheet", label=label, sheet=selected_sheet_label), 3000)
        else:
            self.statusBar().showMessage(t("status.episode_loaded", label=label), 3000)

    def bulk_import_episodes(self) -> None:
        start_index = self._episode_tabs.currentIndex()
        if start_index < 0:
            start_index = 0

        try:
            sources = self._choose_bulk_import_sources()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, t("error.bulk_import"), str(exc))
            return
        if not sources:
            return

        try:
            plan = self._state.preview_bulk_import(start_index, sources)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, t("error.bulk_import_warning"), str(exc))
            return

        if not self._confirm_bulk_import(plan):
            return

        try:
            imported_indexes = self._state.apply_bulk_import(plan)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, t("error.bulk_import"), str(exc))
            return

        selected_index = imported_indexes[0] if imported_indexes else start_index
        self._populate_from_state(selected_index)
        self.refresh_analysis()
        self._update_window_title()
        self.statusBar().showMessage(t("status.bulk_imported", count=len(plan)), 4000)

    def _choose_bulk_import_sources(self) -> list[BulkImportSource]:
        opt_workbook = t("bulk.source.workbook")
        opt_files = t("bulk.source.files")
        mode, accepted = QInputDialog.getItem(
            self,
            t("dialog.bulk_import"),
            t("dialog.bulk_import.msg"),
            [opt_workbook, opt_files, t("bulk.source.dir")],
            editable=False,
        )
        if not accepted:
            return []

        start_dir = self._state.project_path.parent if self._state.project_path else Path.cwd()
        if mode == opt_workbook:
            selected, _ = QFileDialog.getOpenFileName(
                self,
                t("dialog.select_workbook"),
                str(start_dir),
                t("filter.excel"),
            )
            if not selected:
                return []
            return read_bulk_import_sources_from_workbook(Path(selected))

        if mode == opt_files:
            selected, _ = QFileDialog.getOpenFileNames(
                self,
                t("dialog.select_files"),
                str(start_dir),
                t("filter.input_files"),
            )
            if not selected:
                return []
            return read_bulk_import_sources_from_files([Path(item) for item in selected])

        selected_directory = QFileDialog.getExistingDirectory(
            self,
            t("dialog.select_dir"),
            str(start_dir),
        )
        if not selected_directory:
            return []
        return read_bulk_import_sources_from_directory(Path(selected_directory))

    def _confirm_bulk_import(self, plan: list[BulkImportPlanEntry]) -> bool:
        first_target = plan[0].target_label
        overwrite_count = sum(1 for entry in plan if entry.overwrites_content)
        created_count = sum(1 for entry in plan if entry.creates_episode)

        informative_parts = [t("bulk.plan.starts", label=first_target, count=len(plan))]
        if overwrite_count:
            informative_parts.append(t("bulk.plan.overwrites", count=overwrite_count))
        if created_count:
            informative_parts.append(t("bulk.plan.creates", count=created_count))

        preview_lines = []
        for entry in plan:
            target_number = entry.target_index + 1
            suffix_parts = [t("bulk.plan.episode", num=target_number, label=entry.target_label)]
            if entry.overwrites_content:
                suffix_parts.append(t("bulk.plan.overwrites_suffix"))
            elif entry.creates_episode:
                suffix_parts.append(t("bulk.plan.creates_suffix"))
            preview_lines.append(f"{entry.source_name} -> {' · '.join(suffix_parts)}")

        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Question)
        dialog.setWindowTitle(t("dialog.bulk_import.confirm_title"))
        dialog.setText(t("dialog.bulk_import.confirm_msg"))
        dialog.setInformativeText("\n".join([*informative_parts, "", *preview_lines]))
        import_button = dialog.addButton(t("dialog.bulk_import.btn"), QMessageBox.AcceptRole)
        dialog.addButton(QMessageBox.Cancel)
        dialog.setDefaultButton(import_button)
        dialog.exec()
        return dialog.clickedButton() is import_button

    def clear_episode(self, episode_index: int) -> None:
        label = self._state.episode_label(episode_index)
        self._state.clear_episode(episode_index)
        self._episode_editors[episode_index].set_content("")
        self.refresh_analysis()
        self._update_window_title()
        self.statusBar().showMessage(t("status.episode_cleared", label=label), 3000)

    def export_workbook(self) -> None:
        default_name = slugify_filename(self._state.title or "casting")
        start_dir = self._state.project_path.parent if self._state.project_path else Path.cwd()
        selected, _ = QFileDialog.getSaveFileName(
            self,
            t("dialog.export_xlsx"),
            str(start_dir / f"{default_name}.xlsx"),
            t("filter.excel"),
        )
        if not selected:
            return

        target = Path(selected)
        if target.suffix.lower() != ".xlsx":
            target = target.with_suffix(".xlsx")

        try:
            workbook = self._state.export_workbook()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, t("error.export"), str(exc))
            return

        try:
            target.write_bytes(workbook)
        except OSError as exc:
            QMessageBox.critical(self, t("error.export"), str(exc))
            return

        self._apply_analysis(self._state.analysis)
        self.statusBar().showMessage(t("status.export_done", name=target.name), 4000)

    def _maybe_save(self) -> bool:
        if not self._state.dirty:
            return True

        message_box = QMessageBox(self)
        message_box.setIcon(QMessageBox.Warning)
        message_box.setWindowTitle(t("unsaved.title"))
        message_box.setText(t("unsaved.text"))
        message_box.setInformativeText(t("unsaved.info"))
        message_box.setStandardButtons(
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
        )
        message_box.setDefaultButton(QMessageBox.Save)
        result = message_box.exec()

        if result == QMessageBox.Save:
            return self.save_project()
        if result == QMessageBox.Cancel:
            return False
        return True

    def _update_lang_action_checks(self) -> None:
        lang = get_language()
        self._lang_cs_action.setChecked(lang == "cs")
        self._lang_en_action.setChecked(lang == "en")

    def _switch_language(self, lang: str) -> None:
        if get_language() == lang:
            return
        set_language(lang)
        self._rebuild_ui_texts()
        self.refresh_analysis()

    def _rebuild_ui_texts(self) -> None:
        # Actions
        self._new_action.setText(t("action.new_project"))
        self._open_action.setText(t("action.open_project"))
        self._bulk_import_action.setText(t("action.bulk_import"))
        self._save_action.setText(t("action.save_project"))
        self._save_as_action.setText(t("action.save_as"))
        self._recalculate_action.setText(t("action.recalculate"))
        self._export_action.setText(t("action.export_xlsx"))
        self._komplet_focus_action.setText(t("action.komplet_focus"))
        self._komplet_focus_action.setToolTip(t("action.komplet_focus.tip"))
        self._quit_action.setText(t("action.quit"))
        self._lang_cs_action.setText(t("action.lang_cs"))
        self._lang_en_action.setText(t("action.lang_en"))
        self._update_lang_action_checks()
        # Menus
        self._file_menu.setTitle(t("menu.file"))
        self._tools_menu.setTitle(t("menu.tools"))
        self._lang_menu.setTitle(t("menu.language"))
        # Labels & buttons
        self._title_label.setText(t("label.project_name"))
        self._title_edit.setPlaceholderText(t("title.placeholder"))
        self._herci_by_episode_checkbox.setText(t("checkbox.herci_by_episode"))
        self._herci_by_episode_checkbox.setToolTip(t("checkbox.herci_by_episode.tip"))
        self._search_label.setText(t("label.search"))
        self._show_label.setText(t("label.show"))
        self._casting_search_edit.setPlaceholderText(t("search.placeholder"))
        # Filter combo – update text keeping the data values
        self._casting_filter_combo.setItemText(0, t("filter.all"))
        self._casting_filter_combo.setItemText(1, t("filter.unassigned"))
        self._casting_filter_combo.setItemText(2, t("filter.assigned"))
        # Buttons
        self._add_episode_button.setText(t("btn.add_episode"))
        self._rename_episode_button.setText(t("btn.rename_episode"))
        self._remove_episode_button.setText(t("btn.remove_episode"))
        self._validation_action_button.setText(t("btn.unify_name"))
        self._validation_action_button.setToolTip(t("btn.unify_name.tip"))
        # Group boxes
        if self._editor_group:
            self._editor_group.setTitle(t("group.episodes"))
        if self._komplet_group:
            self._komplet_group.setTitle(t("group.komplet"))
        if self._summary_group:
            self._summary_group.setTitle(t("group.actor_summary"))
        if self._validation_group:
            self._validation_group.setTitle(t("group.validation"))
        # Window title
        self._update_window_title()
        # Episode editor widgets
        for editor in self._episode_editors:
            editor.retranslate()

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._maybe_save():
            event.accept()
            return
        event.ignore()
