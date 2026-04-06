from __future__ import annotations

from PySide6.QtCore import QSignalBlocker, Signal
from PySide6.QtWidgets import QHBoxLayout, QLabel, QPushButton, QPlainTextEdit, QVBoxLayout, QWidget

from obsazovani.i18n import t


class EpisodeEditorWidget(QWidget):
    contentChanged = Signal(int, str)
    importRequested = Signal(int)
    clearRequested = Signal(int)

    def __init__(self, episode_index: int, episode_label: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._episode_index = episode_index

        intro_label = QLabel(t("editor.intro"))
        intro_label.setWordWrap(True)

        self._load_button = QPushButton(t("editor.load"))
        self._clear_button = QPushButton(t("editor.clear"))
        self._editor = QPlainTextEdit()
        self._editor.setPlaceholderText(t("editor.placeholder", label=episode_label))

        action_row = QHBoxLayout()
        action_row.addWidget(self._load_button)
        action_row.addWidget(self._clear_button)
        action_row.addStretch(1)

        layout = QVBoxLayout(self)
        layout.addLayout(action_row)
        layout.addWidget(self._editor, 1)
        layout.addWidget(intro_label)

        self._editor.textChanged.connect(self._emit_content)
        self._load_button.clicked.connect(lambda: self.importRequested.emit(self._episode_index))
        self._clear_button.clicked.connect(lambda: self.clearRequested.emit(self._episode_index))

    def content(self) -> str:
        return self._editor.toPlainText()

    def set_content(self, content: str) -> None:
        blocker = QSignalBlocker(self._editor)
        self._editor.setPlainText(content)
        del blocker

    def retranslate(self) -> None:
        self._load_button.setText(t("editor.load"))
        self._clear_button.setText(t("editor.clear"))
        episode_label = self._editor.placeholderText().split(":")[0]
        self._editor.setPlaceholderText(t("editor.placeholder", label=episode_label))

    def _emit_content(self) -> None:
        self.contentChanged.emit(self._episode_index, self._editor.toPlainText())
