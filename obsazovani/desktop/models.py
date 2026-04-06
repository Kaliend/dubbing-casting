from __future__ import annotations

from typing import Any

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QSortFilterProxyModel, Qt, Signal
from PySide6.QtGui import QBrush, QColor

from obsazovani.i18n import t

CASTING_SORT_ROLE = Qt.UserRole + 1
CASTING_FILTER_TEXT_ROLE = Qt.UserRole + 2
CASTING_ASSIGNED_ROLE = Qt.UserRole + 3


class CastingTableModel(QAbstractTableModel):
    assignmentEdited = Signal(str, str, str)

    def __init__(self, parent: object | None = None) -> None:
        super().__init__(parent)
        self._rows: list[dict[str, Any]] = []
        self._episode_labels = ["01"]

    @property
    def episode_count(self) -> int:
        return max(1, len(self._episode_labels))

    @property
    def inputs_column(self) -> int:
        return 1 + self.episode_count

    @property
    def replicas_column(self) -> int:
        return self.inputs_column + 1

    @property
    def actor_column(self) -> int:
        return self.inputs_column + 2

    @property
    def note_column(self) -> int:
        return self.inputs_column + 3

    def set_analysis(self, analysis: dict[str, Any]) -> None:
        self.beginResetModel()
        self._rows = list(analysis.get("complete", []))
        episodes = list(analysis.get("episodes", []))
        self._episode_labels = [
            str(episode.get("label", f"{index + 1:02d}")) or f"{index + 1:02d}"
            for index, episode in enumerate(episodes)
        ] or ["01"]
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return 1 + self.episode_count + 4

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Vertical:
            return section + 1

        if section == 0:
            return t("col.character")
        if 1 <= section <= self.episode_count:
            return self._episode_labels[section - 1]
        if section == self.inputs_column:
            return t("col.inputs")
        if section == self.replicas_column:
            return t("col.replicas")
        if section == self.actor_column:
            return t("col.actor")
        if section == self.note_column:
            return t("col.note")
        return None

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        row = self._rows[index.row()]
        column = index.column()

        if role in (Qt.DisplayRole, Qt.EditRole):
            if column == 0:
                return row.get("character", "")
            if 1 <= column <= self.episode_count:
                episodes = row.get("episodes", [])
                return episodes[column - 1].get("display", "0 / 0") if column - 1 < len(episodes) else "0 / 0"
            if column == self.inputs_column:
                return int(row.get("totalInputs", 0))
            if column == self.replicas_column:
                return int(row.get("totalReplicas", 0))
            if column == self.actor_column:
                return row.get("actor", "")
            if column == self.note_column:
                return row.get("note", "")

        if role == Qt.TextAlignmentRole:
            if 1 <= column <= self.replicas_column:
                return int(Qt.AlignCenter)
            return int(Qt.AlignVCenter | Qt.AlignLeft)

        if role == CASTING_SORT_ROLE:
            return self._sort_value(row, column)

        if role == CASTING_FILTER_TEXT_ROLE:
            return self._filter_text(row)

        if role == CASTING_ASSIGNED_ROLE:
            return bool(str(row.get("actor", "")).strip())

        if role == Qt.BackgroundRole and not row.get("actor"):
            return QBrush(QColor("#fff2e0"))

        if role == Qt.ToolTipRole and not row.get("actor"):
            return t("tooltip.unassigned")

        return None

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        if not index.isValid():
            return Qt.NoItemFlags
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if index.column() in (self.actor_column, self.note_column):
            flags |= Qt.ItemIsEditable
        return flags

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.EditRole) -> bool:
        if role != Qt.EditRole or not index.isValid() or index.column() not in (self.actor_column, self.note_column):
            return False

        row = self._rows[index.row()]
        field = "actor" if index.column() == self.actor_column else "note"
        text = str(value or "")
        if str(row.get(field, "")) == text:
            return False

        row[field] = text
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])
        self.assignmentEdited.emit(str(row.get("character", "")), field, text)
        return True

    def _sort_value(self, row: dict[str, Any], column: int) -> Any:
        if column == 0:
            return str(row.get("character", "")).casefold()

        if 1 <= column <= self.episode_count:
            episodes = row.get("episodes", [])
            if column - 1 < len(episodes):
                episode = episodes[column - 1]
                return (
                    int(episode.get("inputs", 0)),
                    int(episode.get("replicas", 0)),
                    str(row.get("character", "")).casefold(),
                )
            return (0, 0, str(row.get("character", "")).casefold())

        if column == self.inputs_column:
            return int(row.get("totalInputs", 0))

        if column == self.replicas_column:
            return int(row.get("totalReplicas", 0))

        if column == self.actor_column:
            return str(row.get("actor", "")).casefold()

        if column == self.note_column:
            return str(row.get("note", "")).casefold()

        return ""

    def _filter_text(self, row: dict[str, Any]) -> str:
        return " ".join(
            [
                str(row.get("character", "")).casefold(),
                str(row.get("actor", "")).casefold(),
                str(row.get("note", "")).casefold(),
            ]
        ).strip()


class CastingFilterProxyModel(QSortFilterProxyModel):
    FILTER_ALL = "all"
    FILTER_ASSIGNED = "assigned"
    FILTER_UNASSIGNED = "unassigned"

    def __init__(self, parent: object | None = None) -> None:
        super().__init__(parent)
        self._search_text = ""
        self._assignment_filter = self.FILTER_ALL
        self.setDynamicSortFilter(True)
        self.setSortCaseSensitivity(Qt.CaseInsensitive)

    def set_search_text(self, text: str) -> None:
        normalized = str(text or "").casefold().strip()
        if self._search_text == normalized:
            return
        self.beginFilterChange()
        self._search_text = normalized
        self.endFilterChange(QSortFilterProxyModel.Direction.Rows)

    def set_assignment_filter(self, mode: str) -> None:
        normalized = mode if mode in {self.FILTER_ALL, self.FILTER_ASSIGNED, self.FILTER_UNASSIGNED} else self.FILTER_ALL
        if self._assignment_filter == normalized:
            return
        self.beginFilterChange()
        self._assignment_filter = normalized
        self.endFilterChange(QSortFilterProxyModel.Direction.Rows)

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        model = self.sourceModel()
        if model is None:
            return True

        source_index = model.index(source_row, 0, source_parent)
        assigned = bool(model.data(source_index, CASTING_ASSIGNED_ROLE))

        if self._assignment_filter == self.FILTER_ASSIGNED and not assigned:
            return False
        if self._assignment_filter == self.FILTER_UNASSIGNED and assigned:
            return False

        if not self._search_text:
            return True

        haystack = str(model.data(source_index, CASTING_FILTER_TEXT_ROLE) or "")
        return self._search_text in haystack

    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:
        left_value = left.data(CASTING_SORT_ROLE)
        right_value = right.data(CASTING_SORT_ROLE)

        if left_value is None:
            return right_value is not None
        if right_value is None:
            return False

        return left_value < right_value


class ActorSummaryTableModel(QAbstractTableModel):
    def __init__(self, parent: object | None = None) -> None:
        super().__init__(parent)
        self._rows: list[dict[str, Any]] = []

    def set_analysis(self, analysis: dict[str, Any]) -> None:
        self.beginResetModel()
        self._rows = list(analysis.get("actors", []))
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return 3

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Vertical:
            return section + 1
        return [t("col.actor"), t("col.inputs"), t("col.replicas")][section]

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        row = self._rows[index.row()]
        column = index.column()

        if role == Qt.DisplayRole:
            if column == 0:
                return row.get("actor", "")
            if column == 1:
                return int(row.get("totalInputs", 0))
            if column == 2:
                return int(row.get("totalReplicas", 0))

        if role == Qt.TextAlignmentRole:
            if column == 0:
                return int(Qt.AlignVCenter | Qt.AlignLeft)
            return int(Qt.AlignCenter)

        return None


class ValidationTableModel(QAbstractTableModel):
    def __init__(self, parent: object | None = None) -> None:
        super().__init__(parent)
        self._rows: list[dict[str, Any]] = []

    def set_analysis(self, analysis: dict[str, Any]) -> None:
        self.beginResetModel()
        self._rows = list(analysis.get("validations", []))
        self.endResetModel()

    def validation_at(self, row_index: int) -> dict[str, Any] | None:
        if not 0 <= row_index < len(self._rows):
            return None
        return dict(self._rows[row_index])

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return 3

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Vertical:
            return section + 1
        return [t("col.level"), t("col.area"), t("col.detail")][section]

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        row = self._rows[index.row()]
        severity = str(row.get("severity", "info"))
        column = index.column()

        if role == Qt.DisplayRole:
            if column == 0:
                return t("severity.warning") if severity == "warning" else t("severity.info")
            if column == 1:
                return row.get("category", "")
            if column == 2:
                return row.get("message", "")

        if role == Qt.TextAlignmentRole:
            if column == 2:
                return int(Qt.AlignVCenter | Qt.AlignLeft)
            return int(Qt.AlignCenter)

        if role == Qt.BackgroundRole:
            if severity == "warning":
                return QBrush(QColor("#fff2e0"))
            return QBrush(QColor("#eef4ff"))

        if role == Qt.ForegroundRole:
            if severity == "warning":
                return QBrush(QColor("#7a4d00"))
            return QBrush(QColor("#1f3d68"))

        if role == Qt.ToolTipRole:
            return row.get("message", "")

        if role == Qt.UserRole:
            return dict(row)

        return None
