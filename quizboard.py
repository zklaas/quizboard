#!/usr/bin/env python3
"""Quizboard — a 6 × 6 board of point values, questions and answers.

Each column has a topic; each tile shows its points, 1 to 6 going down. Click a
tile to show its question, click the question to show the answer, and click
again to go back to the board, where that tile is now blank.

Each game starts by asking how many teams are playing. Once an answer is shown,
a checkbox per team appears: tick the team(s) that earn the tile's points.
Game ▸ Scores shows the standings at any time, and the final scores appear when
the last tile has been played.

Bilingual mode (View ▸ Bilingual Mode) shows two language versions of the board
side by side — board_en.xlsx and board_fr.xlsx — with each English label and its
French twin at the same size. View ▸ Language switches the interface between
English and French; all text is in i18n.py.

The questions come from an .xlsx workbook — the layout is described in
board_io.py, and File ▸ Save Blank Template… writes a ready-made one.

Launch:
    python3 quizboard.py
    python3 quizboard.py my_board.xlsx
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

from PyQt6.QtCore import QLibraryInfo, QLocale, QRect, Qt, QTimer, QTranslator, QUrl, pyqtSignal
from PyQt6.QtGui import QAction, QActionGroup, QDesktopServices, QFont, QFontMetrics, QIcon, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from board_io import (
    COLUMNS,
    ROWS,
    Board,
    BoardError,
    compare_boards,
    language_of,
    load_board,
    partner_path,
    write_template,
)
from i18n import LANGUAGE_NAMES, LANGUAGES, NB, Message, and_list, other, render, tr, tr_n

if os.name == "nt":                     # Windows: %APPDATA%\Quizboard
    CONFIG_PATH = Path(os.environ.get("APPDATA") or Path.home()) / "Quizboard" / "config.json"
elif sys.platform == "darwin":          # macOS: ~/Library/Application Support/Quizboard
    CONFIG_PATH = Path.home() / "Library" / "Application Support" / "Quizboard" / "config.json"
else:                                  # Linux: the XDG config folder
    CONFIG_PATH = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "quizboard" / "config.json"
XLSX_SUFFIXES = (".xlsx", ".xlsm")
MAX_TEAMS = 8

BACKGROUND = "#0b1233"
TOPIC_FILL = "#16286b"
TILE = "#1f3d99"
TILE_HOVER = "#2b52c4"
TILE_PLAYED = "#121b45"
TILE_MISSING = "#151a2e"
RULE = "#3a4f9c"
GOLD = "#ffcc33"
DIM = "#8a93b8"

STYLE = f"""
QWidget#page {{ background: {BACKGROUND}; }}
QWidget#page QLabel {{ color: white; }}
QWidget#page QFrame#topic {{ background: {TOPIC_FILL}; border-radius: 6px; }}
QWidget#page QLabel#topicText {{ font-weight: bold; }}
QWidget#page QFrame#rule {{ background: {RULE}; border: none; }}
QWidget#page QLabel#heading {{ color: {DIM}; }}
QWidget#page QLabel#answer {{ color: {GOLD}; }}
QWidget#page QLabel#title {{ color: {GOLD}; font-weight: bold; }}
QWidget#page QLabel#leader {{ color: {GOLD}; font-weight: bold; }}
QWidget#page QLabel#hint {{ color: {DIM}; }}
QWidget#page QLabel#missing {{ color: {DIM}; font-style: italic; }}
QWidget#page QCheckBox {{ color: white; spacing: 12px; padding: 4px 14px; }}
QWidget#page QCheckBox:checked {{ color: {GOLD}; }}
QWidget#page QCheckBox::indicator {{
    width: 22px; height: 22px; border: 2px solid {DIM}; border-radius: 5px; background: transparent;
}}
QWidget#page QCheckBox::indicator:checked {{ background: {GOLD}; border-color: {GOLD}; }}
QPushButton#tile {{ border: none; border-radius: 6px; font-weight: bold; }}
QPushButton#tile[state="ready"] {{ background: {TILE}; color: {GOLD}; }}
QPushButton#tile[state="ready"]:hover {{ background: {TILE_HOVER}; }}
QPushButton#tile[state="played"] {{ background: {TILE_PLAYED}; color: {TILE_PLAYED}; }}
QPushButton#tile[state="missing"] {{ background: {TILE_MISSING}; color: #3a4266; }}
QPushButton#action {{
    background: {TOPIC_FILL}; color: white; border: 1px solid {RULE}; border-radius: 6px;
    padding: 8px 18px;
}}
QPushButton#action:hover {{ background: {TILE_HOVER}; }}
QPushButton#action:disabled {{ color: #59618a; border-color: #2a3466; }}
QPushButton#undo {{
    background: {GOLD}; color: {BACKGROUND}; border: none; border-radius: 6px;
    padding: 10px 22px; font-weight: bold; font-size: 15px;
}}
QPushButton#undo:hover {{ background: #ffe08a; }}
"""


def load_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_config(config: dict) -> None:
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(config, indent=2), encoding="utf-8")
    except OSError:
        pass   # remembering the last file is a convenience, never worth an error


def _action_button(text: str = "") -> QPushButton:
    button = QPushButton(text)
    button.setObjectName("action")
    button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    return button


def _undo_button(keep_space: bool) -> QPushButton:
    """The gold Undo button; the window sets its text and shows it while there is something to undo.
    With `keep_space`, its row keeps its height while it is hidden, so nothing jumps when it appears."""
    button = QPushButton("↶  Undo  (Ctrl+Z)")
    button.setObjectName("undo")
    button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    button.setCursor(Qt.CursorShape.PointingHandCursor)
    policy = button.sizePolicy()
    policy.setRetainSizeWhenHidden(keep_space)
    button.setSizePolicy(policy)
    button.hide()
    return button


def _rule(vertical: bool = False) -> QFrame:
    """A 1-pixel line between the two languages."""
    line = QFrame()
    line.setObjectName("rule")
    if vertical:
        line.setFixedWidth(1)
    else:
        line.setFixedHeight(1)
    return line


class Voice:
    """Which language(s) the app speaks.

    `ui` is the interface language: menus, dialogs and the status bar — what the
    host reads. `room` is what the players see on the question and score pages:
    the board's language, or in bilingual mode both languages, the left column's
    first.
    """

    def __init__(self, ui: str):
        self.ui = ui
        self.room = [ui]

    def host(self, key: str, **params) -> str:
        return tr(key, self.ui, **params)

    def host_n(self, key: str, n: int, **params) -> str:
        return tr_n(key, n, self.ui, **params)

    def show(self, key: str, sep: str = " / ", **params) -> str:
        return sep.join(tr(key, language, **params) for language in self.room)

    def show_each(self, make, sep: str = " / ") -> str:
        return sep.join(make(language) for language in self.room)


class FitLabel(QLabel):
    """Word-wrapped plain-text label whose font shrinks until the text fits its box.

    The size policy is Ignored and there is no height-for-width, so the font never
    feeds back into the layout: the layout decides the box, and the label fills it
    with the largest text that fits, up to `cap_ratio` × its own height.
    """

    def __init__(self, cap_ratio: float, name: str = "", min_px: int = 8, parent: QWidget | None = None):
        super().__init__(parent)
        self.cap_ratio = cap_ratio
        self.min_px = min_px
        self.group: FitGroup | None = None
        if name:
            self.setObjectName(name)
        self.setTextFormat(Qt.TextFormat.PlainText)
        self.setWordWrap(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        self.setMinimumSize(1, 1)

    def hasHeightForWidth(self) -> bool:   # noqa: N802 — Qt override
        # A word-wrapped QLabel reports height-for-width, and box layouts honour it as a
        # minimum even under an Ignored policy — the font would then grow its own box.
        return False

    def setText(self, text: str) -> None:   # noqa: N802 — Qt override
        super().setText(text)
        self._refit()

    def resizeEvent(self, event) -> None:   # noqa: N802 — Qt override
        super().resizeEvent(event)
        self._refit()

    def best_size(self) -> int | None:
        """The largest pixel size at which the text fits, or None with no text or no room."""
        rect = self.contentsRect()
        width, height = rect.width() - 2 * self.margin(), rect.height() - 2 * self.margin()
        if width < 4 or height < 4 or not self.text():
            return None
        flags = self.alignment().value | Qt.TextFlag.TextWordWrap.value
        font = QFont(self.font())
        lo, hi = self.min_px, max(self.min_px, int(height * self.cap_ratio))
        best = lo
        while lo <= hi:
            mid = (lo + hi) // 2
            font.setPixelSize(mid)
            box = QFontMetrics(font).boundingRect(QRect(0, 0, width, 1_000_000), flags, self.text())
            if box.width() <= width and box.height() <= height:
                best, lo = mid, mid + 1
            else:
                hi = mid - 1
        return best

    def apply_size(self, px: int) -> None:
        if self.font().pixelSize() != px:
            font = QFont(self.font())
            font.setPixelSize(px)
            self.setFont(font)

    def _refit(self) -> None:
        if self.group is not None:
            self.group.refit()
        else:
            best = self.best_size()
            if best is not None:
                self.apply_size(best)


class FitGroup:
    """Labels that share one font size: the largest at which every active one fits.

    Bilingual mode groups each label with its other-language twin, so neither
    language is ever shown smaller than the other.
    """

    def __init__(self, *labels: FitLabel):
        self.labels = list(labels)
        self.active = list(labels)
        for label in labels:
            label.group = self

    def set_active(self, labels: list[FitLabel]) -> None:
        self.active = list(labels)
        self.refit()

    def refit(self) -> None:
        sizes = [size for size in (label.best_size() for label in self.active) if size is not None]
        if sizes:
            for label in self.labels:
                label.apply_size(min(sizes))


class Tile(QPushButton):
    """One board square. state: 'ready' (shows points), 'played' (blank) or 'missing'."""

    def __init__(self, column: int, points: int):
        super().__init__()
        self.column, self.points = column, points
        self.setObjectName("tile")
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        self.setMinimumSize(20, 16)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.state = ""
        self.set_state("missing")

    def set_state(self, state: str) -> None:
        if state == self.state:
            return
        self.state = state
        self.setText(str(self.points) if state != "played" else "")
        self.setCursor(Qt.CursorShape.PointingHandCursor if state == "ready" else Qt.CursorShape.ArrowCursor)
        self.setProperty("state", state)
        self.style().unpolish(self)
        self.style().polish(self)

    def resizeEvent(self, event) -> None:   # noqa: N802 — Qt override
        super().resizeEvent(event)
        font = QFont(self.font())
        font.setPixelSize(max(8, int(min(self.height() * 0.55, self.width() * 0.4))))
        font.setBold(True)
        self.setFont(font)


class TopicCell(QFrame):
    """A column heading: the topic, and in bilingual mode its other-language name below a rule."""

    def __init__(self):
        super().__init__()
        self.setObjectName("topic")
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(3)
        self.primary = FitLabel(0.28, "topicText")
        self.rule = _rule()
        self.secondary = FitLabel(0.5, "topicText")
        layout.addWidget(self.primary, 1)
        layout.addWidget(self.rule)
        layout.addWidget(self.secondary, 1)
        self.fit = FitGroup(self.primary, self.secondary)
        self.set_bilingual(False)

    def set_bilingual(self, on: bool) -> None:
        self.rule.setVisible(on)
        self.secondary.setVisible(on)
        self.primary.cap_ratio = 0.5 if on else 0.28
        self.fit.set_active([self.primary, self.secondary] if on else [self.primary])


class BoardPage(QWidget):
    tile_clicked = pyqtSignal(int, int)          # column, points
    tile_menu = pyqtSignal(int, int, object)     # column, points, global position

    def __init__(self):
        super().__init__()
        self.setObjectName("page")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        page = QVBoxLayout(self)
        page.setContentsMargins(14, 14, 14, 10)
        page.setSpacing(8)
        grid = QGridLayout()
        grid.setSpacing(8)
        page.addLayout(grid, 1)
        self.cells: list[TopicCell] = []
        self.tiles: dict[tuple[int, int], Tile] = {}
        for c in range(COLUMNS):
            cell = TopicCell()
            grid.addWidget(cell, 0, c)
            self.cells.append(cell)
            grid.setColumnStretch(c, 1)
            for p in range(1, ROWS + 1):
                tile = Tile(c, p)
                tile.clicked.connect(lambda _=False, t=tile: self.tile_clicked.emit(t.column, t.points))
                tile.customContextMenuRequested.connect(
                    lambda pos, t=tile: self.tile_menu.emit(t.column, t.points, t.mapToGlobal(pos))
                )
                grid.addWidget(tile, p, c)
                self.tiles[(c, p)] = tile
        self.topics = [cell.primary for cell in self.cells]
        grid.setRowStretch(0, 13)
        for p in range(1, ROWS + 1):
            grid.setRowStretch(p, 10)
        # Undo stays up between questions, so it gets its own strip and never covers a tile.
        self.undo_button = _undo_button(keep_space=True)
        bar = QHBoxLayout()
        bar.addStretch(1)
        bar.addWidget(self.undo_button)
        page.addLayout(bar)

    def show_board(self, board: Board, state_of, voice: Voice, partner: Board | None = None) -> None:
        for c, cell in enumerate(self.cells):
            cell.set_bilingual(partner is not None)
            cell.primary.setText(board.topics[c] if c < len(board.topics) else "")
            if partner is not None:
                cell.secondary.setText(partner.topics[c] if c < len(partner.topics) else "")
        for (c, p), tile in self.tiles.items():
            clue = board.clue(c, p)
            tile.setToolTip(render(clue.where, voice.ui) if clue else voice.host("tile.missing"))
            tile.set_state(state_of(c, p))


class ClueColumn(QWidget):
    """One language's heading, question and answer on the question page."""

    def __init__(self):
        super().__init__()
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        self.heading = FitLabel(0.6, "heading")
        self.question = FitLabel(0.16)
        self.answer = FitLabel(0.3, "answer")
        layout.addWidget(self.heading, 1)
        layout.addWidget(self.question, 6)
        layout.addWidget(self.answer, 3)


class CluePage(QWidget):
    advance = pyqtSignal()                  # click / Space: show the answer, then back to the board
    back = pyqtSignal()                     # Esc / Back to Board
    award_toggled = pyqtSignal(int, bool)   # team index, ticked

    def __init__(self, voice: Voice):
        super().__init__()
        self.voice = voice
        self.setObjectName("page")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(48, 24, 48, 20)
        layout.setSpacing(12)

        # The board's own language on the left; in bilingual mode the other one on the right.
        self.sides = [ClueColumn(), ClueColumn()]
        self.divider = _rule(vertical=True)
        columns = QHBoxLayout()
        columns.setSpacing(28)
        columns.addWidget(self.sides[0], 1)
        columns.addWidget(self.divider)
        columns.addWidget(self.sides[1], 1)
        layout.addLayout(columns, 10)
        self.heading, self.question, self.answer = self.sides[0].heading, self.sides[0].question, self.sides[0].answer
        self.fits = [FitGroup(*(getattr(side, part) for side in self.sides)) for part in ("heading", "question", "answer")]

        # One checkbox per team, shown once the answer is. The row keeps its space while
        # hidden so revealing the answer doesn't shrink the question.
        self.team_row = QWidget()
        self.team_grid = QGridLayout(self.team_row)
        self.team_grid.setContentsMargins(0, 0, 0, 0)
        self.team_boxes: list[QCheckBox] = []
        self.hint = QLabel()
        self.hint.setObjectName("hint")
        self.hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint.setWordWrap(True)   # two lines in bilingual mode; never widens the window
        layout.addWidget(self.team_row)
        layout.addWidget(self.hint)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.show_answer_button = _action_button()
        self.back_button = _action_button()
        buttons.addWidget(self.show_answer_button)
        buttons.addWidget(self.back_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)
        self.show_answer_button.clicked.connect(self.advance)
        self.back_button.clicked.connect(self.back)
        self.revealed = False
        self._answers: list[tuple[str, str]] = []
        self.set_bilingual(False)

    def set_bilingual(self, on: bool) -> None:
        self.bilingual = on
        self.sides[1].setVisible(on)
        self.divider.setVisible(on)
        for fit in self.fits:
            fit.set_active(fit.labels if on else fit.labels[:1])

    def set_teams(self, names: list[str]) -> None:
        for box in self.team_boxes:
            self.team_grid.removeWidget(box)
            box.deleteLater()
        self.team_boxes = []
        per_row = len(names) if len(names) <= 4 else (len(names) + 1) // 2
        for i, name in enumerate(names):
            box = QCheckBox(name)
            box.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            box.setCursor(Qt.CursorShape.PointingHandCursor)
            box.clicked.connect(lambda checked, i=i: self.award_toggled.emit(i, checked))
            self.team_grid.addWidget(box, i // per_row, i % per_row, Qt.AlignmentFlag.AlignCenter)
            self.team_boxes.append(box)
        policy = self.team_row.sizePolicy()
        policy.setRetainSizeWhenHidden(bool(names))
        self.team_row.setSizePolicy(policy)
        self.team_row.setVisible(self.revealed and bool(names))
        self._fit_team_fonts()

    def set_team_labels(self, labels: list[str]) -> None:
        for box, text in zip(self.team_boxes, labels):
            box.setText(text)

    def show_clue(self, sides: list[tuple[str, str, str, str]], points: int, awarded: set[int]) -> None:
        """sides: (topic, question, answer, missing_from) — the board's language, then in bilingual
        mode the other one. `missing_from` names the file that has no version of this tile."""
        self.revealed = False
        self._answers = []
        for i, (column, (topic, question, answer, missing_from)) in enumerate(zip(self.sides, sides)):
            language = self.voice.room[min(i, len(self.voice.room) - 1)]
            amount = tr_n("points", points, language).replace(" ", NB)   # "5 points" never splits across lines
            column.heading.setText(f"{topic}   ·   {amount}" if topic else amount)
            if missing_from:
                no_question = no_answer = tr("clue.not_in", language, name=missing_from)
            else:
                no_question, no_answer = tr("clue.no_question", language), tr("clue.no_answer", language)
            self._set(column.question, question, no_question, "")
            column.answer.setText("")
            self._answers.append((answer, no_answer))
        for i, box in enumerate(self.team_boxes):
            box.setChecked(i in awarded)
        self.team_row.setVisible(False)
        self.show_answer_button.setText(self.voice.show("clue.show_answer"))
        self.back_button.setText(self.voice.show("clue.back"))
        self.show_answer_button.setEnabled(True)
        self.hint.setText(self.voice.show("clue.hint_question", sep="\n"))

    def reveal(self) -> None:
        self.revealed = True
        for column, (answer, placeholder) in zip(self.sides, self._answers):
            self._set(column.answer, answer, placeholder, "answer")
        self.show_answer_button.setEnabled(False)
        if self.team_boxes:
            self.team_row.setVisible(True)
            keys = "1" if len(self.team_boxes) == 1 else f"1–{len(self.team_boxes)}"
            self.hint.setText(self.voice.show("clue.hint_teams", sep="\n", keys=keys))
        else:
            self.hint.setText(self.voice.show("clue.hint_back", sep="\n"))

    def toggle_team(self, index: int) -> None:
        if self.revealed and 0 <= index < len(self.team_boxes):
            self.team_boxes[index].click()

    @staticmethod
    def _set(label: FitLabel, text: str, placeholder: str, name: str) -> None:
        label.setObjectName(name if text else "missing")
        label.style().unpolish(label)
        label.style().polish(label)
        label.setText(text or placeholder)

    def _fit_team_fonts(self) -> None:
        for box in self.team_boxes:
            font = QFont(box.font())
            font.setPixelSize(max(14, int(self.height() * 0.034)))
            font.setBold(True)
            box.setFont(font)

    def resizeEvent(self, event) -> None:   # noqa: N802 — Qt override
        super().resizeEvent(event)
        self._fit_team_fonts()

    def mousePressEvent(self, event) -> None:   # noqa: N802 — Qt override
        if event.button() == Qt.MouseButton.LeftButton:
            self.advance.emit()
        else:
            super().mousePressEvent(event)


class ScorePage(QWidget):
    """The standings: Game ▸ Scores at any time, and Final Scores when the board is cleared."""

    back = pyqtSignal()
    new_game = pyqtSignal()

    MIN_ROWS = 4   # two or three teams still get rows of a sensible height

    def __init__(self, voice: Voice):
        super().__init__()
        self.voice = voice
        self.setObjectName("page")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.page_layout = QVBoxLayout(self)
        self.page_layout.setContentsMargins(48, 24, 48, 20)
        self.page_layout.setSpacing(12)
        self.title = FitLabel(0.6, "title")
        self.subtitle = FitLabel(0.5, "heading")
        self.table = QWidget()
        self.page_layout.addWidget(self.title, 2)
        self.page_layout.addWidget(self.subtitle, 1)
        self.page_layout.addWidget(self.table, 9)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.back_button = _action_button()
        self.new_game_button = _action_button()
        self.undo_button = _undo_button(keep_space=False)
        buttons.addWidget(self.back_button)
        buttons.addWidget(self.new_game_button)
        buttons.addWidget(self.undo_button)
        buttons.addStretch(1)
        self.page_layout.addLayout(buttons)
        self.back_button.clicked.connect(self.back)
        self.new_game_button.clicked.connect(self.new_game)
        self.final = False

    def show_scores(self, teams: list[str], scores: list[int], final: bool, played: int, total: int) -> None:
        voice = self.voice
        self.final = final
        self.title.setText(voice.show("scores.final" if final else "scores.title"))
        self.back_button.setText(voice.show("scores.back"))
        self.new_game_button.setText(voice.show("scores.new_game"))
        table = QWidget()
        grid = QGridLayout(table)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(28)
        if not teams:
            self.subtitle.setText(voice.show("scores.no_teams", sep="\n"))
        else:
            top = max(scores)
            leaders = [name for name, score in zip(teams, scores) if score == top]
            if not final:
                self.subtitle.setText(voice.show("scores.progress", sep="\n", played=played, total=total))
            elif top == 0:
                self.subtitle.setText(voice.show("scores.none", sep="\n"))
            elif len(leaders) == 1:
                self.subtitle.setText(voice.show_each(
                    lambda lang: tr("scores.winner", lang, name=leaders[0], amount=tr_n("points", top, lang)), "\n"))
            else:
                self.subtitle.setText(voice.show_each(
                    lambda lang: tr("scores.tie", lang, amount=tr_n("points", top, lang),
                                    names=and_list(leaders, lang)), "\n"))
            order = sorted(range(len(teams)), key=lambda i: -scores[i])   # stable: ties keep team order
            for row, i in enumerate(order):
                rank = 1 + sum(1 for s in scores if s > scores[i])
                style = "leader" if scores[i] == top and top > 0 else ""
                cells = (
                    (f"{rank}.", Qt.AlignmentFlag.AlignRight),
                    (teams[i], Qt.AlignmentFlag.AlignLeft),
                    (str(scores[i]), Qt.AlignmentFlag.AlignRight),
                )
                for column, (text, align) in enumerate(cells):
                    label = FitLabel(0.62, style)
                    label.setAlignment(align | Qt.AlignmentFlag.AlignVCenter)
                    label.setText(text)
                    grid.addWidget(label, row, column)
            for row in range(max(len(teams), self.MIN_ROWS)):
                grid.setRowStretch(row, 1)
            grid.setColumnStretch(0, 1)
            grid.setColumnStretch(1, 6)
            grid.setColumnStretch(2, 2)
        self.page_layout.replaceWidget(self.table, table)
        self.table.deleteLater()
        self.table = table


class TeamsDialog(QDialog):
    """Asked at the start of each game: how many teams are playing, and their names."""

    # The window manager draws the title bar, so Qt can't size the dialog to it. WSLg's
    # compositor centres the title across the whole width with the minimise/maximise/close
    # buttons on the right; measured 2026-10-19, the title is drawn about as wide as in
    # Qt's font and needs ~76 px clear on each side — at 211 px "New Game" was clipped
    # to "New Gam". Width = title × 1.15 (a bolder title font) + this allowance.
    TITLE_BAR_ALLOWANCE = 180

    def __init__(self, parent: QWidget, names: list[str], voice: Voice):
        super().__init__(parent)
        self.voice = voice
        self.setWindowTitle(voice.host("teams.title"))
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(voice.host("teams.count")))
        self.count = QSpinBox()
        self.count.setRange(1, MAX_TEAMS)
        self.count.setValue(min(len(names), MAX_TEAMS) if names else 2)
        layout.addWidget(self.count)
        layout.addSpacing(6)
        layout.addWidget(QLabel(voice.host("teams.names")))
        self.form = QFormLayout()
        self.name_edits: list[QLineEdit] = []
        for i in range(MAX_TEAMS):
            edit = QLineEdit(names[i] if i < len(names) else "")
            edit.setPlaceholderText(voice.host("teams.default", n=i + 1))
            self.form.addRow(f"{i + 1}.", edit)
            self.name_edits.append(edit)
        layout.addLayout(self.form)
        # Start Game is our own button: Qt re-labels its standard ones (OK) whenever the
        # language changes, which would undo a setText. Cancel stays standard — Qt translates it.
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel)
        self.start_button = self.buttons.addButton(voice.host("teams.start"), QDialogButtonBox.ButtonRole.AcceptRole)
        self.start_button.setDefault(True)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        title = self.fontMetrics().horizontalAdvance(self.windowTitle())
        self.setMinimumWidth(max(int(title * 1.15) + self.TITLE_BAR_ALLOWANCE, self.minimumSizeHint().width()))
        self.count.valueChanged.connect(self._show_rows)
        self._show_rows(self.count.value())
        self.count.setFocus()
        self.count.selectAll()

    def _show_rows(self, count: int) -> None:
        for i in range(MAX_TEAMS):
            self.form.setRowVisible(i, i < count)
        self.adjustSize()

    def names(self) -> list[str]:
        return [edit.text().strip() or self.voice.host("teams.default", n=i + 1)
                for i, edit in enumerate(self.name_edits[: self.count.value()])]


class StartPage(QWidget):
    def __init__(self, window: QuizboardWindow):
        super().__init__()
        self.setObjectName("page")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        layout = QVBoxLayout(self)
        layout.addStretch(2)
        title = QLabel("Quizboard")
        title_font = QFont(title.font())
        title_font.setPixelSize(56)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setStyleSheet(f"color: {GOLD};")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        self.intro = QLabel()
        self.intro.setObjectName("hint")
        self.intro.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.intro)
        layout.addSpacing(24)
        row = QHBoxLayout()
        row.addStretch(1)
        self.buttons = []
        for key, slot in (("start.open", window.choose_file), ("start.template", window.save_template)):
            button = QPushButton()
            button.setObjectName("action")
            button.clicked.connect(slot)
            row.addWidget(button)
            self.buttons.append((button, key))
        row.addStretch(1)
        layout.addLayout(row)
        layout.addStretch(3)

    def retranslate(self, voice: Voice) -> None:
        self.intro.setText(voice.host("start.intro"))
        for button, key in self.buttons:
            button.setText(voice.host(key))


class QuizboardWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quizboard")
        self.setStyleSheet(STYLE)
        self.setAcceptDrops(True)
        self.resize(1280, 800)
        self.setMinimumSize(640, 420)

        self.config = load_config()
        language = self.config.get("language", "en")
        self.voice = Voice(language if language in LANGUAGES else "en")
        self.qt_translator: QTranslator | None = None
        self.board: Board | None = None
        # Bilingual mode: the other-language version of the board, matched tile by tile.
        # The board that was opened drives the game; the partner only supplies text.
        self.bilingual = bool(self.config.get("bilingual", False))
        self.partner: Board | None = None
        self.partner_warnings: list[Message] = []
        # Played tiles and awards are keyed by (topic casefolded, points), so they survive Reload.
        self.played: set[tuple[str, int]] = set()
        self.awards: dict[tuple[str, int], set[int]] = {}   # tile -> indices of the teams given its points
        self.teams: list[str] = []
        self.current: tuple[int, int] | None = None
        # The question just left — (column, points, was played before, answer shown) — so a
        # stray Esc can be undone. Cleared as soon as anything else changes the board.
        self.undo_state: tuple[int, int, bool, bool] | None = None

        self.start_page = StartPage(self)
        self.board_page = BoardPage()
        self.clue_page = CluePage(self.voice)
        self.score_page = ScorePage(self.voice)
        self.stack = QStackedWidget()
        for page in (self.start_page, self.board_page, self.clue_page, self.score_page):
            self.stack.addWidget(page)
        self.setCentralWidget(self.stack)

        # Undo is on the board (below the tiles) and on the scores page, while there is a question to go back to.
        self.undo_buttons = (self.board_page.undo_button, self.score_page.undo_button)
        for button in self.undo_buttons:
            button.clicked.connect(self.undo)
        self.stack.currentChanged.connect(self._update_undo_button)

        self.board_page.tile_clicked.connect(self._tile_clicked)
        self.board_page.tile_menu.connect(self._tile_menu)
        self.clue_page.advance.connect(self._advance)
        self.clue_page.back.connect(self._back_to_board)
        self.clue_page.award_toggled.connect(self._award_toggled)
        self.score_page.back.connect(self._leave_scores)
        self.score_page.new_game.connect(self.new_game)

        self.status_label = QLabel()
        self.statusBar().addWidget(self.status_label, 1)
        self._build_menus()
        self._install_qt_translator()
        self._retranslate()

        for key in ("Space", "Return", "Enter"):
            QShortcut(QKeySequence(key), self, activated=self._key_advance)
        for key in ("Esc", "Backspace"):
            QShortcut(QKeySequence(key), self, activated=self._key_back)
        for n in range(1, MAX_TEAMS + 1):
            QShortcut(QKeySequence(str(n)), self, activated=lambda i=n - 1: self.clue_page.toggle_team(i))

    # ---- menus and language ----------------------------------------------------

    def _menu(self, key: str) -> QMenu:
        menu = self.menuBar().addMenu("")
        self._texts.append((menu.setTitle, key))
        return menu

    def _action(self, menu: QMenu, key: str, slot, shortcut=None) -> QAction:
        action = QAction(self)
        if shortcut is not None:
            action.setShortcut(QKeySequence(shortcut))
        action.triggered.connect(slot)
        menu.addAction(action)
        self.addAction(action)   # shortcuts keep working with the menu bar hidden (full screen)
        self._texts.append((action.setText, key))
        return action

    def _build_menus(self) -> None:
        self._texts: list[tuple[object, str]] = []   # (setter, i18n key), re-run on a language change
        file_menu = self._menu("menu.file")
        self._action(file_menu, "menu.open", self.choose_file, QKeySequence.StandardKey.Open)
        self.reload_action = self._action(file_menu, "menu.reload", self.reload, "F5")
        self.edit_action = self._action(file_menu, "menu.edit", self.edit_in_excel, "Ctrl+E")
        file_menu.addSeparator()
        self._action(file_menu, "menu.template", self.save_template)
        file_menu.addSeparator()
        self._action(file_menu, "menu.quit", self.close, QKeySequence.StandardKey.Quit)

        game_menu = self._menu("menu.game")
        self.undo_action = self._action(game_menu, "menu.undo", self.undo, QKeySequence.StandardKey.Undo)
        self.undo_action.setEnabled(False)
        game_menu.addSeparator()
        self.new_game_action = self._action(game_menu, "menu.new_game", self.new_game, "Ctrl+N")
        self.scores_action = self._action(game_menu, "menu.scores", lambda _checked=False: self.show_scores(), "S")

        view_menu = self._menu("menu.view")
        self.full_screen_action = self._action(view_menu, "menu.full_screen", self.toggle_full_screen, "F11")
        self.full_screen_action.setCheckable(True)
        view_menu.addSeparator()
        self.bilingual_action = self._action(view_menu, "menu.bilingual", self.toggle_bilingual, "Ctrl+B")
        self.bilingual_action.setCheckable(True)
        self.bilingual_action.setChecked(self.bilingual)
        self.partner_action = self._action(view_menu, "menu.partner", self.choose_partner)
        language_menu = view_menu.addMenu("")
        self._texts.append((language_menu.setTitle, "menu.language"))
        group = QActionGroup(self)
        self.language_actions = {}
        for language in LANGUAGES:
            action = QAction(LANGUAGE_NAMES[language], self, checkable=True)
            action.setChecked(language == self.voice.ui)
            action.triggered.connect(lambda _checked=False, lang=language: self.set_language(lang))
            group.addAction(action)
            language_menu.addAction(action)
            self.language_actions[language] = action
        view_menu.addSeparator()
        self.report_action = self._action(view_menu, "menu.report", self.show_report)
        self._update_actions()

    def _update_actions(self) -> None:
        loaded = self.board is not None
        for action in (self.reload_action, self.edit_action, self.new_game_action, self.scores_action,
                       self.partner_action, self.report_action):
            action.setEnabled(loaded)

    def set_language(self, language: str) -> None:
        """Switch the interface language — menus, dialogs, status bar, and the pages' own text."""
        self.language_actions[language].setChecked(True)
        if language == self.voice.ui:
            return
        self.voice.ui = language
        self.config["language"] = language
        save_config(self.config)
        self._install_qt_translator()
        self._retranslate()

    def _install_qt_translator(self) -> None:
        # Qt's own text — Yes/No/Cancel, the file dialogs — comes from Qt's translation files.
        app = QApplication.instance()
        if self.qt_translator is not None:
            app.removeTranslator(self.qt_translator)
            self.qt_translator = None
        if self.voice.ui == "fr":
            translator = QTranslator(self)
            path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
            if translator.load(QLocale(QLocale.Language.French), "qtbase", "_", path):
                app.installTranslator(translator)
                self.qt_translator = translator

    def _retranslate(self) -> None:
        for setter, key in self._texts:
            setter(self.voice.host(key))
        self.start_page.retranslate(self.voice)
        self._apply_bilingual()
        if self.stack.currentWidget() is self.score_page:
            self.show_scores(final=self.score_page.final)
        self._update_undo_button()

    def _confirm(self, text: str) -> bool:
        return QMessageBox.question(self, "Quizboard", text) == QMessageBox.StandardButton.Yes

    # ---- files -----------------------------------------------------------------

    def _start_dir(self) -> str:
        last = self.config.get("last_file")
        if last and Path(last).parent.is_dir():
            return str(Path(last).parent)
        return str(Path.home())

    def _may_discard_game(self) -> bool:
        return not self._game_in_progress() or self._confirm(self.voice.host("msg.discard"))

    def choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, self.voice.host("dialog.open"), self._start_dir(), self.voice.host("filter.open")
        )
        if path and self._may_discard_game():
            self.open_path(path)

    def save_template(self) -> None:
        voice = self.voice
        path, _ = QFileDialog.getSaveFileName(
            self, voice.host("dialog.save_template"), str(Path(self._start_dir()) / voice.host("template.file")),
            voice.host("filter.save"),
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            write_template(path, voice.ui)
        except OSError as exc:
            QMessageBox.critical(self, "Quizboard", voice.host("msg.template_failed", error=exc))
            return
        QMessageBox.information(self, "Quizboard", voice.host("msg.template_saved", name=Path(path).name))

    def _read_board(self, path: str) -> Board | None:
        name = Path(path).name
        try:
            return load_board(path)
        except BoardError as exc:
            QMessageBox.critical(self, "Quizboard",
                                 self.voice.host("msg.cannot_use", name=name, error=render(exc.args[0], self.voice.ui)))
        except Exception as exc:   # a malformed workbook — say so rather than crash the board
            QMessageBox.critical(self, "Quizboard", self.voice.host("msg.cannot_read", name=name, error=exc))
        return None

    def open_path(self, path: str, keep_game: bool = False) -> bool:
        """Load a board. A new board starts a new game (teams asked); Reload keeps the game."""
        board = self._read_board(path)
        if board is None:
            return False
        if not keep_game:
            self.played.clear()
            self.awards.clear()
        partner = self.partner.path if keep_game and self.partner is not None else None
        self.undo_state = None
        self._update_undo_button()
        self.board = board
        self.partner, self.partner_warnings = None, []
        self.current = None
        self.config["last_file"] = str(Path(path).resolve())
        save_config(self.config)
        self.setWindowTitle(f"Quizboard — {Path(path).name}")
        if self.bilingual:
            self._attach_partner(partner or partner_path(path))
        self._apply_bilingual()
        self.stack.setCurrentWidget(self.board_page)
        self._update_actions()
        if self._all_warnings():
            self.show_report()
        if not keep_game:
            self._start_game(self.ask_teams() or [])
        return True

    def reload(self) -> None:
        if self.board is not None:
            self.open_path(self.board.path, keep_game=True)

    def edit_in_excel(self) -> None:
        if self.board is None:
            return
        path = str(Path(self.board.path).resolve())
        # Under WSL, hand the file to Windows so it opens in Excel; elsewhere use the desktop default.
        if shutil.which("wslpath") and shutil.which("explorer.exe"):
            try:
                windows_path = subprocess.run(
                    ["wslpath", "-w", path], capture_output=True, text=True, check=True
                ).stdout.strip()
                subprocess.Popen(["explorer.exe", windows_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except (OSError, subprocess.CalledProcessError) as exc:
                QMessageBox.warning(self, "Quizboard", self.voice.host("msg.windows_failed", error=exc))
                return
        elif not QDesktopServices.openUrl(QUrl.fromLocalFile(path)):
            QMessageBox.warning(self, "Quizboard", self.voice.host("msg.no_app", path=path))
            return
        self.statusBar().showMessage(self.voice.host("status.excel"), 8000)

    def _all_warnings(self) -> list[str]:
        return (self.board.warnings if self.board is not None else []) + self.partner_warnings

    def show_report(self) -> None:
        if self.board is None:
            return
        voice = self.voice
        name = Path(self.board.path).name
        if self.partner is not None:
            name += f" + {Path(self.partner.path).name}"
        warnings = [render(w, voice.ui) for w in self._all_warnings()]
        counts = {"name": name, "clues": len(self.board.clues), "tiles": COLUMNS * ROWS}
        if not warnings:
            QMessageBox.information(self, "Quizboard", voice.host("report.clean", **counts))
            return
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle(voice.host("report.title"))
        box.setText(voice.host("report.problems", problems=voice.host_n("problems", len(warnings)), **counts))
        box.setInformativeText("\n".join(f"• {w}" for w in warnings[:12]) + ("\n…" if len(warnings) > 12 else ""))
        if len(warnings) > 12:
            box.setDetailedText("\n".join(warnings))
        box.exec()

    def dragEnterEvent(self, event) -> None:   # noqa: N802 — Qt override
        if self._dropped_path(event) is not None:
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:   # noqa: N802 — Qt override
        path = self._dropped_path(event)
        if path is not None:
            event.acceptProposedAction()
            if self._may_discard_game():
                self.open_path(path)

    @staticmethod
    def _dropped_path(event) -> str | None:
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        for url in urls:
            if url.isLocalFile() and url.toLocalFile().lower().endswith(XLSX_SUFFIXES):
                return url.toLocalFile()
        return None

    # ---- bilingual mode --------------------------------------------------------

    def toggle_bilingual(self, on: bool) -> None:
        if on and self.board is not None and self.partner is None:
            path = partner_path(self.board.path) or self._ask_partner_path()
            if path is None or not self._attach_partner(path):
                self.bilingual_action.setChecked(False)
                return
            if self.partner_warnings:
                self.show_report()
        self._set_bilingual(on)

    def choose_partner(self) -> None:
        path = self._ask_partner_path()
        if path is not None and self._attach_partner(path):
            if self.partner_warnings:
                self.show_report()
            self._set_bilingual(True)

    def _ask_partner_path(self) -> str | None:
        name = Path(self.board.path).name
        path, _ = QFileDialog.getOpenFileName(
            self, self.voice.host("dialog.partner", name=name), str(Path(self.board.path).parent),
            self.voice.host("filter.open"),
        )
        return path or None

    def _attach_partner(self, path: str | None) -> bool:
        """Load the other-language board. With no file to load, bilingual mode is switched off
        and the reason is shown — the board is never quietly left in one language."""
        name = Path(self.board.path).name
        if path is None:
            self._set_bilingual(False)
            QMessageBox.information(self, "Quizboard", self.voice.host("msg.no_partner", name=name))
            return False
        if Path(path).resolve() == Path(self.board.path).resolve():
            QMessageBox.warning(self, "Quizboard", self.voice.host("msg.same_file", name=name))
            return False
        partner = self._read_board(path)
        if partner is None:
            self._set_bilingual(False)
            return False
        self.partner = partner
        self.partner_warnings = [Message("partner.warning", name=Path(path).name, warning=w)
                                 for w in partner.warnings if w not in partner.gaps]
        self.partner_warnings += compare_boards(self.board, partner)
        return True

    def _set_bilingual(self, on: bool) -> None:
        self.bilingual = on
        self.bilingual_action.setChecked(on)
        self.config["bilingual"] = on
        save_config(self.config)
        self._apply_bilingual()

    def _bilingual_on(self) -> bool:
        return self.bilingual and self.partner is not None

    def _apply_bilingual(self) -> None:
        # What the room sees: both languages, left column's first — or the board's own language.
        if self._bilingual_on():
            left = language_of(self.board.path) or self.voice.ui
            right = language_of(self.partner.path)
            self.voice.room = [left, right if right not in (None, left) else other(left)]
        else:
            board_language = language_of(self.board.path) if self.board is not None else None
            self.voice.room = [board_language or self.voice.ui]
        self.clue_page.set_bilingual(self._bilingual_on())
        self._refresh_board()
        if self.stack.currentWidget() is self.clue_page and self.current is not None:
            revealed = self.clue_page.revealed
            self._open_clue(*self.current)
            if revealed:
                self.clue_page.reveal()

    # ---- teams and scores ------------------------------------------------------

    def ask_teams(self) -> list[str] | None:
        """The teams for a new game, or None when the dialog is cancelled."""
        dialog = TeamsDialog(self, self.teams or self.config.get("teams", []), self.voice)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        names = dialog.names()
        self.config["teams"] = names
        save_config(self.config)
        return names

    def new_game(self) -> None:
        if self.board is None:
            return
        if self._game_in_progress() and not self._confirm(self.voice.host("msg.new_game")):
            return
        names = self.ask_teams()
        if names is not None:
            self._start_game(names)

    def _start_game(self, teams: list[str]) -> None:
        self.teams = teams
        self.played.clear()
        self.awards.clear()
        self.current = None
        self.undo_state = None
        self._update_undo_button()
        self.clue_page.set_teams(teams)
        self._refresh_board()
        self.stack.setCurrentWidget(self.board_page)

    def _game_in_progress(self) -> bool:
        return bool(self.played or self.awards)

    def scores(self) -> list[int]:
        totals = [0] * len(self.teams)
        for (_, points), winners in self.awards.items():
            for team in winners:
                if team < len(totals):
                    totals[team] += points
        return totals

    def _award_toggled(self, team: int, ticked: bool) -> None:
        if self.current is None:
            return
        key = self._key(*self.current)
        winners = self.awards.setdefault(key, set())
        if ticked:
            winners.add(team)
        else:
            winners.discard(team)
        if not winners:
            del self.awards[key]
        self._update_team_labels()
        self._refresh_board()

    def _update_team_labels(self) -> None:
        self.clue_page.set_team_labels([f"{name}  ({score})" for name, score in zip(self.teams, self.scores())])

    def show_scores(self, final: bool | None = None) -> None:
        if self.board is None:
            return
        if final is None:
            final = self._complete()
        played = sum(1 for (c, p) in self.board.clues if self._key(c, p) in self.played)
        self.score_page.show_scores(self.teams, self.scores(), final, played, len(self.board.clues))
        self.stack.setCurrentWidget(self.score_page)

    def _leave_scores(self) -> None:
        # Scores opened mid-question go back to that question; otherwise to the board.
        self.stack.setCurrentWidget(self.clue_page if self.current is not None else self.board_page)

    # ---- game ------------------------------------------------------------------

    def _key(self, column: int, points: int) -> tuple[str, int]:
        return (self.board.topics[column].casefold(), points)

    def _state(self, column: int, points: int) -> str:
        if self.board is None or self.board.clue(column, points) is None:
            return "missing"
        return "played" if self._key(column, points) in self.played else "ready"

    def _complete(self) -> bool:
        return self.board is not None and bool(self.board.clues) and all(
            self._key(c, p) in self.played for (c, p) in self.board.clues
        )

    def _refresh_board(self) -> None:
        if self.board is None:
            return
        voice = self.voice
        live = {self._key(c, p) for (c, p) in self.board.clues}
        self.played &= live   # a reload may have removed or renamed tiles
        self.awards = {key: winners for key, winners in self.awards.items() if key in live}
        self.board_page.show_board(self.board, self._state, voice, self.partner if self._bilingual_on() else None)
        name = Path(self.board.path).name
        if self._bilingual_on():
            name = voice.host("status.pair", name=name, partner=Path(self.partner.path).name)
        parts = [name, voice.host("status.played", played=len(self.played), total=len(self.board.clues))]
        if self.teams:
            parts.append("   ".join(f"{team}: {score}" for team, score in zip(self.teams, self.scores())))
        else:
            parts.append(voice.host("status.no_teams"))
        warnings = self._all_warnings()
        if warnings:
            parts.append(voice.host_n("status.warnings", len(warnings)))
        self.status_label.setText("   ·   ".join(parts))

    def _tile_clicked(self, column: int, points: int) -> None:
        if self._state(column, points) == "ready":
            self._open_clue(column, points)

    def _tile_menu(self, column: int, points: int, position) -> None:
        state = self._state(column, points)
        if state == "missing":
            return
        menu = QMenu(self)
        show = menu.addAction(self.voice.host("tile.show"))
        toggle = menu.addAction(self.voice.host("tile.mark_unplayed" if state == "played" else "tile.mark_played"))
        chosen = menu.exec(position)
        if chosen is show:
            self._open_clue(column, points)
        elif chosen is toggle:
            was_complete = self._complete()
            self.played ^= {self._key(column, points)}
            self.undo_state = None
            self._refresh_board()
            self._board_or_final_scores(was_complete)
            self._update_undo_button()

    def _open_clue(self, column: int, points: int) -> None:
        self.current = (column, points)
        self.undo_state = None   # another question: the last one can no longer be undone
        clue = self.board.clue(column, points)
        sides = [(clue.topic, clue.question, clue.answer, "")]
        if self._bilingual_on():
            other_clue = self.partner.clue(column, points)
            if other_clue is not None:
                sides.append((other_clue.topic, other_clue.question, other_clue.answer, ""))
            else:
                topic = self.partner.topics[column] if column < len(self.partner.topics) else ""
                sides.append((topic, "", "", Path(self.partner.path).name))
        self.clue_page.show_clue(sides, points, self.awards.get(self._key(column, points), set()))
        self._update_team_labels()
        self.stack.setCurrentWidget(self.clue_page)

    def _advance(self) -> None:
        if self.stack.currentWidget() is not self.clue_page:
            return
        if self.clue_page.revealed:
            self._back_to_board()
        else:
            self.clue_page.reveal()

    def _back_to_board(self) -> None:
        was_complete = self._complete()
        if self.current is not None:
            key = self._key(*self.current)
            self.undo_state = (*self.current, key in self.played, self.clue_page.revealed)
            self.played.add(key)
            self.current = None
        self._refresh_board()
        self._board_or_final_scores(was_complete)
        self._update_undo_button()

    def undo(self) -> None:
        """Back to the question just left (an Esc pressed by mistake), as it was: answer hidden
        or shown, team ticks untouched. The tile is unplayed again unless it already was played."""
        if self.undo_state is None or self.board is None:
            return
        column, points, was_played, revealed = self.undo_state
        if not was_played:
            self.played.discard(self._key(column, points))
        self._open_clue(column, points)
        if revealed:
            self.clue_page.reveal()
        self._refresh_board()
        self._update_undo_button()

    def _update_undo_button(self) -> None:
        available = self.undo_state is not None
        self.undo_action.setEnabled(available)
        for button in self.undo_buttons:
            button.setText(f"↶  {self.voice.show('undo.button')}  (Ctrl+Z)")
            button.setToolTip(self.voice.host("undo.tooltip"))
            button.setVisible(available)

    def _board_or_final_scores(self, was_complete: bool) -> None:
        # The final scores appear once, when the last tile is played.
        if self.teams and not was_complete and self._complete():
            self.show_scores(final=True)
        else:
            self.stack.setCurrentWidget(self.board_page)

    def _key_advance(self) -> None:
        if self.stack.currentWidget() is self.clue_page:
            self._advance()
        elif self.stack.currentWidget() is self.score_page:
            self._leave_scores()

    def _key_back(self) -> None:
        if self.stack.currentWidget() is self.clue_page:
            self._back_to_board()
        elif self.stack.currentWidget() is self.score_page:
            self._leave_scores()
        elif self.isFullScreen():
            self.toggle_full_screen()

    def toggle_full_screen(self) -> None:
        full = not self.isFullScreen()
        self.menuBar().setVisible(not full)
        self.statusBar().setVisible(not full)
        if full:
            self.showFullScreen()
        else:
            self.showNormal()
        self.full_screen_action.setChecked(full)


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    app = QApplication(sys.argv[:1])
    app.setApplicationName("Quizboard")
    app.setDesktopFileName("quizboard")   # GNOME matches the window to quizboard.desktop
    app.setWindowIcon(QIcon(str(Path(__file__).resolve().parent / "icons" / "quizboard.png")))
    window = QuizboardWindow()
    window.show()
    if argv:
        QTimer.singleShot(0, lambda: window.open_path(argv[0]))
    else:
        last = window.config.get("last_file")
        if last and Path(last).is_file():
            QTimer.singleShot(0, lambda: window.open_path(last))
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
