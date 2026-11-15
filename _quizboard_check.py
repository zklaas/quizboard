#!/usr/bin/env python3
"""Quizboard regression checks — reader, writer, board rules and a headless GUI run.

    python3 _quizboard_check.py      (or: ./run.sh --check)

Exits non-zero if any check fails.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import traceback
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent))

from board_io import (   # noqa: E402
    BOARD_HEADER,
    BoardError,
    SheetSpec,
    compare_boards,
    load_board,
    partner_path,
    read_workbook,
    write_template,
    write_workbook,
)

TMP = Path(tempfile.mkdtemp(prefix="quizboard_check_"))
CHECKS = []
_APP = None


def check(fn):
    CHECKS.append(fn)
    return fn


def board_from(rows, name="Board", extra=()):
    path = TMP / f"{len(list(TMP.iterdir()))}.xlsx"
    write_workbook(str(path), [SheetSpec(name, rows)] + list(extra))
    return load_board(str(path))


def full_topic(topic, start=1):
    return [[topic, p, f"{topic} q{p}", f"{topic} a{p}"] for p in range(start, 7)]


def demo_board() -> str:
    """A complete 6 × 6 board: 'Topic A'…'Topic F', 'Question A1' / 'Answer A1' and so on."""
    path = TMP / "demo.xlsx"
    if not path.exists():
        rows = [BOARD_HEADER] + [[f"Topic {t}", p, f"Question {t}{p}", f"Answer {t}{p}"]
                                 for t in "ABCDEF" for p in range(1, 7)]
        write_workbook(str(path), [SheetSpec("Board", rows)])
    return str(path)


def has_warning(board, *fragments):
    return any(all(f in w for f in fragments) for w in board.warnings)


@check
def demo_board_is_complete():
    board = load_board(demo_board())
    assert len(board.topics) == 6 and len(board.clues) == 36 and board.warnings == [], board.warnings
    assert board.clue(0, 1).answer == "Answer A1"


@check
def every_text_has_both_languages_with_the_same_placeholders():
    from string import Formatter

    from i18n import HOW_TO, TEMPLATE_HEADER, TEXT

    def fields(text):
        return sorted(name for _, name, _, _ in Formatter().parse(text) if name)

    for key, (en, fr) in TEXT.items():
        assert en and fr, key
        assert fields(en) == fields(fr), (key, fields(en), fields(fr))
    assert len(HOW_TO["en"]) == len(HOW_TO["fr"]) and len(TEMPLATE_HEADER["en"]) == len(TEMPLATE_HEADER["fr"])


@check
def warnings_and_errors_render_in_french():
    rows = [BOARD_HEADER, ["", 1, "orphan", "x"]] + [r for t in "ABCDEFG" for r in full_topic(t)]
    rows += [["A", 7, "too many points", "x"], ["B", 2, "duplicate", "x"]]
    board = board_from(rows)
    french = [w.render("fr") for w in board.warnings]
    assert any("ligne 2" in w and "sans catégorie" in w for w in french), french
    assert any("«\u00a0G\u00a0»" in w and "n'ont pas été utilisées" in w for w in french), french
    assert any("«\u00a07\u00a0»" in w and "de 1 à 6" in w for w in french), french
    assert any("a déjà une question pour la valeur 2" in w for w in french), french
    assert all(str(w) == w.render("en") for w in board.warnings)   # a Message reads as its English
    try:
        load_board(str(TMP / "absent.xlsx"))
    except BoardError as exc:
        assert exc.args[0].render("fr").startswith("Fichier introuvable"), exc.args[0].render("fr")


@check
def french_template_has_french_headers_tabs_and_instructions():
    path = TMP / "modele.xlsx"
    write_template(str(path), "fr")
    sheets = read_workbook(str(path))
    assert [name for name, _ in sheets] == ["Tableau", "Mode d'emploi"]
    assert sheets[0][1][0] == ["Catégorie", "Points", "Question", "Réponse"]
    assert sheets[1][1][0] == ["Comment remplir la feuille Tableau"]
    board = load_board(str(path))
    assert board.topics == [f"Catégorie {i}" for i in range(1, 7)] and board.clues == {}


@check
def template_has_six_empty_topics():
    path = TMP / "template.xlsx"
    write_template(str(path))
    board = load_board(str(path))
    assert board.topics == [f"Topic {i}" for i in range(1, 7)], board.topics
    assert board.clues == {}
    assert len(board.warnings) == 6 and has_warning(board, "'Topic 1'", "1, 2, 3, 4, 5, 6")


@check
def writer_round_trips_awkward_text():
    texts = ['<b>&"quotes"</b>', "two\nlines", "  leading space", "Réponse — été ✓", "₂O"]
    path = TMP / "awkward.xlsx"
    write_workbook(str(path), [SheetSpec("Odd 'name' & co", [["Text", "Number"]] + [[t, 3] for t in texts])])
    (name, rows), = read_workbook(str(path))
    assert name == "Odd 'name' & co", name
    assert [r[0] for r in rows[1:]] == texts, rows
    assert all(r[1] == "3" for r in rows[1:])


@check
def topic_written_once_carries_down():
    rows = [BOARD_HEADER]
    for t in "ABCDEF":
        block = full_topic(t)
        for row in block[1:]:
            row[0] = ""
        rows += block
    board = board_from(rows)
    assert board.topics == list("ABCDEF") and len(board.clues) == 36 and not board.warnings, board.warnings
    assert board.clue(5, 6).question == "F q6"


@check
def topic_only_rows_start_a_block_and_blank_points_fill_in_order():
    rows = [BOARD_HEADER]
    for t in "ABCDEF":
        rows.append([t, "", "", ""])
        rows += [["", "", f"{t} q{p}", f"{t} a{p}"] for p in range(1, 7)]
    board = board_from(rows)
    assert len(board.clues) == 36 and not board.warnings, board.warnings
    assert board.clue(2, 4).question == "C q4"


@check
def explicit_points_are_placed_by_value_not_order():
    rows = [BOARD_HEADER] + [["A", p, f"q{p}", f"a{p}"] for p in (6, 2, 4, 1, 5, 3)]
    board = board_from(rows)
    assert all(board.clue(0, p).question == f"q{p}" for p in range(1, 7))


@check
def sheet_per_topic_with_french_headers():
    sheets = [SheetSpec(f"Thème {i}", [["Valeur", "Question", "Réponse"]] + [[p, f"q{i}{p}", f"r{i}{p}"] for p in range(1, 7)])
              for i in range(1, 7)]
    path = TMP / "per_sheet.xlsx"
    write_workbook(str(path), sheets)
    board = load_board(str(path))
    assert board.topics == [f"Thème {i}" for i in range(1, 7)], board.topics
    assert len(board.clues) == 36 and not board.warnings, board.warnings


@check
def unread_sheets_are_ignored():
    rows = [BOARD_HEADER] + [r for t in "ABCDEF" for r in full_topic(t)]
    board = board_from(rows, extra=[SheetSpec("Notes", [["Question and Answer go on the Board sheet"]])])
    assert len(board.clues) == 36 and not board.warnings, board.warnings


@check
def problems_are_reported_not_dropped_silently():
    rows = [BOARD_HEADER, ["", 1, "orphan question", "x"]]
    for t in "ABCDEFG":
        rows += full_topic(t)
    rows += [["A", 7, "too many points", "x"], ["B", 2, "duplicate", "x"], ["C", "two", "word points", "x"]]
    rows[2 + 6 * 3 + 2][3] = ""   # D, 3 points: no answer
    board = board_from(rows)
    assert board.topics == list("ABCDEF"), board.topics
    assert has_warning(board, "row 2", "no topic above")
    assert has_warning(board, "'G'", "not used")
    assert has_warning(board, "'7'", "1 to 6")
    assert has_warning(board, "'B' already has a 2-point")
    assert has_warning(board, "'two'")
    assert has_warning(board, "'D' 3 points has a question but no answer")
    assert board.clue(3, 3).question == "D q3"   # kept, just flagged


@check
def missing_cells_and_short_boards_are_reported():
    rows = [BOARD_HEADER] + full_topic("A") + full_topic("B", start=3)
    board = board_from(rows)
    assert len(board.topics) == 2 and len(board.clues) == 10
    assert has_warning(board, "Only 2 topic")
    assert has_warning(board, "'B' has no question for points 1, 2")


@check
def partner_files_are_found_by_language_suffix():
    folder = TMP / "pairs"
    folder.mkdir()
    for name in ("quiz_en.xlsx", "quiz_fr.xlsx", "other_FR.xlsx", "other_EN.xlsx", "lone_en.xlsx", "plain.xlsx"):
        (folder / name).write_bytes(b"")
    assert partner_path(str(folder / "quiz_en.xlsx")) == str(folder / "quiz_fr.xlsx")
    assert partner_path(str(folder / "quiz_fr.xlsx")) == str(folder / "quiz_en.xlsx")
    assert partner_path(str(folder / "other_FR.xlsx")) == str(folder / "other_EN.xlsx")
    assert partner_path(str(folder / "lone_en.xlsx")) is None     # its _fr doesn't exist
    assert partner_path(str(folder / "plain.xlsx")) is None       # no language suffix


@check
def compare_boards_lists_tiles_that_dont_line_up():
    en = board_from([BOARD_HEADER] + [r for t in "ABCDEF" for r in full_topic(t) if r[:2] != ["A", 3]])
    fr = board_from([BOARD_HEADER] + [r for t in "UVWXY" for r in full_topic(t)] + [["Z", 1, "q", "a"]])
    warnings = compare_boards(en, fr)
    fr_name, en_name = Path(fr.path).name, Path(en.path).name
    assert len(warnings) == 2, warnings
    assert warnings[0] == f"Column 1 ('A'): {fr_name} has a question for points 3 that {en_name} lacks — not on the board."
    assert warnings[1] == f"Column 6 ('F'): {fr_name} has no question for points 2, 3, 4, 5, 6."


@check
def unreadable_files_raise_board_error():
    no_header = TMP / "no_header.xlsx"
    write_workbook(str(no_header), [SheetSpec("Board", [["Topic", "Points", "Clue text", "Reply"], ["A", 1, "q", "a"]])])
    not_xlsx = TMP / "old.xls"
    not_xlsx.write_bytes(b"\xd0\xcf\x11\xe0 not a zip")
    for path, fragment in ((no_header, "Question"), (not_xlsx, "not an .xlsx"), (TMP / "absent.xlsx", "not found")):
        try:
            load_board(str(path))
        except BoardError as exc:
            assert fragment in str(exc), (path, exc)
        else:
            raise AssertionError(f"{path.name} should not load")


def _gui_window(teams):
    """A real window, headless, whose New Game dialog answers `teams` (None = cancelled)."""
    from PyQt6.QtWidgets import QApplication

    import quizboard

    global _APP
    quizboard.CONFIG_PATH = TMP / "config.json"   # never touch the real remembered file
    quizboard.CONFIG_PATH.unlink(missing_ok=True)  # each window starts from defaults
    _APP = QApplication.instance() or QApplication([])   # held, or Python collects it
    win = quizboard.QuizboardWindow()
    win.ask_teams = lambda: teams
    win.show()
    _activate(win)
    return win


def _activate(win):
    """Keyboard shortcuts only fire in the active window — and closing any dialog deactivates it."""
    from PyQt6.QtTest import QTest
    win.activateWindow()
    assert QTest.qWaitForWindowActive(win), "window never became active (keyboard shortcuts need it)"


def _expect_dialog(texts, tries=60):
    """Close the next modal dialog once it appears, keeping its text."""
    from PyQt6.QtCore import QTimer
    from PyQt6.QtWidgets import QApplication

    def close(left=tries):
        box = QApplication.activeModalWidget()
        if box is None and left > 0:
            QTimer.singleShot(50, lambda: close(left - 1))
            return
        texts.append(box.text() if box is not None else None)
        if box is not None:
            box.done(0)

    QTimer.singleShot(50, close)


def _label_texts(widget):
    import quizboard
    return [label.text() for label in widget.findChildren(quizboard.FitLabel)]


@check
def gui_tile_question_answer_board_cycle():
    from PyQt6.QtCore import Qt
    from PyQt6.QtTest import QTest

    win = _gui_window(teams=None)                                # New Game dialog cancelled: no scoring
    assert win.stack.currentWidget() is win.start_page
    assert win.open_path(demo_board())
    page, clue = win.board_page, win.clue_page
    assert win.stack.currentWidget() is page and win.teams == []
    assert page.topics[0].text() == "Topic A"
    assert [page.tiles[(0, p)].text() for p in range(1, 7)] == ["1", "2", "3", "4", "5", "6"]

    tile = page.tiles[(0, 1)]
    QTest.mouseClick(tile, Qt.MouseButton.LeftButton)
    assert win.stack.currentWidget() is clue
    assert clue.question.text() == "Question A1" and clue.answer.text() == ""
    QTest.mouseClick(clue.question, Qt.MouseButton.LeftButton)   # click the question: answer appears
    assert clue.answer.text() == "Answer A1" and win.stack.currentWidget() is clue
    assert not clue.team_boxes and not clue.team_row.isVisible()
    QTest.mouseClick(clue.question, Qt.MouseButton.LeftButton)   # click again: back to the board
    assert win.stack.currentWidget() is page
    assert tile.state == "played" and tile.text() == ""
    QTest.mouseClick(tile, Qt.MouseButton.LeftButton)            # a played tile stays put
    assert win.stack.currentWidget() is page

    QTest.mouseClick(page.tiles[(5, 6)], Qt.MouseButton.LeftButton)
    QTest.keyClick(win, Qt.Key.Key_Escape)                       # Esc without the answer still marks it played
    assert win.stack.currentWidget() is page and page.tiles[(5, 6)].state == "played"
    QTest.mouseClick(page.tiles[(2, 3)], Qt.MouseButton.LeftButton)
    QTest.keyClick(win, Qt.Key.Key_Space)
    assert clue.answer.text() == "Answer C3"
    QTest.keyClick(win, Qt.Key.Key_Space)
    assert win.stack.currentWidget() is page and len(win.played) == 3

    win.reload()                                                 # Reload keeps the game's progress
    assert len(win.played) == 3 and page.tiles[(0, 1)].state == "played"

    partial = TMP / "partial.xlsx"                               # problems pop the load report
    write_workbook(str(partial), [SheetSpec("Board", [BOARD_HEADER] + full_topic("A", start=2))])
    reports = []
    _expect_dialog(reports)
    assert win.open_path(str(partial))
    assert reports and reports[0] and "2 problems" in reports[0] and "5 of 36" in reports[0], reports
    assert page.tiles[(0, 1)].state == "missing" and page.tiles[(1, 1)].state == "missing"
    assert "load warnings" in win.status_label.text()
    win.close()


@check
def gui_teams_tick_points_and_final_scores():
    from PyQt6.QtCore import Qt
    from PyQt6.QtTest import QTest

    win = _gui_window(teams=["Red", "Blue", "Green"])
    assert win.open_path(demo_board())                           # opening a board asks for the teams
    page, clue, scores = win.board_page, win.clue_page, win.score_page
    assert win.teams == ["Red", "Blue", "Green"] and len(clue.team_boxes) == 3

    QTest.mouseClick(page.tiles[(0, 1)], Qt.MouseButton.LeftButton)
    assert not clue.team_row.isVisible()                         # no ticking before the answer is shown
    QTest.keyClick(win, Qt.Key.Key_1)
    assert win.scores() == [0, 0, 0]
    QTest.keyClick(win, Qt.Key.Key_Space)
    assert clue.team_row.isVisible()
    red = clue.team_boxes[0]
    QTest.mouseClick(red, Qt.MouseButton.LeftButton, pos=red.rect().center())
    assert win.stack.currentWidget() is clue                     # ticking doesn't leave the question
    assert red.isChecked() and win.scores() == [1, 0, 0] and red.text() == "Red  (1)"
    QTest.keyClick(win, Qt.Key.Key_Space)
    assert win.stack.currentWidget() is page

    QTest.mouseClick(page.tiles[(2, 3)], Qt.MouseButton.LeftButton)
    QTest.keyClick(win, Qt.Key.Key_Space)
    QTest.keyClick(win, Qt.Key.Key_2)                            # number keys tick teams
    QTest.keyClick(win, Qt.Key.Key_3)
    QTest.keyClick(win, Qt.Key.Key_3)                            # ... and untick them
    assert win.scores() == [1, 3, 0]
    QTest.keyClick(win, Qt.Key.Key_Space)

    win._open_clue(0, 1)                                         # reopening shows, and can correct, the ticks
    QTest.keyClick(win, Qt.Key.Key_Space)
    assert clue.team_boxes[0].isChecked() and not clue.team_boxes[1].isChecked()
    QTest.keyClick(win, Qt.Key.Key_1)
    assert win.scores() == [0, 3, 0]
    QTest.keyClick(win, Qt.Key.Key_1)
    assert win.scores() == [1, 3, 0]
    QTest.keyClick(win, Qt.Key.Key_Escape)
    assert "Red: 1   Blue: 3   Green: 0" in win.status_label.text()

    QTest.keyClick(win, Qt.Key.Key_S)                            # Game ▸ Scores, mid-game
    assert win.stack.currentWidget() is scores and not scores.final
    assert scores.title.text() == "Scores" and scores.subtitle.text() == "2 of 36 questions played"
    assert _label_texts(scores.table)[:6] == ["1.", "Blue", "3", "2.", "Red", "1"], _label_texts(scores.table)
    QTest.keyClick(win, Qt.Key.Key_Escape)
    assert win.stack.currentWidget() is page

    QTest.mouseClick(page.tiles[(4, 4)], Qt.MouseButton.LeftButton)
    QTest.keyClick(win, Qt.Key.Key_S)                            # scores opened mid-question ...
    assert win.stack.currentWidget() is scores
    QTest.keyClick(win, Qt.Key.Key_Escape)                       # ... go back to that question
    assert win.stack.currentWidget() is clue and win.current == (4, 4)
    QTest.keyClick(win, Qt.Key.Key_Escape)

    win.reload()                                                 # Reload keeps teams and scores
    assert win.teams == ["Red", "Blue", "Green"] and win.scores() == [1, 3, 0]

    for (c, p), tile in sorted(page.tiles.items()):              # play out the board; Red takes column 0
        if tile.state != "ready":
            continue
        assert win.stack.currentWidget() is page, (c, p)
        QTest.mouseClick(tile, Qt.MouseButton.LeftButton)
        QTest.keyClick(win, Qt.Key.Key_Space)
        if c == 0:
            QTest.keyClick(win, Qt.Key.Key_1)
        QTest.keyClick(win, Qt.Key.Key_Space)
    assert win.stack.currentWidget() is scores and scores.final  # the last tile brings up the final scores
    assert scores.title.text() == "Final Scores"
    assert scores.subtitle.text() == "Red wins with 21 points!", scores.subtitle.text()
    QTest.keyClick(win, Qt.Key.Key_Escape)
    assert win.stack.currentWidget() is page

    win._confirm = lambda text: True
    win.ask_teams = lambda: None                                 # New Game cancelled at the teams question
    win.new_game()
    assert win.scores() == [21, 3, 0] and len(win.played) == 36
    win.ask_teams = lambda: ["Ayes", "Bees"]
    win.new_game()
    assert win.teams == ["Ayes", "Bees"] and win.scores() == [0, 0] and not win.played
    assert len(clue.team_boxes) == 2 and all(t.state == "ready" for t in page.tiles.values())
    win.close()


@check
def gui_score_page_ties_and_teams_dialog():
    import quizboard

    win = _gui_window(teams=None)
    page = win.score_page
    page.show_scores(["X", "Y", "Z"], [5, 2, 5], True, 36, 36)
    assert page.subtitle.text() == "A tie at 5 points: X and Z"
    assert _label_texts(page.table) == ["1.", "X", "5", "1.", "Z", "5", "3.", "Y", "2"], _label_texts(page.table)
    page.show_scores([], [], True, 36, 36)
    assert "New Game" in page.subtitle.text()

    dialog = quizboard.TeamsDialog(win, ["Red"], win.voice)
    assert dialog.count.value() == 1 and dialog.names() == ["Red"]
    dialog.count.setValue(3)
    assert dialog.names() == ["Red", "Team 2", "Team 3"]
    assert dialog.form.isRowVisible(2) and not dialog.form.isRowVisible(3)
    dialog.close()
    win.close()


@check
def gui_bilingual_mode_side_by_side():
    from PyQt6.QtCore import Qt
    from PyQt6.QtTest import QTest

    folder = TMP / "bilingual"
    folder.mkdir()
    en_rows = [BOARD_HEADER] + [[f"Topic {t}", p, f"EN question {t}{p}", f"EN answer {t}{p}"]
                                for t in "ABCDEF" for p in range(1, 7)]
    fr_rows = [["Catégorie", "Points", "Question", "Réponse"]] + [
        [f"Sujet {t}", p, f"Question FR {t}{p}", f"Réponse FR {t}{p}"]
        for t in "ABCDEF" for p in range(1, 7) if (t, p) != ("B", 4)]
    write_workbook(str(folder / "quiz_en.xlsx"), [SheetSpec("Board", en_rows)])
    write_workbook(str(folder / "quiz_fr.xlsx"), [SheetSpec("Tableau", fr_rows)])
    write_workbook(str(folder / "lonely.xlsx"), [SheetSpec("Board", en_rows)])

    win = _gui_window(teams=["Red", "Blue"])
    page, clue = win.board_page, win.clue_page
    left, right = clue.sides
    assert win.open_path(str(folder / "quiz_en.xlsx")) and win.partner is None   # mode starts off
    reports = []
    _expect_dialog(reports)
    win.bilingual_action.trigger()                               # finds quiz_fr.xlsx by its name
    _activate(win)
    assert win.bilingual and win.partner.path.endswith("quiz_fr.xlsx")
    assert reports and "1 problem" in reports[0], reports        # FR has no B-4: reported
    cell = page.cells[0]
    assert (cell.primary.text(), cell.secondary.text()) == ("Topic A", "Sujet A")
    assert cell.primary.font().pixelSize() == cell.secondary.font().pixelSize()   # equal prominence

    QTest.mouseClick(page.tiles[(0, 2)], Qt.MouseButton.LeftButton)
    assert right.isVisible() and left.question.text() == "EN question A2"
    assert right.question.text() == "Question FR A2" and right.heading.text().startswith("Sujet A")
    assert left.answer.text() == right.answer.text() == ""
    QTest.keyClick(win, Qt.Key.Key_Space)                        # both answers appear together
    assert (left.answer.text(), right.answer.text()) == ("EN answer A2", "Réponse FR A2")
    assert left.question.font().pixelSize() == right.question.font().pixelSize()
    QTest.keyClick(win, Qt.Key.Key_1)
    QTest.keyClick(win, Qt.Key.Key_Space)

    QTest.mouseClick(page.tiles[(1, 4)], Qt.MouseButton.LeftButton)   # no French version: said so, in French
    assert right.question.text() == "(Absent de quiz_fr.xlsx)" and right.question.objectName() == "missing"
    QTest.keyClick(win, Qt.Key.Key_Space)
    assert right.answer.text() == "(Absent de quiz_fr.xlsx)" and left.answer.text() == "EN answer B4"
    win.bilingual_action.trigger()                               # off mid-question: the question stays
    assert not right.isVisible() and clue.revealed and left.answer.text() == "EN answer B4"
    QTest.keyClick(win, Qt.Key.Key_Escape)
    win.bilingual_action.trigger()                               # back on: the partner is kept, no prompt
    assert win._bilingual_on() and not page.cells[0].secondary.isHidden()
    assert win.scores() == [2, 0] and len(win.played) == 2      # switching modes never touches the game

    reports.clear()
    _expect_dialog(reports)
    win.reload_action.trigger()                                  # Reload re-reads both files
    _activate(win)
    assert reports and win._bilingual_on() and win.scores() == [2, 0]

    win._confirm = lambda text: True
    reports.clear()
    _expect_dialog(reports)
    assert win.open_path(str(folder / "quiz_fr.xlsx"))           # the opened file goes on the left
    _activate(win)
    assert (page.cells[0].primary.text(), page.cells[0].secondary.text()) == ("Sujet A", "Topic A")
    assert page.tiles[(1, 4)].state == "missing"                 # the opened file decides the tiles
    assert any("quiz_en.xlsx has a question for points 4 that quiz_fr.xlsx lacks" in w
               for w in win.partner_warnings), win.partner_warnings

    infos = []
    _expect_dialog(infos)
    assert win.open_path(str(folder / "lonely.xlsx"))            # no partner: said so, mode off
    _activate(win)
    assert infos and infos[0] and "none was found" in infos[0], infos
    assert not win.bilingual and not win.bilingual_action.isChecked() and not page.cells[0].secondary.isVisible()
    win.close()


@check
def gui_french_interface():
    from PyQt6.QtCore import QEvent, Qt
    from PyQt6.QtTest import QTest
    from PyQt6.QtWidgets import QApplication, QMessageBox

    import quizboard

    win = _gui_window(teams=["Rouge", "Bleu"])
    win.set_language("fr")
    assert [m.title() for m in (a.menu() for a in win.menuBar().actions())] == ["&Fichier", "&Partie", "&Affichage"]
    assert win.start_page.buttons[0][0].text() == "Ouvrir un classeur…"
    box = QMessageBox(QMessageBox.Icon.Question, "t", "t", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
    assert box.button(QMessageBox.StandardButton.Yes).text().replace("&", "") == "Oui"   # Qt's own text too
    dialog = quizboard.TeamsDialog(win, [], win.voice)
    assert dialog.windowTitle() == "Nouvelle partie" and dialog.names() == ["Équipe 1", "Équipe 2"]
    # room for the title in the window manager's title bar (it clipped "New Game" at 211 px)
    assert dialog.minimumWidth() >= dialog.fontMetrics().horizontalAdvance("Nouvelle partie") + 152
    QApplication.sendEvent(dialog.buttons, QEvent(QEvent.Type.LanguageChange))   # Qt re-labels standard buttons
    assert dialog.start_button.text() == "Commencer la partie", dialog.start_button.text()
    dialog.close()

    assert win.open_path(demo_board())                           # no _en/_fr in the name: the room follows the UI
    page, clue, scores = win.board_page, win.clue_page, win.score_page
    assert "0 jouées sur 36" in win.status_label.text()
    QTest.mouseClick(page.tiles[(0, 1)], Qt.MouseButton.LeftButton)
    assert clue.show_answer_button.text() == "Afficher la réponse  (Espace)"
    assert clue.heading.text().endswith("1\u00a0point")
    QTest.keyClick(win, Qt.Key.Key_Space)
    assert clue.hint.text().startswith("Cochez la ou les équipes")
    QTest.keyClick(win, Qt.Key.Key_S)
    assert scores.title.text() == "Pointages" and scores.subtitle.text() == "0 questions jouées sur 36"
    QTest.keyClick(win, Qt.Key.Key_Escape)                       # back to the open question
    win.set_language("en")                                       # switching back retranslates in place
    assert win.stack.currentWidget() is clue and clue.revealed
    assert clue.hint.text().startswith("Tick the team(s)") and win.menuBar().actions()[0].menu().title() == "&File"
    QTest.keyClick(win, Qt.Key.Key_Escape)
    win.close()


@check
def gui_bilingual_mode_speaks_both_languages_to_the_room():
    from PyQt6.QtCore import Qt
    from PyQt6.QtTest import QTest

    folder = TMP / "bilingual_room"
    folder.mkdir()
    for language, header in (("en", BOARD_HEADER), ("fr", ["Catégorie", "Points", "Question", "Réponse"])):
        rows = [header] + [[f"{language} {t}", p, f"{language} q{t}{p}", f"{language} a{t}{p}"]
                           for t in "ABCDEF" for p in range(1, 7)]
        write_workbook(str(folder / f"board_{language}.xlsx"), [SheetSpec("Board", rows)])

    win = _gui_window(teams=["Red", "Blue"])
    win._set_bilingual(True)
    assert win.open_path(str(folder / "board_fr.xlsx")) and win._bilingual_on()
    page, clue, scores = win.board_page, win.clue_page, win.score_page
    assert win.voice.room == ["fr", "en"] and win.voice.ui == "en"   # French file opened: French first
    QTest.mouseClick(page.tiles[(2, 2)], Qt.MouseButton.LeftButton)
    assert clue.show_answer_button.text() == "Afficher la réponse  (Espace) / Show Answer  (Space)"
    assert clue.hint.text().split("\n")[1] == "Click the question or press Space to show the answer"
    assert clue.sides[0].question.text() == "fr qC2" and clue.sides[1].question.text() == "en qC2"
    QTest.keyClick(win, Qt.Key.Key_Space)
    QTest.keyClick(win, Qt.Key.Key_2)
    QTest.keyClick(win, Qt.Key.Key_Space)
    for (c, p), tile in sorted(page.tiles.items()):              # play out the board
        if tile.state == "ready":
            QTest.mouseClick(tile, Qt.MouseButton.LeftButton)
            QTest.keyClick(win, Qt.Key.Key_Space)
            QTest.keyClick(win, Qt.Key.Key_Space)
    assert win.stack.currentWidget() is scores
    assert scores.title.text() == "Résultats finaux / Final Scores"
    assert scores.subtitle.text() == "Blue gagne avec 2 points!\nBlue wins with 2 points!", scores.subtitle.text()
    assert "Menu" not in win.status_label.text() and "36 of 36 played" in win.status_label.text()   # host: English
    win.close()


@check
def gui_undo_after_esc_by_mistake():
    from PyQt6.QtCore import Qt
    from PyQt6.QtTest import QTest

    win = _gui_window(teams=["Red", "Blue"])
    assert win.open_path(demo_board())
    page, clue, scores = win.board_page, win.clue_page, win.score_page
    undo, score_undo = page.undo_button, scores.undo_button
    assert not win.undo_action.isEnabled() and not undo.isVisible()

    QTest.mouseClick(page.tiles[(0, 3)], Qt.MouseButton.LeftButton)
    QTest.keyClick(win, Qt.Key.Key_Escape)                       # oops: Esc before the answer
    assert win.stack.currentWidget() is page and page.tiles[(0, 3)].state == "played"
    assert undo.isVisible() and win.undo_action.isEnabled()
    assert not any(undo.geometry().intersects(t.geometry()) for t in page.tiles.values())   # covers no tile
    QTest.mouseClick(undo, Qt.MouseButton.LeftButton)
    assert win.stack.currentWidget() is clue and win.current == (0, 3) and not clue.revealed
    assert clue.question.text() == "Question A3" and not undo.isVisible()
    assert page.tiles[(0, 3)].state == "ready" and not win.played  # the tile is unplayed again

    QTest.keyClick(win, Qt.Key.Key_Space)
    QTest.keyClick(win, Qt.Key.Key_1)                            # Red earns 3
    QTest.keyClick(win, Qt.Key.Key_Escape)                       # oops again, answer showing
    QTest.keyClick(win, Qt.Key.Key_Z, Qt.KeyboardModifier.ControlModifier)
    assert win.stack.currentWidget() is clue and clue.revealed and clue.answer.text() == "Answer A3"
    assert clue.team_boxes[0].isChecked() and win.scores() == [3, 0]   # ticks kept, not doubled

    QTest.keyClick(win, Qt.Key.Key_Space)                        # leave properly
    QTest.mouseClick(page.tiles[(1, 1)], Qt.MouseButton.LeftButton)
    assert not win.undo_action.isEnabled()                       # another question ends the undo
    QTest.keyClick(win, Qt.Key.Key_Escape)

    QTest.mouseClick(page.tiles[(2, 2)], Qt.MouseButton.LeftButton)
    QTest.keyClick(win, Qt.Key.Key_Escape)
    QTest.qWait(400)
    assert undo.isVisible() and win.undo_action.isEnabled()      # no timeout: it stays until it no longer applies
    win.undo_action.trigger()
    assert win.stack.currentWidget() is clue and win.current == (2, 2)
    QTest.keyClick(win, Qt.Key.Key_Escape)

    for (c, p), tile in sorted(page.tiles.items()):              # play out; the last one left with Esc
        if tile.state == "ready":
            QTest.mouseClick(tile, Qt.MouseButton.LeftButton)
            QTest.keyClick(win, Qt.Key.Key_Escape)
    assert win.stack.currentWidget() is scores and scores.final and score_undo.isVisible()
    QTest.mouseClick(score_undo, Qt.MouseButton.LeftButton)      # undo from the final scores
    assert win.stack.currentWidget() is clue and not win._complete()
    QTest.keyClick(win, Qt.Key.Key_Escape)
    assert win.stack.currentWidget() is scores and scores.final  # and they come back when it's left
    win._confirm = lambda text: True
    win.new_game()                                               # a new game takes the button away
    assert not undo.isVisible() and not win.undo_action.isEnabled()
    win.close()


def main() -> int:
    failed = 0
    for fn in CHECKS:
        try:
            fn()
        except Exception:
            failed += 1
            print(f"FAIL  {fn.__name__}")
            traceback.print_exc()
        else:
            print(f"ok    {fn.__name__}")
    print(f"\n{len(CHECKS) - failed} of {len(CHECKS)} checks passed.")
    shutil.rmtree(TMP, ignore_errors=True)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
