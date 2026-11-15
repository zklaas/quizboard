# Quizboard

A 6 × 6 board of point values, questions and answers, read from an Excel
workbook. Six column topics run across the top; under each topic, six tiles show
their points, 1 to 6 going down.

## Download

Ready-to-run downloads (executable programs), with nothing else to install:

* **Windows:** [Quizboard-Setup-Windows.exe](https://github.com/zklaas/quizboard/releases/latest/download/Quizboard-Setup-Windows.exe)
* **macOS** (Apple silicon): [Quizboard-macOS.dmg](https://github.com/zklaas/quizboard/releases/latest/download/Quizboard-macOS.dmg)
* **Linux:** [Quizboard-Linux-x86_64.AppImage](https://github.com/zklaas/quizboard/releases/latest/download/Quizboard-Linux-x86_64.AppImage)

The first time, Windows or macOS asks before opening it; the
[release page](https://github.com/zklaas/quizboard/releases/latest) says what to click.
To install from the source code at the command line instead, see below.

* Click a tile → its question fills the window.
* Click the question (or press Space) → the answer appears under it.
* Click again (or press Space or Esc) → back to the board, where that tile is now blank.
* Left a question by mistake (a stray Esc)? **Undo** (the gold button below the
  tiles, Ctrl+Z, or Game ▸ Undo) reopens it as it was — answer hidden or shown,
  team ticks kept — and puts the tile back on the board. It stays available until
  the next tile is opened (or a new game / reload), and is on the Final Scores page too.

## Teams and scores

* Each game starts by asking **how many teams are playing** (1–8), with optional
  team names. This happens when a board is opened and on **Game ▸ New Game…**.
  Cancel it to play without keeping score.
* Once an answer is shown, a **checkbox per team** appears under it. Tick the
  team(s) that earn the tile's points (or press 1–8); untick to take them back.
  Each box shows that team's running total.
* Reopening a played tile (right-click ▸ Show Question) shows its ticks, so a
  mistake can be corrected.
* **Game ▸ Scores** (S) shows the standings at any time; Esc goes back to the
  board, or to the question if one was open. The status bar also keeps a running
  total.
* When the last tile has been played, the **Final Scores** appear, naming the
  winner (or the tie).

Standalone — not part of ZIS. Runs on the system `python3` with PyQt6 and
nothing else: the .xlsx reader is standard library only.

## Run

```bash
quizboard                       # reopens the last workbook, or shows the start page
quizboard path/to/board.xlsx
quizboard --check               # regression checks (headless)
```

`quizboard` is a symlink `~/.local/bin/quizboard` → `run.sh`. Without the
symlink: `./run.sh` or `python3 quizboard.py`.

## Bilingual mode

**View ▸ Bilingual Mode** (Ctrl+B) shows two language versions of the board side
by side: each topic heading carries both names, and each question and answer
appears in both languages, English and French at the same size.

* The two versions are separate workbooks named as a pair: `board_en.xlsx` and
  `board_fr.xlsx`. Opening either one finds the other automatically.
* The file you opened goes on the left and decides the game (tiles, points,
  scores); the other supplies the second language, matched tile by tile on
  column and points. Open the `_fr` file to put French on the left.
* A tile missing from the other file shows "(Not in …)" in its column, and the
  load report lists every tile that doesn't line up.
* No partner file? Quizboard says so and turns bilingual mode off.
  **View ▸ Choose Other-Language File…** pairs any two workbooks by hand.
* Switching the mode on or off never touches the game in progress. The setting
  is remembered.
* What the room sees — the buttons and hints on the question page, the scores
  page — is in both languages too, the left column's first.

## Interface language

**View ▸ Language** switches the interface between English and Français: menus,
dialogs, the status bar, the load report and Qt's own buttons (Oui / Non /
Annuler). The choice is remembered. In French, File ▸ Save Blank Template… writes
a French template (Catégorie | Points | Question | Réponse, tabs Tableau and
Mode d'emploi).

The question and score pages speak the board's language: a file named `…_fr.xlsx`
shows French buttons and hints even with an English interface (and both
languages in bilingual mode). All text lives in `i18n.py` as English–French
pairs.

## The spreadsheet

One row per question, under a header row of **Topic | Points | Question | Answer**:

| Topic              | Points | Question                                    | Answer        |
|--------------------|-------:|---------------------------------------------|---------------|
| Canadian Geography | 1      | What is the capital of Canada?              | Ottawa        |
| Canadian Geography | 2      | Which province has the largest population?  | Ontario       |
| …                  |        |                                             |               |
| Ships & the Sea    | 1      | What is the front of a ship called?         | The bow       |

* **File ▸ Save Blank Template…** writes this layout with six placeholder topics,
  points 1–6 each. The `examples/` folder holds a complete bilingual pair.
* Topics become the board columns in the order they first appear. The first six are used.
* Points (1–6) say which row a question sits in. A blank Points cell takes the
  topic's next free row.
* The topic can be written once at the top of its block and left blank below —
  a blank topic carries down. A row with only a topic starts that topic's block.
* With no Topic column, each sheet is one topic, named after its tab.
* Sheets without Question and Answer headers (notes, instructions) are ignored.
  French headers work too: Sujet / Thème / Catégorie, Valeur, Réponse.
* Alt+Enter line breaks inside a cell show on the board. Dates show as "1 July 2017".

Anything that can't be placed — a 7th topic, points outside 1–6, two questions
for the same tile, a question with no answer — is listed in a load report when
the file opens (View ▸ Load Report shows it again). Empty tiles stay on the board,
greyed out.

## Controls

| Key / action           | Does                                                     |
|------------------------|----------------------------------------------------------|
| Click tile             | Show its question                                        |
| Click / Space / Enter  | Show the answer, then go back to the board               |
| 1–8                    | Tick / untick a team once the answer is shown            |
| Esc / Backspace        | Back to the board (the tile still counts as played)      |
| Ctrl+Z                 | Undo — reopen the question just left (Esc by mistake)     |
| S                      | Scores                                                   |
| Ctrl+N                 | New Game — asks for the teams; all tiles and scores reset |
| Right-click a tile     | Show Question again · Mark as Played / Not Played        |
| F11                    | Full screen (hides the menu and status bars); Esc leaves |
| Ctrl+B                 | Bilingual mode on/off · View ▸ Language for Français     |
| F5                     | Reload the workbook(s) — played tiles and scores are kept |
| Ctrl+E                 | Edit in Excel (opens the workbook in Windows)            |
| Ctrl+O                 | Open another workbook (or drop an .xlsx on the window)   |

The last opened workbook, the last team names, bilingual mode and the interface
language are remembered in
`~/.config/quizboard/config.json`. Scores are not saved: closing the window ends the game.

## Files

| File                        | What                                                      |
|-----------------------------|-----------------------------------------------------------|
| `quizboard.py`              | The PyQt6 window: start page, board, question/answer page, teams dialog, scores |
| `board_io.py`               | .xlsx reader/writer (stdlib), the board-building rules, `_en`/`_fr` pairing |
| `i18n.py`                   | Every string in English and French; translatable warning messages |
| `_quizboard_check.py`       | Regression checks: layouts, warnings, translations, headless GUI play-throughs |
| `run.sh`                    | Launcher                                                  |
| `examples/quizboard_example_en.xlsx` / `_fr.xlsx` | A complete bilingual pair (36 questions each) |

## Windows and macOS

The same code runs on Windows and macOS. To install it there:

* **Windows:** see [WINDOWS.md](WINDOWS.md) (`install.ps1`).
* **macOS:** see [MACOS.md](MACOS.md) (`install-macos.sh`).
