"""Quizboard spreadsheet I/O — read a 6 × 6 board of questions from an .xlsx file.

Standard library only (zipfile + ElementTree): an .xlsx file is a zip of XML
parts, and reading cell values out of it needs nothing more, so the app runs on
the system python3 without openpyxl.

Spreadsheet layout
------------------
Any sheet whose header row has a **Question** and an **Answer** column is read.
The other recognised columns are **Topic** and **Points** (French headers such
as Sujet / Thème / Réponse / Valeur work too):

    Topic       | Points | Question                         | Answer
    Geography   | 1      | What is the capital of Canada?   | Ottawa
    Geography   | 2      | ...                              | ...
    ...
    Science     | 1      | ...                              | ...

* The topic may be written once at the top of its block and left blank on the
  rows below it — a blank topic carries down from the row above. A row with a
  topic but no question or answer just starts that topic's block.
* With no Topic column, each sheet is one topic, named after the sheet tab.
* Points are 1 to 6 and say which row of the board the question sits in. A
  blank Points cell takes the topic's next free row.
* Topics become board columns in the order they first appear; the first six are
  used.

Anything the reader has to drop or cannot place is reported in Board.warnings —
nothing is skipped silently.

Bilingual pairs
---------------
Two language versions of a board are paired by file name — board_en.xlsx and
board_fr.xlsx — and matched tile by tile on (column, points). partner_path()
finds the other file; compare_boards() lists the tiles that don't line up.
"""

from __future__ import annotations

import datetime as _dt
import posixpath
import re
import unicodedata
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as _xml_escape

from i18n import HOW_TO, TEMPLATE_HEADER, Message, Quoted, tr

COLUMNS = 6   # topics across the board
ROWS = 6      # point values down the board, 1..ROWS

HEADER_SCAN_ROWS = 20   # how far down a sheet to look for the header row


class BoardError(Exception):
    """The file cannot be read as a board at all. args[0] is an i18n Message."""


@dataclass
class Clue:
    topic: str
    points: int
    question: str
    answer: str
    where: str   # "sheet 'Board', row 7" — an i18n Message, for messages


@dataclass
class Board:
    path: str
    topics: list[str]                                   # board columns, ≤ COLUMNS
    clues: dict[tuple[int, int], Clue] = field(default_factory=dict)   # (column, points)
    warnings: list[str] = field(default_factory=list)
    # The warnings that only say tiles or topics are absent. A bilingual partner's are
    # left out of its report, since compare_boards() says the same thing tile by tile.
    gaps: list[str] = field(default_factory=list)

    def clue(self, column: int, points: int) -> Clue | None:
        return self.clues.get((column, points))


# ---------------------------------------------------------------------------
# Reading the workbook
# ---------------------------------------------------------------------------

_REL_OFFICE_DOC = "/officeDocument"
_REL_WORKSHEET = "/worksheet"
_REL_SHARED_STRINGS = "/sharedStrings"
_REL_STYLES = "/styles"

# Built-in number formats that display a date or time.
_BUILTIN_DATE_FORMATS = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))
_BUILTIN_PERCENT_FORMATS = {9, 10}
_FORMAT_NOISE = re.compile(r'"[^"]*"|\\.|_.|\*.|\[(?!h\]|m\]|s\])[^\]]*\]')
_ESCAPED_CHAR = re.compile(r"_x([0-9A-Fa-f]{4})_")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _children(el: ET.Element, name: str):
    return (c for c in el if _local(c.tag) == name)


def _child(el: ET.Element, name: str) -> ET.Element | None:
    return next(_children(el, name), None)


def _rel_id(el: ET.Element) -> str | None:
    # r:id, whichever relationships namespace (transitional or strict) the file uses.
    for key, value in el.attrib.items():
        if key.startswith("{") and _local(key) == "id":
            return value
    return None


def _read_xml(zf: zipfile.ZipFile, part: str) -> ET.Element:
    return ET.fromstring(zf.read(part))


def _rels(zf: zipfile.ZipFile, part: str) -> list[tuple[str, str]]:
    """(type, absolute target part, id) for each internal relationship of `part`."""
    folder, name = posixpath.split(part)
    rels_part = posixpath.join(folder, "_rels", name + ".rels")
    if rels_part not in zf.namelist():
        return []
    out = []
    for rel in _read_xml(zf, rels_part):
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        out.append((rel.get("Type", ""), target, rel.get("Id", "")))
    return out


def _unescape(text: str) -> str:
    # OOXML writes some control characters as _xHHHH_; Excel line breaks may be \r\n.
    text = _ESCAPED_CHAR.sub(lambda m: chr(int(m.group(1), 16)), text)
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _rich_text(el: ET.Element) -> str:
    """Text of a shared-string <si> or inline <is>: plain <t> plus rich-text runs, no phonetics."""
    parts = []
    for child in el:
        name = _local(child.tag)
        if name == "t":
            parts.append(child.text or "")
        elif name == "r":
            t = _child(child, "t")
            if t is not None:
                parts.append(t.text or "")
    return _unescape("".join(parts))


def _is_date_format(code: str) -> bool:
    code = _FORMAT_NOISE.sub("", code.split(";", 1)[0])
    return bool(re.search(r"[dmyhs]", code, re.IGNORECASE)) and "general" not in code.lower()


def _format_number(raw: str) -> str:
    try:
        value = float(raw)
    except ValueError:
        return raw
    if value.is_integer() and abs(value) < 1e15:
        return str(int(value))
    return format(value, ".15g")


def _format_date(raw: str, date1904: bool) -> str:
    try:
        serial = float(raw)
    except ValueError:
        return raw
    base = _dt.datetime(1904, 1, 1) if date1904 else _dt.datetime(1899, 12, 30)
    try:
        moment = base + _dt.timedelta(days=serial)
    except OverflowError:
        return _format_number(raw)
    if serial < 1:
        return moment.strftime("%H:%M")
    day = f"{moment.day} {moment:%B} {moment.year}"
    if serial.is_integer():
        return day
    return f"{day} {moment:%H:%M}"


class _Styles:
    """Which cell style indices display numbers as dates or percentages."""

    def __init__(self, root: ET.Element | None):
        self.dates: set[int] = set()
        self.percents: set[int] = set()
        if root is None:
            return
        custom = {}
        num_fmts = _child(root, "numFmts")
        if num_fmts is not None:
            for nf in _children(num_fmts, "numFmt"):
                custom[int(nf.get("numFmtId", "0"))] = nf.get("formatCode", "")
        cell_xfs = _child(root, "cellXfs")
        if cell_xfs is None:
            return
        for index, xf in enumerate(_children(cell_xfs, "xf")):
            fmt_id = int(xf.get("numFmtId", "0"))
            code = custom.get(fmt_id)
            if fmt_id in _BUILTIN_DATE_FORMATS or (code is not None and _is_date_format(code)):
                self.dates.add(index)
            elif fmt_id in _BUILTIN_PERCENT_FORMATS or (code is not None and code.rstrip().endswith("%")):
                self.percents.add(index)


def _column_index(ref: str) -> int:
    """'A1' -> 0, 'AB12' -> 27."""
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def _cell_text(c: ET.Element, shared: list[str], styles: _Styles, date1904: bool) -> str:
    kind = c.get("t", "n")
    if kind == "inlineStr":
        is_el = _child(c, "is")
        return _rich_text(is_el) if is_el is not None else ""
    v = _child(c, "v")
    raw = v.text if v is not None and v.text is not None else ""
    if kind == "s":
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return ""
    if kind == "b":
        return "TRUE" if raw == "1" else "FALSE"
    if kind in ("str", "e", "d"):
        return _unescape(raw)
    if raw == "":
        return ""
    style = int(c.get("s", "0"))
    if style in styles.dates:
        return _format_date(raw, date1904)
    if style in styles.percents:
        try:
            return _format_number(repr(float(raw) * 100)) + "%"
        except ValueError:
            return raw
    return _format_number(raw)


def read_workbook(path: str) -> list[tuple[str, list[list[str]]]]:
    """Every sheet as (name, rows); each row a list of cell strings ('' when empty)."""
    try:
        zf = zipfile.ZipFile(path)
    except FileNotFoundError:
        raise BoardError(Message("e.not_found", path=path)) from None
    except PermissionError:
        raise BoardError(Message("e.permission", path=path)) from None
    except zipfile.BadZipFile:
        raise BoardError(Message("e.not_xlsx")) from None
    with zf:
        try:
            workbook_part = next(t for typ, t, _ in _rels(zf, "") if typ.endswith(_REL_OFFICE_DOC))
        except StopIteration:
            workbook_part = "xl/workbook.xml"
        try:
            workbook = _read_xml(zf, workbook_part)
        except KeyError:
            raise BoardError(Message("e.no_workbook")) from None

        rels = _rels(zf, workbook_part)
        by_id = {rid: target for _, target, rid in rels}

        shared: list[str] = []
        styles = _Styles(None)
        for typ, target, _ in rels:
            if typ.endswith(_REL_SHARED_STRINGS) and target in zf.namelist():
                shared = [_rich_text(si) for si in _children(_read_xml(zf, target), "si")]
            elif typ.endswith(_REL_STYLES) and target in zf.namelist():
                styles = _Styles(_read_xml(zf, target))

        pr = _child(workbook, "workbookPr")
        date1904 = pr is not None and pr.get("date1904", "0") in ("1", "true")

        sheets = []
        sheets_el = _child(workbook, "sheets")
        for sheet in (_children(sheets_el, "sheet") if sheets_el is not None else ()):
            target = by_id.get(_rel_id(sheet) or "")
            if target is None or target not in zf.namelist():
                continue
            sheets.append((sheet.get("name", "Sheet"), _sheet_rows(_read_xml(zf, target), shared, styles, date1904)))
        return sheets


def _sheet_rows(root: ET.Element, shared: list[str], styles: _Styles, date1904: bool) -> list[list[str]]:
    data = _child(root, "sheetData")
    if data is None:
        return []
    rows: dict[int, dict[int, str]] = {}
    next_row = 0
    for row in _children(data, "row"):
        r = int(row.get("r", next_row + 1)) - 1
        next_row = r + 1
        cells = rows.setdefault(r, {})
        next_col = 0
        for c in _children(row, "c"):
            ref = c.get("r")
            col = _column_index(ref) if ref else next_col
            next_col = col + 1
            text = _cell_text(c, shared, styles, date1904)
            if text != "":
                cells[col] = text
    if not rows:
        return []
    out = []
    for r in range(max(rows) + 1):
        cells = rows.get(r, {})
        width = max(cells) + 1 if cells else 0
        out.append([cells.get(i, "") for i in range(width)])
    return out


# ---------------------------------------------------------------------------
# Turning sheets into a board
# ---------------------------------------------------------------------------

_HEADER_NAMES = {
    "topic": {"topic", "topics", "column topic", "category", "categories", "column", "subject",
              "theme", "sujet", "categorie", "rubrique", "colonne"},
    "points": {"points", "point", "pts", "pt", "value", "points value", "point value", "score",
               "valeur", "pointage"},
    "question": {"question", "questions", "clue", "clues", "indice"},
    "answer": {"answer", "answers", "response", "reponse", "reponses", "solution"},
}


def _normalise(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^\w\s]", " ", text.lower())
    return " ".join(text.split())


def _find_header(rows: list[list[str]]) -> tuple[int, dict[str, int]] | None:
    for r, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        columns: dict[str, int] = {}
        for c, cell in enumerate(row):
            name = _normalise(cell)
            for role, names in _HEADER_NAMES.items():
                if name in names and role not in columns:
                    columns[role] = c
        if "question" in columns and "answer" in columns:
            return r, columns
    return None


def _cell(row: list[str], col: int | None) -> str:
    if col is None or col >= len(row):
        return ""
    return row[col].strip()


def _parse_points(text: str) -> int | None:
    text = re.sub(r"(?i)\s*(points?|pts?)\s*$", "", text.strip())
    try:
        value = float(text)
    except ValueError:
        return None
    return int(value) if value.is_integer() else None


def _points_list(numbers) -> str:
    numbers = sorted(numbers)
    return ", ".join(str(n) for n in numbers)


def load_board(path: str) -> Board:
    """Read `path` into a Board. Raises BoardError when nothing usable is found."""
    sheets = read_workbook(path)
    if not sheets:
        raise BoardError(Message("e.no_sheets"))

    warnings: list[str] = []
    gaps: list[str] = []
    topic_names: dict[str, str] = {}               # casefolded key -> first spelling
    entries: dict[str, list[tuple[str, str, str, str]]] = {}   # key -> (points, q, a, where)
    read_any = False

    for sheet_name, rows in sheets:
        found = _find_header(rows)
        if found is None:
            continue
        read_any = True
        header_row, cols = found
        topic_col, points_col = cols.get("topic"), cols.get("points")
        current = None if topic_col is not None else sheet_name.strip()
        if current:
            topic_names.setdefault(current.casefold(), current)
            entries.setdefault(current.casefold(), [])
        for r in range(header_row + 1, len(rows)):
            row = rows[r]
            where = Message("where", sheet=sheet_name, row=r + 1)
            topic = _cell(row, topic_col)
            if topic:
                current = topic
                topic_names.setdefault(topic.casefold(), topic)
                entries.setdefault(topic.casefold(), [])
            question, answer = _cell(row, cols["question"]), _cell(row, cols["answer"])
            if not question and not answer:
                continue
            if current is None:
                warnings.append(Message("w.no_topic_above", where=where))
                continue
            entries[current.casefold()].append((_cell(row, points_col), question, answer, where))

    if not read_any:
        raise BoardError(Message("e.no_header"))
    if not topic_names:
        raise BoardError(Message("e.no_topics"))

    keys = list(topic_names)
    if len(keys) > COLUMNS:
        dropped = Quoted(topic_names[k] for k in keys[COLUMNS:])
        warnings.append(Message("w.extra_topics", columns=COLUMNS, topics=dropped))
        keys = keys[:COLUMNS]
    elif len(keys) < COLUMNS:
        warnings.append(Message("w.few_topics", n=len(keys)))
        gaps.append(warnings[-1])

    board = Board(path=path, topics=[topic_names[k] for k in keys])
    for column, key in enumerate(keys):
        topic = topic_names[key]
        placed: dict[int, tuple[str, str, str]] = {}
        unnumbered: list[tuple[str, str, str]] = []
        for points_text, question, answer, where in entries[key]:
            if points_text == "":
                unnumbered.append((question, answer, where))
                continue
            points = _parse_points(points_text)
            if points is None or not 1 <= points <= ROWS:
                warnings.append(Message("w.bad_points", where=where, points=points_text, rows=ROWS))
                continue
            if points in placed:
                warnings.append(Message("w.duplicate", where=where, topic=topic, points=points,
                                        first=placed[points][2]))
                continue
            placed[points] = (question, answer, where)
        for question, answer, where in unnumbered:
            free = next((p for p in range(1, ROWS + 1) if p not in placed), None)
            if free is None:
                warnings.append(Message("w.no_free_row", where=where, topic=topic))
                continue
            placed[free] = (question, answer, where)

        for points, (question, answer, where) in placed.items():
            amount = f"{points} point" if points == 1 else f"{points} points"   # the same in French
            if not question:
                warnings.append(Message("w.no_question", where=where, topic=topic, amount=amount))
            if not answer:
                warnings.append(Message("w.no_answer", where=where, topic=topic, amount=amount))
            board.clues[(column, points)] = Clue(topic, points, question, answer, where)
        missing = [p for p in range(1, ROWS + 1) if p not in placed]
        if missing:
            warnings.append(Message("w.missing_points", topic=topic, points=_points_list(missing)))
            gaps.append(warnings[-1])

    board.warnings = warnings
    board.gaps = gaps
    return board


# ---------------------------------------------------------------------------
# Bilingual pairs: board_en.xlsx + board_fr.xlsx
# ---------------------------------------------------------------------------

_LANGUAGE_SUFFIX = re.compile(r"^(.*)_(en|fr)$", re.IGNORECASE)
_OTHER_LANGUAGE = {"en": "fr", "fr": "en"}


def language_of(path: str) -> str | None:
    """'en' or 'fr' from a file named …_en.xlsx / …_fr.xlsx; None otherwise."""
    match = _LANGUAGE_SUFFIX.match(Path(path).stem)
    return match.group(2).lower() if match else None


def partner_path(path: str) -> str | None:
    """The other-language file of a pair (…_en.xlsx ↔ …_fr.xlsx), if it exists."""
    p = Path(path)
    match = _LANGUAGE_SUFFIX.match(p.stem)
    if match is None:
        return None
    language = match.group(2)
    other = _OTHER_LANGUAGE[language.lower()]
    candidate = p.with_name(f"{match.group(1)}_{other.upper() if language.isupper() else other}{p.suffix}")
    return str(candidate) if candidate.is_file() else None


def compare_boards(primary: Board, secondary: Board) -> list[Message]:
    """Where two language versions of a board don't line up, tile by tile (column, points)."""
    a, b = Path(primary.path).name, Path(secondary.path).name
    warnings = []
    if len(primary.topics) != len(secondary.topics):
        warnings.append(Message("c.topic_count", a=a, n=len(primary.topics), b=b, m=len(secondary.topics)))
    for column in range(COLUMNS):
        name = primary.topics[column] if column < len(primary.topics) else f"column {column + 1}"
        only_primary = [p for p in range(1, ROWS + 1)
                        if (column, p) in primary.clues and (column, p) not in secondary.clues]
        only_secondary = [p for p in range(1, ROWS + 1)
                          if (column, p) in secondary.clues and (column, p) not in primary.clues]
        if only_primary:
            warnings.append(Message("c.missing", column=column + 1, name=name, b=b,
                                    points=_points_list(only_primary)))
        if only_secondary:
            warnings.append(Message("c.extra", column=column + 1, name=name, a=a, b=b,
                                    points=_points_list(only_secondary)))
    return warnings


# ---------------------------------------------------------------------------
# Writing a workbook (the blank template)
# ---------------------------------------------------------------------------

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# cellXfs indices written into styles.xml below.
_STYLE_BODY = 1     # top-aligned, wrapped
_STYLE_HEADER = 2   # bold on a light fill

_STYLES_XML = (
    _XML_DECL
    + f'<styleSheet xmlns="{_MAIN_NS}">'
    '<fonts count="2">'
    '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    "</fonts>"
    '<fills count="3">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFDCE3F2"/><bgColor indexed="64"/></patternFill></fill>'
    "</fills>"
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="3">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
    '<alignment vertical="top" wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    "</cellXfs>"
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)


@dataclass
class SheetSpec:
    name: str
    rows: list[list[str | int]]
    widths: list[float] = field(default_factory=list)   # column widths in characters
    header: bool = True                                  # bold, frozen first row


def _column_letter(index: int) -> str:
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _sheet_xml(spec: SheetSpec, selected: bool) -> str:
    out = [_XML_DECL, f'<worksheet xmlns="{_MAIN_NS}">']
    tab_selected = ' tabSelected="1"' if selected else ""
    out.append(f'<sheetViews><sheetView workbookViewId="0"{tab_selected}>')
    if spec.header:
        out.append('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
    out.append("</sheetView></sheetViews>")
    out.append('<sheetFormatPr defaultRowHeight="15"/>')
    if spec.widths:
        out.append("<cols>")
        for i, width in enumerate(spec.widths, start=1):
            out.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
        out.append("</cols>")
    out.append("<sheetData>")
    for r, row in enumerate(spec.rows, start=1):
        style = _STYLE_HEADER if spec.header and r == 1 else _STYLE_BODY
        cells = []
        for c, value in enumerate(row):
            ref = f"{_column_letter(c)}{r}"
            if value is None or value == "":
                continue
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append(f'<c r="{ref}" s="{style}"><v>{value}</v></c>')
            else:
                text = _INVALID_XML_CHARS.sub("", str(value))
                space = ' xml:space="preserve"' if text != text.strip() or "\n" in text else ""
                cells.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t{space}>{_xml_escape(text)}</t></is></c>')
        out.append(f'<row r="{r}">{"".join(cells)}</row>')
    out.append("</sheetData></worksheet>")
    return "".join(out)


def write_workbook(path: str, sheets: list[SheetSpec]) -> None:
    """Write a plain .xlsx (inline strings, a bold header row) that Excel opens directly."""
    content_types = [
        _XML_DECL,
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    workbook = [_XML_DECL, f'<workbook xmlns="{_MAIN_NS}" xmlns:r="{_REL_NS}"><sheets>']
    workbook_rels = [_XML_DECL, f'<Relationships xmlns="{_PKG_REL_NS}">']
    for i, spec in enumerate(sheets, start=1):
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        name = _xml_escape(spec.name[:31], {'"': "&quot;"})
        workbook.append(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>')
        workbook_rels.append(
            f'<Relationship Id="rId{i}" Type="{_REL_NS}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        )
    workbook_rels.append(
        f'<Relationship Id="rId{len(sheets) + 1}" Type="{_REL_NS}/styles" Target="styles.xml"/>'
    )
    content_types.append("</Types>")
    workbook.append("</sheets></workbook>")
    workbook_rels.append("</Relationships>")
    root_rels = (
        _XML_DECL
        + f'<Relationships xmlns="{_PKG_REL_NS}">'
        f'<Relationship Id="rId1" Type="{_REL_NS}/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "".join(content_types))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", "".join(workbook))
        zf.writestr("xl/_rels/workbook.xml.rels", "".join(workbook_rels))
        zf.writestr("xl/styles.xml", _STYLES_XML)
        for i, spec in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(spec, selected=(i == 1)))


BOARD_HEADER = TEMPLATE_HEADER["en"]
BOARD_WIDTHS = [24, 8, 60, 40]


def write_template(path: str, language: str = "en") -> None:
    """A blank board: six placeholder topics, points 1–6 each, empty questions and answers —
    headers, tab names and instructions in `language`."""
    rows: list[list[str | int]] = [list(TEMPLATE_HEADER[language])]
    for t in range(1, COLUMNS + 1):
        for p in range(1, ROWS + 1):
            rows.append([tr("template.topic", language, n=t), p, "", ""])
    write_workbook(path, [
        SheetSpec(tr("template.board_sheet", language), rows, BOARD_WIDTHS),
        SheetSpec(tr("template.howto_sheet", language), [[line] for line in HOW_TO[language]], [100], header=False),
    ])
