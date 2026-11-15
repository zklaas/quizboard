"""Quizboard's text in English and French.

Every string the app shows — menus, buttons, hints, scores, dialogs, and the
warnings board_io writes about a spreadsheet — is an (English, French) pair in
TEXT. tr() picks one; Message is a warning or error that reads as plain English
wherever a str is expected and can be rendered in French when shown.
"""

from __future__ import annotations

LANGUAGES = ("en", "fr")
LANGUAGE_NAMES = {"en": "English", "fr": "Français"}   # each in its own language, never translated
NB = " "   # French: non-breaking space before « : » and inside « »

TEXT: dict[str, tuple[str, str]] = {
    # ---- menus ---------------------------------------------------------------
    "menu.file": ("&File", "&Fichier"),
    "menu.open": ("&Open Spreadsheet…", "&Ouvrir un classeur…"),
    "menu.reload": ("&Reload", "&Recharger"),
    "menu.edit": ("&Edit in Excel", "&Modifier dans Excel"),
    "menu.template": ("Save Blank &Template…", "Enregistrer un modèle &vierge…"),
    "menu.quit": ("&Quit", "&Quitter"),
    "menu.game": ("&Game", "&Partie"),
    "menu.undo": ("&Undo", "&Annuler"),
    "menu.new_game": ("&New Game…", "&Nouvelle partie…"),
    "menu.scores": ("&Scores", "&Pointages"),
    "menu.view": ("&View", "&Affichage"),
    "menu.full_screen": ("&Full Screen", "&Plein écran"),
    "menu.bilingual": ("&Bilingual Mode", "Mode &bilingue"),
    "menu.partner": ("&Choose Other-Language File…", "&Choisir le fichier dans l'autre langue…"),
    "menu.language": ("&Language", "&Langue"),
    "menu.report": ("Load &Report…", "&Rapport de chargement…"),

    # ---- start page ------------------------------------------------------------
    "start.intro": (
        "Open an Excel workbook with Topic, Points, Question and Answer columns —\n"
        "six topics, points 1 to 6 under each. You can also drop an .xlsx file here.",
        "Ouvrez un classeur Excel avec les colonnes Catégorie, Points, Question et Réponse —\n"
        "six catégories, avec les points 1 à 6 sous chacune. Vous pouvez aussi déposer un fichier .xlsx ici.",
    ),
    "start.open": ("Open Spreadsheet…", "Ouvrir un classeur…"),
    "start.template": ("Save Blank Template…", "Enregistrer un modèle vierge…"),

    # ---- board and question pages ----------------------------------------------
    "points.one": ("{n} point", "{n} point"),
    "points.many": ("{n} points", "{n} points"),
    "clue.show_answer": ("Show Answer  (Space)", "Afficher la réponse  (Espace)"),
    "clue.back": ("Back to Board  (Esc)", "Retour au tableau  (Échap)"),
    "clue.hint_question": (
        "Click the question or press Space to show the answer",
        "Cliquez sur la question ou appuyez sur Espace pour afficher la réponse",
    ),
    "clue.hint_teams": (
        "Tick the team(s) that earn the points (keys {keys}) · then click or press Space to go back to the board",
        "Cochez la ou les équipes qui gagnent les points (touches {keys}) · "
        "puis cliquez ou appuyez sur Espace pour revenir au tableau",
    ),
    "clue.hint_back": (
        "Click or press Space to go back to the board",
        "Cliquez ou appuyez sur Espace pour revenir au tableau",
    ),
    "clue.no_question": ("(No question in the spreadsheet for this tile)", "(Aucune question dans le classeur pour cette case)"),
    "clue.no_answer": ("(No answer in the spreadsheet)", "(Aucune réponse dans le classeur)"),
    "clue.not_in": ("(Not in {name})", "(Absent de {name})"),
    "tile.missing": ("No question in the spreadsheet for this tile", "Aucune question dans le classeur pour cette case"),
    "tile.show": ("Show Question", "Afficher la question"),
    "tile.mark_unplayed": ("Mark as Not Played", "Marquer comme non jouée"),
    "tile.mark_played": ("Mark as Played", "Marquer comme jouée"),
    "undo.button": ("Undo", "Annuler"),
    "undo.tooltip": ("Back to the question you just left", "Revenir à la question que vous venez de quitter"),

    # ---- scores ----------------------------------------------------------------
    "scores.title": ("Scores", "Pointages"),
    "scores.final": ("Final Scores", "Résultats finaux"),
    "scores.back": ("Back  (Esc)", "Retour  (Échap)"),
    "scores.new_game": ("New Game…", "Nouvelle partie…"),
    "scores.no_teams": (
        "No teams are keeping score in this game — Game ▸ New Game… asks for them.",
        "Aucune équipe ne compte les points dans cette partie — Partie ▸ Nouvelle partie… les demande.",
    ),
    "scores.progress": ("{played} of {total} questions played", "{played} questions jouées sur {total}"),
    "scores.none": ("No points were awarded", "Aucun point n'a été accordé"),
    "scores.winner": ("{name} wins with {amount}!", "{name} gagne avec {amount}!"),
    "scores.tie": ("A tie at {amount}: {names}", f"Égalité à {{amount}}{NB}: {{names}}"),
    "and": (" and ", " et "),

    # ---- teams dialog ----------------------------------------------------------
    "teams.title": ("New Game", "Nouvelle partie"),
    "teams.count": ("How many teams are playing?", "Combien d'équipes jouent?"),
    "teams.names": ("Team names (optional):", f"Noms des équipes (facultatif){NB}:"),
    "teams.default": ("Team {n}", "Équipe {n}"),
    "teams.start": ("Start Game", "Commencer la partie"),

    # ---- status bar --------------------------------------------------------------
    "status.pair": ("{name} + {partner} (bilingual)", "{name} + {partner} (bilingue)"),
    "status.played": ("{played} of {total} played", "{played} jouées sur {total}"),
    "status.no_teams": ("no teams (Game ▸ New Game… to keep score)", "aucune équipe (Partie ▸ Nouvelle partie… pour compter les points)"),
    "status.warnings.one": ("{n} load warning (View ▸ Load Report)", "{n} avertissement de chargement (Affichage ▸ Rapport de chargement)"),
    "status.warnings.many": ("{n} load warnings (View ▸ Load Report)", "{n} avertissements de chargement (Affichage ▸ Rapport de chargement)"),
    "status.excel": (
        "Opened in Excel — save it there, then press F5 here to reload.",
        "Ouvert dans Excel — enregistrez-le là-bas, puis appuyez sur F5 ici pour recharger.",
    ),

    # ---- dialogs and messages --------------------------------------------------
    "dialog.open": ("Open Spreadsheet", "Ouvrir un classeur"),
    "dialog.save_template": ("Save Blank Template", "Enregistrer un modèle vierge"),
    "dialog.partner": ("Other-Language Version of {name}", "Version de {name} dans l'autre langue"),
    "filter.open": ("Excel workbooks (*.xlsx *.xlsm);;All files (*)", "Classeurs Excel (*.xlsx *.xlsm);;Tous les fichiers (*)"),
    "filter.save": ("Excel workbook (*.xlsx)", "Classeur Excel (*.xlsx)"),
    "template.file": ("quizboard_template.xlsx", "quizboard_modele.xlsx"),
    "msg.template_failed": ("Could not save the template:\n{error}", f"Impossible d'enregistrer le modèle{NB}:\n{{error}}"),
    "msg.template_saved": (
        "Saved {name}.\n\nFill in the six topics and the questions and answers, "
        "then open it with File ▸ Open Spreadsheet…",
        "{name} est enregistré.\n\nRemplissez les six catégories, les questions et les réponses, "
        "puis ouvrez-le avec Fichier ▸ Ouvrir un classeur…",
    ),
    "msg.cannot_use": ("Could not use {name}:\n\n{error}", f"Impossible d'utiliser {{name}}{NB}:\n\n{{error}}"),
    "msg.cannot_read": ("Could not read {name}:\n\n{error}", f"Impossible de lire {{name}}{NB}:\n\n{{error}}"),
    "msg.windows_failed": ("Could not open the file in Windows:\n{error}", f"Impossible d'ouvrir le fichier dans Windows{NB}:\n{{error}}"),
    "msg.no_app": ("No application is set up to open\n{path}", "Aucune application n'est configurée pour ouvrir\n{path}"),
    "msg.discard": (
        "Open another board? The current game and its scores will be lost.",
        "Ouvrir un autre tableau? La partie en cours et ses pointages seront perdus.",
    ),
    "msg.new_game": (
        "Start a new game? Every tile goes back on the board and all scores return to zero.",
        "Commencer une nouvelle partie? Toutes les cases reviennent au tableau et tous les pointages "
        "reviennent à zéro.",
    ),
    "msg.no_partner": (
        "Bilingual mode needs the other-language version of {name}, and none was found "
        "(it looks for the same name ending in _en or _fr).\n\n"
        "Showing one language. View ▸ Choose Other-Language File… picks it by hand.",
        "Le mode bilingue a besoin de la version de {name} dans l'autre langue, et aucune n'a été trouvée "
        "(il cherche le même nom se terminant par _en ou _fr).\n\n"
        "Affichage dans une seule langue. Affichage ▸ Choisir le fichier dans l'autre langue… "
        "permet de le choisir à la main.",
    ),
    "msg.same_file": (
        "{name} is the board itself — choose its other-language version.",
        "{name} est le tableau lui-même — choisissez sa version dans l'autre langue.",
    ),
    "report.title": ("Quizboard — Load Report", "Quizboard — Rapport de chargement"),
    "report.clean": (
        "{name} loaded with no problems: {clues} of {tiles} tiles have a question.",
        f"{{name}} est chargé sans problème{NB}: {{clues}} cases sur {{tiles}} ont une question.",
    ),
    "report.problems": (
        "{name} loaded with {problems}. {clues} of {tiles} tiles have a question.",
        "{name} est chargé avec {problems}. {clues} cases sur {tiles} ont une question.",
    ),
    "problems.one": ("{n} problem", "{n} problème"),
    "problems.many": ("{n} problems", "{n} problèmes"),
    "partner.warning": ("{name}: {warning}", f"{{name}}{NB}: {{warning}}"),

    # ---- spreadsheet warnings and errors (board_io) ------------------------------
    "where": ("sheet '{sheet}', row {row}", f"feuille «{NB}{{sheet}}{NB}», ligne {{row}}"),
    "w.no_topic_above": (
        "{where}: has a question but no topic above it — skipped.",
        f"{{where}}{NB}: question sans catégorie au-dessus — ignorée.",
    ),
    "w.extra_topics": (
        "The board has {columns} columns; these extra topics were not used: {topics}.",
        f"Le tableau a {{columns}} colonnes; ces catégories en trop n'ont pas été utilisées{NB}: {{topics}}.",
    ),
    "w.few_topics": (
        "Only {n} topic(s) found — the remaining board columns are empty.",
        "Seulement {n} catégorie(s) trouvée(s) — les autres colonnes du tableau sont vides.",
    ),
    "w.bad_points": (
        "{where}: points '{points}' is not a whole number from 1 to {rows} — skipped.",
        f"{{where}}{NB}: les points «{NB}{{points}}{NB}» ne sont pas un nombre entier de 1 à {{rows}} — ignorée.",
    ),
    "w.duplicate": (
        "{where}: '{topic}' already has a {points}-point question ({first}) — this one was skipped.",
        f"{{where}}{NB}: «{NB}{{topic}}{NB}» a déjà une question pour la valeur {{points}} ({{first}}) — "
        "celle-ci a été ignorée.",
    ),
    "w.no_free_row": (
        "{where}: no points given and '{topic}' has no free row left — skipped.",
        f"{{where}}{NB}: aucun point indiqué et «{NB}{{topic}}{NB}» n'a plus de rangée libre — ignorée.",
    ),
    "w.no_question": (
        "{where}: '{topic}' {amount} has an answer but no question.",
        f"{{where}}{NB}: «{NB}{{topic}}{NB}», {{amount}}{NB}: une réponse, mais aucune question.",
    ),
    "w.no_answer": (
        "{where}: '{topic}' {amount} has a question but no answer.",
        f"{{where}}{NB}: «{NB}{{topic}}{NB}», {{amount}}{NB}: une question, mais aucune réponse.",
    ),
    "w.missing_points": (
        "'{topic}' has no question for points {points}.",
        f"«{NB}{{topic}}{NB}» n'a aucune question pour les points {{points}}.",
    ),
    "c.topic_count": ("{a} has {n} topics but {b} has {m}.", "{a} a {n} catégories, mais {b} en a {m}."),
    "c.missing": (
        "Column {column} ('{name}'): {b} has no question for points {points}.",
        f"Colonne {{column}} («{NB}{{name}}{NB}»){NB}: {{b}} n'a aucune question pour les points {{points}}.",
    ),
    "c.extra": (
        "Column {column} ('{name}'): {b} has a question for points {points} that {a} lacks — not on the board.",
        f"Colonne {{column}} («{NB}{{name}}{NB}»){NB}: {{b}} a une question pour les points {{points}} "
        "qui manque dans {a} — absente du tableau.",
    ),
    "e.not_found": ("File not found: {path}", f"Fichier introuvable{NB}: {{path}}"),
    "e.permission": ("Permission denied reading {path}", f"Lecture refusée{NB}: {{path}}"),
    "e.not_xlsx": (
        "This is not an .xlsx workbook. Older .xls files and .csv files are not read — "
        "open it in Excel and use Save As ▸ Excel Workbook (.xlsx).",
        "Ce n'est pas un classeur .xlsx. Les anciens fichiers .xls et les fichiers .csv ne sont pas lus — "
        "ouvrez-le dans Excel et choisissez Enregistrer sous ▸ Classeur Excel (.xlsx).",
    ),
    "e.no_workbook": (
        "The file is a zip archive but has no Excel workbook inside it.",
        "Le fichier est une archive zip, mais il ne contient aucun classeur Excel.",
    ),
    "e.no_sheets": ("The workbook has no sheets.", "Le classeur n'a aucune feuille."),
    "e.no_header": (
        "No sheet has a header row with 'Question' and 'Answer' columns.\n\n"
        "Expected layout (row 1 of a sheet):  Topic | Points | Question | Answer\n"
        "Use File ▸ Save Blank Template… for a ready-made sheet.",
        f"Aucune feuille n'a de ligne d'en-tête avec les colonnes «{NB}Question{NB}» et «{NB}Réponse{NB}».\n\n"
        f"Disposition attendue (ligne 1 d'une feuille){NB}:  Catégorie | Points | Question | Réponse\n"
        "Fichier ▸ Enregistrer un modèle vierge… donne une feuille toute prête.",
    ),
    "e.no_topics": (
        "The Question/Answer sheet has no topics and no questions under its header.",
        "La feuille Question/Réponse n'a ni catégorie ni question sous son en-tête.",
    ),

    # ---- the blank template --------------------------------------------------------
    "template.topic": ("Topic {n}", "Catégorie {n}"),
    "template.board_sheet": ("Board", "Tableau"),
    "template.howto_sheet": ("How to use", "Mode d'emploi"),
}

TEMPLATE_HEADER = {
    "en": ["Topic", "Points", "Question", "Answer"],
    "fr": ["Catégorie", "Points", "Question", "Réponse"],
}

HOW_TO = {
    "en": [
        "How to fill in the Board sheet",
        "",
        "Each row is one question: its topic, its points (1 to 6), the question and the answer.",
        "The six topics become the six board columns, in the order they first appear.",
        "Points say which row of the board the question sits in: 1 at the top, 6 at the bottom.",
        "You may write a topic once and leave the Topic cell blank below it; a blank topic carries down.",
        "A blank Points cell takes the next free row for that topic.",
        "Alt+Enter starts a new line inside a cell; line breaks show on the board.",
        "This sheet is ignored by Quizboard — only sheets with Question and Answer headers are read.",
    ],
    "fr": [
        "Comment remplir la feuille Tableau",
        "",
        f"Chaque ligne correspond à une question{NB}: sa catégorie, ses points (de 1 à 6), la question et la réponse.",
        "Les six catégories deviennent les six colonnes du tableau, dans l'ordre où elles apparaissent.",
        f"Les points indiquent la rangée du tableau où se trouve la question{NB}: 1 en haut, 6 en bas.",
        "Vous pouvez écrire une catégorie une seule fois et laisser vides les cellules Catégorie en dessous; "
        "une catégorie vide reprend celle du dessus.",
        "Une cellule Points vide prend la prochaine rangée libre de sa catégorie.",
        "Alt+Entrée insère un saut de ligne dans une cellule; les sauts de ligne s'affichent sur le tableau.",
        f"Quizboard ignore cette feuille{NB}: seules les feuilles qui ont les en-têtes Question et Réponse sont lues.",
    ],
}


class Message(str):
    """A warning or error: its English text as a plain str, renderable in French too."""

    key: str
    params: dict

    def __new__(cls, key: str, **params) -> Message:
        message = super().__new__(cls, tr(key, "en", **params))
        message.key, message.params = key, params
        return message

    def render(self, language: str) -> str:
        return tr(self.key, language, **self.params)


class Quoted(tuple):
    """Names listed in quotes: 'A', 'B' in English, « A », « B » in French."""


def render(value, language: str) -> str:
    """A Message or Quoted list in `language`; anything else as it is."""
    if isinstance(value, Message):
        return value.render(language)
    if isinstance(value, Quoted):
        quote = (lambda name: f"«{NB}{name}{NB}»") if language == "fr" else (lambda name: f"'{name}'")
        return ", ".join(quote(name) for name in value)
    return str(value)


def tr(key: str, language: str, **params) -> str:
    template = TEXT[key][1 if language == "fr" else 0]
    return template.format(**{name: render(value, language) for name, value in params.items()})


def tr_n(key: str, n: int, language: str, **params) -> str:
    """Singular or plural by `n`: English singular for 1 only, French for 0 and 1."""
    singular = n in (0, 1) if language == "fr" else n == 1
    return tr(f"{key}.{'one' if singular else 'many'}", language, n=n, **params)


def other(language: str) -> str:
    return "en" if language == "fr" else "fr"


def and_list(names: list[str], language: str) -> str:
    if len(names) == 1:
        return names[0]
    return ", ".join(names[:-1]) + tr("and", language) + names[-1]
