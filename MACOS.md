# Quizboard on macOS

The same code runs on macOS, Windows and Linux. Quizboard needs **Python 3.10 or newer**
(macOS's own `python3` is too old):

    brew install python@3.13        # Homebrew: https://brew.sh
    # or the installer from https://www.python.org/downloads/macos/

Then, in Terminal:

    git clone https://github.com/zklaas/quizboard.git Quizboard
    cd Quizboard
    bash install-macos.sh

That creates `.venv` (PyQt6), a `quizboard` command in `~/.local/bin`, and a
**Quizboard** app in your Applications folder (Launchpad, Spotlight, the Dock).

Settings are kept in `~/Library/Application Support/Quizboard`. *Edit in Excel* opens the workbook in whatever program macOS uses for .xlsx files (Excel or Numbers).

To remove it: delete the folder, `~/Applications/Quizboard.app` and `~/.local/bin/quizboard`.

macOS support has not yet been tested on a Mac; please report anything that does
not work.
