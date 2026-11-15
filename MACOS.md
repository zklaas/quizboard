# Quizboard on macOS

The same code runs on macOS, Windows and Linux.

## Where to type the commands

Everything below is typed into **Terminal**: open Applications > Utilities >
Terminal, or press ⌘-Space, type *Terminal* and press Return. Press Return after
each line.

## 1. Install Python

Quizboard needs **Python 3.10 or newer** (the `python3` that comes with macOS is
too old). The simplest way is the installer from
https://www.python.org/downloads/macos/ ; or, with Homebrew (https://brew.sh):

    brew install python@3.13

## 2. Get Quizboard

**With git:**

    git clone https://github.com/zklaas/quizboard.git Quizboard
    cd Quizboard

(The first time, macOS offers to install its command-line tools for `git`: accept,
then run the line again.)

**Without git:** on https://github.com/zklaas/quizboard click **Code > Download ZIP**;
the Mac unpacks it into your Downloads folder (`quizboard-main`). In Terminal type
`cd ` (with a space), drag that folder onto the Terminal window, and press Return.

## 3. Install it

    bash install-macos.sh

That creates `.venv` (PyQt6), a `quizboard` command in `~/.local/bin`, and a
**Quizboard** app in your Applications folder: start it from Launchpad, Spotlight or
the Dock from then on.

Settings are kept in `~/Library/Application Support/Quizboard`. *Edit in Excel* opens the workbook in whatever program macOS uses for .xlsx files (Excel or Numbers).

To remove it: delete the folder, `~/Applications/Quizboard.app` and `~/.local/bin/quizboard`.

If something doesn't work, please write to zklaas@yahoo.ca.
