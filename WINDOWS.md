# Quizboard on Windows

The same code runs on Windows, macOS and Linux. Quizboard needs **Python 3.10 or newer**:

    winget install --id Python.Python.3.13 -e --scope user

Then install with PowerShell:

    git clone https://github.com/zklaas/quizboard.git Quizboard
    cd Quizboard
    powershell -ExecutionPolicy Bypass -File install.ps1

That creates `.venv` (PyQt6), a `quizboard` command on your PATH (open a
new terminal first), and a **Quizboard** Start-menu entry.

Settings are kept in `%APPDATA%\Quizboard`. *Edit in Excel* opens the workbook in whatever program Windows uses for .xlsx files.

To remove it: delete the folder, the Start-menu entry (Quizboard) and
`%LOCALAPPDATA%\Programs\Quizboard`.

Windows support has not yet been tested on a Windows machine; please report anything
that does not work.
