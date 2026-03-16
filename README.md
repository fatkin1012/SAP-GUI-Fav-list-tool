# SAP GUI Tcode to Favorites Automation

This small project automates adding many SAP transaction codes (T-codes) to your Favorites folder and renaming them quickly.

Motivation

Since there are way too many Tcode needed to add to the Favourite folder, And doing it manually and renaming them is so time consuming. So I make an easy script to automize this process

What it does

- Reads a list of T-codes from [tcodes.txt](tcodes.txt)
- Activates your SAP GUI window using a title keyword
- For each T-code, it opens the add entry dialog, inserts the T-code, then renames the new Favorite entry with a short prefix (e.g. "<Tcode> - ")
- Automates keyboard interactions so you can process many entries without manual repetition

Files

- [main.ps1](main.ps1) — PowerShell script performing the automation (window detection + SendKeys sequences)
- [run.bat](run.bat) — Launcher that prompts for the SAP window title keyword, then runs `main.ps1`
- [tcodes.txt](tcodes.txt) — Input file: one T-code per line

Requirements

- Windows
- SAP GUI open and logged in
- PowerShell available (ExecutionPolicy may need adjustment; `run.bat` uses `-ExecutionPolicy Bypass`)
- Focusable SAP Easy Access screen

Usage

1. Open SAP and go to SAP Easy Access.
2. Expand `Favorites` and click/select the target folder where new entries should be added.
3. Edit [tcodes.txt](tcodes.txt) with one T-code per line.
4. Run the launcher (`run.bat`) and enter a part of your SAP window title when prompted (example: `WAP` or `WAP(1)/400`).

From a terminal you can run:

```bat
run.bat
```

Or call the PowerShell script directly:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\main.ps1 -WindowTitleKeyword "WAP"
```

Important notes & precautions

- Do NOT use the keyboard or mouse while the script runs — it sends keystrokes to the SAP window.
- Keep the SAP window visible and the correct Favorites folder selected before starting.
- If no window matches the keyword, the script prints visible window titles to help you choose a better keyword.
- The default window keyword is `WAP` if you press Enter without input in `run.bat`.

Input format

- `tcodes.txt` should contain one transaction code per line. Blank lines are ignored.

Troubleshooting

- No T-codes processed: ensure `tcodes.txt` exists and contains non-empty lines.
- Script cannot find SAP window: run `run.bat` and enter a more specific window title keyword.
- Wrong target folder: select the intended Favorites folder in SAP before starting the script.

If you want, I can also:

- Add a short GIF or screenshots showing the script running.
- Add a safety checklist or a dry-run mode that logs actions instead of sending keystrokes.

---

Created by the project author — use responsibly and test with a small list first.
