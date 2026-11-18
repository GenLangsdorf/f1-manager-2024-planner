# F1 Manager 2024 Planner

Reads your **F1 Manager 2024** career save and builds an Excel workbook that tells you what to do next: what to design and build, who to sign, which facilities to fix, how to stay under the cost cap, and when to take a grid penalty for worn power unit parts. It only reads your saves and never changes them.

## What you get

- **Action plan:** a ranked to-do list with timings, costs and a running total against your spending limit
- **Car vs grid:** every car area compared with the grid median, plus what each remaining circuit needs
- **Designs, stock, spares and manufacturing:** build days, with rushed and emergency costs
- **Staff, facilities, scouting, pit crew and drivers:** contracts, mentality and next-season deals
- **Board:** your objective, and how to meet it without raising next season's target
- **Power unit:** wear against your allocation and when to take new parts
- **Reserve driver plan** and an optional all-female team project
- Works with several careers: it asks which save to use

## Install

1. Install [Python 3.8 or later](https://www.python.org/downloads/). On the first installer screen tick **Add python.exe to PATH**.
2. Install the Excel library: open a command window and run `pip install openpyxl` (or `py -m pip install openpyxl`).
3. Download this repository (**Code > Download ZIP**) and unzip it into its own folder.

You need Windows, F1 Manager 2024 and Excel (LibreOffice Calc also opens the workbook).

## Use

Save your game, then double-click **Update planner.py**. It asks which save to use (if you have more than one), builds `F1 Manager Planner - <your team>.xlsx` in the same folder and opens it. Close the workbook in Excel before updating.

Double-click **Watch saves.py** and leave the window open to rebuild the workbook automatically every time the game saves.

Command-line options (`--newest`, `--list`, `--save-dir`, `--workbook` and others) are listed by `python f1m_update.py --help`.

## Good to know

- The facts come straight from your save; the advice is a set of rules applied to them. Treat it as a well-informed starting point, not the game's own answer.
- Tested on one career with a custom team, game version 1.11. Other versions or save layouts may need an update. Please open an issue if something looks wrong.
- Saves are copied to a temporary folder and read there. Nothing is written back to the game's folders, and nothing is sent over the internet. Back up your saves before first use anyway.

## Disclaimer

Fan-made and unofficial. Not affiliated with, endorsed by or supported by Codemasters, Frontier Developments, Formula 1 or the FIA. F1 Manager 2024 is a trademark of its owners. Use at your own risk.

Licensed under the [MIT licence](LICENSE).
