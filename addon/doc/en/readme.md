\# NVDA Battery Report



\*\*NVDA Battery Report\*\* is an \*\*open-source\*\* NVDA add-on that generates and reads the official \*\*Windows battery report\*\* (`powercfg /batteryreport`) in a fully accessible way. It speaks a quick summary (battery health and capacities), organizes data into clear sections, and keeps a local history.



---



\## Features



\* \*\*One-click Windows report\*\*, generated locally via `powercfg`.

\* \*\*Spoken summary\*\* at the end (battery health and capacities).

\* \*\*Clear sections\*\*:



&nbsp; \* Overview

&nbsp; \* Installed battery

&nbsp; \* Recent usage (7 days)

&nbsp; \* \*\*Battery drain (7 days)\*\* — start, state, duration, and \*\*energy drained\*\*

&nbsp; \* Capacity history

&nbsp; \* Usage history

&nbsp; \* Battery life estimates (with averages)

\* \*\*Screen-reader-friendly lists\*\*: each table row becomes plain text with a \*\*column legend\*\*.

\* \*\*Sorting\*\* (newest/oldest) and \*\*row limits\*\* (10, 20, 30…).

\* \*\*Copy selected\*\* line to the clipboard.

\* \*\*Open original HTML\*\* report for verification.

\* \*\*Multi-language\*\* (.po/.mo).



---



\## Changelog



\### 1.0 (latest)



\* Official Windows report generation.

\* Sections with legends, sorting, and row limits.

\* Health calculation using \*\*Design capacity (factory)\*\* vs \*\*Maximum charge capacity (current)\*\*.

\* \*\*Battery drain (7 days)\*\* with clear “no records” messaging.

\* Battery life estimates with averages.

\* Integrated history, copy to clipboard, and open original HTML.



---



\## Installation



1\. \*\*Install via NVDA’s Add-ons menu\*\* (Add-on Store or local file).

&nbsp;  NVDA → \*\*Tools → Add-ons\*\* → choose \*\*Get Add-ons\*\* (Store) \*or\* \*\*Install…\*\* and select the `.nvda-addon` file.

2\. Restart NVDA when prompted.



---



\## How to Use



1\. Open the add-on: \*\*Tools → NVDA Battery Report\*\*.

&nbsp;  \*(No default gesture — see “Shortcuts” to set one.)\*

2\. Click \*\*Generate report\*\*. When it finishes, NVDA announces a summary and the item appears in \*\*History\*\*.

3\. Select a report in \*\*History\*\* and press \*\*View details\*\*.

4\. In \*\*Details\*\*:



&nbsp;  \* Pick a \*\*Section\*\*.

&nbsp;  \* For large tables, adjust \*\*Rows\*\* and \*\*Order\*\* (Newest/Oldest).

&nbsp;  \* The \*\*list\*\* shows each row; the \*\*Description\*\* repeats it as \*\*plain text\*\* with the \*\*column legend\*\*.

&nbsp;  \* Use \*\*Copy selected\*\* to place the line on the clipboard.

&nbsp;  \* \*\*Open raw HTML\*\* to view the original Windows report.



---



\## Shortcuts (assign in NVDA)



1\. NVDA → \*\*Preferences → Input gestures\*\*

2\. Search for \*\*“NVDA Battery Report”\*\*

3\. Bind your preferred gesture (e.g., `NVDA+Shift+B`)



---



\## Key Concepts



\* \*\*Design capacity (factory):\*\* energy the battery supported \*\*when new\*\*.

\* \*\*Maximum charge capacity (current):\*\* energy it supports \*\*today\*\*, after wear.

\* \*\*Battery health\*\* = Maximum charge capacity ÷ Design capacity × 100%.



\*\*Battery drain (7 days)\*\* lists when the device \*\*consumed energy on battery\*\*: the \*\*start time\*\*, \*\*state\*\*, \*\*duration\*\*, and the \*\*energy drained\*\* in that interval.



---



\## Where Files Are Stored



\* HTML reports: `…\\addons\\NVDABatteryReport\\globalPlugins\\battery\_reports\\`

\* History (JSON): `…\\addons\\NVDABatteryReport\\globalPlugins\\battery\_history.json`

&nbsp; \*(Inside the user’s NVDA profile.)\*



---



\## Translating \& Contributing



\* Translation files live under `addon/locale/` (`.po/.mo`).

\* PRs for fixes, features, or translations are welcome!



\### Build from Source (SCons)



```bash

git clone https://github.com/leoguimaoficial/NVDA-BatteryReport.git

cd NVDA-BatteryReport

pip install scons markdown

SCons   # or: python -m SCons

```



This produces the `.nvda-addon` package and compiles translations.



---



\## FAQ



\* \*\*Do I need internet?\*\* No. The report is generated locally (`powercfg`).

\* \*\*No battery in my desktop—does it work?\*\* The report may say no battery and some sections will be empty.

\* \*\*Why do some fields show “–”?\*\* Windows didn’t record enough data for that period.

\* \*\*Order differs from the HTML file.\*\* The add-on sorts by \*\*actual date\*\* (newest first). You can switch it in \*\*Order\*\*.

\* \*\*Dates look odd.\*\* They follow your \*\*Windows regional format\*\*.



---



\## Troubleshooting



\* \*\*“powercfg.exe not found.”\*\* Check Windows; `powercfg.exe` resides in `System32/Sysnative`.

\* \*\*Error opening raw HTML.\*\* Ensure the file exists in the reports folder; generate a new report if needed.

\* \*\*Empty/“unknown” items.\*\* Generate a fresh report; sections without data show “No entries for the last 7 days.”





