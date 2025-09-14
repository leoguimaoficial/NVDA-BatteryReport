# -*- coding: UTF-8 -*-

def _(arg):
    return arg

addon_info = {
    "addon_name": "NVDABatteryReport",
    "addon_summary": _("NVDA BatteryReport"),
    "addon_description": _("Generate accessible Windows battery reports (health, capacity, usage, and life estimates) directly in NVDA. Friendly interface, integrated history, and multi-language support."),
    "addon_version": "1.0",
    "addon_author": "Leo Guima",
    "addon_url": "https://github.com/leoguimaoficial/NVDA-BatteryReport",
    "addon_sourceURL": "https://github.com/leoguimaoficial/NVDA-BatteryReport",
    "addon_docFileName": "readme.html",
    "addon_minimumNVDAVersion": "2021.1",
    "addon_lastTestedNVDAVersion": "2025.2",
    "addon_updateChannel": None,
    "addon_license": "GPL v3",
    "addon_licenseURL": "https://www.gnu.org/licenses/gpl-3.0.html",
}

pythonSources = [
    "addon/*.py",
    "addon/globalPlugins/batteryReport.py"
]

i18nSources = pythonSources + ["buildVars.py"]

excludedFiles = []

baseLanguage = "en"

markdownExtensions = []

brailleTables = {}
symbolDictionaries = {}
