import os
import re
import json
import threading
import subprocess
import locale
from datetime import datetime
from html import unescape

import wx
import gui
import ui
import addonHandler
import globalPluginHandler
from scriptHandler import script

addonHandler.initTranslation()

ADDON_DIR = os.path.dirname(__file__)
REPORTS_DIR = os.path.join(ADDON_DIR, "battery_reports")
HISTORY_FILE = os.path.join(ADDON_DIR, "battery_history.json")
DIALOG_TITLE = _("NVDA Battery Report")
EMPTY_HISTORY_MSG = _("No battery reports found.")
os.makedirs(REPORTS_DIR, exist_ok=True)


def _load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _powercfg_path():
    win = os.environ.get("SystemRoot", r"C:\\Windows")
    sysnative = os.path.join(win, "Sysnative", "powercfg.exe")
    system32 = os.path.join(win, "System32", "powercfg.exe")
    return sysnative if os.path.isfile(sysnative) else system32

POWERCFG = _powercfg_path()


def generate_battery_report():
    if not os.path.isfile(POWERCFG):
        raise FileNotFoundError(_("powercfg.exe not found."))
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(REPORTS_DIR, f"battery_report_{ts}.html")
    cmd = [POWERCFG, "/batteryreport", "/output", out_path]
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=False)
    _, stderr = p.communicate()
    if p.returncode != 0:
        raise RuntimeError(stderr.strip() or _("Failed to run powercfg."))
    if not os.path.isfile(out_path):
        raise RuntimeError(_("Battery report file was not created."))
    try:
        with open(out_path, "r", encoding="utf-8", errors="replace") as f:
            html = f.read()
    except UnicodeDecodeError:
        with open(out_path, "r", encoding="cp1252", errors="replace") as f:
            html = f.read()
    return out_path, html


def _collapse(s):
    return re.sub(r"\s+", " ", unescape(s), flags=re.S).strip()


def _find_table_by_header(html, header_text):
    pat = rf"<h2[^>]*>\s*{re.escape(header_text)}\s*</h2>\s*(?:<(?:div|canvas)[^>]*>.*?</(?:div|canvas)>\s*)*<table[^>]*>(.*?)</table>"
    m = re.search(pat, html, flags=re.I | re.S)
    return m.group(1) if m else ""


def _text_in_td_after(label, html):
    pat = r"<t[dh][^>]*>\s*(?:<span[^>]*>)?\s*" + re.escape(label) + r"\s*(?:</span>)?\s*</t[dh]>\s*<t[dh][^>]*>(.*?)</t[dh]>"
    m = re.search(pat, html, flags=re.I | re.S)
    if not m:
        return ""
    val = re.sub(r"<[^>]+>", "", m.group(1), flags=re.S)
    return _collapse(val)


def _table_rows(html_table_inner):
    rows = []
    for row in re.findall(r"<tr[^>]*>(.*?)</tr>", html_table_inner, flags=re.I | re.S):
        cells = [_collapse(re.sub(r"<[^>]+>", "", c, flags=re.S)) for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", row, flags=re.I | re.S)]
        if cells:
            rows.append(cells)
    return rows


def _to_mWh(s):
    if not s:
        return None
    digits = re.findall(r"\d", s)
    return int("".join(digits)) if digits else None


try:
    locale.setlocale(locale.LC_TIME, "")
except Exception:
    pass

_DT_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$")
_PERIOD_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})$")


def _parse_dt(s):
    m = _DT_RE.match(s)
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    hh, mm, ss = m.group(4), m.group(5), m.group(6)
    if hh is None:
        return datetime(y, mo, d)
    return datetime(y, mo, d, int(hh), int(mm), int(ss or 0))


def _fmt_date_local(d):
    try:
        return d.strftime("%x")
    except Exception:
        return d.strftime("%d/%m/%Y")


def _fmt_dt_local(d):
    try:
        return f"{d.strftime('%x')} {d.strftime('%X')}".strip()
    except Exception:
        return d.strftime("%d/%m/%Y %H:%M:%S")


def _localize_cell(label, value):
    if not value:
        return value
    lab = (label or "").strip().upper()
    if lab in {"START TIME", "TIME", "DATE", "START"}:
        dt = _parse_dt(value)
        return _fmt_dt_local(dt) if dt else value
    if lab == "PERIOD":
        m = _PERIOD_RE.match(value)
        if m:
            d1 = _parse_dt(m.group(1)); d2 = _parse_dt(m.group(2))
            s1 = _fmt_date_local(d1) if d1 else m.group(1)
            s2 = _fmt_date_local(d2) if d2 else m.group(2)
            return f"{s1} - {s2}"
        dt = _parse_dt(value)
        return _fmt_date_local(dt) if dt else value
    dt = _parse_dt(value)
    return _fmt_dt_local(dt) if dt else value


def parse_battery_report(html):
    if not html:
        return {}
    raw = _collapse(html)
    m_head = re.search(r"<h1[^>]*>\s*Battery report\s*</h1>\s*<table[^>]*>(.*?)</table>", raw, flags=re.I | re.S)
    head_tbl = m_head.group(1) if m_head else ""
    header = {
        "Computer name": _text_in_td_after("COMPUTER NAME", head_tbl),
        "System product name": _text_in_td_after("SYSTEM PRODUCT NAME", head_tbl),
        "BIOS": _text_in_td_after("BIOS", head_tbl),
        "OS build": _text_in_td_after("OS BUILD", head_tbl),
        "Platform role": _text_in_td_after("PLATFORM ROLE", head_tbl),
        "Connected standby": _text_in_td_after("CONNECTED STANDBY", head_tbl),
        "Report time": _text_in_td_after("REPORT TIME", head_tbl),
    }
    inst_tbl = _find_table_by_header(raw, "Installed batteries") or ""
    installed = {
        "Name": _text_in_td_after("NAME", inst_tbl),
        "Manufacturer": _text_in_td_after("MANUFACTURER", inst_tbl),
        "Serial number": _text_in_td_after("SERIAL NUMBER", inst_tbl),
        "Chemistry": _text_in_td_after("CHEMISTRY", inst_tbl),
        "Design capacity": _text_in_td_after("DESIGN CAPACITY", inst_tbl),
        "Full charge capacity": _text_in_td_after("FULL CHARGE CAPACITY", inst_tbl),
        "Cycle count": _text_in_td_after("CYCLE COUNT", inst_tbl),
    }
    design_mWh = _to_mWh(installed.get("Design capacity")) if installed.get("Design capacity") else None
    full_mWh = _to_mWh(installed.get("Full charge capacity")) if installed.get("Full charge capacity") else None
    health_pct = round((full_mWh / float(design_mWh)) * 100.0, 2) if (design_mWh and full_mWh) else None

    recent_tbl = _find_table_by_header(raw, "Recent usage")
    usage_tbl = _find_table_by_header(raw, "Battery usage")
    hist_tbl = _find_table_by_header(raw, "Usage history")
    cap_tbl = _find_table_by_header(raw, "Battery capacity history")
    life_tbl = _find_table_by_header(raw, "Battery life estimates")

    return {
        "header": header,
        "installed": installed,
        "design_mWh": design_mWh,
        "full_mWh": full_mWh,
        "health_pct": health_pct,
        "recent_usage": _table_rows(recent_tbl),
        "battery_usage": _table_rows(usage_tbl),
        "usage_history": _table_rows(hist_tbl),
        "capacity_history": _table_rows(cap_tbl),
        "life_estimates": _table_rows(life_tbl),
    }


def format_summary(info):
    ts = _fmt_dt_local(datetime.now())
    hp = info.get("health_pct"); dm = info.get("design_mWh"); fm = info.get("full_mWh")
    if hp is not None and dm and fm:
        return _("{ts} — Health {hp}% ({full:,}/{des:,} mWh)").format(ts=ts, hp=hp, full=fm, des=dm)
    return _("{ts} — Battery report").format(ts=ts)


def _parse_hms_to_secs(s):
    if not s:
        return None
    m = re.search(r"(\d+):(\d{2}):(\d{2})", s)
    if not m:
        return None
    h, mnt, sec = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return h * 3600 + mnt * 60 + sec


def _secs_to_hms(secs):
    h = secs // 3600
    m = (secs % 3600) // 60
    s = secs % 60
    return f"{h}:{m:02d}:{s:02d}"


def _is_all_nulls(row):
    return all(re.fullmatch(r"[-–—\s]*", (c or "")) for c in row)


def _upper_set(lst):
    return {(c or "").strip().upper() for c in lst}


def _build_items_from_table(rows, expected_headers, label_map=None, date_key_kind=None):
    items = []
    header_idx = None
    for i, r in enumerate(rows):
        if expected_headers.issubset(_upper_set(r)):
            header_idx = i
            break
    if header_idx is None:
        return items
    headers = rows[header_idx]
    for r in rows[header_idx + 1:]:
        if expected_headers.issubset(_upper_set(r)):
            break
        if _is_all_nulls(r):
            continue
        pairs = []
        key_dt = None
        for j, h in enumerate(headers):
            if j >= len(r):
                break
            raw_val = r[j]
            lab = (label_map.get(h.strip().upper(), h) if label_map else h)
            val = _localize_cell(h, raw_val)
            pairs.append(f"{lab}: {val}" if val else f"{lab}:")
            if key_dt is None:
                if date_key_kind == "start":
                    dt = _parse_dt(raw_val)
                    if dt:
                        key_dt = dt
                elif date_key_kind == "period":
                    m = _PERIOD_RE.match(raw_val)
                    if m:
                        dt = _parse_dt(m.group(2))
                        if dt:
                            key_dt = dt
                    else:
                        dt = _parse_dt(raw_val)
                        if dt:
                            key_dt = dt
        items.append((key_dt, " | ".join(pairs)))
    return items


class DetailsDialog(wx.Dialog):
    SECTIONS = (
        ("overview", _("Overview")),
        ("installed", _("Installed battery")),
        ("recent", _("Recent usage (last 7 days)")),
        ("battery_usage", _("Battery usage (last 7 days)")),
        ("capacity_history", _("Capacity history")),
        ("usage_history", _("Usage history")),
        ("life_estimates", _("Battery life estimates")),
    )

    def __init__(self, parent, info):
        super().__init__(parent, title=_("Battery report details"), size=(1020, 700))
        self.info = info
        pnl = wx.Panel(self)
        self.sectionLabel = wx.StaticText(pnl, label=_("&Section:"))
        self.section = wx.Choice(pnl, choices=[label for key, label in self.SECTIONS])
        self.section.SetSelection(0)
        self.rowsLabel = wx.StaticText(pnl, label=_("&Rows:"))
        self.rowsChoice = wx.Choice(pnl)
        self.orderLabel = wx.StaticText(pnl, label=_("&Order:"))
        self.orderChoice = wx.Choice(pnl, choices=[_("Newest first"), _("Oldest first")])
        self.orderChoice.SetSelection(0)
        self.listLabel = wx.StaticText(pnl, label=_("&Items:"))
        self.list = wx.ListBox(pnl, name=_("Items list"))
        self.descLabel = wx.StaticText(pnl, label=_("&Description:"))
        self.desc = wx.TextCtrl(pnl, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_DONTWRAP | wx.HSCROLL | wx.BORDER_SUNKEN)
        self.desc.SetValue(_("Select an item to read its details."))
        self.btn_copy = wx.Button(pnl, label=_("&Copy selected"))
        self.btn_open_raw = wx.Button(pnl, label=_("&Open raw HTML"))
        self.btn_close = wx.Button(pnl, id=wx.ID_CANCEL, label=_("&Close"))

        gridTop = wx.FlexGridSizer(2, 8, 5, 5)
        gridTop.Add(self.sectionLabel, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.TOP, 10)
        gridTop.Add(self.section, 0, wx.EXPAND | wx.TOP, 8)
        gridTop.Add(self.rowsLabel, 0, wx.ALIGN_CENTER_VERTICAL | wx.TOP, 10)
        gridTop.Add(self.rowsChoice, 0, wx.TOP, 8)
        gridTop.Add(self.orderLabel, 0, wx.ALIGN_CENTER_VERTICAL | wx.TOP, 10)
        gridTop.Add(self.orderChoice, 0, wx.TOP, 8)
        gridTop.Add(self.listLabel, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        gridTop.Add((1, 1))

        lhs = wx.BoxSizer(wx.VERTICAL)
        lhs.Add(gridTop, 0, wx.EXPAND | wx.RIGHT, 10)
        lhs.Add(self.list, 1, wx.ALL | wx.EXPAND, 10)

        rhs = wx.BoxSizer(wx.VERTICAL)
        rhs.Add(self.descLabel, 0, wx.TOP | wx.RIGHT | wx.LEFT, 10)
        rhs.Add(self.desc, 1, wx.ALL | wx.EXPAND, 10)

        hs = wx.BoxSizer(wx.HORIZONTAL)
        hs.Add(lhs, 1, wx.EXPAND)
        hs.Add(rhs, 1, wx.EXPAND)

        bs = wx.BoxSizer(wx.HORIZONTAL)
        bs.Add(self.btn_copy, 0, wx.LEFT | wx.BOTTOM, 10)
        bs.Add(self.btn_open_raw, 0, wx.LEFT | wx.BOTTOM, 10)
        bs.AddStretchSpacer()
        bs.Add(self.btn_close, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        root = wx.BoxSizer(wx.VERTICAL)
        root.Add(hs, 1, wx.EXPAND)
        root.Add(bs, 0, wx.EXPAND)
        pnl.SetSizer(root)

        self._sections, self._section_lengths, self._prefix_counts, self._legends = self._build_sections(info)
        self._apply_section("overview")

        self.section.Bind(wx.EVT_CHOICE, self._on_section_changed)
        self.rowsChoice.Bind(wx.EVT_CHOICE, self._on_rows_order)
        self.orderChoice.Bind(wx.EVT_CHOICE, self._on_rows_order)
        self.list.Bind(wx.EVT_LISTBOX, self._on_select)
        self.btn_copy.Bind(wx.EVT_BUTTON, self._copy_selected)
        self.btn_open_raw.Bind(wx.EVT_BUTTON, lambda e: self._open_latest_html())
        self.btn_close.Bind(wx.EVT_BUTTON, lambda e: self.EndModal(wx.ID_CANCEL))
        self.Bind(wx.EVT_CHAR_HOOK, self._on_key)

        self.section.MoveAfterInTabOrder(self.sectionLabel)
        self.rowsChoice.MoveAfterInTabOrder(self.section)
        self.orderChoice.MoveAfterInTabOrder(self.rowsChoice)
        self.list.MoveAfterInTabOrder(self.orderChoice)
        self.desc.MoveAfterInTabOrder(self.list)
        self.btn_copy.MoveAfterInTabOrder(self.desc)
        self.btn_open_raw.MoveAfterInTabOrder(self.btn_copy)
        self.btn_close.MoveAfterInTabOrder(self.btn_open_raw)
        self.section.SetFocus()

    def _build_sections(self, info):
        def add(items, label, value, desc):
            if value is None or value == "":
                return
            line = f"{label}: {value}"
            items.append((None, line, _("{label}: {value}\n\n{desc}").format(label=label, value=value, desc=desc)))

        sections = {}
        lengths = {}
        prefixes = {}
        legends = {
            "recent": _("Columns: Start time | State | Source | Remaining"),
            "battery_usage": _("Columns: Start time | State | Duration | Energy drained"),
            "capacity_history": _("Columns: Period | Full charge capacity | Design capacity"),
            "usage_history": _("Columns: Period | Active | Connected standby"),
            "life_estimates": _("Battery life estimates\nBattery life estimates based on observed drains\nColumns: Period | At full charge — Active, Connected standby | At design capacity — Active, Connected standby"),
        }

        items = []
        h = info.get("header", {})
        add(items, _("Computer name"), h.get("Computer name"), _("Computer name is the Windows name of this device."))
        add(items, _("System product name"), h.get("System product name"), _("Model reported by the system firmware (BIOS/UEFI)."))
        add(items, _("BIOS"), h.get("BIOS"), _("Firmware version and date."))
        add(items, _("OS build"), h.get("OS build"), _("Windows build installed on this system."))
        add(items, _("Platform role"), h.get("Platform role"), _("Device role, e.g., Mobile or Desktop."))
        add(items, _("Connected standby"), h.get("Connected standby"), _("Whether modern standby is supported."))
        add(items, _("Report time"), _localize_cell("START TIME", h.get("Report time")), _("Timestamp when this report was generated."))
        dm = info.get("design_mWh"); fm = info.get("full_mWh"); hp = info.get("health_pct")
        if hp is not None and dm and fm:
            add(items, _("Battery health"), f"{hp} %", _("Battery health = Full charge capacity / Design capacity."))
            add(items, _("Design capacity (mWh)"), f"{dm:,}", _("Factory-specified maximum energy in milliwatt-hours."))
            add(items, _("Full charge capacity (mWh)"), f"{fm:,}", _("Current maximum energy (mWh) after wear."))
        sections["overview"] = items
        lengths["overview"] = len(items)
        prefixes["overview"] = 0

        items = []
        inst = info.get("installed", {})
        add(items, _("Battery name"), inst.get("Name"), _("Identifier for the installed battery."))
        add(items, _("Manufacturer"), inst.get("Manufacturer"), _("Battery vendor reported by firmware."))
        add(items, _("Serial number"), inst.get("Serial number"), _("Battery serial number."))
        add(items, _("Chemistry"), inst.get("Chemistry"), _("Battery chemistry code as reported by the system."))
        add(items, _("Design capacity"), inst.get("Design capacity"), _("Factory-specified maximum energy (mWh)."))
        add(items, _("Full charge capacity"), inst.get("Full charge capacity"), _("Current maximum energy (mWh) after wear."))
        cc = inst.get("Cycle count")
        if cc and cc not in ("-", "—"):
            add(items, _("Cycle count"), cc, _("Number of full charge–discharge cycles recorded."))
        sections["installed"] = items
        lengths["installed"] = len(items)
        prefixes["installed"] = 0

        rec_items = _build_items_from_table(
            info.get("recent_usage", []),
            {"START TIME", "STATE", "SOURCE", "CAPACITY REMAINING"},
            label_map={"START TIME": _("Start time"), "STATE": _("State"), "SOURCE": _("Source"), "CAPACITY REMAINING": _("Remaining")},
            date_key_kind="start",
        )
        bat_items = _build_items_from_table(
            info.get("battery_usage", []),
            {"START TIME", "STATE", "DURATION", "ENERGY DRAINED"},
            label_map={"START TIME": _("Start time"), "STATE": _("State"), "DURATION": _("Duration"), "ENERGY DRAINED": _("Energy drained")},
            date_key_kind="start",
        )
        cap_items = _build_items_from_table(
            info.get("capacity_history", []),
            {"PERIOD", "FULL CHARGE CAPACITY", "DESIGN CAPACITY"},
            label_map={"PERIOD": _("Period"), "FULL CHARGE CAPACITY": _("Full charge capacity"), "DESIGN CAPACITY": _("Design capacity")},
            date_key_kind="period",
        )
        use_items = _build_items_from_table(
            info.get("usage_history", []),
            {"PERIOD", "ACTIVE", "CONNECTED STANDBY"},
            label_map={"PERIOD": _("Period"), "ACTIVE": _("Active"), "CONNECTED STANDBY": _("Connected standby")},
            date_key_kind="period",
        )

        life_rows = info.get("life_estimates", [])
        life_items = []
        for r in life_rows:
            if _is_all_nulls(r):
                continue
            if {"PERIOD", "ACTIVE", "CONNECTED STANDBY"}.issubset(_upper_set(r)):
                continue
            if len(r) >= 6:
                period = _localize_cell("PERIOD", r[0])
                tP = _("Period"); tFC = _("At full charge"); tDC = _("At design capacity"); tA = _("Active"); tCS = _("Connected standby")
                line = f"{tP}: {period} | {tFC} — {tA}: {r[1]}, {tCS}: {r[2]} | {tDC} — {tA}: {r[4]}, {tCS}: {r[5]}"
                key_dt = None
                m = _PERIOD_RE.match(r[0])
                if m:
                    dt = _parse_dt(m.group(2))
                    if dt:
                        key_dt = dt
                else:
                    dt = _parse_dt(r[0])
                    if dt:
                        key_dt = dt
                life_items.append((key_dt, line))
        fc_act = []; fc_cs = []; dc_act = []; dc_cs = []
        for r in life_rows:
            if len(r) >= 6 and not _is_all_nulls(r) and not {"PERIOD", "ACTIVE", "CONNECTED STANDBY"}.issubset(_upper_set(r)):
                s1 = _parse_hms_to_secs(r[1]); s2 = _parse_hms_to_secs(r[2]); s3 = _parse_hms_to_secs(r[4]); s4 = _parse_hms_to_secs(r[5])
                if s1 is not None: fc_act.append(s1)
                if s2 is not None: fc_cs.append(s2)
                if s3 is not None: dc_act.append(s3)
                if s4 is not None: dc_cs.append(s4)
        def avg(lst):
            return _secs_to_hms(sum(lst)//len(lst)) if lst else _("-")
        if life_items:
            tAVG = _("Average"); tFC = _("At full charge"); tDC = _("At design capacity"); tA = _("Active"); tCS = _("Connected standby")
            avg_line = f"{tAVG} | {tFC} — {tA}: {avg(fc_act)}, {tCS}: {avg(fc_cs)} | {tDC} — {tA}: {avg(dc_act)}, {tCS}: {avg(dc_cs)}"
            life_items = [(None, avg_line)] + life_items
            prefixes["life_estimates"] = 1
        else:
            prefixes["life_estimates"] = 0

        def finalize(items, empty_msg, legend_key):
            if not items:
                line = empty_msg
                return [(None, line, f"{line}\n\n{legends.get(legend_key, '')}")]
            items_sorted = sorted(items, key=lambda x: (x[0] or datetime.min), reverse=True)
            return [(k, line, f"{line}\n\n{legends.get(legend_key, '')}") for (k, line) in items_sorted]

        sections["recent"] = finalize(rec_items, _("No entries for the last 7 days."), "recent")
        sections["battery_usage"] = finalize(bat_items, _("No entries for the last 7 days."), "battery_usage")
        sections["capacity_history"] = finalize(cap_items, _("No data."), "capacity_history")
        sections["usage_history"] = finalize(use_items, _("No data."), "usage_history")
        sections["life_estimates"] = [(k, s, f"{s}\n\n{legends['life_estimates']}") for (k, s) in life_items]

        for key in ("recent", "battery_usage", "capacity_history", "usage_history", "life_estimates"):
            lengths[key] = max(0, len(sections[key]) - prefixes.get(key, 0))
            prefixes.setdefault(key, 0)

        return sections, lengths, prefixes, legends

    def _apply_section(self, key):
        self._toggle_rows_controls(key)
        self._populate_rows_choice(key)
        self._refresh_list(key)

    def _toggle_rows_controls(self, key):
        table_like = key in {"usage_history", "capacity_history", "life_estimates"}
        self.rowsLabel.Enable(table_like)
        self.rowsChoice.Enable(table_like)
        self.orderLabel.Enable(table_like)
        self.orderChoice.Enable(table_like)

    def _populate_rows_choice(self, key):
        self.rowsChoice.Clear()
        n = self._section_lengths.get(key, 0)
        if key not in {"usage_history", "capacity_history", "life_estimates"}:
            return
        maxOpt = max(10, ((n + 9)//10)*10)
        opts = [str(x) for x in range(10, maxOpt + 1, 10)]
        for o in opts:
            self.rowsChoice.Append(o)
        default_val = min(30, int(opts[-1])) if opts else 10
        try:
            self.rowsChoice.SetStringSelection(str(default_val))
        except Exception:
            if opts:
                self.rowsChoice.SetSelection(len(opts)-1)

    def _get_current_key(self):
        idx = self.section.GetSelection()
        return self.SECTIONS[idx][0]

    def _refresh_list(self, key=None):
        key = key or self._get_current_key()
        items = list(self._sections.get(key, []))
        if key in {"usage_history", "capacity_history", "life_estimates"}:
            take = None
            try:
                take = int(self.rowsChoice.GetStringSelection())
            except Exception:
                pass
            if self.orderChoice.GetSelection() == 1:
                prefix = self._prefix_counts.get(key, 0)
                if prefix:
                    head = items[:prefix]
                    rest = list(reversed(items[prefix:]))
                    items = head + rest
                else:
                    items = list(reversed(items))
            prefix = self._prefix_counts.get(key, 0)
            base = items[prefix:]
            n = len(base)
            if take is None:
                take = n
            kept = base[:min(take, n)]
            items = (items[:prefix] if prefix else []) + kept
        self._current_items = items
        self.list.Clear()
        if items:
            self.list.InsertItems([ln for _, ln, _ in items], 0)
            self.list.SetSelection(0)
            self._update_desc_from_selection()
        else:
            self.desc.SetValue(_("No data for this section."))
            self.desc.SetInsertionPoint(0)

    def _on_section_changed(self, evt):
        self._apply_section(self._get_current_key())

    def _on_rows_order(self, evt):
        self._refresh_list()

    def _on_select(self, evt):
        self._update_desc_from_selection()

    def _update_desc_from_selection(self):
        idx = self.list.GetSelection()
        if idx == wx.NOT_FOUND:
            return
        items = self._current_items
        desc = items[idx][2] if idx < len(items) else ""
        self.desc.SetValue(desc or _("No description available."))
        self.desc.SetInsertionPoint(0)

    def _copy_selected(self, evt):
        idx = self.list.GetSelection()
        if idx == wx.NOT_FOUND:
            wx.CallLater(120, lambda: ui.message(_("Please select an item to copy.")))
            return
        text = self._current_items[idx][1]
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(text))
            wx.TheClipboard.Close()
            wx.CallLater(120, lambda: ui.message(_("Copied to clipboard.")))
        else:
            wx.CallLater(120, lambda: ui.message(_("Failed to open clipboard.")))

    def _open_latest_html(self):
        try:
            files = [os.path.join(REPORTS_DIR, f) for f in os.listdir(REPORTS_DIR) if f.lower().endswith('.html')]
            if not files:
                return
            latest = max(files, key=os.path.getmtime)
            import webbrowser
            webbrowser.open(latest)
        except Exception:
            pass

    def _on_key(self, event):
        if event.GetKeyCode() == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        event.Skip()


class BatteryReportDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title=DIALOG_TITLE, size=(640, 620))
        self.worker = None
        self.items = _load_json(HISTORY_FILE, []) or []
        pnl = wx.Panel(self)
        self.info = wx.StaticText(pnl, label=_("Click Generate to create a Windows battery report using powercfg."))
        self.btn_generate = wx.Button(pnl, label=_("&Generate report"))
        self.lst = wx.ListBox(pnl, name=_("Battery reports history"))
        self.btn_view = wx.Button(pnl, label=_("&View details"))
        self.btn_delete = wx.Button(pnl, label=_("&Delete"))
        self.btn_clear = wx.Button(pnl, label=_("&Clear history"))
        btn_close = wx.Button(pnl, id=wx.ID_CLOSE, label=_("&Close"))
        v = wx.BoxSizer(wx.VERTICAL)
        v.Add(self.info, 0, wx.ALL | wx.EXPAND, 10)
        v.Add(self.btn_generate, 0, wx.ALL | wx.EXPAND, 10)
        v.Add(wx.StaticText(pnl, label=_("History:")), 0, wx.LEFT, 10)
        v.Add(self.lst, 1, wx.ALL | wx.EXPAND, 10)
        v.Add(self.btn_view, 0, wx.LEFT | wx.EXPAND, 10)
        v.Add(self.btn_delete, 0, wx.LEFT | wx.EXPAND, 10)
        v.Add(self.btn_clear, 0, wx.LEFT | wx.EXPAND, 10)
        v.Add(btn_close, 0, wx.ALL | wx.ALIGN_RIGHT, 10)
        pnl.SetSizer(v)
        self.btn_generate.Bind(wx.EVT_BUTTON, self._on_generate)
        self.btn_view.Bind(wx.EVT_BUTTON, self._on_view)
        self.btn_delete.Bind(wx.EVT_BUTTON, self._on_delete)
        self.btn_clear.Bind(wx.EVT_BUTTON, self._on_clear)
        btn_close.Bind(wx.EVT_BUTTON, self._on_close)
        self.lst.Bind(wx.EVT_LISTBOX, lambda e: self._update_buttons())
        self.Bind(wx.EVT_CHAR_HOOK, self._on_key)
        self.Bind(wx.EVT_CLOSE, self._on_close)
        if not self.items:
            self.lst.Append(EMPTY_HISTORY_MSG)
        else:
            for it in self.items:
                self.lst.Append(it.get("summary", ""))
        self._update_buttons()

    def _update_buttons(self):
        empty = (self.lst.GetCount() == 1 and self.lst.GetString(0) == EMPTY_HISTORY_MSG)
        if empty:
            self.btn_view.Enable(False)
            self.btn_delete.Enable(False)
            self.btn_clear.Enable(False)
        else:
            has_sel = self.lst.GetSelection() != wx.NOT_FOUND
            self.btn_view.Enable(has_sel)
            self.btn_delete.Enable(has_sel)
            self.btn_clear.Enable(bool(self.items))

    def _on_key(self, evt):
        if evt.GetKeyCode() == wx.WXK_ESCAPE:
            self._on_close(evt)
            return
        evt.Skip()

    def _on_generate(self, evt):
        if self.worker and self.worker.is_alive():
            return
        self.btn_generate.Enable(False)
        self.info.SetLabel(_("Generating report... Please wait."))
        self.worker = threading.Thread(target=self._worker_thread, daemon=True)
        self.worker.start()

    def _worker_thread(self):
        try:
            path, html = generate_battery_report()
            info = parse_battery_report(html)
            wx.CallAfter(self._finish, path, info)
        except Exception as e:
            wx.CallAfter(self._error, str(e))

    def _finish(self, path, info):
        self.btn_generate.Enable(True)
        summary = format_summary(info)
        if self.lst.GetCount() == 1 and self.lst.GetString(0) == EMPTY_HISTORY_MSG:
            self.lst.Delete(0)
        self.lst.InsertItems([summary], 0)
        self.items.insert(0, {"summary": summary, "path": path, "info": info})
        _save_json(HISTORY_FILE, self.items[:100])
        self._update_buttons()
        self.info.SetLabel(summary.replace("\n", "  "))
        hp = info.get("health_pct"); dm = info.get("design_mWh"); fm = info.get("full_mWh")
        if hp is not None and dm and fm:
            msg = _("Report generated. Battery health: {hp}% ({full:,}/{des:,} mWh)").format(hp=hp, full=fm, des=dm)
        else:
            msg = _("Report generated.")
        wx.MessageBox(msg, _("Battery report"), style=wx.OK | wx.ICON_INFORMATION)

    def _error(self, msg):
        self.btn_generate.Enable(True)
        self.info.SetLabel(_("Error: {m}").format(m=msg))

    def _on_view(self, evt):
        sel = self.lst.GetSelection()
        if sel == wx.NOT_FOUND:
            return
        item = self.items[sel]
        dlg = DetailsDialog(self, item.get("info", {}))
        dlg.ShowModal(); dlg.Destroy()

    def _on_delete(self, evt):
        sel = self.lst.GetSelection()
        if sel == wx.NOT_FOUND:
            return
        dlg = wx.MessageDialog(self, _("Are you sure you want to delete this report?"), _("Confirm delete"), style=wx.YES_NO | wx.ICON_WARNING)
        if dlg.ShowModal() == wx.ID_YES:
            try:
                path = self.items[sel].get("path")
                if path and os.path.isfile(path):
                    os.remove(path)
            except Exception:
                pass
            self.lst.Delete(sel)
            del self.items[sel]
            _save_json(HISTORY_FILE, self.items)
            if not self.items:
                self.lst.Append(EMPTY_HISTORY_MSG)
            self._update_buttons()
        dlg.Destroy()

    def _on_clear(self, evt):
        if not self.items:
            return
        dlg = wx.MessageDialog(self, _("Clear all reports and delete files?"), _("Clear history"), style=wx.YES_NO | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            for it in self.items:
                try:
                    p = it.get("path")
                    if p and os.path.isfile(p):
                        os.remove(p)
                except Exception:
                    pass
            self.items.clear()
            _save_json(HISTORY_FILE, [])
            self.lst.Clear()
            self.lst.Append(EMPTY_HISTORY_MSG)
            self._update_buttons()
        dlg.Destroy()

    def _on_close(self, evt=None):
        if self.worker and self.worker.is_alive():
            self.worker.join(timeout=2)
        self.Destroy()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("NVDA Battery Report")

    def __init__(self):
        super().__init__()
        self._toolsMenuId = wx.NewId()
        gui.mainFrame.sysTrayIcon.toolsMenu.Append(self._toolsMenuId, _("NVDA Battery Report"))
        gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self.on_tools_menu, id=self._toolsMenuId)

    @script(description=_("Opens the NVDA BatteryReport dialog."), category=_("NVDA Battery Report"))
    def script_showUI(self, gesture):
        self._launch_dialog()

    def on_tools_menu(self, event):
        self._launch_dialog()

    def _launch_dialog(self):
        dlg = BatteryReportDialog(wx.GetApp().GetTopWindow())
        dlg.Show(); dlg.Raise(); wx.CallAfter(dlg.btn_generate.SetFocus)

    # no default gesture; user sets it in Input gestures
