"""
Affiliate Team P/L Dashboard — Daily Refresh Script
Runs via GitHub Actions at 1pm IST (07:30 UTC) every day.

Steps:
  1. Fetch Google Sheets gviz export (public, no auth needed)
  2. Parse current-month rows for each vertical
  3. Calculate daily deltas vs yesterday's snapshot
  4. Update snapshots.json (rolling 30-day history)
  5. Write data.json (consumed by index.html live-data loader)
"""

import json
import sys
import urllib.request
from datetime import date, timedelta
from html.parser import HTMLParser

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
SHEET_ID   = "1TEhR-l5S_-u9Z2J7q7VeBWjmIKvg0mGVe1mdYvI6Dus"
SHEET_TABS = {"education": "Education", "autos": "Autos", "summary": "Sheet1"}

TODAY     = date.today().isoformat()          # e.g. "2026-04-15"
YESTERDAY = (date.today() - timedelta(days=1)).isoformat()
MONTH_NUM = date.today().month
MONTH_LBL = date.today().strftime("%B %Y")    # e.g. "April 2026"
PERIOD    = date.today().strftime("%-d %b")   # e.g. "15 Apr"


# ─────────────────────────────────────────────
# HTML TABLE PARSER
# ─────────────────────────────────────────────
class TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell, self._in_cell = [], [], "", False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th"):
            self._in_cell = True
            self._cell = ""

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self._row.append(self._cell.strip())
            self._in_cell = False
        elif tag == "tr" and self._row:
            self.rows.append(self._row)

    def handle_data(self, data):
        if self._in_cell:
            self._cell += data


def fetch_and_parse(sheet_name):
    url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        f"/gviz/tq?tqx=out:html&sheet={sheet_name}"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            html = resp.read().decode("utf-8")
    except Exception as e:
        print(f"  WARNING: Could not fetch sheet '{sheet_name}': {e}")
        return [], []

    parser = TableParser()
    parser.feed(html)
    if len(parser.rows) < 2:
        print(f"  WARNING: Sheet '{sheet_name}' returned no data rows.")
        return [], []

    headers = [h.strip() for h in parser.rows[0]]
    return headers, parser.rows[1:]


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def to_num(val):
    """Parse '$1,234.56', '(500)', '-$200' etc. → float."""
    if not val or val.strip() in ("", "-", "N/A", "n/a"):
        return 0.0
    v = val.replace("$", "").replace(",", "").replace("%", "").strip()
    if v.startswith("(") and v.endswith(")"):
        v = "-" + v[1:-1]
    try:
        return float(v)
    except ValueError:
        return 0.0


MONTH_ABBRS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

def is_current_month(cell):
    """Return True if the Month cell refers to the current calendar month."""
    c = cell.lower()
    for abbr, num in MONTH_ABBRS.items():
        if abbr in c:
            return num == MONTH_NUM
    return False


def row_to_dict(headers, row):
    return {h: (row[i] if i < len(row) else "") for i, h in enumerate(headers)}


def col(d, *keys):
    """Return first matching key value from a row dict (case-insensitive fallback)."""
    for k in keys:
        if k in d:
            return d[k]
    kl = {key.lower(): val for key, val in d.items()}
    for k in keys:
        if k.lower() in kl:
            return kl[k.lower()]
    return ""


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
print(f"\n=== Affiliate P/L Dashboard Refresh — {TODAY} ===\n")

# ── 1. Load existing snapshots ────────────────
with open("snapshots.json") as f:
    snap_store = json.load(f)

history = snap_store.get("snapshots", [])
yesterday_snap = next((s for s in history if s["date"] == YESTERDAY), None)
print(f"Yesterday snapshot: {'found' if yesterday_snap else 'NOT found — daily deltas will be N/A'}")


# ── 2. Fetch sheets ───────────────────────────
print("\nFetching Google Sheets data...")

edu_hdr,     edu_rows     = fetch_and_parse(SHEET_TABS["education"])
autos_hdr,   autos_rows   = fetch_and_parse(SHEET_TABS["autos"])
summary_hdr, summary_rows = fetch_and_parse(SHEET_TABS["summary"])

print(f"  Education : {len(edu_rows)} rows")
print(f"  Autos     : {len(autos_rows)} rows")
print(f"  Summary   : {len(summary_rows)} rows")


# ── 3. Parse rows into structured advertiser lists ────
def parse_advertisers(headers, rows, sub_label=""):
    """Return list of advertiser dicts for current month rows."""
    advs = []
    for row in rows:
        d = row_to_dict(headers, row)
        month_cell = col(d, "Month", "month")
        if not is_current_month(month_cell):
            continue
        name = col(d, "Advertiser", "advertiser")
        if not name:
            continue
        advs.append({
            "name":       name,
            "sub":        sub_label or col(d, "Vertical", "vertical"),
            "book":       col(d, "Supply/Demand Book", "Book", "book"),
            "vertical":   col(d, "Vertical", "vertical"),
            "payout":     col(d, "Payout Type", "payout"),
            "target":     to_num(col(d, "P&L Target", "pl target")),
            "netRev":     to_num(col(d, "Net Revenue", "Gross Revenue", "net revenue", "gross revenue")),
            "tac":        to_num(col(d, "TAC", "tac")),
            "plActual":   to_num(col(d, "P&L Actual", "pl actual")),
            "plPacing":   to_num(col(d, "P&L EOM Pacing", "eom pacing", "pacing")),
            "rag":        col(d, "RAG Status", "rag", "status"),
            "blocker":    col(d, "Top Blocker", "blocker"),
            "action":     col(d, "Next Action", "action"),
        })
    return advs


edu_advs     = parse_advertisers(edu_hdr, edu_rows)
autos_advs   = parse_advertisers(autos_hdr, autos_rows)
summary_advs = parse_advertisers(summary_hdr, summary_rows)

print(f"\n  Current-month advertisers parsed:")
print(f"    Education : {len(edu_advs)}")
print(f"    Autos     : {len(autos_advs)}")
print(f"    Summary   : {len(summary_advs)}")


# ── 4. Aggregate by vertical ──────────────────
def agg(advs, *vertical_keywords):
    """Sum net revenue, TAC, P&L for rows whose vertical contains any keyword."""
    matched = [a for a in advs
               if any(kw.lower() in a["vertical"].lower() for kw in vertical_keywords)]
    return {
        "netRevenue": round(sum(a["netRev"]   for a in matched)),
        "tac":        round(sum(a["tac"]       for a in matched)),
        "plActual":   round(sum(a["plActual"]  for a in matched)),
    }


# Education & Autos: always use their dedicated tabs — these have Net Revenue (5% AAC-adjusted).
# Sheet1 has Gross Revenue (pre-AAC) for those verticals, so we exclude them from summary_advs.
# PTP, Others, Demand Book have no dedicated tabs — Sheet1 is correct for those.
summary_other = [
    a for a in summary_advs
    if not any(kw in a["vertical"].lower() for kw in ("edu", "education", "autos"))
]
all_advs = edu_advs + autos_advs + summary_other

today_verticals = {
    "edu_leadgen":   agg(all_advs, "Edu: LeadGen", "edu leadgen", "education leadgen"),
    "edu_clicks":    agg(all_advs, "Edu: Clicks",  "edu clicks",  "education clicks"),
    "autos_clicks":  agg(all_advs, "Autos: Clicks", "autos clicks"),
    "autos_leadgen": agg(all_advs, "Autos: LeadGen","autos leadgen"),
    "ptp":           agg(all_advs, "PTP"),
    "others":        agg(all_advs, "Others"),
    "demand":        agg(all_advs, "Ad.Tech", "Demand", "demand"),
}

# Rollup groups
supply_keys = ["edu_leadgen", "edu_clicks", "autos_clicks", "autos_leadgen", "ptp", "others"]
demand_keys = ["demand"]

def rollup(keys):
    return {
        "netRevenue": sum(today_verticals[k]["netRevenue"] for k in keys),
        "tac":        sum(today_verticals[k]["tac"]        for k in keys),
        "plActual":   sum(today_verticals[k]["plActual"]   for k in keys),
    }

supply_totals = rollup(supply_keys)
demand_totals = rollup(demand_keys)
grand_totals  = {
    "netRevenue": supply_totals["netRevenue"] + demand_totals["netRevenue"],
    "tac":        supply_totals["tac"]        + demand_totals["tac"],
    "plActual":   supply_totals["plActual"]   + demand_totals["plActual"],
}

print(f"\n  Aggregated totals:")
print(f"    Net Revenue : ${grand_totals['netRevenue']:>10,.0f}")
print(f"    TAC         : ${grand_totals['tac']:>10,.0f}")
print(f"    P&L         : ${grand_totals['plActual']:>10,.0f}")


# ── 5. Calculate daily deltas ─────────────────
deltas = {}
if yesterday_snap:
    yv = {v["name"]: v for v in yesterday_snap.get("verticals", [])}

    NAME_MAP = {
        "edu_leadgen":   "Edu: LeadGen",
        "edu_clicks":    "Edu: Clicks",
        "autos_clicks":  "Autos: Clicks",
        "autos_leadgen": "Autos: LeadGen",
        "ptp":           "PTP",
        "others":        "Others",
        "demand":        "Ad.Tech",
    }
    for key, snap_name in NAME_MAP.items():
        y = yv.get(snap_name, {"netRev": 0, "tac": 0, "pl": 0})
        t = today_verticals[key]
        deltas[key] = {
            "netRevenue": t["netRevenue"] - y.get("netRev", 0),
            "tac":        t["tac"]        - y.get("tac",    0),
            "plActual":   t["plActual"]   - y.get("pl",     0),
        }
    print(f"\n  Daily deltas calculated vs {YESTERDAY}.")
else:
    print("\n  No daily deltas — first snapshot or gap in history.")


# ── 6. Update snapshots.json ──────────────────
new_snap = {
    "date":       TODAY,
    "capturedAt": f"{TODAY}T13:00:00+05:30",
    "dataSource": "github_actions_gviz",
    "overview": {
        "netRev":      grand_totals["netRevenue"],
        "tac":         grand_totals["tac"],
        "pl":          grand_totals["plActual"],
    },
    "verticals": [
        {"book": "Supply", "name": "Edu: LeadGen",    **{k: today_verticals["edu_leadgen"][k]   for k in ("netRevenue","tac","plActual")}},
        {"book": "Supply", "name": "Edu: Clicks",     **{k: today_verticals["edu_clicks"][k]    for k in ("netRevenue","tac","plActual")}},
        {"book": "Supply", "name": "Autos: Clicks",   **{k: today_verticals["autos_clicks"][k]  for k in ("netRevenue","tac","plActual")}},
        {"book": "Supply", "name": "Autos: LeadGen",  **{k: today_verticals["autos_leadgen"][k] for k in ("netRevenue","tac","plActual")}},
        {"book": "Supply", "name": "PTP",             **{k: today_verticals["ptp"][k]           for k in ("netRevenue","tac","plActual")}},
        {"book": "Supply", "name": "Others",          **{k: today_verticals["others"][k]        for k in ("netRevenue","tac","plActual")}},
        {"book": "Demand", "name": "Ad.Tech",         **{k: today_verticals["demand"][k]        for k in ("netRevenue","tac","plActual")}},
    ],
}
# Re-key fields to match existing snapshot schema (netRev/tac/pl)
for v in new_snap["verticals"]:
    v["netRev"] = v.pop("netRevenue")
    v["pl"]     = v.pop("plActual")

# Remove today's entry if exists (idempotent), append, keep last 30
history = [s for s in history if s["date"] != TODAY]
history.append(new_snap)
history = sorted(history, key=lambda s: s["date"])[-30:]

snap_store["snapshots"] = history
snap_store["lastUpdated"] = TODAY

with open("snapshots.json", "w") as f:
    json.dump(snap_store, f, indent=2)
print("\n  snapshots.json updated.")


# ── 7. Build data.json ────────────────────────
def fmt_date(d):
    return date.fromisoformat(d).strftime("%-d %b %Y") if d else ""


VERTICAL_META = {
    "edu_leadgen":   {"book": "Supply", "name": "Edu: LeadGen",   "tag": "vt-edu"},
    "edu_clicks":    {"book": "Supply", "name": "Edu: Clicks",    "tag": "vt-edu"},
    "autos_clicks":  {"book": "Supply", "name": "Autos: Clicks",  "tag": "vt-autos"},
    "autos_leadgen": {"book": "Supply", "name": "Autos: LeadGen", "tag": "vt-autos"},
    "ptp":           {"book": "Supply", "name": "PTP",            "tag": "vt-ptp"},
    "others":        {"book": "Supply", "name": "Others",         "tag": "vt-other"},
    "demand":        {"book": "Demand", "name": "Ad.Tech",        "tag": "vt-demand"},
}


def build_daily_verticals():
    if not deltas:
        return []
    rows = []
    for key, meta in VERTICAL_META.items():
        if key not in deltas:
            continue
        d = deltas[key]
        pl = d["plActual"]
        rows.append({
            **meta,
            "netRev":   d["netRevenue"],
            "tac":      d["tac"],
            "pl":       pl,
            "pacing":   None,
            "target":   None,
            "variance": f"+${pl:,.0f}" if pl >= 0 else f"-${abs(pl):,.0f}",
            "focus":    "Up"   if pl >= 0 else "Down",
            "focusCls": "focus-low" if pl >= 0 else "focus-high",
        })
    return rows


def build_daily_overview():
    if not deltas:
        return None
    d_supply = [deltas[k] for k in supply_keys if k in deltas]
    d_demand = [deltas[k] for k in demand_keys if k in deltas]
    s_rev = sum(x["netRevenue"] for x in d_supply)
    d_rev = sum(x["netRevenue"] for x in d_demand)
    s_pl  = sum(x["plActual"]   for x in d_supply)
    d_pl  = sum(x["plActual"]   for x in d_demand)
    t_rev = s_rev + d_rev
    t_tac = sum(x["tac"] for x in deltas.values())
    t_pl  = s_pl + d_pl
    return {
        "netRev":    round(t_rev),
        "tac":       round(t_tac),
        "pl":        round(t_pl),
        "supplyRev": round(s_rev),
        "demandRev": round(d_rev),
        "supplyPL":  round(s_pl),
        "demandPL":  round(d_pl),
        "pacing":    None,
        "target":    None,
        "variance":  None,
    }


def build_edu_advertisers():
    result = []
    for a in edu_advs:
        sub = "LeadGen" if "leadgen" in a["vertical"].lower() else "Clicks"
        result.append({
            "sub":     sub,
            "name":    a["name"],
            "target":  round(a["target"]),
            "netRev":  round(a["netRev"]),
            "tac":     round(a["tac"]),
            "pl":      round(a["plActual"]),
            "pacing":  round(a["plPacing"]) if a["plPacing"] else None,
            "rag":     a["rag"],
            "blocker": a["blocker"],
            "action":  a["action"],
        })
    return result


def build_autos_advertisers():
    result = []
    for a in autos_advs:
        sub = "Clicks" if "clicks" in a["vertical"].lower() else "LeadGen"
        result.append({
            "sub":     sub,
            "name":    a["name"],
            "target":  round(a["target"]),
            "netRev":  round(a["netRev"]),
            "tac":     round(a["tac"]),
            "pl":      round(a["plActual"]),
            "pacing":  round(a["plPacing"]) if a["plPacing"] else None,
            "rag":     a["rag"],
            "blocker": a["blocker"],
        })
    return result


data_json = {
    "lastUpdated": f"Auto-refreshed {date.today().strftime('%-d %b %Y')} 1:00pm IST",
    "refreshDate": TODAY,
    "mtd": {
        "overview": {
            "netRev":    grand_totals["netRevenue"],
            "tac":       grand_totals["tac"],
            "pl":        grand_totals["plActual"],
            "supplyRev": supply_totals["netRevenue"],
            "demandRev": demand_totals["netRevenue"],
            "supplyPL":  supply_totals["plActual"],
            "demandPL":  demand_totals["plActual"],
        },
        "eduAdvertisers":   build_edu_advertisers(),
        "autosAdvertisers": build_autos_advertisers(),
    },
    "daily": {
        "label":    f"Daily Change vs {fmt_date(YESTERDAY)}" if deltas else "Daily Change (no prior snapshot)",
        "period":   "Daily",
        "overview": build_daily_overview(),
        "verticals": build_daily_verticals(),
        "_note":    f"Delta = {TODAY} MTD minus {YESTERDAY} MTD. Source: GitHub Actions gviz refresh.",
    } if deltas else None,
}

with open("data.json", "w") as f:
    json.dump(data_json, f, indent=2)

print("  data.json written.")
print(f"\n=== Refresh complete — P&L: ${grand_totals['plActual']:,.0f} | Net Rev: ${grand_totals['netRevenue']:,.0f} ===\n")
