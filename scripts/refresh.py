"""
Affiliate Team P/L Dashboard — Daily Refresh Script
Runs via GitHub Actions at 1pm IST (07:30 UTC) every day.

Steps:
  1. Fetch Google Sheets gviz export from ALL 5 dedicated tabs (public, no auth needed)
  2. Parse current-month rows for each vertical
  3. Calculate daily deltas vs yesterday's snapshot
  4. Update snapshots.json (rolling 30-day history)
  5. Write data.json (consumed by index.html live-data loader)

Tab → Revenue column mapping (confirmed from sheet inspection):
  Education              → "Net Revenue"   (5% AAC-adjusted)  entity col: "Advertiser"
  Autos                  → "Net Revenue"   (5% AAC-adjusted)  entity col: "Advertiser"
  PTP                    → "Net Revenue"   (5% AAC-adjusted)  entity col: "Publisher"  ← different!
  Others/ Performance    → "Net Revenue"   (5% AAC-adjusted)  entity col: "Advertiser"
  Demand Book            → "Gross Revenue" (correct for demand) entity col: "Advertiser"

Sheet1 is NOT used for any vertical — it contains Gross Revenue for all rows
and would inflate Education/Autos/PTP/Others figures.
"""

import json
import urllib.parse
import urllib.request
from datetime import date, timedelta
from html.parser import HTMLParser

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
SHEET_ID = "1TEhR-l5S_-u9Z2J7q7VeBWjmIKvg0mGVe1mdYvI6Dus"

# Exact tab names as they appear in the Google Sheet
SHEET_TABS = {
    "education": "Education",
    "autos":     "Autos",
    "ptp":       "PTP",
    "others":    "Others/ Performance Offers",   # note: space before "Performance"
    "demand":    "Demand Book",
}

TODAY     = date.today().isoformat()
YESTERDAY = (date.today() - timedelta(days=1)).isoformat()
MONTH_NUM = date.today().month


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


def fetch_and_parse(tab_name):
    """Fetch gviz HTML export for a sheet tab and return (headers, data_rows)."""
    encoded = urllib.parse.quote(tab_name, safe="")
    url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        f"/gviz/tq?tqx=out:html&sheet={encoded}"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            html = resp.read().decode("utf-8")
    except Exception as e:
        print(f"  WARNING: Could not fetch tab '{tab_name}': {e}")
        return [], []

    parser = TableParser()
    parser.feed(html)
    if len(parser.rows) < 2:
        print(f"  WARNING: Tab '{tab_name}' returned no data rows.")
        return [], []

    headers = [h.strip() for h in parser.rows[0]]
    return headers, parser.rows[1:]


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def to_num(val):
    """Parse '$1,234.56', '(500)', '-$200', '' → float."""
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
    c = cell.lower()
    for abbr, num in MONTH_ABBRS.items():
        if abbr in c:
            return num == MONTH_NUM
    return False


def row_to_dict(headers, row):
    return {h: (row[i] if i < len(row) else "") for i, h in enumerate(headers)}


def col(d, *keys):
    """Return first matching value from dict — exact match first, then lowercase."""
    for k in keys:
        if k in d:
            return d[k]
    dl = {k.lower(): v for k, v in d.items()}
    for k in keys:
        if k.lower() in dl:
            return dl[k.lower()]
    return ""


# ─────────────────────────────────────────────
# PARSE ADVERTISERS / PUBLISHERS
# ─────────────────────────────────────────────
def parse_rows(headers, rows):
    """
    Parse current-month rows into structured dicts.
    Handles both 'Advertiser' (Education/Autos/Others/Demand) and
    'Publisher' (PTP) as the entity name column.
    Revenue column: tries 'Net Revenue' first, falls back to 'Gross Revenue'
    (Demand Book uses Gross Revenue — correct for that vertical).
    """
    result = []
    for row in rows:
        d = row_to_dict(headers, row)

        month_cell = col(d, "Month", "month")
        if not is_current_month(month_cell):
            continue

        # Entity name — Advertiser OR Publisher
        name = col(d, "Advertiser", "Publisher", "advertiser", "publisher")
        if not name or name.upper() == "TOTAL":
            continue

        result.append({
            "name":     name,
            "vertical": col(d, "Vertical", "vertical"),
            "book":     col(d, "Supply/Demand Book", "Book", "book"),
            "payout":   col(d, "Payout Type", "payout"),
            "target":   to_num(col(d, "P&L Target",      "pl target")),
            "netRev":   to_num(col(d, "Net Revenue",  "Gross Revenue",
                                       "net revenue", "gross revenue")),
            "tac":      to_num(col(d, "TAC", "tac")),
            "plActual": to_num(col(d, "P&L Actual",       "pl actual")),
            "plPacing": to_num(col(d, "P&L EOM Pacing", "eom pacing", "pacing")),
            "rag":      col(d, "RAG Status", "RAG", "rag", "status"),
            "blocker":  col(d, "Top Blocker", "blocker"),
            "action":   col(d, "Next Action", "action"),
        })
    return result


def agg(rows, *vertical_keywords):
    """Sum netRev, TAC, P&L for rows whose vertical contains any of the keywords."""
    matched = [r for r in rows
               if any(kw.lower() in r["vertical"].lower() for kw in vertical_keywords)]
    return {
        "netRevenue": round(sum(r["netRev"]   for r in matched)),
        "tac":        round(sum(r["tac"]       for r in matched)),
        "plActual":   round(sum(r["plActual"]  for r in matched)),
    }


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
print(f"\n=== Affiliate P/L Dashboard Refresh — {TODAY} ===\n")

# ── 1. Load existing snapshots ────────────────
with open("snapshots.json") as f:
    snap_store = json.load(f)

history = snap_store.get("snapshots", [])
yesterday_snap = next((s for s in history if s["date"] == YESTERDAY), None)
print(f"Yesterday snapshot : {'found (' + YESTERDAY + ')' if yesterday_snap else 'NOT found — daily deltas will be N/A'}")


# ── 2. Fetch all dedicated tabs ───────────────
print("\nFetching Google Sheets tabs...")
edu_hdr,    edu_rows    = fetch_and_parse(SHEET_TABS["education"])
autos_hdr,  autos_rows  = fetch_and_parse(SHEET_TABS["autos"])
ptp_hdr,    ptp_rows    = fetch_and_parse(SHEET_TABS["ptp"])
others_hdr, others_rows = fetch_and_parse(SHEET_TABS["others"])
demand_hdr, demand_rows = fetch_and_parse(SHEET_TABS["demand"])

print(f"  Education              : {len(edu_rows)} rows")
print(f"  Autos                  : {len(autos_rows)} rows")
print(f"  PTP                    : {len(ptp_rows)} rows")
print(f"  Others/Perf. Offers    : {len(others_rows)} rows")
print(f"  Demand Book            : {len(demand_rows)} rows")


# ── 3. Parse into structured row lists ────────
edu_rows_p    = parse_rows(edu_hdr,    edu_rows)
autos_rows_p  = parse_rows(autos_hdr,  autos_rows)
ptp_rows_p    = parse_rows(ptp_hdr,    ptp_rows)
others_rows_p = parse_rows(others_hdr, others_rows)
demand_rows_p = parse_rows(demand_hdr, demand_rows)

print(f"\n  Current-month rows parsed:")
print(f"    Education           : {len(edu_rows_p)}")
print(f"    Autos               : {len(autos_rows_p)}")
print(f"    PTP                 : {len(ptp_rows_p)}")
print(f"    Others              : {len(others_rows_p)}")
print(f"    Demand Book         : {len(demand_rows_p)}")

# All rows — each vertical sourced exclusively from its own tab
all_rows = edu_rows_p + autos_rows_p + ptp_rows_p + others_rows_p + demand_rows_p


# ── 4. Aggregate by vertical ──────────────────
today_verticals = {
    "edu_leadgen":   agg(edu_rows_p,    "LeadGen"),
    "edu_clicks":    agg(edu_rows_p,    "Clicks"),
    "autos_clicks":  agg(autos_rows_p,  "Clicks"),
    "autos_leadgen": agg(autos_rows_p,  "LeadGen"),
    "ptp":           agg(ptp_rows_p,    "PTP"),
    "others":        agg(others_rows_p, "Others", "Performance"),
    "demand":        agg(demand_rows_p, "Ad.Tech", "Demand"),
}

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

print(f"\n  Aggregated totals (all from dedicated tabs, 5% AAC-adjusted):")
print(f"    Net Revenue : ${grand_totals['netRevenue']:>12,.0f}")
print(f"    TAC         : ${grand_totals['tac']:>12,.0f}")
print(f"    P&L         : ${grand_totals['plActual']:>12,.0f}")
print(f"\n  By vertical:")
for k, v in today_verticals.items():
    print(f"    {k:<18}: P&L ${v['plActual']:>10,.0f}  |  Rev ${v['netRevenue']:>10,.0f}  |  TAC ${v['tac']:>10,.0f}")


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
    "dataSource": "github_actions_gviz_dedicated_tabs",
    "overview": {
        "netRev": grand_totals["netRevenue"],
        "tac":    grand_totals["tac"],
        "pl":     grand_totals["plActual"],
    },
    "verticals": [
        {"book": "Supply", "name": "Edu: LeadGen",   "netRev": today_verticals["edu_leadgen"]["netRevenue"],   "tac": today_verticals["edu_leadgen"]["tac"],   "pl": today_verticals["edu_leadgen"]["plActual"]},
        {"book": "Supply", "name": "Edu: Clicks",    "netRev": today_verticals["edu_clicks"]["netRevenue"],    "tac": today_verticals["edu_clicks"]["tac"],    "pl": today_verticals["edu_clicks"]["plActual"]},
        {"book": "Supply", "name": "Autos: Clicks",  "netRev": today_verticals["autos_clicks"]["netRevenue"],  "tac": today_verticals["autos_clicks"]["tac"],  "pl": today_verticals["autos_clicks"]["plActual"]},
        {"book": "Supply", "name": "Autos: LeadGen", "netRev": today_verticals["autos_leadgen"]["netRevenue"], "tac": today_verticals["autos_leadgen"]["tac"], "pl": today_verticals["autos_leadgen"]["plActual"]},
        {"book": "Supply", "name": "PTP",            "netRev": today_verticals["ptp"]["netRevenue"],           "tac": today_verticals["ptp"]["tac"],           "pl": today_verticals["ptp"]["plActual"]},
        {"book": "Supply", "name": "Others",         "netRev": today_verticals["others"]["netRevenue"],        "tac": today_verticals["others"]["tac"],        "pl": today_verticals["others"]["plActual"]},
        {"book": "Demand", "name": "Ad.Tech",        "netRev": today_verticals["demand"]["netRevenue"],        "tac": today_verticals["demand"]["tac"],        "pl": today_verticals["demand"]["plActual"]},
    ],
}

history = [s for s in history if s["date"] != TODAY]
history.append(new_snap)
history = sorted(history, key=lambda s: s["date"])[-30:]
snap_store["snapshots"] = history
snap_store["lastUpdated"] = TODAY

with open("snapshots.json", "w") as f:
    json.dump(snap_store, f, indent=2)
print("\n  snapshots.json updated.")


# ── 7. Build data.json ────────────────────────
VERTICAL_META = {
    "edu_leadgen":   {"book": "Supply", "name": "Edu: LeadGen",   "tag": "vt-edu"},
    "edu_clicks":    {"book": "Supply", "name": "Edu: Clicks",    "tag": "vt-edu"},
    "autos_clicks":  {"book": "Supply", "name": "Autos: Clicks",  "tag": "vt-autos"},
    "autos_leadgen": {"book": "Supply", "name": "Autos: LeadGen", "tag": "vt-autos"},
    "ptp":           {"book": "Supply", "name": "PTP",            "tag": "vt-ptp"},
    "others":        {"book": "Supply", "name": "Others",         "tag": "vt-other"},
    "demand":        {"book": "Demand", "name": "Ad.Tech",        "tag": "vt-demand"},
}


def build_adv_list(rows, sub_fn=None):
    result = []
    for r in rows:
        sub = sub_fn(r) if sub_fn else r["vertical"]
        result.append({
            "sub":     sub,
            "name":    r["name"],
            "target":  round(r["target"]),
            "netRev":  round(r["netRev"]),
            "tac":     round(r["tac"]),
            "pl":      round(r["plActual"]),
            "pacing":  round(r["plPacing"]) if r["plPacing"] else None,
            "rag":     r["rag"],
            "blocker": r["blocker"],
            "action":  r["action"],
        })
    return result


def build_daily_overview():
    if not deltas:
        return None
    d_supply = [deltas[k] for k in supply_keys if k in deltas]
    d_demand = [deltas[k] for k in demand_keys if k in deltas]
    s_rev = sum(x["netRevenue"] for x in d_supply)
    d_rev = sum(x["netRevenue"] for x in d_demand)
    return {
        "netRev":    round(s_rev + d_rev),
        "tac":       round(sum(x["tac"]      for x in deltas.values())),
        "pl":        round(sum(x["plActual"]  for x in deltas.values())),
        "supplyRev": round(s_rev),
        "demandRev": round(d_rev),
        "supplyPL":  round(sum(x["plActual"] for x in d_supply)),
        "demandPL":  round(sum(x["plActual"] for x in d_demand)),
        "pacing": None, "target": None, "variance": None,
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
            "pacing":   None, "target": None,
            "variance": f"+${pl:,.0f}" if pl >= 0 else f"-${abs(pl):,.0f}",
            "focus":    "Up"        if pl >= 0 else "Down",
            "focusCls": "focus-low" if pl >= 0 else "focus-high",
        })
    return rows


today_display = date.today().strftime("%-d %b %Y")
prev_display  = date.fromisoformat(YESTERDAY).strftime("%-d %b %Y") if deltas else ""

data_json = {
    "lastUpdated": f"Auto-refreshed {today_display} 1:00pm IST",
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
        "eduAdvertisers":   build_adv_list(
            edu_rows_p,
            sub_fn=lambda r: "LeadGen" if "leadgen" in r["vertical"].lower() else "Clicks"
        ),
        "autosAdvertisers": build_adv_list(
            autos_rows_p,
            sub_fn=lambda r: "Clicks" if "clicks" in r["vertical"].lower() else "LeadGen"
        ),
        "ptpPublishers":    build_adv_list(ptp_rows_p),
        "demandAdvertisers": build_adv_list(demand_rows_p),
        "othersAdvertisers": build_adv_list(others_rows_p),
    },
    "daily": {
        "label":     f"Daily Change vs {prev_display}",
        "period":    "Daily",
        "overview":  build_daily_overview(),
        "verticals": build_daily_verticals(),
        "_note":     f"Delta = {TODAY} MTD minus {YESTERDAY} MTD. All figures 5% AAC-adjusted from dedicated tabs.",
    } if deltas else None,
}

with open("data.json", "w") as f:
    json.dump(data_json, f, indent=2)

print("  data.json written.")
print(f"\n=== Refresh complete — P&L: ${grand_totals['plActual']:,.0f} | Net Rev: ${grand_totals['netRevenue']:,.0f} ===\n")
