#!/usr/bin/env python3
"""
Generate the VD Rebate dashboard data (index.html) with a per-MONTH history so the
page can switch between months and view each one separately.

- Actuals: Sales_Invoices_All.parquet, ivqty>0. One data block is built for EVERY
  (year, month) present in the parquet, up to the cutoff month (today, or the
  override passed on the command line).
- Targets: "Quantity Targets.xlsx" -> "Qty Targets" sheet -> the column whose header
  matches the month name (seasonally-adjusted monthly requirement per product/size).
  Each size-line target is split across its colour groups by the ACTUAL sales mix for
  that month (falls back to the T3 tier mix if a line has no sales). The seasonal
  plan is the current one; it is applied to each month as a consistent yardstick.
- Reps: AC/AP/BM/BV/NP are named; every other smref (incl. blank) -> "Internal".
  Company totals == sum of the 6 buckets (exact reconciliation, asserted per month).
- Injects ALL_MONTHS / MONTH_KEYS / CURRENT_KEY / REP_NAMES into dashboard/index.html.
  KPIs, rep buttons, month selector and labels are all computed client-side.

Run:  python build_vd_dashboard.py            (all months up to the current month)
      python build_vd_dashboard.py 2026 6      (override cutoff month, for testing)
"""
import sys, re, json, datetime
from functools import lru_cache
from pathlib import Path
import pandas as pd

ONE = str(Path.home() / "OneDrive" / "1.Projects" / "1.Olympic Paints")  # machine-agnostic (Administrator / quint)
PARQUET = Path(ONE) / "3.Resources" / "16.Sales and Other data" / "Sales_Invoices_All.parquet"
XLSX    = Path(ONE) / "2.Areas" / "1. Sales" / "New Sales Targets and Pricing" / "Quantity Targets.xlsx"
HTML    = Path(__file__).resolve().parent / "dashboard" / "index.html"

NAMED = ["AC", "AP", "BM", "BV", "NP"]
MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

# vol_key -> (size_line key used in Qty Targets, T3 tier qty for fallback split, T3 price)
TIER = {
 "Decor Colours 20L":("Decor 20L",14000,255), "Decor White/Cream 20L":("Decor 20L",5500,235),
 "Decor Colours 5L":("Decor 5L",1700,80),     "Decor White/Cream 5L":("Decor 5L",750,70),
 "High Gloss Colours 1L":("HG 1L",18000,60),   "High Gloss White/Black 1L":("HG 1L",5000,65),
 "High Gloss Colours 5L":("HG 5L",14500,250),  "High Gloss White/Black 5L":("HG 5L",4500,240),
 "QD Enamel Colours 1L":("QD 1L",9000,65),     "QD Enamel Colours 5L":("QD 5L",12500,250),
 "Thinner 750ml":("Thinner 750ml",6150,250),   "Thinner 5L":("Thinner 5L",5000,135),
 "Varnish Copal 1L":("Varnish 1L",2150,75),    "Varnish Colours 1L":("Varnish 1L",7300,90),
 "Varnish Copal 5L":("Varnish 5L",800,270),    "Varnish Colours 5L":("Varnish 5L",2500,295),
 "Universal Roof 20L":("Roof 20L",1700,450),   "Universal Roof 5L":("Roof 5L",1000,120),
}
ORDER = list(TIER.keys())
PRICE = {vk: p for vk, (ln, q, p) in TIER.items()}
# size-line -> row label / size in the Qty Targets sheet (header row of each block)
LINE_ROW = {  # size_line : (block_label, size_label)
 "Decor 20L":("Decor","20LT"), "Decor 5L":("Decor","5LT"),
 "HG 1L":("High Gloss","1LT"), "HG 5L":("High Gloss","5LT"),
 "QD 1L":("Q.D. Enamel","1LT"), "QD 5L":("Q.D. Enamel","5LT"),
 "Thinner 5L":("Thinner","5LT"), "Thinner 750ml":("Thinner","750ML X 12"),
 "Roof 20L":("Universal Roof Paint","20LT"), "Roof 5L":("Universal Roof Paint","5LT"),
 "Varnish 1L":("Wood Varnish","1LT"), "Varnish 5L":("Wood Varnish","5LT"),
}

def size_of(pn):
    if re.search(r'\b20\s?L', pn) or '20LT' in pn: return '20L'
    if re.search(r'\b5\s?L', pn) or '5LT' in pn:  return '5L'
    if '750ML' in pn or '750 ML' in pn:           return '750ML'
    if '500ML' in pn or '500 ML' in pn:           return '500ML'
    if re.search(r'\b1\s?L', pn) or '1LT' in pn:  return '1L'
    return None

def categorize(pn):
    if not isinstance(pn, str): return None
    pn = pn.upper(); sz = size_of(pn)
    wb = bool(re.search(r'\bWHITE\b|\bBLACK\b', pn)); wc = bool(re.search(r'\bWHITE\b|\bCREAM\b', pn))
    if 'HIGH GLOSS' in pn:
        if sz not in ('1L','5L'): return None  # 500ml and 20L not in VD programme
        # CREAM and IVORY share the White/Black tier (similar neutral price point)
        wb_ext = bool(re.search(r'\bWHITE\b|\bBLACK\b|\bCREAM\b|\bIVORY\b', pn))
        return f"High Gloss {'White/Black' if wb_ext else 'Colours'} {sz}"
    if re.search(r'Q\.?\s*D\b', pn):
        if sz not in ('1L','5L') or wb: return None
        return f"QD Enamel Colours {sz}"
    if 'DECOR' in pn and 'MASTER' not in pn:
        if sz not in ('20L','5L'): return None
        return f"Decor {'White/Cream' if wc else 'Colours'} {sz}"
    if 'VARNISH' in pn:
        if sz not in ('1L','5L'): return None
        return f"Varnish {'Copal' if 'COPAL' in pn else 'Colours'} {sz}"
    if 'ROOF' in pn and ('UNIVERSAL' in pn or 'UNIV' in pn):
        if sz not in ('20L','5L'): return None
        return f"Universal Roof {sz}"
    if 'THINNER' in pn:  # catches THINNER and THINNERS
        if sz == '5L':    return "Thinner 5L"
        if sz == '750ML': return "Thinner 750ml"
        if '12X' in pn:   return "Thinner 750ml"
    return None

@lru_cache(maxsize=None)
def month_targets(month_name):
    """Read the Qty Targets sheet and return {size_line: monthly_qty} for the month."""
    d = pd.read_excel(XLSX, sheet_name="Qty Targets", header=None)
    hdr = d.iloc[0].tolist()
    col = next((i for i, v in enumerate(hdr) if str(v).strip().lower() == month_name.lower()), None)
    if col is None:
        raise SystemExit(f"Month column '{month_name}' not found in Qty Targets sheet headers: {hdr}")
    # map (block_label, size_label) -> value by scanning rows; block label is in col 0 with blanks under it
    cur_block = None
    vals = {}
    for _, row in d.iterrows():
        c0 = row[0]
        if isinstance(c0, str) and c0.strip() and c0.strip() not in ('20LT','5LT','1LT','750ML X 12'):
            cur_block = c0.strip()
        if isinstance(c0, str) and c0.strip() in ('20LT','5LT','1LT','750ML X 12'):
            vals[(cur_block, c0.strip())] = row[col]
    out = {}
    for line, (blk, sz) in LINE_ROW.items():
        v = vals.get((blk, sz))
        out[line] = float(v) if v is not None and not pd.isna(v) else 0.0
    return out

def compute_month(df, year, month, today):
    """Build the {t3, rep, config} data block for one (year, month)."""
    month_name = MONTHS[month-1]
    cur = df[(df.year == year) & (df.month == month)].copy()
    cur['vk'] = cur.prodname.map(categorize)
    ins = cur[cur.vk.notna()].copy()

    line_tgt = month_targets(month_name)
    grp = ins.groupby('vk').agg(q=('ivqty','sum'), nett=('ivnett','sum')).to_dict('index')
    line_q, line_tier = {}, {}
    for vk, (ln, tq, pr) in TIER.items():
        line_q[ln]    = line_q.get(ln, 0) + (grp.get(vk, {}).get('q', 0) or 0)
        line_tier[ln] = line_tier.get(ln, 0) + tq

    rows = []
    for vk in ORDER:
        ln, tq, price = TIER[vk]
        g = grp.get(vk, {}); curv = int(round(g.get('q', 0) or 0))
        # split month line-target across colour groups by ACTUAL mix; fallback to tier mix
        share = (g.get('q', 0) or 0) / line_q[ln] if line_q[ln] > 0 else tq / line_tier[ln]
        tgt = round(line_tgt.get(ln, 0) * share)
        bl = round(g['nett']/g['q'], 2) if g.get('q') else 0
        pct = round(curv/tgt*1000)/10 if tgt > 0 else None
        rows.append(dict(vol_key=vk, current_volume=curv, t3_adjusted=tgt, peak_t3=tgt,
            t3_threshold_total=tgt, pct_to_t3=pct, pct_to_t3_adj=pct, gap_units=tgt-curv,
            t3_unit_price_avg=price, blended_price=bl,
            blended_vs_t3=round(bl-price, 2) if bl else 0,
            snapshot_date=today.isoformat()))

    rep_data = {}
    for rep in NAMED + ['Internal']:
        g = ins[ins.rep == rep].groupby('vk').agg(v=('ivqty','sum'), nett=('ivnett','sum'),
            cust=('accno','nunique'))
        rep_data[rep] = {vk: dict(vol_key=vk, rep_volume=int(round(x.v)),
            customer_count=int(x.cust), blended_price=round(x.nett/x.v, 2))
            for vk, x in g.iterrows()}

    is_current = (year == today.year and month == today.month)
    config = {"snapshot_date": today.isoformat(), "month_name": month_name,
              "year": year, "month": month, "key": f"{year}-{month:02d}",
              "is_current": is_current}

    total  = sum(r['current_volume'] for r in rows)
    repsum = sum(sum(v['rep_volume'] for v in rep_data[rep].values()) for rep in rep_data)
    assert repsum == total, f"reconcile fail {year}-{month:02d}: {repsum} != {total}"
    return {"t3": rows, "rep": rep_data, "config": config}, total

def main():
    today = datetime.date.today()
    if len(sys.argv) == 3:
        cy, cm = int(sys.argv[1]), int(sys.argv[2])
    else:
        cy, cm = today.year, today.month
    cutoff = (cy, cm)
    current_key = f"{cy}-{cm:02d}"

    df = pd.read_parquet(PARQUET)
    df = df[df.ivqty > 0].copy()
    df['rep'] = df.smref.where(df.smref.isin(NAMED), 'Internal')

    # every (year, month) with data, up to the cutoff month, newest first
    yms = sorted({(int(y), int(m)) for y, m in zip(df.year, df.month) if (int(y), int(m)) <= cutoff},
                 reverse=True)

    all_months, keys_desc, totals = {}, [], {}
    for (y, m) in yms:
        data, total = compute_month(df, y, m, today)
        key = data["config"]["key"]
        all_months[key] = data
        keys_desc.append(key)
        totals[key] = total

    if current_key not in all_months and keys_desc:
        current_key = keys_desc[0]

    names = {"AC":"Aboo Cassim","AP":"Amit Patel","BM":"Byron Minnie",
             "BV":"Bhadresh Vallabh","NP":"Nikhil Panchal","Internal":"Internal"}

    # ── inject ──
    html = HTML.read_text(encoding="utf-8")
    j = lambda o: json.dumps(o, separators=(',', ':'), ensure_ascii=False)
    html = re.sub(r'const ALL_MONTHS = \{.*?\};\nconst MONTH_KEYS',
                  'const ALL_MONTHS = '+j(all_months)+';\nconst MONTH_KEYS', html, count=1, flags=re.S)
    html = re.sub(r'const MONTH_KEYS = \[.*?\];\nconst CURRENT_KEY',
                  'const MONTH_KEYS = '+j(keys_desc)+';\nconst CURRENT_KEY', html, count=1, flags=re.S)
    html = re.sub(r'const CURRENT_KEY = ".*?";\nconst REP_NAMES',
                  'const CURRENT_KEY = '+j(current_key)+';\nconst REP_NAMES', html, count=1, flags=re.S)
    html = re.sub(r'const REP_NAMES  = \{.*?\};', 'const REP_NAMES  = '+j(names)+';', html, count=1, flags=re.S)
    HTML.write_text(html, encoding="utf-8")

    print(f"VD dashboard rebuilt: {len(all_months)} months, current={current_key} ({today.isoformat()})")
    for k in keys_desc:
        print(f"  {k}: total={totals[k]}")

if __name__ == "__main__":
    main()
