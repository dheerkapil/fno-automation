#!/usr/bin/env python3
"""
fno_eod.py - NSE end-of-day F&O buildup screener.

Rebuilds, in ~40 seconds and a few hundred KB, what
consolidated_manual_date_fno.xlsm did with Power Query + tens of thousands of
SUMIFS over a 26 MB workbook:

  * pulls the last N trading sessions straight from NSE archives (UDiFF format)
  * builds a rolling 20-session window: cash (delivery %, qty/trade),
    futures (settle, OI, OI change %), options (CE/PE OI in lots)
  * classifies every symbol/day as LB / SB / SC / LU using price change vs
    open-interest change against a configurable threshold
  * writes one small .xlsx (all sheets), one .csv per sheet, and a PNG of the
    buildup grid
  * optionally ships all of it to Telegram

Nothing is stored between runs: the cache directory is temporary, so the
footprint is the size of the output files only.

Usage:
    python fno_eod.py --days 20 --out out/
    python fno_eod.py --date 2026-11-06 --telegram
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import os
import shutil
import sys
import time
import zipfile

import pandas as pd
import requests

# --------------------------------------------------------------------------- #
# config
# --------------------------------------------------------------------------- #
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}

URL_FO = "https://nsearchives.nseindia.com/content/fo/BhavCopy_NSE_FO_0_0_0_{d}_F_0000.csv.zip"
URL_SEC = "https://nsearchives.nseindia.com/products/content/sec_bhavdata_full_{d}.csv"
URL_IND = "https://nsearchives.nseindia.com/content/indices/ind_nifty500list.csv"

CACHE = os.environ.get("FNO_CACHE", "cache")
THRESHOLD = float(os.environ.get("FNO_THRESHOLD", "0.08"))   # 8% OI change
SIGNAL_COLORS = {"LB": "C6EFCE", "SB": "FFC7CE", "SC": "BDD7EE", "LU": "FFE699"}
SIGNAL_MEANING = {
    "LB": "Long buildup   (price up, OI up)",
    "SB": "Short buildup  (price down, OI up)",
    "SC": "Short covering (price up, OI down)",
    "LU": "Long unwinding (price down, OI down)",
}


# --------------------------------------------------------------------------- #
# fetch
# --------------------------------------------------------------------------- #
def looks_valid(url: str, blob: bytes) -> bool:
    """NSE sometimes answers 200 with an HTML page; reject anything that is not
    the file we asked for."""
    if url.endswith(".zip"):
        return blob[:2] == b"PK"
    head = blob[:400].lstrip(b"\xef\xbb\xbf").decode("utf-8", "ignore").upper()
    first = head.split("\n")[0]          # "SYMBOL, SERIES, ..." or "Company Name,...,SYMBOL,..."
    return "SYMBOL" in first


def fetch(url: str, tries: int = 4, timeout: int = 60) -> bytes | None:
    """GET with retries. Returns None on a 404 (holiday / not published yet)."""
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            if not looks_valid(url, r.content):
                raise ValueError(f"unexpected payload ({r.headers.get('content-type')})")
            return r.content
        except Exception as exc:                      # noqa: BLE001
            last = exc
            time.sleep(2 * (i + 1))
    print(f"  ! failed {url}: {last}", file=sys.stderr)
    return None


def cached(name: str, url: str) -> bytes | None:
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, name)
    if os.path.exists(path) and os.path.getsize(path) > 0:
        with open(path, "rb") as fh:
            return fh.read()
    data = fetch(url)
    if data:
        with open(path, "wb") as fh:
            fh.write(data)
    return data


def sec_url(d: dt.date) -> str:
    return URL_SEC.format(d=d.strftime("%d%m%Y"))


def fo_url(d: dt.date) -> str:
    return URL_FO.format(d=d.strftime("%Y%m%d"))


def industry_map() -> dict:
    """Symbol -> industry, the lookup the master kept on its general sheet."""
    blob = cached("ind_nifty500list.csv", URL_IND)
    if not blob:
        return {}
    try:
        df = pd.read_csv(io.BytesIO(blob))
        df.columns = [c.strip() for c in df.columns]
        return dict(zip(df["Symbol"].astype(str).str.strip(),
                        df["Industry"].astype(str).str.strip()))
    except Exception:                                   # noqa: BLE001
        return {}


def trading_dates(end: dt.date, n: int) -> list[dt.date]:
    """Walk backwards from `end` until n sessions with a cash bhavcopy exist."""
    out: list[dt.date] = []
    d = end
    guard = 0
    while len(out) < n and guard < n * 3 + 15:
        guard += 1
        if d.weekday() < 5:                      # skip Sat/Sun without a request
            # a session counts only if BOTH cash and F&O files exist: NSE
            # publishes no UDiFF bhavcopy for special (muhurat) sessions
            if cached(f"sec_{d:%Y%m%d}.csv", sec_url(d)) and \
               cached(f"fo_{d:%Y%m%d}.zip", fo_url(d)):
                out.append(d)
        d -= dt.timedelta(days=1)
    return sorted(out)


# --------------------------------------------------------------------------- #
# parse one session
# --------------------------------------------------------------------------- #
def read_sec(blob: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(blob))
    df.columns = [c.strip() for c in df.columns]
    df = df[df["SERIES"].astype(str).str.strip() == "EQ"]
    num = ["CLOSE_PRICE", "TTL_TRD_QNTY", "NO_OF_TRADES", "DELIV_QTY", "DELIV_PER"]
    for c in num:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["SYMBOL"] = df["SYMBOL"].astype(str).str.strip()
    return df[["SYMBOL"] + num]


def read_fo(blob: bytes) -> pd.DataFrame:
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        name = [n for n in z.namelist() if n.lower().endswith(".csv")][0]
        df = pd.read_csv(io.BytesIO(z.read(name)))
    df.columns = [c.strip() for c in df.columns]
    keep = ["FinInstrmTp", "TckrSymb", "XpryDt", "OptnTp", "SttlmPric", "ClsPric",
            "OpnIntrst", "ChngInOpnIntrst", "TtlTradgVol", "TtlNbOfTxsExctd",
            "NewBrdLotQty", "UndrlygPric"]
    df = df[[c for c in keep if c in df.columns]].copy()
    df["TckrSymb"] = df["TckrSymb"].astype(str).str.strip()
    for c in ["SttlmPric", "ClsPric", "OpnIntrst", "ChngInOpnIntrst",
              "TtlTradgVol", "TtlNbOfTxsExctd", "NewBrdLotQty", "UndrlygPric"]:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["XpryDt"] = pd.to_datetime(df["XpryDt"], errors="coerce")
    return df


# --------------------------------------------------------------------------- #
# build the window
# --------------------------------------------------------------------------- #
def build(dates: list[dt.date]):
    cash_rows, fut_rows, opt_rows = [], [], []
    for d in dates:
        sec = cached(f"sec_{d:%Y%m%d}.csv", sec_url(d))
        fo = cached(f"fo_{d:%Y%m%d}.zip", fo_url(d))
        if not sec or not fo:
            print(f"  - {d} skipped (file not published)")
            continue
        c = read_sec(sec)
        c["date"] = d
        cash_rows.append(c)

        f = read_fo(fo)
        fut = f[f["FinInstrmTp"] == "STF"]
        if not fut.empty:
            fut = fut.assign(date=d)
            fut_rows.append(fut)

        opt = f[f["FinInstrmTp"] == "STO"].copy()
        if not opt.empty:
            opt["lot"] = opt["NewBrdLotQty"].replace(0, pd.NA)
            opt["oi_lots"] = opt["OpnIntrst"] / opt["lot"]
            opt = opt.assign(date=d)
            opt_rows.append(opt)

    cash = pd.concat(cash_rows, ignore_index=True) if cash_rows else pd.DataFrame()
    fut = pd.concat(fut_rows, ignore_index=True) if fut_rows else pd.DataFrame()
    opt = pd.concat(opt_rows, ignore_index=True) if opt_rows else pd.DataFrame()
    return cash, fut, opt


def label(dates: list[dt.date]) -> dict:
    """D1..Dn, oldest first, like the workbook's date_table."""
    return {d: f"D{i+1}" for i, d in enumerate(dates)}


# --------------------------------------------------------------------------- #
# the sheets
# --------------------------------------------------------------------------- #
def sheets(cash, fut, opt, dates, threshold, raw=False):
    """Rebuild every sheet of the master workbook, full width.

    Each sheet keeps the master's shape: one row per symbol, then the D1..D20
    grid for every measure it carried - not just a 20-day summary.
    """
    lab = label(dates)
    cols = [f"D{i+1}" for i in range(len(dates))]
    out = {}

    # ---------------- cash -------------------------------------------------
    cash = cash[cash["date"].isin(lab)].copy()
    cash["day"] = cash["date"].map(lab)
    px = cash.pivot_table(index="SYMBOL", columns="day", values="CLOSE_PRICE", aggfunc="last").reindex(columns=cols)
    dv = cash.pivot_table(index="SYMBOL", columns="day", values="DELIV_PER", aggfunc="last").reindex(columns=cols)
    qty = cash.pivot_table(index="SYMBOL", columns="day", values="TTL_TRD_QNTY", aggfunc="sum").reindex(columns=cols)
    trd = cash.pivot_table(index="SYMBOL", columns="day", values="NO_OF_TRADES", aggfunc="sum").reindex(columns=cols)
    delq = cash.pivot_table(index="SYMBOL", columns="day", values="DELIV_QTY", aggfunc="sum").reindex(columns=cols)
    qt = qty / trd.replace(0, pd.NA)
    cash_close = px.ffill(axis=1)

    # the master carries four 20-session blocks here - price, volume, delivery %
    # and qty/trade - each with its own min/max pair, plus the two window
    # aggregates that sat between them.
    ca = pd.DataFrame(index=px.index)
    ca["min_price"] = px.min(axis=1)
    ca["max_price"] = px.max(axis=1)
    for c in cols:
        ca[f"Close_{c}"] = px[c]
    ca["min_vol"] = qty.min(axis=1)
    ca["max_vol"] = qty.max(axis=1)
    for c in cols:
        ca[f"Volume_{c}"] = qty[c]
    ca["deliv%_20D"] = delq.sum(axis=1) / qty.sum(axis=1).replace(0, pd.NA) * 100
    ca["min_deliv"] = dv.min(axis=1)
    ca["max_deliv"] = dv.max(axis=1)
    for c in cols:
        ca[f"Deliv%_{c}"] = dv[c]
    ca["qty/trade_20D"] = qty.sum(axis=1) / trd.sum(axis=1).replace(0, pd.NA)
    ca["min_qtytrade"] = qt.min(axis=1)
    ca["max_qtytrade"] = qt.max(axis=1)
    for c in cols:
        ca[f"Qty/Trade_{c}"] = qt[c]
    ind = industry_map()
    ca.insert(0, "industry", [ind.get(sym, "") for sym in ca.index])
    # assigned after the futures universe is known (see below)

    # ---------------- futures ---------------------------------------------
    fut = fut[fut["date"].isin(lab)].copy()
    fut["day"] = fut["date"].map(lab)
    fut = fut.dropna(subset=["OpnIntrst"]).sort_values(["TckrSymb", "XpryDt", "date"])
    fut["px_chg%"] = (fut.groupby(["TckrSymb", "XpryDt"], group_keys=False)["SttlmPric"]
                         .pct_change(fill_method=None) * 100)
    near = fut.loc[fut.groupby(["TckrSymb", "day"])["XpryDt"].idxmin()]
    price = near.pivot_table(index="TckrSymb", columns="day", values="SttlmPric", aggfunc="last").reindex(columns=cols)
    px_pct = near.pivot_table(index="TckrSymb", columns="day", values="px_chg%", aggfunc="last").reindex(columns=cols)
    lot = near.pivot_table(index="TckrSymb", columns="day", values="NewBrdLotQty", aggfunc="last").reindex(columns=cols)

    agg = fut.groupby(["TckrSymb", "day"]).agg(
        oi=("OpnIntrst", "sum"), coi=("ChngInOpnIntrst", "sum"),
        trades=("TtlNbOfTxsExctd", "sum")).reset_index()
    agg["oi_pct"] = agg["coi"] / (agg["oi"] - agg["coi"]).replace(0, pd.NA) * 100
    agg["ct_per_trade"] = agg["coi"] / agg["trades"].replace(0, pd.NA)
    oi = agg.pivot_table(index="TckrSymb", columns="day", values="oi", aggfunc="last").reindex(columns=cols)
    oi_pct = agg.pivot_table(index="TckrSymb", columns="day", values="oi_pct", aggfunc="last").reindex(columns=cols)
    per_trade = (agg.pivot_table(index="TckrSymb", columns="day", values="ct_per_trade",
                                 aggfunc="last").reindex(columns=cols)
                    .div(lot.replace(0, pd.NA)))

    fa = pd.DataFrame(index=price.index)
    fa["min"] = price.min(axis=1)
    fa["max"] = price.max(axis=1)
    fa["lot"] = lot[cols[-1]]
    for c in cols:
        fa[f"Settle_{c}"] = price[c]
    for c in cols:
        fa[f"PxChg%_{c}"] = px_pct[c]
    for c in cols:
        fa[f"OI_{c}"] = oi[c]
    for c in cols:
        fa[f"OIchg%_{c}"] = oi_pct[c]
    for c in cols:
        fa[f"OIperTrade_{c}"] = per_trade[c]
    fa["OIperTrade_20D_avg"] = per_trade.mean(axis=1)      # the master's BP column
    out["fut_alys_data"] = fa.round(2).sort_index().rename_axis("SYMBOL")
    # the master's cash sheet lists the F&O universe, not all ~2,800 cash symbols
    out["cash_alys_fno"] = ca.loc[ca.index.intersection(price.index)].round(2).sort_index()

    # ---------------- options ---------------------------------------------
    op = opt[opt["date"].isin(lab)].copy()
    op["day"] = op["date"].map(lab)
    op["oi_lots"] = pd.to_numeric(op["oi_lots"], errors="coerce")
    op = op.dropna(subset=["oi_lots"])
    ce = pe = None
    if not op.empty:
        per_exp = (op.groupby(["TckrSymb", "day", "XpryDt", "OptnTp"])["oi_lots"]
                     .sum().reset_index())
        tot = per_exp.groupby(["TckrSymb", "day", "XpryDt"])["oi_lots"].sum().reset_index()
        dom = tot.loc[tot.groupby(["TckrSymb", "day"])["oi_lots"].idxmax()]
        dom = dom.merge(per_exp, on=["TckrSymb", "day", "XpryDt"], suffixes=("_tot", ""))
        ce = (dom[dom["OptnTp"] == "CE"].pivot_table(index="TckrSymb", columns="day",
                                                     values="oi_lots", aggfunc="sum")
              .reindex(columns=cols))
        pe = (dom[dom["OptnTp"] == "PE"].pivot_table(index="TckrSymb", columns="day",
                                                     values="oi_lots", aggfunc="sum")
              .reindex(columns=cols))
        oa = pd.DataFrame(index=ce.index.union(pe.index))
        for c in cols:
            oa[f"CE_lots_{c}"] = ce[c]
        for c in cols:
            oa[f"PE_lots_{c}"] = pe[c]
        for c in cols:
            oa[f"PCR_{c}"] = (pe[c] / ce[c].replace(0, pd.NA))
        out["op_alys_data"] = oa.round(2).sort_index().rename_axis("SYMBOL")

    # ---------------- buildup + divergence --------------------------------
    idx = price.index.intersection(oi_pct.index)
    sig = pd.DataFrame("", index=idx, columns=cols, dtype=object)
    for c in cols:
        p, o = px_pct.get(c), oi_pct.get(c)
        if p is None or o is None:
            continue
        s = pd.Series("", index=idx, dtype=object)
        up, dn = p > 0, p < 0
        s[up & (o > threshold * 100)] = "LB"
        s[dn & (o > threshold * 100)] = "SB"
        s[up & (o < -threshold * 100)] = "SC"
        s[dn & (o < -threshold * 100)] = "LU"
        sig[c] = s

    div = pd.DataFrame(index=idx, columns=cols, dtype=float)
    for c in cols:
        spot = cash_close.reindex(idx).get(c)
        if spot is not None:
            div[c] = ((price.get(c) - spot) / spot).abs() * 100
    div.columns = [f"{c}_div%" for c in cols]
    out["buildup_summary"] = sig.join(div.round(2)).sort_index().rename_axis("SYMBOL")

    # ---------------- one-stock panel + normalised twin -------------------
    focus = pick_focus(sig, oi_pct, cols)

    metrics = [
        ("Date", pd.Series({c: d.strftime("%d-%b-%Y") for c, d in zip(cols, dates)})),
        ("Cash Price", cash_close.loc[focus] if focus in cash_close.index else None),
        ("Volume", qty.loc[focus] if focus in qty.index else None),
        ("Delivery%", dv.loc[focus] if focus in dv.index else None),
        ("Cash Qty/Trade", qt.loc[focus] if focus in qt.index else None),
        ("Futures Price", price.loc[focus] if focus in price.index else None),
        ("Cumulative OI", oi.loc[focus] if focus in oi.index else None),
        ("OI chg %", oi_pct.loc[focus] if focus in oi_pct.index else None),
        ("FNO Contracts/Trade", per_trade.loc[focus] if focus in per_trade.index else None),
    ]
    if ce is not None and focus in ce.index:
        metrics += [("CE OI lots", ce.loc[focus]), ("PE OI lots", pe.loc[focus]),
                    ("PCR", pe.loc[focus] / ce.loc[focus].replace(0, pd.NA))]
    if focus in sig.index:
        metrics += [("Signal", sig.loc[focus])]

    panel = pd.DataFrame({name: (series.reindex(cols) if isinstance(series, pd.Series)
                                 else pd.Series(dtype=float))
                          for name, series in metrics}).T
    panel.columns = cols
    panel.index.name = focus
    out["stock_specific_data"] = panel.round(2)

    norm = panel.drop(index=["Date", "Signal"], errors="ignore").copy()
    for name in norm.index:
        row = pd.to_numeric(norm.loc[name], errors="coerce")
        if row.notna().sum() < 2 or row.max() == row.min():
            norm.loc[name] = 0.0
            continue
        norm.loc[name] = ((row - row.min()) / (row.max() - row.min())).round(3)
    out["normalized_data"] = norm

    # ---------------- control panel ---------------------------------------
    rows = [{"Days": c, "Date": d.strftime("%d-%b-%Y"), "ReportDate": d.strftime("%d%m%Y")}
            for c, d in zip(cols, dates)]
    rows += [
        {"Days": "", "Date": "", "ReportDate": ""},
        {"Days": "Threshold Buildup Change", "Date": threshold, "ReportDate": ""},
        {"Days": "Sessions", "Date": len(dates), "ReportDate": ""},
        {"Days": "F&O symbols", "Date": len(price.index), "ReportDate": ""},
        {"Days": "Cash symbols", "Date": len(px.index), "ReportDate": ""},
        {"Days": "Focus symbol", "Date": focus, "ReportDate": ""},
        {"Days": "Generated", "Date": dt.datetime.now().strftime("%d-%b-%Y %H:%M"),
         "ReportDate": ""},
    ]
    out["general"] = pd.DataFrame(rows)

    if raw:
        # the master's three hidden data sheets, for the latest session only -
        # enough to eyeball the source rows without shipping 620k option lines
        last_day = dates[-1]
        out["cash_data"] = (cash[cash["date"] == last_day]
                            .drop(columns=["date", "day"], errors="ignore")
                            .sort_values("SYMBOL"))
        out["futures_data"] = (fut[fut["date"] == last_day]
                               .drop(columns=["date", "day"], errors="ignore")
                               .sort_values("TckrSymb"))
        out["options_data"] = (opt[opt["date"] == last_day]
                               .drop(columns=["date", "day"], errors="ignore")
                               .sort_values(["TckrSymb", "XpryDt", "OptnTp"]))

    out["_focus"] = focus
    return out


def pick_focus(sig, oi_pct, cols):
    """The symbol with the biggest absolute OI move on the last session."""
    last = oi_pct[cols[-1]].dropna()
    if last.empty:
        return sig.index[0] if len(sig.index) else ""
    return last.abs().idxmax()


# --------------------------------------------------------------------------- #
# render
# --------------------------------------------------------------------------- #
SUPPORT_COLS = {"min_price", "max_price", "min_vol", "max_vol", "min_deliv", "max_deliv",
                "min_qtytrade", "max_qtytrade", "deliv%_20D", "qty/trade_20D",
                "lot", "OIperTrade_20D_avg", "min", "max"}


def _day_index(name: str):
    """D7 out of 'Close_D7' / 'OIchg%_D7' / 'D7' / 'D7_div%'."""
    for tok in name.replace("_div%", "").split("_"):
        if tok.startswith("D") and tok[1:].isdigit():
            return int(tok[1:])
    return None


def write_xlsx(out: dict, path: str, show_days: int = 8, master_layout: bool = True):
    focus = out.pop("_focus", None)
    order = ["general", "cash_alys_fno", "fut_alys_data", "op_alys_data",
             "buildup_summary", "stock_specific_data", "normalized_data"]
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        for name in [n for n in order if n in out] + [n for n in out if n not in order]:
            if name.startswith("_"):
                continue
            out[name].to_excel(xl, sheet_name=name[:31])
    # light styling: signal colours, colour scales, freeze panes, widths
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils import get_column_letter
    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 22
        for col in ws.iter_cols(min_row=1, max_row=1):
            col[0].font = Font(bold=True)
            col[0].alignment = Alignment(horizontal="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2, min_col=2):
            for cell in row:
                if isinstance(cell.value, str) and cell.value in SIGNAL_COLORS:
                    cell.fill = PatternFill("solid", fgColor=SIGNAL_COLORS[cell.value])
                    cell.alignment = Alignment(horizontal="center")
        # ---- presentation: hide support columns and the older sessions -----
        headers = [c.value for c in ws[1]]
        if master_layout:
            for idx, name in enumerate(headers, start=1):
                if not isinstance(name, str):
                    continue
                di = _day_index(name)
                hide = False
                if name in SUPPORT_COLS:
                    hide = True                     # min/max/aggregate = support
                elif di is not None and ws.title in ("cash_alys_fno", "fut_alys_data",
                                                     "op_alys_data"):
                    hide = di <= 20 - show_days     # older sessions stay, hidden
                elif di is not None and ws.title == "buildup_summary":
                    # the signal grid stays whole; only the older divergence
                    # columns fold away, as in the master
                    hide = name.endswith("_div%") and di <= 20 - show_days
                if hide:
                    ws.column_dimensions[get_column_letter(idx)].hidden = True

        # ---- colour scales, one per day column (master's rule shape) --------
        if ws.title in ("cash_alys_fno", "fut_alys_data", "op_alys_data"):
            for idx, name in enumerate(headers, start=1):
                if not isinstance(name, str) or _day_index(name) is None:
                    continue
                col = get_column_letter(idx)
                ws.conditional_formatting.add(
                    f"{col}2:{col}{ws.max_row}",
                    ColorScaleRule(start_type="min", start_color="F8696B",
                                   mid_type="percentile", mid_value=50, mid_color="FFEB84",
                                   end_type="max", end_color="63BE7B"))
        if ws.title == "buildup_summary":
            for idx, name in enumerate(headers, start=1):
                if isinstance(name, str) and name.endswith("_div%"):
                    col = get_column_letter(idx)
                    ws.conditional_formatting.add(
                        f"{col}2:{col}{ws.max_row}",
                        ColorScaleRule(start_type="min", start_color="63BE7B",
                                       mid_type="percentile", mid_value=50,
                                       mid_color="FFEB84", end_type="max",
                                       end_color="F8696B"))
        # one scale per metric row on the single-stock panels (master: E4:XFD4)
        if ws.title in ("stock_specific_data", "normalized_data") and ws.max_column > 2:
            for r in range(2, ws.max_row + 1):
                ws.conditional_formatting.add(
                    f"B{r}:{get_column_letter(ws.max_column)}{r}",
                    ColorScaleRule(start_type="min", start_color="F8696B",
                                   mid_type="percentile", mid_value=50, mid_color="FFEB84",
                                   end_type="max", end_color="63BE7B"))

    # line chart of the normalised panel, the master's "Plotted Data"
    if "normalized_data" in wb.sheetnames:
        ws = wb["normalized_data"]
        if ws.max_row > 1 and ws.max_column > 2:
            ch = LineChart()
            ch.title = "Normalised (0-1) - one stock, 20 sessions"
            ch.height, ch.width = 9, 24
            data = Reference(ws, min_col=2, max_col=ws.max_column,
                             min_row=1, max_row=ws.max_row)
            cats = Reference(ws, min_col=2, max_col=ws.max_column, min_row=1, max_row=1)
            ch.add_data(data, from_rows=True, titles_from_data=True)
            ch.set_categories(cats)
            ws.add_chart(ch, f"A{ws.max_row + 3}")
    wb.save(path)
    out["_focus"] = focus


def write_png(sig: pd.DataFrame, path: str, max_rows: int = 60):
    """The buildup grid as a phone-friendly image: last 5 sessions."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.colors import ListedColormap
    except Exception:                                   # noqa: BLE001
        return None

    cols = [c for c in sig.columns if not c.endswith("_div%")][-5:]
    if not cols:
        return None
    grid = sig[cols]
    keep = grid.replace("", pd.NA).dropna(how="all").index
    if len(keep) == 0:
        return None
    keep = list(keep)[:max_rows]
    codes = {"LB": 2, "SB": 1, "SC": 0, "LU": 3, "": 4}
    data = grid.loc[keep].map(lambda v: codes.get(v, 4)).values
    cmap = ListedColormap(["#BDD7EE", "#FFC7CE", "#C6EFCE", "#FFE699", "#FFFFFF"])

    fig, ax = plt.subplots(figsize=(5.2, max(3, 0.16 * len(keep) + 1.2)))
    ax.imshow(data, cmap=cmap, vmin=0, vmax=4, aspect="auto")
    ax.set_xticks(range(len(cols)))
    ax.set_xticklabels(cols, fontsize=7)
    ax.set_yticks(range(len(keep)))
    ax.set_yticklabels(keep, fontsize=5)
    ax.set_title("Buildup: LB green | SB red | SC blue | LU yellow", fontsize=7)
    for spine in ax.spines.values():
        spine.set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=160)
    plt.close(fig)
    return path


def summary_text(sig: pd.DataFrame, div: pd.DataFrame, dates) -> str:
    cols = [c for c in sig.columns if not c.endswith("_div%")]
    last = cols[-1]

    s = sig[last]
    lines = [f"NSE F&O buildup - {dates[-1]:%d-%b-%Y}  (D{len(cols)} of {len(cols)})", ""]
    for code in ("LB", "SB", "SC", "LU"):
        names = list(s[s == code].index)
        lines.append(f"{code} {SIGNAL_MEANING[code]}: {len(names)}")
        if names:
            lines.append("   " + ", ".join(names[:25]) + (" ..." if len(names) > 25 else ""))
        lines.append("")
    d = div[f"{last}_div%"].dropna().sort_values(ascending=False).head(8)
    if len(d):
        lines.append("Widest futures-vs-spot gaps: "
                     + ", ".join(f"{k} {v:.2f}%" for k, v in d.items()))
    return "\n".join(lines)


def inference_text(cash_sheet, fut_sheet, op_sheet, sig, cols) -> str:
    """The reading of the sheets: biggest OI moves, delivery, PCR extremes."""
    last = cols[-1]
    lines = []
    oi_col, px_col = f"OIchg%_{last}", f"PxChg%_{last}"
    if oi_col in fut_sheet.columns:
        oi = fut_sheet[oi_col].dropna()
        big = oi.reindex(oi.abs().sort_values(ascending=False).index).head(6)
        parts = []
        for sym, v in big.items():
            s = sig.get(last, pd.Series(dtype=object)).get(sym, "")
            px = fut_sheet[px_col].get(sym) if px_col in fut_sheet.columns else None
            parts.append(f"{sym} {v:+.1f}%" + (f" px{px:+.1f}%" if pd.notna(px) else "")
                         + (f" {s}" if s else ""))
        lines.append("Largest OI moves: " + "; ".join(parts))

    dv = cash_sheet.get(f"Deliv%_{last}")
    if dv is not None:
        top = dv.dropna().sort_values(ascending=False).head(5)
        lines.append("Highest delivery: " + ", ".join(f"{k} {v:.0f}%" for k, v in top.items()))

    ind_col = cash_sheet.get("industry")
    sig_last = sig.get(last) if last in sig.columns else None
    if ind_col is not None and sig_last is not None:
        hit = sig_last[sig_last != ""]
        if len(hit):
            sectors = ind_col.reindex(hit.index).dropna()
            counts = sectors[sectors != ""].value_counts().head(3)
            if len(counts):
                lines.append("Sector read: " + ", ".join(
                    f"{k} {v}" for k, v in counts.items()))

    pcr = op_sheet.get(f"PCR_{last}")
    if pcr is not None:
        p = pcr.dropna()
        if len(p):
            lines.append(f"PCR: lowest {p.idxmin()} {p.min():.2f}, "
                         f"highest {p.idxmax()} {p.max():.2f}, "
                         f"median {p.median():.2f}")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# telegram
# --------------------------------------------------------------------------- #
def tg_check(r, what: str) -> dict:
    """A 401/400 from Telegram must fail the run, not pass silently."""
    try:
        body = r.json()
    except Exception:                                   # noqa: BLE001
        body = {}
    if r.status_code != 200 or not body.get("ok"):
        raise RuntimeError(f"telegram {what} failed: HTTP {r.status_code} "
                           f"{body.get('description', r.text[:200])}")
    return body


def tg_send(token: str, chat: str, text: str) -> int:
    sent = 0
    for chunk in [text[i:i + 3900] for i in range(0, len(text), 3900)]:
        r = requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
                          data={"chat_id": chat, "text": chunk}, timeout=60)
        tg_check(r, "sendMessage")
        sent += 1
    return sent


def tg_doc(token: str, chat: str, path: str, caption: str = "") -> int:
    with open(path, "rb") as fh:
        r = requests.post(f"https://api.telegram.org/bot{token}/sendDocument",
                          data={"chat_id": chat, "caption": caption[:1000]},
                          files={"document": (os.path.basename(path), fh)}, timeout=180)
    tg_check(r, "sendDocument")
    return 1


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="last session, YYYY-MM-DD (default: latest published)")
    ap.add_argument("--days", type=int, default=20, help="window length (default 20)")
    ap.add_argument("--out", default="out", help="output directory")
    ap.add_argument("--telegram", action="store_true", help="deliver to Telegram")
    ap.add_argument("--show-days", type=int, default=8,
                    help="how many of the 20 sessions stay visible; older ones are "
                         "kept in the sheet but hidden (default 8, like the master)")
    ap.add_argument("--full-layout", action="store_true",
                    help="show every column instead of the master's hidden-column layout")
    ap.add_argument("--raw", action="store_true",
                    help="also include cash/futures/options source rows for the "
                         "latest session, as the master's hidden data sheets")
    ap.add_argument("--csv", action="store_true",
                    help="also send every sheet as a CSV (off by default: the "
                         "workbook already holds all of them)")
    ap.add_argument("--wait", type=int, default=0,
                    help="minutes to keep retrying if today's file is not up yet")
    ap.add_argument("--keep-cache", action="store_true")
    args = ap.parse_args()

    end = dt.date.fromisoformat(args.date) if args.date else dt.date.today()
    os.makedirs(args.out, exist_ok=True)

    print(f"resolving last {args.days} trading sessions up to {end} ...")
    dates = trading_dates(end, args.days)
    cash_up = bool(cached(f"sec_{end:%Y%m%d}.csv", sec_url(end)))
    if args.wait and end.weekday() < 5 and cash_up and dates and dates[-1] < end:
        deadline = time.time() + args.wait * 60
        print(f"  latest session is {dates[-1]}, waiting up to {args.wait} min for {end} ...")
        while time.time() < deadline and dates[-1] < end:
            time.sleep(120)
            dates = trading_dates(end, args.days)
    if len(dates) < 5:
        sys.exit("could not resolve enough sessions - NSE archive unreachable?")
    print(f"  {dates[0]} .. {dates[-1]}  ({len(dates)} sessions)")

    print("downloading + parsing ...")
    cash, fut, opt = build(dates)
    print(f"  cash {len(cash):,} rows | futures {len(fut):,} | options {len(opt):,}")

    print("computing sheets ...")
    out = sheets(cash, fut, opt, dates, THRESHOLD, raw=args.raw)

    xlsx = os.path.join(args.out, f"fno_eod_{dates[-1]:%Y%m%d}.xlsx")
    write_xlsx(out, xlsx, show_days=args.show_days, master_layout=not args.full_layout)
    sig = out["buildup_summary"]
    div_cols = [c for c in sig.columns if c.endswith("_div%")]
    div = sig[div_cols]
    png = write_png(sig, os.path.join(args.out, f"buildup_{dates[-1]:%Y%m%d}.png"))

    files = [xlsx] + ([png] if png else [])
    csvs = []
    for name, df in out.items():
        if name.startswith("_"):
            continue
        p = os.path.join(args.out, f"{name}_{dates[-1]:%Y%m%d}.csv")
        df.to_csv(p)
        csvs.append(p)
    if args.csv:
        files += csvs

    text = summary_text(sig.drop(columns=div_cols), div, dates)
    extra = inference_text(out.get("cash_alys_fno", pd.DataFrame()),
                           out.get("fut_alys_data", pd.DataFrame()),
                           out.get("op_alys_data", pd.DataFrame()),
                           sig.drop(columns=div_cols), [c for c in sig.columns if not c.endswith("_div%")])
    if extra:
        text = text + "\n" + extra
    with open(os.path.join(args.out, "summary.txt"), "w") as fh:
        fh.write(text)
    print("\n" + text)
    total = sum(os.path.getsize(f) for f in files)
    print(f"\n{len(files)} files to send, {total/1024:.0f} KB "
          f"(+{len(csvs)} CSVs kept on disk) -> {args.out}/")

    if args.telegram:
        token = os.environ.get("TELEGRAM_BOT_TOKEN") or os.environ.get("TELEGRAM_TOKEN")
        chat = os.environ.get("TELEGRAM_CHAT_ID")
        if not token or not chat:
            sys.exit("TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID not set")
        msgs = tg_send(token, chat, text)
        docs = sum(tg_doc(token, chat, f) for f in files)
        print(f"sent to Telegram: {msgs} message(s), {docs} document(s), all ok")

    if not args.keep_cache:
        shutil.rmtree(CACHE, ignore_errors=True)


if __name__ == "__main__":
    main()
