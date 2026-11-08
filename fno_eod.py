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
import re
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
# exact fills lifted from the master's conditional formats (dxf 5,6,7,8)
SIGNAL_COLORS = {"LB": "00B050", "SB": "FF0000", "SC": "92D050", "LU": "ED7D31"}
# theme indices in the master resolve to accent2 / accent6 / accent4, not the
# blues I assumed: LU is orange, and the two highlights are pale green and gold
HILITE_ABOVE = "E2F0D9"     # theme accent6 (70AD47), tint 0.80 - above average
HILITE_BOTH = "FFE699"      # theme accent4 (FFC000), tint 0.60 - both above
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
        # probe every calendar day, weekends included: muhurat sessions fall on
        # a Sunday and publish a normal bhavcopy
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
    # parked last so the four blocks land on the master's F / AC / BA / BY
    ca["industry"] = [ind.get(sym, "") for sym in ca.index]
    # assigned after the futures universe is known (see below)

    # ---------------- futures ---------------------------------------------
    fut = fut[fut["date"].isin(lab)].copy()
    fut["day"] = fut["date"].map(lab)
    fut = fut.dropna(subset=["OpnIntrst"]).sort_values(["TckrSymb", "XpryDt", "date"])
    # fractions, not percent numbers: the master stores 0.0402 and formats it
    # as 4.02% (its cash sheet is the opposite - 38.24 means 38.24%)
    fut["px_chg%"] = (fut.groupby(["TckrSymb", "XpryDt"], group_keys=False)["SttlmPric"]
                         .pct_change(fill_method=None))
    # the master pins ONE expiry per symbol for the whole window - the nearest
    # expiry as of the last session - rather than rolling day by day
    last_day = cols[-1]
    pinned = (fut[fut["day"] == last_day].groupby("TckrSymb")["XpryDt"].min()
                 .rename("pin"))
    fut = fut.merge(pinned, left_on="TckrSymb", right_index=True, how="left")
    near = fut[fut["XpryDt"] == fut["pin"]]
    price = near.pivot_table(index="TckrSymb", columns="day", values="SttlmPric", aggfunc="last").reindex(columns=cols)
    px_pct = near.pivot_table(index="TckrSymb", columns="day", values="px_chg%", aggfunc="last").reindex(columns=cols)
    lot = near.pivot_table(index="TckrSymb", columns="day", values="NewBrdLotQty", aggfunc="last").reindex(columns=cols)

    agg = fut.groupby(["TckrSymb", "day"]).agg(
        oi=("OpnIntrst", "sum"), coi=("ChngInOpnIntrst", "sum"),
        trades=("TtlNbOfTxsExctd", "sum"), vol=("TtlTradgVol", "sum")).reset_index()
    agg["oi_pct"] = agg["coi"] / (agg["oi"] - agg["coi"]).replace(0, pd.NA)
    # the master's ABS(CHG_IN_OI / CONTRACTS): absolute, per contract traded
    agg["ct_per_trade"] = agg["coi"].abs() / agg["vol"].replace(0, pd.NA)
    oi = agg.pivot_table(index="TckrSymb", columns="day", values="oi", aggfunc="last").reindex(columns=cols)
    oi_pct = agg.pivot_table(index="TckrSymb", columns="day", values="oi_pct", aggfunc="last").reindex(columns=cols)
    per_trade = agg.pivot_table(index="TckrSymb", columns="day", values="ct_per_trade",
                                aggfunc="last").reindex(columns=cols)

    # column order copied from the master: price, OI % change, futures TQ/NT
    # (with its average), cumulative OI, then price % change
    fa = pd.DataFrame(index=price.index)
    fa["min"] = price.min(axis=1)
    fa["max"] = price.max(axis=1)
    for c in cols:
        fa[f"Settle_{c}"] = price[c]
    for c in cols:
        fa[f"OIchg%_{c}"] = oi_pct[c]
    for c in cols:
        fa[f"OIperTrade_{c}"] = per_trade[c]
    fa["OIperTrade_20D_avg"] = per_trade.mean(axis=1)      # the master's BP column
    for c in cols:
        fa[f"OI_{c}"] = oi[c]
    fa["lot"] = lot[cols[-1]]        # parked last so the five blocks land on
                                     # the master's F / AA / AV / BR / CN
    out["fut_alys_data"] = _round_frame(fa).sort_index().rename_axis("SYMBOL")
    # the master's cash sheet lists the F&O universe, not all ~2,800 cash symbols
    out["cash_alys_fno"] = ca.loc[ca.index.intersection(price.index)].round(2).sort_index()

    # ---------------- options ---------------------------------------------
    op = opt[opt["date"].isin(lab)].copy()
    op["day"] = op["date"].map(lab)
    # the master's options block is the CHANGE in OI, in lots, signed -
    # not the OI level
    op["oi_lots"] = (pd.to_numeric(op["ChngInOpnIntrst"], errors="coerce")
                     / op["lot"].replace(0, pd.NA))
    op = op.dropna(subset=["oi_lots"])
    ce = pe = None
    if not op.empty:
        # every expiry summed - the master's SUMIFS filters symbol, date and
        # option type only, never expiry
        ce = (op[op["OptnTp"] == "CE"].pivot_table(index="TckrSymb", columns="day",
                                                   values="oi_lots", aggfunc="sum")
              .reindex(columns=cols))
        pe = (op[op["OptnTp"] == "PE"].pivot_table(index="TckrSymb", columns="day",
                                                   values="oi_lots", aggfunc="sum")
              .reindex(columns=cols))
        oa = pd.DataFrame(index=ce.index.union(pe.index))
        # CE and PE side by side for each day, the master's layout - not two
        # separate 20-column blocks
        for c in cols:
            oa[f"CE_lots_{c}"] = ce[c]
            oa[f"PE_lots_{c}"] = pe[c]
        oa = oa.round(0).astype("Int64")     # real integers, not rounded floats
        out["op_alys_data"] = oa.sort_index().rename_axis("SYMBOL")

    # ---------------- buildup + divergence --------------------------------
    idx = price.index.intersection(oi_pct.index)
    sig = pd.DataFrame("", index=idx, columns=cols, dtype=object)
    for c in cols:
        p, o = px_pct.get(c), oi_pct.get(c)
        if p is None or o is None:
            continue
        s = pd.Series("", index=idx, dtype=object)
        up, dn = p > 0, p < 0
        s[up & (o > threshold)] = "LB"
        s[dn & (o > threshold)] = "SB"
        s[up & (o < -threshold)] = "SC"
        s[dn & (o < -threshold)] = "LU"
        sig[c] = s

    div = pd.DataFrame(index=idx, columns=cols, dtype=float)
    for c in cols:
        spot = cash_close.reindex(idx).get(c)
        if spot is not None:
            div[c] = ((price.get(c) - spot) / spot).abs()
    div.columns = [f"{c}_div%" for c in cols]
    out["buildup_summary"] = _round_frame(sig.join(div)).sort_index().rename_axis("SYMBOL")

    # ---------------- one-stock panel + normalised twin -------------------
    focus = pick_focus(sig, oi_pct, cols)

    metrics = [
        ("Date", pd.Series({c: d.strftime("%d.%m.%y") for c, d in zip(cols, dates)})),
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
        metrics += [("CE OI lots", ce.loc[focus]), ("PE OI lots", pe.loc[focus])]
    if focus in sig.index:
        metrics += [("Signal", sig.loc[focus])]

    # every symbol x metric x day, so the sheet's dropdown can drive INDEX/MATCH
    src = {"Cash Price": cash_close, "Volume": qty, "Delivery%": dv,
           "Cash Qty/Trade": qt, "Futures Price": price, "Cumulative OI": oi,
           "OI chg %": oi_pct, "FNO TQ/NT": per_trade}
    if ce is not None:
        src["CE OI lots"] = ce
        src["PE OI lots"] = pe
    syms = sorted(price.index)      # F&O universe only - keeps the sheet small
    data = {}
    for mname, frame in src.items():
        if mname in ("CE OI lots", "PE OI lots"):
            frame = frame.round(0).astype("Int64")
        for c in cols:
            data[f"{mname}|{c}"] = (frame.reindex(syms)[c] if c in frame.columns
                                    else pd.Series(pd.NA, index=syms))
    wide = pd.DataFrame(data, index=syms)
    out["panel_data"] = _round_frame(wide)

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
                "lot", "OIperTrade_20D_avg", "min", "max", "industry"}

SHORT = {"Close": "Px", "Volume": "Vol", "Deliv%": "Del%", "Qty/Trade": "Q/T",
         "Settle": "Fut", "PxChg%": "Chg%", "OI": "OI", "OIchg%": "OI%",
         "OIperTrade": "OI/T", "CE_lots": "CE", "PE_lots": "PE", "PCR": "PCR",
         "Deliv": "DelQ", "Vol": "Vol", "SIGNAL": "SIG", "DIVERGENCE": "DIV"}
SCALAR_SHORT = {"min_price": "min", "max_price": "max", "min_vol": "min", "max_vol": "max",
                "min_deliv": "min", "max_deliv": "max", "min_qtytrade": "min",
                "max_qtytrade": "max", "deliv%_20D": "Del%20", "qty/trade_20D": "Q/T20",
                "lot": "lot", "OIperTrade_20D_avg": "OI/T20", "min": "min", "max": "max",
                "industry": "industry"}
SEP_RGB = "002060"


def _cellval(v):
    """pandas NA / numpy scalars -> plain Excel values."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v.item() if hasattr(v, "item") else v


def _round_frame(df, frac_prefixes=("OIchg%", "PxChg%"), frac_suffixes=("_div%",)):
    """Fractions must not go through round(2) - 0.0074 would collapse to 0.01."""
    out = df.copy()
    for c in out.columns:
        if out[c].dtype.kind not in "if":      # leave the signal strings alone
            continue
        frac = (str(c).startswith(frac_prefixes) or str(c).endswith(frac_suffixes))
        out[c] = out[c].round(6 if frac else 2)
    return out


def _day_index(name: str):
    """D7 out of 'Close_D7' / 'OIchg%_D7' / 'D7' / 'D7_div%'."""
    for tok in str(name).replace("_div%", "").split("_"):
        if tok.startswith("D") and tok[1:].isdigit():
            return int(tok[1:])
    return None


def _plan(df):
    """Contiguous groups, in the frame's own order: scalars stay where they sit,
    day columns stay together per block. This reproduces the master's interleaved
    layout - min/max, then its 20 days, then the next block."""
    groups = []
    for c in df.columns:
        d = _day_index(c)
        if d is None:
            key, label = f"__{c}", ""
        elif re.fullmatch(r"D\d+", str(c)):
            key, label = "SIGNAL", "SIGNAL"
        elif str(c).endswith("_div%"):
            key, label = "DIVERGENCE", "DIVERGENCE"
        else:
            key = re.sub(r"_D\d+$", "", str(c))
            label = key
        if groups and groups[-1][0] == key:
            groups[-1][1].append(c)
        else:
            groups.append([key, [c], label])
    return [(label, cols) for _, cols, label in groups]


def write_workbook(out: dict, path: str, dates, show_days: int = 10,
                   master_layout: bool = True):
    """Write the workbook in the master's shape: a blue separator column between
    datasets, a group header, D-labels, a row of real dates, colour scales per
    row inside each dataset, and all sheets at 61% zoom."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
    from openpyxl.styles import Border, Side
    from openpyxl.utils import get_column_letter

    out = {k: v for k, v in out.items() if not k.startswith("_")}
    order = ["cash_alys_fno", "fut_alys_data", "op_alys_data",
             "buildup_summary", "stock_specific_data", "normalized_data",
             "panel_data"]
    HIDDEN = {"normalized_data", "panel_data", "chart_data"}
    wb = Workbook()
    wb.remove(wb.active)
    date_lbl = {f"D{i+1}": d.strftime("%d.%m.%y") for i, d in enumerate(dates)}

    for name in [n for n in order if n in out] + [n for n in out if n not in order]:
        df = out[name]
        ws = wb.create_sheet(name[:31])
        ws.sheet_view.zoomScale = 61
        ws.sheet_view.zoomScaleNormal = 61
        grid = name in ("cash_alys_fno", "fut_alys_data", "op_alys_data",
                        "nifty_delivery", "buildup_summary")
        n_rows = len(df)
        col = 2                       # column A is the master's hidden spacer
        hide_cols, block_spans, colmap = [1], [], {}

        # index column (grid sheets only - the plain sheets label rows in A)
        if grid:
            ws.cell(row=2, column=col, value=df.index.name or "SYMBOL")
            for i, v in enumerate(df.index, start=4):
                ws.cell(row=i, column=col, value=str(v))
            ws.column_dimensions[get_column_letter(col)].width = 15
            col += 1

        if grid:
            single_header = False
            if name == "op_alys_data":
                cols_list = list(df.columns)
                groups = [(f"D{i+1}", cols_list[i*2:i*2+2]) for i in range(len(cols_list)//2)]
                single_header = True
            else:
                groups = _plan(df)
            for gname, cols in groups:
                if gname:                       # day-block: give it a blue divider
                    ws.column_dimensions[get_column_letter(col)].width = 2.2
                    for r in range(1, n_rows + 4):
                        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=SEP_RGB)
                    col += 1
                first = col
                for c in cols:
                    d = _day_index(c)
                    if d is not None:
                        pair_pe = str(c).startswith("PE_lots")
                        if name == "op_alys_data":
                            # every column says what it is; the day and date sit
                            # merged across the CE+PE pair on the line below
                            ws.cell(row=2, column=col, value="PE" if pair_pe else "CE")
                            ws.cell(row=3, column=col,
                                    value="" if pair_pe else f"D{d}  {date_lbl.get(f'D{d}', '')}")
                            if pair_pe:
                                ws.merge_cells(start_row=3, start_column=col - 1,
                                               end_row=3, end_column=col)
                        else:
                            ws.cell(row=2, column=col, value="" if pair_pe else f"D{d}")
                            ws.cell(row=3, column=col,
                                    value="" if pair_pe else date_lbl.get(f"D{d}", ""))
                        # only these two sheets fold the older sessions away;
                        # options, buildup and nifty_delivery show all 20 days
                        if master_layout and d <= 20 - show_days and \
                           name in ("cash_alys_fno", "fut_alys_data"):
                            hide_cols.append(col)
                    else:
                        ws.cell(row=2, column=col, value=SCALAR_SHORT.get(c, str(c)[:8]))
                        if master_layout and c in SUPPORT_COLS:
                            hide_cols.append(col)
                    colmap[c] = col
                    pct = (str(c).startswith(("OIchg%", "PxChg%"))
                           or str(c).endswith("_div%"))
                    whole = (str(c).startswith("OIperTrade")
                             or str(c).startswith(("CE_lots", "PE_lots")))
                    for i, v in enumerate(df[c].tolist(), start=4):
                        cell = ws.cell(row=i, column=col, value=_cellval(v))
                        if isinstance(v, (int, float)):
                            cell.number_format = ("0.00%" if pct else
                                                  "#,##0" if whole else "#,##0.##")
                    ws.column_dimensions[get_column_letter(col)].width = 8
                    col += 1
                last = col - 1
                if last >= first:
                    if not single_header:
                        ws.merge_cells(start_row=1, start_column=first, end_row=1, end_column=last)
                        ws.cell(row=1, column=first, value=gname)
                    block_spans.append((gname, first, last))
            if name == "op_alys_data" and block_spans:
                f0, l0 = block_spans[0][1], block_spans[-1][2]
                ws.merge_cells(start_row=1, start_column=f0, end_row=1, end_column=l0)
                ws.cell(row=1, column=f0, value="CE / PE OI change (lots)")

            # per-row colour scale inside each dataset - but only on the blocks
            # the master scales. Delivery % and qty/trade are NOT colour-scaled
            # there: they carry the blue/grey above-average highlight instead,
            # and a scale on top would drown it out.
            SCALED = {"cash_alys_fno": {"Close", "Volume"},
                      "fut_alys_data": {"Settle", "OIchg%", "OIperTrade", "OI"},
                      "buildup_summary": {"DIVERGENCE"},
                      "op_alys_data": set()}      # CE/PE get their own split scales
            allowed = SCALED.get(name)
            for gname, first, last in block_spans:
                if not gname or gname == "SIGNAL":
                    continue
                if allowed is not None and gname not in allowed:
                    continue
                for r in range(4, n_rows + 4):
                    ws.conditional_formatting.add(
                        f"{get_column_letter(first)}{r}:{get_column_letter(last)}{r}",
                        ColorScaleRule(start_type="min", start_color="F8696B",
                                       mid_type="percentile", mid_value=50,
                                       mid_color="FFEB84", end_type="max",
                                       end_color="63BE7B"))
            # delivery % and qty/trade: the master's two-tier highlight
            if name == "cash_alys_fno" and "deliv%_20D" in colmap:
                agd = get_column_letter(colmap["deliv%_20D"])
                agq = get_column_letter(colmap.get("qty/trade_20D", colmap["deliv%_20D"]))
                first_row, last_row = 4, n_rows + 3
                for c, ci in colmap.items():
                    if not re.search(r"_D\d+$", str(c)):
                        continue
                    L = get_column_letter(ci)
                    if str(c).startswith("Deliv%"):
                        q = colmap.get(re.sub(r"^Deliv%", "Qty/Trade", c))
                        if q:
                            Q = get_column_letter(q)
                            ws.conditional_formatting.add(
                                f"{L}{first_row}:{L}{last_row}",
                                FormulaRule(formula=[f"AND({L}{first_row}>${agd}{first_row},"
                                                     f"{Q}{first_row}>${agq}{first_row})"],
                                            fill=PatternFill("solid", bgColor=HILITE_BOTH)))
                        ws.conditional_formatting.add(
                            f"{L}{first_row}:{L}{last_row}",
                            FormulaRule(formula=[f"{L}{first_row}>${agd}{first_row}"],
                                        fill=PatternFill("solid", bgColor=HILITE_ABOVE)))
                    elif str(c).startswith("Qty/Trade"):
                        ws.conditional_formatting.add(
                            f"{L}{first_row}:{L}{last_row}",
                            FormulaRule(formula=[f"{L}{first_row}>${agq}{first_row}"],
                                        fill=PatternFill("solid", bgColor=HILITE_ABOVE)))

            # futures: thin box where OI per contract beats its own average
            if name == "fut_alys_data" and "OIperTrade_20D_avg" in colmap:
                ag = get_column_letter(colmap["OIperTrade_20D_avg"])
                pt = [get_column_letter(ci) for c, ci in sorted(colmap.items(), key=lambda kv: kv[1])
                      if str(c).startswith("OIperTrade_") and not str(c).endswith("_avg")]
                oi = [get_column_letter(ci) for c, ci in sorted(colmap.items(), key=lambda kv: kv[1])
                      if str(c).startswith("OIchg%_")]
                if pt and oi:
                    ws.conditional_formatting.add(
                        f"{oi[0]}4:{oi[-1]}{n_rows + 3}",
                        FormulaRule(formula=[f"${pt[0]}4>${ag}4"],
                                    border=Border(left=Side(style="thin"), right=Side(style="thin"),
                                                  top=Side(style="thin"), bottom=Side(style="thin"))))

            # options: CE and PE are separate datasets - scale each stock's CE
            # cells together and its PE cells together, never across the pair
            if name == "op_alys_data":
                ce_cols = [get_column_letter(ci) for c, ci in colmap.items()
                           if str(c).startswith("CE_lots")]
                pe_cols = [get_column_letter(ci) for c, ci in colmap.items()
                           if str(c).startswith("PE_lots")]
                for r in range(4, n_rows + 4):
                    for group in (ce_cols, pe_cols):
                        if not group:
                            continue
                        refs = " ".join(f"{g}{r}" for g in group)
                        ws.conditional_formatting.add(
                            refs, ColorScaleRule(start_type="min", start_color="F8696B",
                                                 mid_type="percentile", mid_value=50,
                                                 mid_color="FFEB84", end_type="max",
                                                 end_color="63BE7B"))

            # signal fills
            if name == "buildup_summary":
                for gname, first, last in block_spans:
                    if gname != "SIGNAL":
                        continue
                    c0 = get_column_letter(first)
                    for code, rgb in SIGNAL_COLORS.items():
                        ws.conditional_formatting.add(
                            f"{c0}4:{get_column_letter(last)}{n_rows + 3}",
                            FormulaRule(formula=[f'{c0}4="{code}"'],
                                        fill=PatternFill("solid", bgColor=rgb)))
        elif name == "stock_specific_data" and "panel_data" in out:
            wide = out["panel_data"]
            syms = list(wide.index)
            ws.column_dimensions["A"].width = 22
            ws.cell(row=1, column=1, value="Stock")
            ws.cell(row=1, column=2, value=str(df.index.name or syms[0]))
            for j, c in enumerate(df.columns, start=2):
                ws.cell(row=2, column=j, value=c)
                ws.cell(row=3, column=j, value=date_lbl.get(c, ""))
                ws.column_dimensions[get_column_letter(j)].width = 11
            metrics = list(df.index)
            # panel_data: symbols in column A from row 4, data from column B
            first_sym_row, last_sym_row = 4, 4 + len(syms) - 1
            last_data_col = get_column_letter(2 + len(wide.columns) - 1)
            cols_index = {c: i + 1 for i, c in enumerate(wide.columns)}
            for i, m in enumerate(metrics, start=4):
                ws.cell(row=i, column=1, value=str(m))
                for j, c in enumerate(df.columns, start=2):
                    if m == "Date":
                        ws.cell(row=i, column=j, value=df.loc[m, c])
                        continue
                    key = f"{m}|{c}"
                    if key not in cols_index:
                        continue
                    k = cols_index[key]
                    ws.cell(row=i, column=j,
                            value=f"=IFERROR(INDEX(panel_data!$B${first_sym_row}:${last_data_col}${last_sym_row},"
                                  f"MATCH($B$1,panel_data!$A${first_sym_row}:$A${last_sym_row},0),{k}),\"\")")
                    ws.cell(row=i, column=j).number_format = (
                        "0.00%" if str(m) == "OI chg %" else
                        "#,##0" if "OI lots" in str(m) else "#,##0.##")
            from openpyxl.worksheet.datavalidation import DataValidation
            dv = DataValidation(type="list",
                                formula1=f"=panel_data!$A$4:$A${4 + len(syms) - 1}",
                                allow_blank=False, showDropDown=False)
            ws.add_data_validation(dv)
            dv.add("B1")
            hide_cols = []

        else:
            # plain sheets: row labels in A, single header row, no blocks
            for i, v in enumerate(df.index, start=4):
                ws.cell(row=i, column=1, value=str(v))
            ws.column_dimensions["A"].width = 20
            hide_cols = []
            for c in df.columns:
                ws.cell(row=2, column=col, value=SHORT.get(str(c), str(c))[:14])
                for i, v in enumerate(df[c].tolist(), start=4):
                    cell = ws.cell(row=i, column=col, value=_cellval(v))
                    if isinstance(v, float):
                        cell.number_format = "#,##0.##"
                ws.column_dimensions[get_column_letter(col)].width = 12
                col += 1
            if name in ("stock_specific_data", "normalized_data"):
                for r in range(4, n_rows + 4):
                    ws.conditional_formatting.add(
                        f"B{r}:{get_column_letter(col-1)}{r}",
                        ColorScaleRule(start_type="min", start_color="F8696B",
                                       mid_type="percentile", mid_value=50,
                                       mid_color="FFEB84", end_type="max",
                                       end_color="63BE7B"))

        for c in hide_cols:
            ws.column_dimensions[get_column_letter(c)].hidden = True
        for r in (1, 2, 3):
            ws.row_dimensions[r].height = 13
        for row in ws.iter_rows(min_row=1, max_row=3):
            for cell in row:
                cell.font = Font(bold=False)
                cell.alignment = Alignment(horizontal="center", vertical="center")
        if grid:
            ws.column_dimensions["A"].width = 2      # hidden spacer on grid sheets
            ws.freeze_panes = "C4"
        if master_layout and name in HIDDEN:
            ws.sheet_state = "hidden"

    # chart_data: same rows as the stock panel, offset into bands, but as
    # formulas - so picking another symbol moves the lines as well
    if "stock_specific_data" in wb.sheetnames:
        src_ws = wb["stock_specific_data"]
        cd = wb.create_sheet("chart_data")
        cd.sheet_view.zoomScale = 61
        metrics = [src_ws.cell(r, 1).value for r in range(4, src_ws.max_row + 1)]
        ncols = src_ws.max_column
        for j in range(2, ncols + 1):
            col = get_column_letter(j)
            cd.cell(row=2, column=j, value=src_ws.cell(2, j).value)
            cd.cell(row=3, column=j, value=src_ws.cell(3, j).value)
        for i, m in enumerate(metrics):
            cd.cell(row=4 + i, column=1, value=m)
            lastc = get_column_letter(ncols)
            for j in range(2, ncols + 1):
                col = get_column_letter(j)
                r = 4 + i
                # normalise the row to 0-1 then lift it into its own band
                cd.cell(row=r, column=j,
                        value=(f"=IFERROR(('stock_specific_data'!{col}{r}"
                               f"-MIN('stock_specific_data'!$B{r}:${lastc}{r}))"
                               f"/(MAX('stock_specific_data'!$B{r}:${lastc}{r})"
                               f"-MIN('stock_specific_data'!$B{r}:${lastc}{r}))+{i},\"\")"))
        cd.column_dimensions["A"].width = 22
        cd.sheet_state = "hidden"

    if "stock_specific_data" in wb.sheetnames and "chart_data" in wb.sheetnames:
        ws = wb["stock_specific_data"]
        src = wb["chart_data"]
        if src.max_row >= 5 and src.max_column >= 3:
            from openpyxl.chart import LineChart, Reference
            from openpyxl.chart.marker import Marker
            ch = LineChart()
            ch.title = None                     # the master's chart has no title
            ch.height, ch.width = 16, 30
            # start the range at column A so the row label becomes the series
            # name - otherwise the legend shows the first value instead
            data = Reference(src, min_col=1, max_col=src.max_column,
                             min_row=2, max_row=src.max_row)
            cats = Reference(src, min_col=2, max_col=src.max_column, min_row=3, max_row=3)
            ch.add_data(data, from_rows=True, titles_from_data=True)
            ch.set_categories(cats)
            ch.legend.position = "b"            # legend under the plot
            ch.y_axis.majorGridlines = None     # no heavy horizontal rules
            ch.x_axis.majorGridlines = None
            for ser in ch.series:
                ser.marker = Marker(symbol="circle", size=4)
                ser.smooth = False
            ws.add_chart(ch, f"A{ws.max_row + 3}")

    if "normalized_data" in wb.sheetnames:
        ws = wb["normalized_data"]
        if ws.max_row > 4 and ws.max_column > 2:
            from openpyxl.chart import LineChart, Reference
            ch = LineChart()
            ch.title = "Normalised 0-1 - one stock, 20 sessions"
            ch.height, ch.width = 16, 30
            data = Reference(ws, min_col=2, max_col=ws.max_column, min_row=2, max_row=ws.max_row)
            ch.add_data(data, from_rows=True, titles_from_data=True)
            ws.add_chart(ch, f"A{ws.max_row + 3}")
    wb.save(path)


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
    lines = [f"NSE F&O buildup - {dates[-1]:%d.%m.%y}  (D{len(cols)} of {len(cols)})", ""]
    for code in ("LB", "SB", "SC", "LU"):
        names = list(s[s == code].index)
        lines.append(f"{code} {SIGNAL_MEANING[code]}: {len(names)}")
        if names:
            lines.append("   " + ", ".join(names[:25]) + (" ..." if len(names) > 25 else ""))
        lines.append("")
    d = div[f"{last}_div%"].dropna().sort_values(ascending=False).head(8)
    if len(d):
        lines.append("Widest futures-vs-spot gaps: "
                     + ", ".join(f"{k} {v*100:.2f}%" for k, v in d.items()))
    return "\n".join(lines)


def inference_text(cash_sheet, fut_sheet, op_sheet, sig, cols) -> str:
    """The reading of the sheets: biggest OI moves and delivery extremes."""
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
            parts.append(f"{sym} {v*100:+.1f}%" + (f" px{px*100:+.1f}%" if pd.notna(px) else "")
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
    ap.add_argument("--show-days", type=int, default=10,
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
    ap.add_argument("--skip-if-sent", action="store_true",
                    help="exit quietly when the latest session was already delivered "
                         "(lets the schedule run daily without resending on holidays)")
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
    if args.skip_if_sent:
        shard = os.path.join("data", "history", f"{dates[-1]}.csv.gz")
        if os.path.exists(shard):
            print(f"{dates[-1]} already delivered - nothing to do")
            return

    print("downloading + parsing ...")
    cash, fut, opt = build(dates)
    print(f"  cash {len(cash):,} rows | futures {len(fut):,} | options {len(opt):,}")

    print("computing sheets ...")
    out = sheets(cash, fut, opt, dates, THRESHOLD, raw=args.raw)

    xlsx = os.path.join(args.out, f"fno_eod_{dates[-1]:%Y%m%d}.xlsx")
    write_workbook(out, xlsx, dates, show_days=args.show_days,
                   master_layout=not args.full_layout)
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
