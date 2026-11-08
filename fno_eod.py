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
    return head.startswith("SYMBOL")


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
def sheets(cash, fut, opt, dates, threshold):
    """All output sheets. Every series is computed *within its own expiry* so
    rollover days do not produce fake OI-collapse signals."""
    lab = label(dates)
    cols = [f"D{i+1}" for i in range(len(dates))]
    out = {}

    # ---------------- cash: delivery %, qty per trade, min/max -------------
    cash = cash[cash["date"].isin(lab)].copy()
    cash["day"] = cash["date"].map(lab)
    px = cash.pivot_table(index="SYMBOL", columns="day", values="CLOSE_PRICE", aggfunc="last")
    dv = cash.pivot_table(index="SYMBOL", columns="day", values="DELIV_PER", aggfunc="last")
    qty = cash.pivot_table(index="SYMBOL", columns="day", values="TTL_TRD_QNTY", aggfunc="sum")
    trd = cash.pivot_table(index="SYMBOL", columns="day", values="NO_OF_TRADES", aggfunc="sum")
    px, dv, qty, trd = (df.reindex(columns=cols) for df in (px, dv, qty, trd))
    qt = qty / trd.replace(0, pd.NA)
    cash_close = px.ffill(axis=1)          # last traded close, carried forward

    ca = pd.DataFrame(index=px.index)
    ca["min"] = px.min(axis=1)
    ca["max"] = px.max(axis=1)
    ca["close"] = cash_close[cols[-1]]
    ca["deliv%_20DMA"] = dv.mean(axis=1)
    ca["qty/trade_20DMA"] = qt.mean(axis=1)
    ca["turnover_cr_20D"] = (qty.sum(axis=1) * ca["close"] / 1e7).round(2)
    out["cash_alys_fno"] = ca.round(2).sort_index()

    # ---------------- futures: settle, OI, OI change % ---------------------
    # price comes from the nearest series (its own day-over-day change), while
    # OI is aggregated across *all* series - that is roll-neutral, so expiry
    # week no longer looks like a mass unwind.
    fut = fut[fut["date"].isin(lab)].copy()
    fut["day"] = fut["date"].map(lab)
    fut = fut.dropna(subset=["OpnIntrst"]).sort_values(["TckrSymb", "XpryDt", "date"])
    fut["px_chg%"] = (fut.groupby(["TckrSymb", "XpryDt"], group_keys=False)["SttlmPric"]
                         .pct_change(fill_method=None) * 100)
    near = fut.loc[fut.groupby(["TckrSymb", "day"])["XpryDt"].idxmin()]
    price = near.pivot_table(index="TckrSymb", columns="day", values="SttlmPric",
                             aggfunc="last").reindex(columns=cols)
    px_pct = near.pivot_table(index="TckrSymb", columns="day", values="px_chg%",
                              aggfunc="last").reindex(columns=cols)
    lot = near.pivot_table(index="TckrSymb", columns="day", values="NewBrdLotQty",
                           aggfunc="last").reindex(columns=cols)

    agg = fut.groupby(["TckrSymb", "day"]).agg(
        oi=("OpnIntrst", "sum"), coi=("ChngInOpnIntrst", "sum"),
        trades=("TtlNbOfTxsExctd", "sum")).reset_index()
    agg["oi_pct"] = agg["coi"] / (agg["oi"] - agg["coi"]).replace(0, pd.NA) * 100
    oi = agg.pivot_table(index="TckrSymb", columns="day", values="oi",
                         aggfunc="last").reindex(columns=cols)
    oi_pct = agg.pivot_table(index="TckrSymb", columns="day", values="oi_pct",
                             aggfunc="last").reindex(columns=cols)
    # change in contracts per trade (workbook's futures "TQ/NT", unit-corrected)
    agg["ct_per_trade"] = agg["coi"] / agg["trades"].replace(0, pd.NA)
    per_trade = agg.pivot_table(index="TckrSymb", columns="day", values="ct_per_trade",
                                aggfunc="last").reindex(columns=cols)
    per_trade = per_trade.div(lot.replace(0, pd.NA))

    fa = pd.DataFrame(index=price.index)
    fa["min"] = price.min(axis=1)
    fa["max"] = price.max(axis=1)
    fa["settle"] = price[cols[-1]]
    fa["oi_units"] = oi[cols[-1]]
    fa["oi_chg%_last"] = oi_pct[cols[-1]]
    fa["oi_contracts_per_trade_20DMA"] = per_trade.mean(axis=1)
    fa["lot"] = lot[cols[-1]]
    out["fut_alys_data"] = fa.round(2).sort_index()

    # ---------------- options: CE / PE OI in lots -------------------------
    op = opt[opt["date"].isin(lab)].copy()
    op["day"] = op["date"].map(lab)
    op["oi_lots"] = pd.to_numeric(op["oi_lots"], errors="coerce")
    op = op.dropna(subset=["oi_lots"])
    if not op.empty:
        per_exp = (op.groupby(["TckrSymb", "day", "XpryDt", "OptnTp"])["oi_lots"]
                     .sum().reset_index())
        tot = (per_exp.groupby(["TckrSymb", "day", "XpryDt"])["oi_lots"].sum()
                      .reset_index())
        dom = tot.loc[tot.groupby(["TckrSymb", "day"])["oi_lots"].idxmax()]
        dom = dom.merge(per_exp, on=["TckrSymb", "day", "XpryDt"], suffixes=("_tot", ""))
        ce = (dom[dom["OptnTp"] == "CE"].pivot_table(index="TckrSymb", columns="day",
                                                     values="oi_lots", aggfunc="sum")
              .reindex(columns=cols))
        pe = (dom[dom["OptnTp"] == "PE"].pivot_table(index="TckrSymb", columns="day",
                                                     values="oi_lots", aggfunc="sum")
              .reindex(columns=cols))
        oa = pd.concat({"CE_lots": ce, "PE_lots": pe}, axis=1)
        oa = oa.reindex(columns=[f"{k}_D{i+1}" for i in range(len(cols))
                                 for k in ("CE_lots", "PE_lots")])
        oa["PE_by_CE_last"] = (pe[cols[-1]] / ce[cols[-1]].replace(0, pd.NA)).round(2)
        out["op_alys_data"] = oa.round(0).sort_index()

    # ---------------- buildup summary: LB / SB / SC / LU ------------------
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
        if spot is None:
            continue
        div[c] = ((price.get(c) - spot) / spot).abs() * 100
    div.columns = [f"{c}_div%" for c in cols]
    out["buildup_summary"] = sig.join(div.round(2)).sort_index()

    # ---------------- general: settings + date table ----------------------
    rows = [{"Days": c, "Date": d.strftime("%d-%b-%Y"), "ReportDate": d.strftime("%d%m%Y")}
            for c, d in zip(cols, dates)]
    rows += [
        {"Days": "", "Date": "", "ReportDate": ""},
        {"Days": "Threshold Buildup Change", "Date": threshold, "ReportDate": ""},
        {"Days": "Sessions", "Date": len(dates), "ReportDate": ""},
        {"Days": "Generated", "Date": dt.datetime.now().strftime("%d-%b-%Y %H:%M"),
         "ReportDate": ""},
    ]
    out["general"] = pd.DataFrame(rows)

    # ---------------- one-stock panel -------------------------------------
    focus = pick_focus(sig, oi_pct, cols)
    last = cols[-1]
    panel = [
        ("Cash Price", cash_close.get(last)),
        ("Delivery%", dv.get(last)),
        ("Qty/Trade", qt.get(last)),
        ("Futures Price", price.get(last)),
        ("OI (units)", oi.get(last)),
        ("OI chg %", oi_pct.get(last)),
        ("Signal", sig.get(last)),
    ]
    data = {}
    for name, series in panel:
        if series is not None and focus in series.index:
            data[name] = [series.get(focus)]
    out["stock_specific_data"] = pd.DataFrame(data, index=[focus]).T if data else pd.DataFrame()
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
def write_xlsx(out: dict, path: str):
    focus = out.pop("_focus", None)
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        for name, df in out.items():
            df.to_excel(xl, sheet_name=name[:31])
    # light styling: signal colours, freeze panes, widths
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
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
    out = sheets(cash, fut, opt, dates, THRESHOLD)

    xlsx = os.path.join(args.out, f"fno_eod_{dates[-1]:%Y%m%d}.xlsx")
    write_xlsx(out, xlsx)
    sig = out["buildup_summary"]
    div_cols = [c for c in sig.columns if c.endswith("_div%")]
    div = sig[div_cols]
    png = write_png(sig, os.path.join(args.out, f"buildup_{dates[-1]:%Y%m%d}.png"))

    files = [xlsx] + ([png] if png else [])
    for name, df in out.items():
        if name.startswith("_"):
            continue
        p = os.path.join(args.out, f"{name}_{dates[-1]:%Y%m%d}.csv")
        df.to_csv(p)
        files.append(p)

    text = summary_text(sig.drop(columns=div_cols), div, dates)
    with open(os.path.join(args.out, "summary.txt"), "w") as fh:
        fh.write(text)
    print("\n" + text)
    total = sum(os.path.getsize(f) for f in files)
    print(f"\n{len(files)} files, {total/1024:.0f} KB total -> {args.out}/")

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
