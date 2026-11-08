#!/usr/bin/env python3
"""
backfill.py - build and maintain a compact per-session history for fno_eod.

fno_eod.py answers "what do the last 20 sessions look like". This answers
"what did the last 16 months look like": one row per F&O symbol per session,
every metric the screener uses.

Storage shape matters: each session is written as its own small gzipped CSV
(about 25 KB), never rewritten afterwards. Git stores each shard once, so a
daily run adds ~25 KB to the repo instead of committing a new 9 MB blob every
day. Combine on demand when you actually want one file.

    python backfill.py --from 2024-07-08      # first run: full backfill
    python backfill.py --since-last           # daily: append newer sessions
    python backfill.py --combine              # data/history.csv.gz for analysis

NSE only publishes derivatives in the current UDiFF format from 2024-07-08, so
that is the earliest date backfillable with one consistent schema. Re-running
is safe: shards that already exist are skipped.
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import os
import sys

import pandas as pd

import fno_eod as fe

DEFAULT_START = dt.date(2024, 7, 8)          # first UDiFF session
HIST_DIR = "data/history"
COLS = ["date", "symbol", "close", "volume", "trades", "deliv_pct", "qty_per_trade",
        "fut_price", "fut_oi", "fut_oi_pct", "fut_volume", "lot",
        "ce_lots", "pe_lots", "pcr"]


def one_session(d: dt.date, all_symbols: bool = False) -> pd.DataFrame | None:
    """Per-symbol metrics for a single session, or None if not published."""
    sec = fe.cached(f"sec_{d:%Y%m%d}.csv", fe.sec_url(d))
    fo = fe.cached(f"fo_{d:%Y%m%d}.zip", fe.fo_url(d))
    if not sec or not fo:
        return None

    cash = fe.read_sec(sec).set_index("SYMBOL")
    out = pd.DataFrame(index=cash.index)
    out["close"] = cash["CLOSE_PRICE"]
    out["volume"] = cash["TTL_TRD_QNTY"]
    out["trades"] = cash["NO_OF_TRADES"]
    out["deliv_pct"] = cash["DELIV_QTY"] / cash["TTL_TRD_QNTY"].replace(0, pd.NA) * 100
    out["qty_per_trade"] = cash["TTL_TRD_QNTY"] / cash["NO_OF_TRADES"].replace(0, pd.NA)

    f = fe.read_fo(fo)

    fut = f[f["FinInstrmTp"] == "STF"].dropna(subset=["OpnIntrst"])
    if not fut.empty:
        near = fut.loc[fut.groupby("TckrSymb")["XpryDt"].idxmin()].set_index("TckrSymb")
        out["fut_price"] = near["SttlmPric"]
        out["lot"] = near["NewBrdLotQty"]
        agg = fut.groupby("TckrSymb").agg(oi=("OpnIntrst", "sum"),
                                          coi=("ChngInOpnIntrst", "sum"),
                                          vol=("TtlTradgVol", "sum"))
        out["fut_oi"] = agg["oi"]
        out["fut_volume"] = agg["vol"]
        out["fut_oi_pct"] = agg["coi"] / (agg["oi"] - agg["coi"]).replace(0, pd.NA) * 100

    opt = f[f["FinInstrmTp"] == "STO"].copy()
    if not opt.empty:
        opt["oi_lots"] = opt["OpnIntrst"] / opt["NewBrdLotQty"].replace(0, pd.NA)
        per_exp = (opt.groupby(["TckrSymb", "XpryDt", "OptnTp"])["oi_lots"].sum()
                      .reset_index())
        tot = per_exp.groupby(["TckrSymb", "XpryDt"])["oi_lots"].sum().reset_index()
        dom = tot.loc[tot.groupby("TckrSymb")["oi_lots"].idxmax()]
        dom = dom.merge(per_exp, on=["TckrSymb", "XpryDt"], suffixes=("_tot", ""))
        ce = dom[dom["OptnTp"] == "CE"].set_index("TckrSymb")["oi_lots"]
        pe = dom[dom["OptnTp"] == "PE"].set_index("TckrSymb")["oi_lots"]
        out["ce_lots"] = ce
        out["pe_lots"] = pe
        out["pcr"] = pe / ce.replace(0, pd.NA)

    out = out.dropna(subset=["close"])
    if not all_symbols:
        # F&O universe only: ~220 rows a session instead of ~2,100
        out = out[out["fut_price"].notna()]
    out["date"] = d.isoformat()
    out.index.name = "symbol"
    return out.reset_index()[COLS]


def existing(directory: str) -> set[str]:
    return {os.path.basename(p)[:-7] for p in glob.glob(os.path.join(directory, "*.csv.gz"))}


def save_session(df: pd.DataFrame, directory: str, d: dt.date):
    os.makedirs(directory, exist_ok=True)
    df = df.copy()
    for c in df.columns:
        if df[c].dtype.kind == "f":
            df[c] = df[c].round(3)
    df.sort_values("symbol").to_csv(
        os.path.join(directory, f"{d.isoformat()}.csv.gz"), index=False, compression="gzip")


def combine(directory: str, out_path: str) -> pd.DataFrame:
    files = sorted(glob.glob(os.path.join(directory, "*.csv.gz")))
    if not files:
        return pd.DataFrame(columns=COLS)
    df = pd.concat([pd.read_csv(p, dtype={"symbol": str}) for p in files], ignore_index=True)
    df = df.drop_duplicates(subset=["date", "symbol"], keep="last").sort_values(["date", "symbol"])
    df.to_csv(out_path, index=False, compression="gzip")
    return df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="start", default=None, help="first session, YYYY-MM-DD")
    ap.add_argument("--to", dest="end", default=None, help="last session, default today")
    ap.add_argument("--since-last", action="store_true",
                    help="only sessions newer than the newest shard already stored")
    ap.add_argument("--dir", default=HIST_DIR)
    ap.add_argument("--combine", action="store_true",
                    help="also write data/history.csv.gz for analysis")
    ap.add_argument("--max-sessions", type=int, default=0, help="cap for testing")
    ap.add_argument("--all-symbols", action="store_true")
    args = ap.parse_args()

    have = existing(args.dir)
    end = dt.date.fromisoformat(args.end) if args.end else dt.date.today()

    if args.since_last:
        if not have:
            print("no history yet - run a full backfill first; nothing to do")
            return
        start = dt.date.fromisoformat(max(have)) + dt.timedelta(days=1)
    else:
        start = dt.date.fromisoformat(args.start) if args.start else DEFAULT_START

    print(f"shards: {len(have)} session(s)"
          f" ({min(have) if have else '-'} .. {max(have) if have else '-'})")
    print(f"target: {start} .. {end}")

    n_sessions = int((end - start).days * 5 / 7) + 10 if start <= end else 0
    dates = fe.trading_dates(end, n_sessions) if n_sessions else []
    todo = [d for d in dates if d >= start and d.isoformat() not in have]
    if args.max_sessions:
        todo = todo[:args.max_sessions]
    print(f"{len(todo)} session(s) to fetch")

    done = 0
    for d in todo:
        try:
            rows = one_session(d, args.all_symbols)
        except Exception as exc:                        # noqa: BLE001
            print(f"  ! {d}: {type(exc).__name__}: {exc}", file=sys.stderr)
            continue
        if rows is None:
            print(f"  - {d}: not published")
            continue
        save_session(rows, args.dir, d)
        done += 1
        if done % 20 == 0:
            print(f"  {done}/{len(todo)}  {d}  {len(rows)} symbols", flush=True)

    have = existing(args.dir)
    print(f"\n{done} session(s) written; {len(have)} shards on disk"
          f" ({min(have) if have else '-'} .. {max(have) if have else '-'})")

    if args.combine:
        out = os.path.join(os.path.dirname(args.dir.rstrip("/")) or ".", "history.csv.gz")
        df = combine(args.dir, out)
        print(f"combined: {len(df):,} rows, {df['symbol'].nunique()} symbols, "
              f"{os.path.getsize(out)/1e6:.1f} MB -> {out}")


if __name__ == "__main__":
    main()
