# NSE EOD F&O buildup screener — automated, free, zero-storage

This replaces `consolidated_manual_date_fno.xlsm` end to end. Instead of Power
Query pulling CSVs into a 26 MB macro workbook that you then have to open, the
pipeline downloads the last 20 sessions straight from NSE archives, recomputes
the same sheets in about 40 seconds, and pushes them to your Telegram chat.
Nothing is kept between runs, so the footprint is the ~400 KB of output you
receive.

---

## 1. Why this design

| Requirement | How it is met |
|---|---|
| Free of cost | GitHub Actions on a public repo (unlimited minutes), Telegram Bot API, NSE archive files. No paid API, no server, no database. |
| Least storage | Nothing is persisted. The 20-session cache is deleted at the end of every run. Repo holds ~40 KB of code. Each run downloads ~34 MB and throws it away. |
| Fully automated | One cron entry. No Excel, no manual download, no refresh button. |
| Sheet by sheet on Telegram | One message with the day's signals, then the workbook, then each sheet as its own CSV, then a PNG of the buildup grid. |

Storage comparison: the old workbook was **26 MB** on disk (and a 133 MB sheet
inside it) and grew every time you refreshed. This is **~0.4 MB per day**, kept
only in your Telegram chat.

---

## 2. Quickstart (about 10 minutes, one time)

1. **Create a Telegram bot.** Message `@BotFather` → `/newbot` → pick a name.
   Copy the token (`123456:ABC-...`).
2. **Get your chat id.** Send your new bot any message, then open
   `https://api.telegram.org/bot<TOKEN>/getUpdates` and read
   `"chat":{"id":123456789}`. (Or message `@userinfobot`.)
3. **Create a public GitHub repo** and push this folder:

   ```bash
   git init && git add . && git commit -m "fno eod"
   git branch -M main
   git remote add origin https://github.com/<you>/fno-eod.git
   git push -u origin main
   ```

4. **Add the two secrets** — repo → Settings → Secrets and variables → Actions:
   `TELEGRAM_TOKEN` and `TELEGRAM_CHAT_ID`.
5. **Test it now** — Actions tab → *NSE EOD F&O to Telegram* → Run workflow.
   You should get the message and files in under two minutes.

That's it. It runs every weekday at 19:40 IST.

### Run it locally instead

```bash
pip install -r requirements.txt
export TELEGRAM_TOKEN=... TELEGRAM_CHAT_ID=...
python fno_eod.py --days 20 --telegram --out out
# or a specific day, no Telegram:
python fno_eod.py --date 2026-11-06 --out out
```

---

## 3. What you receive

| File | What it is |
|---|---|
| text message | Counts per signal for the latest session, the symbol list in each bucket, and the widest futures-vs-spot gaps. Readable on a phone lock screen. |
| `fno_eod_YYYYMMDD.xlsx` | All six sheets, signal cells colour-coded exactly like the old workbook. |
| `buildup_YYYYMMDD.png` | The buildup grid for the last 5 sessions — LB green, SB red, SC blue, LU yellow. |
| `cash_alys_fno_*.csv` | Per symbol: 20-session min/max/close, delivery % 20DMA, qty-per-trade 20DMA, turnover. |
| `fut_alys_data_*.csv` | Futures: settle, total OI, last OI change %, OI contracts per trade 20DMA, lot size. |
| `op_alys_data_*.csv` | CE and PE OI in lots, day by day, plus the last PE/CE ratio. |
| `buildup_summary_*.csv` | The 20-day LB/SB/SC/LU grid and the futures-vs-spot divergence %. |
| `general_*.csv` | Date table (D1..D20) and the threshold actually used. |
| `stock_specific_data_*.csv` | One-stock panel for the symbol with the biggest OI move. |

Signal rules (same as your workbook, threshold from `FNO_THRESHOLD`, default 8%):

| | OI up | OI down |
|---|---|---|
| **price up** | LB long buildup | SC short covering |
| **price down** | SB short buildup | LU long unwinding |

---

## 4. What changed from the old workbook (and why)

* **Dead download links.** `fo<DDMMMYYYY>bhav.csv.zip` and `CD_Bhavcopy<DDMMYY>.zip`
  were retired by NSE in July 2024 and now return 404. The pipeline uses the
  UDiFF files that replaced them:
  `BhavCopy_NSE_FO_0_0_0_YYYYMMDD_F_0000.csv.zip` and
  `sec_bhavdata_full_DDMMYYYY.csv`. Both verified working.
* **Rollover-safe OI.** Futures OI is summed across all series and the OI-change
  percentage is computed on that total, while price comes from the nearest
  series. Without this, expiry week produces hundreds of fake "long unwinding"
  signals as positions roll.
* **Expiry filter that actually filters.** Options use the dominant expiry
  (largest total OI) for that symbol and day, so CE and PE legs are always from
  the same expiry. The old sheet summed every expiry together.
* **Unit fix.** OI change is divided by lot size before dividing by trade count,
  so the futures "TQ/NT" column is contracts per trade rather than a
  lot-size-scaled number.
* **No local paths.** Nothing points at `D:\kapil\...`. Data is fetched, not read
  off a disk.
* **No pinbar section.** Your pinbar macro pointed at a table
  (`ind_nifty500list__2`) that no longer exists in the file. If you want that
  screener back, it is a 20-line addition — say the word.

---

## 5. Reliability notes

* **Timing.** GitHub's cron is best-effort and can start 5–30 minutes late. The
  script is scheduled at 19:40 IST with `--wait 25`, so if NSE is late it keeps
  retrying until the file appears.
* **60-day rule.** GitHub disables scheduled workflows on repos with no activity
  for 60 days. The workflow here writes an artifact each run, but to be safe
  click *Run workflow* once every couple of months, or add a monthly no-op commit.
* **Datacenter IPs.** NSE occasionally throttles cloud IPs. If you start seeing
  HTTP 403 or empty downloads, run the same script on any always-on machine at
  home (a Raspberry Pi, or an old Android phone with Termux) — it is one file
  plus `pip install -r requirements.txt`.
* **Holidays.** Sessions are discovered by probing the archive, so holidays,
  weekends and special muhurat sessions (which have no UDiFF file) are skipped
  automatically; the window is always 20 real sessions.
* **Freshness check.** The header of every message carries the session date. If
  it is not today's date, NSE had not published yet — just re-run.

---

## 6. Alternatives if GitHub Actions does not suit you

| Option | Free? | Notes |
|---|---|---|
| **GitHub Actions (this repo)** | yes, unlimited on public repos | No server to maintain, zero storage. Best default. |
| **Oracle Cloud Always Free VM** (Ampere A1) | yes, forever | Exact timing via a systemd timer, your own IP, 200 GB block storage included. More setup. |
| **Home device** (Pi / Termux) | yes | Best if NSE blocks cloud IPs. Needs the device to be on at 19:30 IST. |
| **Cloudflare Workers cron** | yes | Not suitable: no pandas, 3 MB bundle limit, no long-running jobs. |

---

## 7. Tuning

```bash
FNO_THRESHOLD=0.05   # more signals (5% OI change)
FNO_THRESHOLD=0.12   # fewer, stronger signals
--days 20            # window length; the sheets are built for 20
--date 2026-11-06    # reproduce any past session
```

The whole thing is one file, `fno_eod.py`, roughly 400 lines. `sheets()` is the
only function you need to touch to add a column or a new screen.
