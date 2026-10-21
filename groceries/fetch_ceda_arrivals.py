"""
Fetch arrivals data from the CEDA Agri-Market API (the same request the web page makes).
Run on YOUR machine:  pip install requests pandas
Examples:
  python fetch_ceda_arrivals.py --commodity tomato --calc m
  python fetch_ceda_arrivals.py --commodity tomato --calc d --start 2024-05-01 --end 2024-08-31

Outputs (created next to the script):
  raw/      untouched JSON response for every call (never edit these)
  interim/  generic flattened CSV (column names exactly as the API gives them)
  request_log.csv  one row per call: when, payload, HTTP status, size, record count
"""
import argparse
import datetime as dt
import json
import time
from pathlib import Path

import pandas as pd
import requests

URL = "https://agmarknet.ceda.ashoka.edu.in/api/quantities"

# commodity_id 78 came from your captured request (page was set to Tomato).
# Add onion / potato ids after capturing the same request with that commodity selected.
COMMODITIES = {
    "tomato": 78,
    "onion": 23,
    "potato": 24,
}

HEADERS = {
    "Content-Type": "application/json",
    "Referer": "https://agmarknet.ceda.ashoka.edu.in/",
    "User-Agent": "Mozilla/5.0 (CS215 student project; contact: your-email@example.com)",
}

BASE = Path(__file__).parent
RAW, INTERIM = BASE / "raw", BASE / "interim"
RAW.mkdir(exist_ok=True)
INTERIM.mkdir(exist_ok=True)


def fetch(commodity_id, calc, start, end, state_id=0, district_id=0):
    payload = {
        "state_id": state_id,          # 0 = All India (as in your request)
        "commodity_id": commodity_id,
        "district_id": district_id,    # 0 = all districts
        "calculation_type": calc,      # "m" = monthly (confirmed); "d" = daily (UNVERIFIED guess)
        "start_date": start,
        "end_date": end,
    }
    resp = requests.post(URL, headers=HEADERS, json=payload, timeout=60)
    return payload, resp


def find_records(obj):
    """Return the first list of dicts found anywhere in the JSON, else None."""
    if isinstance(obj, list) and obj and isinstance(obj[0], dict):
        return obj
    if isinstance(obj, dict):
        for v in obj.values():
            found = find_records(v)
            if found is not None:
                return found
    return None


def year_chunks(start, end):
    """Split [start, end] into calendar-year pieces (useful for daily pulls)."""
    s, e = dt.date.fromisoformat(start), dt.date.fromisoformat(end)
    out = []
    while s <= e:
        stop = min(dt.date(s.year, 12, 31), e)
        out.append((s.isoformat(), stop.isoformat()))
        s = stop + dt.timedelta(days=1)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commodity", required=True, choices=COMMODITIES.keys())
    ap.add_argument("--calc", default="m", choices=["m", "d"])
    ap.add_argument("--start", default="2018-12-31")  # same as the web page's request
    ap.add_argument("--end", default="2026-09-29")
    ap.add_argument("--by-year", action="store_true", help="one request per calendar year")
    ap.add_argument("--pause", type=float, default=2.0, help="seconds between requests")
    args = ap.parse_args()

    cid = COMMODITIES[args.commodity]
    ranges = year_chunks(args.start, args.end) if args.by_year else [(args.start, args.end)]
    pull_ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_rows, frames = [], []

    for s, e in ranges:
        payload, resp = fetch(cid, args.calc, s, e)
        stem = f"{args.commodity}_{args.calc}_{s}_{e}_{pull_ts}"
        (RAW / f"{stem}.raw.txt").write_bytes(resp.content)  # exact bytes, whatever they are
        n_rec = None
        if resp.ok:
            try:
                data = resp.json()
                (RAW / f"{stem}.json").write_text(json.dumps(data, indent=1))
                records = find_records(data)
                if records is not None:
                    df = pd.json_normalize(records)
                    df.insert(0, "_request_start", s)
                    df.insert(1, "_request_end", e)
                    frames.append(df)
                    n_rec = len(df)
                else:
                    print(f"[{s}..{e}] no list of records found; top-level type "
                          f"{type(data).__name__}, keys: "
                          f"{list(data)[:10] if isinstance(data, dict) else 'n/a'}")
            except ValueError:
                print(f"[{s}..{e}] response was not JSON; see raw/{stem}.raw.txt")
        log_rows.append({
            "pull_time": pull_ts, "url": URL, "payload": json.dumps(payload),
            "http_status": resp.status_code, "bytes": len(resp.content), "n_records": n_rec,
        })
        print(f"[{s}..{e}] HTTP {resp.status_code}, {len(resp.content)} bytes, records: {n_rec}")
        time.sleep(args.pause)

    log = pd.DataFrame(log_rows)
    log_path = BASE / "request_log.csv"
    log.to_csv(log_path, mode="a", header=not log_path.exists(), index=False)

    if frames:
        out = pd.concat(frames, ignore_index=True)
        out_path = INTERIM / f"{args.commodity}_{args.calc}_flat_{pull_ts}.csv"
        out.to_csv(out_path, index=False)
        print(f"\nSaved {out.shape[0]} rows x {out.shape[1]} cols -> {out_path}")
        print("Columns:", list(out.columns))
        print(out.head(10).to_string())
        print(out.tail(5).to_string())
    else:
        print("\nNothing flattened. Check raw/ and request_log.csv, then send me what you see.")


if __name__ == "__main__":
    main()