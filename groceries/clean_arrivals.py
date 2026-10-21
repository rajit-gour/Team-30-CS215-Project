"""
Clean daily CEDA arrivals (output of fetch_ceda_arrivals.py --calc d) into monthly series.

Usage:
  python clean_arrivals.py --interim interim --out cleaned

Reads  interim/*_d_flat_*.csv   (columns: t, cmdty, qty, _request_start, _request_end)
Writes cleaned/arrivals_daily_flagged.csv   every reported day + coverage_break flag
       cleaned/official_monthly.csv         team long-format schema (+ day counts)
       cleaned/data_quality_report.csv      runs of missing days / suspect days
Rules: absence is never zero; flagged days are excluded from the monthly mean;
       the monthly outcome is MEAN tonnes per valid day (not a sum), so partial months stay comparable.
"""
import argparse
import glob
from pathlib import Path

import numpy as np
import pandas as pd

GEOGRAPHY = "All India"
SOURCE_ID = "ceda_agmarknet_arrivals"
BREAK_RATIO = 0.20     # day is suspect if arrivals < 20% of the centred 365-day median
OK_COVERAGE = 0.80     # month "ok" if valid days / days in month >= this
MIN_COVERAGE = 0.50    # below this the month is "invalid" (outcome left blank)


def load(interim):
    files = sorted(glob.glob(str(Path(interim) / "*_d_flat_*.csv")))
    if not files:
        raise SystemExit(f"No daily files found in {interim} (expected *_d_flat_*.csv)")
    parts = []
    for f in files:
        d = pd.read_csv(f)
        d["source_file"] = Path(f).name
        parts.append(d)
    df = pd.concat(parts, ignore_index=True)
    df = df.rename(columns={"t": "date", "cmdty": "commodity", "qty": "arrivals_tonnes"})
    df["date"] = pd.to_datetime(df["date"])
    df["arrivals_tonnes"] = pd.to_numeric(df["arrivals_tonnes"], errors="coerce")

    # Overlapping pulls: keep the latest file's value, warn if values disagree.
    df = df.sort_values(["commodity", "date", "source_file"])
    dup = df.duplicated(["commodity", "date"], keep=False)
    if dup.any():
        spread = df[dup].groupby(["commodity", "date"])["arrivals_tonnes"].agg(lambda s: s.max() - s.min())
        n_diff = int((spread > 1e-6).sum())
        print(f"Note: {int(dup.sum())} rows were duplicated across files; {n_diff} dates had differing values.")
    return df.drop_duplicates(["commodity", "date"], keep="last").reset_index(drop=True)


def flag_days(df):
    out = []
    for com, g in df.groupby("commodity"):
        g = g.set_index("date").sort_index()
        ref = g["arrivals_tonnes"].rolling("365D", center=True, min_periods=20).median()
        g["reference_median"] = ref
        g["coverage_break"] = g["arrivals_tonnes"] < BREAK_RATIO * ref
        g["arrivals_missing_value"] = g["arrivals_tonnes"].isna()
        out.append(g.reset_index())
    return pd.concat(out, ignore_index=True)


def quality_runs(daily):
    """List runs of (a) calendar days with no row, (b) reported-but-suspect days."""
    rows = []
    for com, g in daily.groupby("commodity"):
        full = pd.date_range(g["date"].min(), g["date"].max(), freq="D")
        have = set(g["date"])
        missing = pd.Series([d for d in full if d not in have], dtype="datetime64[ns]")
        sus = g.loc[g["coverage_break"], "date"].reset_index(drop=True)
        for kind, s in (("no_row", missing), ("suspect_low_arrivals", sus)):
            if s.empty:
                continue
            run = (s.diff().dt.days != 1).cumsum()
            for _, r in s.groupby(run):
                rows.append({"commodity": com, "issue": kind, "start": r.iloc[0].date(),
                             "end": r.iloc[-1].date(), "n_days": len(r)})
    return pd.DataFrame(rows).sort_values(["commodity", "start"]) if rows else pd.DataFrame()


def monthly(daily):
    rows = []
    for com, g in daily.groupby("commodity"):
        months = pd.period_range(g["date"].min(), g["date"].max(), freq="M")
        g = g.assign(month=g["date"].dt.to_period("M"))
        for m in months:
            gm = g[g["month"] == m]
            valid = gm[~gm["coverage_break"] & ~gm["arrivals_missing_value"]]
            dim = m.days_in_month
            cov = len(valid) / dim
            if len(gm) == 0:
                flag = "missing_no_data"
            elif cov >= OK_COVERAGE:
                flag = "ok"
            elif cov >= MIN_COVERAGE:
                flag = "partial_coverage"
            else:
                flag = "invalid_coverage_break" if gm["coverage_break"].any() else "invalid_sparse"
            value = valid["arrivals_tonnes"].mean() if flag in ("ok", "partial_coverage") else np.nan
            rows.append({
                "month": m.start_time.strftime("%Y-%m-01"),
                "geography": GEOGRAPHY,
                "series_id": f"{com.lower()}_arrivals_mean_daily",
                "outcome_name": f"{com} wholesale arrivals, mean per reported day",
                "outcome_value": value,
                "outcome_unit": "tonnes/day (unit assumed from portal; unverified)",
                "source_id": SOURCE_ID,
                "quality_flag": flag,
                "days_reported": len(gm),
                "days_valid": len(valid),
                "days_in_month": dim,
            })
    out = pd.DataFrame(rows)
    assert not out.duplicated(["month", "series_id", "geography"]).any()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--interim", default="interim")
    ap.add_argument("--out", default="cleaned")
    a = ap.parse_args()
    Path(a.out).mkdir(exist_ok=True)

    raw = load(a.interim)
    print(f"Raw daily rows: {len(raw)}  cols: {raw.shape[1]}  "
          f"range: {raw['date'].min().date()} to {raw['date'].max().date()}")
    daily = flag_days(raw)
    qr = quality_runs(daily)
    mon = monthly(daily)

    daily.to_csv(Path(a.out) / "arrivals_daily_flagged.csv", index=False)
    mon.to_csv(Path(a.out) / "official_monthly.csv", index=False)
    qr.to_csv(Path(a.out) / "data_quality_report.csv", index=False)

    print("\nMonthly flags:\n", mon["quality_flag"].value_counts().to_string())
    print("\nData quality runs:\n", qr.to_string(index=False) if len(qr) else "none")
    print("\nMonthly table:\n", mon[["month", "series_id", "outcome_value", "quality_flag",
                                    "days_reported", "days_valid"]].to_string(index=False))


if __name__ == "__main__":
    main()