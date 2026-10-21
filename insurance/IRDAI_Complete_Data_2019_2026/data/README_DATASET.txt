TEAM 30: IRDAI LIFE INSURANCE DATASET

PRIMARY STUDY PERIOD: January 2019 to September 2026 inclusive (93 calendar months).
SOURCE: 95 original IRDAI Excel files submitted by Rajit Singh Gour, converted to a machine-readable form for analysis; original files unchanged.

FILES:
- insurance_monthly_2019_2026.csv: principal counts by month (single and non-single premium, total), with provenance and processing status.
- insurance_monthly_enriched.csv: added insurer splits, individual first-year premium, sum assured where reported. Premium and sum assured units are Rs crore; policy counts are counts.
- insurer_current_month_metrics_long.csv: current-month detailed metrics by insurer and category, with references to source cells (monthly period source column detection reviewed for anomalies).
- insurer_all_reported_metrics_long.csv: comprehensive numeric archival extraction from selected English-language worksheet of each month's report, including historical comparison, financial YTD, changes, market share, plus all categories and insurers. Includes the printed period headings, which may contain errors.
- source_file_map.csv: original file names (IRDAI_001 etc.).
- recovery_and_correction_log.csv: special cases and original-source cross-checks.
- validation_results.csv: reported Grand Total reconciliations and LIC vs private subtotals.
- metrics_coverage_by_month.csv: monthly record-level coverage.
- detailed_extraction_issues.csv: remaining structural extraction issues, if any.

IMPORTANT: There is a documented July 2022 source whose heading incorrectly refers to June, but values align with the July reference in the following year's report. March 2019 included English data on lower rows of a sheet otherwise headed in Hindi. Several months were recovered using the prior-year comparison column of a later report. Such sources are detailed in the correction log.

QUALITY: Main 93-month national individual policy totals reconcile to the respective Grand Total rows from the selected insurer-level worksheets. This constitutes internal consistency checking, not independent external verification. See validation_results.csv for LIC+private reconciliation and any exceptions. The dataset has not been matched with Google Trends; no predictive modelling has been performed. The counts reflect policies issued, not unique customers or all insurance-related consumer decisions.
