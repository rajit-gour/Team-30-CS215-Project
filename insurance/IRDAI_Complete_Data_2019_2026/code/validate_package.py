from pathlib import Path
import pandas as pd, numpy as np, zipfile, json
base=Path('/mnt/data/rajit_complete/final_output')
main=pd.read_csv(base/'insurance_monthly_2019_2026.csv')
d=pd.read_csv(base/'insurer_current_month_metrics_long.csv',low_memory=False)
# Pivot 4 categories/insurer group key
subset=d[(d.insurer.isin(['Grand Total','LIC of India','Life Insurance Corporation of India','Private Total']))&(d.category.isin(['individual_single','individual_non_single']))&(d.metric.isin(['first_year_premium_rs_crore','new_policies_count','sum_assured_rs_crore']))].copy()
subset['key']=subset.insurer.map({'Grand Total':'national','LIC of India':'lic','Life Insurance Corporation of India':'lic','Private Total':'private'})+'_'+subset.category.map({'individual_single':'single','individual_non_single':'non_single'})+'_'+subset.metric.map({'new_policies_count':'policies','first_year_premium_rs_crore':'premium_crore','sum_assured_rs_crore':'sum_assured_crore'})
w=subset.pivot_table(index='month',columns='key',values='numeric_value',aggfunc='first').reset_index()
combined=main.merge(w,on='month',how='left')
for prefix in ('national','lic','private'):
 for suffix in ('policies','premium_crore','sum_assured_crore'):
  a=f'{prefix}_single_{suffix}';b=f'{prefix}_non_single_{suffix}'
  if a in combined and b in combined:
   combined[f'{prefix}_individual_total_{suffix}']=combined[a]+combined[b]
validation=[]
for _,r in combined.iterrows():
 mm=r.month
 main_equal=(int(r['individual_single'])==int(r['national_single_policies']) and int(r['individual_non_single'])==int(r['national_non_single_policies']))
 licprivate_ok=True; premium_ok=True
 for cat in ['single','non_single']:
  def comp(suf):
   vs=[r.get(f'{p}_{cat}_{suf}',float('nan')) for p in ['national','lic','private']]
   if pd.isna(vs).any():return 'MISSING'
   gap=vs[0]-vs[1]-vs[2]
   return 'PASS' if abs(gap)<(0.01 if suf!='policies' else 0.5) else f'FAIL gap {gap:.6g}'
  lp=comp('policies');pr=comp('premium_crore')
  if lp!='PASS':licprivate_ok=False
  if pr!='PASS':premium_ok=False
 validation.append(dict(month=mm,main_vs_grandtotal='PASS' if main_equal else 'FAIL',lic_plus_private_policy_total='PASS' if licprivate_ok else 'MISSING_OR_FAIL',lic_plus_private_premium_total='PASS' if premium_ok else 'MISSING_OR_FAIL'))
combined.to_csv(base/'insurance_monthly_enriched.csv',index=False)
v=pd.DataFrame(validation);v.to_csv(base/'validation_results.csv',index=False)
stats={ 'source_workbooks':95,'months':len(main),'start':str(main.month.min()),'end':str(main.month.max()),'mean_individual_policies':round(float(main.individual_total.mean()),1),'median_individual_policies':int(main.individual_total.median()),'min':int(main.individual_total.min()),'max':int(main.individual_total.max()),'main_total_crosschecks_passed':int((v.main_vs_grandtotal=='PASS').sum()),'lic_private_count_reconciliations_passed':int((v.lic_plus_private_policy_total=='PASS').sum()),'lic_private_premium_reconciliations_passed':int((v.lic_plus_private_premium_total=='PASS').sum()),'insurer_current_metric_records':len(d),'other_all_period_metric_records':sum(1 for _ in open(base/'insurer_all_reported_metrics_long.csv',encoding='utf-8-sig'))-1 }
(base/'dataset_statistics.json').write_text(json.dumps(stats,indent=2),encoding='utf-8')
print(stats);print('failed:',v[(v!='PASS').any(axis=1)].to_string(index=False)[:1800]);print('columns',len(combined.columns))
(base/'README_DATASET.txt').write_text('''TEAM 30: IRDAI LIFE INSURANCE DATASET\n\nPRIMARY STUDY PERIOD: January 2019 to September 2026 inclusive (93 calendar months).\nSOURCE: 95 original IRDAI Excel files submitted by Rajit Singh Gour, converted to a machine-readable form for analysis; original files unchanged.\n\nFILES:\n- insurance_monthly_2019_2026.csv: principal counts by month (single and non-single premium, total), with provenance and processing status.\n- insurance_monthly_enriched.csv: added insurer splits, individual first-year premium, sum assured where reported. Premium and sum assured units are Rs crore; policy counts are counts.\n- insurer_current_month_metrics_long.csv: current-month detailed metrics by insurer and category, with references to source cells (monthly period source column detection reviewed for anomalies).\n- insurer_all_reported_metrics_long.csv: comprehensive numeric archival extraction from selected English-language worksheet of each month's report, including historical comparison, financial YTD, changes, market share, plus all categories and insurers. Includes the printed period headings, which may contain errors.\n- source_file_map.csv: original file names (IRDAI_001 etc.).\n- recovery_and_correction_log.csv: special cases and original-source cross-checks.\n- validation_results.csv: reported Grand Total reconciliations and LIC vs private subtotals.\n- metrics_coverage_by_month.csv: monthly record-level coverage.\n- detailed_extraction_issues.csv: remaining structural extraction issues, if any.\n\nIMPORTANT: There is a documented July 2022 source whose heading incorrectly refers to June, but values align with the July reference in the following year's report. March 2019 included English data on lower rows of a sheet otherwise headed in Hindi. Several months were recovered using the prior-year comparison column of a later report. Such sources are detailed in the correction log.\n\nQUALITY: Main 93-month national individual policy totals reconcile to the respective Grand Total rows from the selected insurer-level worksheets. This constitutes internal consistency checking, not independent external verification. See validation_results.csv for LIC+private reconciliation and any exceptions. The dataset has not been matched with Google Trends; no predictive modelling has been performed. The counts reflect policies issued, not unique customers or all insurance-related consumer decisions.\n''',encoding='utf-8')
# package zipped, files are roughly 80MB before compression
path='/mnt/data/Rajit_IRDAI_Complete_Data_2019_2026.zip'
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
 for p in sorted(base.iterdir()):
  if p.suffix in ('.csv','.txt','.json'):z.write(p,'data/'+p.name)
 z.write('/mnt/data/rajit_complete/finalize_data.py','code/finalize_data.py')
 z.write('/mnt/data/rajit_complete/assemble_all_metrics.py','code/assemble_all_metrics.py')
 z.write('/mnt/data/rajit_complete/validate_package.py','code/validate_package.py')
print('package MB',round(Path(path).stat().st_size/1024/1024,2))
