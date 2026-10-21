from pathlib import Path
from openpyxl import load_workbook
import pandas as pd, csv, re, collections
base=Path('/mnt/data/rajit_complete');out=base/'final_output'
df=pd.read_csv(out/'insurance_monthly_2019_2026.csv'); print('rows',len(df),flush=True)
exceptions=[]; summ=[]; validation=[]
fix={ '2019-03': ('2019',10), '2021-06':('2021',10), '2021-07':('2021',10), '2022-07':('Sheet3',10), '2022-09':('Sept 2023',9), '2023-07':('2024',9), '2022-06':('2022',10)}

def category(x):
 t=str(x or '').lower().replace('–','-').replace('—','-')
 if 'individual' in t and ('non-single' in t or 'non single' in t): return 'individual_non_single'
 if 'individual' in t and 'single' in t:return 'individual_single'
 if 'group' in t and ('yearly' in t or 'renewable' in t):return 'group_yearly_renewable'
 if 'group' in t and ('non-single' in t or 'non single' in t):return 'group_non_single'
 if 'group' in t and 'single' in t:return 'group_single'
 return None

def group(x):
 s=str(x or '').lower()
 if 'premium' in s or 'प्रीमियम' in s:return 'first_year_premium_rs_crore'
 if 'polic' in s or 'पॉलिसि' in s:return 'new_policies_count'
 if 'lives covered' in s:return 'group_lives_covered_count'
 if 'sum assured' in s:return 'sum_assured_rs_crore'
 return None

raw_all=out/'insurer_all_reported_metrics_long.csv';current_all=out/'insurer_current_month_metrics_long.csv'
fields=['month','source_id','sheet','insurer','record_type','category','row_label','metric','period_label','period_type','numeric_value','source_cell']
with raw_all.open('w',encoding='utf-8-sig',newline='') as fa,current_all.open('w',encoding='utf-8-sig',newline='') as fc:
 wa=csv.DictWriter(fa,fieldnames=fields);wc=csv.DictWriter(fc,fieldnames=fields);wa.writeheader();wc.writeheader()
 for _,rec in df.iterrows():
  month=rec.month; filename=rec.source_file
  p=base/'raw_data'/filename
  if not p.exists():p=base/'converted_xls'/filename
  try:wb=load_workbook(p,read_only=True,data_only=True)
  except Exception as e:exceptions.append((month,filename,'cannot read workbook',str(e)));continue
  if month in fix:
   sheet_hint,policy_col=fix[month]
   sheets=[s for s in wb if sheet_hint.lower() in s.title.lower()]
  else:
   sheets=[s for s in wb if str(rec.sheet).lower()==s.title.lower()]
   policy_col=None
  if not sheets:exceptions.append((month,filename,'cannot find sheet',str(rec.sheet)));continue
  s=sheets[-1];rows=list(s.iter_rows(max_row=450,max_col=46,values_only=True))
  hdr=None
  for i,r in enumerate(rows[:10]):
   if sum('for ' in str(x or '').lower() or 'up to ' in str(x or '').lower() for x in r)>=2:hdr=i;break
  if hdr is None and month=='2019-03':hdr=2
  if hdr is None:exceptions.append((month,filename,'unrecognized period headings',s.title));continue
  # Read first row of each group boundary; where a report omits monthly columns for policies, capture reported monthly-derived values with source note in correction log.
  hrow=rows[hdr]
  gh=rows[hdr-1] if hdr else []
  metrics={};currentmetric='unclassified'
  for j in range(len(hrow)):
   m=group(gh[j] if j<len(gh) else None)
   if m:currentmetric=m
   metrics[j]=currentmetric
  # column-to-time map. Numeric columns normally begin at index 2, each with dedicated period labels.
  pp={}
  for j,v in enumerate(hrow):
   text=str(v or '')
   z=text.lower()
   if 'growth' in z or 'व्रुद्धि' in z:ptype='growth_percentage'
   elif 'market share' in z or 'मार्केट शेयर' in z:ptype='market_share_percentage'
   elif 'up to' in z or ' तक' in z:ptype='fiscal_ytd'
   elif 'for ' in z or 'के लिये' in z:ptype='monthly'
   else:ptype='other'
   pp[j]=(text,ptype)
  # For 2023 July erroneous 'up to' label at monthly column 9/10 and 2019 March Hindi headers special case.
  if month=='2023-07':
   pp[9]=(pp[9][0],'monthly_previous_year'); pp[10]=(pp[10][0],'monthly')
  if month=='2019-03':
   for j in (2,3,9,10,16,17,23,24): pp[j]=(pp[j][0],'monthly' if j%7==3 else 'monthly_previous_year')
  if month=='2022-07':
   for j in (3,10,17,24): pp[j]=(pp[j][0],'monthly_corrected_from_july_comparison')
  selected_cols=([3,7,11,15] if month in ('2019-04','2020-04') else [2,9,16,23] if month in ('2022-09','2023-07') else [3,10,17,24])
  insurer='';count_rows=0;rowtypes=collections.Counter();insurers=set()
  for ri,r in enumerate(rows[hdr+1:],start=hdr+2):
   if len(r)<2:continue
   label=str(r[1] or '').strip()
   if not label:continue
   cat=category(label)
   if cat:rtype='category';rowtypes[rtype]+=1
   elif r[0] is not None or label.lower() in ('grand total','private total','lic of india'):
    insurer=label;rtype='insurer_total';insurers.add(insurer);rowtypes[rtype]+=1
   else:
    # ignore text-only subtitle rows; still preserve numeric rows if present
    if not any(isinstance(x,(int,float)) for x in r[2:]):continue
    rtype='other_numeric_row';rowtypes[rtype]+=1
   for j,x in enumerate(r):
    if j<2 or not isinstance(x,(int,float)) or isinstance(x,bool):continue
    if j>=len(gh):continue
    period,pt=pp.get(j,('','other'))
    obj={'month':month,'source_id':filename,'sheet':s.title,'insurer':insurer,'record_type':rtype,'category':cat or '', 'row_label':label,'metric':metrics[j],'period_label':period,'period_type':pt,'numeric_value':x,'source_cell':f'{chr(65+j) if j<26 else chr(65+j//26-1)+chr(65+j%26)}{ri}'}
    wa.writerow(obj);count_rows+=1
    # select current period columns: for normal files infer per metric band by policy col index and 7-wide patterns; special 2023 July infer correction.
    if j in selected_cols and (pt in ('monthly','monthly_corrected_from_july_comparison','monthly_previous_year') or month in ('2019-04','2020-04')):
     wc.writerow(obj)
  summ.append({'month':month,'source_file':filename,'source_sheet':s.title,'insurer_labels':len(insurers),'numeric_metrics':count_rows,'category_rows':rowtypes['category']})
  if len(insurers)<10 or count_rows<100:
   exceptions.append((month,filename,'low records',f'insurers {len(insurers)} metrics {count_rows}'))
  if len(summ)%15==0:print('processed',len(summ),flush=True)

pd.DataFrame(summ).to_csv(out/'metrics_coverage_by_month.csv',index=False)
pd.DataFrame(exceptions,columns=['month','file','problem','detail']).to_csv(out/'detailed_extraction_issues.csv',index=False)
print('written',len(summ),'months; issues',len(exceptions),'output sizes',raw_all.stat().st_size,current_all.stat().st_size)
