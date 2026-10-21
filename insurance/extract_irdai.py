#!/usr/bin/env python3
"""IRDAI batch extractor. Python standard library only (no pip install)
Run from project root: python extract_irdai.py
Preserves EVERY populated cell from ALL sheets (including reference/Hindi sheets)
and exports structured CURRENT English New Business Statement data separately.
Never guesses numeric fields from changed layouts; reports exceptions.
"""
import csv, re, json, sys, hashlib, traceback
from pathlib import Path
from datetime import datetime
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from posixpath import normpath, dirname, join

_NS = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
       'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
       'p':'http://schemas.openxmlformats.org/package/2006/relationships'}

def _colindex(s):
    n=0
    for ch in s.upper():
        if 'A' <= ch <= 'Z': n=n*26+ord(ch)-64
        else: break
    return n-1

def _xml(z, name):
    return ET.fromstring(z.read(name))

def load_xlsx(path):
    """Read cached XLSX cell values without third-party dependencies.
    Returns list of (sheet name, rows), limited to A1:CB400 as in v3.
    Original XLSX files are never modified.
    """
    result=[]
    with ZipFile(path) as z:
        names=set(z.namelist())
        shared=[]
        if 'xl/sharedStrings.xml' in names:
            root=_xml(z,'xl/sharedStrings.xml')
            for si in root.findall('m:si',_NS):
                shared.append(''.join(t.text or '' for t in si.findall('.//m:t',_NS)))
        wb=_xml(z,'xl/workbook.xml')
        rel=_xml(z,'xl/_rels/workbook.xml.rels')
        rels={item.get('Id'):item.get('Target') for item in rel}
        for sh in wb.findall('m:sheets/m:sheet',_NS):
            relid=sh.get('{'+_NS['r']+'}id')
            target=rels.get(relid)
            if not target: continue
            sheetpath=target.lstrip('/') if target.startswith('/') else normpath(join('xl',target))
            if sheetpath not in names: continue
            root=_xml(z,sheetpath)
            rows=[[None]*80 for _ in range(400)]
            for cell in root.findall('.//m:sheetData/m:row/m:c',_NS):
                ref=cell.get('r','')
                cm=re.match(r'([A-Z]+)([0-9]+)',ref)
                if not cm: continue
                i=int(cm.group(2))-1; j=_colindex(cm.group(1))
                if i<0 or i>=400 or j<0 or j>=80: continue
                kind=cell.get('t','')
                value_node=cell.find('m:v',_NS)
                inline_node=cell.find('m:is',_NS)
                if kind=='inlineStr' and inline_node is not None:
                    val=''.join(t.text or '' for t in inline_node.findall('.//m:t',_NS))
                elif value_node is None:
                    continue
                else:
                    raw=value_node.text or ''
                    if kind=='s':
                        try: val=shared[int(raw)]
                        except (ValueError,IndexError): val=raw
                    elif kind in ('str','e','d'):
                        val=raw
                    elif kind=='b': val=(raw=='1')
                    else:
                        try: val=float(raw)
                        except ValueError: val=raw
                rows[i][j]=val
            result.append((sh.get('name',''),rows))
    return result


def load_xls(path):
    """Read legacy binary Excel files (.xls). Requires: pip install xlrd"""
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError("Legacy .xls requires xlrd: python -m pip install xlrd") from exc
    wb=xlrd.open_workbook(str(path), on_demand=True)
    result=[]
    for sh in wb.sheets()[:40]:
        rows=[[None]*80 for _ in range(400)]
        for i in range(min(sh.nrows,400)):
            for j in range(min(sh.ncols,80)):
                cell=sh.cell(i,j)
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK): continue
                v=cell.value
                if cell.ctype==xlrd.XL_CELL_DATE:
                    try: v=xlrd.xldate_as_datetime(v,wb.datemode).isoformat()
                    except Exception: pass
                rows[i][j]=v
        result.append((sh.name, rows))
    wb.release_resources()
    return result

def load_workbook(path):
    if path.suffix.lower()=='.xls': return load_xls(path)
    return load_xlsx(path)

ROOT=Path(__file__).resolve().parent
RAW=ROOT/'raw_data'; OUT=ROOT/'output'; OUT.mkdir(exist_ok=True)
KEYS=('Individual Single Premium','Individual Non-Single Premium','Group Single Premium','Group Non-Single Premium','Group Yearly Renewable Premium')
MONTHS={m.lower():i for i,m in enumerate(['January','February','March','April','May','June','July','August','September','October','November','December'],1)}
MONTH_ALIASES={**MONTHS, **{m[:3]:v for m,v in MONTHS.items()},'sept':9}

def norm(x): return re.sub(r'\s+',' ',str(x or '').strip()).lower()
def month_parse(text):
    t=str(text or '')
    # Prefer literal full month or English abbreviation in date headers/titles.
    # Do not use generic isolated year mentions as evidence of reporting month.
    for hit in re.finditer(r'\b(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,10})[,.]?\s*(20\d{2})\b',t,re.I):
        mm=MONTH_ALIASES.get(hit.group(2).lower())
        if mm and 1<=int(hit.group(1))<=31:return f'{int(hit.group(3)):04d}-{mm:02d}'
    for hit in re.finditer(r'\b([A-Za-z]{3,10})\s*[,./-]?\s*(20\d{2})\b',t,re.I):
        mm=MONTH_ALIASES.get(hit.group(1).lower())
        if mm:return f'{int(hit.group(2)):04d}-{mm:02d}'
    hit=re.search(r'\b(\d{1,2})[./-](\d{1,2})[./-](20\d{2})\b',t)
    if hit and 1<=int(hit.group(2))<=12:return f'{int(hit.group(3)):04d}-{int(hit.group(2)):02d}'
    return None

def label_kind(s):
    z=norm(s)
    if 'individual' in z and ('non-single' in z or 'non single' in z): return 'individual_non_single'
    if 'individual' in z and 'single' in z:return 'individual_single'
    if 'group' in z and ('yearly renewable' in z or 'renewable' in z):return 'group_yearly_renewable'
    if 'group' in z and ('non-single' in z or 'non single' in z):return 'group_non_single'
    if 'group' in z and 'single' in z:return 'group_single'
    return None
def valstr(x):
    if x is None:return ''
    if isinstance(x,datetime): return x.isoformat()
    return str(x)
def save(name,rows,fields):
    with (OUT/name).open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.DictWriter(f,fieldnames=fields,extrasaction='ignore');w.writeheader();w.writerows(rows)
def header_metrics(rows):
    """Find metric group bands via header row, then tag each column using its OWN period header.
    Do not infer metric meaning based on fixed column positions."""
    best=None
    for h in range(min(len(rows),12)):
        rr=rows[h]
        if sum(bool(re.search(r'For\s+\w+\s*,?\s*20\d{2}|Up\s+to',valstr(x),re.I)) for x in rr)>=3:
            best=h;break
    if best is None:raise ValueError('No recognisable month/cumulative header row')
    metrics=[];group='unknown'
    for i,x in enumerate(rows[best-1] if best>0 else []):
        z=norm(x)
        if 'first year premium' in z or z=='premium':group='first_year_premium_rs_crore'
        elif 'no. of policies' in z or 'number of policies' in z:group='number_of_policies_schemes'
        elif 'lives covered' in z:group='group_lives_covered'
        elif 'sum assured' in z:group='sum_assured'
        metrics.append(group)
    # For merged headings the preceding row may have sparse heading cells; carry groups forward.
    heads=[valstr(x).strip() for x in rows[best]]
    return best,[(i,(metrics[i] if i<len(metrics) else 'unknown'),heads[i]) for i in range(len(heads)) if heads[i] and i>=2]
def extract(file):
    wb=load_workbook(file)
    cells=[];structured=[];metadata=[];issues=[]
    sha=hashlib.sha256(file.read_bytes()).hexdigest()
    for sheetnum, (sheet_name, rows) in enumerate(wb[:40]):
        actual=[(i,j,x) for i,r in enumerate(rows) for j,x in enumerate(r) if x is not None and str(x).strip()!='']
        for i,j,v in actual:
            cells.append({'source_file':file.name,'sha256':sha,'sheet':sheet_name,'row_1based':i+1,'col_1based':j+1,'value':valstr(v),'value_type':type(v).__name__})
        st=' '.join(valstr(x) for r in rows[:3] for x in r[:5])
        # Sheet title is more reliable than secondary/reference dates in header text.
        month=month_parse(sheet_name) or month_parse(st)
        current=('new business statement' in norm(st) and ('as at' in norm(sheet_name) or 'new business' in norm(sheet_name)) and 'हिंदी' not in sheet_name)
        metadata.append({'file':file.name,'sheet':sheet_name,'sheet_month':month or '', 'populated_cells':len(actual),'current_new_business_candidate':current})
        if not current:continue
        # Reporting month is determined from workbook/sheet content ONLY.
        # Filenames are arbitrary and never used as a reason to reject data.
        if not month:
            issues.append({'file':file.name,'sheet':sheet_name,'issue':'Cannot detect reporting month'});continue
        try: hdr, columns=header_metrics(rows)
        except Exception as e:
            issues.append({'file':file.name,'sheet':sheet_name,'issue':str(e)});continue
        insurer=None;section=''
        for r_idx in range(hdr+1,len(rows)):
            r=rows[r_idx];rawlabel=valstr(r[1] if len(r)>1 else '').strip()
            if not rawlabel:continue
            kind=label_kind(rawlabel)
            if not kind:
                # insurer header rows typically have numeric S.No, plus category rows have empty S.No.
                if (r[0] is not None and str(r[0]).strip() and (str(r[0]).strip().replace('.','',1).isdigit() or norm(rawlabel) in ['grand total','private total','lic of india'])) or norm(rawlabel) in ['grand total','private total','lic of india']:
                    insurer=rawlabel
                continue
            if not insurer:
                issues.append({'file':file.name,'sheet':sheet_name,'issue':f'Category without insurer row {r_idx+1}'});continue
            for col,metric,period in columns:
                if col>=len(r) or r[col] is None or valstr(r[col]).strip()=='':continue
                v=r[col]
                structured.append({'month':month,'source_file':file.name,'sheet':sheet_name,'insurer':insurer,'category':kind,'metric':metric,'period_header':period,'value':valstr(v),'numeric_value':v if isinstance(v,(int,float)) and not isinstance(v,bool) else '', 'source_cell':f'{letters(col)}{r_idx+1}'})
        if not any(z['month']==month for z in structured):issues.append({'file':file.name,'sheet':sheet_name,'issue':'No structured detail rows found'})
    return cells,structured,metadata,issues

def letters(n):
    s='';n+=1
    while n:n,k=divmod(n-1,26);s=chr(65+k)+s
    return s

def main():
    fs=sorted(x for x in RAW.rglob('*') if x.is_file() and x.suffix.lower() in ('.xlsx','.xls') and not x.name.startswith('~$'))
    if not fs:sys.exit('No Excel (.xlsx/.xls) files in raw_data/')
    allcells=[];allvalues=[];sheets=[];issues=[]
    for f in fs:
        try:
            a,b,c,d=extract(f);allcells+=a;allvalues+=b;sheets+=c;issues+=d
            print('OK',f.name,'structured records',len(b),'archived cells',len(a),flush=True)
        except Exception as e:
            issues.append({'file':f.name,'sheet':'','issue':f'IMPORT ERROR {e}'});print('CHECK',f.name,e)
    save('all_workbook_cells.csv',allcells,['source_file','sha256','sheet','row_1based','col_1based','value','value_type'])
    save('all_reported_metrics.csv',allvalues,['month','source_file','sheet','insurer','category','metric','period_header','value','numeric_value','source_cell'])
    save('workbook_inventory.csv',sheets,['file','sheet','sheet_month','populated_cells','current_new_business_candidate'])
    # Key selected monthly metrics: preserve detailed values in long format; do not silently invent missing values.
    current=[]
    for z in allvalues:
        header=norm(z['period_header']);m=z['month'];mon=int(m[-2:]);year=int(m[:4]);
        match=re.match(r'for\s+(?:the\s+month\s+of\s+)?([a-z]+)\s*,?\s*(20\d{2})',header)
        if match and MONTH_ALIASES.get(match.group(1))==mon and int(match.group(2))==year and 'growth' not in header:
            current.append(z)
    save('current_month_detail.csv',current,['month','source_file','sheet','insurer','category','metric','period_header','value','numeric_value','source_cell'])
    # main analytic series: grand total rows for two individual categories for policy count metric.
    grouped={}; seen={}
    for z in current:
        if norm(z['insurer'])!='grand total' or z['metric']!='number_of_policies_schemes' or z['category'] not in ('individual_single','individual_non_single'):continue
        try:v=float(z['numeric_value']);assert v>=0 and v.is_integer()
        except:issues.append({'file':z['source_file'],'sheet':z['sheet'],'issue':'Invalid national individual policy count '+z['source_cell']});continue
        m=z['month']; k=(m,z['category'])
        if k in seen:issues.append({'file':z['source_file'],'sheet':z['sheet'],'issue':'Duplicate month+category '+str(k)});continue
        seen[k]=z['source_file'];grouped.setdefault(m,{})[z['category']]=int(v)
    summary=[]
    for m in sorted(grouped):
        a=grouped[m];ok=all(k in a for k in ('individual_single','individual_non_single'))
        summary.append({'month':m,'individual_single':a.get('individual_single',''),'individual_non_single':a.get('individual_non_single',''),'individual_total':sum(a.values()) if ok else '', 'status':'OK' if ok else 'INCOMPLETE'})
        if not ok: issues.append({'file':'','sheet':'','issue':'Missing grand-total component '+m})
    save('insurance_monthly.csv',summary,['month','individual_single','individual_non_single','individual_total','status'])
    # Compare recognized metrics across reports without inventing missing values.
    # Grand Total combinations are the core audit keys; any additional metrics
    # observed in any month are retained, and gaps get their own report.
    expected=set((z['category'],z['metric']) for z in current if norm(z['insurer'])=='grand total')
    observed={}
    source_by_month={}
    for z in current:
        if norm(z['insurer'])=='grand total':
            observed.setdefault(z['month'],set()).add((z['category'],z['metric']))
            source_by_month.setdefault(z['month'],set()).add(z['source_file'])
    missing=[]
    for m in sorted(observed):
        for category,metric in sorted(expected-observed[m]):
            missing.append({'month':m,'source_files':'; '.join(sorted(source_by_month[m])),
                'category':category,'metric':metric,'reason':'Not found in parsed current-month Grand Total'})
    save('missing_columns.csv',missing,['month','source_files','category','metric','reason'])
    # Report all duplicates explicitly, including duplicate reports for a single month.
    month_files={}
    for z in sheets:
        if z['current_new_business_candidate'] and z['sheet_month']:
            month_files.setdefault(z['sheet_month'],set()).add(z['file'])
    duplicates=[{'month':m,'file_count':len(files),'files':'; '.join(sorted(files))}
        for m,files in sorted(month_files.items()) if len(files)>1]
    save('duplicate_months.csv',duplicates,['month','file_count','files'])
    for z in duplicates:
        issues.append({'file':z['files'],'sheet':'','issue':'Multiple source files claim reporting month '+z['month']})
    save('extraction_issues.csv',issues,['file','sheet','issue'])
    print(f'DONE: {len(fs)} files, {len(summary)} months, {len(current)} current metrics, {len(allcells)} archived cells, {len(issues)} issues, {len(missing)} missing metrics, {len(duplicates)} duplicate months')
    if issues:print('Inspect output/extraction_issues.csv before using series')
if __name__=='__main__':main()
