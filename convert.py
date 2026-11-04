#!/usr/bin/env python3
"""convert.py - يحوّل ملف الإكسل إلى data.enc.json (مشفّر) للداشبورد.
الاستخدام (زي ما الـ workflow بيشغّله):  python convert.py book.xlsx data.json
- الباسورد من Secret اسمه DASH_PASSWORD
- المشروع من متغير PROJECT (الافتراضي: دريم). سيبه فاضي ('') لكل المشاريع.
الملف الناتج data.enc.json بيتكتب في نفس مكان الاسكريبت."""
import base64, gzip, json, os, re, sys
import pandas as pd
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

DEFAULT_PROJECT = 'دريم'
XLSX = sys.argv[1] if len(sys.argv) > 1 else 'book.xlsx'
PW = os.environ.get('DASH_PASSWORD', '')
if len(PW) < 8:
    sys.exit('DASH_PASSWORD secret مش موجود أو قصير (8 حروف على الأقل).')
ONLY = os.environ['PROJECT'].strip() if 'PROJECT' in os.environ else DEFAULT_PROJECT
HERE = os.path.dirname(os.path.abspath(__file__))
df = pd.read_excel(XLSX, sheet_name=0, header=2)
df.columns = [str(c).strip() for c in df.columns]
df = df.iloc[:, :25]
df = df[df['اسم المشروع'].notna()].copy()
df['اسم المشروع'] = df['اسم المشروع'].astype(str).str.strip()
if ONLY: df = df[df['اسم المشروع'] == ONLY].copy()
df['t'] = df['نوع الدفعة'].fillna('').astype(str).str.strip()
stats = {}

# update date: cell Y2 of the sheet
raw2 = pd.read_excel(XLSX, sheet_name=0, header=None, nrows=2)
upd = pd.to_datetime(raw2.iloc[1, 24], errors='coerce')
if pd.isna(upd): upd = pd.Timestamp.today()
updated = upd.strftime('%Y-%m-%d')

def norm_channel(v):
    if pd.isna(v) or str(v).strip() == '': return 'غير محدد'
    toks = [x.strip() for x in re.split(r'\s+-\s+|-(?=\s*[A-Za-zء-ي])', str(v)) if x.strip()]
    out = []
    for t in toks:
        l = t.lower().replace(' ', '')
        if l.startswith('eg'):
            m = re.search(r'\d+', l); n = 'EG ' + (m.group(0).zfill(3) if m else '')
        elif 'cib' in l: n = 'CIB'
        elif 'saib' in l: n = 'Saib'
        elif 'كاش' in l: n = 'كاش'
        elif 'الكويت' in l: n = 'الكويت الوطني'
        elif 'الاهلي' in l: n = 'الاهلي المصري'
        elif 'بنكمصر' in l: n = 'بنك مصر'
        elif l == 'mid': n = 'ميد بنك'
        elif 'كابتن' in l: n = 'الكابتن مازن'
        else: n = 'أخرى'
        if n not in out: out.append(n)
    return out[0] if len(out) == 1 else 'أكثر من جهة'

def is_inst(t): return 'قسط' in t and 'صيانة' not in t and 'جراج' not in t
def is_con(t): return t in ('تعاقد', 'مقدم', 'تعاقد + جراج')

projects = df['اسم المشروع'].value_counts().index.tolist()
pidx = {p: i for i, p in enumerate(projects)}
D = lambda x: x.strftime('%Y-%m-%d')
r2 = lambda x: round(float(x), 2)

pay = df[df['t'] != ''].copy()
# ---- collected
col = pay[pay['المحصل'].fillna(0) > 0].copy()
col['d'] = col['تاريخ السداد'].fillna(col['تاريخ الاستحقاق'])
stats['collected_rows_no_date'] = int(col['d'].isna().sum())
col = col[col['d'].notna()]
inst = [[D(r['d']), r2(r['المحصل']), norm_channel(r['البنك المحصل به']), pidx[r['اسم المشروع']]]
        for _, r in col[col['t'].map(is_inst)].iterrows()]
con = [[D(r['d']), r2(r['المحصل']), norm_channel(r['البنك المحصل به']), pidx[r['اسم المشروع']],
        str(r['القطعة']).strip() if pd.notna(r['القطعة']) else '', str(r['اسم العميل']).strip() if pd.notna(r['اسم العميل']) else '']
       for _, r in col[col['t'].map(is_con)].iterrows()]

# ---- remaining installments
rem_df = pay[(pay['t'].map(is_inst) | pay['t'].str.contains('صيانة')) & (pay['المتبقى'].fillna(0) > 0)]
stats['remaining_rows_no_date'] = int(rem_df['تاريخ الاستحقاق'].isna().sum())
rem = [[D(r['تاريخ الاستحقاق']) if pd.notna(r['تاريخ الاستحقاق']) else '', r2(r['المتبقى']),
        str(r['القطعة']).strip() if pd.notna(r['القطعة']) else '',
        str(r['اسم العميل']).strip() if pd.notna(r['اسم العميل']) else '', pidx[r['اسم المشروع']],
        1 if 'صيانة' in r['t'] else 0]
       for _, r in rem_df.iterrows()]

# ---- sales: every row carrying اجمالى الوحدة (the unit total may sit on the header row
# or on the first payment row). Unit total = sum of those rows per (project, unit, client);
# matches the reference sheet "Sheet1". Contract date = first contract date, else first payment date.
df['unit'] = df['القطعة'].astype(str).str.strip()
df['client'] = df['اسم العميل'].fillna('').astype(str).str.strip()
h = df[df['اجمالى الوحدة'].fillna(0) > 0].copy()
def utype(v):
    v = str(v).strip()
    return v if v in ('فيلا', 'شقة', 'ارض') else 'غير محدد'
sales = []; nodate = 0
for (pj, unit, cl), g in h.groupby(['اسم المشروع', 'unit', 'client'], sort=False):
    d = g['تاريخ التعاقد'].dropna()
    if d.empty:
        allrows = df[(df['اسم المشروع'] == pj) & (df['unit'] == unit) & (df['client'] == cl)]
        d = allrows['تاريخ الاستحقاق'].dropna()
    if d.empty: nodate += 1; continue
    ty = g['فيلا- ارض'].map(utype); ty = ty[ty != 'غير محدد']
    sales.append([D(d.iloc[0]), unit, ty.iloc[0] if len(ty) else 'غير محدد', r2(g['اجمالى الوحدة'].sum()), cl, pidx[pj]])
stats['sales_no_date_skipped'] = nodate


# ---- maintenance & infrastructure (pure types only)
MI = {'صيانة': 0, 'الصيانة': 0, 'البنية التحتية': 1, 'بنية تحتية': 1}
mi_df = pay[pay['t'].isin(MI.keys())]
mi = [[D(r['تاريخ الاستحقاق']) if pd.notna(r['تاريخ الاستحقاق']) else '',
       r2(r['القيمة'] if pd.notna(r['القيمة']) else (r['المحصل'] if pd.notna(r['المحصل']) else 0)),
       r2(r['المحصل']) if pd.notna(r['المحصل']) else 0,
       r2(r['المتبقى']) if pd.notna(r['المتبقى']) else 0,
       str(r['القطعة']).strip() if pd.notna(r['القطعة']) else '',
       str(r['اسم العميل']).strip() if pd.notna(r['اسم العميل']) else '',
       pidx[r['اسم المشروع']], MI[r['t']]] for _, r in mi_df.iterrows()]

payload = dict(updated=updated, projects=projects, inst=inst, con=con, rem=rem, sales=sales, mi=mi)
raw = gzip.compress(json.dumps(payload, ensure_ascii=False, separators=(',', ':')).encode('utf-8'), 9)
salt, iv = os.urandom(16), os.urandom(12)
key = PBKDF2HMAC(hashes.SHA256(), 32, salt, 200000).derive(PW.encode('utf-8'))
ct = AESGCM(key).encrypt(iv, raw, None)
blob = json.dumps({'s': base64.b64encode(salt).decode(), 'i': base64.b64encode(iv).decode(), 'c': base64.b64encode(ct).decode()})

open(os.path.join(HERE, 'data.enc.json'), 'w', encoding='utf-8').write(blob)
print('OK updated', updated, 'projects', projects, 'inst', len(inst), 'con', len(con), 'rem', len(rem), 'sales', len(sales), 'mi', len(mi), stats)
