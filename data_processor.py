"""
data_processor.py  —  Invesmate Analytics Dashboard
Parses all files and returns JSON-serialisable dicts for HTML templates.

Offline section requires 3 files:
  1. Seminar Updated Sheet  — attendance, seat-book, student info
  2. Conversion List        — orders, payments, courses purchased
  3. Leads Report           — lead source, campaign, stage, owner, etc.

Online section requires:
  4. Free_Class_Lead_Report.xlsx — BCMB + INSG sheets
"""

import pandas as pd
import numpy as np
import re
from datetime import datetime

# ──────────────────────────────────────────────────────────────────────────────
# LOW-LEVEL HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def _n(val):
    try:
        v = float(val)
        return 0 if (np.isnan(v) or np.isinf(v)) else v
    except Exception:
        return 0

def _s(val):
    if val is None:
        return ''
    try:
        if isinstance(val, float) and np.isnan(val):
            return ''
    except Exception:
        pass
    return str(val).strip()

def _d(val):
    try:
        if pd.isna(val):
            return ''
    except Exception:
        pass
    try:
        if isinstance(val, (datetime, pd.Timestamp)):
            return val.strftime('%Y-%m-%d')
        s = str(val).strip()
        if '/' in s:
            s = re.split(r'\s{2,}', s)[0].strip()
            return pd.to_datetime(s, dayfirst=True).strftime('%Y-%m-%d')
        return s[:10] if len(s) >= 10 else ''
    except Exception:
        return ''

def _col(df, *keywords, exact=False, exclude=None):
    excl = [e.lower() for e in (exclude or [])]
    cols = [str(c) for c in df.columns]
    for kw in keywords:
        kw_l = kw.lower()
        for c in cols:
            c_l = c.lower()
            if any(e in c_l for e in excl):
                continue
            if (exact and c_l == kw_l) or (not exact and kw_l in c_l):
                return c
    return None

def safe_numeric(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)

def clean_mobile(x):
    if pd.isna(x):
        return None
    s = re.sub(r'\D', '', str(x))
    return s[-10:] if len(s) >= 10 else None

def parse_date_series(series):
    for fmt in ['%d-%b-%Y','%d-%b-%y','%d/%m/%Y','%Y-%m-%d','%d-%m-%Y','%b-%d-%Y','%d %b %Y']:
        try:
            parsed = pd.to_datetime(series, format=fmt, errors='coerce')
            if parsed.notna().any():
                return parsed
        except Exception:
            pass
    return pd.to_datetime(series, errors='coerce', dayfirst=True)

def normalize_status(status):
    if not status:
        return ''
    s = str(status).strip().lower()
    if s in ['paid','completed','success','active','converted']:
        return 'Active'
    if s in ['partial','partially paid','in progress']:
        return 'Partially Converted'
    if s in ['failed','cancelled','canceled','inactive','pending']:
        return 'Inactive'
    return str(status).strip()

COMBO_MATCH = 'Power Of Trading & Investing Combo Course'

# ──────────────────────────────────────────────────────────────────────────────
# FILE LOADER
# ──────────────────────────────────────────────────────────────────────────────

def _load_file(file_obj, name=''):
    name = (name or '').lower()
    try:
        if name.endswith('.csv'):
            try:
                return pd.read_csv(file_obj)
            except Exception:
                file_obj.seek(0)
                return pd.read_csv(file_obj, encoding='latin1')
        if name.endswith(('.xlsx','.xls')):
            return pd.read_excel(file_obj, sheet_name=0)
        try:
            return pd.read_excel(file_obj, sheet_name=0)
        except Exception:
            try: file_obj.seek(0)
            except Exception: pass
            try:
                return pd.read_csv(file_obj)
            except Exception:
                file_obj.seek(0)
                return pd.read_csv(file_obj, encoding='latin1')
    except Exception as e:
        raise ValueError(f'Error reading file ({name}): {e}') from e

def _detect(df, *candidates):
    norm = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in norm:
            return norm[key]
    return None

# ──────────────────────────────────────────────────────────────────────────────
# TRAINER NORMALISATION (online files)
# ──────────────────────────────────────────────────────────────────────────────

TRAINER_MAP = {
    'rohitava majumdar':'Rohitava Majumder','rohitav majumder':'Rohitava Majumder',
    'rohitava majumder**':'Rohitava Majumder','debargha  saha':'Debargho Saha',
    'debargha saha':'Debargho Saha','debargho\u00a0saha':'Debargho Saha',
    'pratim kumer chakraborty':'Pratim Kumar Chakraborty',
    'hironmoy laheri':'Hironmoy Lahiri','hironmoy lahiri\u00a0':'Hironmoy Lahiri',
    'sandipan das':'Sandipan Kumar Das',
    'kunal saha (special advanced class)':'Kunal Saha',
    'sayan sarker(special advanced class)':'Sayan Sarker',
}

def _norm_trainer(raw):
    parts = [p.strip() for p in re.split(r',|&|\n', str(raw)) if p.strip()]
    out = []
    for p in parts:
        p = re.sub(r'\s*\(Special Advanced Class\)\s*','',p,flags=re.I).strip()
        p = re.sub(r'\s+',' ',p)
        out.append(TRAINER_MAP.get(p.lower(),p))
    return ', '.join(dict.fromkeys(out))

# ──────────────────────────────────────────────────────────────────────────────
# ONLINE FILE — BCMB + INSIGNIA
# ──────────────────────────────────────────────────────────────────────────────

_SHEET_SKIP = {
    'log','hitting','call','re-target','retarget','backup','rough','comparison',
    'summary','offline','forx','fund','hindi','invesmeet','simplify','monitoring',
    'lead wise','joining','percentage','day to day','sheet1','8_45','sunday','tuesday','friday',
}

def _pick_sheet(xl, keyword):
    candidates = [s for s in xl.sheet_names
                  if keyword in s.lower() and not any(sk in s.lower() for sk in _SHEET_SKIP)]
    candidates.sort(key=len)
    return candidates[0] if candidates else None

def _parse_bcmb(xl, sheet_name):
    if not sheet_name: return []
    df = xl.parse(sheet_name, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    c_trainer = _col(df,'trainer',exact=True) or _col(df,'trainer',exclude=['re-target','retarget'])
    c_type    = _col(df,'type',exact=True) or _col(df,'location',exact=True)
    c_date    = _col(df,'date',exact=True) or _col(df,'date',exclude=['web','hitting','batch'])
    c_tgt     = _col(df,'targeted',exact=True) or _col(df,'targeted',exclude=['to','%','re-','retarget','dialed','visited','regist','over','seat','new','old'])
    c_reg     = _col(df,'registered',exact=True) or _col(df,'registered',exclude=['%','to'])
    c_o30     = _col(df,'over 30 min',exact=True) or _col(df,'over 30',exclude=['%','to'])
    c_sb      = _col(df,'seat booked',exact=True) or _col(df,'seat booked',exclude=['%','to','amount'])
    c_join    = _col(df,'total joined',exact=True) or _col(df,'joined',exclude=['%','re-','new','old','semi'])
    c_rev     = _col(df,'seat booking amount') or _col(df,'course amount')
    records = []
    for _, row in df.iterrows():
        dv = _d(row.get(c_date,'')) if c_date else ''
        tg = int(_n(row.get(c_tgt,0))) if c_tgt else 0
        if not dv or tg < 1: continue
        tr = _norm_trainer(_s(row.get(c_trainer,'Unknown')) if c_trainer else 'Unknown')
        ty = _s(row.get(c_type,'Live')) if c_type else 'Live'
        rg = int(_n(row.get(c_reg,0))) if c_reg else 0
        o3 = int(_n(row.get(c_o30,0))) if c_o30 else 0
        sb = int(_n(row.get(c_sb,0)))  if c_sb  else 0
        jn = int(_n(row.get(c_join,0)))if c_join else 0
        rv = int(_n(row.get(c_rev,0))) if c_rev  else 0
        t  = ty.upper()
        wt = ('Rec' if 'REC' in t else 'Backup' if 'BACKUP' in t or 'BACK' in t
              else 'Practice' if 'PRACTICE' in t else 'Cancel' if 'CANCEL' in t
              else 'Live\n(ZOOM)' if 'ZOOM' in t else 'Live')
        if rv == 0 and sb > 0: rv = sb * 5632
        records.append({'date':dv,'yearMonth':dv[:7],'trainer':tr,'course':'BCMB','type':wt,
                        'mode':'Online','targeted':tg,'registered':rg,'over30':o3,
                        'seatBooked':sb,'joined':jn,'revenue':rv,'expenses':0,'surplus':rv})
    return sorted(records, key=lambda r: r['date'])

def _parse_insg(xl, sheet_name):
    if not sheet_name: return []
    df = xl.parse(sheet_name, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    c_trainer = _col(df,'trainer',exact=True)
    c_type    = _col(df,'type',exact=True)
    c_date    = _col(df,'date',exact=True) or _col(df,'date',exclude=['web','hitting','hidden','batch'])
    c_tgt     = _col(df,'targated',exact=True) or _col(df,'targeted',exact=True) or _col(df,'targated',exclude=['%','to'])
    c_reg     = _col(df,'registered',exact=True) or _col(df,'registered',exclude=['%','to'])
    c_o30     = _col(df,'over 30 min',exact=True) or _col(df,'over 30',exclude=['%','to'])
    c_sb      = _col(df,'seat booked',exact=True) or _col(df,'seat booked',exclude=['%','to'])
    c_join    = _col(df,'unique viewer') or _col(df,'total joined') or _col(df,'joined',exclude=['%'])
    records = []
    for _, row in df.iterrows():
        dv = _d(row.get(c_date,'')) if c_date else ''
        tg = int(_n(row.get(c_tgt,0))) if c_tgt else 0
        if not dv or tg < 1: continue
        tr = _norm_trainer(_s(row.get(c_trainer,'Unknown')) if c_trainer else 'Unknown')
        ty = _s(row.get(c_type,'Live')) if c_type else 'Live'
        sb = int(_n(row.get(c_sb,0))) if c_sb else 0
        records.append({'date':dv,'yearMonth':dv[:7],'trainer':tr,'course':'INSIGNIA',
                        'type':'Rec' if 'REC' in ty.upper() else 'Live','mode':'Online',
                        'targeted':tg,'registered':int(_n(row.get(c_reg,0))) if c_reg else 0,
                        'over30':int(_n(row.get(c_o30,0))) if c_o30 else 0,
                        'seatBooked':sb,'joined':int(_n(row.get(c_join,0))) if c_join else 0,
                        'revenue':sb*8999,'expenses':0,'surplus':sb*8999})
    return sorted(records, key=lambda r: r['date'])

def parse_webinar_file(file_obj):
    xl = pd.ExcelFile(file_obj)
    return (_parse_bcmb(xl, _pick_sheet(xl,'bcmb')),
            _parse_insg(xl, _pick_sheet(xl,'insg') or _pick_sheet(xl,'insignia')))

# ──────────────────────────────────────────────────────────────────────────────
# OFFLINE FILES — Student matching (mirrors app.py logic exactly)
# ──────────────────────────────────────────────────────────────────────────────

def normalize_sales_rep(raw):
    s = _s(raw)
    if not s or s.lower() in {'nan','none','null'}:
        return ''
    s = re.sub(r'\s+', ' ', s).strip(' -,:;')
    if not s:
        return ''
    return ' '.join(part.upper() if len(part) <= 3 and part.isalpha() else part.capitalize() for part in s.split())

def normalize_course_name(raw):
    s = _s(raw)
    if not s or s.lower() in {'nan','none','null'}:
        return ''
    s = re.sub(r'\s+', ' ', s.replace(' ',' ')).strip(' -,:;')
    return s

def compute_order_due(total_amount, paid_amount, total_due):
    total = _n(total_amount)
    paid = _n(paid_amount)
    due = _n(total_due)
    if total > 0:
        derived = round(max(total - paid, 0), 2)
        if due <= 0 and paid < total:
            return derived
        if abs(due - derived) > 1 and derived > 0:
            return derived
    return round(max(due, 0), 2)

def compute_student_financials(order_df):
    if order_df is None or order_df.empty:
        return {'total_fees':0.0,'total_paid':0.0,'total_due':0.0,'order_count':0,
                'converted':False,'fully_paid':False,'has_partial':False}
    total_fees = float(order_df['effective_total_amount'].fillna(0).sum())
    total_paid = float(order_df['paid_amount'].fillna(0).sum())
    total_due  = float(order_df['effective_due'].fillna(0).sum())
    # trust recomputed due over status labels
    total_due = round(max(total_due, 0.0), 2)
    total_fees = round(max(total_fees, total_paid + total_due), 2)
    fully_paid = total_paid > 0 and total_due <= 1
    return {
        'total_fees': total_fees,
        'total_paid': round(total_paid, 2),
        'total_due': total_due,
        'order_count': int(len(order_df)),
        'converted': fully_paid,
        'fully_paid': fully_paid,
        'has_partial': total_paid > 0 and not fully_paid,
    }

def _build_lead_mapping(leads_file=None, leads_name=''):
    lead_map = pd.DataFrame()
    meta = {'mobile':None,'convfrom':None,'source':None,'campaign':None,'status':None,
            'stage':None,'owner':None,'state':None,'attempted':None,'service':None,
            'email':None,'remarks':None,'name':None}
    if leads_file is None:
        return lead_map, meta, []
    leads = _load_file(leads_file, leads_name)
    leads.columns = [str(c).strip() for c in leads.columns]
    meta['mobile']    = _detect(leads,'phone','Phone','mobile','Mobile','Contact','Contact Number','Phone Number')
    meta['convfrom']  = _detect(leads,'converted_from','ConvertedFrom','lead_type','LeadType')
    meta['source']    = _detect(leads,'leadsource','lead_source','LeadSource','Source')
    meta['campaign']  = _detect(leads,'campaign_name','Campaign','CampaignName')
    meta['status']    = _detect(leads,'leadstatus','lead_status','LeadStatus','Status')
    meta['stage']     = _detect(leads,'stage_name','StageName','Stage')
    meta['owner']     = _detect(leads,'leadownername','LeadOwner','lead_owner','Owner')
    meta['state']     = _detect(leads,'state','State','Province')
    meta['attempted'] = _detect(leads,'Attempted/Unattempted','attempted','Attempted')
    meta['service']   = _detect(leads,'servicename','ServiceName','service_name')
    meta['email']     = _detect(leads,'email','Email')
    meta['remarks']   = _detect(leads,'remarks','Remarks','Notes')
    meta['name']      = _detect(leads,'name','Name','StudentName')
    if meta['mobile']:
        leads['mobile_clean'] = leads[meta['mobile']].apply(clean_mobile)
        if meta.get('attempted'):
            leads['_attempt_rank'] = leads[meta['attempted']].astype(str).str.lower().map({'attempted':1,'unattempted':0}).fillna(0)
            leads = leads.sort_values(['_attempt_rank'], ascending=False)
        lead_map = leads.dropna(subset=['mobile_clean']).drop_duplicates('mobile_clean', keep='first').set_index('mobile_clean')
    return lead_map, meta, leads.to_dict('records')

def _lead_row_from_index(row, meta):
    def gs(key):
        col = meta.get(key)
        return _s(row[col]) if col and col in row.index and pd.notna(row[col]) else ''
    wt = gs('convfrom')
    src = gs('source')
    if not wt:
        src_u = src.upper()
        wt = 'Webinar' if ('WBN' in src_u or 'WEBINAR' in src_u) else ('Non Webinar' if src else '')
    return {
        'webinar_type': wt,
        'lead_source': gs('source'),
        'campaign_name': gs('campaign'),
        'lead_status': gs('status'),
        'stage_name': gs('stage'),
        'lead_owner': gs('owner'),
        'state': gs('state'),
        'attempted': gs('attempted'),
        'service_name_lead': gs('service'),
        'email': gs('email'),
        'remarks': gs('remarks'),
        'lead_name': gs('name'),
    }

def _blank_lead_row():
    return {'webinar_type':'','lead_source':'','campaign_name':'','lead_status':'','stage_name':'',
            'lead_owner':'','state':'','attempted':'','service_name_lead':'','email':'',
            'remarks':'','lead_name':''}

def _get_lead(possible_mobiles, lead_map, meta):
    if lead_map.empty:
        return _blank_lead_row()
    for mob in possible_mobiles:
        if mob and mob in lead_map.index:
            row = lead_map.loc[mob]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
            return _lead_row_from_index(row, meta)
    return _blank_lead_row()

def parse_offline_files(seminar_updated_file, conversion_file, leads_file,
                        sem_name='', conv_name='', leads_name=''):
    """Parse Seminar Updated Sheet + Conversion List + Leads Report."""
    if seminar_updated_file is None or conversion_file is None:
        return [], [], [], {}

    try:
        sem = _load_file(seminar_updated_file, sem_name)
        sem.columns = [str(c).strip() for c in sem.columns]
    except Exception as e:
        return [], [], [], {'error': f'Seminar file: {e}'}

    c_mobile   = _detect(sem,'Mobile','Phone','mobile','phone','Contact')
    c_altmob   = _detect(sem,'Alternate Number','Alt Mobile','alternate_number','Alternate Mobile','Alternative Mobile')
    c_name     = _detect(sem,'NAME','Name','Student Name','name')
    c_place    = _detect(sem,'Place','Location','Venue','City','place')
    c_trainer  = _detect(sem,'Trainer / Presenter','Trainer','Presenter','trainer')
    c_semdate  = _detect(sem,'Seminar Date','Date','seminar_date','Event Date')
    c_session  = _detect(sem,'Session','session','Batch','Time')
    c_attended = _detect(sem,'Is Attended ?','Is Attended?','Attended','attended','is_attended','IsAttended','ATTENDED','Attendance','Present','attend')
    c_amount   = _detect(sem,'Amount Paid','Seat Amount','seat_amount','amount_paid','SeatAmount')

    sem['mobile_clean']     = sem[c_mobile].apply(clean_mobile) if c_mobile else None
    sem['alt_mobile_clean'] = sem[c_altmob].apply(clean_mobile) if c_altmob else None
    sem['seminar_date']     = parse_date_series(sem[c_semdate]) if c_semdate else pd.NaT
    sem['seat_book_amount'] = safe_numeric(sem[c_amount]) if c_amount else 0
    sem['attended_flag']    = (sem[c_attended].astype(str).str.strip().str.upper().isin(['YES','TRUE','1','Y'])) if c_attended else False

    attendees = sem[
        ((sem['mobile_clean'].notna()) | (sem['alt_mobile_clean'].notna()))
    ].copy()
    if c_name:
        attendees = attendees[(attendees[c_name].notna()) | (attendees['attended_flag']) | (attendees['seat_book_amount'] > 0)]
    attendees = attendees.drop_duplicates(subset=['mobile_clean','alt_mobile_clean','seminar_date','seat_book_amount'], keep='first').reset_index(drop=True)

    try:
        conv = _load_file(conversion_file, conv_name)
        conv.columns = [str(c).strip() for c in conv.columns]
    except Exception as e:
        return [], [], [], {'error': f'Conversion file: {e}'}

    cc_mobile   = _detect(conv,'phone','Phone','mobile','Mobile','Contact')
    cc_service  = _detect(conv,'service_name','Service Name','Course','course_name','ServiceName')
    cc_orderdt  = _detect(conv,'order_date','Order Date','OrderDate','Date')
    cc_payrec   = _detect(conv,'payment_received','Payment Received','PaymentReceived','amount_paid')
    cc_gst      = _detect(conv,'total_gst','GST','gst','TotalGST')
    cc_due      = _detect(conv,'total_due','Due','total_due_amount','TotalDue')
    cc_trainer  = _detect(conv,'trainer','Trainer')
    cc_salesrep = _detect(conv,'sales_rep_name','Sales Rep','SalesRep','sales_rep')
    cc_mode     = _detect(conv,'payment_mode','Payment Mode','mode')
    cc_status   = _detect(conv,'status','Status')
    cc_orderid  = _detect(conv,'orderID','Order ID','order_id','OrderId')
    cc_total    = _detect(conv,'total_amount','Total Amount','TotalAmount','course_amount','CourseAmount','total_fees','TotalFees','course_fee','CourseFee','total','Total')

    conv['mobile_clean']       = conv[cc_mobile].apply(clean_mobile) if cc_mobile else None
    conv['order_date_clean']   = (pd.to_datetime(conv[cc_orderdt], errors='coerce', utc=True).dt.tz_localize(None) if cc_orderdt else pd.NaT)
    conv['paid_amount']        = safe_numeric(conv[cc_payrec]) if cc_payrec else 0
    conv['total_gst']          = safe_numeric(conv[cc_gst]) if cc_gst else 0
    conv['total_due']          = safe_numeric(conv[cc_due]) if cc_due else 0
    conv['total_amount']       = safe_numeric(conv[cc_total]) if cc_total else 0
    conv['effective_total_amount'] = conv['total_amount']
    conv['effective_due']      = [compute_order_due(t, p, d) for t, p, d in zip(conv['total_amount'], conv['paid_amount'], conv['total_due'])]
    conv['service_name_clean'] = conv[cc_service].apply(normalize_course_name) if cc_service else ''
    conv['trainer_clean']      = conv[cc_trainer].apply(_norm_trainer) if cc_trainer else ''
    conv['sales_rep_clean']    = conv[cc_salesrep].apply(normalize_sales_rep) if cc_salesrep else ''
    conv['payment_mode_clean'] = conv[cc_mode].astype(str).str.strip() if cc_mode else ''
    conv['status_clean']       = conv[cc_status].astype(str).str.strip() if cc_status else ''
    conv['order_id_clean']     = conv[cc_orderid].astype(str).str.strip() if cc_orderid else ''

    lead_map = pd.DataFrame()
    lead_meta = {}
    try:
        lead_map, lead_meta, _ = _build_lead_mapping(leads_file, leads_name)
    except Exception:
        lead_map, lead_meta = pd.DataFrame(), {}

    student_rows = []
    order_rows = []

    for _, row in attendees.iterrows():
        mob = row.get('mobile_clean')
        alt_mob = row.get('alt_mobile_clean')
        possible = [m for m in [mob, alt_mob] if m]
        sem_dt = row['seminar_date']
        lead_data = _get_lead(possible, lead_map, lead_meta)
        entry = {
            'name': _s(row.get(c_name,'')) if c_name else '',
            'mobile': mob or alt_mob or '',
            'place': _s(row.get(c_place,'')) if c_place else '',
            'trainer': _norm_trainer(_s(row.get(c_trainer,'')) if c_trainer else ''),
            'seminar_date': sem_dt.strftime('%Y-%m-%d') if pd.notna(sem_dt) else '',
            'seminar_month': sem_dt.strftime('%Y-%m') if pd.notna(sem_dt) else '',
            'session': _s(row.get(c_session,'')).upper() if c_session else '',
            'attended': bool(row.get('attended_flag',False)),
            'seat_book_amount': float(row.get('seat_book_amount',0) or 0),
            'seat_booked': bool(float(row.get('seat_book_amount',0) or 0) > 0),
            'primary_course':'','primary_order_date':'','primary_paid':0.0,'primary_total':0.0,'primary_due':0.0,
            'primary_gst':0.0,'primary_mode':'','primary_status':'','additional_courses':[],
            'additional_paid':0.0,'additional_due':0.0,'converted':False,'has_partial':False,
            'sales_rep':'','match_reason':'',
            'total_paid':0.0,'total_due':0.0,'total_fees':0.0,'order_count':0,'fully_paid':False,
        }
        entry.update(lead_data)

        all_orders = (conv[conv['mobile_clean'].isin(possible)].sort_values(['order_date_clean','paid_amount','effective_total_amount'], ascending=[True,False,False]).copy()) if possible else pd.DataFrame()
        if not all_orders.empty:
            all_orders = all_orders[~((all_orders['paid_amount']<=0) & (all_orders['effective_total_amount']<=0) & (all_orders['effective_due']<=0) & (all_orders['service_name_clean']==''))]
        if not possible:
            entry['match_reason'] = 'No mobile'
        elif all_orders.empty:
            entry['match_reason'] = 'No conversion row'
        elif pd.notna(sem_dt) and all_orders['order_date_clean'].notna().any():
            entry['match_reason'] = 'Matched (post seminar)' if (all_orders['order_date_clean'] >= sem_dt).any() else 'Matched (pre seminar)'
        else:
            entry['match_reason'] = 'Matched'

        if not all_orders.empty:
            valid_after = all_orders[(all_orders['order_date_clean'] >= sem_dt)] if pd.notna(sem_dt) and all_orders['order_date_clean'].notna().any() else pd.DataFrame()
            primary_pool = valid_after if not valid_after.empty else all_orders
            primary_pool = primary_pool.sort_values(['paid_amount','effective_total_amount','order_date_clean'], ascending=[False,False,True])
            pti_pool = primary_pool[primary_pool['service_name_clean'].str.contains(COMBO_MATCH, na=False, case=False)]
            primary = pti_pool.iloc[0] if not pti_pool.empty else primary_pool.iloc[0]
            fin = compute_student_financials(all_orders)
            entry.update(fin)
            entry['converted'] = fin['converted']
            entry['fully_paid'] = fin['fully_paid']
            entry['has_partial'] = fin['has_partial']
            entry['primary_course'] = primary['service_name_clean']
            entry['primary_order_date'] = primary['order_date_clean'].strftime('%Y-%m-%d') if pd.notna(primary['order_date_clean']) else ''
            entry['primary_paid'] = float(primary['paid_amount'])
            entry['primary_total'] = float(primary['effective_total_amount'])
            entry['primary_due'] = float(primary['effective_due'])
            entry['primary_gst'] = float(primary['total_gst'])
            entry['primary_mode'] = _s(primary['payment_mode_clean'])
            entry['primary_status'] = normalize_status(primary['status_clean'])
            entry['sales_rep'] = normalize_sales_rep(primary['sales_rep_clean'])
            others = all_orders[all_orders.index != primary.name].copy()
            entry['additional_courses'] = [{
                'course': _s(o['service_name_clean']),
                'paid': float(o['paid_amount']),
                'due': float(o['effective_due']),
                'gst': float(o['total_gst']),
                'mode': _s(o['payment_mode_clean']),
                'status': normalize_status(o['status_clean']),
                'order_date': o['order_date_clean'].strftime('%Y-%m-%d') if pd.notna(o['order_date_clean']) else '',
                'sales_rep': normalize_sales_rep(o['sales_rep_clean']),
                'order_id': _s(o['order_id_clean']),
            } for _, o in others.iterrows()]
            entry['additional_paid'] = float(others['paid_amount'].sum())
            entry['additional_due'] = float(others['effective_due'].sum())
            for _, o in all_orders.iterrows():
                order_rows.append({
                    'name': entry['name'], 'mobile': entry['mobile'], 'place': entry['place'],
                    'seminar_date': entry['seminar_date'], 'course': _s(o['service_name_clean']),
                    'order_date': o['order_date_clean'].strftime('%Y-%m-%d') if pd.notna(o['order_date_clean']) else '',
                    'order_month': o['order_date_clean'].strftime('%Y-%m') if pd.notna(o['order_date_clean']) else '',
                    'paid_amount': float(o['paid_amount']), 'total_amount': float(o['effective_total_amount']),
                    'total_due': float(o['effective_due']), 'total_gst': float(o['total_gst']),
                    'payment_mode': _s(o['payment_mode_clean']), 'status': normalize_status(o['status_clean']),
                    'sales_rep': normalize_sales_rep(o['sales_rep_clean']), 'is_primary': bool(o.name == primary.name),
                    'order_id': _s(o['order_id_clean']), 'trainer': _s(o['trainer_clean']) or entry['trainer'],
                })
        student_rows.append(entry)

    meta_map = {}
    for s in student_rows:
        key = (s['seminar_date'], s['place'])
        if key not in meta_map:
            meta_map[key] = {'date':s['seminar_date'],'month':s['seminar_month'],'place':s['place'],
                             'location':s['place'],'trainer':s['trainer'],'session':s['session'],
                             'total':0,'attended':0,'seat_booked':0,'sb_seminar':0,
                             'seat_book_amount':0.0,'converted':0,'paid':0.0,'due':0.0,
                             'actual_revenue':0.0,'expenses':0.0}
        m = meta_map[key]
        m['total'] += 1
        m['attended'] += 1 if s['attended'] else 0
        m['seat_booked'] += 1 if s['seat_booked'] else 0
        m['sb_seminar'] += 1 if s['seat_booked'] else 0
        m['seat_book_amount'] += s['seat_book_amount']
        m['converted'] += 1 if s['converted'] else 0
        m['paid'] += s['total_paid']
        m['due'] += s['total_due']
        m['actual_revenue'] += s['total_paid']
    seminar_meta = sorted(meta_map.values(), key=lambda r: r['date'])

    total = len(student_rows)
    conv_count = sum(1 for s in student_rows if s['converted'])
    t_paid = sum(s['total_paid'] for s in student_rows)
    t_due = sum(s['total_due'] for s in student_rows)

    course_stats = {}
    for o in order_rows:
        c = normalize_course_name(o.get('course'))
        if not c:
            continue
        if c not in course_stats:
            course_stats[c] = {'count':0,'paid':0.0,'due':0.0,'fully_paid':0,'is_primary':False,'total_amount':0.0}
        course_stats[c]['count'] += 1
        course_stats[c]['paid'] += float(o.get('paid_amount',0) or 0)
        course_stats[c]['due'] += float(o.get('total_due',0) or 0)
        course_stats[c]['total_amount'] += float(o.get('total_amount',0) or 0)
        if float(o.get('total_due',0) or 0) <= 1 and float(o.get('paid_amount',0) or 0) > 0:
            course_stats[c]['fully_paid'] += 1
        if o.get('is_primary'):
            course_stats[c]['is_primary'] = True
    course_stats = dict(sorted(course_stats.items(), key=lambda x: -x[1]['paid']))

    sr_stats = {}
    for o in order_rows:
        rep = normalize_sales_rep(o.get('sales_rep'))
        if not rep:
            continue
        if rep not in sr_stats:
            sr_stats[rep] = {'deals':0,'revenue':0.0,'due':0.0,'active':0,'avg_deal':0.0,'courses':set()}
        sr_stats[rep]['deals'] += 1
        sr_stats[rep]['revenue'] += float(o.get('paid_amount',0) or 0)
        sr_stats[rep]['due'] += float(o.get('total_due',0) or 0)
        if normalize_status(o.get('status','')) == 'Active':
            sr_stats[rep]['active'] += 1
        if o.get('course'):
            sr_stats[rep]['courses'].add(o['course'])
    for rep, d in sr_stats.items():
        d['avg_deal'] = round(d['revenue']/d['deals'], 2) if d['deals'] else 0
        d['courses'] = sorted(d['courses'])
        d['top_course'] = d['courses'][0] if d['courses'] else ''
    sr_stats = dict(sorted(sr_stats.items(), key=lambda x: -x[1]['revenue'])[:25])

    def _init_bucket(defaults):
        return defaults.copy()

    loc_stats, lead_src_stats, stage_stats, trainer_stats, monthly = {}, {}, {}, {}, {}
    for s in student_rows:
        loc = s['place'] or 'Unknown'
        loc_stats.setdefault(loc, {'total':0,'converted':0,'paid':0.0,'due':0.0,'seat_booked':0})
        loc_stats[loc]['total'] += 1
        loc_stats[loc]['converted'] += 1 if s['converted'] else 0
        loc_stats[loc]['paid'] += s['total_paid']
        loc_stats[loc]['due'] += s['total_due']
        loc_stats[loc]['seat_booked'] += 1 if s['seat_booked'] else 0

        src = s.get('lead_source') or 'Unknown'
        if src == 'nan': src = 'Unknown'
        lead_src_stats.setdefault(src, {'count':0,'converted':0,'revenue':0.0})
        lead_src_stats[src]['count'] += 1
        lead_src_stats[src]['converted'] += 1 if s['converted'] else 0
        lead_src_stats[src]['revenue'] += s['total_paid']

        stg = s.get('stage_name') or ''
        if stg and stg != 'nan':
            stage_stats.setdefault(stg, {'count':0,'converted':0})
            stage_stats[stg]['count'] += 1
            stage_stats[stg]['converted'] += 1 if s['converted'] else 0

        tr = s['trainer'] or 'Unknown'
        trainer_stats.setdefault(tr, {'total':0,'converted':0,'paid':0.0,'seat_booked':0})
        trainer_stats[tr]['total'] += 1
        trainer_stats[tr]['converted'] += 1 if s['converted'] else 0
        trainer_stats[tr]['paid'] += s['total_paid']
        trainer_stats[tr]['seat_booked'] += 1 if s['seat_booked'] else 0

        m = s['seminar_month'] or 'Unknown'
        monthly.setdefault(m, {'total':0,'converted':0,'paid':0.0,'seat_booked':0,'seat_amount':0.0,'due':0.0})
        monthly[m]['total'] += 1
        monthly[m]['converted'] += 1 if s['converted'] else 0
        monthly[m]['paid'] += s['total_paid']
        monthly[m]['due'] += s['total_due']
        monthly[m]['seat_booked'] += 1 if s['seat_booked'] else 0
        monthly[m]['seat_amount'] += s['seat_book_amount']

    lead_status_stats, campaign_stats = {}, {}
    for s in student_rows:
        ls = s.get('lead_status') or 'Unknown'
        if ls in ('', 'nan'): ls = 'Unknown'
        lead_status_stats.setdefault(ls, {'count':0,'converted':0})
        lead_status_stats[ls]['count'] += 1
        lead_status_stats[ls]['converted'] += 1 if s['converted'] else 0

        cp = s.get('campaign_name') or ''
        if cp and cp != 'nan':
            campaign_stats.setdefault(cp, {'count':0,'converted':0,'revenue':0.0})
            campaign_stats[cp]['count'] += 1
            campaign_stats[cp]['converted'] += 1 if s['converted'] else 0
            campaign_stats[cp]['revenue'] += s['total_paid']

    agg = {
        'total_attendees': total,
        'num_attended': sum(1 for s in student_rows if s['attended']),
        'num_seat_booked': sum(1 for s in student_rows if s['seat_booked']),
        'num_seminars': len(set((s['seminar_date'], s['place']) for s in student_rows)),
        'num_locations': len(set(s['place'] for s in student_rows if s['place'])),
        'converted': conv_count,
        'conversion_rate': round(conv_count / max(sum(1 for s in student_rows if s['seat_booked']),1) * 100, 1) if student_rows else 0,
        'attended_rate': round(sum(1 for s in student_rows if s['attended']) / total * 100, 1) if total else 0,
        'seat_to_conv_rate': round(conv_count / max(sum(1 for s in student_rows if s['seat_booked']),1) * 100, 1) if student_rows else 0,
        'total_paid': round(t_paid, 2),
        'total_due': round(t_due, 2),
        'seat_book_count': sum(1 for s in student_rows if s['seat_booked']),
        'seat_book_amount': round(sum(s['seat_book_amount'] for s in student_rows), 2),
        'fully_paid': sum(1 for s in student_rows if s['fully_paid']),
        'has_due': sum(1 for s in student_rows if s['has_partial'] or (s['total_due'] > 1 and s['total_paid'] > 0)),
        'additional_revenue': round(sum(s['additional_paid'] for s in student_rows), 2),
        'avg_paid': round(t_paid / conv_count, 2) if conv_count else 0,
        'webinar_leads': sum(1 for s in student_rows if s.get('webinar_type') == 'Webinar'),
        'non_webinar_leads': sum(1 for s in student_rows if s.get('webinar_type') == 'Non Webinar'),
        'attempted': sum(1 for s in student_rows if str(s.get('attempted','')).lower() == 'attempted'),
        'unattempted': sum(1 for s in student_rows if str(s.get('attempted','')).lower() == 'unattempted'),
        'unique_courses': len(course_stats),
        'course_stats': course_stats,
        'sales_rep_stats': sr_stats,
        'location_stats': dict(sorted(loc_stats.items(), key=lambda x: -x[1]['paid'])[:40]),
        'lead_source_stats': lead_src_stats,
        'lead_status_stats': lead_status_stats,
        'campaign_stats': campaign_stats,
        'stage_stats': stage_stats,
        'trainer_stats': trainer_stats,
        'monthly_trend': dict(sorted(monthly.items())),
    }
    return student_rows, order_rows, seminar_meta, agg

def build_online_intelligence(conversion_file=None, leads_file=None, conv_name='', leads_name=''):
    if conversion_file is None and leads_file is None:
        return [], [], {}
    online_leads = []
    online_orders = []
    lead_map, lead_meta = pd.DataFrame(), {}
    try:
        lead_map, lead_meta, _ = _build_lead_mapping(leads_file, leads_name)
    except Exception:
        pass
    if not lead_map.empty:
        for mob, row in lead_map.iterrows():
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
            lead = _lead_row_from_index(row, lead_meta)
            online_leads.append({'mobile': mob, **lead})
    def is_online_lead(lead):
        wt = (lead.get('webinar_type') or '').lower()
        src = (lead.get('lead_source') or '').lower()
        svc = (lead.get('service_name_lead') or '').lower()
        camp = (lead.get('campaign_name') or '').lower()
        blob = ' '.join([wt, src, svc, camp])
        return any(tok in blob for tok in ['webinar','wbn','free class','zoom','bcmb','insignia'])
    online_leads_filtered = [r for r in online_leads if is_online_lead(r)]
    online_mobiles = {r['mobile'] for r in online_leads_filtered if r.get('mobile')}
    if conversion_file is not None:
        conv = _load_file(conversion_file, conv_name)
        conv.columns = [str(c).strip() for c in conv.columns]
        cc_mobile   = _detect(conv,'phone','Phone','mobile','Mobile','Contact')
        cc_service  = _detect(conv,'service_name','Service Name','Course','course_name','ServiceName')
        cc_orderdt  = _detect(conv,'order_date','Order Date','OrderDate','Date')
        cc_payrec   = _detect(conv,'payment_received','Payment Received','PaymentReceived','amount_paid')
        cc_due      = _detect(conv,'total_due','Due','total_due_amount','TotalDue')
        cc_total    = _detect(conv,'total_amount','Total Amount','TotalAmount','course_amount','CourseAmount','total_fees','TotalFees','course_fee','CourseFee','total','Total')
        cc_salesrep = _detect(conv,'sales_rep_name','Sales Rep','SalesRep','sales_rep')
        cc_status   = _detect(conv,'status','Status')
        cc_mode     = _detect(conv,'payment_mode','Payment Mode','mode')
        cc_orderid  = _detect(conv,'orderID','Order ID','order_id','OrderId')
        conv['mobile_clean'] = conv[cc_mobile].apply(clean_mobile) if cc_mobile else None
        conv['paid_amount'] = safe_numeric(conv[cc_payrec]) if cc_payrec else 0
        conv['total_amount'] = safe_numeric(conv[cc_total]) if cc_total else 0
        conv['total_due'] = safe_numeric(conv[cc_due]) if cc_due else 0
        conv['effective_due'] = [compute_order_due(t,p,d) for t,p,d in zip(conv['total_amount'], conv['paid_amount'], conv['total_due'])]
        conv['order_date_clean'] = pd.to_datetime(conv[cc_orderdt], errors='coerce').dt.strftime('%Y-%m-%d') if cc_orderdt else ''
        conv['course_clean'] = conv[cc_service].apply(normalize_course_name) if cc_service else ''
        conv['sales_rep_clean'] = conv[cc_salesrep].apply(normalize_sales_rep) if cc_salesrep else ''
        conv['status_clean'] = conv[cc_status].apply(normalize_status) if cc_status else ''
        conv['payment_mode_clean'] = conv[cc_mode].astype(str).str.strip() if cc_mode else ''
        conv['order_id_clean'] = conv[cc_orderid].astype(str).str.strip() if cc_orderid else ''
        for _, r in conv.iterrows():
            mob = r.get('mobile_clean')
            lead = _get_lead([mob], lead_map, lead_meta) if mob else _blank_lead_row()
            if mob in online_mobiles or is_online_lead(lead):
                online_orders.append({
                    'mobile': mob or '', 'course': _s(r.get('course_clean')), 'order_date': _s(r.get('order_date_clean')),
                    'paid_amount': float(r.get('paid_amount',0) or 0), 'total_amount': float(r.get('total_amount',0) or 0),
                    'total_due': float(r.get('effective_due',0) or 0), 'payment_mode': _s(r.get('payment_mode_clean')),
                    'sales_rep': normalize_sales_rep(r.get('sales_rep_clean')), 'status': _s(r.get('status_clean')),
                    'order_id': _s(r.get('order_id_clean')), **lead,
                })
    if not online_leads_filtered and online_orders:
        seen = set()
        for o in online_orders:
            mob = o.get('mobile')
            if mob and mob not in seen:
                seen.add(mob)
                online_leads_filtered.append({
                    'mobile': mob,
                    'webinar_type': o.get('webinar_type') or 'Webinar',
                    'lead_source': o.get('lead_source') or 'Conversion List',
                    'campaign_name': o.get('campaign_name') or '',
                    'lead_status': o.get('lead_status') or o.get('status') or '',
                    'stage_name': o.get('stage_name') or '',
                    'lead_owner': o.get('lead_owner') or '',
                    'state': o.get('state') or '',
                    'attempted': o.get('attempted') or '',
                    'service_name_lead': o.get('service_name_lead') or o.get('course') or '',
                    'email': o.get('email') or '',
                    'remarks': o.get('remarks') or '',
                })
        online_mobiles = {r['mobile'] for r in online_leads_filtered if r.get('mobile')}
    student_fin = {}
    for o in online_orders:
        mob = o.get('mobile') or f"order-{len(student_fin)}"
        student_fin.setdefault(mob, {'paid':0.0,'due':0.0,'fees':0.0})
        student_fin[mob]['paid'] += float(o.get('paid_amount',0) or 0)
        student_fin[mob]['due'] += float(o.get('total_due',0) or 0)
        student_fin[mob]['fees'] += float(o.get('total_amount',0) or 0)
    converted_mobiles = {m for m,v in student_fin.items() if v['paid'] > 0 and v['due'] <= 1}
    partial_mobiles = {m for m,v in student_fin.items() if v['paid'] > 0 and v['due'] > 1}
    not_converted_mobiles = {m for m in online_mobiles if m not in converted_mobiles}
    src_stats, camp_stats, rep_stats, course_stats, status_stats, monthly = {}, {}, {}, {}, {}, {}
    for l in online_leads_filtered:
        src = l.get('lead_source') or 'Unknown'
        src_stats.setdefault(src, {'count':0,'converted':0,'revenue':0.0})
        src_stats[src]['count'] += 1
        if l.get('mobile') in converted_mobiles:
            src_stats[src]['converted'] += 1
            src_stats[src]['revenue'] += student_fin.get(l.get('mobile'),{}).get('paid',0)
        camp = l.get('campaign_name') or ''
        if camp:
            camp_stats.setdefault(camp, {'count':0,'converted':0,'revenue':0.0})
            camp_stats[camp]['count'] += 1
            if l.get('mobile') in converted_mobiles:
                camp_stats[camp]['converted'] += 1
                camp_stats[camp]['revenue'] += student_fin.get(l.get('mobile'),{}).get('paid',0)
        ls = l.get('lead_status') or 'Unknown'
        status_stats.setdefault(ls, {'count':0,'converted':0})
        status_stats[ls]['count'] += 1
        status_stats[ls]['converted'] += 1 if l.get('mobile') in converted_mobiles else 0
    for o in online_orders:
        c = o.get('course') or 'Unknown'
        course_stats.setdefault(c, {'count':0,'paid':0.0,'due':0.0,'total_amount':0.0})
        course_stats[c]['count'] += 1
        course_stats[c]['paid'] += float(o.get('paid_amount',0) or 0)
        course_stats[c]['due'] += float(o.get('total_due',0) or 0)
        course_stats[c]['total_amount'] += float(o.get('total_amount',0) or 0)
        rep = o.get('sales_rep') or ''
        if rep:
            rep_stats.setdefault(rep, {'deals':0,'revenue':0.0,'due':0.0})
            rep_stats[rep]['deals'] += 1
            rep_stats[rep]['revenue'] += float(o.get('paid_amount',0) or 0)
            rep_stats[rep]['due'] += float(o.get('total_due',0) or 0)
        month = (o.get('order_date') or '')[:7]
        if month:
            monthly.setdefault(month, {'paid':0.0,'due':0.0,'orders':0})
            monthly[month]['paid'] += float(o.get('paid_amount',0) or 0)
            monthly[month]['due'] += float(o.get('total_due',0) or 0)
            monthly[month]['orders'] += 1
    agg = {
        'total_online_leads': len(online_leads_filtered),
        'total_online_converted_leads': len(converted_mobiles),
        'total_online_not_converted_leads': len(not_converted_mobiles),
        'online_conversion_rate': round(len(converted_mobiles) / max(len(online_leads_filtered),1) * 100, 1) if online_leads_filtered else 0,
        'total_online_paid': round(sum(v['paid'] for v in student_fin.values()), 2),
        'total_online_due': round(sum(v['due'] for v in student_fin.values()), 2),
        'fully_paid_online_students': len(converted_mobiles),
        'partial_payment_online_students': len(partial_mobiles),
        'webinar_leads': len(online_leads_filtered),
        'non_webinar_leads': len([r for r in online_leads if not is_online_lead(r)]),
        'attempted': len([r for r in online_leads_filtered if str(r.get('attempted','')).lower() == 'attempted']),
        'unattempted': len([r for r in online_leads_filtered if str(r.get('attempted','')).lower() == 'unattempted']),
        'unique_sales_reps': len(rep_stats),
        'unique_courses': len(course_stats),
        'source_stats': src_stats,
        'campaign_stats': camp_stats,
        'sales_rep_stats': rep_stats,
        'course_stats': course_stats,
        'lead_status_stats': status_stats,
        'monthly_trend': dict(sorted(monthly.items())),
    }
    return online_leads_filtered, online_orders, agg

# ──────────────────────────────────────────────────────────────────────────────
# MAIN ENTRY POINT
# ──────────────────────────────────────────────────────────────────────────────

def process_all(webinar_file=None, seminar_file=None, attendee_file=None,
                seminar_updated_file=None, conversion_file=None, leads_file=None,
                sem_name='', conv_name='', leads_name=''):
    errors = []
    try:
        bcmb, insg = parse_webinar_file(webinar_file) if webinar_file else ([],[])
    except Exception as e:
        errors.append(f'Webinar file: {e}')
        bcmb, insg = [], []
    try:
        students, orders, seminar_meta, agg = parse_offline_files(seminar_updated_file, conversion_file, leads_file, sem_name, conv_name, leads_name)
    except Exception as e:
        errors.append(f'Offline files: {e}')
        students, orders, seminar_meta, agg = [], [], [], {}
    try:
        online_leads, online_orders, online_agg = build_online_intelligence(conversion_file, leads_file, conv_name, leads_name)
    except Exception as e:
        errors.append(f'Online intelligence: {e}')
        online_leads, online_orders, online_agg = [], [], {}
    offline_rows = [{'date':s['date'],'yearMonth':s['month'],'trainer':s['trainer'],'location':s['place'],
                     'course':'OFFLINE','type':'Offline','mode':'Offline','targeted':s['total'],'registered':s['attended'],
                     'over30':s['attended'],'seatBooked':s.get('sb_seminar', s.get('seat_booked', 0)),'joined':s.get('sb_seminar', s.get('seat_booked', 0)),
                     'revenue':s.get('actual_revenue', s.get('paid', 0)),'expenses':s.get('expenses',0),'surplus':s.get('actual_revenue', s.get('paid', 0)) - s.get('expenses',0)}
                    for s in seminar_meta]
    return {
        'bcmb': bcmb, 'insg': insg, 'offline': offline_rows, 'seminar': seminar_meta,
        'students': students, 'orders': orders, 'seminar_meta': seminar_meta, 'offline_agg': agg,
        'online_leads': online_leads, 'online_orders': online_orders, 'online_agg': online_agg,
        'att_summary': {}, 'ct_stats': agg.get('course_stats',{}), 'sr_stats': agg.get('sales_rep_stats',{}),
        'loc_stats': agg.get('location_stats',{}), 'conversion_stats': {}, 'leads_stats': {}, 'seminar_updated': [],
        'errors': errors,
        'stats': {'bcmb_count':len(bcmb),'insg_count':len(insg),'seminar_count':len(seminar_meta),
                  'locations':len(set(s['place'] for s in seminar_meta)) if seminar_meta else 0,
                  'students':agg.get('total_attendees',0),'conversions':agg.get('converted',0),
                  'leads':online_agg.get('total_online_leads',0)}
    }
