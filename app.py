import streamlit as st
import pandas as pd
import calendar
import random
import io
import os
import xlsxwriter
import base64
import requests
import time
from datetime import date, datetime

# --- 1. KONFIGURÁCIA ---
SVIATKY_2026 = {
    date(2026,1,1), date(2026,1,6), date(2026,4,3), date(2026,4,6),
    date(2026,5,1), date(2026,5,8), date(2026,7,5), date(2026,8,29),
    date(2026,9,1), date(2026,9,15), date(2026,11,1), date(2026,11,17),
    date(2026,12,24), date(2026,12,25), date(2026,12,26)
}
PRIO_LIST = ['C2', 'W1', 'W2', 'Z1', 'Z2', 'G', 'GH', 'SH']
START_REF = date(2026, 3, 1)
CYKLY = {1: "DNVDNVVV", 2: "VVDNVDNV", 3: "VDNVVVDN", 4: "NVVVDNVD"}

REPO_USER = "haluskamartin76"
REPO_NAME = "planovac-smien"
FILE_PATH = "databaza_pozicii.xlsx"

# --- 2. POMOCNÉ FUNKCIE ---
def parse_days(s):
    res = set()
    if pd.isna(s) or str(s).lower() == 'nan' or str(s).strip() == "": 
        return res
    try:
        clean_s = str(s).replace('[', '').replace(']', '').replace("'", "").replace('"', '').strip()
        parts = clean_s.replace(' ', '').replace('.0', '').replace('.', ',').split(',')
        for p in parts:
            if not p or not any(c.isdigit() for c in p): continue
            if '-' in p:
                a, b = p.split('-')
                res.update(range(int(float(a)), int(float(b)) + 1))
            else: res.add(int(float(p)))
    except: pass
    return res

def short_label(lbl):
    m = {"C1":"C", "C2":"C", "Z1":"Z", "Z2":"Z", "W1":"W", "W2":"W", "W_EXTRA":"W", "IR":"R", "IP":"K"}
    return m.get(lbl, lbl)

def moze_nastupit(idx, d, smena, poz, vysledky):
    if smena == 'D' and d > 1 and idx in vysledky[d-1]['N']: return False
    if poz == 'C1': return True
    if d > 2 and idx in vysledky[d-1][smena] and idx in vysledky[d-2][smena]: return False
    return True

def get_prioritized_people(df_db, curr_d, smena_target, hod_fond_sofar, fond_limit, is_75_poz=False):
    pool = []
    for i in range(len(df_db)):
        try:
            z_val = int(float(df_db.iloc[i]['Zmena']))
            idx = df_db.index[i]
            ma_cyk = CYKLY[z_val][(curr_d - START_REF).days % 8] == smena_target
            penalty = 10000 if hod_fond_sofar[idx] >= fond_limit else 0
            fond_score = -hod_fond_sofar[idx] if is_75_poz else hod_fond_sofar[idx]
            pool.append((idx, (0 if ma_cyk else 1, penalty, fond_score, random.random())))
        except: continue
    return [x for x in sorted(pool, key=lambda x: x)]

def load_from_github():
    try:
        # Používame raw URL pre priame stiahnutie súboru
        url = f"https://githubusercontent.com{REPO_USER}/{REPO_NAME}/main/{FILE_PATH}?t={int(time.time())}"
        res = requests.get(url)
        if res.status_code == 200:
            return io.BytesIO(res.content)
        return None
    except:
        return None

def push_to_github(df_data, df_volno):
    if "GITHUB_TOKEN" not in st.secrets:
        st.error("Chýba GITHUB_TOKEN!")
        return False
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_data.to_excel(writer, sheet_name='Data', index=False)
        df_volno.to_excel(writer, sheet_name='Volno', index=False)
    content = output.getvalue()
    token = st.secrets["GITHUB_TOKEN"]
    url = f"https://github.com{REPO_USER}/{REPO_NAME}/contents/{FILE_PATH}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    try:
        res = requests.get(url, headers=headers)
        sha = res.json().get('sha') if res.status_code == 200 else None
        payload = {
            "message": f"Update {datetime.now().strftime('%d.%m.%Y %H:%M')}",
            "content": base64.b64encode(content).decode(),
            "sha": sha
        }
        r = requests.put(url, json=payload, headers=headers)
        if r.status_code in [200, 201]:
            st.success("✅ Úspešne uložené! Počkajte pár sekúnd pred generovaním.")
            return True
        return False
    except: return False

# --- 3. LOGIKA GENERÁTORA ---
def generuj_final_streamlit(m, r, fond_limit, parl_active, p_from, p_to, df_volno_edited, use_extra_w, df_db):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    ws = workbook.add_worksheet("Plán")
    ws_miss = workbook.add_worksheet("Neobsadené")
    
    fmt_b = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': False, 'font_size': 9}
    fmt_sep = {**fmt_b, 'bottom': 2}
    z_fmts = {str(i+1): workbook.add_format({**fmt_sep, 'bg_color': c, 'font_color': fc}) 
              for i, (c, fc) in enumerate(zip(['#B2B2B2','#FF0000','#FFFF00','#003399'], ['white','white','black','white']))}
    f_d, f_kz, f_v = workbook.add_format({**fmt_sep, 'bg_color': '#339933', 'font_color': 'white'}), workbook.add_format({**fmt_sep, 'bg_color': '#0066FF', 'font_color': 'white'}), workbook.add_format({**fmt_sep, 'bg_color': '#00FFCC'})
    fmt_num, fmt_low = workbook.add_format({**fmt_sep, 'num_format': '#,##0.0'}), workbook.add_format({**fmt_sep, 'bg_color': '#FF9900', 'num_format': '#,##0.0'})
    f_c1_d = workbook.add_format({**fmt_b, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
    f_c1_n = workbook.add_format({**fmt_sep, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})

    _, days_count = calendar.monthrange(r, m)
    vysledky, hod_fond_sofar = {d: {'D': {}, 'N': {}} for d in range(1, days_count + 1)}, {idx: 0.0 for idx in df_db.index}
    abs_map = {str(row.get('Priezvisko', '')).strip(): {'d': parse_days(row.get('Dovolenka','')), 'kz': parse_days(row.get('KZ','')), 'v': parse_days(row.get('Volno',''))} for _, row in df_volno_edited.iterrows()}

    for i in range(len(df_db)):
        try:
            row, idx = df_db.iloc[i], df_db.index[i]
            ab = abs_map.get(str(row['Priezvisko']).strip(), {'d':set(), 'kz':set(), 'v':set()})
            z_os = int(float(row['Zmena']))
            for d_v in (ab['d'] | ab['kz']):
                if 1 <= d_v <= days_count:
                    if CYKLY[z_os][(date(r, m, d_v) - START_REF).days % 8] in ['D', 'N']: hod_fond_sofar[idx] += 11.5
        except: continue

    for d in range(1, days_count + 1):
        curr_d = date(r, m, d)
        is_workday = curr_d.weekday() < 5 and curr_d not in SVIATKY_2026
        if is_workday:
            nas_z8 = False
            for col_f in ["Priorita_Z8", "Z8"]:
                if nas_z8: break
                for idx in get_prioritized_people(df_db, curr_d, 'D', hod_fond_sofar, fond_limit, True):
                    if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                    priez = df_db.loc[idx, 'Priezvisko']
                    if isinstance(priez, pd.Series): priez = priez.iloc[0]
                    ab = abs_map.get(str(priez).strip(), {'d':set(), 'kz':set(), 'v':set()})
                    if d not in (ab['d']|ab['kz']|ab['v']) and str(df_db.loc[idx].get(col_f,'Nie')).lower()=='áno':
                        vysledky[d]['D'][idx] = "Z8"; hod_fond_sofar[idx] += 7.5; nas_z8 = True; break

        for smena in ['D', 'N']:
            for idx in df_db.index:
                if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                try:
                    z_val = df_db.loc[idx, 'Zmena']
                    if isinstance(z_val, pd.Series): z_val = z_val.iloc[0]
                    z_os = int(float(z_val))
                    if CYKLY[z_os][(curr_d - START_REF).days % 8] == smena and str(df_db.loc[idx].get('C1','Nie')).lower() == 'áno':
                        priez = df_db.loc[idx, 'Priezvisko']
                        if isinstance(priez, pd.Series): priez = priez.iloc[0]
                        ab = abs_map.get(str(priez).strip(), {'d':set(), 'kz':set(), 'v':set()})
                        if d not in (ab['d']|ab['kz']|ab['v']) and moze_nastupit(idx, d, smena, 'C1', vysledky):
                            vysledky[d][smena][idx] = 'C1'; hod_fond_sofar[idx] += 11.5; break
                except: continue

            for p_n in ['ZT', 'NB']:
                if p_n in vysledky[d][smena].values() or (p_n == 'NB' and smena == 'D' and is_workday): continue
                pool, nas = get_prioritized_people(df_db, curr_d, smena, hod_fond_sofar, fond_limit), False
                for idx in pool:
                    if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                    z_val = df_db.loc[idx, 'Zmena']
                    if isinstance(z_val, pd.Series): z_val = z_val.iloc[0]
                    z_os = int(float(z_val))
                    if str(df_db.loc[idx].get(f"Priorita_{p_n}",'Nie')).lower() == 'áno' and CYKLY[z_os][(curr_d - START_REF).days % 8] == smena:
                        priez = df_db.loc[idx, 'Priezvisko']
                        if isinstance(priez, pd.Series): priez = priez.iloc[0]
                        ab = abs_map.get(str(priez).strip(), {'d':set(), 'kz':set(), 'v':set()})
                        if d not in (ab['d']|ab['kz']|ab['v']) and moze_nastupit(idx, d, smena, p_n, vysledky):
                            vysledky[d][smena][idx] = p_n; hod_fond_sofar[idx] += 11.5; nas = True; break
                if not nas:
                    for idx in pool:
                        if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                        z_val = df_db.loc[idx, 'Zmena']
                        if isinstance(z_val, pd.Series): z_val = z_val.iloc[0]
                        z_os = int(float(z_val))
                        if str(df_db.loc[idx].get(p_n,'Nie')).lower() == 'áno' and CYKLY[z_os][(curr_d - START_REF).days % 8] == smena:
                            priez = df_db.loc[idx, 'Priezvisko']
                            if isinstance(priez, pd.Series): priez = priez.iloc[0]
                            ab = abs_map.get(str(priez).strip(), {'d':set(), 'kz':set(), 'v':set()})
                            if d not in (ab['d']|ab['kz']|ab['v']) and moze_nastupit(idx, d, smena, p_n, vysledky):
                                vysledky[d][smena][idx] = p_n; hod_fond_sofar[idx] += 11.5; nas = True; break

            if smena == 'D' and is_workday:
                specs = (['TP', 'S1', 'S2', 'S3'] if parl_active and p_from <= curr_d <= p_to else []) + (['W_EXTRA'] if use_extra_w else []) + ['M']
                for poz in specs:
                    if poz in vysledky[d]['D'].values(): continue
                    for idx in get_prioritized_people(df_db, curr_d, 'D', hod_fond_sofar, fond_limit):
                        if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                        priez = df_db.loc[idx, 'Priezvisko']
                        if isinstance(priez, pd.Series): priez = priez.iloc[0]
                        ab = abs_map.get(str(priez).strip(), {'d':set(), 'kz':set(), 'v':set()})
                        p_col = poz if poz != 'W_EXTRA' else 'W1'
                        if d not in (ab['d']|ab['kz']|ab['v']) and str(df_db.loc[idx].get(p_col,'Nie')).lower() == 'áno' and moze_nastupit(idx, d, 'D', poz, vysledky):
                            vysledky[d]['D'][idx] = poz; hod_fond_sofar[idx] += 11.5; break

            for poz in PRIO_LIST:
                if poz in vysledky[d][smena].values(): continue
                for idx in get_prioritized_people(df_db, curr_d, smena, hod_fond_sofar, fond_limit):
                    if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                    priez = df_db.loc[idx, 'Priezvisko']
                    if isinstance(priez, pd.Series): priez = priez.iloc[0]
                    ab = abs_map.get(str(priez).strip(), {'d':set(), 'kz':set(), 'v':set()})
                    if d not in (ab['d']|ab['kz']|ab['v']) and str(df_db.loc[idx].get(poz,'Nie')).lower() == 'áno' and moze_nastupit(idx, d, smena, poz, vysledky):
                        vysledky[d][smena][idx] = poz; hod_fond_sofar[idx] += 11.5; break

        if is_workday:
            wa = (((curr_d - START_REF).days // 7) % 2 == 0)
            trg = "IR" if (wa and curr_d.weekday() <= 1) or (not wa and curr_d.weekday() >= 2) else "IP"
            for idx in get_prioritized_people(df_db, curr_d, 'D', hod_fond_sofar, fond_limit, True):
                if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                priez = df_db.loc[idx, 'Priezvisko']
                if isinstance(priez, pd.Series): priez = priez.iloc[0]
                ab = abs_map.get(str(priez).strip(), {'d':set(), 'kz':set(), 'v':set()})
                if d not in (ab['d']|ab['kz']|ab['v']):
                    fx = trg if str(df_db.loc[idx].get(trg,'Nie')).lower() == 'áno' else next((p for p in ['X'] if str(df_db.loc[idx].get(p,'Nie')).lower() == 'áno'), None)
                    if fx: vysledky[d]['D'][idx] = fx; hod_fond_sofar[idx] += 7.5

    # ZÁPIS EXCEL
    ws.set_column(0, 0, 25); ws.set_column(days_count+1, days_count+2, 10); ZZ = days_count + 10
    for d in range(1, days_count + 1): ws.set_column(d, d, 3.5)
    col_bg_map = {d: ('#40B4EE' if date(r,m,d) in SVIATKY_2026 else ('#FFCC66' if date(r,m,d).weekday()==5 else ('#CC9900' if date(r,m,d).weekday()==6 else '#FFFFFF'))) for d in range(1, days_count+1)}
    for d in range(1, days_count+1): ws.write(0, d, d, workbook.add_format({**fmt_b, 'bg_color': col_bg_map[d]}))
    ws.write(0, days_count+1, "Sumár", workbook.add_format({'bold':True, 'border':1}))
    ws.write(0, days_count+2, "Rozdiel", workbook.add_format({'bold':True, 'border':1}))

    for i in range(len(df_db)):
        row, idx, zebra, row_ptr = df_db.iloc[i], df_db.index[i], ('#FFFF00' if i % 2 == 1 else '#FFFFFF'), i*2+1
        z_num = str(int(float(row['Zmena'])))
        ws.merge_range(row_ptr, 0, row_ptr+1, 0, f"{row['Priezvisko']} {row['Meno']}", z_fmts.get(z_num, z_fmts['1']))
        ws.write(row_ptr, ZZ, int(float(row['Zmena'])))
        for d in range(1, days_count + 1):
            bg = col_bg_map[d] if col_bg_map[d] != '#FFFFFF' else zebra
            ab = abs_map.get(str(row['Priezvisko']).strip(), {'d':set(), 'kz':set(), 'v':set()})
            cyk_char = CYKLY[int(float(row['Zmena']))][(date(r, m, d) - START_REF).days % 8]
            if d in ab['d']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'D', f_d)
            elif d in ab['kz']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'KZ', f_kz)
            elif d in ab['v']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'V', f_v)
            else:
                ps, ns = short_label(vysledky[d]['D'].get(idx, "")), short_label(vysledky[d]['N'].get(idx, ""))
                ws.write(row_ptr, d, ps, f_c1_d if ps=='C' else workbook.add_format({**fmt_b, 'bg_color': bg, 'bold': bool(ps) and cyk_char != 'D'}))
                ws.write(row_ptr+1, d, ns, f_c1_n if ns=='C' else workbook.add_format({**fmt_sep, 'bg_color': bg, 'bold': bool(ns) and cyk_char != 'N'}))
        
        # FULL FORMULA (Obnovená tvoja pôvodná logika z Colabu)
        r_ex, zz_col = row_ptr+1, xlsxwriter.utility.xl_col_to_name(ZZ)
        cyk_f = f"CHOOSE({zz_col}{r_ex},\"{CYKLY[1]}\",\"{CYKLY[2]}\",\"{CYKLY[3]}\",\"{CYKLY[4]}\")"
        f_parts = [f"IF(OR({xlsxwriter.utility.xl_col_to_name(d)}{r_ex}=\"D\",{xlsxwriter.utility.xl_col_to_name(d)}{r_ex}=\"KZ\"),IF(OR(MID({cyk_f},{(date(r,m,d)-START_REF).days%8+1},1)=\"D\",MID({cyk_f},{(date(r,m,d)-START_REF).days%8+1},1)=\"N\"),11.5,0),0)" for d in range(1, days_count+1)]
        sc, ec = xlsxwriter.utility.xl_col_to_name(1), xlsxwriter.utility.xl_col_to_name(days_count)
        full_formula = f"=(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"*\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"R\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"K\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"X\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"Z8\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"D\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"KZ\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"V\")*11.5)+({'+'.join(f_parts)})"
        
        ws.merge_range(row_ptr, days_count+1, row_ptr+1, days_count+1, full_formula, fmt_num)
        sum_c = xlsxwriter.utility.xl_rowcol_to_cell(row_ptr, days_count+1)
        ws.merge_range(row_ptr, days_count+2, row_ptr+1, days_count+2, f"={fond_limit}-{sum_c}", fmt_num)
        ws.conditional_format(row_ptr, days_count+2, row_ptr+1, days_count+2, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_low})

    ws_miss.write_row(0, 0, ["Deň", "Smena", "Pozícia"], workbook.add_format({'bold':True, 'border':1}))
    m_row = 1
    for d in range(1, days_count+1):
        for smena in ['D', 'N']:
            curr_obs = vysledky[d][smena].values()
            prio_check = PRIO_LIST + (['Z8'] if smena == 'D' and date(r,m,d).weekday()<5 and date(r,m,d) not in SVIATKY_2026 else [])
            for p in prio_check:
                if p not in curr_obs: ws_miss.write_row(m_row, 0, [d, smena, p]); m_row += 1
    workbook.close()
    return output.getvalue(), f"Plan_{m}_{r}.xlsx"

# --- 4. UI ---
st.set_page_config(page_title="Smart Plánovač 2026", layout="wide")
st.title("🚀 Smart Plánovač 2026")

db_file = load_from_github()
if db_file:
    ex = pd.ExcelFile(db_file, engine='openpyxl')
    df_db_raw = ex.parse('Data').dropna(subset=['Priezvisko'])
    df_v_raw = ex.parse('Volno') if 'Volno' in ex.sheet_names else pd.DataFrame(columns=['Priezvisko', 'Meno', 'Dovolenka', 'KZ', 'Volno'])
    for col in ['Dovolenka', 'KZ', 'Volno']: df_v_raw[col] = df_v_raw[col].fillna("").astype(str).replace('nan', '')

    t1, t2 = st.tabs(["📊 Plánovanie", "⚙️ Databáza"])
    with t2:
        df_db_edit = st.data_editor(df_db_raw, use_container_width=True, key="db_ed", num_rows="dynamic")
        if st.button("💾 ULOŽIŤ PERSONÁL NA GITHUB"):
            push_to_github(df_db_edit, df_v_raw)
    with t1:
        c1, c2, c3, c4 = st.columns(4)
        mes = c1.selectbox("Mesiac", range(1, 13), index=datetime.now().month-1)
        fon = c2.number_input("Fond", value=155.0)
        parl, extra_w = c3.checkbox("Parlament", True), c4.checkbox("Extra W", True)
        
        _, last_day = calendar.monthrange(2026, mes)
        cp1, cp2 = st.columns(2)
        p_od = cp1.date_input("Parlament od", date(2026, mes, 1), format="DD.MM.YYYY")
        p_do = cp2.date_input("Parlament do", date(2026, mes, last_day), format="DD.MM.YYYY")
        
        df_v_edit = st.data_editor(df_v_raw, use_container_width=True, key="v_ed", num_rows="dynamic")
        if st.button("💾 ULOŽIŤ ABSENCIE NA GITHUB"):
            push_to_github(df_db_edit, df_v_edit)
        
        if st.button("🚀 GENEROVAŤ PLÁN", type="primary", use_container_width=True):
            xlsx, name = generuj_final_streamlit(mes, 2026, fon, parl, p_od, p_do, df_v_edit, extra_w, df_db_edit)
            st.download_button("📥 STIAHNUŤ EXCEL", data=xlsx, file_name=name, use_container_width=True)
else:
    st.error("Chyba pripojenia k databáze na GitHube. Skontrolujte GITHUB_TOKEN.")
