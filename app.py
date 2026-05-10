import streamlit as st
import pandas as pd
import calendar
import random
import io
import os
import xlsxwriter
import base64
import requests
from datetime import date, datetime

# --- KONFIGURÁCIA ---
SVIATKY_2026 = {
    date(2026,1,1), date(2026,1,6), date(2026,4,3), date(2026,4,6),
    date(2026,5,1), date(2026,5,8), date(2026,7,5), date(2026,8,29),
    date(2026,9,1), date(2026,9,15), date(2026,11,1), date(2026,11,17),
    date(2026,12,24), date(2026,12,25), date(2026,12,26)
}
PRIO_LIST = ['C2', 'W1', 'W2', 'Z1', 'Z2', 'G', 'GH', 'SH']
START_REF = date(2026, 3, 1)
CYKLY = {1: "DNVDNVVV", 2: "VVDNVDNV", 3: "VDNVVVDN", 4: "NVVVDNVD"}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILENAME = os.path.join(BASE_DIR, 'databaza_pozicii.xlsx')

# --- FUNKCIA NA ZÁPIS NA GITHUB (Voliteľná) ---
def save_to_github(file_content, filename):
    if "GITHUB_TOKEN" not in st.secrets:
        return False
    
    token = st.secrets["GITHUB_TOKEN"]
    repo = "TVOJE_MENO/TVOJ_REPOZITAR" # SEM DOPLŇ SVOJU CESTU
    url = f"https://github.com{repo}/contents/{filename}"
    
    # Získanie SHA súboru (povinné pre update)
    resp = requests.get(url, headers={"Authorization": f"token {token}"})
    sha = resp.json().get('sha') if resp.status_code == 200 else None
    
    content_b64 = base64.b64encode(file_content).decode()
    data = {
        "message": f"Aktualizácia dát {datetime.now()}",
        "content": content_b64,
        "sha": sha
    }
    r = requests.put(url, json=data, headers={"Authorization": f"token {token}"})
    return r.status_code in [200, 201]

# --- POMOCNÉ FUNKCIE ---
def parse_days(s):
    res = set()
    if s is None or str(s).lower() == 'nan' or str(s).strip() == "": return res
    try:
        s = str(s).replace('[', '').replace(']', '').replace("'", "").replace('"', '')
        parts = s.replace(' ', '').replace('.0', '').replace('.', ',').split(',')
        for p in parts:
            if not p or not any(char.isdigit() for char in p): continue
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
    for idx in df_db.index:
        z_val = df_db.loc[idx, 'Zmena']
        ma_cyk = CYKLY[int(z_val)][(curr_d - START_REF).days % 8] == smena_target
        penalty = 10000 if hod_fond_sofar[idx] >= fond_limit else 0
        fond_score = -hod_fond_sofar[idx] if is_75_poz else hod_fond_sofar[idx]
        pool.append((idx, (0 if ma_cyk else 1, penalty, fond_score, random.random())))
    return [x for x in sorted(pool, key=lambda x: x)]

# --- GENEROVANIE SMIEN ---
def generuj_final_streamlit(m, r, fond_limit, parl_active, p_from, p_to, df_volno_edited, use_extra_w, df_db):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        ws = workbook.add_worksheet("Plán")
        ws_miss = workbook.add_worksheet("Neobsadené")

        # Formáty
        fmt_b = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': False, 'font_size': 9}
        fmt_sep = {**fmt_b, 'bottom': 2}
        z_fmts = {str(i+1): workbook.add_format({**fmt_sep, 'bg_color': c, 'font_color': fc})
                  for i, c, fc in zip(range(4), ['#B2B2B2','#FF0000','#FFFF00','#003399'], ['white','white','black','white'])}
        f_d, f_kz, f_v = workbook.add_format({**fmt_sep, 'bg_color': '#339933', 'font_color': 'white'}), workbook.add_format({**fmt_sep, 'bg_color': '#0066FF', 'font_color': 'white'}), workbook.add_format({**fmt_sep, 'bg_color': '#00FFCC'})
        fmt_num = workbook.add_format({**fmt_sep, 'num_format': '#,##0.0'})
        f_c1_d = workbook.add_format({**fmt_b, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
        f_c1_n = workbook.add_format({**fmt_sep, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
        fmt_low = workbook.add_format({**fmt_sep, 'bg_color': '#FF9900', 'num_format': '#,##0.0'})

        _, days_count = calendar.monthrange(r, m)
        vysledky = {d: {'D': {}, 'N': {}} for d in range(1, days_count + 1)}
        hod_fond_sofar = {idx: 0.0 for idx in df_db.index}

        absencie_map = {str(row['Priezvisko']).strip(): {
            'd': parse_days(row['Dovolenka']),
            'kz': parse_days(row['KZ']),
            'v': parse_days(row['Volno'])
        } for _, row in df_volno_edited.iterrows()}

        for idx in df_db.index:
            priezvisko = str(df_db.loc[idx, 'Priezvisko']).strip()
            ab = absencie_map.get(priezvisko, {'d':set(), 'kz':set(), 'v':set()})
            z_os = int(df_db.loc[idx, 'Zmena'])
            for d_val in (ab['d'] | ab['kz']):
                if d_val <= days_count:
                    if CYKLY[z_os][(date(r, m, d_val) - START_REF).days % 8] in ['D', 'N']:
                        hod_fond_sofar[idx] += 11.5

        # LOGIKA (Tu je tvoj kompletný blok priraďovania z predošlého kódu...)
        for d in range(1, days_count + 1):
            curr_d = date(r, m, d)
            is_workday = curr_d.weekday() < 5 and curr_d not in SVIATKY_2026
            
            # Priradenie Z8, C1, ZT, NB, PRIO_LIST, atď. (Tvoja 100% logika)
            # ... (vložiť tvoj blok s priraďovaním) ...

        # Zápis (Identický s tvojím zebra formátom a vzorcami)
        # ...

    return output.getvalue(), f"Plan_{m}_{r}.xlsx"

# --- STREAMLIT UI ---
st.set_page_config(page_title="Plánovač Smien 2026", layout="wide")
st.title("🚀 Smart Plánovač 2026")

if os.path.exists(DB_FILENAME):
    ex = pd.ExcelFile(DB_FILENAME, engine='openpyxl')
    df_db_raw = ex.parse('Data').dropna(subset=['Priezvisko'])
    df_v_raw = ex.parse('Volno') if 'Volno' in ex.sheet_names else pd.DataFrame()
    
    # Vyčistenie stĺpcov pre editor
    for col in ['Dovolenka', 'KZ', 'Volno']:
        if col in df_v_raw.columns:
            df_v_raw[col] = df_v_raw[col].apply(lambda x: "" if pd.isna(x) or str(x).lower() == 'nan' else str(x))

    tab1, tab2 = st.tabs(["📊 Generovanie Plánu", "⚙️ Správa Databázy"])

    with tab2:
        st.subheader("Editácia personálnej databázy")
        df_db_edited = st.data_editor(df_db_raw, use_container_width=True, key="db_editor")
    
    with tab1:
        col1, col2, col3, col4 = st.columns(4)
        mesiac = col1.selectbox("Mesiac", range(1, 13), index=2)
        fond = col2.number_input("Fond hodín", value=155.0)
        parl = col3.checkbox("Parlament", value=True)
        extra_w = col4.checkbox("Extra W", value=True)

        st.subheader("📅 Tabuľka absencií")
        df_v_edited = st.data_editor(df_v_raw, use_container_width=True, key="v_editor")

        if st.button("🚀 GENEROVAŤ A STIAHNUŤ PLÁN", use_container_width=True, type="primary"):
            xlsx_data, name = generuj_final_streamlit(mesiac, 2026, fond, parl, date(2026,3,10), date(2026,3,20), df_v_edited, extra_w, df_db_edited)
            st.success("✅ Plán vygenerovaný!")
            st.download_button("📥 STIAHNUŤ EXCEL", data=xlsx_data, file_name=name, use_container_width=True)
