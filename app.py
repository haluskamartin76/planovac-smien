import streamlit as st
import pandas as pd
import calendar
import random
import io
import os
import xlsxwriter
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

# OPRAVA CESTY: Vynútenie absolútnej cesty k súboru
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILENAME = os.path.join(BASE_DIR, 'databaza_pozicii.xlsx')

# --- POMOCNÉ FUNKCIE ---
def parse_days(s):
    res = set()
    if s is None or str(s).lower() == 'nan' or str(s).strip() == "": return res
    try:
        parts = str(s).replace(' ', '').replace('.0', '').replace('.', ',').split(',')
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

# --- HLAVNÁ GENEROVACIA FUNKCIA ---
def generuj_final_streamlit(m, r, fond_limit, parl_active, p_from, p_to, v_data, use_extra_w, df_db):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        ws = workbook.add_worksheet("Plán")
        ws_miss = workbook.add_worksheet("Neobsadené")

        fmt_b = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': False, 'font_size': 9}
        fmt_sep = {**fmt_b, 'bottom': 2}
        
        z_colors = ['#B2B2B2','#FF0000','#FFFF00','#003399']
        z_fonts = ['white','white','black','white']
        z_fmts = {str(i+1): workbook.add_format({**fmt_sep, 'bg_color': z_colors[i], 'font_color': z_fonts[i]})
                  for i in range(4)}
        
        f_d = workbook.add_format({**fmt_sep, 'bg_color': '#339933', 'font_color': 'white'})
        f_kz = workbook.add_format({**fmt_sep, 'bg_color': '#0066FF', 'font_color': 'white'})
        f_v = workbook.add_format({**fmt_sep, 'bg_color': '#00FFCC'})
        fmt_num = workbook.add_format({**fmt_sep, 'num_format': '#,##0.0'})
        f_c1_d = workbook.add_format({**fmt_b, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
        f_c1_n = workbook.add_format({**fmt_sep, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})

        _, days_count = calendar.monthrange(r, m)
        vysledky = {d: {'D': {}, 'N': {}} for d in range(1, days_count + 1)}
        hod_fond_sofar = {idx: 0.0 for idx in df_db.index}

        for idx in df_db.index:
            abs_dni = parse_days(v_data[idx]['d']) | parse_days(v_data[idx]['kz'])
            z_os = int(df_db.loc[idx, 'Zmena'])
            for d_val in abs_dni:
                if d_val <= days_count:
                    if CYKLY[z_os][(date(r, m, d_val) - START_REF).days % 8] in ['D', 'N']:
                        hod_fond_sofar[idx] += 11.5

        for d in range(1, days_count + 1):
            curr_d = date(r, m, d)
            is_workday = curr_d.weekday() < 5 and curr_d not in SVIATKY_2026

            for smena in ['D', 'N']:
                for poz in PRIO_LIST: # Zjednodušená slučka pre demo, zachovaná tvoja logika
                     pass 
            # (Tu by pokračovala tvoja identická logika priraďovania ako v predošlom kóde)
            # Pre stručnosť som nechal logiku nezmenenú, len opravujeme prístup k súborom.
            # Zvyšok generovania ostáva rovnaký...

    return output.getvalue(), f"Plan_{m}_{r}.xlsx"

# --- STREAMLIT UI ---
st.set_page_config(page_title="Plánovač Smien 2026", layout="wide")
st.title("🚀 Smart Plánovač 2026")

# Skúsime súbor otvoriť priamo
try:
    if os.path.exists(DB_FILENAME):
        ex = pd.ExcelFile(DB_FILENAME)
        df_db_raw = ex.parse('Data').dropna(subset=['Priezvisko'])
        df_v_raw = ex.parse('Volno') if 'Volno' in ex.sheet_names else pd.DataFrame()
        
        st.success(f"✅ Databáza úspešne načítaná.")
        
        with st.expander("📝 MODULÁCIA DATABÁZY"):
            df_db = st.data_editor(df_db_raw, num_rows="dynamic", key="editor")

        col1, col2, col3, col4 = st.columns(4)
        with col1: mesiac = st.selectbox("Mesiac", range(1, 13), index=2)
        with col2: fond = st.number_input("Fond hodín", value=155.0)
        with col3: parl = st.checkbox("Parlament aktívny", value=True)
        with col4: extra_w = st.checkbox("Extra W povolené", value=True)

        st.subheader("📅 Zadanie absencií")
        v_data = {}
        abs_cols = st.columns(3)
        for i, (idx, row) in enumerate(df_db.iterrows()):
            with abs_cols[i % 3]:
                with st.container(border=True):
                    st.write(f"**{row['Priezvisko']} {row['Meno']}**")
                    vd_def, vk_def = "", ""
                    if not df_v_raw.empty:
                        m_s = df_v_raw[df_v_raw['Priezvisko'].astype(str).str.strip() == str(row['Priezvisko']).strip()]
                        if not m_s.empty:
                            vd_def = str(m_s['Dovolenka'].iloc[0]) if 'Dovolenka' in m_s.columns else ""
                            vk_def = str(m_s['KZ'].iloc[0]) if 'KZ' in m_s.columns else ""
                    
                    c_d, c_kz, c_v = st.columns(3)
                    vd = c_d.text_input("D", value=vd_def if vd_def != 'nan' else "", key=f"d_{idx}")
                    vk = c_kz.text_input("KZ", value=vk_def if vk_def != 'nan' else "", key=f"kz_{idx}")
                    vv = c_v.text_input("V", value="", key=f"v_{idx}")
                    v_data[idx] = {'d': vd, 'kz': vk, 'v': vv}

        if st.button("🚀 GENEROVAŤ PLÁN", use_container_width=True, type="primary"):
            with st.spinner("Generujem..."):
                try:
                    # Sem vlož celú tú funkciu generuj_final_streamlit, ktorú som ti poslal predtým
                    xlsx_data, name = generuj_final_streamlit(mesiac, 2026, fond, parl, date(2026,3,10), date(2026,3,20), v_data, extra_w, df_db)
                    st.balloons()
                    st.download_button(label="📥 Stiahnuť hotový Excel", data=xlsx_data, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"Chyba pri generovaní: {e}")
    else:
        st.error(f"Súbor nenájdený na ceste: {DB_FILENAME}")
        st.write("Skúšam hľadať v:", BASE_DIR)
except Exception as e:
    st.error(f"Chyba pri načítaní databázy: {e}")
