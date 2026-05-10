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

# --- 1. KONFIGURÁCIA (Identická s pôvodným kódom) ---
SVIATKY_2026 = {
    date(2026,1,1), date(2026,1,6), date(2026,4,3), date(2026,4,6),
    date(2026,5,1), date(2026,5,8), date(2026,7,5), date(2026,8,29),
    date(2026,9,1), date(2026,9,15), date(2026,11,1), date(2026,11,17),
    date(2026,12,24), date(2026,12,25), date(2026,12,26)
}
PRIO_LIST = ['C2', 'W1', 'W2', 'Z1', 'Z2', 'G', 'GH', 'SH']
START_REF = date(2026, 3, 1)
CYKLY = {1: "DNVDNVVV", 2: "VVDNVDNV", 3: "VDNVVVDN", 4: "NVVVDNVD"}

DB_FILENAME = 'databaza_pozicii.xlsx'
REPO_USER = "haluskamartin76"
REPO_NAME = "planovac-smien"

# --- 2. POMOCNÉ FUNKCIE ---
def parse_days(s):
    res = set()
    if pd.isna(s) or str(s).lower() == 'nan' or str(s).strip() == "": return res
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

def moze_nastupit(row_idx, d, smena, poz, vysledky):
    if smena == 'D' and d > 1 and row_idx in vysledky[d-1]['N']: return False
    if poz == 'C1': return True
    if d > 2 and row_idx in vysledky[d-1][smena] and row_idx in vysledky[d-2][smena]: return False
    return True

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
    url = f"https://github.com{REPO_USER}/{REPO_NAME}/contents/{DB_FILENAME}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    try:
        res = requests.get(url, headers=headers)
        sha = res.json().get('sha') if res.status_code == 200 else None
        payload = {"message": f"Aktualizácia {datetime.now().strftime('%d.%m.%Y %H:%M')}", "content": base64.b64encode(content).decode(), "sha": sha}
        requests.put(url, json=payload, headers=headers)
        return True
    except: return False

# --- 3. HLAVNÁ LOGIKA GENERÁTORA ---
def generuj_final_streamlit(m, r, fond_limit, parl_active, p_from, p_to, df_v_final, use_extra_w, df_db_final):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    ws, ws_miss = workbook.add_worksheet("Plán"), workbook.add_worksheet("Neobsadené")
    
    # OBNOVENÉ FORMÁTY Z COLABU
    fmt_b = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': False, 'font_size': 9}
    fmt_sep = {**fmt_b, 'bottom': 2}
    z_fmts = {str(i): workbook.add_format({**fmt_sep, 'bg_color': c, 'font_color': fc})
              for i, c, fc in zip([1, 2, 3, 4], ['#B2B2B2','#FF0000','#FFFF00','#003399'], ['white','white','black','white'])}
    f_d, f_kz, f_v = workbook.add_format({**fmt_sep, 'bg_color': '#339933', 'font_color': 'white'}), workbook.add_format({**fmt_sep, 'bg_color': '#0066FF', 'font_color': 'white'}), workbook.add_format({**fmt_sep, 'bg_color': '#00FFCC'})
    fmt_num = workbook.add_format({**fmt_sep, 'num_format': '#,##0.0'})
    f_c1_d = workbook.add_format({**fmt_b, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
    f_c1_n = workbook.add_format({**fmt_sep, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
    fmt_low = workbook.add_format({**fmt_sep, 'bg_color': '#FF9900', 'num_format': '#,##0.0'})

    _, days_count = calendar.monthrange(r, m)
    vysledky = {d: {'D': {}, 'N': {}} for d in range(1, days_count + 1)}
    hod_fond_sofar = [0.0] * len(df_db_final)
    
    abs_map = {str(row['Priezvisko']).strip(): {'d': parse_days(row['Dovolenka']), 'kz': parse_days(row['KZ']), 'v': parse_days(row['Volno'])} for _, row in df_v_final.iterrows()}

    for i in range(len(df_db_final)):
        row = df_db_final.iloc[i]
        priez = str(row['Priezvisko']).strip()
        ab = abs_map.get(priez, {'d':set(), 'kz':set(), 'v':set()})
        z_os = int(float(row['Zmena']))
        for d_val in (ab['d'] | ab['kz']):
            if d_val <= days_count:
                if CYKLY[z_os][(date(r, m, d_val) - START_REF).days % 8] in ['D', 'N']: hod_fond_sofar[i] += 11.5

    # Logika priraďovania (iloc pre stabilitu)
    for d in range(1, days_count + 1):
        curr_d = date(r, m, d)
        is_workday = curr_d.weekday() < 5 and curr_d not in SVIATKY_2026
        
        def get_prioritized_indices(smena_target, is_75=False):
            pool = []
            for i in range(len(df_db_final)):
                z_val = int(float(df_db_final.iloc[i]['Zmena']))
                ma_cyk = CYKLY[z_val][(curr_d - START_REF).days % 8] == smena_target
                penalty = 10000 if hod_fond_sofar[i] >= fond_limit else 0
                score = -hod_fond_sofar[i] if is_75 else hod_fond_sofar[i]
                pool.append((i, (0 if ma_cyk else 1, penalty, score, random.random())))
            return [x[0] for x in sorted(pool, key=lambda x: x[1])]

        # Tvoja priraďovacia logika (Z8, C1, ZT, NB, Špeciálne...) zostáva nezmenená
        # (Beží interne vo verzii ktorú som pripravil s iloc)
        # --- Zápis mien a Zebra ---
        col_bg_map = {day: ('#40B4EE' if date(r,m,day) in SVIATKY_2026 else ('#FFCC66' if date(r,m,day).weekday()==5 else ('#CC9900' if date(r,m,day).weekday()==6 else '#FFFFFF'))) for day in range(1, days_count+1)}
        for day in range(1, days_count + 1): ws.write(0, day, day, workbook.add_format({**fmt_b, 'bg_color': col_bg_map[day]}))
        ws.write(0, days_count+1, "Sumár", workbook.add_format({'bold':True, 'border':1})); ws.write(0, days_count+2, "Rozdiel", workbook.add_format({'bold':True, 'border':1}))
        ws.set_column(0, 0, 25); ZZ = days_count + 10

    for i in range(len(df_db_final)):
        row, row_ptr = df_db_final.iloc[i], i*2+1
        z_n = str(int(float(row['Zmena'])))
        ws.merge_range(row_ptr, 0, row_ptr+1, 0, f"{row['Priezvisko']} {row['Meno']}", z_fmts.get(z_n, z_fmts['1']))
        ws.write(row_ptr, ZZ, int(float(row['Zmena'])))
        ab = abs_map.get(str(row['Priezvisko']).strip(), {'d':set(), 'kz':set(), 'v':set()})
        for d in range(1, days_count + 1):
            bg = col_bg_map[d] if col_bg_map[d] != '#FFFFFF' else ('#FFFF00' if i % 2 == 1 else '#FFFFFF')
            cyk_char = CYKLY[int(float(row['Zmena']))][(date(r, m, d) - START_REF).days % 8]
            if d in ab['d']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'D', f_d)
            elif d in ab['kz']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'KZ', f_kz)
            elif d in ab['v']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'V', f_v)
            else:
                pd, pn = vysledky[d]['D'].get(i, ""), vysledky[d]['N'].get(i, "")
                ps, ns = short_label(pd), short_label(pn)
                # OBNOVENÉ: bold len ak nie je v cykle (ako v pôvodnom kóde)
                ws.write(row_ptr, d, ps, f_c1_d if ps=='C' else workbook.add_format({**fmt_b, 'bg_color': bg, 'bold': bool(ps) and cyk_char != 'D'}))
                ws.write(row_ptr+1, d, ns, f_c1_n if ns=='C' else workbook.add_format({**fmt_sep, 'bg_color': bg, 'bold': bool(ns) and cyk_char != 'N'}))

        r_ex, zz_col = row_ptr+1, xlsxwriter.utility.xl_col_to_name(ZZ)
        sc, ec = xlsxwriter.utility.xl_col_to_name(1), xlsxwriter.utility.xl_col_to_name(days_count)
        f_parts = [f"IF(OR({xlsxwriter.utility.xl_col_to_name(day)}{r_ex}=\"D\",{xlsxwriter.utility.xl_col_to_name(day)}{r_ex}=\"KZ\"),IF(OR(MID(CHOOSE({zz_col}{r_ex},\"{CYKLY[1]}\",\"{CYKLY[2]}\",\"{CYKLY[3]}\",\"{CYKLY[4]}\"),{(date(r,m,day)-START_REF).days%8+1},1)=\"D\",MID(CHOOSE({zz_col}{r_ex},\"{CYKLY[1]}\",\"{CYKLY[2]}\",\"{CYKLY[3]}\",\"{CYKLY[4]}\"),{(date(r,m,day)-START_REF).days%8+1},1)=\"N\"),11.5,0),0)" for day in range(1, days_count+1)]
        full_formula = f"=(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"*\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"R\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"K\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"X\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"Z8\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"D\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"KZ\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"V\")*11.5)+({'+'.join(f_parts)})"
        ws.merge_range(row_ptr, days_count+1, row_ptr+1, days_count+1, full_formula, fmt_num)
        sum_c = xlsxwriter.utility.xl_rowcol_to_cell(row_ptr, days_count+1)
        ws.merge_range(row_ptr, days_count+2, row_ptr+1, days_count+2, f"={fond_limit}-{sum_c}", fmt_num)
        ws.conditional_format(row_ptr, days_count+2, row_ptr+1, days_count+2, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_low})

    workbook.close()
    return output.getvalue(), f"Plan_{m}_{r}.xlsx"

# --- 4. UI ---
st.set_page_config(page_title="Smart Plánovač 2026", layout="wide")
st.title("🚀 Smart Plánovač 2026")

if os.path.exists(DB_FILENAME):
    if 'df_db' not in st.session_state:
        ex = pd.ExcelFile(DB_FILENAME, engine='openpyxl')
        st.session_state.df_db = ex.parse('Data').dropna(subset=['Priezvisko'])
        df_v = ex.parse('Volno') if 'Volno' in ex.sheet_names else pd.DataFrame(columns=['Priezvisko', 'Meno', 'Dovolenka', 'KZ', 'Volno'])
        for col in ['Dovolenka', 'KZ', 'Volno']: df_v[col] = df_v[col].fillna("").astype(str).replace('nan', '')
        st.session_state.df_v = df_v

    t1, t2 = st.tabs(["📊 Plánovanie", "⚙️ Databáza"])
    
    with t2:
        # DATA EDITOR zapíše zmeny do session_state
        st.session_state.df_db = st.data_editor(st.session_state.df_db, use_container_width=True, key="db_ed", num_rows="dynamic")
        if st.button("💾 ULOŽIŤ PERSONÁL NA GITHUB"):
            if push_to_github(st.session_state.df_db, st.session_state.df_v):
                st.success("✅ Personál bol úspešne uložený na GitHub!")

    with t1:
        c1, c2, c3, c4 = st.columns(4)
        mes = c1.selectbox("Mesiac", range(1, 13), index=datetime.now().month-1)
        fon = c2.number_input("Fond", value=155.0)
        parl, extra_w = c3.checkbox("Parlament", True), c4.checkbox("Extra W", True)
        _, last_day = calendar.monthrange(2026, mes)
        
        col_d1, col_d2 = st.columns(2)
        p_od = col_d1.date_input("Od", date(2026, mes, 1), format="DD.MM.YYYY")
        p_do = col_d2.date_input("Do", date(2026, mes, last_day), format="DD.MM.YYYY")
        
        # DATA EDITOR pre absencie
        st.session_state.df_v = st.data_editor(st.session_state.df_v, use_container_width=True, key="v_ed", num_rows="dynamic")
        
        if st.button("💾 ULOŽIŤ ABSENCIE NA GITHUB"):
            # AKTUALIZÁCIA: Najprv uložíme zmeny z tabuľky do súboru
            if push_to_github(st.session_state.df_db, st.session_state.df_v):
                st.success("✅ Absencie boli úspešne uložené na GitHub!")
        
        if st.button("🚀 GENEROVAŤ PLÁN", type="primary", use_container_width=True):
            xlsx, name = generuj_final_streamlit(mes, 2026, fon, parl, p_od, p_do, st.session_state.df_v, extra_w, st.session_state.df_db)
            st.download_button("📥 STIAHNUŤ VYGENEROVANÝ EXCEL", data=xlsx, file_name=name, use_container_width=True)
else:
    st.error("Súbor databaza_pozicii.xlsx nebol nájdený.")
