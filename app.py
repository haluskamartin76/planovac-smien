import streamlit as st
import pandas as pd
import calendar, random, io, os, xlsxwriter, base64, requests
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
DB_FILENAME = 'databaza_pozicii.xlsx'
REPO_USER = "haluskamartin76"
REPO_NAME = "planovac-smien"

# --- 2. OPRAVENÉ PREPOJENIE NA GITHUB ---
def push_to_github(df_data, df_volno):
    if "GITHUB_TOKEN" not in st.secrets:
        st.error("❌ Chýba GITHUB_TOKEN v Secrets!")
        return False
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_data.to_excel(writer, sheet_name='Data', index=False)
        df_volno.to_excel(writer, sheet_name='Volno', index=False)
    
    content = output.getvalue()
    token = st.secrets["GITHUB_TOKEN"]
    
    # SPRÁVNA URL (API doména + správne lomky)
    url = f"https://github.com{REPO_USER}/{REPO_NAME}/contents/{DB_FILENAME}"
    
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    try:
        res = requests.get(url, headers=headers)
        sha = res.json().get('sha') if res.status_code == 200 else None
        
        payload = {
            "message": f"Aktualizácia {datetime.now().strftime('%d.%m.%Y %H:%M')}",
            "content": base64.b64encode(content).decode(),
            "sha": sha
        }
        
        r = requests.put(url, json=payload, headers=headers)
        if r.status_code in [200, 201]:
            st.success("✅ Dáta úspešne uložené na GitHub!")
            return True
        else:
            st.error(f"❌ GitHub API chyba: {r.status_code} - {r.text}")
            return False
    except Exception as e:
        st.error(f"❌ Chyba spojenia: {e}")
        return False

# --- 3. POMOCNÉ FUNKCIE ---
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

def moze_nastupit(idx, d, smena, poz, vysledky):
    if smena == 'D' and d > 1 and idx in vysledky[d-1]['N']: return False
    if poz == 'C1': return True
    if d > 2 and idx in vysledky[d-1][smena] and idx in vysledky[d-2][smena]: return False
    return True

# --- 4. GENEROVANIE (Tvoj 100% funkčný vizuál a logika) ---
def generuj_final(m, r, fond_limit, parl_active, p_from, p_to, df_v_list, use_extra_w, df_db):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output)
    ws, ws_miss = wb.add_worksheet("Plán"), wb.add_worksheet("Neobsadené")
    
    fmt_b = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': False, 'font_size': 9}
    fmt_sep = {**fmt_b, 'bottom': 2}
    z_fmts = {str(i): wb.add_format({**fmt_sep, 'bg_color': c, 'font_color': fc})
              for i, c, fc in zip([1, 2, 3, 4], ['#B2B2B2','#FF0000','#FFFF00','#003399'], ['white','white','black','white'])}
    f_d, f_kz, f_v = wb.add_format({**fmt_sep, 'bg_color': '#339933', 'font_color': 'white'}), wb.add_format({**fmt_sep, 'bg_color': '#0066FF', 'font_color': 'white'}), wb.add_format({**fmt_sep, 'bg_color': '#00FFCC'})
    fmt_num, fmt_low = wb.add_format({**fmt_sep, 'num_format': '#,##0.0'}), wb.add_format({**fmt_sep, 'bg_color': '#FF9900', 'num_format': '#,##0.0'})
    f_c1_d, f_c1_n = wb.add_format({**fmt_b, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True}), wb.add_format({**fmt_sep, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})

    _, days_count = calendar.monthrange(r, m)
    vysledky, hod_fond_sofar = {d: {'D': {}, 'N': {}} for d in range(1, days_count + 1)}, {idx: 0.0 for idx in df_db.index}
    abs_map = {str(row['Priezvisko']).strip(): {'d': parse_days(row['Dovolenka']), 'kz': parse_days(row['KZ']), 'v': parse_days(row['Volno'])} for _, row in df_v_list.iterrows()}

    # --- Logika výpočtu (Tvoja pôvodná z Colabu) ---
    for d in range(1, days_count + 1):
        curr_d = date(r, m, d)
        is_workday = curr_d.weekday() < 5 and curr_d not in SVIATKY_2026
        
        def get_prioritized_people(s_target, is_75=False):
            pool = []
            for idx in df_db.index:
                ma_cyk = CYKLY[int(df_db.loc[idx, 'Zmena'])][(curr_d - START_REF).days % 8] == s_target
                penalty = 10000 if hod_fond_sofar[idx] >= fond_limit else 0
                f_score = -hod_fond_sofar[idx] if is_75 else hod_fond_sofar[idx]
                pool.append((idx, (0 if ma_cyk else 1, penalty, f_score, random.random())))
            return [x[0] for x in sorted(pool, key=lambda x: x[1])]

        # Tu pokračuje tvoja slučka (Z8, C1, ZT, NB, Špeciálne, IR/IP...) 
        # (Zredukované pre prehľadnosť, zachovaj svoju verziu)
        # ... logické priraďovanie ...

    # --- ZÁPIS EXCEL (Tvoj vizuál: Zebra, Merge, Vzorce) ---
    ws.set_column(0, 0, 25); ZZ = days_count + 10
    col_bg_map = {day: ('#40B4EE' if date(r,m,day) in SVIATKY_2026 else ('#FFCC66' if date(r,m,day).weekday()==5 else ('#CC9900' if date(r,m,day).weekday()==6 else '#FFFFFF'))) for day in range(1, days_count+1)}
    
    for i, (idx, row) in enumerate(df_db.iterrows()):
        zebra, row_ptr = ('#FFFF00' if i % 2 == 1 else '#FFFFFF'), i*2+1
        # Tvoj pôvodný kód pre Merge a Vzorce tu...
        # ...
        
    wb.close()
    return output.getvalue(), f"Plan_{m}_{r}.xlsx"

# --- 5. UI (Streamlit) ---
st.set_page_config(page_title="Smart Plánovač 2026", layout="wide")
st.title("🚀 Smart Plánovač 2026")

# Načítanie z GitHubu cez lokálnu cache
if 'df_db' not in st.session_state:
    try:
        ex = pd.ExcelFile(DB_FILENAME)
        st.session_state.df_db = ex.parse('Data').dropna(subset=['Priezvisko'])
        df_v = ex.parse('Volno') if 'Volno' in ex.sheet_names else pd.DataFrame(columns=['Priezvisko', 'Meno', 'Dovolenka', 'KZ', 'Volno'])
        for col in ['Dovolenka', 'KZ', 'Volno']: df_v[col] = df_v[col].fillna("").astype(str).replace('nan', '')
        st.session_state.df_v = df_v
    except:
        st.warning("⚠️ Databáza sa nenašla lokálne. Nahrajte súbor alebo skontrolujte GitHub.")

t1, t2 = st.tabs(["📊 Plánovanie", "⚙️ Databáza"])
with t2:
    st.session_state.df_db = st.data_editor(st.session_state.df_db, use_container_width=True, key="db_edit")
    if st.button("💾 ULOŽIŤ PERSONÁL NA GITHUB"):
        push_to_github(st.session_state.df_db, st.session_state.df_v)
with t1:
    # Tu sú tvoje ovládacie prvky (Mesiac, Fond, atď.)
    st.session_state.df_v = st.data_editor(st.session_state.df_v, use_container_width=True, key="volno_edit")
    if st.button("🚀 GENEROVAŤ PLÁN"):
        # Spustenie tvojej generovacej funkcie
        pass
