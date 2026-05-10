import streamlit as st
import pandas as pd
import calendar, random, io, os, xlsxwriter, base64, requests
from datetime import date, datetime

# --- 1. KONFIGURÁCIA (Identická s tvojím originálom) ---
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

def moze_nastupit(idx, d, smena, poz, vysledky):
    if smena == 'D' and d > 1 and idx in vysledky[d-1]['N']: return False
    if poz == 'C1': return True
    if d > 2 and idx in vysledky[d-1][smena] and idx in vysledky[d-2][smena]: return False
    return True

def push_to_github(df_data, df_volno):
    if "GITHUB_TOKEN" not in st.secrets:
        st.error("Chýba GITHUB_TOKEN v Secrets!")
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
        payload = {"message": f"Update {datetime.now()}", "content": base64.b64encode(content).decode(), "sha": sha}
        r = requests.put(url, json=payload, headers=headers)
        return r.status_code in [200, 201]
    except: return False

# --- 3. GENEROVANIE ---
def generuj_final(m, r, fond_limit, parl_active, p_from, p_to, df_v_list, use_extra_w, df_db):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output)
    ws, ws_miss = wb.add_worksheet("Plán"), wb.add_worksheet("Neobsadené")
    
    fmt_b = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': False, 'font_size': 9}
    fmt_sep = {**fmt_b, 'bottom': 2}
    z_fmts = {str(i): wb.add_format({**fmt_sep, 'bg_color': c, 'font_color': fc})
              for i, c, fc in zip([1, 2, 3, 4], ['#B2B2B2','#FF0000','#FFFF00','#003399'], ['white','white','black','white'])}
    f_d, f_kz, f_v = wb.add_format({**fmt_sep, 'bg_color': '#339933', 'font_color': 'white'}), wb.add_format({**fmt_sep, 'bg_color': '#0066FF', 'font_color': 'white'}), wb.add_format({**fmt_sep, 'bg_color': '#00FFCC'})
    fmt_num = wb.add_format({**fmt_sep, 'num_format': '#,##0.0'})
    f_c1_d = wb.add_format({**fmt_b, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
    f_c1_n = wb.add_format({**fmt_sep, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
    fmt_low = wb.add_format({**fmt_sep, 'bg_color': '#FF9900', 'num_format': '#,##0.0'})

    _, days_count = calendar.monthrange(r, m)
    vysledky, hod_fond_sofar = {d: {'D': {}, 'N': {}} for d in range(1, days_count + 1)}, {idx: 0.0 for idx in df_db.index}
    abs_map = {str(row['Priezvisko']).strip(): {'d': parse_days(row['Dovolenka']), 'kz': parse_days(row['KZ']), 'v': parse_days(row['Volno'])} for _, row in df_v_list.iterrows()}

    for idx in df_db.index:
        priez = str(df_db.loc[idx, 'Priezvisko']).strip()
        ab = abs_map.get(priez, {'d':set(), 'kz':set()})
        z_os = int(df_db.loc[idx, 'Zmena'])
        for d_val in (ab['d'] | ab['kz']):
            if d_val <= days_count:
                if CYKLY[z_os][(date(r, m, d_val) - START_REF).days % 8] in ['D', 'N']: hod_fond_sofar[idx] += 11.5

    for d in range(1, days_count + 1):
        curr_d = date(r, m, d)
        is_workday = curr_d.weekday() < 5 and curr_d not in SVIATKY_2026
        def get_prioritized_people(smena_target, is_75_poz=False):
            pool = []
            for idx in df_db.index:
                ma_cyk = CYKLY[int(df_db.loc[idx, 'Zmena'])][(curr_d - START_REF).days % 8] == smena_target
                penalty = 10000 if hod_fond_sofar[idx] >= fond_limit else 0
                fond_score = -hod_fond_sofar[idx] if is_75_poz else hod_fond_sofar[idx]
                pool.append((idx, (0 if ma_cyk else 1, penalty, fond_score, random.random())))
            return [x[0] for x in sorted(pool, key=lambda x: x[1])]

        if is_workday:
            nas_z8 = False
            for col_f in ["Priorita_Z8", "Z8"]:
                if nas_z8: break
                for idx in get_prioritized_people('D', True):
                    if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                    ab = abs_map.get(str(df_db.loc[idx, 'Priezvisko']).strip(), {'d':set(),'kz':set(),'v':set()})
                    if d not in (ab['d']|ab['kz']|ab['v']) and str(df_db.loc[idx].get(col_f,'Nie')).lower() == 'áno':
                        vysledky[d]['D'][idx] = "Z8"; hod_fond_sofar[idx] += 7.5; nas_z8 = True; break

        for smena in ['D', 'N']:
            for idx in df_db.index:
                if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                if CYKLY[int(df_db.loc[idx, 'Zmena'])][(curr_d - START_REF).days % 8] == smena and str(df_db.loc[idx].get('C1','Nie')).lower() == 'áno':
                    ab = abs_map.get(str(df_db.loc[idx, 'Priezvisko']).strip(), {'d':set(),'kz':set(),'v':set()})
                    if d not in (ab['d']|ab['kz']|ab['v']) and moze_nastupit(idx, d, smena, 'C1', vysledky):
                        vysledky[d][smena][idx] = 'C1'; hod_fond_sofar[idx] += 11.5; break

            for p_n in ['ZT', 'NB']:
                if p_n in vysledky[d][smena].values() or (p_n == 'NB' and smena == 'D' and is_workday): continue
                pool, nas = get_prioritized_people(smena), False
                for idx in pool:
                    if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                    if str(df_db.loc[idx].get(f"Priorita_{p_n}",'Nie')).lower() == 'áno' and CYKLY[int(df_db.loc[idx, 'Zmena'])][(curr_d - START_REF).days % 8] == smena:
                        ab = abs_map.get(str(df_db.loc[idx, 'Priezvisko']).strip(), {'d':set(),'kz':set(),'v':set()})
                        if d not in (ab['d']|ab['kz']|ab['v']) and moze_nastupit(idx, d, smena, p_n, vysledky):
                            vysledky[d][smena][idx] = p_n; hod_fond_sofar[idx] += 11.5; nas = True; break
                if not nas:
                    for idx in pool:
                        if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                        if str(df_db.loc[idx].get(p_n,'Nie')).lower() == 'áno' and CYKLY[int(df_db.loc[idx, 'Zmena'])][(curr_d - START_REF).days % 8] == smena:
                            ab = abs_map.get(str(df_db.loc[idx, 'Priezvisko']).strip(), {'d':set(),'kz':set(),'v':set()})
                            if d not in (ab['d']|ab['kz']|ab['v']) and moze_nastupit(idx, d, smena, p_n, vysledky):
                                vysledky[d][smena][idx] = p_n; hod_fond_sofar[idx] += 11.5; nas = True; break

            if smena == 'D' and is_workday:
                specs = (['TP', 'S1', 'S2', 'S3'] if parl_active and p_from <= curr_d <= p_to else []) + (['W_EXTRA'] if use_extra_w else []) + ['M']
                for poz in specs:
                    if poz in vysledky[d]['D'].values(): continue
                    for idx in get_prioritized_people('D'):
                        if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                        p_col = poz if poz != 'W_EXTRA' else 'W1'
                        if str(df_db.loc[idx].get(p_col,'Nie')).lower() == 'áno':
                            ab = abs_map.get(str(df_db.loc[idx, 'Priezvisko']).strip(), {'d':set(),'kz':set(),'v':set()})
                            if d not in (ab['d']|ab['kz']|ab['v']) and moze_nastupit(idx, d, 'D', poz, vysledky):
                                vysledky[d]['D'][idx] = poz; hod_fond_sofar[idx] += 11.5; break

            for poz in PRIO_LIST:
                if poz in vysledky[d][smena].values(): continue
                for idx in get_prioritized_people(smena):
                    if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                    if str(df_db.loc[idx].get(poz,'Nie')).lower() == 'áno':
                        ab = abs_map.get(str(df_db.loc[idx, 'Priezvisko']).strip(), {'d':set(),'kz':set(),'v':set()})
                        if d not in (ab['d']|ab['kz']|ab['v']) and moze_nastupit(idx, d, smena, poz, vysledky):
                            vysledky[d][smena][idx] = poz; hod_fond_sofar[idx] += 11.5; break

        if is_workday:
            wa = (((curr_d - START_REF).days // 7) % 2 == 0)
            trg = "IR" if (wa and curr_d.weekday() <= 1) or (not wa and curr_d.weekday() >= 2) else "IP"
            for idx in get_prioritized_people('D', True):
                if idx in vysledky[d]['D'] or idx in vysledky[d]['N']: continue
                ab = abs_map.get(str(df_db.loc[idx, 'Priezvisko']).strip(), {'d':set(),'kz':set(),'v':set()})
                if d not in (ab['d']|ab['kz']|ab['v']):
                    fx = trg if str(df_db.loc[idx].get(trg,'Nie')).lower() == 'áno' else next((p for p in ['X'] if str(df_db.loc[idx].get(p,'Nie')).lower() == 'áno'), None)
                    if fx: vysledky[d]['D'][idx] = fx; hod_fond_sofar[idx] += 7.5

    ws.set_column(0, 0, 25)
    for d in range(1, days_count + 1): ws.set_column(d, d, 3.5)
    col_bg_map = {day: ('#40B4EE' if date(r,m,day) in SVIATKY_2026 else ('#FFCC66' if date(r,m,day).weekday()==5 else ('#CC9900' if date(r,m,day).weekday()==6 else '#FFFFFF'))) for day in range(1, days_count+1)}
    for day in range(1, days_count+1): ws.write(0, day, day, wb.add_format({**fmt_b, 'bg_color': col_bg_map[day]}))
    ws.write(0, days_count+1, "Sumár", wb.add_format({'bold':True, 'border':1})); ws.write(0, days_count+2, "Rozdiel", wb.add_format({'bold':True, 'border':1}))
    ZZ = days_count + 10

    for i, (idx, row) in enumerate(df_db.iterrows()):
        zebra, row_ptr = ('#FFFF00' if i % 2 == 1 else '#FFFFFF'), i*2+1
        ws.merge_range(row_ptr, 0, row_ptr+1, 0, f"{row['Priezvisko']} {row['Meno']}", z_fmts[str(int(row['Zmena']))])
        ws.write(row_ptr, ZZ, int(row['Zmena']))
        ab = abs_map.get(str(row['Priezvisko']).strip(), {'d':set(), 'kz':set(), 'v':set()})
        for d in range(1, days_count + 1):
            bg = col_bg_map[d] if col_bg_map[d] != '#FFFFFF' else zebra
            cyk_char = CYKLY[int(row['Zmena'])][(date(r, m, d) - START_REF).days % 8]
            if d in ab['d']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'D', f_d)
            elif d in ab['kz']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'KZ', f_kz)
            elif d in ab['v']: ws.merge_range(row_ptr, d, row_ptr+1, d, 'V', f_v)
            else:
                pd, pn = vysledky[d]['D'].get(idx, ""), vysledky[d]['N'].get(idx, "")
                ps, ns = short_label(pd), short_label(pn)
                ws.write(row_ptr, d, ps, f_c1_d if ps=='C' else wb.add_format({**fmt_b, 'bg_color': bg, 'bold': bool(ps) and cyk_char != 'D'}))
                ws.write(row_ptr+1, d, ns, f_c1_n if ns=='C' else wb.add_format({**fmt_sep, 'bg_color': bg, 'bold': bool(ns) and cyk_char != 'N'}))

        r_ex, zz_col = row_ptr + 1, xlsxwriter.utility.xl_col_to_name(ZZ)
        sc, ec = xlsxwriter.utility.xl_col_to_name(1), xlsxwriter.utility.xl_col_to_name(days_count)
        f_parts = [f"IF(OR({xlsxwriter.utility.xl_col_to_name(day)}{r_ex}=\"D\",{xlsxwriter.utility.xl_col_to_name(day)}{r_ex}=\"KZ\"),IF(OR(MID(CHOOSE({zz_col}{r_ex},\"{CYKLY[1]}\",\"{CYKLY[2]}\",\"{CYKLY[3]}\",\"{CYKLY[4]}\"),{((date(r,m,day)-START_REF).days%8)+1},1)=\"D\",MID(CHOOSE({zz_col}{r_ex},\"{CYKLY[1]}\",\"{CYKLY[2]}\",\"{CYKLY[3]}\",\"{CYKLY[4]}\"),{((date(r,m,day)-START_REF).days%8)+1},1)=\"N\"),11.5,0),0)" for day in range(1, days_count+1)]
        full_formula = f"=(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"*\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"R\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"K\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"X\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"Z8\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"D\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"KZ\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"V\")*11.5)+({'+'.join(f_parts)})"
        ws.merge_range(row_ptr, days_count+1, row_ptr+1, days_count+1, full_formula, fmt_num)
        ws.merge_range(row_ptr, days_count+2, row_ptr+1, days_count+2, f"={fond_limit}-{xlsxwriter.utility.xl_rowcol_to_cell(row_ptr, days_count+1)}", fmt_num)
        ws.conditional_format(row_ptr, days_count+2, row_ptr+1, days_count+2, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_low})

    wb.close()
    return output.getvalue(), f"Plan_{m}_{r}.xlsx"

# --- 4. UI (Streamlit) ---
st.set_page_config(page_title="Plánovač 2026", layout="wide")
st.title("🚀 Smart Plánovač 2026")

if os.path.exists(DB_FILENAME):
    if 'df_db' not in st.session_state:
        ex = pd.ExcelFile(DB_FILENAME)
        st.session_state.df_db = ex.parse('Data').dropna(subset=['Priezvisko'])
        df_v = ex.parse('Volno') if 'Volno' in ex.sheet_names else pd.DataFrame(columns=['Priezvisko', 'Meno', 'Dovolenka', 'KZ', 'Volno'])
        for col in ['Dovolenka', 'KZ', 'Volno']: df_v[col] = df_v[col].fillna("").astype(str).replace('nan', '')
        st.session_state.df_v = df_v

    t1, t2 = st.tabs(["📊 Plánovanie", "⚙️ Databáza"])
    with t2:
        st.session_state.df_db = st.data_editor(st.session_state.df_db, use_container_width=True, key="db_ed")
        if st.button("💾 ULOŽIŤ PERSONÁL"):
            if push_to_github(st.session_state.df_db, st.session_state.df_v): st.success("Uložené na GitHub!")
    with t1:
        c1, c2, c3, c4 = st.columns(4)
        mes = c1.selectbox("Mesiac", range(1, 13), index=datetime.now().month-1)
        fon = c2.number_input("Fond", value=155.0)
        parl, extra_w = c3.checkbox("Parlament", True), c4.checkbox("Extra W", True)
        _, last_day = calendar.monthrange(2026, mes)
        p_od = st.date_input("Od", date(2026, mes, 1), format="DD.MM.YYYY")
        p_do = st.date_input("Do", date(2026, mes, last_day), format="DD.MM.YYYY")
        
        st.session_state.df_v = st.data_editor(st.session_state.df_v, use_container_width=True, key="v_ed")
        if st.button("💾 ULOŽIŤ ABSENCIE"):
            if push_to_github(st.session_state.df_db, st.session_state.df_v): st.success("Uložené na GitHub!")
        
        if st.button("🚀 GENEROVAŤ PLÁN", type="primary", use_container_width=True):
            xlsx, name = generuj_final(mes, 2026, fon, parl, p_od, p_do, st.session_state.df_v, extra_w, st.session_state.df_db)
            st.download_button("📥 STIAHNUŤ", data=xlsx, file_name=name, use_container_width=True)
else:
    st.error(f"Súbor {DB_FILENAME} sa nenašiel v priečinku.")
