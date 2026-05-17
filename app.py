import streamlit as st
import pandas as pd
import calendar, random, io, os, xlsxwriter
from datetime import date, datetime

# --- DYNAMICKÁ KONFIGURÁCIA SVIATKOV ---
def ziskaj_sviatky(rok):
    a = rok % 19
    b = rok % 4
    c = rok % 7
    d = (19 * a + 24) % 30
    e = (2 * b + 4 * c + 6 * d + 5) % 7
    velka_noc_dni = 22 + d + e
    
    if velka_noc_dni > 31:
        mesiac_vn = 4
        den_vn = velka_noc_dni - 31
    else:
        mesiac_vn = 3
        den_vn = velka_noc_dni

    if d == 29 and e == 6:
        den_vn = 19; mesiac_vn = 4
    elif d == 28 and e == 6 and a > 10:
        den_vn = 18; mesiac_vn = 4

    v_nedela = date(rok, mesiac_vn, den_vn)
    v_piatok = date.fromordinal(v_nedela.toordinal() - 2)
    v_pondelok = date.fromordinal(v_nedela.toordinal() + 1)

    return {
        date(rok, 1, 1), date(rok, 1, 6), v_piatok, v_pondelok,
        date(rok, 5, 1), date(rok, 5, 8), date(rok, 7, 5), date(rok, 8, 29),
        date(rok, 9, 1), date(rok, 9, 15), date(rok, 11, 1), date(rok, 11, 17),
        date(rok, 12, 24), date(rok, 12, 25), date(rok, 12, 26)
    }

PRIO_LIST = ['C2', 'W1', 'W2', 'Z1', 'Z2', 'G', 'GH', 'SH']
START_REF = date(2026, 3, 1)
CYKLY = {1: "DNVDNVVV", 2: "VVDNVDNV", 3: "VDNVVVDN", 4: "NVVVDNVD"}

def parse_days(s):
    res = set()
    if pd.isna(s) or str(s).lower() == 'nan' or str(s).strip() == "" or str(s).lower() == 'none': return res
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
    m = {
        "C1":"C", "C2":"C", 
        "Z1":"Z", "Z2":"Z", 
        "W1":"W", "W2":"W", "W_EXTRA":"W", 
        "S1":"S", "S2":"S", "S3":"S", 
        "IR":"R", "IP":"K"
    }
    return m.get(lbl, lbl)

def moze_nastupit(zamestnanec_id, d, smena, poz, vysledky):
    if smena == 'D' and d > 1 and zamestnanec_id in vysledky[d-1]['N']: return False
    if poz == 'C1': return True
    if d > 2 and zamestnanec_id in vysledky[d-1][smena] and zamestnanec_id in vysledky[d-2][smena]: return False
    return True

def generuj_final(m, r, fond_limit, parl_active, p_from, p_to, df_v_edit, use_extra_w, df_db):
    sviatky_aktualne = ziskaj_sviatky(r)
    df_db = df_db.sort_values(by='Povodne_Poradie')
    
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output)
    ws = wb.add_worksheet("Plán")
    ws_miss = wb.add_worksheet("Neobsadené")

    fmt_b = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': False, 'font_size': 9}
    fmt_sep = {**fmt_b, 'bottom': 2}
    z_fmts = {str(i): wb.add_format({**fmt_sep, 'bg_color': c, 'font_color': fc})
              for i, c, fc in zip([1, 2, 3, 4], ['#B2B2B2','#FF0000','#FFFF00','#003399'], ['white','white','black','white'])}

    # Základné formáty pre absencie
    f_d = wb.add_format({**fmt_sep, 'bg_color': '#339933', 'font_color': 'white'})
    f_kz = wb.add_format({**fmt_sep, 'bg_color': '#0066FF', 'font_color': 'white'})
    f_v = wb.add_format({**fmt_sep, 'bg_color': '#00FFCC'})
    
    # OPRAVENÉ: Kľúčové slovo zmenené na pattern_background_color podľa špecifikácie xlsxwriter
    f_d_vzor = wb.add_format({**fmt_sep, 'bg_color': '#339933', 'font_color': 'white', 'pattern': 4, 'pattern_background_color': '#226622'})
    f_kz_vzor = wb.add_format({**fmt_sep, 'bg_color': '#0066FF', 'font_color': 'white', 'pattern': 4, 'pattern_background_color': '#003399'})
    f_v_vzor = wb.add_format({**fmt_sep, 'bg_color': '#00FFCC', 'pattern': 4, 'pattern_background_color': '#00BB99'})

    fmt_num = wb.add_format({**fmt_sep, 'num_format': '#,##0.0'})
    f_c1_d = wb.add_format({**fmt_b, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
    f_c1_n = wb.add_format({**fmt_sep, 'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
    fmt_low = wb.add_format({**fmt_sep, 'bg_color': '#FF9900', 'num_format': '#,##0.0'})

    _, days_count = calendar.monthrange(r, m)
    vysledky = {d: {'D': {}, 'N': {}} for d in range(1, days_count + 1)}
    
    ludia = []
    for i, row in df_db.iterrows():
        ludia.append({
            'id': i,
            'Priezvisko': str(row['Priezvisko']).strip(),
            'Meno': str(row.get('Meno', '')).strip(),
            'Zmena': int(row['Zmena']),
            'C1': str(row.get('C1', 'Nie')).lower(),
            'Priorita_ZT': str(row.get('Priorita_ZT', 'Nie')).lower(),
            'ZT': str(row.get('ZT', 'Nie')).lower(),
            'Priorita_NB': str(row.get('Priorita_NB', 'Nie')).lower(),
            'NB': str(row.get('NB', 'Nie')).lower(),
            'Priorita_Z8': str(row.get('Priorita_Z8', 'Nie')).lower(),
            'Z8': str(row.get('Z8', 'Nie')).lower(),
            'TP': str(row.get('TP', 'Nie')).lower(),
            'S1': str(row.get('S1', 'Nie')).lower(),
            'S2': str(row.get('S2', 'Nie')).lower(),
            'S3': str(row.get('S3', 'Nie')).lower(),
            'W1': str(row.get('W1', 'Nie')).lower(),
            'M': str(row.get('M', 'Nie')).lower(),
            'IR': str(row.get('IR', 'Nie')).lower(),
            'IP': str(row.get('IP', 'Nie')).lower(),
            'X': str(row.get('X', 'Nie')).lower(),
            'original_row': row
        })

    hod_fond_sofar = {p['id']: 0.0 for p in ludia}
    neobsadene_zaznamy = []

    abs_map = {}
    for _, row in df_v_edit.iterrows():
        priez = str(row['Priezvisko']).strip()
        abs_map[priez] = {'d': parse_days(row['Dovolenka']), 'kz': parse_days(row['KZ']), 'v': parse_days(row['Volno'])}

    for p in ludia:
        ab = abs_map.get(p['Priezvisko'], {'d':set(), 'kz':set()})
        abs_dni = ab['d'] | ab['kz']
        for d_val in abs_dni:
            if d_val <= days_count:
                if CYKLY[p['Zmena']][(date(r, m, d_val) - START_REF).days % 8] in ['D', 'N']:
                    hod_fond_sofar[p['id']] += 11.5

    for d in range(1, days_count + 1):
        curr_d = date(r, m, d)
        is_workday = curr_d.weekday() < 5 and curr_d not in sviatky_aktualne

        def get_prioritized_people(smena_target, is_75_poz=False):
            pool = []
            for p in ludia:
                ma_cyk = CYKLY[p['Zmena']][(curr_d - START_REF).days % 8] == smena_target
                penalty = 10000 if hod_fond_sofar[p['id']] >= fond_limit else 0
                fond_score = -hod_fond_sofar[p['id']] if is_75_poz else hod_fond_sofar[p['id']]
                pool.append((p, (0 if ma_cyk else 1, penalty, fond_score, random.random())))
            return [x[0] for x in sorted(pool, key=lambda x: x[1])]

        if is_workday:
            nas_z8 = False
            for col_f in ["Priorita_Z8", "Z8"]:
                if nas_z8: break
                for p in get_prioritized_people('D', True):
                    if p['id'] in vysledky[d]['D'] or p['id'] in vysledky[d]['N']: continue
                    ab = abs_map.get(p['Priezvisko'], {'d':set(),'kz':set(),'v':set()})
                    cv = ab['d'] | ab['kz'] | ab['v']
                    if p[col_f] == 'áno':
                        if d not in cv:
                            vysledky[d]['D'][p['id']] = "Z8"; hod_fond_sofar[p['id']] += 7.5; nas_z8 = True; break
                        else:
                            neobsadene_zaznamy.append((d, 'D', 'Z8'))

        for smena in ['D', 'N']:
            for p in ludia:
                if p['id'] in vysledky[d]['D'] or p['id'] in vysledky[d]['N']: continue
                if CYKLY[p['Zmena']][(curr_d - START_REF).days % 8] == smena and p['C1'] == 'áno':
                    ab = abs_map.get(p['Priezvisko'], {'d':set(),'kz':set(),'v':set()})
                    cv = ab['d'] | ab['kz'] | ab['v']
                    if d not in cv and moze_nastupit(p['id'], d, smena, 'C1', vysledky):
                        vysledky[d][smena][p['id']] = 'C1'; hod_fond_sofar[p['id']] += 11.5; break
                    elif d in cv and moze_nastupit(p['id'], d, smena, 'C1', vysledky):
                        neobsadene_zaznamy.append((d, smena, 'C1'))

            for p_n in ['ZT', 'NB']:
                if p_n in vysledky[d][smena].values() or (p_n == 'NB' and smena == 'D' and is_workday): continue
                nas = False
                pool = get_prioritized_people(smena)
                for p in pool:
                    if p['id'] in vysledky[d]['D'] or p['id'] in vysledky[d]['N']: continue
                    if p[f"Priorita_{p_n}"] == 'áno' and CYKLY[p['Zmena']][(curr_d - START_REF).days % 8] == smena:
                        ab = abs_map.get(p['Priezvisko'], {'d':set(),'kz':set(),'v':set()})
                        cv = ab['d'] | ab['kz'] | ab['v']
                        if d not in cv and moze_nastupit(p['id'], d, smena, p_n, vysledky):
                            vysledky[d][smena][p['id']] = p_n; hod_fond_sofar[p['id']] += 11.5; nas = True; break
                        elif d in cv and moze_nastupit(p['id'], d, smena, p_n, vysledky):
                            neobsadene_zaznamy.append((d, smena, p_n))
                if not nas:
                    for p in pool:
                        if p['id'] in vysledky[d]['D'] or p['id'] in vysledky[d]['N']: continue
                        if p[p_n] == 'áno' and CYKLY[p['Zmena']][(curr_d - START_REF).days % 8] == smena:
                            ab = abs_map.get(p['Priezvisko'], {'d':set(),'kz':set(),'v':set()})
                            cv = ab['d'] | ab['kz'] | ab['v']
                            if d not in cv and moze_nastupit(p['id'], d, smena, p_n, vysledky):
                                vysledky[d][smena][p['id']] = p_n; hod_fond_sofar[p['id']] += 11.5; nas = True; break
                            elif d in cv and moze_nastupit(p['id'], d, smena, p_n, vysledky):
                                neobsadene_zaznamy.append((d, smena, p_n))

            if smena == 'D' and is_workday:
                specs = (['TP', 'S1', 'S2', 'S3'] if parl_active and p_from <= curr_d <= p_to and curr_d.weekday() not in [0, 5, 6] else []) + (['W_EXTRA'] if use_extra_w else []) + ['M']
                for poz in specs:
                    if poz in vysledky[d]['D'].values(): continue
                    for p in get_prioritized_people('D'):
                        if p['id'] in vysledky[d]['D'] or p['id'] in vysledky[d]['N']: continue
                        ab = abs_map.get(p['Priezvisko'], {'d':set(),'kz':set(),'v':set()})
                        cv = ab['d'] | ab['kz'] | ab['v']
                        p_col = poz if poz != 'W_EXTRA' else 'W1'
                        if p[p_col] == 'áno' and moze_nastupit(p['id'], d, 'D', poz, vysledky):
                            if d not in cv:
                                vysledky[d]['D'][p['id']] = poz; hod_fond_sofar[p['id']] += 11.5; break
                            else:
                                neobsadene_zaznamy.append((d, 'D', poz))

            for poz in PRIO_LIST:
                if poz in vysledky[d][smena].values(): continue
                for p in get_prioritized_people(smena):
                    if p['id'] in vysledky[d]['D'] or p['id'] in vysledky[d]['N']: continue
                    ab = abs_map.get(p['Priezvisko'], {'d':set(),'kz':set(),'v':set()})
                    cv = ab['d'] | ab['kz'] | ab['v']
                    if p['original_row'].get(poz,'Nie').lower() == 'áno' and moze_nastupit(p['id'], d, smena, poz, vysledky):
                        if d not in cv:
                            vysledky[d][smena][p['id']] = poz; hod_fond_sofar[p['id']] += 11.5; break
                        else:
                            neobsadene_zaznamy.append((d, smena, poz))

        if is_workday:
            wa = (((curr_d - START_REF).days // 7) % 2 == 0)
            trg = "IR" if (wa and curr_d.weekday() <= 1) or (not wa and curr_d.weekday() >= 2) else "IP"
            for p in get_prioritized_people('D', True):
                if p['id'] in vysledky[d]['D'] or p['id'] in vysledky[d]['N']: continue
                ab = abs_map.get(p['Priezvisko'], {'d':set(),'kz':set(),'v':set()})
                cv = ab['d'] | ab['kz'] | ab['v']
                if d not in cv:
                    fx = trg if p[trg] == 'áno' else next((x for x in ['X'] if p[x] == 'áno'), None)
                    if fx: vysledky[d]['D'][p['id']] = fx; hod_fond_sofar[p['id']] += 7.5
                else:
                    fx = trg if p[trg] == 'áno' else next((x for x in ['X'] if p[x] == 'áno'), None)
                    if fx: neobsadene_zaznamy.append((d, 'D', fx))

    ws.set_column(0, 0, 25)
    for d in range(1, days_count + 1): ws.set_column(d, d, 3.5)
    ws.set_column(days_count+1, days_count+2, 10)
    ZZ = days_count + 10
    ws.set_column(ZZ, ZZ, None, None, {'hidden': True})

    col_bg_map = {d: ('#40B4EE' if date(r,m,d) in sviatky_aktualne else ('#FFCC66' if date(r,m,d).weekday()==5 else ('#CC9900' if date(r,m,d).weekday()==6 else '#FFFFFF'))) for d in range(1, days_count+1)}
    for d in range(1, days_count + 1):
        ws.write(0, d, d, wb.add_format({**fmt_b, 'bg_color': col_bg_map[d]}))
    ws.write(0, days_count+1, "Sumár", wb.add_format({'bold':True, 'border':1}))
    ws.write(0, days_count+2, "Rozdiel", wb.add_format({'bold':True, 'border':1}))

    for i, p in enumerate(ludia):
        zebra = '#FFFF00' if i % 2 == 1 else '#FFFFFF'
        row_ptr = i*2+1
        ws.merge_range(row_ptr, 0, row_ptr+1, 0, f"{p['Priezvisko']} {p['Meno']}", z_fmts[str(p['Zmena'])])
        ws.write(row_ptr, ZZ, p['Zmena'])
        ab = abs_map.get(p['Priezvisko'], {'d':set(), 'kz':set(), 'v':set()})
        
        for d in range(1, days_count + 1):
            bg = col_bg_map[d] if col_bg_map[d] != '#FFFFFF' else zebra
            cyk_char = CYKLY[p['Zmena']][(date(r, m, d) - START_REF).days % 8]
            ma_mat_smenu = cyk_char in ['D', 'N']
            
            if d in ab['d']:
                fmt = f_d_vzor if ma_mat_smenu else f_d
                ws.merge_range(row_ptr, d, row_ptr+1, d, 'D', fmt)
            elif d in ab['kz']:
                fmt = f_kz_vzor if ma_mat_smenu else f_kz
                ws.merge_range(row_ptr, d, row_ptr+1, d, 'KZ', fmt)
            elif d in ab['v']:
                fmt = f_v_vzor if ma_mat_smenu else f_v
                ws.merge_range(row_ptr, d, row_ptr+1, d, 'V', fmt)
            else:
                pd, pn = vysledky[d]['D'].get(p['id'], ""), vysledky[d]['N'].get(p['id'], "")
                ps, ns = short_label(pd), short_label(pn)
                ws.write(row_ptr, d, ps, f_c1_d if ps=='C' else wb.add_format({**fmt_b, 'bg_color': bg, 'bold': bool(ps) and cyk_char != 'D'}))
                ws.write(row_ptr+1, d, ns, f_c1_n if ns=='C' else wb.add_format({**fmt_sep, 'bg_color': bg, 'bold': bool(ns) and cyk_char != 'N'}))

        r_ex, zz_col = row_ptr + 1, xlsxwriter.utility.xl_col_to_name(ZZ)
        sc, ec = xlsxwriter.utility.xl_col_to_name(1), xlsxwriter.utility.xl_col_to_name(days_count)
        f_parts = [f"IF(OR({xlsxwriter.utility.xl_col_to_name(d)}{r_ex}=\"D\",{xlsxwriter.utility.xl_col_to_name(d)}{r_ex}=\"KZ\"),IF(OR(MID(CHOOSE({zz_col}{r_ex},\"{CYKLY[1]}\",\"{CYKLY[2]}\",\"{CYKLY[3]}\",\"{CYKLY[4]}\"),{((date(r,m,d)-START_REF).days%8)+1},1)=\"D\",MID(CHOOSE({zz_col}{r_ex},\"{CYKLY[1]}\",\"{CYKLY[2]}\",\"{CYKLY[3]}\",\"{CYKLY[4]}\"),{((date(r,m,d)-START_REF).days%8)+1},1)=\"N\"),11.5,0),0)" for d in range(1, days_count+1)]
        full_formula = f"=(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"*\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"R\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"K\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"X\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"Z8\")*4)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"D\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"KZ\")*11.5)-(COUNTIF({sc}{r_ex}:{ec}{r_ex+1},\"V\")*11.5)+({'+'.join(f_parts)})"
        ws.merge_range(row_ptr, days_count+1, row_ptr+1, days_count+1, full_formula, fmt_num)
        sum_c = xlsxwriter.utility.xl_rowcol_to_cell(row_ptr, days_count+1)
        ws.merge_range(row_ptr, days_count+2, row_ptr+1, days_count+2, f"={fond_limit}-{sum_c}", fmt_num)
        ws.conditional_format(row_ptr, days_count+2, row_ptr+1, days_count+2, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_low})

    # --- ZÁPIS NEOBSADENÝCH POZÍCIÍ NA SAMOSTATNÝ HÁROK ---
    ws_miss.write_row(0, 0, ["Deň", "Smena", "Pozícia"], wb.add_format({'bold':True, 'border':1}))
    
    unikatne_neobsadene = sorted(list(set(neobsadene_zaznamy)), key=lambda x: x)
    
    skutocny_row = 1
    for den_m, smena_m, poz_m in unikatne_neobsadene:
        aktualne_obsadene = [short_label(x) for x in vysledky[den_m][smena_m].values()]
        if short_label(poz_m) not in aktualne_obsadene:
            ws_miss.write_row(skutocny_row, 0, [den_m, smena_m, short_label(poz_m)])
            skutocny_row += 1
    
    wb.close()
    return output.getvalue(), f"Plan_{m}_{r}.xlsx"

# --- UI STREAMLIT ---
st.set_page_config(page_title="Plánovač 2026", layout="wide")
st.title("🚀 Plánovač Zmien")

uploaded_file = st.file_uploader("Nahraj databaza_pozicii.xlsx", type="xlsx")

if uploaded_file:
    ex = pd.ExcelFile(uploaded_file)
    
    df_db = ex.parse('Data').dropna(subset=['Priezvisko'])
    df_db['Povodne_Poradie'] = range(len(df_db))
    
    df_v = ex.parse('Volno') if 'Volno' in ex.sheet_names else pd.DataFrame(columns=['Priezvisko', 'Meno', 'Dovolenka', 'KZ', 'Volno'])
    for col in ['Dovolenka', 'KZ', 'Volno']: df_v[col] = df_v[col].fillna("").astype(str).replace(['nan', 'None'], '')
    
    df_v = df_v.merge(df_db[['Priezvisko', 'Meno', 'Povodne_Poradie']], on=['Priezvisko', 'Meno'], how='left')

    c1, c2, c3, c4, c5 = st.columns(5)
    rok = c1.number_input("Rok", min_value=2020, max_value=2100, value=2026)
    mes = c2.selectbox("Mesiac", range(1, 13), index=2)
    fon = c3.number_input("Fond", value=155.0)
    parl, extra_w = c4.checkbox("Parlament", True), c5.checkbox("Extra W", True)
    
    _, last_day = calendar.monthrange(rok, mes)
    p_od = st.date_input("Parlament Od", date(rok, mes, 1), format="DD/MM/YYYY")
    p_do = st.date_input("Parlament Do", date(rok, mes, last_day), format="DD/MM/YYYY")
    
    st.subheader("Uprav absencie (zoradené abecedne pre pohodlie)")
    
    df_v_alphabetical = df_v.sort_values(by=['Priezvisko', 'Meno'])
    
    df_v_edit = st.data_editor(
        df_v_alphabetical, 
        use_container_width=True, 
        key="volno_edit",
        column_config={
            "Priezvisko": st.column_config.Column(disabled=True),
            "Meno": st.column_config.Column(disabled=True),
            "Povodne_Poradie": None 
        }
    )
    
    if st.button("🚀 GENEROVAŤ PLÁN", type="primary", use_container_width=True):
        xlsx, name = generuj_final(mes, rok, fon, parl, p_od, p_do, df_v_edit, extra_w, df_db)
        st.download_button("📥 STIAHNUŤ VYGENEROVANÝ PLÁN", data=xlsx, file_name=name, use_container_width=True)
