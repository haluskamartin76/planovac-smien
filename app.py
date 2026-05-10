# --- UPRAVENÉ NAČÍTANIE DÁT S AUTOMATICKOU OČISTOU ---
if os.path.exists(DB_FILENAME):
    try:
        ex = pd.ExcelFile(DB_FILENAME, engine='openpyxl')
        
        # Načítame a hneď orežeme prázdne riadky a stĺpce
        df_db_raw = ex.parse('Data').dropna(how='all', axis=0).dropna(how='all', axis=1)
        # Necháme len riadky, kde je vyplnené Priezvisko
        df_db_raw = df_db_raw.dropna(subset=['Priezvisko'])
        
        df_v_raw = ex.parse('Volno').dropna(how='all', axis=0).dropna(how='all', axis=1)
        if not df_v_raw.empty:
            df_v_raw = df_v_raw.dropna(subset=['Priezvisko'])
            
        # OREŽEME STĹPCE - necháme len tie, ktoré kód reálne používa
        # Týmto sa zbavíme všetkých skrytých stĺpcov napravo
        potrebne_stlpce = ['Priezvisko', 'Meno', 'Zmena', 'C1', 'Z8', 'ZT', 'NB', 'Priorita_Z8', 'Priorita_ZT', 'Priorita_NB', 'IR', 'IP', 'X', 'TP', 'S1', 'S2', 'S3', 'M'] + PRIO_LIST
        existujuce = [c for c in potrebne_stlpce if c in df_db_raw.columns]
        df_db_raw = df_db_raw[existujuce]

        st.success(f"✅ Databáza vyčistená a načítaná ({len(df_db_raw)} zamestnancov).")
