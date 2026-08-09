import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import graphviz
import streamlit.components.v1 as components
from datetime import datetime, date
import json
from io import BytesIO
from docx import Document

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="WBS/OBS Manager & EVM", layout="wide")
st.title("🏗️ Project Workflow & EVM Controller")

# --- SISTEMA DI LOGIN SICURO (TRAMITE SECRETS) ---
try:
    USER_ID = st.secrets["USER_ID"]
    PASSWORD = st.secrets["PASSWORD"]
except KeyError:
    st.error("⚠️ Errore di sistema: Credenziali non trovate. Configura i 'Secrets' di Streamlit.")
    st.stop()

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown("<br><br><h2 style='text-align: center;'>🔒 Accesso Riservato GECO</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            user_input = st.text_input("ID Utente")
            pass_input = st.text_input("Password", type="password")
            submit = st.form_submit_button("Accedi al Gestionale", use_container_width=True)
            
            if submit:
                if user_input == USER_ID and pass_input == PASSWORD:
                    st.session_state.logged_in = True
                    st.rerun() 
                else:
                    st.error("Credenziali errate. Riprova.")                
    st.stop()

# --- 1. INIZIALIZZAZIONE DATI (Session State) ---
if 'wbs_data' not in st.session_state:
    st.session_state.wbs_data = pd.DataFrame({
        'ID_WBS': ['1', '1.1', '1.2', '2', '2.1', '2.2', '3', '3.1', '3.1.1', '3.1.2', '3.2', '4', '4.1'],
        'Attività': [
            'Scavi', 'Scavi con mezzi meccanici', 'Scavi a mano', 
            'Strutture', 'Strutture in fondazione', 'Strutture in elevazione', 
            'Murature', 'Tamponatura esterna', 'Muratura a cassa vuota', 'Muratura in blocchi CLS', 'Tramezzatura interna',
            'Impianti', 'Impianto Elettrico'
        ],
        'Inizio_Previsto': [None, date(2026, 9, 1), date(2026, 9, 15), None, date(2026, 10, 1), date(2026, 10, 15), None, None, date(2026, 11, 1), date(2026, 11, 10), date(2026, 11, 20), None, date(2026, 12, 1)],
        'Fine_Prevista': [None, date(2026, 9, 14), date(2026, 9, 30), None, date(2026, 10, 14), date(2026, 11, 1), None, None, date(2026, 11, 9), date(2026, 11, 19), date(2026, 11, 30), None, date(2026, 12, 15)],
        'Inizio_Effettivo': [None, date(2026, 9, 2), None, None, None, None, None, None, None, None, None, None, None],
        'Fine_Effettiva': [None, date(2026, 9, 16), None, None, None, None, None, None, None, None, None, None, None],
        'BAC_Budget': [0.0, 5000.0, 2000.0, 0.0, 15000.0, 20000.0, 0.0, 0.0, 3000.0, 4000.0, 5000.0, 0.0, 8000.0],
        '%_Completamento': [0, 100, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        'AC_Costo_Reale': [0.0, 5200.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        'ID_OBS_Assegnato': [None, '1.1', '1.2', None, '1.1', '1.1', None, None, '1.2', '1.2', '1.2', None, '1.1'],
        'Predecessori': ['', '', '1.1', '', '', '2.1', '', '', '2.2', '3.1.1', '3.1.2', '', '2.2, 3.1.2']
    })
    
if 'obs_data' not in st.session_state:
    st.session_state.obs_data = pd.DataFrame({
        'ID_OBS': ['1.1', '1.2'],
        'Ruolo': ['Capo Cantiere', 'Strutturista'],
        'Risorsa': ['Mario Rossi', 'Studio Tecnico'],
        'Tipo_Contratto': ['Appalto ▾', 'Sub appalto ▾'], 
        'Note': ['', 'Ricordare DURC']            
    })
    
if 'registro_data' not in st.session_state:
    st.session_state.registro_data = pd.DataFrame({
        'Data': [date(2026, 9, 5), date(2026, 9, 10)],
        'N_Doc': ['FATT-01', 'FATT-02'],
        'Fornitore': ['Mario Rossi', 'Nolo Scavi Srl'],
        'Voce_WBS': ['1.1 - Scavi con mezzi meccanici', '1.1 - Scavi con mezzi meccanici'],
        'Importo_Netto': [2000.0, 3200.0],
        'Descrizione': ['Acconto lavori', 'Nolo escavatore']
    })

if 'capa_data' not in st.session_state:
    st.session_state.capa_data = pd.DataFrame({
        'Data_Apertura': [pd.Timestamp.today().date()],
        'ID_WBS_Rif': ['1.1 - Scavi con mezzi meccanici'],
        'Tipo_Azione': ['Correttiva ▾'],
        'Descrizione': ['Esempio: Valutare sostituzione fornitore per ritardi accumulati.'],
        'Responsabile_OBS': ['1.1 - Capo Cantiere'],
        'Stato': ['Aperto ▾']
    })

if 'archivio_progetti' not in st.session_state:
    st.session_state.archivio_progetti = {}
if 'nome_progetto_attivo' not in st.session_state:
    st.session_state.nome_progetto_attivo = "Progetto_01"

# --- 2. MOTORI MATEMATICI (Definiti Prima dei Tab) ---

def aggiorna_costi_reali():
    df_reg = st.session_state.registro_data.copy()
    df_reg['ID_WBS_calc'] = df_reg['Voce_WBS'].astype(str).apply(lambda x: x.split(' - ')[0] if ' - ' in x else None)
    costi_raggruppati = df_reg.groupby('ID_WBS_calc')['Importo_Netto'].sum().reset_index()
    cost_map = dict(zip(costi_raggruppati['ID_WBS_calc'], costi_raggruppati['Importo_Netto']))
    wbs = st.session_state.wbs_data
    wbs['AC_Costo_Reale'] = wbs['ID_WBS'].apply(lambda x: cost_map.get(str(x), 0.0))
    st.session_state.wbs_data = wbs

def calcola_evm(df, data_status):
    oggi = pd.to_datetime(data_status).date()
    df['BAC_Budget'] = pd.to_numeric(df['BAC_Budget'], errors='coerce').fillna(0.0)
    df['%_Completamento'] = pd.to_numeric(df['%_Completamento'], errors='coerce').fillna(0.0)
    df['AC_Costo_Reale'] = pd.to_numeric(df['AC_Costo_Reale'], errors='coerce').fillna(0.0)
    
    def calcola_pv(row):
        try:
            inizio_ts = pd.to_datetime(row['Inizio_Previsto'])
            fine_ts = pd.to_datetime(row['Fine_Prevista'])
            bac = float(row['BAC_Budget'])
            
            if pd.isna(inizio_ts) or pd.isna(fine_ts) or bac == 0: 
                return 0.0
                
            inizio = inizio_ts.date()
            fine = fine_ts.date()
            
            if oggi >= fine: return bac 
            if oggi <= inizio: return 0.0 
            
            giorni_totali = (fine - inizio).days
            giorni_trascorsi = (oggi - inizio).days
            
            if giorni_totali <= 0: return bac
            return bac * (giorni_trascorsi / giorni_totali)
            
        except Exception:
            return 0.0

    df['PV'] = df.apply(calcola_pv, axis=1)
    df['EV'] = df['BAC_Budget'] * (df['%_Completamento'] / 100.0)
    df['CV'] = df['EV'] - df['AC_Costo_Reale'] 
    df['SV'] = df['EV'] - df['PV']             
    
    df['SPI'] = df.apply(lambda x: (x['EV'] / x['PV']) if x['PV'] > 0 else (1.0 if x['EV']==0 else 1.1), axis=1)
    df['CPI'] = df.apply(lambda x: (x['EV'] / x['AC_Costo_Reale']) if x['AC_Costo_Reale'] > 0 else (1.0 if x['EV']==0 else 1.1), axis=1)

    df['EAC'] = df.apply(lambda x: x['BAC_Budget'] / x['CPI'] if x['CPI'] > 0 else x['BAC_Budget'], axis=1)
    df['ETC'] = df['EAC'] - df['AC_Costo_Reale']
    df['VAC'] = df['BAC_Budget'] - df['EAC']
    
    return df

def genera_dati_scurve(df_wbs, df_reg, data_status):
    oggi = pd.to_datetime(data_status).date()
    
    date_inizio = pd.to_datetime(df_wbs['Inizio_Previsto']).dropna().dt.date
    date_fine = pd.to_datetime(df_wbs['Fine_Prevista']).dropna().dt.date
    
    if date_inizio.empty or date_fine.empty:
        return None
        
    min_date = date_inizio.min()
    max_date = date_fine.max()
    
    date_range = pd.date_range(start=min_date, end=max_date)
    
    df_reg_calc = df_reg.copy()
    df_reg_calc['Data'] = pd.to_datetime(df_reg_calc['Data'], errors='coerce').dt.date
    ac_daily = df_reg_calc.groupby('Data')['Importo_Netto'].sum().to_dict()
    
    dati = []
    cum_ac = 0.0
    
    for d_ts in date_range:
        d = d_ts.date()
        pv_giorno = 0.0
        ev_giorno = 0.0
        
        for _, row in df_wbs.iterrows():
            bac = float(row['BAC_Budget']) if pd.notna(row['BAC_Budget']) else 0.0
            
            ip = pd.to_datetime(row['Inizio_Previsto']).date() if pd.notna(row['Inizio_Previsto']) else None
            fp = pd.to_datetime(row['Fine_Prevista']).date() if pd.notna(row['Fine_Prevista']) else None
            if ip and fp and bac > 0:
                if d >= fp: 
                    pv_giorno += bac
                elif d > ip:
                    giorni_tot = (fp - ip).days
                    if giorni_tot > 0:
                        pv_giorno += bac * ((d - ip).days / giorni_tot)
                        
            if d <= oggi:
                ev_attuale = float(row['EV']) if 'EV' in row else 0.0
                ie = pd.to_datetime(row['Inizio_Effettivo']).date() if pd.notna(row['Inizio_Effettivo']) else ip
                if ie and ev_attuale > 0:
                    if d >= oggi:
                        ev_giorno += ev_attuale
                    elif d > ie:
                        giorni_lav = (oggi - ie).days
                        if giorni_lav > 0:
                            ev_giorno += ev_attuale * ((d - ie).days / giorni_lav)
        
        if d <= oggi:
            cum_ac += ac_daily.get(d, 0.0)
            ac_val = cum_ac
            ev_val = ev_giorno
        else:
            ac_val = None
            ev_val = None
            
        dati.append({
            'Data': d,
            'PV (Valore Pianificato)': pv_giorno,
            'EV (Valore Guadagnato)': ev_val,
            'AC (Costo Reale)': ac_val
        })
        
    return pd.DataFrame(dati)

def calcola_cpm(df_wbs):
    # Filtriamo solo le lavorazioni operative
    df_wp = df_wbs[df_wbs['ID_WBS'].astype(str).str.contains('\.')].copy()
    cpm_nodes = {}
    
    # SETUP: Inizializziamo i nodi e calcoliamo le durate previste
    for _, row in df_wp.iterrows():
        node_id = str(row['ID_WBS']).strip()
        inizio = pd.to_datetime(row['Inizio_Previsto'], errors='coerce')
        fine = pd.to_datetime(row['Fine_Prevista'], errors='coerce')
        
        # Calcoliamo i giorni lavorativi
        if pd.notna(inizio) and pd.notna(fine):
            durata = max((fine - inizio).days, 1) # Minimo 1 giorno
        else:
            durata = 1
            
        preds = [p.strip() for p in str(row['Predecessori']).split(',')] if pd.notna(row['Predecessori']) and str(row['Predecessori']).strip() != '' else []
        
        cpm_nodes[node_id] = {
            'durata': durata,
            'preds': preds,
            'succs': [],
            'ES': 0, 'EF': 0, 'LS': 0, 'LF': 0, 'slack': 0,
            'is_critical': False
        }
        
    # Popoliamo i successori
    for n_id, data in cpm_nodes.items():
        for p_id in data['preds']:
            if p_id in cpm_nodes:
                cpm_nodes[p_id]['succs'].append(n_id)
                
    # FASE 1: FORWARD PASS (Andata)
    changed = True
    while changed:
        changed = False
        for n_id, data in cpm_nodes.items():
            max_ef = 0
            for p_id in data['preds']:
                if p_id in cpm_nodes:
                    max_ef = max(max_ef, cpm_nodes[p_id]['EF'])
            new_es = max_ef
            new_ef = new_es + data['durata']
            if new_es != data['ES'] or new_ef != data['EF']:
                data['ES'] = new_es
                data['EF'] = new_ef
                changed = True
                
    # FASE 2: BACKWARD PASS (Ritorno)
    project_duration = max([data['EF'] for data in cpm_nodes.values()], default=0)
    for n_id, data in cpm_nodes.items():
        data['LF'] = project_duration
        data['LS'] = data['LF'] - data['durata']
        
    changed = True
    while changed:
        changed = False
        for n_id, data in cpm_nodes.items():
            min_ls = data['LF'] 
            if len(data['succs']) > 0:
                min_ls = min([cpm_nodes[s_id]['LS'] for s_id in data['succs'] if s_id in cpm_nodes])
            new_lf = min_ls
            new_ls = new_lf - data['durata']
            if new_lf != data['LF'] or new_ls != data['LS']:
                data['LF'] = new_lf
                data['LS'] = new_ls
                changed = True
                
    # FASE 3: MARGINI E CRITICITÀ
    for n_id, data in cpm_nodes.items():
        data['slack'] = data['LS'] - data['ES']
        if data['slack'] <= 0:
            data['is_critical'] = True
            
    return cpm_nodes

# --- 3. ESECUZIONE CALCOLI INIZIALI ---
aggiorna_costi_reali()
st.session_state.wbs_data = calcola_evm(st.session_state.wbs_data, pd.Timestamp.today().date())

# --- SIDEBAR: GESTIONE PROGETTI A SCOMPARSA ---
with st.sidebar:
    st.header("📂 Gestione Progetti")
    
    st.session_state.nome_progetto_attivo = st.text_input("Nome Progetto Attuale", value=st.session_state.nome_progetto_attivo)
    
    c_save, c_dup = st.columns(2)
    if c_save.button("💾 Salva", use_container_width=True):
        st.session_state.archivio_progetti[st.session_state.nome_progetto_attivo] = {
            "wbs": st.session_state.wbs_data.copy(),
            "obs": st.session_state.obs_data.copy(),
            "registro": st.session_state.registro_data.copy()
        }
        st.success("Progetto salvato!")
        
    if c_dup.button("📑 Duplica", use_container_width=True):
        nuovo_nome = f"{st.session_state.nome_progetto_attivo}_Copia"
        st.session_state.archivio_progetti[nuovo_nome] = {
            "wbs": st.session_state.wbs_data.copy(),
            "obs": st.session_state.obs_data.copy(),
            "registro": st.session_state.registro_data.copy()
        }
        st.session_state.nome_progetto_attivo = nuovo_nome
        st.success("Progetto duplicato!")
        st.rerun()

    if st.session_state.archivio_progetti:
        st.divider()
        st.write("🔄 **Progetti in memoria (Sessione attuale)**")
        prog_selezionato = st.selectbox("Seleziona da caricare", options=list(st.session_state.archivio_progetti.keys()), label_visibility="collapsed")
        if st.button("📂 Apri Progetto", use_container_width=True):
            st.session_state.wbs_data = st.session_state.archivio_progetti[prog_selezionato]["wbs"].copy()
            st.session_state.obs_data = st.session_state.archivio_progetti[prog_selezionato]["obs"].copy()
            st.session_state.registro_data = st.session_state.archivio_progetti[prog_selezionato]["registro"].copy()
            st.session_state.nome_progetto_attivo = prog_selezionato
            st.rerun()

    st.divider()
    
    if st.button("📄 Nuovo Progetto (Reset Dati)", use_container_width=True):
        for key in ['wbs_data', 'obs_data', 'registro_data']:
            if key in st.session_state:
                del st.session_state[key]
        st.session_state.nome_progetto_attivo = "Nuovo_Progetto"
        st.rerun()
        
    st.divider()
    
    st.write("💾 **Archiviazione su PC**")
    try:
        progetto_export = {
            "wbs": json.loads(st.session_state.wbs_data.to_json(orient="records", date_format="iso")),
            "obs": json.loads(st.session_state.obs_data.to_json(orient="records")),
            "registro": json.loads(st.session_state.registro_data.to_json(orient="records", date_format="iso"))
        }
        json_string = json.dumps(progetto_export, indent=4)
        
        st.download_button(
            label="⬇️ Scarica (.json)",
            data=json_string,
            file_name=f"{st.session_state.nome_progetto_attivo}.json",
            mime="application/json",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Errore esportazione: {e}")
    
    uploaded_file = st.file_uploader("📤 Carica da PC", type=['json'], label_visibility="collapsed")
    
    if uploaded_file is not None:
        try:
            dati_caricati = json.load(uploaded_file)
            st.session_state.wbs_data = pd.DataFrame(dati_caricati['wbs'])
            st.session_state.obs_data = pd.DataFrame(dati_caricati['obs'])
            if 'registro' in dati_caricati:
                st.session_state.registro_data = pd.DataFrame(dati_caricati['registro'])
            
            colonne_date_wbs = ['Inizio_Previsto', 'Fine_Prevista', 'Inizio_Effettivo', 'Fine_Effettiva']
            for col in colonne_date_wbs:
                if col in st.session_state.wbs_data.columns:
                    st.session_state.wbs_data[col] = pd.to_datetime(st.session_state.wbs_data[col]).dt.date
                    
            if 'registro_data' in st.session_state and 'Data' in st.session_state.registro_data.columns:
                st.session_state.registro_data['Data'] = pd.to_datetime(st.session_state.registro_data['Data']).dt.date
            
            st.session_state.nome_progetto_attivo = uploaded_file.name.replace(".json", "")
            st.success("Dati ripristinati!")
            
        except Exception as e:
            st.error(f"File JSON non valido. {e}")
            
    st.divider()
    
    if st.button("🚪 Esci (Logout)", type="primary", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()


# --- CREAZIONE TAB ---
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🗂️ WBS (Lavorazioni)", 
    "👥 OBS (Risorse)", 
    "🕸️ Nodi & Matrice", 
    "📅 Cronoprogramma", 
    "📈 EVM & Cash Flow",
    "🧾 Reg. Contabile",
    "🛠️ Direzione & CAPA"
])

# --- TAB 1: SETUP WBS (Solo Lavorazioni) ---
with tab1:
    st.header("WBS - Work Breakdown Structure")
    
    df = st.session_state.wbs_data
    df['Durata_Prevista (gg)'] = (pd.to_datetime(df['Fine_Prevista']) - pd.to_datetime(df['Inizio_Previsto'])).dt.days
    
    is_root = ~df['ID_WBS'].astype(str).str.contains('\.')
    radici = df[is_root]
    
    df_aggiornato = pd.DataFrame()
    
    for _, radice in radici.iterrows():
        id_radice = str(radice['ID_WBS'])
        discendenti = df[df['ID_WBS'].astype(str).str.startswith(f"{id_radice}.")]
        tot_budget = discendenti['BAC_Budget'].sum()
        
        with st.expander(f"📁 {id_radice} - {radice['Attività']} (Budget Raggruppato: € {tot_budget:,.2f})", expanded=True):
            
            discendenti_modificati = st.data_editor(
                discendenti,
                key=f"editor_wbs_{id_radice}",
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                disabled=["Durata_Prevista (gg)", "ID_WBS", "AC_Costo_Reale"],
                column_config={
                    "Predecessori": st.column_config.TextColumn(
                        "Predecessori (WP)",
                        help="ID dei WP che devono finire prima (es. 1.1, 1.2)"
                    ),
                    "Inizio_Previsto": st.column_config.DateColumn("Inizio Previsto"),
                    "Fine_Prevista": st.column_config.DateColumn("Fine Prevista"),
                    "Inizio_Effettivo": st.column_config.DateColumn("Inizio Effettivo"),
                    "Fine_Effettiva": st.column_config.DateColumn("Fine Effettiva")
                }
            )
            
            radice_aggiornata = radice.copy()
            radice_aggiornata['BAC_Budget'] = discendenti_modificati['BAC_Budget'].sum()
            radice_aggiornata['AC_Costo_Reale'] = discendenti_modificati['AC_Costo_Reale'].sum()
            
            df_aggiornato = pd.concat([df_aggiornato, pd.DataFrame([radice_aggiornata]), discendenti_modificati], ignore_index=True)
            
    with st.form("aggiungi_padre"):
        st.write("Aggiungi nuova Macro-Categoria")
        c1, c2, c3 = st.columns([2, 5, 2])
        nuovo_id = c1.text_input("ID (es. 5)")
        nuova_att = c2.text_input("Nome Categoria")
        if c3.form_submit_button("➕ Aggiungi"):
            if nuovo_id and nuova_att:
                nuova_riga = pd.DataFrame([{
                    'ID_WBS': nuovo_id, 'Attività': nuova_att, 'BAC_Budget': 0.0, 
                    '%_Completamento': 0, 'AC_Costo_Reale': 0.0
                }])
                st.session_state.wbs_data = pd.concat([st.session_state.wbs_data, nuova_riga], ignore_index=True)
                st.rerun()

    if not df_aggiornato.empty:
        st.session_state.wbs_data = df_aggiornato

# --- TAB 2: SETUP OBS (Solo Risorse) ---
with tab2:
    st.header("OBS - Organization Breakdown Structure")
    
    with st.expander("⚙️ Gestione Colonne Aggiuntive", expanded=False):
        c1, c2 = st.columns([3, 1])
        nuova_col = c1.text_input("Nome nuova colonna (es. Telefono, Qualifica, Partita IVA)")
        if c2.button("➕ Crea") and nuova_col:
            if nuova_col not in st.session_state.obs_data.columns:
                st.session_state.obs_data[nuova_col] = "" 
                st.rerun()
        
        colonne_base = ['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note']
        colonne_custom = [c for c in st.session_state.obs_data.columns if c not in colonne_base]
        
        if colonne_custom:
            st.divider()
            c3, c4, c5 = st.columns([2, 2, 1])
            col_da_modificare = c3.selectbox("Colonna da rinominare", options=colonne_custom)
            nuovo_nome = c4.text_input("Nuovo nome intestazione")
            if c5.button("✏️ Modifica") and nuovo_nome:
                st.session_state.obs_data.rename(columns={col_da_modificare: nuovo_nome}, inplace=True)
                st.rerun()

    st.session_state.obs_data = st.data_editor(
        st.session_state.obs_data, 
        column_config={
            "Tipo_Contratto": st.column_config.SelectboxColumn(
                "Tipo Contratto",
                options=["Appalto ▾", "Sub appalto ▾"],
                required=True
            )
        },
        num_rows="dynamic", 
        use_container_width=True, 
        hide_index=True
    )
        
# --- TAB 3: MATRICE E GRAFO A NODI ---
with tab3:
    st.header("Incrocio Logico (Work Packages e Percorso Critico)")
    
    # Attiviamo l'algoritmo CPM in background
    cpm_data = calcola_cpm(st.session_state.wbs_data)
    
    mostra_relazioni = st.toggle("👁️ Mostra Relazioni tra WP (Interferenze)", value=True)
    
    graph = graphviz.Digraph(engine='dot')
    graph.attr(rankdir='LR', ranksep='1.5', nodesep='0.8', splines='spline')
    graph.attr('node', fontname='Helvetica', fontsize='10', margin='0.2')
    
    # --- NODI OBS ---
    for _, row in st.session_state.obs_data.iterrows():
        label_html = f"<<TABLE BORDER='0' CELLBORDER='0' CELLSPACING='2'>"
        label_html += f"<TR><TD><B>{row['Ruolo']}</B></TD></TR>"
        label_html += f"<TR><TD>({row['Risorsa']})</TD></TR>"
        
        colonne_base = ['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note']
        colonne_custom = [col for col in st.session_state.obs_data.columns if col not in colonne_base]
        
        for col in colonne_custom:
            valore = row[col]
            if pd.notna(valore) and str(valore).strip() != "":
                label_html += f"<TR><TD><FONT POINT-SIZE='9' COLOR='gray30'>{col}: {valore}</FONT></TD></TR>"
        label_html += "</TABLE>>"
        
        graph.node(
            f"OBS_{row['ID_OBS']}",  
            label=label_html, 
            shape='rect', 
            style='rounded,filled', 
            fillcolor='#E1F5FE', 
            color='#0288D1',     
            penwidth='1.5'
        )
        
    # --- NODI WBS E PERCORSO CRITICO ---
    df_wp_reali = st.session_state.wbs_data[st.session_state.wbs_data['ID_WBS'].astype(str).str.contains('\.')]
    valid_wbs_ids = set(df_wp_reali['ID_WBS'].astype(str))
    
    for _, row in df_wp_reali.iterrows():
        attivita = str(row['Attività'])
        budget = float(row['BAC_Budget'])
        costo_reale = float(row['AC_Costo_Reale'])
        completamento = float(row['%_Completamento'])
        
        # Recuperiamo i dati CPM per questo nodo specifico
        wp_cpm = cpm_data.get(str(row['ID_WBS']).strip(), {})
        margine = wp_cpm.get('slack', 0)
        is_critical = wp_cpm.get('is_critical', False)
        
        inizio_str = row['Inizio_Previsto'].strftime('%d/%m/%Y') if pd.notna(row['Inizio_Previsto']) else "N/D"
        fine_str = row['Fine_Prevista'].strftime('%d/%m/%Y') if pd.notna(row['Fine_Prevista']) else "N/D"
        
        # HTML arricchito: Margine in rosso se critico
        testo_margine = f"<FONT COLOR='#D32F2F'><B>Margine: {margine} gg</B></FONT>" if is_critical else f"<FONT COLOR='#388E3C'>Margine: {margine} gg</FONT>"
        
        wp_html = f"<<TABLE BORDER='0' CELLBORDER='0' CELLSPACING='4'>"
        wp_html += f"<TR><TD COLSPAN='2'><B>{row['ID_WBS']} - {attivita}</B></TD></TR>"
        wp_html += f"<TR><TD ALIGN='LEFT'>Inizio: {inizio_str}</TD><TD ALIGN='RIGHT'>Fine: {fine_str}</TD></TR>"
        wp_html += f"<TR><TD ALIGN='LEFT'>Budget: &euro; {budget:,.2f}</TD><TD ALIGN='RIGHT'>AC: &euro; {costo_reale:,.2f}</TD></TR>"
        wp_html += f"<TR><TD ALIGN='LEFT'>Avanzamento: {completamento:.1f}%</TD><TD ALIGN='RIGHT'>{testo_margine}</TD></TR>"
        wp_html += "</TABLE>>"
        
        # Gestione Stile (Barra di progresso)
        if completamento >= 100:
            stile = 'rounded,filled'
            colore_sfondo = '#C8E6C9' 
        elif completamento <= 0:
            stile = 'rounded,filled'
            colore_sfondo = 'white'   
        else:
            stile = 'rounded,striped'
            quota_verde = completamento / 100.0
            colore_sfondo = f"#C8E6C9;{quota_verde}:white"
            
        # --- APPLICAZIONE STILE PERCORSO CRITICO ---
        bordo_colore = '#D32F2F' if is_critical else '#388E3C'  # Rosso se critico, altrimenti verde scuro
        spessore_bordo = '3.0' if is_critical else '1.5'        # Più spesso se critico
        
        graph.node(
            f"WBS_{row['ID_WBS']}", 
            label=wp_html, 
            shape='rect', 
            style=stile, 
            fillcolor=colore_sfondo, 
            color=bordo_colore,     
            penwidth=spessore_bordo
        )
        
        # Assegnazioni OBS (Linee grigie)
        if pd.notna(row['ID_OBS_Assegnato']):
            obs_ids = str(row['ID_OBS_Assegnato']).split(',')
            for o_id in obs_ids:
                if o_id.strip():
                    graph.edge(f"OBS_{o_id.strip()}", f"WBS_{row['ID_WBS']}", color='#757575', penwidth='1.5', arrowsize='0.8')
                    
        # Connessioni WP (Il Fiume Logico)
        if mostra_relazioni and 'Predecessori' in row and pd.notna(row['Predecessori']):
            preds = str(row['Predecessori']).split(',')
            for p_id in preds:
                p_id = p_id.strip()
                if p_id in valid_wbs_ids:
                    pred_is_critical = cpm_data.get(p_id, {}).get('is_critical', False)
                    
                    # Se ENTRAMBI i nodi sono sul percorso critico, coloriamo il cavo di rosso spesso
                    if is_critical and pred_is_critical:
                        colore_cavo = '#D32F2F' # Rosso fuoco
                        stile_cavo = 'solid'
                        spessore_cavo = '2.5'
                        freccia = '1.0'
                    else:
                        colore_cavo = '#FF9800' # Arancione standard
                        stile_cavo = 'dashed'
                        spessore_cavo = '1.0'
                        freccia = '0.6'
                        
                    graph.edge(
                        f"WBS_{p_id}", 
                        f"WBS_{row['ID_WBS']}", 
                        color=colore_cavo,  
                        style=stile_cavo,   
                        penwidth=spessore_cavo,   
                        arrowsize=freccia
                    )

    # Rendering interattivo
    try:
        raw_svg = graph.pipe(format='svg').decode('utf-8')
        svg_data = raw_svg[raw_svg.find('<svg'):]
        
        html_code = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <script src="https://cdn.jsdelivr.net/npm/svg-pan-zoom@3.6.1/dist/svg-pan-zoom.min.js"></script>
            <style>
                body {{ margin: 0; padding: 0; overflow: hidden; background-color: #fafafa; }}
                #svg-container {{ width: 100vw; height: 100vh; display: flex; justify-content: center; align-items: center; }}
                svg {{ width: 100% !important; height: 100% !important; }}
            </style>
        </head>
        <body>
            <div id="svg-container">
                {svg_data}
            </div>
            <script>
                window.onload = function() {{
                    var svgElement = document.querySelector('svg');
                    if (svgElement) {{
                        svgElement.setAttribute('id', 'grafo-interattivo');
                        svgElement.removeAttribute('width');
                        svgElement.removeAttribute('height');
                        var panZoom = svgPanZoom('#grafo-interattivo', {{
                            zoomEnabled: true,
                            controlIconsEnabled: true,
                            fit: true,
                            center: true,
                            minZoom: 0.1,
                            maxZoom: 10,
                            mouseWheelZoomEnabled: true
                        }});
                    }} else {{
                        document.getElementById('svg-container').innerHTML = "Errore grafico SVG.";
                    }}
                }};
            </script>
        </body>
        </html>
        """
        components.html(html_code, height=600)
    except Exception as e:
        st.error(f"Errore nella generazione del grafo: {e}")
        st.graphviz_chart(graph)

    # --- NUOVA LEGENDA DEL GRAFO ---
    st.divider()
    st.subheader("📖 Legenda del Grafo")
    
    col_leg1, col_leg2 = st.columns(2)
    
    with col_leg1:
        st.markdown("""
        **NODI E FIGURE**
        * 🟦 **Riquadro Azzurro:** Risorsa/Ruolo (OBS) assegnato al cantiere.
        * 🟩 **Riquadro Verde:** Work Package (WBS). Il riempimento interno funge da barra di caricamento e indica la **% di avanzamento** reale.
        * 🟥 **Bordo Rosso Spesso:** Attività sul **Percorso Critico** (Margine = 0 gg). Attenzione: un ritardo in questo blocco ritarderà la fine dell'intero progetto!
        """)
        
    with col_leg2:
        st.markdown("""
        **CAVI E COLLEGAMENTI**
        * 🔗 **Freccia Grigia Continua:** Indica quale Risorsa (OBS) è incaricata di eseguire quale Lavorazione (WBS).
        * 🔀 **Freccia Arancione Tratteggiata:** Relazione logica standard (es. *L'attività B inizia dopo l'attività A*).
        * 🚨 **Freccia Rossa Spessa:** Il flusso del **Percorso Critico**. Segue esattamente la catena logica di attività che determina la durata totale del cantiere.
        """)

# --- TAB 4: CRONOPROGRAMMA (GANTT) ---
with tab4:
    st.header("Cronoprogramma Lavori")
    
    c1, c2 = st.columns([1, 2])
    vista = c1.selectbox("Seleziona Vista", ["Progetto (Baseline)", "Esecuzione (Esecutivo)", "Comparativa"])
    
    data_status_gantt = c2.date_input("📅 Data di Rilevamento (Simulazione avanzamento cantiere)", value=date(2026, 10, 15))
    
    df_gantt = st.session_state.wbs_data.copy()
    df_gantt = df_gantt[df_gantt['ID_WBS'].astype(str).str.contains('\.')] 
    
    df_gantt['Inizio_Previsto'] = pd.to_datetime(df_gantt['Inizio_Previsto'])
    df_gantt['Fine_Prevista'] = pd.to_datetime(df_gantt['Fine_Prevista'])
    df_gantt['Inizio_Effettivo'] = pd.to_datetime(df_gantt['Inizio_Effettivo'])
    
    df_gantt['Fine_Effettiva'] = pd.to_datetime(df_gantt['Fine_Effettiva']).fillna(pd.to_datetime(data_status_gantt))
    
    fig = go.Figure()
    
    if vista in ["Progetto (Baseline)", "Comparativa"]:
        durata_prevista_ms = (df_gantt['Fine_Prevista'] - df_gantt['Inizio_Previsto']).dt.total_seconds() * 1000
        
        fig.add_trace(go.Bar(
            x=durata_prevista_ms,
            y=df_gantt['Attività'],
            base=df_gantt['Inizio_Previsto'],
            orientation='h',
            name='Baseline',
            width=0.4, 
            marker=dict(color='rgba(0, 0, 255, 0.4)') if vista == "Comparativa" else dict(color='blue')
        ))
        
    if vista in ["Esecuzione (Esecutivo)", "Comparativa"]:
        df_esec = df_gantt.dropna(subset=['Inizio_Effettivo']).copy()
        durata_effettiva_ms = (df_esec['Fine_Effettiva'] - df_esec['Inizio_Effettivo']).dt.total_seconds() * 1000
        
        fig.add_trace(go.Bar(
            x=durata_effettiva_ms,
            y=df_esec['Attività'],
            base=df_esec['Inizio_Effettivo'],
            orientation='h',
            name='Esecutivo',
            width=0.2, 
            marker=dict(color='red')
        ))
        
    fig.update_layout(
        barmode='overlay', 
        height=600, 
        bargap=0.3, 
        xaxis_title="Linea Temporale", 
        yaxis_title="WBS", 
        yaxis={'autorange': 'reversed'},
        xaxis_type='date' 
    )
    st.plotly_chart(fig, use_container_width=True)

# --- TAB 5: EVM E CASH FLOW ---
with tab5:
    st.header("Controllo Costi e Analisi EVM")
    
    data_status_evm = st.date_input("📅 Data di Stato (Status Date) per l'analisi EVM:", value=date(2026, 10, 15))
    
    df_completo = st.session_state.wbs_data.copy()
    df_evm = df_completo[df_completo['ID_WBS'].astype(str).str.contains('\.')].copy()
    
    df_evm = calcola_evm(df_evm, data_status_evm)
    
    tot_bac = df_evm['BAC_Budget'].sum()
    tot_pv = df_evm['PV'].sum()
    tot_ev = df_evm['EV'].sum()
    tot_ac = df_evm['AC_Costo_Reale'].sum()
    
    tot_eac = df_evm['EAC'].sum()
    tot_etc = df_evm['ETC'].sum()
    tot_vac = df_evm['VAC'].sum()
    
    cpi_globale = tot_ev / tot_ac if tot_ac > 0 else 1.0
    spi_globale = tot_ev / tot_pv if tot_pv > 0 else 1.0
    perc_completamento = (tot_ev / tot_bac * 100) if tot_bac > 0 else 0.0
    perc_pianificata = (tot_pv / tot_bac * 100) if tot_bac > 0 else 0.0
    
    st.markdown("### 📊 Stato Attuale (Consuntivo)")
    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
    col_m1.metric("Budget Totale (BAC)", f"€ {tot_bac:,.0f}")
    col_m2.metric("Lavoro Eseguito (EV)", f"€ {tot_ev:,.0f}")
    col_m3.metric("Costi Sostenuti (AC)", f"€ {tot_ac:,.0f}")
    col_m4.metric("Avanzamento Globale", f"{perc_completamento:.1f}%", delta=f"Pianificato: {perc_pianificata:.1f}%", delta_color="off")
    col_m5.metric("SPI (Tempi)", f"{spi_globale:.2f}", delta="In ritardo" if spi_globale < 1 else "In anticipo", delta_color="inverse")
    
    st.markdown("### 🔮 Previsioni a Finire (Proiezioni)")
    col_p1, col_p2, col_p3, col_p4 = st.columns(4)
    col_p1.metric("Costo Finale Stimato (EAC)", f"€ {tot_eac:,.0f}", delta="Proiezione a fine lavori", delta_color="off")
    col_p2.metric("Costo Residuo (ETC)", f"€ {tot_etc:,.0f}", delta="Capitale ancora necessario", delta_color="off")
    col_p3.metric("Varianza a Finire (VAC)", f"€ {tot_vac:,.0f}", delta="Perdita Stimata" if tot_vac < 0 else "Risparmio Stimato", delta_color="normal")
    col_p4.metric("CPI (Costi)", f"{cpi_globale:.2f}", delta="Over-budget" if cpi_globale < 1 else "Under-budget", delta_color="inverse")
    
    st.divider()

    # --- GRAFICO EVM: CURVA AD S (S-CURVE) ---
    st.subheader("📈 Curva ad S (Andamento Temporale di Progetto)")
    
    df_scurve = genera_dati_scurve(df_evm, st.session_state.registro_data, data_status_evm)
    
    if df_scurve is not None and not df_scurve.empty:
        fig_scurve = px.line(
            df_scurve, 
            x='Data', 
            y=['PV (Valore Pianificato)', 'EV (Valore Guadagnato)', 'AC (Costo Reale)'],
            color_discrete_map={
                'PV (Valore Pianificato)': 'blue',
                'EV (Valore Guadagnato)': 'green',
                'AC (Costo Reale)': 'red'
            },
            labels={'value': 'Importo (€)', 'variable': 'Metrica EVM'}
        )
        
        # --- NOVITÀ: AGGIUNTA PROIEZIONI FUTURE (FORECAST) ---
        df_past = df_scurve[df_scurve['Data'] <= data_status_evm]
        
        if not df_past.empty:
            last_ac = df_past.iloc[-1]['AC (Costo Reale)']
            last_ev = df_past.iloc[-1]['EV (Valore Guadagnato)']
            last_pv = df_past.iloc[-1]['PV (Valore Pianificato)']
            
            min_date = df_scurve['Data'].min()
            max_date = df_scurve['Data'].max()
            giorni_pianificati = (max_date - min_date).days
            
            spi_effettivo = last_ev / last_pv if last_pv > 0 else 1.0
            
            if spi_effettivo > 0:
                giorni_stimati = int(giorni_pianificati / spi_effettivo)
            else:
                giorni_stimati = giorni_pianificati
                
            giorni_stimati = min(giorni_stimati, giorni_pianificati * 3) 
            data_fine_stimata = min_date + pd.Timedelta(days=giorni_stimati)
            
            fig_scurve.add_trace(go.Scatter(
                x=[data_status_evm, data_fine_stimata],
                y=[last_ac, tot_eac],
                mode='lines',
                line=dict(color='red', dash='dot', width=2),
                name='Proiezione Costi (verso EAC)'
            ))
            
            fig_scurve.add_trace(go.Scatter(
                x=[data_status_evm, data_fine_stimata],
                y=[last_ev, tot_bac],
                mode='lines',
                line=dict(color='green', dash='dot', width=2),
                name='Proiezione Lavoro (verso BAC)'
            ))
            
            fig_scurve.update_xaxes(range=[min_date, max(max_date, data_fine_stimata) + pd.Timedelta(days=5)])

        fig_scurve.update_layout(
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=50, b=20)
        )
        
        fig_scurve.add_vline(x=str(data_status_evm), line_width=2, line_dash="dash", line_color="gray", annotation_text="Data di Rilevamento")
        
        st.plotly_chart(fig_scurve, use_container_width=True)
    else:
        st.info("ℹ️ Non ci sono ancora date di pianificazione sufficienti per generare la Curva ad S.")
    
    st.divider()
    
    # --- GRAFICO A BARRE ---
    st.subheader("Raffronto Costi per Attività")
    fig_evm = go.Figure(data=[
        go.Bar(name='BAC (Budget)', x=df_evm['Attività'], y=df_evm['BAC_Budget'], marker_color='lightgray', text=df_evm['BAC_Budget'], texttemplate='€ %{text:,.0f}', textposition='outside', textangle=-90),
        go.Bar(name='EV (Valore Guadagnato)', x=df_evm['Attività'], y=df_evm['EV'], marker_color='green', text=df_evm['EV'], texttemplate='€ %{text:,.0f}', textposition='outside', textangle=-90),
        go.Bar(name='AC (Costo Reale)', x=df_evm['Attività'], y=df_evm['AC_Costo_Reale'], marker_color='red', text=df_evm['AC_Costo_Reale'], texttemplate='€ %{text:,.0f}', textposition='outside', textangle=-90)
    ])
    fig_evm.update_layout(barmode='group', margin=dict(t=80), uniformtext_minsize=9, uniformtext_mode='hide')
    st.plotly_chart(fig_evm, use_container_width=True)
        
    # --- TABELLA E LEGENDA ---
    col_KPI, col_LEGENDA = st.columns([7, 3]) 
    with col_KPI:
        st.subheader("Indicatori di Performance (KPI)")
        df_kpi = df_evm[['Attività', '%_Completamento', 'CPI', 'SPI', 'CV']].copy()
        
        def color_kpi(val):
            if isinstance(val, (int, float)):
                if val < 0.95: return 'color: red; font-weight: bold;'
                elif val >= 1.0: return 'color: green'
            return ''
            
        st.dataframe(df_kpi.style.map(color_kpi, subset=['CPI', 'SPI']).format({'CPI': "{:.2f}", 'SPI': "{:.2f}", 'CV': "€ {:.2f}"}), use_container_width=True)

    with col_LEGENDA:
        st.subheader("Legenda EVM")
        st.markdown("""
        * **CPI:** Efficienza costi (<1 sforamento budget)
        * **SPI:** Efficienza tempi (<1 in ritardo)
        * **CV:** Varianza Costi Assoluta
        """)
        
    st.divider()
    
    # --- MOTORE AI ANALIZZATORE DIREZIONALE ---
    st.subheader("🤖 Analizzatore Direzionale (AI-Assist)")
    
    soglia_allerta = 0.95
    critici_costo = df_evm[df_evm['CPI'] < soglia_allerta]
    critici_tempo = df_evm[df_evm['SPI'] < soglia_allerta]
    
    if critici_costo.empty and critici_tempo.empty:
        st.success("✅ **Progetto in Salute:** Tutti i parametri (Tempi e Costi) sono entro i margini di tolleranza pianificati. Nessuna criticità rilevata.")
    else:
        st.warning("⚠️ **Attenzione: Rilevati scostamenti rispetto alla baseline di progetto.** Analisi suggerita:")
        
        for _, row in critici_tempo.iterrows():
            st.error(f"⏳ **Ritardo Schedulazione su '{row['Attività']}':** (SPI = {row['SPI']:.2f})")
            st.markdown(f"> *Il Work Package sta generando meno valore del previsto. Dato lo scostamento, **devi accelerare la produzione**.*")
            st.markdown(f"> * **Soluzioni suggerite:** Verifica la disponibilità della risorsa ({row['ID_OBS_Assegnato']}), valuta di approvare lavoro straordinario o affianca un sub-appaltatore per recuperare il gap prima che intacchi il percorso critico (CPM).*")
            
        for _, row in critici_costo.iterrows():
            st.error(f"💸 **Sforamento Budget su '{row['Attività']}':** (CPI = {row['CPI']:.2f})")
            st.markdown(f"> *Hai speso **€ {row['AC_Costo_Reale']:,.2f}** per produrre un valore equivalente di soli **€ {row['EV']:,.2f}**. Stai perdendo marginalità.*")
            st.markdown(f"> * **Soluzioni suggerite:** Analizza immediatamente le bolle di accompagnamento e il Registro Contabile. Possibili cause: inefficienza della manodopera, aumento prezzi materiali imprevisto, o errata valutazione del BAC iniziale.*")
        
# --- TAB 6: REGISTRO CONTABILE ---
with tab6:
    st.header("Registro Contabile")
    st.markdown("Inserisci qui le fatture e i SAL. Gli importi netti si sommeranno automaticamente aggiornando la voce *AC_Costo_Reale* nella WBS.")
    
    df_wbs = st.session_state.wbs_data
    leaf_wbs = df_wbs[df_wbs['ID_WBS'].astype(str).str.contains('\.')]
    wbs_options = [f"{row['ID_WBS']} - {row['Attività']}" for _, row in leaf_wbs.iterrows()]
    
    edited_registro = st.data_editor(
        st.session_state.registro_data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Data": st.column_config.DateColumn("Data Registrazione"),
            "N_Doc": st.column_config.TextColumn("N° Doc/Fattura"),
            "Fornitore": st.column_config.TextColumn("Fornitore (OBS)"),
            "Descrizione": st.column_config.TextColumn("Descrizione / Note"),
            "Importo_Netto": st.column_config.NumberColumn(
                "Importo Netto (€)", 
                format="€ %.2f", 
                min_value=0.0
            ),
            "Voce_WBS": st.column_config.SelectboxColumn(
                "Attività WBS (Destinazione) ▾",
                help="Seleziona la lavorazione di riferimento",
                options=wbs_options, 
                required=True
            )
        }
    )
    
    if not edited_registro.equals(st.session_state.registro_data):
        st.session_state.registro_data = edited_registro
        st.rerun()

# --- TAB 7: DIREZIONE LAVORI, CAPA & REPORTISTICA ---
with tab7:
    st.header("Direzione Lavori: Interventi (CAPA) e Simulazioni")
    
    # --- PREPARAZIONE DATI DINAMICI ---
    df_wbs_capa = st.session_state.wbs_data
    leaf_wbs_capa = df_wbs_capa[df_wbs_capa['ID_WBS'].astype(str).str.contains('\.')]
    wbs_options_capa = [f"{row['ID_WBS']} - {row['Attività']}" for _, row in leaf_wbs_capa.iterrows()]
    
    df_obs_capa = st.session_state.obs_data
    obs_options_capa = [f"{row['ID_OBS']} - {row['Ruolo']}" for _, row in df_obs_capa.iterrows()]
    
    # ---------------------------------------------------------
    # SEZIONE 1: REGISTRO DEGLI INTERVENTI (ACTION LOG)
    # ---------------------------------------------------------
    st.subheader("1. Registro Azioni Correttive e Preventive (CAPA)")
    
    edited_capa = st.data_editor(
        st.session_state.capa_data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Data_Apertura": st.column_config.DateColumn("Data Segnalazione"),
            "ID_WBS_Rif": st.column_config.SelectboxColumn("Attività WBS (Rif.)", options=wbs_options_capa, required=True),
            "Tipo_Azione": st.column_config.SelectboxColumn("Tipo", options=["Correttiva ▾", "Preventiva ▾"], required=True),
            "Descrizione": st.column_config.TextColumn("Descrizione Intervento / Ordine", width="large"),
            "Responsabile_OBS": st.column_config.SelectboxColumn("Assegnato a (OBS)", options=obs_options_capa, required=True),
            "Stato": st.column_config.SelectboxColumn("Stato", options=["Aperto ▾", "In Lavorazione ▾", "Chiuso ▾"], required=True)
        }
    )
    if not edited_capa.equals(st.session_state.capa_data):
        st.session_state.capa_data = edited_capa
        st.rerun()

    st.divider()
    
    # ---------------------------------------------------------
    # SEZIONE 2: SIMULATORE WHAT-IF
    # ---------------------------------------------------------
    with st.expander("🔬 2. Ambiente di Simulazione (What-If Analysis)"):
        st.markdown("Simula l'impatto economico di un'azione correttiva sul Costo Finale Stimato (EAC) prima di approvarla.")
        
        c_sim1, c_sim2 = st.columns(2)
        wp_scelto = c_sim1.selectbox("Seleziona Work Package da simulare", options=wbs_options_capa)
        extra_costo = c_sim2.number_input("Iniezione Extra Costo per risolvere l'anomalia (€)", value=0.0, step=500.0)
        
        if wp_scelto:
            wp_id = wp_scelto.split(' - ')[0]
            
            # DataFrame temporaneo per simulazione
            df_simulazione = st.session_state.wbs_data.copy()
            df_simulazione['BAC_Budget'] = pd.to_numeric(df_simulazione['BAC_Budget'], errors='coerce').fillna(0.0)
            df_simulazione['%_Completamento'] = pd.to_numeric(df_simulazione['%_Completamento'], errors='coerce').fillna(0.0)
            df_simulazione['AC_Costo_Reale'] = pd.to_numeric(df_simulazione['AC_Costo_Reale'], errors='coerce').fillna(0.0)
            
            # Applichiamo la simulazione all'AC
            indice_riga = df_simulazione.index[df_simulazione['ID_WBS'] == wp_id].tolist()
            if indice_riga:
                idx = indice_riga[0]
                df_simulazione.at[idx, 'AC_Costo_Reale'] += extra_costo
                
            # Ricalcoli
            df_sim_calc = calcola_evm(df_simulazione[df_simulazione['ID_WBS'].astype(str).str.contains('\.')].copy(), pd.Timestamp.today().date())
            df_reale_calc = calcola_evm(st.session_state.wbs_data[st.session_state.wbs_data['ID_WBS'].astype(str).str.contains('\.')].copy(), pd.Timestamp.today().date())
            
            eac_attuale = df_reale_calc['EAC'].sum()
            eac_simulato = df_sim_calc['EAC'].sum()
            ac_attuale_tot = df_reale_calc['AC_Costo_Reale'].sum()
            ac_simulato_tot = df_sim_calc['AC_Costo_Reale'].sum()
            delta_eac = eac_simulato - eac_attuale
            
            c_res1, c_res2 = st.columns([1, 2])
            
            with c_res1:
                st.metric(
                    label="Nuovo Costo Finale (EAC Simulato)", 
                    value=f"€ {eac_simulato:,.2f}", 
                    delta=f"Variazione: € {delta_eac:,.2f}" if delta_eac != 0 else "Nessun impatto", 
                    delta_color="inverse"
                )
                st.markdown(f"*Costo attuale sostenuto: € {ac_attuale_tot:,.2f}*")
                st.markdown(f"*Costo istantaneo simulato: € {ac_simulato_tot:,.2f}*")
            
            with c_res2:
                # Mini grafico vettoriale: Forchetta di proiezione tra Attuale e Simulato
                fig_sim = go.Figure()
                
                # Proiezione Attuale (Rossa)
                fig_sim.add_trace(go.Scatter(
                    x=["Oggi", "Fine Lavori"],
                    y=[ac_attuale_tot, eac_attuale],
                    mode='lines+markers+text',
                    name='Traiettoria Attuale',
                    line=dict(color='red', dash='dash', width=3),
                    text=[f"€ {ac_attuale_tot:,.0f}", f"€ {eac_attuale:,.0f}"],
                    textposition="bottom right"
                ))
                
                # Proiezione Simulata (Blu)
                fig_sim.add_trace(go.Scatter(
                    x=["Oggi", "Fine Lavori"],
                    y=[ac_simulato_tot, eac_simulato],
                    mode='lines+markers+text',
                    name='Traiettoria Simulata',
                    line=dict(color='blue', dash='solid', width=3),
                    text=[f"€ {ac_simulato_tot:,.0f}", f"€ {eac_simulato:,.0f}"],
                    textposition="top left"
                ))
                
                fig_sim.update_layout(
                    title="Forchetta di Variazione Costi a Finire",
                    height=250,
                    margin=dict(l=10, r=10, t=30, b=10),
                    yaxis_title="Importo (€)",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig_sim, use_container_width=True)

    st.divider()

    # ---------------------------------------------------------
    # SEZIONE 3: EXPORT REPORT DIREZIONALE IN WORD (.DOCX)
    # ---------------------------------------------------------
    st.subheader("3. Stampa Verbale di Direzione Lavori")
    
    col_f1, col_f2 = st.columns([1, 2])
    filtro_stampa = col_f1.radio("Quali interventi includere nel verbale?", ["Tutti i registrati", "Solo l'ultimo inserito", "Intervallo di date"])
    
    df_stampa = st.session_state.capa_data.copy()
    df_stampa['Data_Apertura'] = pd.to_datetime(df_stampa['Data_Apertura']).dt.date
    
    if filtro_stampa == "Solo l'ultimo inserito":
        df_stampa = df_stampa.tail(1)
    elif filtro_stampa == "Intervallo di date":
        d_start = col_f2.date_input("Da data:", value=pd.Timestamp.today().date())
        d_end = col_f2.date_input("A data:", value=pd.Timestamp.today().date())
        df_stampa = df_stampa[(df_stampa['Data_Apertura'] >= d_start) & (df_stampa['Data_Apertura'] <= d_end)]

    if st.button("📄 Genera Verbale WORD (.docx)", use_container_width=True, type="primary"):
        
        # Recupero i totali EVM correnti per il report
        df_evm_rep = calcola_evm(st.session_state.wbs_data[st.session_state.wbs_data['ID_WBS'].astype(str).str.contains('\.')].copy(), pd.Timestamp.today().date())
        tot_ev_rep = df_evm_rep['EV'].sum()
        tot_ac_rep = df_evm_rep['AC_Costo_Reale'].sum()
        tot_pv_rep = df_evm_rep['PV'].sum()
        cpi_rep = tot_ev_rep / tot_ac_rep if tot_ac_rep > 0 else 1.0
        spi_rep = tot_ev_rep / tot_pv_rep if tot_pv_rep > 0 else 1.0
        eac_rep = df_evm_rep['EAC'].sum()
        
        # Creazione del Documento Word
        doc = Document()
        
        # Intestazione Documento
        doc.add_heading('VERBALE DI DIREZIONE LAVORI', 0)
        doc.add_paragraph(f"Progetto: {st.session_state.nome_progetto_attivo}")
        doc.add_paragraph(f"Data emissione verbale: {pd.Timestamp.today().strftime('%d/%m/%Y')}")
        
        # Sezione EVM
        doc.add_heading('1. Stato Avanzamento Lavori (EVM)', level=1)
        p = doc.add_paragraph()
        p.add_run(f"CPI (Efficienza Costi): {cpi_rep:.2f}\n").bold = True
        p.add_run(f"SPI (Efficienza Tempi): {spi_rep:.2f}\n").bold = True
        p.add_run(f"Costo Finale Stimato (EAC): € {eac_rep:,.2f}").bold = True
        doc.add_paragraph("Nota: Un indicatore inferiore a 1.00 indica un superamento del budget o un ritardo sui tempi.")
        
        # Sezione Interventi (Tabella)
        doc.add_heading('2. Disposizioni e Azioni (CAPA)', level=1)
        
        if not df_stampa.empty:
            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Data'
            hdr_cells[1].text = 'Attività WBS'
            hdr_cells[2].text = 'Tipo'
            hdr_cells[3].text = 'Descrizione Intervento'
            hdr_cells[4].text = 'Stato'
            
            for _, row in df_stampa.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(row['Data_Apertura'])
                row_cells[1].text = str(row['ID_WBS_Rif'])
                row_cells[2].text = str(row['Tipo_Azione'])
                row_cells[3].text = str(row['Descrizione']) + f"\n(Assegnato: {row['Responsabile_OBS']})"
                row_cells[4].text = str(row['Stato'])
        else:
            doc.add_paragraph("Nessun intervento registrato nel periodo selezionato.")
            
        doc.add_paragraph("\n\nFirma Direzione Lavori\n_________________________")
        
        # Salvataggio nel buffer in RAM
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # Bottone di Download file Word
        st.download_button(
            label="⬇️ Clicca qui per scaricare il file Word pronto per la firma",
            data=buffer,
            file_name=f"Verbale_{st.session_state.nome_progetto_attivo}_{pd.Timestamp.today().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="secondary"
        )
