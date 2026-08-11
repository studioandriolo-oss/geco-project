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
import html
import base64

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="WBS/OBS Manager & EVM", layout="wide")

st.markdown("""
    <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 0rem;
        }
    </style>
""", unsafe_allow_html=True)

col_logo, col_title = st.columns([1, 9], vertical_alignment="center")

with col_logo:
    try:
        with open("logo_base64.txt", "r") as f:
            logo_str = f.read().strip()
        st.image(base64.b64decode(logo_str), width=120)
    except Exception:
        pass

with col_title:
    st.title("Project Workflow & EVM Controller")

# --- SISTEMA DI LOGIN SICURO ---
try:
    USER_ID = st.secrets["USER_ID"]
    PASSWORD = st.secrets["PASSWORD"]
except KeyError:
    st.error("⚠️ Errore di sistema: Credenziali non trovate. Configura i 'Secrets' di Streamlit.")
    st.stop()

if st.query_params.get("auth") == "valid":
    st.session_state.logged_in = True
elif 'logged_in' not in st.session_state:
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
                    st.query_params["auth"] = "valid" 
                    st.rerun() 
                else:
                    st.error("Credenziali errate. Riprova.")                
    st.stop()

# --- 1. INIZIALIZZAZIONE DATI E CHIAVI DI SICUREZZA ---
for key in ['editor_key_obs', 'editor_key_reg', 'editor_key_capa']:
    if key not in st.session_state:
        st.session_state[key] = 0

if 'wbs_data' not in st.session_state:
    st.session_state.wbs_data = pd.DataFrame([{
        'ID_WBS': '1', 'Attività': 'Progetto Principale', 
        'Inizio_Previsto': None, 'Fine_Prevista': None, 
        'Inizio_Effettivo': None, 'Fine_Effettiva': None, 
        'BAC_Budget': 0.0, '%_Completamento': 0.0, 
        'AC_Costo_Reale': 0.0, 'ID_OBS_Assegnato': None, 'Predecessori': ''
    }])
    
if 'obs_data' not in st.session_state:
    st.session_state.obs_data = pd.DataFrame(columns=['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note'])
    
if 'registro_data' not in st.session_state:
    st.session_state.registro_data = pd.DataFrame(columns=['Data', 'N_Doc', 'Fornitore', 'Voce_WBS', 'Importo_Netto', 'Descrizione'])

if 'capa_data' not in st.session_state:
    st.session_state.capa_data = pd.DataFrame(columns=['Data_Apertura', 'ID_WBS_Rif', 'Tipo_Azione', 'Descrizione', 'Responsabile_OBS', 'Stato'])

if 'archivio_progetti' not in st.session_state:
    st.session_state.archivio_progetti = {}
if 'nome_progetto_attivo' not in st.session_state:
    st.session_state.nome_progetto_attivo = "Nuovo_Progetto"


# --- 2. MOTORI MATEMATICI ---
def aggiorna_gerarchia(df):
    df_calc = df.copy()
    df_calc['BAC_Budget'] = pd.to_numeric(df_calc['BAC_Budget'], errors='coerce').fillna(0.0)
    df_calc['AC_Costo_Reale'] = pd.to_numeric(df_calc['AC_Costo_Reale'], errors='coerce').fillna(0.0)
    df_calc['%_Completamento'] = pd.to_numeric(df_calc['%_Completamento'], errors='coerce').fillna(0.0)
    
    for col in ['Inizio_Previsto', 'Fine_Prevista', 'Inizio_Effettivo', 'Fine_Effettiva']:
        if col in df_calc.columns:
            df_calc[col] = df_calc[col].astype(object)
            
    ids = df_calc['ID_WBS'].astype(str).tolist()
    foglie = [uid for uid in ids if not any(other.startswith(uid + '.') for other in ids if other != uid)]
    df_calc['Is_Leaf'] = df_calc['ID_WBS'].astype(str).isin(foglie)
    
    df_calc['Livello'] = df_calc['ID_WBS'].astype(str).apply(lambda x: len(x.split('.')))
    df_calc = df_calc.sort_values(by='Livello', ascending=False)
    
    for index, row in df_calc.iterrows():
        uid = str(row['ID_WBS'])
        if not row['Is_Leaf']:
            discendenti = df_calc[df_calc['ID_WBS'].astype(str).str.startswith(uid + '.') & df_calc['Is_Leaf']]
            if not discendenti.empty:
                df_calc.at[index, 'BAC_Budget'] = discendenti['BAC_Budget'].sum()
                df_calc.at[index, 'AC_Costo_Reale'] = discendenti['AC_Costo_Reale'].sum()
                
                inizio_min = pd.to_datetime(discendenti['Inizio_Previsto']).min()
                fine_max = pd.to_datetime(discendenti['Fine_Prevista']).max()
                if pd.notna(inizio_min): df_calc.at[index, 'Inizio_Previsto'] = inizio_min.date()
                if pd.notna(fine_max): df_calc.at[index, 'Fine_Prevista'] = fine_max.date()
                
                inizio_eff_min = pd.to_datetime(discendenti['Inizio_Effettivo']).min()
                fine_eff_max = pd.to_datetime(discendenti['Fine_Effettiva']).max()
                if pd.notna(inizio_eff_min): df_calc.at[index, 'Inizio_Effettivo'] = inizio_eff_min.date()
                if pd.notna(fine_eff_max): df_calc.at[index, 'Fine_Effettiva'] = fine_eff_max.date()
                
                tot_bac = discendenti['BAC_Budget'].sum()
                if tot_bac > 0:
                    df_calc.at[index, '%_Completamento'] = (discendenti['BAC_Budget'] * discendenti['%_Completamento']).sum() / tot_bac
                else:
                    df_calc.at[index, '%_Completamento'] = discendenti['%_Completamento'].mean()
                    
    df_calc['sort_key'] = df_calc['ID_WBS'].astype(str).apply(lambda x: '.'.join([p.zfill(5) for p in x.split('.')]))
    df_calc = df_calc.sort_values(by='sort_key').drop(columns=['sort_key', 'Is_Leaf', 'Livello']).reset_index(drop=True)
    return df_calc

def modifica_struttura(id_target, azione):
    df = st.session_state.wbs_data.copy()
    def get_sort_key(wbs_id):
        return [int(x) if x.isdigit() else x for x in str(wbs_id).split('.')]
    df['sort_key'] = df['ID_WBS'].apply(get_sort_key)
    df = df.sort_values(by='sort_key').reset_index(drop=True)
    df['Livello'] = df['ID_WBS'].apply(lambda x: len(str(x).split('.')))
    
    if azione == 'elimina':
        mask = (df['ID_WBS'].astype(str) == id_target) | (df['ID_WBS'].astype(str).str.startswith(f"{id_target}."))
        df = df[~mask]
        if df.empty:
            df = pd.DataFrame([{'ID_WBS': '1', 'Attività': 'Progetto Principale', 'BAC_Budget': 0.0, '%_Completamento': 0.0, 'AC_Costo_Reale': 0.0, 'Livello': 1}])
            st.session_state.wbs_data = df.drop(columns=['Livello'])
            st.rerun()
    elif azione in ['su', 'giu', 'destra', 'sinistra']:
        ids = df['ID_WBS'].astype(str).tolist()
        if id_target not in ids: return
        idx = ids.index(id_target)
        livello_target = df.at[idx, 'Livello']
        end_idx = idx + 1
        while end_idx < len(df) and df.at[end_idx, 'Livello'] > livello_target: end_idx += 1
        blocco_target = list(range(idx, end_idx))
        
        if azione == 'destra':
            if idx > 0 and df.at[idx - 1, 'Livello'] >= livello_target: df.loc[blocco_target, 'Livello'] += 1
        elif azione == 'sinistra':
            if livello_target > 1: df.loc[blocco_target, 'Livello'] -= 1
        elif azione == 'su':
            prev_idx = idx - 1
            while prev_idx >= 0 and df.at[prev_idx, 'Livello'] > livello_target: prev_idx -= 1
            if prev_idx >= 0 and df.at[prev_idx, 'Livello'] == livello_target:
                blocco_prev = list(range(prev_idx, idx))
                new_order = list(range(len(df)))
                new_order[prev_idx:end_idx] = blocco_target + blocco_prev
                df = df.iloc[new_order].reset_index(drop=True)
        elif azione == 'giu':
            next_idx = end_idx
            if next_idx < len(df) and df.at[next_idx, 'Livello'] == livello_target:
                next_end = next_idx + 1
                while next_end < len(df) and df.at[next_end, 'Livello'] > livello_target: next_end += 1
                blocco_next = list(range(next_idx, next_end))
                new_order = list(range(len(df)))
                new_order[idx:next_end] = blocco_next + blocco_target
                df = df.iloc[new_order].reset_index(drop=True)

    nuovi_id = []
    counters = {} 
    for idx, row in df.iterrows():
        liv = row['Livello']
        if idx == 0: liv = 1
        else:
            prev_liv = df.at[idx-1, 'Livello']
            if liv > prev_liv + 1: liv = prev_liv + 1 
        df.at[idx, 'Livello'] = liv
        counters[liv] = counters.get(liv, 0) + 1
        for k in list(counters.keys()):
            if k > liv: counters[k] = 0 
        nuovo_id = ".".join([str(counters[i]) for i in range(1, liv + 1)])
        nuovi_id.append(nuovo_id)
        
    old_ids = df['ID_WBS'].astype(str).tolist()
    mapping = dict(zip(old_ids, nuovi_id))
    
    def aggiorna_preds(val):
        if not val or pd.isna(val) or str(val).strip() in ['', 'None', 'nan']: return val
        preds = [p.strip() for p in str(val).split(',')]
        new_preds = []
        for p in preds:
            parts = p.split(' - ', 1)
            vecchio_id = parts[0].strip()
            nuovo_id = mapping.get(vecchio_id, vecchio_id)
            if len(parts) > 1: new_preds.append(f"{nuovo_id} - {parts[1]}")
            else: new_preds.append(nuovo_id)
        return ', '.join(new_preds)
        
    df['ID_WBS'] = nuovi_id
    if 'Predecessori' in df.columns: df['Predecessori'] = df['Predecessori'].apply(aggiorna_preds)
    df = df.drop(columns=['Livello', 'sort_key'])
    st.session_state.wbs_data = df.copy()
    st.session_state.wbs_data = aggiorna_gerarchia(st.session_state.wbs_data)
    st.rerun()
    
def get_foglie(df):
    ids = df['ID_WBS'].astype(str).tolist()
    foglie = [uid for uid in ids if not any(other.startswith(uid + '.') for other in ids if other != uid)]
    return df[df['ID_WBS'].astype(str).isin(foglie)].copy()

def aggiorna_costi_reali():
    df_reg = st.session_state.registro_data.copy()
    wbs = st.session_state.wbs_data
    if not df_reg.empty:
        df_reg['ID_WBS_calc'] = df_reg['Voce_WBS'].astype(str).apply(
            lambda x: str(x).split(' - ')[0].strip() if pd.notna(x) and str(x).strip() not in ['', 'None', 'nan'] else None
        )
        costi_raggruppati = df_reg.groupby('ID_WBS_calc')['Importo_Netto'].sum().reset_index()
        cost_map = dict(zip(costi_raggruppati['ID_WBS_calc'], costi_raggruppati['Importo_Netto']))
        wbs['AC_Costo_Reale'] = wbs['ID_WBS'].apply(lambda x: cost_map.get(str(x), 0.0))
    else:
        wbs['AC_Costo_Reale'] = 0.0
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
            if pd.isna(inizio_ts) or pd.isna(fine_ts) or bac == 0: return 0.0
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
    if date_inizio.empty or date_fine.empty: return None
    min_date = date_inizio.min()
    max_date = date_fine.max()
    date_range = pd.date_range(start=min_date, end=max_date)
    
    df_reg_calc = df_reg.copy()
    if not df_reg_calc.empty:
        df_reg_calc['Data'] = pd.to_datetime(df_reg_calc['Data'], errors='coerce').dt.date
        ac_daily = df_reg_calc.groupby('Data')['Importo_Netto'].sum().to_dict()
    else:
        ac_daily = {}
    
    dati = []
    cum_ac = 0.0
    for d_ts in date_range:
        d = d_ts.date()
        pv_giorno = ev_giorno = 0.0
        for _, row in df_wbs.iterrows():
            bac = float(row['BAC_Budget']) if pd.notna(row['BAC_Budget']) else 0.0
            ip = pd.to_datetime(row['Inizio_Previsto']).date() if pd.notna(row['Inizio_Previsto']) else None
            fp = pd.to_datetime(row['Fine_Prevista']).date() if pd.notna(row['Fine_Prevista']) else None
            if ip and fp and bac > 0:
                if d >= fp: pv_giorno += bac
                elif d > ip:
                    giorni_tot = (fp - ip).days
                    if giorni_tot > 0: pv_giorno += bac * ((d - ip).days / giorni_tot)
            if d <= oggi:
                ev_attuale = float(row['EV']) if 'EV' in row else 0.0
                ie = pd.to_datetime(row['Inizio_Effettivo']).date() if pd.notna(row['Inizio_Effettivo']) else ip
                if ie and ev_attuale > 0:
                    if d >= oggi: ev_giorno += ev_attuale
                    elif d > ie:
                        giorni_lav = (oggi - ie).days
                        if giorni_lav > 0: ev_giorno += ev_attuale * ((d - ie).days / giorni_lav)
        if d <= oggi:
            cum_ac += ac_daily.get(d, 0.0)
            ac_val, ev_val = cum_ac, ev_giorno
        else:
            ac_val = ev_val = None
        dati.append({'Data': d, 'PV (Valore Pianificato)': pv_giorno, 'EV (Valore Guadagnato)': ev_val, 'AC (Costo Reale)': ac_val})
    return pd.DataFrame(dati)

def calcola_cpm(df_wbs):
    df_wp = get_foglie(df_wbs)
    cpm_nodes = {}
    for _, row in df_wp.iterrows():
        node_id = str(row['ID_WBS']).strip()
        inizio = pd.to_datetime(row['Inizio_Previsto'], errors='coerce')
        fine = pd.to_datetime(row['Fine_Prevista'], errors='coerce')
        durata = max((fine - inizio).days + 1, 1) if pd.notna(inizio) and pd.notna(fine) else 1
        pred_val = str(row.get('Predecessori', '')).strip()
        preds = []
        if pred_val and pred_val.lower() not in ['none', 'nan', 'null']:
            for p in pred_val.split(','):
                p_id = p.split(' - ')[0].strip()
                if p_id.endswith('.0'): p_id = p_id[:-2]
                if p_id: preds.append(p_id)
        cpm_nodes[node_id] = {'durata': durata, 'preds': preds, 'succs': [], 'ES': 0, 'EF': 0, 'LS': 0, 'LF': 0, 'slack': 0, 'is_critical': False}
    for n_id, data in cpm_nodes.items():
        for p_id in data['preds']:
            if p_id in cpm_nodes: cpm_nodes[p_id]['succs'].append(n_id)
    changed = True
    while changed:
        changed = False
        for n_id, data in cpm_nodes.items():
            max_ef = 0
            for p_id in data['preds']:
                if p_id in cpm_nodes: max_ef = max(max_ef, cpm_nodes[p_id]['EF'])
            new_es = max_ef
            new_ef = new_es + data['durata']
            if new_es != data['ES'] or new_ef != data['EF']:
                data['ES'] = new_es; data['EF'] = new_ef; changed = True
    project_duration = max([data['EF'] for data in cpm_nodes.values()], default=0)
    for n_id, data in cpm_nodes.items():
        data['LF'] = project_duration
        data['LS'] = data['LF'] - data['durata']
    changed = True
    while changed:
        changed = False
        for n_id, data in cpm_nodes.items():
            min_ls = data['LF'] 
            if len(data['succs']) > 0: min_ls = min([cpm_nodes[s_id]['LS'] for s_id in data['succs'] if s_id in cpm_nodes])
            new_lf = min_ls
            new_ls = new_lf - data['durata']
            if new_lf != data['LF'] or new_ls != data['LS']:
                data['LF'] = new_lf; data['LS'] = new_ls; changed = True
    for n_id, data in cpm_nodes.items():
        data['slack'] = data['LS'] - data['ES']
        if data['slack'] <= 0: data['is_critical'] = True
    return cpm_nodes

# --- 3. ESECUZIONE CALCOLI INIZIALI ---
aggiorna_costi_reali()
st.session_state.wbs_data = aggiorna_gerarchia(st.session_state.wbs_data)
st.session_state.wbs_data = calcola_evm(st.session_state.wbs_data, pd.Timestamp.today().date())

# --- CREAZIONE TAB ---
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🗂️ WBS (Lavorazioni)", "👥 OBS (Risorse)", "🕸️ Nodi & Matrice", 
    "📅 Cronoprogramma", "📈 Earned Value & Cash Flow", "🧾 Reg. Contabile", "🛠️ Direzione & CAPA"
])

# --- TAB 1: SETUP WBS ---
with tab1:
    st.header("WBS - Work Breakdown Structure")
    st.markdown('*Usa i pulsanti sotto ogni capitolo per spostare e rientrare le voci.*')
    st.subheader("🔀 Organizzatore Capitoli")
    c_sel, c_btn1, c_btn2, c_btn3 = st.columns([3, 1, 1, 1])
    df_padri = st.session_state.wbs_data[~st.session_state.wbs_data['ID_WBS'].astype(str).str.contains(r'\.')]
    lista_capitoli = df_padri['ID_WBS'].astype(str) + " - " + df_padri['Attività'].astype(str)
    nodo_scelto = c_sel.selectbox("Seleziona una voce", options=lista_capitoli, label_visibility="collapsed")
    if nodo_scelto:
        id_scelto = nodo_scelto.split(' - ')[0]
        if c_btn1.button("⬆️ Su", use_container_width=True): modifica_struttura(id_scelto, 'su')
        if c_btn2.button("⬇️ Giù", use_container_width=True): modifica_struttura(id_scelto, 'giu')
        if c_btn3.button("🗑️ Elimina", use_container_width=True): modifica_struttura(id_scelto, 'elimina')
            
    st.divider()
    df = st.session_state.wbs_data.copy()
    colonne_date = ['Inizio_Previsto', 'Fine_Prevista', 'Inizio_Effettivo', 'Fine_Effettiva']
    for col in colonne_date:
        if col in df.columns: df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
    df['Durata_Prevista (gg)'] = (pd.to_datetime(df['Fine_Prevista']) - pd.to_datetime(df['Inizio_Previsto'])).dt.days + 1
    radici = df[~df['ID_WBS'].astype(str).str.contains(r'\.')]
    df_aggiornato = pd.DataFrame()
    lista_obs_dropdown = [""] + [f"{row['ID_OBS']} - {row['Risorsa']}" for _, row in st.session_state.obs_data.iterrows() if pd.notna(row['ID_OBS'])]
    lista_wbs_dropdown = [""] + [f"{row['ID_WBS']} - {row['Attività']}" for _, row in df.iterrows() if pd.notna(row['ID_WBS'])]
    
    for idx_riga, radice in radici.iterrows():
        id_radice = str(radice['ID_WBS'])
        discendenti = df[df['ID_WBS'].astype(str).str.startswith(f"{id_radice}.")]
        with st.expander(f"📁 {id_radice} - {radice['Attività']} (Budget Raggruppato: € {radice['BAC_Budget']:,.2f})", expanded=True):
            colonne_bloccate = ["ID_WBS", "Durata_Prevista (gg)", "AC_Costo_Reale", "PV", "EV", "CV", "SV", "SPI", "CPI", "EAC", "ETC", "VAC"]
            colonne_bloccate = [col for col in colonne_bloccate if col in discendenti.columns]
            discendenti_modificati = st.data_editor(
                discendenti, key=f"editor_wbs_idx_{idx_riga}_id_{id_radice}", num_rows="dynamic", use_container_width=True, hide_index=True, disabled=colonne_bloccate, 
                column_config={
                    "ID_WBS": st.column_config.TextColumn("ID WBS (Auto)"),
                    "Predecessori": st.column_config.SelectboxColumn("Predecessore ▾", options=lista_wbs_dropdown),
                    "ID_OBS_Assegnato": st.column_config.SelectboxColumn("Risorsa ▾", options=lista_obs_dropdown),
                    "Inizio_Previsto": st.column_config.DateColumn("Inizio Previsto"),
                    "Fine_Prevista": st.column_config.DateColumn("Fine Prevista"),
                    "Inizio_Effettivo": st.column_config.DateColumn("Inizio Effettivo"),
                    "Fine_Effettiva": st.column_config.DateColumn("Fine Effettiva")
                }
            )
            for i_row, row_mod in discendenti_modificati.iterrows():
                val_id = str(row_mod['ID_WBS']).strip()
                if val_id in ['', 'None', 'nan']: discendenti_modificati.at[i_row, 'ID_WBS'] = f"{id_radice}.999{i_row}"
            df_aggiornato = pd.concat([df_aggiornato, pd.DataFrame([radice]), discendenti_modificati], ignore_index=True)
            if not discendenti.empty:
                c_sel_int, c1, c2, c3, c4 = st.columns([3, 1.5, 1.5, 1, 1])
                nodo_locale = c_sel_int.selectbox("Seleziona", options=discendenti['ID_WBS'].astype(str) + " - " + discendenti['Attività'].astype(str), key=f"sel_move_{id_radice}", label_visibility="collapsed")
                if nodo_locale:
                    id_loc = nodo_locale.split(' - ')[0]
                    if c1.button("⬅️ Estrai", key=f"l_{id_radice}", use_container_width=True): modifica_struttura(id_loc, 'sinistra')
                    if c2.button("➡️ Rientra", key=f"r_{id_radice}", use_container_width=True): modifica_struttura(id_loc, 'destra')
                    if c3.button("⬆️ Su", key=f"u_{id_radice}", use_container_width=True): modifica_struttura(id_loc, 'su')
                    if c4.button("⬇️ Giù", key=f"d_{id_radice}", use_container_width=True): modifica_struttura(id_loc, 'giu')

    with st.form("aggiungi_padre"):
        st.write("Aggiungi Capitolo Principale")
        c1, c2 = st.columns([4, 1])
        nuova_att = c1.text_input("Nome", placeholder="Es. Impianti Elettrici")
        if c2.form_submit_button("➕ Aggiungi Capitolo"):
            if nuova_att:
                is_root_calc = ~st.session_state.wbs_data['ID_WBS'].astype(str).str.contains(r'\.')
                nuovo_id = str(len(st.session_state.wbs_data[is_root_calc]) + 1)
                st.session_state.wbs_data = pd.concat([st.session_state.wbs_data, pd.DataFrame([{'ID_WBS': nuovo_id, 'Attività': nuova_att, 'BAC_Budget': 0.0, '%_Completamento': 0.0, 'AC_Costo_Reale': 0.0}])], ignore_index=True)
                modifica_struttura('1', 'rinumera') 

    st.divider()
    if st.button("🔄 RICALCOLA ALBERO WBS", type="primary", use_container_width=True):
        st.session_state.wbs_data = df_aggiornato.copy() 
        modifica_struttura('1', 'rinumera')
        
# --- TAB 2: SETUP OBS ---
with tab2:
    st.header("OBS - Organization Breakdown Structure")
    st.info("✅ **Autosave Attivo:** Tutte le modifiche fatte in questa tabella vengono scritte istantaneamente nel progetto. **Non ci sono pulsanti da cliccare per salvare.**")
    
    with st.expander("⚙️ Gestione Colonne Aggiuntive", expanded=False):
        c1, c2 = st.columns([3, 1])
        nuova_col = c1.text_input("Nome nuova colonna")
        if c2.button("➕ Crea") and nuova_col:
            if nuova_col not in st.session_state.obs_data.columns:
                st.session_state.obs_data[nuova_col] = "" 
                st.session_state.editor_key_obs += 1
                st.rerun()
        
        colonne_custom = [c for c in st.session_state.obs_data.columns if c not in ['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note']]
        if colonne_custom:
            c3, c4, c5 = st.columns([2, 2, 1])
            col_da_modificare = c3.selectbox("Colonna da rinominare", options=colonne_custom)
            nuovo_nome = c4.text_input("Nuovo nome intestazione")
            if c5.button("✏️ Modifica") and nuovo_nome:
                st.session_state.obs_data.rename(columns={col_da_modificare: nuovo_nome}, inplace=True)
                st.session_state.editor_key_obs += 1
                st.rerun()

    # Questo è il protocollo corretto: st.session_state sovrascrive sé stesso
    st.session_state.obs_data = st.data_editor(
        st.session_state.obs_data, 
        key=f"obs_editor_{st.session_state.editor_key_obs}",
        column_config={
            "Tipo_Contratto": st.column_config.SelectboxColumn("Tipo Contratto", options=["Appalto ▾", "Sub appalto ▾"], required=True)
        },
        num_rows="dynamic", use_container_width=True, hide_index=True
    )
        
# --- TAB 3: MATRICE E GRAFO A NODI ---
with tab3:
    st.header("Incrocio Logico (Work Packages e Percorso Critico)")
    cpm_data = calcola_cpm(st.session_state.wbs_data)
    mostra_relazioni = st.toggle("👁️ Mostra Relazioni", value=True)
    graph = graphviz.Digraph(engine='dot')
    graph.attr(rankdir='LR', ranksep='1.5', nodesep='0.8', splines='spline')
    graph.attr('node', fontname='Helvetica', fontsize='10', margin='0.2')
    
    for _, row in st.session_state.obs_data.iterrows():
        ruolo_safe = str(row.get('Ruolo', '')).replace('<', '').replace('>', '').replace('&', 'e')
        risorsa_safe = str(row.get('Risorsa', '')).replace('<', '').replace('>', '').replace('&', 'e')
        label_html = f"<<TABLE BORDER='0' CELLBORDER='0' CELLSPACING='2'><TR><TD><B>{ruolo_safe}</B></TD></TR><TR><TD>({risorsa_safe})</TD></TR>"
        for col in [c for c in st.session_state.obs_data.columns if c not in ['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note']]:
            valore = str(row.get(col, '')).replace('<', '').replace('>', '').replace('&', 'e')
            if pd.notna(row.get(col)) and valore.strip() != "":
                label_html += f"<TR><TD><FONT POINT-SIZE='9' COLOR='gray30'>{str(col).replace('<', '').replace('>', '').replace('&', 'e')}: {valore}</FONT></TD></TR>"
        label_html += "</TABLE>>"
        obs_id = str(row.get('ID_OBS', '')).strip()
        if obs_id.endswith('.0'): obs_id = obs_id[:-2]
        if obs_id: graph.node(f"OBS_{obs_id}", label=label_html, shape='rect', style='rounded,filled', fillcolor='#E1F5FE', color='#0288D1', penwidth='1.5')
        
    df_wp_reali = get_foglie(st.session_state.wbs_data)
    for _, row in df_wp_reali.iterrows():
        wbs_id = str(row.get('ID_WBS', '')).strip()
        if not wbs_id: continue
        attivita = str(row.get('Attività', '')).replace('<', '').replace('>', '').replace('&', 'e')
        budget, costo_reale, completamento = float(row['BAC_Budget']), float(row['AC_Costo_Reale']), float(row['%_Completamento'])
        wp_cpm = cpm_data.get(wbs_id, {})
        margine, is_critical = wp_cpm.get('slack', 0), wp_cpm.get('is_critical', False)
        inizio_val, fine_val = pd.to_datetime(row.get('Inizio_Previsto'), errors='coerce'), pd.to_datetime(row.get('Fine_Prevista'), errors='coerce')
        inizio_str = inizio_val.strftime('%d/%m/%Y') if pd.notna(inizio_val) else "N/D"
        fine_str = fine_val.strftime('%d/%m/%Y') if pd.notna(fine_val) else "N/D"
        testo_margine = f"<FONT COLOR='#D32F2F'><B>Margine: {margine} gg</B></FONT>" if is_critical else f"<FONT COLOR='#388E3C'>Margine: {margine} gg</FONT>"
        wp_html = f"<<TABLE BORDER='0' CELLBORDER='0' CELLSPACING='4'><TR><TD COLSPAN='2'><B>{wbs_id} - {attivita}</B></TD></TR><TR><TD ALIGN='LEFT'>Inizio: {inizio_str}</TD><TD ALIGN='RIGHT'>Fine: {fine_str}</TD></TR><TR><TD ALIGN='LEFT'>Budget: € {budget:,.2f}</TD><TD ALIGN='RIGHT'>AC: € {costo_reale:,.2f}</TD></TR><TR><TD ALIGN='LEFT'>Avanzamento: {completamento:.1f}%</TD><TD ALIGN='RIGHT'>{testo_margine}</TD></TR></TABLE>>"
        
        stile, colore_sfondo = ('rounded,filled', '#C8E6C9') if completamento >= 100 else ('rounded,filled', 'white') if completamento <= 0 else ('rounded,striped', f"#C8E6C9;{completamento / 100.0}:white")
        bordo_colore, spessore_bordo = ('#D32F2F', '3.0') if is_critical else ('#388E3C', '1.5')
        graph.node(f"WBS_{wbs_id}", label=wp_html, shape='rect', style=stile, fillcolor=colore_sfondo, color=bordo_colore, penwidth=spessore_bordo)
        
        obs_val = str(row.get('ID_OBS_Assegnato', '')).strip()
        if obs_val and obs_val.lower() not in ['none', 'nan', 'null']:
            for o in [o.strip() for o in obs_val.split(',')]:
                o_id = o.split(' - ')[0].strip()
                if o_id.endswith('.0'): o_id = o_id[:-2]
                if o_id: graph.edge(f"OBS_{o_id}", f"WBS_{wbs_id}", color='#757575', penwidth='1.5', arrowsize='0.8')
                    
        pred_val = str(row.get('Predecessori', '')).strip()
        if mostra_relazioni and pred_val and pred_val.lower() not in ['none', 'nan', 'null']:
            for p in [p.strip() for p in pred_val.split(',')]:
                p_id = p.split(' - ')[0].strip()
                if p_id.endswith('.0'): p_id = p_id[:-2]
                if p_id in set(df_wp_reali['ID_WBS'].astype(str)):
                    pred_is_critical = cpm_data.get(p_id, {}).get('is_critical', False)
                    colore_cavo, stile_cavo, spessore_cavo, freccia = ('#D32F2F', 'solid', '2.5', '1.0') if is_critical and pred_is_critical else ('#FF9800', 'dashed', '1.0', '0.6')
                    graph.edge(f"WBS_{p_id}", f"WBS_{wbs_id}", color=colore_cavo, style=stile_cavo, penwidth=spessore_cavo, arrowsize=freccia)

    try:
        raw_svg = graph.pipe(format='svg').decode('utf-8')
        if '<svg' in raw_svg:
            components.html(f"""<!DOCTYPE html><html><head><script src="https://cdn.jsdelivr.net/npm/svg-pan-zoom@3.6.1/dist/svg-pan-zoom.min.js"></script><style>body {{ margin: 0; padding: 0; overflow: hidden; background-color: #fafafa; }} #svg-container {{ width: 100vw; height: 100vh; display: flex; justify-content: center; align-items: center; }} svg {{ width: 100% !important; height: 100% !important; }}</style></head><body><div id="svg-container">{raw_svg[raw_svg.find('<svg'):]}</div><script>window.onload = function() {{ var svgElement = document.querySelector('svg'); if (svgElement) {{ svgElement.setAttribute('id', 'grafo-interattivo'); svgElement.removeAttribute('width'); svgElement.removeAttribute('height'); svgPanZoom('#grafo-interattivo', {{ zoomEnabled: true, controlIconsEnabled: true, fit: true, center: true, minZoom: 0.1, maxZoom: 10, mouseWheelZoomEnabled: true }}); }} }};</script></body></html>""", height=600)
        else:
            st.info("Il grafo logico è vuoto.")
    except Exception as e:
        st.error(f"Errore generazione grafo: {e}")
        st.graphviz_chart(graph)
        
# --- TAB 4: CRONOPROGRAMMA (GANTT) ---
with tab4:
    st.header("Cronoprogramma Lavori")
    c1, c2 = st.columns([1, 2])
    vista = c1.selectbox("Seleziona Vista", ["Progetto (Baseline)", "Esecuzione (Esecutivo)", "Comparativa"])
    data_status_gantt = c2.date_input("📅 Data Rilevamento", value=pd.Timestamp.today().date())
    
    df_gantt = get_foglie(st.session_state.wbs_data).copy().dropna(subset=['Inizio_Previsto', 'Fine_Prevista'])
    if not df_gantt.empty:
        df_gantt['Inizio_Previsto'] = pd.to_datetime(df_gantt['Inizio_Previsto'])
        df_gantt['Fine_Prevista'] = pd.to_datetime(df_gantt['Fine_Prevista'])
        df_gantt['Inizio_Effettivo'] = pd.to_datetime(df_gantt['Inizio_Effettivo'])
        df_gantt['Fine_Effettiva'] = pd.to_datetime(df_gantt['Fine_Effettiva']).fillna(pd.to_datetime(data_status_gantt))
        
        fig = go.Figure()
        if vista in ["Progetto (Baseline)", "Comparativa"]:
            fig.add_trace(go.Bar(x=(df_gantt['Fine_Prevista'] - df_gantt['Inizio_Previsto'] + pd.Timedelta(days=1)).dt.total_seconds() * 1000, y=df_gantt['ID_WBS'].astype(str) + " - " + df_gantt['Attività'], base=df_gantt['Inizio_Previsto'], orientation='h', name='Baseline', width=0.4, marker=dict(color='rgba(0, 0, 255, 0.4)' if vista == "Comparativa" else 'blue')))
        if vista in ["Esecuzione (Esecutivo)", "Comparativa"]:
            df_esec = df_gantt.dropna(subset=['Inizio_Effettivo']).copy()
            if not df_esec.empty:
                fig.add_trace(go.Bar(x=(df_esec['Fine_Effettiva'] - df_esec['Inizio_Effettivo'] + pd.Timedelta(days=1)).dt.total_seconds() * 1000, y=df_esec['ID_WBS'].astype(str) + " - " + df_esec['Attività'], base=df_esec['Inizio_Effettivo'], orientation='h', name='Esecutivo', width=0.2, marker=dict(color='red')))
        fig.update_layout(barmode='overlay', height=600, bargap=0.3, xaxis_title="Linea Temporale", yaxis_title="Lavorazioni (WBS)", yaxis={'autorange': 'reversed'}, xaxis_type='date')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("⚠️ Il cronoprogramma è vuoto.")
        
# --- TAB 5: EVM ---
with tab5:
    st.header("Controllo Costi e Analisi EVM")
    data_status_evm = st.date_input("📅 Data di Stato (EVM):", value=date(2026, 10, 15))
    df_evm = calcola_evm(get_foglie(st.session_state.wbs_data), data_status_evm)
    tot_bac, tot_pv, tot_ev, tot_ac = df_evm['BAC_Budget'].sum(), df_evm['PV'].sum(), df_evm['EV'].sum(), df_evm['AC_Costo_Reale'].sum()
    tot_eac, tot_etc, tot_vac = df_evm['EAC'].sum(), df_evm['ETC'].sum(), df_evm['VAC'].sum()
    cpi_globale = tot_ev / tot_ac if tot_ac > 0 else 1.0
    spi_globale = tot_ev / tot_pv if tot_pv > 0 else 1.0
    
    col_box1, col_box2 = st.columns(2)
    with col_box1:
        with st.container(border=True):
            st.markdown("#### 📊 Stato Attuale")
            c1, c2, c3 = st.columns(3)
            c1.metric("BAC", f"€ {tot_bac:,.0f}")
            c2.metric("EV", f"€ {tot_ev:,.0f}")
            c3.metric("AC", f"€ {tot_ac:,.0f}")
            st.divider()
            c4, c5 = st.columns(2)
            c4.metric("Avanzamento", f"{(tot_ev / tot_bac * 100) if tot_bac > 0 else 0.0:.1f}%")
            c5.metric("SPI", f"{spi_globale:.2f}")

    with col_box2:
        with st.container(border=True):
            st.markdown("#### 🔮 Previsioni")
            c1, c2, c3 = st.columns(3)
            c1.metric("EAC", f"€ {tot_eac:,.0f}")
            c2.metric("ETC", f"€ {tot_etc:,.0f}")
            c3.metric("VAC", f"€ {tot_vac:,.0f}")
            st.divider()
            c4, c5 = st.columns(2)
            c4.metric("CPI", f"{cpi_globale:.2f}")
            
    st.divider()
    df_scurve = genera_dati_scurve(df_evm, st.session_state.registro_data, data_status_evm)
    if df_scurve is not None and not df_scurve.empty:
        fig_scurve = px.line(df_scurve, x='Data', y=['PV (Valore Pianificato)', 'EV (Valore Guadagnato)', 'AC (Costo Reale)'], color_discrete_map={'PV (Valore Pianificato)': 'blue', 'EV (Valore Guadagnato)': 'green', 'AC (Costo Reale)': 'red'}, labels={'value': 'Importo (€)', 'variable': 'Metrica'})
        df_past = df_scurve[df_scurve['Data'] <= data_status_evm]
        if not df_past.empty:
            min_date, max_date = df_scurve['Data'].min(), df_scurve['Data'].max()
            spi_effettivo = df_past.iloc[-1]['EV (Valore Guadagnato)'] / df_past.iloc[-1]['PV (Valore Pianificato)'] if df_past.iloc[-1]['PV (Valore Pianificato)'] > 0 else 1.0
            giorni_stimati = min(int((max_date - min_date).days / spi_effettivo) if spi_effettivo > 0 else (max_date - min_date).days, (max_date - min_date).days * 3) 
            data_fine_stimata = min_date + pd.Timedelta(days=giorni_stimati)
            fig_scurve.add_trace(go.Scatter(x=[data_status_evm, data_fine_stimata], y=[df_past.iloc[-1]['AC (Costo Reale)'], tot_eac], mode='lines', line=dict(color='red', dash='dot'), name='Proiez. Costi'))
            fig_scurve.add_trace(go.Scatter(x=[data_status_evm, data_fine_stimata], y=[df_past.iloc[-1]['EV (Valore Guadagnato)'], tot_bac], mode='lines', line=dict(color='green', dash='dot'), name='Proiez. Lavoro'))
        fig_scurve.update_layout(hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        st.plotly_chart(fig_scurve, use_container_width=True)

# --- TAB 6: REGISTRO CONTABILE ---
with tab6:
    st.header("Registro Contabile")
    st.info("✅ **Autosave Attivo:** Le fatture inserite finiscono immediatamente nel database e si collegano alla WBS senza bisogno di premere alcun tasto.")
    
    leaf_wbs = get_foglie(st.session_state.wbs_data)
    wbs_options = [f"{row['ID_WBS']} - {row['Attività']}" for _, row in leaf_wbs.iterrows()]
    obs_options = [""] + [f"{row['ID_OBS']} - {row['Risorsa']}" for _, row in st.session_state.obs_data.iterrows() if pd.notna(row['ID_OBS'])]
    
    st.session_state.registro_data = st.data_editor(
        st.session_state.registro_data,
        key=f"reg_editor_{st.session_state.editor_key_reg}",
        num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Data": st.column_config.DateColumn("Data Registrazione"),
            "N_Doc": st.column_config.TextColumn("N° Doc/Fattura"),
            "Fornitore": st.column_config.SelectboxColumn("Fornitore (OBS) ▾", options=obs_options),
            "Descrizione": st.column_config.TextColumn("Descrizione / Note"),
            "Importo_Netto": st.column_config.NumberColumn("Importo Netto (€)", format="€ %.2f", min_value=0.0),
            "Voce_WBS": st.column_config.SelectboxColumn("Attività WBS (Destinazione) ▾", options=wbs_options, required=True)
        }
    )
        
# --- TAB 7: CAPA ---
with tab7:
    st.header("Direzione Lavori: Interventi (CAPA) e Simulazioni")
    
    leaf_wbs_capa = get_foglie(st.session_state.wbs_data)
    wbs_options_capa = [f"{row['ID_WBS']} - {row['Attività']}" for _, row in leaf_wbs_capa.iterrows()]
    obs_options_capa = [f"{row['ID_OBS']} - {row['Risorsa']}" for _, row in st.session_state.obs_data.iterrows()]
    
    st.subheader("1. Registro Azioni Correttive e Preventive (CAPA)")
    st.info("✅ **Autosave Attivo**")
    
    st.session_state.capa_data = st.data_editor(
        st.session_state.capa_data,
        key=f"capa_editor_{st.session_state.editor_key_capa}",
        num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Data_Apertura": st.column_config.DateColumn("Data Segnalazione"),
            "ID_WBS_Rif": st.column_config.SelectboxColumn("Attività WBS", options=wbs_options_capa, required=True),
            "Tipo_Azione": st.column_config.SelectboxColumn("Tipo", options=["Correttiva ▾", "Preventiva ▾"], required=True),
            "Descrizione": st.column_config.TextColumn("Descrizione", width="large"),
            "Responsabile_OBS": st.column_config.SelectboxColumn("Risorsa (OBS)", options=obs_options_capa, required=True),
            "Stato": st.column_config.SelectboxColumn("Stato", options=["Aperto ▾", "In Lavorazione ▾", "Chiuso ▾"], required=True)
        }
    )
    
    with st.expander("🔬 2. Ambiente di Simulazione (Crashing)"):
        c_sim1, c_sim2, c_sim3 = st.columns([2, 1.5, 1.5])
        wp_scelto = c_sim1.selectbox("Seleziona Work Package da simulare", options=wbs_options_capa)
        extra_costo = c_sim2.number_input("Iniezione Extra Costo (€)", value=0.0, step=500.0)
        giorni_risparmiati = c_sim3.number_input("Stima recupero ritardo (Giorni)", value=0, step=1, min_value=0)
        
        if wp_scelto:
            wp_id = wp_scelto.split(' - ')[0]
            df_simulazione = st.session_state.wbs_data.copy()
            df_simulazione['AC_Costo_Reale'] = pd.to_numeric(df_simulazione['AC_Costo_Reale'], errors='coerce').fillna(0.0)
            indice_riga = df_simulazione.index[df_simulazione['ID_WBS'] == wp_id].tolist()
            if indice_riga: df_simulazione.at[indice_riga[0], 'AC_Costo_Reale'] += extra_costo
                
            oggi = pd.Timestamp.today().date()
            df_sim_calc, df_reale_calc = calcola_evm(get_foglie(df_simulazione), oggi), calcola_evm(get_foglie(st.session_state.wbs_data), oggi)
            
            eac_attuale, eac_simulato = df_reale_calc['EAC'].sum(), df_sim_calc['EAC'].sum()
            ac_attuale_tot, ac_simulato_tot = df_reale_calc['AC_Costo_Reale'].sum(), df_sim_calc['AC_Costo_Reale'].sum()
            
            min_date, max_date = pd.to_datetime(df_reale_calc['Inizio_Previsto']).min().date(), pd.to_datetime(df_reale_calc['Fine_Prevista']).max().date()
            if pd.notna(min_date) and pd.notna(max_date):
                spi_attuale = df_reale_calc['EV'].sum() / df_reale_calc['PV'].sum() if df_reale_calc['PV'].sum() > 0 else 1.0
                data_fine_attuale = min_date + pd.Timedelta(days=min(int((max_date - min_date).days / spi_attuale) if spi_attuale > 0 else (max_date - min_date).days, (max_date - min_date).days * 3))
                data_fine_simulata = data_fine_attuale - pd.Timedelta(days=giorni_risparmiati)
            else:
                data_fine_attuale = data_fine_simulata = oggi
            
            c_res1, c_res2 = st.columns([1, 2])
            with c_res1:
                st.metric("EAC Simulato", f"€ {eac_simulato:,.0f}", delta=f"Delta: € {eac_simulato - eac_attuale:,.0f}", delta_color="inverse")
                st.metric("Data di Consegna", data_fine_simulata.strftime('%d/%m/%Y'), delta=f"Anticipo: {giorni_risparmiati} gg", delta_color="normal")
            
            with c_res2:
                fig_sim = go.Figure()
                fig_sim.add_trace(go.Scatter(x=[oggi, data_fine_attuale], y=[ac_attuale_tot, eac_attuale], mode='lines+markers+text', name='Attuale', line=dict(color='red', dash='dash', width=2)))
                fig_sim.add_trace(go.Scatter(x=[oggi, data_fine_simulata], y=[ac_simulato_tot, eac_simulato], mode='lines+markers+text', name='Simulato', line=dict(color='blue', width=3)))
                st.plotly_chart(fig_sim, use_container_width=True)

    st.subheader("3. Stampa Verbale di Direzione Lavori")
    if st.button("📄 Genera Verbale WORD (.docx)", use_container_width=True, type="primary"):
        doc = Document()
        doc.add_heading('VERBALE DI DIREZIONE LAVORI', 0)
        doc.add_paragraph(f"Progetto: {st.session_state.nome_progetto_attivo}")
        if not st.session_state.capa_data.empty:
            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'
            for i, h in enumerate(['Data', 'WBS', 'Tipo', 'Descrizione', 'Stato']): table.rows[0].cells[i].text = h
            for _, row in st.session_state.capa_data.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text, row_cells[1].text, row_cells[2].text, row_cells[3].text, row_cells[4].text = str(row['Data_Apertura']), str(row['ID_WBS_Rif']), str(row['Tipo_Azione']), str(row['Descrizione']), str(row['Stato'])
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.download_button("⬇️ Scarica il file Word", data=buffer, file_name=f"Verbale_{st.session_state.nome_progetto_attivo}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# --- SIDEBAR: GESTIONE PROGETTI E EXPORT (A Fine Script per massima precisione!) ---
with st.sidebar:
    st.header("📂 Gestione Progetti")
    st.session_state.nome_progetto_attivo = st.text_input("Nome Progetto Attuale", value=st.session_state.nome_progetto_attivo)
    
    if st.button("📄 Nuovo Progetto (Reset Dati)", use_container_width=True):
        for key in ['wbs_data', 'obs_data', 'registro_data', 'capa_data']:
            if key in st.session_state: del st.session_state[key]
        st.session_state.editor_key_obs += 1
        st.session_state.editor_key_reg += 1
        st.session_state.editor_key_capa += 1
        st.session_state.nome_progetto_attivo = "Nuovo_Progetto"
        st.rerun()
        
    st.divider()
    
    st.write("💾 **Esporta su PC**")
    try:
        progetto_export = {
            "wbs": json.loads(st.session_state.wbs_data.to_json(orient="records", date_format="iso")),
            "obs": json.loads(st.session_state.obs_data.to_json(orient="records")),
            "registro": json.loads(st.session_state.registro_data.to_json(orient="records", date_format="iso")),
            "capa": json.loads(st.session_state.capa_data.to_json(orient="records", date_format="iso"))
        }
        json_string = json.dumps(progetto_export, indent=4)
        
        st.download_button(
            label="⬇️ Scarica File (.json)",
            data=json_string,
            file_name=f"{st.session_state.nome_progetto_attivo}.json",
            mime="application/json",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Errore esportazione: {e}")
    
    st.divider()
    uploaded_file = st.file_uploader("📤 Carica File da PC", type=['json'], label_visibility="collapsed")
    
    if uploaded_file is not None:
        if 'ultimo_file_caricato' not in st.session_state or st.session_state.ultimo_file_caricato != uploaded_file.file_id:
            try:
                dati_caricati = json.load(uploaded_file)
                
                # --- CARICAMENTO E SICUREZZA ---
                df_wbs = pd.DataFrame(dati_caricati.get('wbs', []))
                if df_wbs.empty:
                    df_wbs = pd.DataFrame([{'ID_WBS': '1', 'Attività': 'Progetto Principale', 'BAC_Budget': 0.0, '%_Completamento': 0.0, 'AC_Costo_Reale': 0.0, 'ID_OBS_Assegnato': None, 'Predecessori': ''}])
                for col in ['Inizio_Previsto', 'Fine_Prevista', 'Inizio_Effettivo', 'Fine_Effettiva']:
                    if col in df_wbs.columns: df_wbs[col] = pd.to_datetime(df_wbs[col], errors='coerce').dt.date
                st.session_state.wbs_data = df_wbs
                
                df_obs = pd.DataFrame(dati_caricati.get('obs', []))
                if df_obs.empty:
                    df_obs = pd.DataFrame(columns=['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note'])
                st.session_state.obs_data = df_obs
                
                df_reg = pd.DataFrame(dati_caricati.get('registro', []))
                if df_reg.empty:
                    df_reg = pd.DataFrame(columns=['Data', 'N_Doc', 'Fornitore', 'Voce_WBS', 'Importo_Netto', 'Descrizione'])
                if 'Data' in df_reg.columns: df_reg['Data'] = pd.to_datetime(df_reg['Data'], errors='coerce').dt.date
                st.session_state.registro_data = df_reg
                
                df_capa = pd.DataFrame(dati_caricati.get('capa', []))
                if df_capa.empty:
                    df_capa = pd.DataFrame(columns=['Data_Apertura', 'ID_WBS_Rif', 'Tipo_Azione', 'Descrizione', 'Responsabile_OBS', 'Stato'])
                if 'Data_Apertura' in df_capa.columns: df_capa['Data_Apertura'] = pd.to_datetime(df_capa['Data_Apertura'], errors='coerce').dt.date
                st.session_state.capa_data = df_capa
                
                st.session_state.wbs_data = aggiorna_gerarchia(st.session_state.wbs_data)
                st.session_state.nome_progetto_attivo = uploaded_file.name.replace(".json", "")
                st.session_state.ultimo_file_caricato = uploaded_file.file_id
                
                # ROTAZIONE CHIAVI: Forza Streamlit a cancellare la memoria grafica vecchia
                st.session_state.editor_key_obs += 1
                st.session_state.editor_key_reg += 1
                st.session_state.editor_key_capa += 1
                
                st.success("Dati ripristinati correttamente!")
                st.rerun() 
                
            except Exception as e:
                st.error(f"Errore critico durante la lettura del File JSON: {e}")
