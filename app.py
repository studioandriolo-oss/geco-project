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
st.set_page_config(page_title="WBS/OBS Manager & EVM", layout="wide", initial_sidebar_state="expanded")

# --- RIMOZIONE TOTALE HEADER, SIDEBAR E STILE COLONNA COMPATTA ---
st.markdown("""
<style>
header[data-testid="stHeader"] {display: none !important;}
#MainMenu, footer, [data-testid="stSidebar"], [data-testid="collapsedControl"] {display: none !important;}
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 1rem !important;
    max-width: 98% !important;
}
.btn-compatto button {
    padding: 0.2rem 0.5rem !important;
    min-height: 30px !important;
    font-size: 0.85rem !important;
    margin-bottom: 0px !important;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# IL MOTORE (PRIMA DELLA GRAFICA)
# ==========================================

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

# --- 1. INIZIALIZZAZIONE DATI ---
if 'wbs_data' not in st.session_state:
    st.session_state.wbs_data = pd.DataFrame([{
        'ID_WBS': '1', 
        'Attività': 'Progetto Principale', 
        'Inizio_Previsto': None, 'Fine_Prevista': None, 
        'Inizio_Effettivo': None, 'Fine_Effettiva': None, 
        'BAC_Budget': 0.0, '%_Completamento': 0.0, 
        'AC_Costo_Reale': 0.0, 'ID_OBS_Assegnato': None, 'Predecessori': ''
    }])
    
if 'obs_data' not in st.session_state:
    st.session_state.obs_data = pd.DataFrame(columns=[
        'ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note'
    ])
    
if 'registro_data' not in st.session_state:
    st.session_state.registro_data = pd.DataFrame(columns=[
        'Data', 'N_Doc', 'Fornitore', 'Voce_WBS', 'Importo_Netto', 'Descrizione'
    ])

if 'capa_data' not in st.session_state:
    st.session_state.capa_data = pd.DataFrame(columns=[
        'Data_Apertura', 'ID_WBS_Rif', 'Tipo_Azione', 'Descrizione', 'Responsabile_OBS', 'Stato', 'Rischio_Associato'
    ])

if 'rischi_data' not in st.session_state:
    st.session_state.rischi_data = pd.DataFrame(columns=[
        'ID_WBS_Rif', 'Descrizione_Rischio', 'Probabilità (1-5)', 'Impatto (1-5)', 'Stato'
    ])

if 'sal_data' not in st.session_state:
    st.session_state.sal_data = pd.DataFrame(columns=['Data_Emissione', 'Descrizione_SAL', 'Importo_Euro', 'Stato_Pagamento'])

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
        df = df[~mask].reset_index(drop=True) 
        
        if df.empty:
            df = pd.DataFrame([{'ID_WBS': '1', 'Attività': 'Progetto Principale', 'BAC_Budget': 0.0, '%_Completamento': 0.0, 'AC_Costo_Reale': 0.0, 'Livello': 1}])
            st.session_state.wbs_data = df.drop(columns=['Livello'])
            st.session_state['tracker_id'] = None
            for k in list(st.session_state.keys()):
                if k.startswith("editor_wbs_"): 
                    del st.session_state[k]
            st.rerun()
            
    elif azione in ['su', 'giu', 'destra', 'sinistra']:
        ids = df['ID_WBS'].astype(str).tolist()
        if id_target not in ids: return
        idx = ids.index(id_target)
        livello_target = df.at[idx, 'Livello']
        
        end_idx = idx + 1
        while end_idx < len(df) and df.at[end_idx, 'Livello'] > livello_target:
            end_idx += 1
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
    if 'Predecessori' in df.columns:
        df['Predecessori'] = df['Predecessori'].apply(aggiorna_preds)
        
    df = df.drop(columns=['Livello', 'sort_key'])
    
    st.session_state.wbs_data = df.copy()
    st.session_state.wbs_data = aggiorna_gerarchia(st.session_state.wbs_data)
    
    if azione != 'elimina':
        st.session_state['tracker_id'] = mapping.get(id_target, id_target)
    else:
        st.session_state['tracker_id'] = None
        
    for k in list(st.session_state.keys()):
        if k.startswith("editor_wbs_"):
            del st.session_state[k]
            
    st.rerun()
    
def get_foglie(df):
    ids = df['ID_WBS'].astype(str).tolist()
    foglie = [uid for uid in ids if not any(other.startswith(uid + '.') for other in ids if other != uid)]
    return df[df['ID_WBS'].astype(str).isin(foglie)].copy()

def aggiorna_costi_reali():
    df_reg = st.session_state.registro_data.copy()
    if not df_reg.empty:
        df_reg['ID_WBS_calc'] = df_reg['Voce_WBS'].astype(str).apply(
            lambda x: str(x).split(' - ')[0].strip() if pd.notna(x) and str(x).strip() not in ['', 'None', 'nan'] else None
        )
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
    if not df_reg_calc.empty:
        df_reg_calc['Data'] = pd.to_datetime(df_reg_calc['Data'], errors='coerce').dt.date
        ac_daily = df_reg_calc.groupby('Data')['Importo_Netto'].sum().to_dict()
    else:
        ac_daily = {}
    
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
        
        cpm_nodes[node_id] = {
            'durata': durata, 'preds': preds, 'succs': [],
            'ES': 0, 'EF': 0, 'LS': 0, 'LF': 0, 'slack': 0, 'is_critical': False
        }
        
    for n_id, data in cpm_nodes.items():
        for p_id in data['preds']:
            if p_id in cpm_nodes:
                cpm_nodes[p_id]['succs'].append(n_id)
                
    changed = True
    loop_counter = 0
    while changed and loop_counter < 1000:
        loop_counter += 1
        changed = False
        for n_id, data in cpm_nodes.items():
            max_ef = 0
            for p_id in data['preds']:
                if p_id in cpm_nodes:
                    max_ef = max(max_ef, cpm_nodes[p_id]['EF'])
            new_es = max_ef
            new_ef = new_es + data['durata']
            if new_es != data['ES'] or new_ef != data['EF']:
                data['ES'] = new_es; data['EF'] = new_ef; changed = True
                
    project_duration = max([data['EF'] for data in cpm_nodes.values()], default=0)
    for n_id, data in cpm_nodes.items():
        data['LF'] = project_duration
        data['LS'] = data['LF'] - data['durata']
        
    changed = True
    loop_counter = 0
    while changed and loop_counter < 1000:
        loop_counter += 1
        changed = False
        for n_id, data in cpm_nodes.items():
            min_ls = data['LF'] 
            if len(data['succs']) > 0:
                min_ls = min([cpm_nodes[s_id]['LS'] for s_id in data['succs'] if s_id in cpm_nodes])
            new_lf = min_ls
            new_ls = new_lf - data['durata']
            if new_lf != data['LF'] or new_ls != data['LS']:
                data['LF'] = new_lf; data['LS'] = new_ls; changed = True
                
    for n_id, data in cpm_nodes.items():
        data['slack'] = data['LS'] - data['ES']
        if data['slack'] <= 0:
            data['is_critical'] = True
            
    return cpm_nodes

# --- 3. ESECUZIONE CALCOLI INIZIALI ---
aggiorna_costi_reali()
st.session_state.wbs_data = aggiorna_gerarchia(st.session_state.wbs_data)
st.session_state.wbs_data = calcola_evm(st.session_state.wbs_data, pd.Timestamp.today().date())

# ==========================================
# IMPOSTAZIONE GRAFICA E LARGHEZZA DINAMICA
# ==========================================

# 1. Creiamo una "memoria" per la larghezza del pannello (default 1.2)
if 'pannello_sx' not in st.session_state:
    st.session_state.pannello_sx = 1.2

# 2. Le colonne ora "respirano" in base al valore del cursore
col_save, col_sviluppo = st.columns([st.session_state.pannello_sx, 10])

# ==========================================
# COLONNA DI SINISTRA (PANNELLO DI CONTROLLO)
# ==========================================
with col_save:
    # 3. Il cursore che comanda la larghezza in tempo reale!
    st.slider("↔️ Regola Pannello", min_value=1.0, max_value=5.0, step=0.1, key="pannello_sx", help="Trascina per allargare o restringere la colonna")

# ==========================================
# COLONNA DI SINISTRA (PANNELLO DI CONTROLLO)
# ==========================================
with col_save:
    st.caption("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;PROGETTO")
    
    st.session_state.nome_progetto_attivo = st.text_area("Nome Progetto", value=st.session_state.nome_progetto_attivo, label_visibility="collapsed", height=100)
    
    st.markdown('<div class="btn-compatto">', unsafe_allow_html=True)
    
    # --- 1. MEMORIA DI SESSIONE ---
    
    if st.button("💾 Salva", use_container_width=True):
        st.session_state.archivio_progetti[st.session_state.nome_progetto_attivo] = {
            "wbs": st.session_state.wbs_data.copy(),
            "obs": st.session_state.obs_data.copy(),
            "registro": st.session_state.registro_data.copy(),
            "capa": st.session_state.capa_data.copy(),
            "sal": st.session_state.sal_data.copy(),
            "conflitti_ignorati": st.session_state.conflitti_ignorati.copy()
        }
        st.success("Salvato!")
        
    if st.button("📑 Duplica", use_container_width=True):
        nuovo_nome = f"{st.session_state.nome_progetto_attivo}_Copia"
        st.session_state.archivio_progetti[nuovo_nome] = {
            "wbs": st.session_state.wbs_data.copy(),
            "obs": st.session_state.obs_data.copy(),
            "registro": st.session_state.registro_data.copy(),
            "capa": st.session_state.capa_data.copy(),
            "sal": st.session_state.sal_data.copy()
        }
        st.session_state.nome_progetto_attivo = nuovo_nome
        st.rerun()

    if st.session_state.archivio_progetti:
        prog_selezionato = st.selectbox("Apri Progetto", options=list(st.session_state.archivio_progetti.keys()), label_visibility="collapsed")
        if st.button("📂 Apri Selezionato", use_container_width=True):
            st.session_state.wbs_data = st.session_state.archivio_progetti[prog_selezionato]["wbs"].copy()
            st.session_state.obs_data = st.session_state.archivio_progetti[prog_selezionato]["obs"].copy()
            st.session_state.registro_data = st.session_state.archivio_progetti[prog_selezionato]["registro"].copy()
            st.session_state.capa_data = st.session_state.archivio_progetti[prog_selezionato]["capa"].copy()
            st.session_state.nome_progetto_attivo = prog_selezionato
            for k in list(st.session_state.keys()):
                if k.startswith("editor_wbs_"):
                    del st.session_state[k]
            st.rerun()

    if st.button("📄 Nuovo", use_container_width=True):
        st.session_state.nome_progetto_attivo = "Nuovo_Progetto"
        for key in ['wbs_data', 'obs_data', 'registro_data', 'capa_data']:
            if key in st.session_state:
                del st.session_state[key]
        for k in list(st.session_state.keys()):
            if k.startswith("editor_wbs_"):
                del st.session_state[k]
        st.rerun()
        
    # --- 2. ARCHIVIAZIONE SU PC (JSON) ---
    st.divider()
    st.caption("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ARCHIVIO PC")
    
    try:
        progetto_export = {
            "wbs": json.loads(st.session_state.wbs_data.to_json(orient="records", date_format="iso")),
            "obs": json.loads(st.session_state.obs_data.to_json(orient="records")),
            "registro": json.loads(st.session_state.registro_data.to_json(orient="records", date_format="iso")),
            "capa": json.loads(st.session_state.capa_data.to_json(orient="records", date_format="iso")),
            "rischi": json.loads(st.session_state.rischi_data.to_json(orient="records")),
            "sal": json.loads(st.session_state.sal_data.to_json(orient="records", date_format="iso")),
            "conflitti_ignorati": list(st.session_state.conflitti_ignorati),
            
        }
        json_string = json.dumps(progetto_export, indent=4)
        
        st.download_button(
            label="⬇️ Scarica",
            data=json_string,
            file_name=f"{st.session_state.nome_progetto_attivo}.json",
            mime="application/json",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Errore: {e}")
    
    uploaded_file = st.file_uploader("📤 Carica da PC", type=['json'], label_visibility="collapsed")
    if uploaded_file is not None:
        if 'ultimo_file_caricato' not in st.session_state or st.session_state.ultimo_file_caricato != uploaded_file.file_id:
            try:
                dati_caricati = json.load(uploaded_file)
                
                df_wbs = pd.DataFrame(dati_caricati.get('wbs', []))
                if df_wbs.empty:
                    df_wbs = pd.DataFrame([{'ID_WBS': '1', 'Attività': 'Progetto Principale', 'BAC_Budget': 0.0, '%_Completamento': 0.0, 'AC_Costo_Reale': 0.0, 'ID_OBS_Assegnato': None, 'Predecessori': ''}])
                st.session_state.wbs_data = df_wbs
                
                df_obs = pd.DataFrame(dati_caricati.get('obs', []))
                if df_obs.empty:
                    df_obs = pd.DataFrame(columns=['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note'])
                st.session_state.obs_data = df_obs
                
                df_reg = pd.DataFrame(dati_caricati.get('registro', []))
                if df_reg.empty:
                    df_reg = pd.DataFrame(columns=['Data', 'N_Doc', 'Fornitore', 'Voce_WBS', 'Importo_Netto', 'Descrizione'])
                st.session_state.registro_data = df_reg
                
                df_capa = pd.DataFrame(dati_caricati.get('capa', []))
                if df_capa.empty:
                    df_capa = pd.DataFrame(columns=['Data_Apertura', 'ID_WBS_Rif', 'Tipo_Azione', 'Descrizione', 'Responsabile_OBS', 'Stato', 'Rischio_Associato'])
                st.session_state.capa_data = df_capa

                df_rischi = pd.DataFrame(dati_caricati.get('rischi', []))
                if df_rischi.empty:
                    df_rischi = pd.DataFrame(columns=['ID_WBS_Rif', 'Descrizione_Rischio', 'Probabilità (1-5)', 'Impatto (1-5)', 'Stato'])
                st.session_state.rischi_data = df_rischi

                df_sal = pd.DataFrame(dati_caricati.get('sal', []))
                if df_sal.empty:
                    df_sal = pd.DataFrame(columns=['Data_Emissione', 'Descrizione_SAL', 'Importo_Euro', 'Stato_Pagamento'])
                st.session_state.sal_data = df_sal

                st.session_state.conflitti_ignorati = dati_caricati.get('conflitti_ignorati', [])
                
                for col in ['Inizio_Previsto', 'Fine_Prevista', 'Inizio_Effettivo', 'Fine_Effettiva']:
                    if col in st.session_state.wbs_data.columns:
                        st.session_state.wbs_data[col] = pd.to_datetime(st.session_state.wbs_data[col], errors='coerce').dt.date
                if 'Data' in st.session_state.registro_data.columns:
                    st.session_state.registro_data['Data'] = pd.to_datetime(st.session_state.registro_data['Data'], errors='coerce').dt.date
                if 'Data_Apertura' in st.session_state.capa_data.columns:
                    st.session_state.capa_data['Data_Apertura'] = pd.to_datetime(st.session_state.capa_data['Data_Apertura'], errors='coerce').dt.date
                
                st.session_state.wbs_data = aggiorna_gerarchia(st.session_state.wbs_data)
                st.session_state.nome_progetto_attivo = uploaded_file.name.replace(".json", "")
                st.session_state.ultimo_file_caricato = uploaded_file.file_id
                
                for k in list(st.session_state.keys()):
                    if k.startswith("editor_wbs_"):
                        del st.session_state[k]
                st.rerun() 
            except Exception as e:
                st.error(f"Errore critico durante la lettura: {e}")
                
    st.markdown('</div>', unsafe_allow_html=True)
    st.divider()
    st.caption("Versione 1.0")

   # ==========================================
    # 🤖 GIANFRANCO SUGGERISCE 🤖
    # ==========================================
    st.divider()
    st.markdown("#### 🤖 GIANFRANCO SUGGERISCE")
    
    # Inizializziamo la memoria delle eccezioni consentite dal DL
    if 'conflitti_ignorati' not in st.session_state:
        st.session_state.conflitti_ignorati = []
    
    df_ai = get_foglie(st.session_state.wbs_data).copy()
    soglia_allerta = 0.95
    
    if df_ai.empty or df_ai['BAC_Budget'].sum() == 0:
        st.info("In attesa di dati.")
    else:
        # Rilevamento criticità EVM
        critici_costo = df_ai[df_ai['CPI'] < soglia_allerta]
        critici_tempo = df_ai[df_ai['SPI'] < soglia_allerta]
        
        # Rilevamento Rischi Attivi
        df_rischi_ai = st.session_state.rischi_data
        rischi_attivi = pd.DataFrame()
        if not df_rischi_ai.empty:
            rischi_attivi = df_rischi_ai[df_rischi_ai['Stato'].isin(['Attivo ▾', 'Monitorato ▾'])]
            
        # ===================================================
        # RILEVAMENTO CONFLITTI RISORSE (CON "IGNORE BUTTON")
        # ===================================================
        conflitti_risorse = []
        df_res = df_ai.dropna(subset=['Inizio_Previsto', 'Fine_Prevista']).copy()
        df_res = df_res[df_res['ID_OBS_Assegnato'].astype(str).str.strip().astype(bool)]
        df_res = df_res[~df_res['ID_OBS_Assegnato'].astype(str).isin(['None', 'nan'])]
        
        if not df_res.empty:
            df_res['Inizio_Previsto'] = pd.to_datetime(df_res['Inizio_Previsto'])
            df_res['Fine_Prevista'] = pd.to_datetime(df_res['Fine_Prevista'])
            
            for obs_val, group in df_res.groupby('ID_OBS_Assegnato'):
                tasks = group.to_dict('records')
                for i in range(len(tasks)):
                    for j in range(i + 1, len(tasks)):
                        t1, t2 = tasks[i], tasks[j]
                        inizio1, fine1 = t1['Inizio_Previsto'], t1['Fine_Prevista']
                        inizio2, fine2 = t2['Inizio_Previsto'], t2['Fine_Prevista']
                        
                        if (inizio1 <= fine2) and (inizio2 <= fine1):
                            # Creiamo una "Targa" univoca per questa sovrapposizione (es. "2.1_2.2")
                            conflitto_id = f"{t1['ID_WBS']}_{t2['ID_WBS']}"
                            
                            # Aggiungiamo all'allarme SOLO se non è stato ignorato dal DL
                            if conflitto_id not in st.session_state.conflitti_ignorati:
                                nome_risorsa = str(obs_val).split(' - ')[1] if ' - ' in str(obs_val) else str(obs_val)
                                sovrapposizione_inizio = max(inizio1, inizio2).strftime('%d/%m')
                                sovrapposizione_fine = min(fine1, fine2).strftime('%d/%m')
                                
                                conflitti_risorse.append({
                                    'ID_Univoco': conflitto_id,
                                    'Risorsa': nome_risorsa,
                                    'WBS1': f"{t1['ID_WBS']} ({t1['Attività']})",
                                    'WBS2': f"{t2['ID_WBS']} ({t2['Attività']})",
                                    'Date': f"dal {sovrapposizione_inizio} al {sovrapposizione_fine}"
                                })

        # --- GESTIONE DEGLI ALLARMI ---
        if critici_costo.empty and critici_tempo.empty and rischi_attivi.empty and not conflitti_risorse:
            st.success("✅ **Progetto in salute!** Nessuna azione richiesta.")
        else:
            # 1. Allarmi di Schedulazione (Tempi)
            if not critici_tempo.empty:
                for _, row in critici_tempo.iterrows():
                    with st.expander(f"⏳ WBS {row['ID_WBS']}: Ritardo"):
                        st.markdown(f"**Situazione:** Efficienza tempi al **{row['SPI']*100:.0f}%**.<br>**Azione:** Usa il Simulatore (*Tab 7*) per testare un *Crashing* oppure scala le date in *Tab 1*.", unsafe_allow_html=True)
            
            # 2. Allarmi Finanziari (Costi)
            if not critici_costo.empty:
                for _, row in critici_costo.iterrows():
                    with st.expander(f"💸 WBS {row['ID_WBS']}: Over-Budget"):
                        st.markdown(f"**Situazione:** Efficienza costi a **{row['CPI']:.2f}**.<br>**Azione:** Verifica le spese nel *Tab 6*. Se la spesa è irreversibile, apri una CAPA in *Tab 7*.", unsafe_allow_html=True)
                        
            # 3. Allarmi Rischio (Scudo)
            if not rischi_attivi.empty:
                with st.expander(f"⚠️ {len(rischi_attivi)} Rischi Attivi"):
                    st.markdown("**Situazione:** Capitale congelato nel Fondo Imprevisti.<br>**Azione:** Metti in atto le mitigazioni di cantiere e chiudi l'Azione in *Tab 7* per sbloccare i fondi.", unsafe_allow_html=True)
                    
            # 4. Allarmi Sovraccarico Risorse (Conflitti)
            if conflitti_risorse:
                with st.expander(f"👷 {len(conflitti_risorse)} Conflitti Risorse"):
                    st.markdown("**Situazione:** Una risorsa è stata assegnata a lavorazioni sovrapposte.")
                    for c in conflitti_risorse:
                        st.error(f"**{c['Risorsa']}** lavora su:\n* {c['WBS1']}\n* {c['WBS2']}\n*(Accavallamento: {c['Date']})*")
                        # IL BOTTONE PER IGNORARE!
                        if st.button("👁️ Consenti Sovrapposizione", key=f"ignora_{c['ID_Univoco']}", help="Nascondi questo allarme: accetto che lavorino in contemporanea."):
                            st.session_state.conflitti_ignorati.append(c['ID_Univoco'])
                            st.rerun()
                            
            # 5. NUOVO: Registro dei Conflitti Tollerati (Invisibili)
            if st.session_state.conflitti_ignorati:
                with st.expander(f"👁️ {len(st.session_state.conflitti_ignorati)} Conflitti Tollerati (Nascosti)"):
                    st.markdown("Hai autorizzato le seguenti sovrapposizioni:")
                    for conf_id in st.session_state.conflitti_ignorati:
                        wbs_split = conf_id.split('_')
                        st.markdown(f"- WBS **{wbs_split[0]}** + WBS **{wbs_split[1]}**")
                        if st.button("🔄 Ripristina Allarme", key=f"ripristina_{conf_id}"):
                            st.session_state.conflitti_ignorati.remove(conf_id)
                            st.rerun()


# ==========================================
# COLONNA DI DESTRA (IL MOTORE DELL'APP)
# ==========================================
with col_sviluppo:

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

    # --- CREAZIONE TAB ---
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
        "🗂️ 1-WBS (Lavorazioni)", 
        "👥 2-OBS (Risorse)", 
        "🕸️ 3-Nodi & Matrice", 
        "📅 4-Cronoprogramma", 
        "📈 5-Earned Value & Cash Flow",
        "🧾 6-Reg. Contabile",
        "🛠️ 7-Direzione & CAPA",
        "⚠️ 8-Matrice Rischi",
        "📚 9-Guida & Glossario"
    ])
        
    # --- TAB 1: SETUP WBS ---
    with tab1:
        st.header("WBS - Work Breakdown Structure")
        st.markdown('*I numeri ID sono bloccati per garantire l\'integrità. Usa i pulsanti sotto ogni capitolo per spostare e rientrare le voci.*')
        
        st.subheader("🔀 Organizzatore Capitoli")
        c_sel, c_btn1, c_btn2, c_btn3 = st.columns([3, 1, 1, 1])
        
        df_padri = st.session_state.wbs_data[~st.session_state.wbs_data['ID_WBS'].astype(str).str.contains(r'\.')]
        lista_capitoli = list(df_padri['ID_WBS'].astype(str) + " - " + df_padri['Attività'].astype(str))
        
        # Inseguimento voce Capitolo
        tracker = st.session_state.get('tracker_id', None)
        idx_padre = 0
        if tracker:
            for i, opz in enumerate(lista_capitoli):
                if opz.split(' - ')[0] == tracker:
                    idx_padre = i
                    break
                    
        nodo_scelto = c_sel.selectbox("Seleziona Capitolo", options=lista_capitoli, index=idx_padre, label_visibility="collapsed")
        
        if nodo_scelto:
            id_scelto = nodo_scelto.split(' - ')[0]
            if c_btn1.button("⬆️ Su", use_container_width=True): modifica_struttura(id_scelto, 'su')
            if c_btn2.button("⬇️ Giù", use_container_width=True): modifica_struttura(id_scelto, 'giu')
            if c_btn3.button("🗑️ Elimina", use_container_width=True): modifica_struttura(id_scelto, 'elimina')
                
        st.divider()
        
        df = st.session_state.wbs_data.copy()
        
        for col in ['Inizio_Previsto', 'Fine_Prevista', 'Inizio_Effettivo', 'Fine_Effettiva']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
                
        df['Durata_Prevista (gg)'] = (pd.to_datetime(df['Fine_Prevista']) - pd.to_datetime(df['Inizio_Previsto'])).dt.days + 1
        
        radici = df[~df['ID_WBS'].astype(str).str.contains(r'\.')]
        df_aggiornato = pd.DataFrame()
        
        lista_obs_dropdown = [""] + [f"{row['ID_OBS']} - {row['Risorsa']}" for _, row in st.session_state.obs_data.iterrows() if pd.notna(row['ID_OBS'])]
        lista_wbs_dropdown = [""] + [f"{row['ID_WBS']} - {row['Attività']}" for _, row in df.iterrows() if pd.notna(row['ID_WBS'])]
        
        for idx_riga, radice in radici.iterrows():
            id_radice = str(radice['ID_WBS'])
            discendenti = df[df['ID_WBS'].astype(str).str.startswith(f"{id_radice}.")]
            tot_budget = radice['BAC_Budget']
            
            with st.expander(f"📁 {id_radice} - {radice['Attività']} (Budget Totale Raggruppato: € {tot_budget:,.2f})", expanded=True):
                
                colonne_bloccate = ["ID_WBS", "Durata_Prevista (gg)", "AC_Costo_Reale", "PV", "EV", "CV", "SV", "SPI", "CPI", "EAC", "ETC", "VAC"]
                colonne_bloccate = [col for col in colonne_bloccate if col in discendenti.columns]
                
                discendenti_modificati = st.data_editor(
                    discendenti,
                    key=f"editor_wbs_idx_{idx_riga}_id_{id_radice}",
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    disabled=colonne_bloccate, 
                    column_config={
                        "ID_WBS": st.column_config.TextColumn("ID WBS (Auto)"),
                        "Predecessori": st.column_config.SelectboxColumn("Predecessore ▾", options=lista_wbs_dropdown),
                        "ID_OBS_Assegnato": st.column_config.SelectboxColumn("Risorsa Assegnata ▾", options=lista_obs_dropdown),
                        "Inizio_Previsto": st.column_config.DateColumn("Inizio Previsto"),
                        "Fine_Prevista": st.column_config.DateColumn("Fine Prevista"),
                        "Inizio_Effettivo": st.column_config.DateColumn("Inizio Effettivo"),
                        "Fine_Effettiva": st.column_config.DateColumn("Fine Effettiva")
                    }
                )
                
                for i_row, row_mod in discendenti_modificati.iterrows():
                    val_id = str(row_mod['ID_WBS']).strip()
                    if val_id in ['', 'None', 'nan']:
                        discendenti_modificati.at[i_row, 'ID_WBS'] = f"{id_radice}.999{i_row}"
                
                df_aggiornato = pd.concat([df_aggiornato, pd.DataFrame([radice]), discendenti_modificati], ignore_index=True)
                
                if not discendenti.empty:
                    st.markdown("↕️ **Sposta / Modifica Livello:**")
                    c_sel_int, c1, c2, c3, c4 = st.columns([3, 1.5, 1.5, 1, 1])
                    
                    opzioni_locali = list(discendenti['ID_WBS'].astype(str) + " - " + discendenti['Attività'].astype(str))
                    
                    # Inseguimento voce Sotto-Capitolo
                    idx_locale = 0
                    if tracker:
                        for i, opz in enumerate(opzioni_locali):
                            if opz.split(' - ')[0] == tracker:
                                idx_locale = i
                                break
                                
                    nodo_locale = c_sel_int.selectbox("Seleziona voce da muovere", options=opzioni_locali, index=idx_locale, key=f"sel_move_{id_radice}", label_visibility="collapsed")
                    
                    if nodo_locale:
                        id_loc = nodo_locale.split(' - ')[0]
                        if c1.button("⬅️ Rendi Padre", key=f"l_{id_radice}", use_container_width=True): modifica_struttura(id_loc, 'sinistra')
                        if c2.button("➡️ Rendi Figlio", key=f"r_{id_radice}", use_container_width=True): modifica_struttura(id_loc, 'destra')
                        if c3.button("⬆️ Su", key=f"u_{id_radice}", use_container_width=True): modifica_struttura(id_loc, 'su')
                        if c4.button("⬇️ Giù", key=f"d_{id_radice}", use_container_width=True): modifica_struttura(id_loc, 'giu')

        with st.form("aggiungi_padre"):
            st.write("Aggiungi un nuovo Capitolo Principale")
            c1, c2 = st.columns([4, 1])
            nuova_att = c1.text_input("Nome", placeholder="Es. Impianti Elettrici")
            if c2.form_submit_button("➕ Aggiungi Capitolo"):
                if nuova_att:
                    is_root_calc = ~st.session_state.wbs_data['ID_WBS'].astype(str).str.contains(r'\.')
                    nuovo_id = str(len(st.session_state.wbs_data[is_root_calc]) + 1)
                    nuova_riga = pd.DataFrame([{'ID_WBS': nuovo_id, 'Attività': nuova_att, 'BAC_Budget': 0.0, '%_Completamento': 0.0, 'AC_Costo_Reale': 0.0}])
                    st.session_state.wbs_data = pd.concat([st.session_state.wbs_data, nuova_riga], ignore_index=True)
                    modifica_struttura('1', 'rinumera') 

        st.divider()
        st.warning("⚠️ **Hai aggiunto nuove lavorazioni nelle tabelle?** Clicca il tasto qui sotto per far assegnare al sistema la numerazione definitiva e riallineare l'albero WBS.")
        if st.button("💾 SALVA INSERIMENTI E RICALCOLA ALBERO", type="primary", use_container_width=True, key="btn_salva_mega_wbs"):
            
            # --- 1. INIZIO INNESTO CAPA (Agisce su df_aggiornato) ---
            df_capa_check = st.session_state.capa_data
            wbs_bloccate = []
            if not df_capa_check.empty and 'ID_WBS_Rif' in df_capa_check.columns and 'Stato' in df_capa_check.columns:
                capa_attive = df_capa_check[df_capa_check['Stato'].isin(['Aperto ▾', 'In Lavorazione ▾'])]
                if not capa_attive.empty:
                    wbs_bloccate = capa_attive['ID_WBS_Rif'].astype(str).apply(lambda x: x.split(' - ')[0].strip()).unique().tolist()
            
            allarmi_blocco = []
            for idx, row in df_aggiornato.iterrows():
                wbs_id = str(row.get('ID_WBS', '')).strip()
                try:
                    completamento = float(row.get('%_Completamento', 0))
                except:
                    completamento = 0.0
                    
                if wbs_id in wbs_bloccate and completamento >= 100:
                    df_aggiornato.at[idx, '%_Completamento'] = 99.0
                    allarmi_blocco.append(wbs_id)
            # --- FINE INNESTO CAPA ---

            # 2. Salvataggio del dataframe corretto
            st.session_state.wbs_data = df_aggiornato
            
            # 3. Pulizia della cache di Streamlit
            for k in list(st.session_state.keys()):
                if k.startswith("editor_wbs_"):
                    del st.session_state[k]
                    
            # 4. Ricalcolo struttura
            modifica_struttura('1', 'rinumera')
            
            # 5. Feedback a schermo
            if allarmi_blocco:
                st.error(f"🚧 BLOCCO QUALITÀ: WBS {', '.join(allarmi_blocco)} bloccate al 99% per CAPA aperte.")
            else:
                st.success("✅ Dati salvati e albero ricalcolato!")
                
            st.rerun()

    # --- TAB 2: SETUP OBS ---
    with tab2:
        st.header("OBS - Organization Breakdown Structure")
        
        with st.expander("⚙️ Gestione Colonne Aggiuntive", expanded=False):
            c1, c2 = st.columns([3, 1])
            nuova_col = c1.text_input("Nome nuova colonna (es. Telefono, Qualifica, Partita IVA)")
            if c2.button("➕ Crea") and nuova_col:
                if nuova_col not in st.session_state.obs_data.columns:
                    st.session_state.obs_data[nuova_col] = "" 
                    st.rerun()
            
            colonne_custom = [c for c in st.session_state.obs_data.columns if c not in ['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note']]
            if colonne_custom:
                st.divider()
                c3, c4, c5 = st.columns([2, 2, 1])
                col_da_modificare = c3.selectbox("Colonna da rinominare", options=colonne_custom)
                nuovo_nome = c4.text_input("Nuovo nome intestazione")
                if c5.button("✏️ Modifica") and nuovo_nome:
                    st.session_state.obs_data.rename(columns={col_da_modificare: nuovo_nome}, inplace=True)
                    st.rerun()

        edited_obs = st.data_editor(
            st.session_state.obs_data, 
            column_config={
                "Tipo_Contratto": st.column_config.SelectboxColumn("Tipo Contratto", options=["Appalto ▾", "Sub appalto ▾"])
            },
            num_rows="dynamic", use_container_width=True, hide_index=True
        )
        
        st.divider()
        if st.button("💾 CONFERMA E SALVA ANAGRAFICA (OBS)", type="primary", use_container_width=True):
            st.session_state.obs_data = edited_obs
            st.success("✅ Dati anagrafici salvati con successo!")
            st.rerun()

    # --- TAB 3: MATRICE E GRAFO A NODI ---
    with tab3:
        st.header("Percorso Logico - Work Packages e Percorso Critico")
        
        try:
            cpm_data = calcola_cpm(st.session_state.wbs_data)
            mostra_relazioni = st.toggle("👁️ Mostra Percorso Critico (CPM)", value=True)
            
            graph = graphviz.Digraph(engine='dot')
            graph.attr(rankdir='LR', ranksep='1.5', nodesep='0.8', splines='spline')
            graph.attr('node', fontname='Helvetica', fontsize='10', margin='0.2')
            
            # Scudo totale contro caratteri che fanno impazzire il motore grafico
            def pulisci_testo(testo):
                if pd.isna(testo) or testo is None: return ""
                t = str(testo)
                t = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
                t = t.replace('\n', '<BR/>') # Questo impedisce i crash se premi "invio" nelle descrizioni
                return t
                
            for _, row in st.session_state.obs_data.iterrows():
                ruolo = pulisci_testo(row.get('Ruolo', ''))
                risorsa = pulisci_testo(row.get('Risorsa', ''))
                
                if not ruolo and not risorsa: continue
                
                label_html = f"<<TABLE BORDER='0' CELLBORDER='0' CELLSPACING='2'>"
                label_html += f"<TR><TD><B>{ruolo}</B></TD></TR>"
                label_html += f"<TR><TD>({risorsa})</TD></TR>"
                
                colonne_custom = [col for col in st.session_state.obs_data.columns if col not in ['ID_OBS', 'Ruolo', 'Risorsa', 'Tipo_Contratto', 'Note']]
                for col in colonne_custom:
                    valore = pulisci_testo(row.get(col, ''))
                    if pd.notna(row.get(col)) and valore.strip() != "":
                        col_safe = pulisci_testo(col)
                        label_html += f"<TR><TD><FONT POINT-SIZE='9' COLOR='gray30'>{col_safe}: {valore}</FONT></TD></TR>"
                label_html += "</TABLE>>"
                
                obs_id = str(row.get('ID_OBS', '')).strip()
                if obs_id.endswith('.0'): obs_id = obs_id[:-2]
                
                if obs_id:
                    graph.node(f"OBS_{obs_id}", label=label_html, shape='rect', style='rounded,filled', fillcolor='#E1F5FE', color='#0288D1', penwidth='1.5')
                
            df_wp_reali = get_foglie(st.session_state.wbs_data)
            valid_wbs_ids = set(df_wp_reali['ID_WBS'].astype(str))
            
            if df_wp_reali.empty or len(valid_wbs_ids) == 0:
                graph.node("Vuoto", label="Nessuna lavorazione inserita.", shape="rect")
                
            # --- AGGIORNAMENTO GRAFO: INTEGRAZIONE RISCHI ---
            df_rischi = st.session_state.rischi_data
            
            # --- AGGIORNAMENTO GRAFO: INTEGRAZIONE RISCHI ---
            df_rischi = st.session_state.rischi_data
            
            for _, row in df_wp_reali.iterrows():
                wbs_id = str(row.get('ID_WBS', '')).strip()
                if not wbs_id or wbs_id in ['nan', 'None']: continue
                
                attivita = pulisci_testo(row.get('Attività', ''))
                
                # 1. Base Styling del CPM (Percorso Critico Temporale)
                wp_cpm = cpm_data.get(wbs_id, {})
                is_critical = wp_cpm.get('is_critical', False)
                margine = wp_cpm.get('slack', 0)
                
                if is_critical:
                    testo_margine = f"<FONT COLOR='#D32F2F'><B>Margine: {margine} gg</B></FONT>"
                    bordo_colore, spessore_bordo = '#D32F2F', '3.0'
                else:
                    testo_margine = f"<FONT COLOR='#388E3C'>Margine: {margine} gg</FONT>"
                    bordo_colore, spessore_bordo = '#388E3C', '1.5'

                # 2. Sovrascrittura Rischio (SOLO se Attivo o Monitorato)
                rischi_wbs = df_rischi[df_rischi['ID_WBS_Rif'].astype(str).str.startswith(wbs_id + " -")]
                rischi_attivi = rischi_wbs[rischi_wbs['Stato'].isin(['Attivo ▾', 'Monitorato ▾'])]
                
                punteggio_max = 0
                if not rischi_attivi.empty:
                    prob = pd.to_numeric(rischi_attivi['Probabilità (1-5)'], errors='coerce').fillna(0)
                    imp = pd.to_numeric(rischi_attivi['Impatto (1-5)'], errors='coerce').fillna(0)
                    punteggio_max = (prob * imp).max()

                # 3. Applicazione Alert Visivo
                tag_rischio = ""
                if punteggio_max >= 15:
                    bordo_colore, spessore_bordo = '#FF0000', '4.0' # Rosso Fuoco per emergenze
                    tag_rischio = " <FONT COLOR='#FF0000'><b>[💣 ALTO RISCHIO]</b></FONT>"
                elif punteggio_max >= 8:
                    if not is_critical: 
                        bordo_colore, spessore_bordo = '#FF9800', '2.5' # Arancione
                    tag_rischio = " <FONT COLOR='#FF9800'><b>[⚠️ RISCHIO]</b></FONT>"

                # 4. Parametri classici (Budget, Costi, ecc.)
                try: budget = float(str(row.get('BAC_Budget', 0)).replace(',', '.'))
                except: budget = 0.0
                try: costo_reale = float(str(row.get('AC_Costo_Reale', 0)).replace(',', '.'))
                except: costo_reale = 0.0
                try: completamento = float(str(row.get('%_Completamento', 0)).replace(',', '.'))
                except: completamento = 0.0
                
                inizio_val = pd.to_datetime(row.get('Inizio_Previsto'), errors='coerce')
                inizio_str = inizio_val.strftime('%d/%m/%Y') if pd.notna(inizio_val) else "N/D"
                fine_val = pd.to_datetime(row.get('Fine_Prevista'), errors='coerce')
                fine_str = fine_val.strftime('%d/%m/%Y') if pd.notna(fine_val) else "N/D"
                
                # Assemblaggio Nodo HTML
                wp_html = f"<<TABLE BORDER='0' CELLBORDER='0' CELLSPACING='4'>"
                wp_html += f"<TR><TD COLSPAN='2'><B>{wbs_id} - {attivita}</B>{tag_rischio}</TD></TR>"
                wp_html += f"<TR><TD ALIGN='LEFT'>Inizio: {inizio_str}</TD><TD ALIGN='RIGHT'>Fine: {fine_str}</TD></TR>"
                wp_html += f"<TR><TD ALIGN='LEFT'>Budget: {budget:,.0f} Euro</TD><TD ALIGN='RIGHT'>AC: {costo_reale:,.0f} Euro</TD></TR>"
                wp_html += f"<TR><TD ALIGN='LEFT'>Avanzamento: {completamento:.0f}%</TD><TD ALIGN='RIGHT'>{testo_margine}</TD></TR>"
                wp_html += "</TABLE>>"
                
                # Colore Barra Progressione
                if completamento >= 100:
                    stile, colore_sfondo = 'rounded,filled', '#C8E6C9' 
                elif completamento <= 0:
                    stile, colore_sfondo = 'rounded,filled', 'white'   
                else:
                    stile = 'rounded,striped'
                    quota_verde = max(0.01, min(0.99, completamento / 100.0))
                    colore_sfondo = f"#C8E6C9;{str(round(quota_verde, 3)).replace(',', '.')}:white"
                    
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
                        if p_id in valid_wbs_ids:
                            pred_is_critical = cpm_data.get(p_id, {}).get('is_critical', False)
                            col_cavo, stile_cavo, spes_cavo, freccia = ('#D32F2F', 'solid', '2.5', '1.0') if (is_critical and pred_is_critical) else ('#FF9800', 'dashed', '1.0', '0.6')
                            graph.edge(f"WBS_{p_id}", f"WBS_{wbs_id}", color=col_cavo, style=stile_cavo, penwidth=spes_cavo, arrowsize=freccia)

            # RENDERIZZAZIONE SICURA
            raw_svg = graph.pipe(format='svg').decode('utf-8')
            if '<svg' in raw_svg:
                svg_data = raw_svg[raw_svg.find('<svg'):]
                html_code = f"""
                <!DOCTYPE html><html><head>
                <script src="https://cdn.jsdelivr.net/npm/svg-pan-zoom@3.6.1/dist/svg-pan-zoom.min.js"></script>
                <style>body {{ margin: 0; overflow: hidden; background-color: #fafafa; }} #svg-container {{ width: 100vw; height: 100vh; display: flex; justify-content: center; align-items: center; }} svg {{ width: 100% !important; height: 100% !important; }}</style>
                </head><body><div id="svg-container">{svg_data}</div>
                <script>window.onload = function() {{ 
                    var s = document.querySelector('svg'); 
                    if (s) {{ 
                        s.setAttribute('id', 'grafo'); 
                        s.removeAttribute('width'); 
                        s.removeAttribute('height'); 
                        var panZoom = svgPanZoom('#grafo', {{ 
                            zoomEnabled: true, 
                            controlIconsEnabled: true, 
                            fit: false, 
                            center: true,
                            minZoom: 0.1,
                            maxZoom: 10,
                            mouseWheelZoomEnabled: true
                        }}); 
                    }} 
                }};</script>
                </body></html>
                """
                components.html(html_code, height=900)
            else:
                st.warning("Grafo generato ma vuoto.")
                
        except Exception as e:
            st.error(f"⚠️ ERRORE TECNICO DURANTE IL DISEGNO DEL GRAFO: {e}")
            
        st.divider()
        st.subheader("📖 Legenda del Grafo")
        col_leg1, col_leg2 = st.columns(2)
        with col_leg1:
            st.markdown("**NODI E FIGURE**\n* 🟦 **Riquadro Azzurro:** Risorsa/Ruolo (OBS)\n* 🟩 **Riquadro Verde:** Work Package (WBS)\n* 🟥 **Bordo Rosso Spesso:** Percorso Critico (Margine = 0 gg)")
        with col_leg2:
            st.markdown("**CAVI E COLLEGAMENTI**\n* 🔗 **Freccia Grigia Continua:** Assegnazione Risorsa\n* 🔀 **Freccia Arancione Tratteggiata:** Relazione logica\n* 🚨 **Freccia Rossa Spessa:** Flusso del Percorso Critico")

    # --- TAB 4: CRONOPROGRAMMA (GANTT) ---
    with tab4:
        st.header("Cronoprogramma Lavori (Gantt EVM-Aware)")
        st.markdown("Le barre della fase esecutiva cambiano automaticamente colore in base alle performance reali (Indice SPI).")
        
        c1, c2, c3 = st.columns([1, 1, 1])
        vista = c1.selectbox("Seleziona Vista", ["Progetto (Baseline)", "Esecuzione (Esecutivo)", "Comparativa"])
        mostra_frecce = c2.toggle("🔗 Mostra Frecce Dipendenze", value=True)
        data_status_gantt = c3.date_input("📅 Data Rilevamento", value=pd.Timestamp.today().date())
        
        df_gantt = get_foglie(st.session_state.wbs_data).copy().dropna(subset=['Inizio_Previsto', 'Fine_Prevista'])
        
        if not df_gantt.empty:
            df_gantt['Inizio_Previsto'] = pd.to_datetime(df_gantt['Inizio_Previsto'])
            df_gantt['Fine_Prevista'] = pd.to_datetime(df_gantt['Fine_Prevista'])
            df_gantt['Inizio_Effettivo'] = pd.to_datetime(df_gantt['Inizio_Effettivo'])
            df_gantt['Fine_Effettiva'] = pd.to_datetime(df_gantt['Fine_Effettiva']).fillna(pd.to_datetime(data_status_gantt))
            
            cpm_data = calcola_cpm(st.session_state.wbs_data)
            # Calcoliamo l'EVM in background per avere gli indici di performance aggiornati a questa data
            df_evm_gantt = calcola_evm(get_foglie(st.session_state.wbs_data), data_status_gantt)
            
            fig = go.Figure()
            
            # 1. TRACCIA BASELINE (Azzurra)
            if vista in ["Progetto (Baseline)", "Comparativa"]:
                fig.add_trace(go.Bar(
                    x=(df_gantt['Fine_Prevista'] - df_gantt['Inizio_Previsto'] + pd.Timedelta(days=1)).dt.total_seconds() * 1000, 
                    y=df_gantt['ID_WBS'].astype(str) + " - " + df_gantt['Attività'], 
                    base=df_gantt['Inizio_Previsto'], 
                    orientation='h', 
                    name='Baseline (Pianificato)', 
                    width=0.4, 
                    marker=dict(color='rgba(30, 136, 229, 0.4)' if vista == "Comparativa" else '#1E88E5')
                ))
                
            # 2. TRACCIA ESECUTIVA "PARLANTE"
            if vista in ["Esecuzione (Esecutivo)", "Comparativa"]:
                df_esec = df_gantt.dropna(subset=['Inizio_Effettivo']).copy()
                if not df_esec.empty:
                    # Forziamo il formato testo per evitare mancati incroci
                    df_esec['ID_WBS'] = df_esec['ID_WBS'].astype(str).str.strip()
                    df_evm_clean = df_evm_gantt[['ID_WBS', 'SPI', '%_Completamento']].copy()
                    df_evm_clean['ID_WBS'] = df_evm_clean['ID_WBS'].astype(str).str.strip()
                    
                    # CANCELLIAMO LE VECCHIE COLONNE PER EVITARE I DOPPIONI (_x e _y)
                    df_esec = df_esec.drop(columns=['SPI', '%_Completamento'], errors='ignore')
                    
                    # Ora uniamo i dati puliti
                    df_esec = df_esec.merge(df_evm_clean, on='ID_WBS', how='left')
                    
                    def colora_gantt(row):
                        spi = row['SPI']
                        if pd.isna(spi) or row['%_Completamento'] == 0: return '#9E9E9E' # Grigio se non ancora valutabile
                        if spi >= 1.0: return '#4CAF50' # Verde (In anticipo/Puntuale)
                        if spi >= 0.90: return '#FF9800' # Arancione (Lieve ritardo)
                        return '#D32F2F' # Rosso (Ritardo grave)
                        
                    colori_barre = df_esec.apply(colora_gantt, axis=1).tolist()
                    testi_hover = df_esec['SPI'].apply(lambda x: f"SPI: {x:.2f}" if pd.notna(x) else "").tolist()

                    fig.add_trace(go.Bar(
                        x=(df_esec['Fine_Effettiva'] - df_esec['Inizio_Effettivo'] + pd.Timedelta(days=1)).dt.total_seconds() * 1000, 
                        y=df_esec['ID_WBS'].astype(str) + " - " + df_esec['Attività'], 
                        base=df_esec['Inizio_Effettivo'], 
                        orientation='h', 
                        name='Esecutivo',
                        text=testi_hover,
                        textposition='inside',
                        insidetextanchor='middle',
                        textfont=dict(color='white', size=11, family='Arial Black'),
                        width=0.4 if vista == "Esecuzione (Esecutivo)" else 0.2, 
                        marker=dict(color=colori_barre)
                    ))
                    
                    # Tracce fantasma per creare la legenda dei colori
                    fig.add_trace(go.Scatter(x=[None], y=[None], mode='markers', marker=dict(size=12, color='#4CAF50', symbol='square'), name='🟢 Puntuale/Anticipo (SPI ≥ 1)'))
                    fig.add_trace(go.Scatter(x=[None], y=[None], mode='markers', marker=dict(size=12, color='#FF9800', symbol='square'), name='🟠 Lieve Ritardo (SPI 0.9-1.0)'))
                    fig.add_trace(go.Scatter(x=[None], y=[None], mode='markers', marker=dict(size=12, color='#D32F2F', symbol='square'), name='🔴 Ritardo Grave (SPI < 0.9)'))
                    
            # 3. FRECCE DIPENDENZE
            if mostra_frecce:
                for _, row in df_gantt.iterrows():
                    wbs_id = str(row['ID_WBS']).strip()
                    succ_y = wbs_id + " - " + str(row['Attività'])
                    
                    if vista == "Esecuzione (Esecutivo)" and pd.notna(row['Inizio_Effettivo']):
                        succ_start = row['Inizio_Effettivo']
                    else:
                        succ_start = row['Inizio_Previsto']
                    
                    preds = str(row.get('Predecessori', '')).strip()
                    if preds and preds.lower() not in ['none', 'nan', 'null']:
                        for p in preds.split(','):
                            p_id = p.split(' - ')[0].strip()
                            if p_id.endswith('.0'): p_id = p_id[:-2]
                            
                            pred_row = df_gantt[df_gantt['ID_WBS'].astype(str) == p_id]
                            if not pred_row.empty:
                                pred_y = p_id + " - " + str(pred_row.iloc[0]['Attività'])
                                
                                if vista == "Esecuzione (Esecutivo)" and pd.notna(pred_row.iloc[0]['Fine_Effettiva']):
                                    pred_end = pred_row.iloc[0]['Fine_Effettiva'] + pd.Timedelta(days=1)
                                else:
                                    pred_end = pred_row.iloc[0]['Fine_Prevista'] + pd.Timedelta(days=1)
                                
                                is_critical = cpm_data.get(wbs_id, {}).get('is_critical', False)
                                pred_is_critical = cpm_data.get(p_id, {}).get('is_critical', False)
                                
                                arrow_color = '#D32F2F' if (is_critical and pred_is_critical) else '#9E9E9E'
                                
                                fig.add_annotation(
                                    x=succ_start, y=succ_y, ax=pred_end, ay=pred_y,
                                    xref='x', yref='y', axref='x', ayref='y',
                                    showarrow=True, arrowhead=2, arrowsize=1, arrowwidth=1.5,
                                    arrowcolor=arrow_color, opacity=0.8,
                                    standoff=2, startstandoff=2
                                )
            
            altezza_dinamica = max(500, len(df_gantt) * 45)
            
            fig.update_layout(
                barmode='overlay', 
                height=altezza_dinamica, 
                bargap=0.3, 
                xaxis_title="Linea Temporale", 
                yaxis_title="Lavorazioni (WBS)", 
                yaxis={'autorange': 'reversed'}, 
                xaxis_type='date',
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, bgcolor="rgba(255,255,255,0.8)")
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("⚠️ Il cronoprogramma è vuoto. Inserisci le date di Inizio e Fine nel Tab 1.")
            
    # --- TAB 5: EVM E CASH FLOW ---
    with tab5:
        st.header("Controllo Costi e Analisi EVM")
        
        data_status_evm = st.date_input("📅 Data di Stato (Status Date):", value=pd.Timestamp.today().date())
        df_evm = calcola_evm(get_foglie(st.session_state.wbs_data), data_status_evm)
        
        tot_bac, tot_pv, tot_ev, tot_ac = df_evm['BAC_Budget'].sum(), df_evm['PV'].sum(), df_evm['EV'].sum(), df_evm['AC_Costo_Reale'].sum()
        tot_eac, tot_etc, tot_vac = df_evm['EAC'].sum(), df_evm['ETC'].sum(), df_evm['VAC'].sum()
        cpi_globale = tot_ev / tot_ac if tot_ac > 0 else 1.0
        spi_globale = tot_ev / tot_pv if tot_pv > 0 else 1.0
        perc_completamento = (tot_ev / tot_bac * 100) if tot_bac > 0 else 0.0
        perc_pianificata = (tot_pv / tot_bac * 100) if tot_bac > 0 else 0.0
        
        # --- CALCOLO CONTINGENCY RESERVE (EMV DAI RISCHI) ---
        df_rischi = st.session_state.rischi_data.copy()
        contingency_reserve = 0.0
        
        if not df_rischi.empty:
            # Peschiamo SOLO i rischi attivi o monitorati
            rischi_attivi = df_rischi[df_rischi['Stato'].isin(['Attivo ▾', 'Monitorato ▾'])]
            for _, r_row in rischi_attivi.iterrows():
                wbs_id_full = str(r_row.get('ID_WBS_Rif', ''))
                wbs_id = wbs_id_full.split(' - ')[0].strip()
                
                # Cerchiamo il budget di quella specifica attività
                wbs_match = df_evm[df_evm['ID_WBS'].astype(str) == wbs_id]
                wbs_budget = float(wbs_match.iloc[0]['BAC_Budget']) if not wbs_match.empty else 0.0
                
                prob = pd.to_numeric(r_row['Probabilità (1-5)'], errors='coerce')
                imp = pd.to_numeric(r_row['Impatto (1-5)'], errors='coerce')
                
                if pd.notna(prob) and pd.notna(imp) and wbs_budget > 0:
                    prob_pct = prob * 0.20        # Prob: 1=20%, 5=100%
                    imp_pct = (imp - 1) * 0.05    # Impatto: 1=0%, 5=20% del budget
                    
                    emv = wbs_budget * prob_pct * imp_pct
                    contingency_reserve += emv
                    
        eac_risk_adjusted = tot_eac + contingency_reserve
        
        # --- CRUSCOTTO UI ---
        col_box1, col_box2 = st.columns(2)
        with col_box1:
            with st.container(border=True):
                st.markdown("#### 📊 Stato Attuale (Consuntivo)")
                c1, c2, c3 = st.columns(3)
                c1.metric("Budget Totale (BAC)", f"€ {tot_bac:,.0f}")
                c2.metric("Lavoro Eseguito (EV)", f"€ {tot_ev:,.0f}")
                c3.metric("Costi Sostenuti (AC)", f"€ {tot_ac:,.0f}")
                st.divider()
                c4, c5 = st.columns(2)
                c4.metric("Avanzamento Globale", f"{perc_completamento:.1f}%", delta=f"Pianificato: {perc_pianificata:.1f}%", delta_color="off")
                c5.metric("SPI (Tempi)", f"{spi_globale:.2f}", delta="In ritardo" if spi_globale < 1 else "In anticipo", delta_color="inverse")

        with col_box2:
            with st.container(border=True):
                st.markdown("#### 🔮 Previsioni a Finire (Proiezioni)")
                c1, c2, c3 = st.columns(3)
                c1.metric("Costo Stimato (EAC)", f"€ {tot_eac:,.0f}", delta="Puro EVM matematico", delta_color="off")
                c2.metric("Costo Residuo (ETC)", f"€ {tot_etc:,.0f}", delta="Capitale necessario", delta_color="off")
                c3.metric("Varianza a Finire (VAC)", f"€ {tot_vac:,.0f}", delta="Perdita" if tot_vac < 0 else "Risparmio", delta_color="normal")
                st.divider()
                c4, c5 = st.columns(2)
                c4.metric("CPI (Costi)", f"{cpi_globale:.2f}", delta="Over-budget" if cpi_globale < 1 else "Under-budget", delta_color="inverse")
        
        # ========================================================
        # CRUSCOTTO CASH FLOW NETTO
        # ========================================================
        st.divider()
        st.markdown("### 💶 Esposizione Finanziaria - Cash Flow Netto")
        
        # Calcolo sicuro delle uscite totali leggendo direttamente la colonna Actual Cost (AC)
        tot_uscite_ac = df_evm['AC_Costo_Reale'].sum() if not df_evm.empty else 0.0
        
        # Sommiamo le entrate SOLO se i SAL sono stati effettivamente incassati
        df_sal_calcolo = st.session_state.sal_data
        tot_entrate_sal = 0.0
        if not df_sal_calcolo.empty:
            sal_pagati = df_sal_calcolo[df_sal_calcolo['Stato_Pagamento'] == 'Pagato ▾']
            tot_entrate_sal = pd.to_numeric(sal_pagati['Importo_Euro'], errors='coerce').fillna(0).sum()
            
        cash_flow_netto = tot_entrate_sal - tot_uscite_ac
        
        c_cf1, c_cf2, c_cf3 = st.columns(3)
        c_cf1.metric("Totale Uscite Sostenute (AC)", f"€ {tot_uscite_ac:,.0f}")
        c_cf2.metric("Totale Incassi (SAL Pagati)", f"€ {tot_entrate_sal:,.0f}")
        c_cf3.metric("Cash Flow Netto (Liquidità)", f"€ {cash_flow_netto:,.0f}", 
                     delta="Esposizione Negativa!" if cash_flow_netto < 0 else "Liquidità Positiva", 
                     delta_color="normal" if cash_flow_netto >= 0 else "inverse")
                     
        if cash_flow_netto < 0:
            st.error(f"⚠️ **ALLERTA LIQUIDITÀ:** L'esposizione finanziaria del cantiere è negativa per **€ {abs(cash_flow_netto):,.0f}**. Stai anticipando capitale rispetto a quanto incassato. Valuta l'emissione immediata di un nuovo SAL per ripristinare il flusso di cassa.")
        elif cash_flow_netto > 0:
            st.success(f"✅ **LIQUIDITÀ ATTIVA:** Il cantiere si sta autofinanziando correttamente. Hai un margine di cassa positivo di **€ {cash_flow_netto:,.0f}**.")
        
        # --- SCUDO FINANZIARIO  ---

        st.divider()
        # 1. Motore di calcolo EMV (Expected Monetary Value)
        df_rischi = st.session_state.rischi_data.copy()
        fondo_imprevisti = 0.0
        conteggio_attivi = 0
        
        if not df_rischi.empty:
            # Peschiamo SOLO i rischi non ancora risolti
            rischi_attivi = df_rischi[df_rischi['Stato'].isin(['Attivo ▾', 'Monitorato ▾'])]
            conteggio_attivi = len(rischi_attivi)
            
            if conteggio_attivi > 0:
                # Mappatura standard Punteggio -> Percentuale
                prob_map = {1: 0.10, 2: 0.30, 3: 0.50, 4: 0.70, 5: 0.90} # 10% - 90%
                imp_map = {1: 0.02, 2: 0.05, 3: 0.10, 4: 0.15, 5: 0.20} # 2% - 20% del Budget
                
                for _, r in rischi_attivi.iterrows():
                    id_wbs = str(r['ID_WBS_Rif']).split(' - ')[0].strip()
                    try:
                        prob = int(r['Probabilità (1-5)'])
                        imp = int(r['Impatto (1-5)'])
                        
                        if prob in prob_map and imp in imp_map:
                            # Troviamo il Budget originale della singola WBS interessata
                            wbs_row = df_evm[df_evm['ID_WBS'].astype(str) == id_wbs]
                            bac_wbs = wbs_row['BAC_Budget'].values[0] if not wbs_row.empty else 0.0
                            
                            # Calcolo EMV per questo specifico rischio
                            emv = bac_wbs * prob_map[prob] * imp_map[imp]
                            fondo_imprevisti += emv
                    except:
                        pass
        
        eac_risk_adjusted = tot_eac + fondo_imprevisti

        # 2. Interfaccia Visiva dello Scudo
        st.markdown("### 🛡️ Scudo Finanziario (Risk-Adjusted EVM)")
        with st.expander("Metodologia Contingency Reserve (Fondo Imprevisti)", expanded=True):
            st.markdown(f'''
            **Integrazione EMV (Expected Monetary Value):**
            Il sistema non si limita a calcolare la proiezione dei costi attuale (EAC Tradizionale), ma "congela" dinamicamente una quota di capitale in base ai pericoli registrati nel **Risk Register**. 
            Per ogni rischio ancora *Attivo* o *Monitorato*, il motore trasforma Probabilità e Impatto in percentuali, calcolando la penale attesa sul Budget (BAC) della specifica lavorazione.
            
            * **Formula EMV Singolo:** `BAC Lavorazione` × `Probabilità (%)` × `Impatto (%)`
            * **Formula EAC Risk-Adjusted:** `EAC Tradizionale` + `Σ (Somma EMV Attivi)`
            
            *(N.B. Appena modifichi lo stato di un rischio in "Mitigato" o "Chiuso" nel Tab 8, il Fondo si svuoterà automaticamente, liberando quel capitale per il committente).*
            ''')
            
            c_scudo1, c_scudo2, c_scudo3 = st.columns(3)
            c_scudo1.metric("EAC Tradizionale (Puro)", f"€ {tot_eac:,.0f}", help="Costo stimato a fine progetto senza considerare i rischi futuri.")
            c_scudo2.metric("Fondo Imprevisti (EMV Totale)", f"€ {fondo_imprevisti:,.0f}", delta=f"{conteggio_attivi} Rischi aperti", delta_color="inverse")
            c_scudo3.metric("EAC Risk-Adjusted", f"€ {eac_risk_adjusted:,.0f}", delta="Peggior Scenario Probabile", delta_color="off")
        
        st.divider()

        st.subheader("📈 Andamento di Progetto & Proiezioni")
        
        df_scurve = genera_dati_scurve(df_evm, st.session_state.registro_data, data_status_evm)
        if df_scurve is not None and not df_scurve.empty:
            fig_scurve = px.line(df_scurve, x='Data', y=['PV (Valore Pianificato)', 'EV (Valore Guadagnato)', 'AC (Costo Reale)'], color_discrete_map={'PV (Valore Pianificato)': 'blue', 'EV (Valore Guadagnato)': 'green', 'AC (Costo Reale)': 'red'}, labels={'value': 'Importo (€)', 'variable': 'Metrica EVM'})
            df_past = df_scurve[df_scurve['Data'] <= data_status_evm]
            if not df_past.empty:
                min_date, max_date = df_scurve['Data'].min(), df_scurve['Data'].max()
                spi_effettivo = df_past.iloc[-1]['EV (Valore Guadagnato)'] / df_past.iloc[-1]['PV (Valore Pianificato)'] if df_past.iloc[-1]['PV (Valore Pianificato)'] > 0 else 1.0
                giorni_stimati = min(int((max_date - min_date).days / spi_effettivo) if spi_effettivo > 0 else (max_date - min_date).days, (max_date - min_date).days * 3) 
                data_fine_stimata = min_date + pd.Timedelta(days=giorni_stimati)
                
                # Proiezione Base
                fig_scurve.add_trace(go.Scatter(x=[data_status_evm, data_fine_stimata], y=[df_past.iloc[-1]['AC (Costo Reale)'], tot_eac], mode='lines', line=dict(color='red', dash='dot', width=2), name='Proiezione Costi (EVM Puro)'))
                # Proiezione Risk-Adjusted (se ci sono rischi)
                if contingency_reserve > 0:
                    fig_scurve.add_trace(go.Scatter(x=[data_status_evm, data_fine_stimata], y=[df_past.iloc[-1]['AC (Costo Reale)'], eac_risk_adjusted], mode='lines', line=dict(color='darkred', dash='solid', width=3), name='Proiezione Risk-Adjusted'))
                
                fig_scurve.add_trace(go.Scatter(x=[data_status_evm, data_fine_stimata], y=[df_past.iloc[-1]['EV (Valore Guadagnato)'], tot_bac], mode='lines', line=dict(color='green', dash='dot', width=2), name='Proiezione Lavoro (Tempi)'))
                fig_scurve.update_xaxes(range=[min_date, max(max_date, data_fine_stimata) + pd.Timedelta(days=5)])
            fig_scurve.update_layout(hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), margin=dict(t=50, b=20))
            st.plotly_chart(fig_scurve, use_container_width=True)
        else:
            st.info("ℹ️ Date insufficienti per la Curva ad S.")
        
        st.divider()
        st.subheader("Raffronto Costi per Attività")
        if not df_evm.empty:
            fig_evm = go.Figure(data=[
                go.Bar(name='BAC', x=df_evm['Attività'], y=df_evm['BAC_Budget'], marker_color='lightgray', text=df_evm['BAC_Budget'], texttemplate='€ %{text:,.0f}', textposition='outside', textangle=-90),
                go.Bar(name='EV', x=df_evm['Attività'], y=df_evm['EV'], marker_color='green', text=df_evm['EV'], texttemplate='€ %{text:,.0f}', textposition='outside', textangle=-90),
                go.Bar(name='AC', x=df_evm['Attività'], y=df_evm['AC_Costo_Reale'], marker_color='red', text=df_evm['AC_Costo_Reale'], texttemplate='€ %{text:,.0f}', textposition='outside', textangle=-90)
            ])
            fig_evm.update_layout(barmode='group', margin=dict(t=80), uniformtext_minsize=9, uniformtext_mode='hide')
            st.plotly_chart(fig_evm, use_container_width=True)
            
        col_KPI, col_LEGENDA = st.columns([7, 3]) 
        with col_KPI:
            st.subheader("Indicatori di Performance (KPI)")
            df_kpi = df_evm[['Attività', '%_Completamento', 'CPI', 'SPI', 'CV']].copy()
            def color_kpi(val):
                if isinstance(val, (int, float)):
                    if val < 0.95: return 'color: red; font-weight: bold;'
                    elif val >= 1.0: return 'color: green'
                return ''
            if not df_kpi.empty:
                st.dataframe(df_kpi.style.map(color_kpi, subset=['CPI', 'SPI']).format({'CPI': "{:.2f}", 'SPI': "{:.2f}", 'CV': "€ {:.2f}"}), use_container_width=True)

        with col_LEGENDA:
            st.subheader("Legenda EVM")
            st.markdown("* **CPI:** Efficienza costi (<1 sforamento budget)\n* **SPI:** Efficienza tempi (<1 in ritardo)\n* **CV:** Varianza Costi Assoluta")
            
    # --- TAB 6: GESTIONE FINANZIARIA E CASH FLOW ---
    with tab6:
        st.header("Gestione Finanziaria e Liquidità")
        
        tab_uscite, tab_entrate = st.tabs(["💸 Uscite (Costi Reali - AC)", "💰 Entrate (SAL Certificati)"])
        
        with tab_uscite:
            st.subheader("Brogliaccio Spese e Fatture Fornitori")
            st.markdown("Le spese registrate qui alimentano il Costo Reale (AC) delle singole WBS.")
            
            df_reg = st.session_state.registro_data.copy()
            if not df_reg.empty:
                df_reg['Data'] = pd.to_datetime(df_reg['Data'], errors='coerce').dt.date
            else:
                df_reg['Data'] = pd.Series(dtype='object')
            st.session_state.registro_data = df_reg
            
            leaf_wbs_reg = get_foglie(st.session_state.wbs_data)
            wbs_options = [f"{row['ID_WBS']} - {row['Attività']}" for _, row in leaf_wbs_reg.iterrows()]
            
            edited_registro = st.data_editor(
                st.session_state.registro_data, num_rows="dynamic", use_container_width=True, hide_index=True,
                column_config={
                    "Data": st.column_config.DateColumn("Data Fattura / Spesa"),
                    "ID_WBS_Rif": st.column_config.SelectboxColumn("Attività WBS (Rif.)", options=wbs_options),
                    "Descrizione_Costo": st.column_config.TextColumn("Descrizione", width="medium"),
                    "Importo_Euro": st.column_config.NumberColumn("Importo (€)", format="€ %.2f")
                }
            )
            if st.button("💾 SALVA REGISTRO USCITE", type="primary", use_container_width=True):
                st.session_state.registro_data = edited_registro
                st.success("✅ Registro Spese salvato e costi riallineati!")
                st.rerun()

        with tab_entrate:
            st.subheader("Emissione SAL e Incassi Committenza")
            st.markdown("Registra gli Stati Avanzamento Lavori emessi e pagati. Solo i SAL in stato **'Pagato'** concorreranno al calcolo della liquidità attiva (Cash Flow).")
            
            df_sal = st.session_state.sal_data.copy()
            if not df_sal.empty:
                df_sal['Data_Emissione'] = pd.to_datetime(df_sal['Data_Emissione'], errors='coerce').dt.date
            else:
                df_sal['Data_Emissione'] = pd.Series(dtype='object')
            
            edited_sal = st.data_editor(
                df_sal, num_rows="dynamic", use_container_width=True, hide_index=True,
                column_config={
                    "Data_Emissione": st.column_config.DateColumn("Data SAL"),
                    "Descrizione_SAL": st.column_config.TextColumn("Riferimento / Note", width="medium"),
                    "Importo_Euro": st.column_config.NumberColumn("Importo (€)", format="€ %.2f", min_value=0.0),
                    "Stato_Pagamento": st.column_config.SelectboxColumn("Stato", options=["Emesso ▾", "Pagato ▾"])
                }
            )
            if st.button("💾 SALVA REGISTRO ENTRATE (SAL)", type="primary", use_container_width=True):
                st.session_state.sal_data = edited_sal
                st.success("✅ Registro Incassi aggiornato con successo!")
                st.rerun()

        # ========================================================
        # GRAFICO: ANDAMENTO FLUSSI DI CASSA (CUMULATIVO)
        # ========================================================
        st.divider()
        st.subheader("📊 Andamento Flussi di Cassa (Cash Flow Storico)")
        
        df_grafico_uscite = st.session_state.registro_data.copy()
        df_grafico_entrate = st.session_state.sal_data.copy()
        
        oggi = pd.Timestamp.today().normalize()
        
        # 1. Preparazione Dati Uscite (Spese Cumulative) - ORA RICONOSCE "Importo_Netto"
        if not df_grafico_uscite.empty:
            # Creiamo una lista di tutti i possibili nomi che hai usato per i costi
            possibili_nomi_costo = ['Importo_Netto', 'Importo_Euro', 'Costo', 'Importo']
            col_imp_uscite = next((col for col in possibili_nomi_costo if col in df_grafico_uscite.columns), None)
            
            # Stessa cosa per la data
            possibili_nomi_data = ['Data', 'Data Fattura / Spesa', 'Data_Fattura']
            col_data_uscite = next((col for col in possibili_nomi_data if col in df_grafico_uscite.columns), None)
            
            if col_data_uscite and col_imp_uscite:
                df_grafico_uscite['Data_Grafico'] = pd.to_datetime(df_grafico_uscite[col_data_uscite], errors='coerce')
                df_grafico_uscite[col_imp_uscite] = pd.to_numeric(df_grafico_uscite[col_imp_uscite], errors='coerce').fillna(0)
                df_grafico_uscite = df_grafico_uscite.dropna(subset=['Data_Grafico'])
                df_grafico_uscite = df_grafico_uscite.sort_values('Data_Grafico')
                df_grafico_uscite['Cumulato'] = df_grafico_uscite[col_imp_uscite].cumsum()
            else:
                df_grafico_uscite = pd.DataFrame()
            
        # 2. Preparazione Dati Entrate (Solo SAL Pagati, Cumulativi)
        if not df_grafico_entrate.empty and 'Stato_Pagamento' in df_grafico_entrate.columns:
            df_grafico_entrate = df_grafico_entrate[df_grafico_entrate['Stato_Pagamento'] == 'Pagato ▾'].copy()
            col_imp_entrate = 'Importo_Euro' if 'Importo_Euro' in df_grafico_entrate.columns else 'Importo'
            
            if 'Data_Emissione' in df_grafico_entrate.columns and col_imp_entrate in df_grafico_entrate.columns:
                df_grafico_entrate['Data_Grafico'] = pd.to_datetime(df_grafico_entrate['Data_Emissione'], errors='coerce')
                df_grafico_entrate[col_imp_entrate] = pd.to_numeric(df_grafico_entrate[col_imp_entrate], errors='coerce').fillna(0)
                df_grafico_entrate = df_grafico_entrate.dropna(subset=['Data_Grafico'])
                df_grafico_entrate = df_grafico_entrate.sort_values('Data_Grafico')
                df_grafico_entrate['Cumulato'] = df_grafico_entrate[col_imp_entrate].cumsum()
            else:
                df_grafico_entrate = pd.DataFrame()

        # 3. Disegno del Grafico
        if df_grafico_uscite.empty and df_grafico_entrate.empty:
            st.info("Aggiungi spese o incassi (e ricordati di cliccare SALVA) per visualizzare il grafico.")
        else:
            fig_cf = go.Figure()
            
            # Traccia Costi (Rossa)
            if not df_grafico_uscite.empty:
                # Estensione linea fino a oggi per continuità visiva
                ultima_data_u = df_grafico_uscite['Data_Grafico'].max()
                if ultima_data_u < oggi:
                    nuova_riga = pd.DataFrame({'Data_Grafico': [oggi], 'Cumulato': [df_grafico_uscite['Cumulato'].iloc[-1]]})
                    df_grafico_uscite = pd.concat([df_grafico_uscite, nuova_riga], ignore_index=True)
                    
                fig_cf.add_trace(go.Scatter(
                    x=df_grafico_uscite['Data_Grafico'],
                    y=df_grafico_uscite['Cumulato'],
                    mode='lines+markers',
                    name='Uscite Cumulate (AC)',
                    line=dict(color='#D32F2F', width=3, shape='hv'), # 'hv' = a gradini
                    fill='tozeroy',
                    fillcolor='rgba(211, 47, 47, 0.1)'
                ))
                
            # Traccia Incassi (Verde)
            if not df_grafico_entrate.empty:
                # Estensione linea fino a oggi per continuità visiva
                ultima_data_e = df_grafico_entrate['Data_Grafico'].max()
                if ultima_data_e < oggi:
                    nuova_riga = pd.DataFrame({'Data_Grafico': [oggi], 'Cumulato': [df_grafico_entrate['Cumulato'].iloc[-1]]})
                    df_grafico_entrate = pd.concat([df_grafico_entrate, nuova_riga], ignore_index=True)

                fig_cf.add_trace(go.Scatter(
                    x=df_grafico_entrate['Data_Grafico'],
                    y=df_grafico_entrate['Cumulato'],
                    mode='lines+markers',
                    name='Incassi Cumulati (SAL)',
                    line=dict(color='#4CAF50', width=3, shape='hv'), # 'hv' = a gradini
                    fill='tozeroy',
                    fillcolor='rgba(76, 175, 80, 0.1)'
                ))
                
            fig_cf.update_layout(
                xaxis_title="Linea Temporale",
                yaxis_title="Importo Cumulato (€)",
                hovermode="x unified",
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, bgcolor="rgba(255,255,255,0.8)")
            )
            
            st.plotly_chart(fig_cf, use_container_width=True)
            
    # --- TAB 7: DIREZIONE LAVORI, CAPA & REPORTISTICA ---
    with tab7:
        st.header("Direzione Lavori: Interventi (CAPA) e Simulazioni")
        
        leaf_wbs_capa = get_foglie(st.session_state.wbs_data)
        wbs_options_capa = [f"{row['ID_WBS']} - {row['Attività']}" for _, row in leaf_wbs_capa.iterrows()]
        obs_options_capa = [f"{row['ID_OBS']} - {row['Risorsa']}" for _, row in st.session_state.obs_data.iterrows()]
        
        # NUOVO: Recuperiamo i rischi registrati dal Tab 8 per popolare la tendina
        lista_rischi = [""] + [str(r) for r in st.session_state.rischi_data['Descrizione_Rischio'].dropna().unique() if str(r).strip() != ""]
        
        st.subheader("1. Registro Azioni Correttive e Preventive (CAPA)")
        
        df_capa = st.session_state.capa_data.copy()
        
        # Retro-compatibilità: Se carichi un vecchio progetto senza questa colonna, la crea al volo
        if 'Rischio_Associato' not in df_capa.columns:
            df_capa['Rischio_Associato'] = ""
            
        if not df_capa.empty:
            df_capa['Data_Apertura'] = pd.to_datetime(df_capa['Data_Apertura'], errors='coerce').dt.date
        else:
            df_capa['Data_Apertura'] = pd.Series(dtype='object')
        st.session_state.capa_data = df_capa
        
        edited_capa = st.data_editor(
            st.session_state.capa_data, num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={
                "Data_Apertura": st.column_config.DateColumn("Data Segnalazione"),
                "ID_WBS_Rif": st.column_config.SelectboxColumn("Attività WBS (Rif.)", options=wbs_options_capa),
                "Tipo_Azione": st.column_config.SelectboxColumn("Tipo", options=["Correttiva ▾", "Preventiva ▾"]),
                "Descrizione": st.column_config.TextColumn("Descrizione Intervento / Ordine", width="medium"),
                "Responsabile_OBS": st.column_config.SelectboxColumn("Risorsa (OBS)", options=obs_options_capa),
                "Rischio_Associato": st.column_config.SelectboxColumn("Rischio da mitigare ▾", options=lista_rischi), # LA NUOVA TENDINA
                "Stato": st.column_config.SelectboxColumn("Stato", options=["Aperto ▾", "In Lavorazione ▾", "Chiuso ▾"])
            }
        )
        
        st.divider()
        st.warning("⚠️ **Ricordati di cliccare il tasto rosso qui sotto dopo aver inserito i dati!**")
        if st.button("💾 SALVA REGISTRO CAPA", type="primary", use_container_width=True):
            st.session_state.capa_data = edited_capa
            
            # --- AUTO-MITIGAZIONE RISCHI (IL PILOTA AUTOMATICO) ---
            rischi_aggiornati = False
            df_rischi = st.session_state.rischi_data.copy()
            
            for _, capa_row in edited_capa.iterrows():
                # Se l'azione di cantiere è stata CHIUSA e c'era un rischio collegato
                if capa_row['Stato'] == 'Chiuso ▾' and pd.notna(capa_row.get('Rischio_Associato')) and str(capa_row.get('Rischio_Associato')).strip() != "":
                    rischio_target = str(capa_row['Rischio_Associato']).strip()
                    
                    # Cerca questo rischio nel Tab 8. Se è ancora Attivo o Monitorato...
                    mask = (df_rischi['Descrizione_Rischio'] == rischio_target) & (~df_rischi['Stato'].isin(['Mitigato ▾', 'Chiuso ▾']))
                    if mask.any():
                        # ...abbatti la sua Probabilità a 1 e settalo come Mitigato!
                        df_rischi.loc[mask, 'Probabilità (1-5)'] = 1
                        df_rischi.loc[mask, 'Stato'] = 'Mitigato ▾'
                        rischi_aggiornati = True
            
            if rischi_aggiornati:
                st.session_state.rischi_data = df_rischi
                st.success("✅ Interventi salvati! Il pilota automatico ha **Mitigato** i rischi associati alle azioni chiuse nel Tab 8.")
            else:
                st.success("✅ Interventi salvati con successo nel database!")
                
            st.rerun()

        # ------------------------------
        # SIMULAZIONE
        # ------------------------------
        
        with st.expander("🔬 2. Ambiente di Simulazione (Compromesso Costi / Tempi)"):
            st.markdown("Simula l'impatto di un'azione correttiva senza alterare i principi base dell'EVM. Valuta costi di *Crashing* e l'incidenza sui Costi Indiretti di cantiere.")
            
            c_sim1, c_sim2, c_sim3, c_sim4 = st.columns([2, 1.5, 1.5, 1.5])
            wp_scelto = c_sim1.selectbox("Seleziona Work Package da simulare", options=wbs_options_capa)
            extra_costo = c_sim2.number_input("Costo Extra Diretto (€)", value=0.0, step=500.0, help="Costo vivo aggiuntivo (es. premi, straordinari). Intacca il CPI e fa salire l'EAC.")
            var_giorni = c_sim3.number_input("Variazione Tempi (Giorni)", value=0, step=1, help="Usa numeri NEGATIVI per anticipare (es. -5), POSITIVI per ritardare (es. 5).")
            costo_indiretto_gg = c_sim4.number_input("Costi Indiretti (€/gg)", value=0.0, step=50.0, help="Costi fissi per ogni giorno di ritardo (es. gru, baraccamenti, penali)")
            
            if wp_scelto:
                wp_id = wp_scelto.split(' - ')[0]
                
                # --- CALCOLO REALE (BASE) ---
                oggi = pd.Timestamp.today().date()
                df_reale_calc = calcola_evm(get_foglie(st.session_state.wbs_data), oggi)
                
                eac_attuale = df_reale_calc['EAC'].sum()
                ac_attuale_tot = df_reale_calc['AC_Costo_Reale'].sum()
                
                # --- SIMULAZIONE EVM PURA ---
                df_simulazione = st.session_state.wbs_data.copy()
                indice_riga = df_simulazione.index[df_simulazione['ID_WBS'] == wp_id].tolist()
                if indice_riga:
                    idx = indice_riga[0]
                    # Iniettare costo diretto abbassa il CPI e alza l'EAC secondo le regole pure dell'EVM
                    df_simulazione.at[idx, 'AC_Costo_Reale'] = pd.to_numeric(df_simulazione.at[idx, 'AC_Costo_Reale'], errors='coerce') + extra_costo
                    
                df_sim_calc = calcola_evm(get_foglie(df_simulazione), oggi)
                eac_simulato = df_sim_calc['EAC'].sum()
                ac_simulato_tot = df_sim_calc['AC_Costo_Reale'].sum()
                
                # --- SIMULAZIONE TEMPI E COSTI INDIRETTI ---
                cpm_nodes = calcola_cpm(st.session_state.wbs_data)
                is_wp_critical = cpm_nodes.get(wp_id, {}).get('is_critical', False)
                
                min_date = pd.to_datetime(df_reale_calc['Inizio_Previsto']).min().date()
                max_date = pd.to_datetime(df_reale_calc['Fine_Prevista']).max().date()
                
                if pd.notna(min_date) and pd.notna(max_date):
                    giorni_pianificati = (max_date - min_date).days
                    tot_pv = df_reale_calc['PV'].sum()
                    tot_ev = df_reale_calc['EV'].sum()
                    spi_attuale = tot_ev / tot_pv if tot_pv > 0 else 1.0
                    
                    giorni_stimati_attuali = int(giorni_pianificati / spi_attuale) if spi_attuale > 0 else giorni_pianificati
                    giorni_stimati_attuali = min(giorni_stimati_attuali, giorni_pianificati * 3)
                    
                    data_fine_attuale = min_date + pd.Timedelta(days=giorni_stimati_attuali)
                    
                    # Applica i giorni solo se la voce è critica
                    if is_wp_critical:
                        data_fine_simulata = data_fine_attuale + pd.Timedelta(days=var_giorni)
                        giorni_delta = var_giorni
                    else:
                        data_fine_simulata = data_fine_attuale
                        giorni_delta = 0
                else:
                    data_fine_attuale = data_fine_simulata = oggi
                    giorni_delta = 0
                    
                # Calcolo separato per gli Indiretti
                impatto_indiretti = giorni_delta * costo_indiretto_gg
                costo_totale_effettivo = eac_simulato + impatto_indiretti
                
                # --- MESSAGGI DI ALLERTA INTELLIGENTI ---
                if var_giorni != 0 and not is_wp_critical:
                    st.warning(f"⚠️ **Attenzione:** Stai modificando un'attività NON CRITICA. Cambiare i tempi di '{wp_scelto.split(' - ')[1]}' non sposterà la data di fine cantiere, quindi non inciderà sui costi indiretti.")
                elif var_giorni < 0 and is_wp_critical:
                    st.success(f"✅ **Crashing Efficace:** L'attività è critica. Anticipando di {abs(var_giorni)} giorni la consegna, generi un risparmio extra (indiretto) di € {abs(impatto_indiretti):,.0f}.")
                elif var_giorni > 0 and is_wp_critical:
                    st.error(f"🚨 **Allerta Ritardo:** Un ritardo critico di {var_giorni} giorni farà slittare il cantiere, aggiungendo € {impatto_indiretti:,.0f} ai tuoi costi fissi.")
                
                c_res1, c_res2, c_res3 = st.columns([1, 1, 1.5])
                
                with c_res1:
                    st.metric("EAC Simulato (Puro EVM)", f"€ {eac_simulato:,.0f}", delta=f"Deriva CPI: € {eac_simulato - eac_attuale:,.0f}", delta_color="inverse")
                    
                    if giorni_delta < 0:
                        label_delta_giorni = f"Anticipo: {abs(giorni_delta)} gg"
                        colore_delta = "normal"
                    elif giorni_delta > 0:
                        label_delta_giorni = f"Ritardo: {giorni_delta} gg"
                        colore_delta = "inverse"
                    else:
                        label_delta_giorni = "Nessuna variazione globale"
                        colore_delta = "off"
                        
                    st.metric("Nuova Data di Consegna", data_fine_simulata.strftime('%d/%m/%Y') if pd.notna(data_fine_simulata) else "N/D", delta=label_delta_giorni, delta_color=colore_delta)
                
                with c_res2:
                    st.metric("Impatto Costi Indiretti", f"€ {impatto_indiretti:,.0f}", delta="Risparmio da anticipo" if impatto_indiretti < 0 else ("Penale / Spesa extra" if impatto_indiretti > 0 else "Nessun impatto"), delta_color="inverse")
                    st.metric("Costo Finale Effettivo", f"€ {costo_totale_effettivo:,.0f}", delta=f"Delta vs Attuale: € {costo_totale_effettivo - eac_attuale:,.0f}", delta_color="inverse")

                with c_res3:
                    fig_sim = go.Figure()
                    
                    # Traiettoria Attuale (Rossa)
                    fig_sim.add_trace(go.Scatter(
                        x=[oggi, data_fine_attuale], 
                        y=[ac_attuale_tot, eac_attuale], 
                        mode='lines+markers+text', 
                        name='EAC Attuale', 
                        line=dict(color='red', dash='dash', width=2), 
                        text=["", f"€ {eac_attuale:,.0f}"], 
                        textposition="bottom right"
                    ))
                    
                    # Traiettoria Simulata (Blu) - FIX: Ora parte dallo stesso punto di oggi (AC Attuale) e punta dritta al nuovo traguardo!
                    fig_sim.add_trace(go.Scatter(
                        x=[oggi, data_fine_simulata], 
                        y=[ac_attuale_tot, costo_totale_effettivo], 
                        mode='lines+markers+text', 
                        name='EAC + Indiretti', 
                        line=dict(color='blue', dash='solid', width=3), 
                        text=["", f"€ {costo_totale_effettivo:,.0f}"], 
                        textposition="top left"
                    ))
                    
                    fig_sim.update_layout(
                        title="Impatto Strategico Globale", 
                        height=280, 
                        margin=dict(l=10, r=10, t=35, b=10), 
                        yaxis_title="Costo (€)", 
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    st.plotly_chart(fig_sim, use_container_width=True)
                    
    # ---------------------------------------------------------
        
        st.subheader("3. Stampa Verbale di Direzione Lavori")
        
        # 1. Filtri per scegliere cosa stampare
        col_f1, col_f2 = st.columns([1, 2])
        filtro_stampa = col_f1.radio("Quali interventi includere nel verbale?", ["Tutti i registrati", "Solo l'ultimo inserito", "Intervallo di date"])
        
        df_stampa = st.session_state.capa_data.copy()
        if not df_stampa.empty:
            df_stampa['Data_Apertura'] = pd.to_datetime(df_stampa['Data_Apertura']).dt.date
            
            if filtro_stampa == "Solo l'ultimo inserito":
                df_stampa = df_stampa.tail(1)
            elif filtro_stampa == "Intervallo di date":
                d_start = col_f2.date_input("Da data:", value=pd.Timestamp.today().date())
                d_end = col_f2.date_input("A data:", value=pd.Timestamp.today().date())
                df_stampa = df_stampa[(df_stampa['Data_Apertura'] >= d_start) & (df_stampa['Data_Apertura'] <= d_end)]

        # 2. Generazione automatica in background del documento WORD
        df_evm_rep = calcola_evm(get_foglie(st.session_state.wbs_data), pd.Timestamp.today().date())
        tot_ev_rep = df_evm_rep['EV'].sum()
        tot_ac_rep = df_evm_rep['AC_Costo_Reale'].sum()
        tot_pv_rep = df_evm_rep['PV'].sum()
        cpi_rep = tot_ev_rep / tot_ac_rep if tot_ac_rep > 0 else 1.0
        spi_rep = tot_ev_rep / tot_pv_rep if tot_pv_rep > 0 else 1.0
        eac_rep = df_evm_rep['EAC'].sum()
        
        doc = Document()
        doc.add_heading('VERBALE DI DIREZIONE LAVORI', 0)
        doc.add_paragraph(f"Progetto: {st.session_state.nome_progetto_attivo}")
        doc.add_paragraph(f"Data emissione verbale: {pd.Timestamp.today().strftime('%d/%m/%Y')}")
        
        doc.add_heading('1. Stato Avanzamento Lavori (EVM)', level=1)
        p = doc.add_paragraph()
        p.add_run(f"CPI (Efficienza Costi): {cpi_rep:.2f}\n").bold = True
        p.add_run(f"SPI (Efficienza Tempi): {spi_rep:.2f}\n").bold = True
        p.add_run(f"Costo Finale Stimato (EAC): {eac_rep:,.2f} Euro").bold = True
        doc.add_paragraph("Nota: Un indicatore inferiore a 1.00 indica un superamento del budget o un ritardo sui tempi.")
        
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
        
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # 3. Tasto nativo di download (fuori da altri bottoni!)
        st.download_button(
            label="⬇️ Scarica il file Word pronto per la firma",
            data=buffer,
            file_name=f"Verbale_{st.session_state.nome_progetto_attivo}_{pd.Timestamp.today().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
    # ---------------------------------------------------
    # --- TAB 8: MATRICE DEI RISCHI (RISK MANAGEMENT) ---
    # ---------------------------------------------------
    
    with tab8:
        st.header("Matrice di Rischio Strategico")
        st.markdown("Valuta e mappa i rischi associati alle singole lavorazioni (WBS) assegnando un punteggio da **1 (Minimo)** a **5 (Massimo)**.")
        
        leaf_wbs_rischi = get_foglie(st.session_state.wbs_data)
        wbs_options_rischi = [f"{row['ID_WBS']} - {row['Attività']}" for _, row in leaf_wbs_rischi.iterrows()]
        
        st.subheader("1. Registro dei Rischi (Risk Register)")
        
        edited_rischi = st.data_editor(
            st.session_state.rischi_data, num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={
                "ID_WBS_Rif": st.column_config.SelectboxColumn("Attività WBS (Rif.)", options=wbs_options_rischi),
                "Descrizione_Rischio": st.column_config.TextColumn("Descrizione / Nome del Rischio", width="large"),
                "Probabilità (1-5)": st.column_config.NumberColumn("Probabilità (1=Rara, 5=Certa)", min_value=1, max_value=5, step=1),
                "Impatto (1-5)": st.column_config.NumberColumn("Impatto (1=Lieve, 5=Critico)", min_value=1, max_value=5, step=1),
                "Stato": st.column_config.SelectboxColumn("Stato", options=["Attivo ▾", "Monitorato ▾", "Mitigato ▾", "Chiuso ▾"])
            }
        )
        
        st.divider()
        st.warning("⚠️ **Ricordati di cliccare il tasto rosso qui sotto dopo aver modificato o aggiunto i rischi!**")
        if st.button("💾 SALVA REGISTRO RISCHI E AGGIORNA GRAFICO", type="primary", use_container_width=True):
            st.session_state.rischi_data = edited_rischi
            st.success("✅ Rischi mappati con successo!")
            st.rerun()
            
        st.divider()
        st.subheader("2. Mappa di Calore Operativa (Heatmap)")
        
        # IL TUO NUOVO INTERRUTTORE
        mostra_nomi = st.toggle("👁️ Mostra nomi completi delle attività sul grafico", value=False)
        
        df_plot = st.session_state.rischi_data.copy()
        df_plot = df_plot.dropna(subset=['Probabilità (1-5)', 'Impatto (1-5)'])
        
        if not df_plot.empty:
            df_plot['Probabilità (1-5)'] = pd.to_numeric(df_plot['Probabilità (1-5)'])
            df_plot['Impatto (1-5)'] = pd.to_numeric(df_plot['Impatto (1-5)'])
            df_plot['Punteggio'] = df_plot['Probabilità (1-5)'] * df_plot['Impatto (1-5)']
            
            fig_risk = go.Figure()
            
            # --- 4 QUADRANTI COLORATI ---
            fig_risk.add_shape(type="rect", x0=0.5, y0=0.5, x1=3, y1=3, fillcolor="#E8F5E9", line_width=0, layer="below")
            fig_risk.add_shape(type="rect", x0=0.5, y0=3, x1=3, y1=5.5, fillcolor="#FFFDE7", line_width=0, layer="below")
            fig_risk.add_shape(type="rect", x0=3, y0=0.5, x1=5.5, y1=3, fillcolor="#FFFDE7", line_width=0, layer="below")
            fig_risk.add_shape(type="rect", x0=3, y0=3, x1=5.5, y1=5.5, fillcolor="#FFEBEE", line_width=0, layer="below")
            
            # Linee tratteggiate a Croce
            fig_risk.add_hline(y=3, line_dash="dash", line_color="gray", line_width=2)
            fig_risk.add_vline(x=3, line_dash="dash", line_color="gray", line_width=2)
            
            # --- TRACCE FANTASMA PER LA LEGENDA GRAFICA ---
            fig_risk.add_trace(go.Scatter(x=[None], y=[None], mode='markers', marker=dict(size=18, color='#E8F5E9', symbol='square', line=dict(color='gray', width=1)), name='Safe-Zone (Rischio Controllato)'))
            fig_risk.add_trace(go.Scatter(x=[None], y=[None], mode='markers', marker=dict(size=18, color='#FFFDE7', symbol='square', line=dict(color='gray', width=1)), name='Area di Attenzione'))
            fig_risk.add_trace(go.Scatter(x=[None], y=[None], mode='markers', marker=dict(size=18, color='#FFEBEE', symbol='square', line=dict(color='gray', width=1)), name='Area Critica'))
            
            # Logica dell'interruttore: Testo Lungo vs ID Corto
            if mostra_nomi:
                etichette_punti = df_plot['ID_WBS_Rif'].apply(lambda x: str(x) if pd.notna(x) else "")
            else:
                etichette_punti = df_plot['ID_WBS_Rif'].apply(lambda x: str(x).split(' - ')[0] if pd.notna(x) else "")

            # --- I PUNTI DEI RISCHI (COLORATI IN BASE ALLO STATO) ---
            colori_stato = {"Attivo ▾": "#F44336", "Monitorato ▾": "#FF9800", "Mitigato ▾": "#4CAF50", "Chiuso ▾": "#9E9E9E"}
            df_plot['Colore_Punto'] = df_plot['Stato'].map(colori_stato).fillna("#607D8B")
            
            fig_risk.add_trace(go.Scatter(
                x=df_plot['Probabilità (1-5)'], y=df_plot['Impatto (1-5)'],
                mode='markers+text',
                text=etichette_punti,
                textposition="top center",
                textfont=dict(size=12, color='DarkSlateGrey', family="Arial Black"),
                marker=dict(size=18, color=df_plot['Colore_Punto'], line=dict(width=2, color='white')),
                showlegend=False,
                hovertemplate="<b>WBS: %{text}</b><br>Rischio: %{customdata[0]}<br>Probabilità: %{x}<br>Impatto: %{y}<br>Punteggio: %{customdata[2]}<br>Stato: %{customdata[1]}<extra></extra>",
                customdata=df_plot[['Descrizione_Rischio', 'Stato', 'Punteggio']]
            ))
            
            # --- IL PALLINO ROSSO GIGANTE DEL RISCHIO MEDIO ---
            media_prob = df_plot['Probabilità (1-5)'].mean()
            media_imp = df_plot['Impatto (1-5)'].mean()
            
            fig_risk.add_trace(go.Scatter(
                x=[media_prob], 
                y=[media_imp],
                mode='markers+text',
                text=["RISCHIO MEDIO"],
                textposition="bottom center",
                textfont=dict(size=14, color='DarkRed', family="Arial Black"),
                marker=dict(size=35, color='rgba(255, 0, 0, 0.7)', line=dict(width=3, color='DarkRed')),
                showlegend=False,
                hovertemplate="<b>RISCHIO GLOBALE MEDIO</b><br>Probabilità Media: %{x:.2f}<br>Impatto Medio: %{y:.2f}<extra></extra>"
            ))
            
            fig_risk.update_layout(
                xaxis=dict(title="<b>Probabilità di Accadimento (1-5)</b>", range=[0.5, 5.5], dtick=1, gridcolor='white', zeroline=False),
                yaxis=dict(title="<b>Impatto sul Progetto (1-5)</b>", range=[0.5, 5.5], dtick=1, gridcolor='white', zeroline=False),
                height=650, plot_bgcolor='white', margin=dict(l=60, r=40, t=40, b=60),
                legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01, bgcolor="rgba(255, 255, 255, 0.9)", bordercolor="lightgray", borderwidth=1)
            )
            
            st.plotly_chart(fig_risk, use_container_width=True)
            
            st.markdown("**Legenda Stato Rischi:** 🔴 `Attivo` | 🟠 `Monitorato` | 🟢 `Mitigato` | ⚪ `Chiuso`")
        else:
            st.info("ℹ️ Compila i valori numerici (da 1 a 5) nella colonna Probabilità e Impatto della tabella qui sopra per generare la matrice.")

# ========================================================
    # --- TAB 9: MANUALE OPERATIVO, FORMULARIO & FAQ V.2.0 ---
    # ========================================================
    with tab9:
        st.header("📖 Manuale Operativo & Knowledge Base (Versione 2.0)")
        st.markdown("Benvenuto nella centrale di controllo della documentazione di cantiere. Questo manuale guida l'utente attraverso l'architettura dei dati, le formule matematiche e le logiche di automazione integrate nell'applicazione.")

        # Sotto-sezioni del manuale per massima pulizia visiva
        t_sec1, t_sec2, t_sec3, t_sec4, t_sec5 = st.tabs([
            "🧭 Flusso di Lavoro", 
            "📚 Glossario Tecnico",
            "📐 Formulario EVM & Finanza", 
            "⚠️ Casi Studio & Falsi Ingippi", 
            "🛠️ Roadmap Versione 2.0"
        ])

        # --- SEZIONE 1: FLUSSO DI LAVORO ---
        with t_sec1:
            st.subheader("Il Ciclo di Vita del Progetto nell'App")
            st.markdown("Il software è progettato in modo che le informazioni viaggino automaticamente tra le varie sezioni, creando un ciclo continuo di pianificazione, misurazione, allerta e correzione.")
            
            st.markdown("""
            1. **Tab 2 (OBS & Risorse):** Inserisci le imprese, le maestranze e le attrezzature disponibili. Sono i soggetti che animeranno il cantiere.
            2. **Tab 1 (WBS - Lavorazioni):** Struttura l'albero delle attività. Assegna i budget (BAC), i predecessori e collega ciascuna lavorazione alla risorsa responsabile (OBS) e alle date previste.
            3. **Tab 4 (Gantt & Monitoraggio):** Controlla l'allineamento temporale. Le barre si coloreranno automaticamente in base all'efficienza (SPI).
            4. **Tab 6 (Gestione Finanziaria):** 
               * *Uscite:* Registra fatture e costi reali associandoli alle WBS.
               * *Entrate:* Emetti e traccia i SAL certificati e pagati dalla committenza.
            5. **Tab 7 (Rischi & CAPA):** Gestisci il fondo imprevisti e apri azioni correttive (Non-Conformità) qualora qualcosa non rispetti gli standard qualitativi.
            6. **Tab 5 & Radar (GIANFRANCO CONSIGLIA):** Monitora il cruscotto di controllo per verificare l'esposizione di cassa (Cash Flow) e le allerte di sovraccarico risorse.
            """)
            
            import graphviz
            diag_guida = graphviz.Digraph(engine='dot')
            diag_guida.attr(rankdir='LR', splines='ortho', nodesep='0.6', ranksep='1.0')
            diag_guida.attr('node', shape='box', style='rounded,filled', fontname='Helvetica', fontsize='10', margin='0.15')
            
            # Nodi Input Base (Grigi)
            with diag_guida.subgraph(name='cluster_input') as c:
                c.attr(label='FASE 1: INPUT DATI', style='dashed', color='gray')
                c.node('T1', 'TAB 1: WBS (Lavorazioni)\nStruttura, Budget, Date,\n% Completamento & Cancello 99%', fillcolor='#F5F5F5')
                c.node('T2', 'TAB 2: OBS (Risorse)\nAnagrafiche e Assegnazioni', fillcolor='#F5F5F5')
                c.node('T6', 'TAB 6: FINANZA & CASH FLOW\n- Uscite (Fatture e AC)\n- Entrate (SAL e Incassi Pagati)', fillcolor='#F5F5F5')

            # Nodi Motore e Output (Colorati)
            with diag_guida.subgraph(name='cluster_analisi') as c:
                c.attr(label='FASE 2: ANALISI E PREVISIONI', style='dashed', color='gray')
                c.node('T3', 'TAB 3: GRAFO E CPM\nPercorso Critico e Margini', fillcolor='#FFE0B2', color='#FBC02D', penwidth='2')
                c.node('T4', 'TAB 4: GANTT\nBaseline vs Esecutivo\nColorazione SPI in tempo reale', fillcolor='#FFE0B2', color='#FBC02D', penwidth='2')
                c.node('T5', 'TAB 5: EVM & CASH FLOW\n- Indicatori (CPI / SPI)\n- Esposizione di Cassa Netta\n- Grafico Storico a Gradoni', fillcolor='#C8E6C9', color='#43A047', penwidth='2')

            # Nodi Gestione Rischi e Controllo
            with diag_guida.subgraph(name='cluster_controllo') as c:
                c.attr(label='FASE 3: DIREZIONE LAVORI & GIANFRANCO CONSIGLIA', style='dashed', color='gray')
                c.node('RADAR', 'RADAR (GIANFRANCO CONSIGLIA)\n- Rilevamento Sovraccarichi\n- Gestione Deroghe / Ignora', fillcolor='#E1BEE7', color='#8E24AA', penwidth='2')
                c.node('T8', 'TAB 8: MATRICE RISCHI\nHeatmap e Fondo Imprevisti', fillcolor='#FFCDD2', color='#E53935', penwidth='2')
                c.node('T7', 'TAB 7: DIREZIONE & CAPA\n- Registro Non-Conformità\n- Blocco Qualità 99% WBS', fillcolor='#BBDEFB', color='#1E88E5', penwidth='2')
            
            # Archi Relazionali (Frecce)
            diag_guida.edge('T2', 'T1', ' Assegnazione', color='gray')
            diag_guida.edge('T1', 'T3', ' Predecessori', color='gray')
            diag_guida.edge('T1', 'T4', ' Schedulazione', color='gray')
            diag_guida.edge('T1', 'T5', ' Budget (BAC) & Valore (EV)', color='gray')
            diag_guida.edge('T2', 'RADAR', ' Verifica Sovrapposizioni', color='#8E24AA', fontcolor='#8E24AA')
            diag_guida.edge('T1', 'RADAR', ' Incrocio Date / Risorse', color='#8E24AA', fontcolor='#8E24AA')
            diag_guida.edge('T6', 'T5', ' Costo Reale (AC) & SAL', color='gray')
            
            # Archi Rischio, CAPA e Blocchi
            diag_guida.edge('T8', 'T3', ' Allerta Visiva', color='#E53935', fontcolor='#E53935')
            diag_guida.edge('T8', 'T5', ' Fondo Imprevisti', color='#E53935', fontcolor='#E53935')
            diag_guida.edge('T8', 'T7', ' Attiva Mitigazione', color='#E53935', fontcolor='#E53935')
            diag_guida.edge('T7', 'T1', ' 🔒 Blocca WBS a 99%\nse CAPA Aperta', color='#1E88E5', fontcolor='#1E88E5', style='bold')
            
            st.graphviz_chart(diag_guida, use_container_width=True)
            
            st.markdown("""
            ### 📌 Legenda dei Flussi Automatici:
            * **Da Tab 1, 2 a Radar (GIANFRANCO CONSIGLIA):** Il sistema controlla in tempo reale se la stessa risorsa è impegnata su più fronti nello stesso periodo, segnalando l'eventuale sovraccarico (con opzione di deroga).
            * **Da Tab 7 a Tab 1 (Cancello di Qualità):** Se esiste una CAPA attiva su una WBS, il sistema impedisce matematicamente di certificarla al 100%, bloccandola al 99% finché il problema non viene chiuso.
            * **Da Tab 6 a Tab 5 (Cash Flow):** Le uscite (costi reali) e le entrate (SAL pagati) alimentano la curva cumulativa e l'indicatore di esposizione finanziaria netta.
            """)

        # --- 2. GLOSSARIO EVM ---
        with t_sec2:
            st.subheader("2. Glossario EVM (Earned Value Management)")
        
            c_glos1, c_glos2 = st.columns(2)
            with c_glos1:
                with st.expander("Valori Base (I Pilastri)"):
                    st.markdown("""
                    * **BAC (Budget at Completion):** Il Budget totale pianificato per l'intero progetto o lavorazione.
                    * **PV (Planned Value):** Il valore del lavoro che *dovrebbe* essere stato completato fino ad oggi secondo il cronoprogramma. (Si calcola proiettando il BAC nel tempo).
                    * **EV (Earned Value):** Il valore del lavoro *effettivamente* completato fino ad oggi. È la metrica più importante: indica i soldi che il cantiere ha realmente "guadagnato" producendo.
                    * **AC (Actual Cost):** I costi reali effettivamente sostenuti per il lavoro svolto fino ad oggi (fatture, ore manodopera, materiali).
                    """)
            with c_glos2:
                with st.expander("Indicatori di Performance (KPI)"):
                    st.markdown("""
                    * **CV (Cost Variance):** Varianza dei costi. Se è negativa, stai spendendo più del previsto.
                    * **SV (Schedule Variance):** Varianza dei tempi. Se è negativa, sei in ritardo sul cronoprogramma.
                    * **CPI (Cost Performance Index):** Efficienza dei costi. Valore ideale: ≥ 1.0. Se vale 0.8, significa che per ogni Euro speso stai producendo solo 80 centesimi di valore.
                    * **SPI (Schedule Performance Index):** Efficienza dei tempi. Valore ideale: ≥ 1.0. Se vale 0.9, stai viaggiando al 90% della velocità prevista.
                    """)
                
            with st.expander("Previsioni e Rischio (Proiezioni)"):
                st.markdown("""
                * **EAC (Estimate At Completion):** Costo totale stimato a fine progetto, ricalcolato in base all'efficienza attuale (CPI). Ti dice quanto ti costerà davvero il cantiere se continui a lavorare come stai facendo oggi.
                * **ETC (Estimate To Complete):** I fondi residui necessari per finire il lavoro da oggi in poi.
                * **VAC (Variance At Completion):** Scostamento finale previsto (BAC - EAC). Se è negativo, il progetto si chiuderà in perdita rispetto al budget iniziale.
                * **EMV (Expected Monetary Value):** Valore Monetario Atteso. Trasforma i punteggi di rischio in valuta, creando un *Fondo Imprevisti* dinamico.
                * **EAC Risk-Adjusted:** L'EAC classico sommato al Fondo Imprevisti. È lo scenario finanziario più prudente.
                """)
            
        # --- SEZIONE 3: FORMULARIO ---
        with t_sec3:
            st.subheader("Matematica e Indicatori di Performance (EVM)")
            st.markdown("Il motore calcola in tempo reale lo stato di salute del cantiere utilizzando le metriche standard internazionali dell'**Earned Value Management** e della finanza di progetto:")
            with st.expander("Valori Base (I Pilastri)"):
                st.latex(r"EV = BAC \times \% \text{ Avanzamento Fisico}")
                st.caption("**EV (Earned Value)")
                st.markdown("Planned Value: *sono i costi definiti da progetto/computo, compilati manualmente in WBS*")
                st.markdown("BAC (Budget at Completion): *Il Budget totale pianificato*")
                
                
            with st.expander("Indicatori di Performance (KPI)"):
                st.latex(r"CV = EV - AC")
                st.caption("**CV (Cost Variance):** Variazione dei costi (Valore Guadagnato meno Costo Reale).")
                st.latex(r"SV = EV - PV")
                st.caption("**SV (Schedule Variance)")
                st.latex(r"CPI = \frac{EV}{AC}")
                st.caption("**CPI (Cost Performance Index):** Efficienza economica. Se $< 0.95$, si sta spendendo più del budget previsto.")
                st.latex(r"SPI = \frac{EV}{PV}")
                st.caption("**SPI (Schedule Performance Index):** Efficienza temporale. Se $< 0.95$, il cantiere è in ritardo rispetto al cronoprogramma.")
                
            with st.expander("Previsioni e Rischio (Proiezioni)"):
                st.latex(r"EAC = \frac{BAC}{CPI}")
                st.caption("**EAC (Estimate At Completion)")
                st.latex(r"ETC = EAC - AC")
                st.caption("**ETC (Estimate To Complete)")
                st.latex(r"VAC = BAC - EAC")
                st.caption("**VAC (Variance At Completion)")    
                st.latex(r"\text{EAC Risk-Adjusted} = EAC + \sum (\text{Budget}_\text{WBS} \times \text{Probabilità}_\text{Rischio} \times \text{Impatto}_\text{Rischio})")
                st.caption("**Quanta liquidità manca al competamento in considerazione dei rischi nella loro probabilità e impatto")
                st.latex(r"CF_{netto} = \sum Entrate_{SAL} - \sum Uscite_{AC}")
                st.caption("**Cash Flow Netto:** Esposizione di cassa al netto dei pagamenti ricevuti.")

        # --- SEZIONE 4: CASISTICHE E FALSI INGHIPPI ---
        with t_sec4:
            st.subheader("🔍 Guida pratica ai 'Falsi Ingippi' e Blocchi di Sicurezza")
            st.markdown("In questa sezione spieghiamo i comportamenti automatizzati del software che potrebbero sembrare anomalie, ma che in realtà sono **controlli di sicurezza attivi**.")

            with st.expander("🚧 1. Il mistero del '99%' (Blocco Qualità CAPA)"):
                st.markdown("""
                * **La Situazione:** Hai impostato una lavorazione al `100%` nel Tab 1, ma il sistema la corregge d'ufficio al `99%` e ti mostra un errore rosso.
                * **Perché accade:** C'è una **CAPA (Non-Conformità) aperta o in lavorazione** nel Tab 7 associata a quella specifica WBS. 
                * **La Logica:** Il software impedisce al Direttore Lavori di chiudere contabilmente un'attività finché il problema di qualità o sicurezza non è stato formalmente risolto.
                * **Come sbloccarlo:** Vai nel Tab 7, verifica l'azione correttiva e imposta lo stato su **'Chiuso'**. Torna nel Tab 1 e potrai finalmente certificare il 100%.
                """)

            with st.expander("📉 2. Il grafico del Cash Flow non si aggiorna o mostra solo un punto"):
                st.markdown("""
                * **La Situazione:** Hai inserito i SAL nel Tab 6, ma il grafico a gradoni non mostra le linee o la curva delle uscite è piatta.
                * **Perché accade:** 
                  1. Nel Tab 6 (Entrate), hai digitato i dati ma **non hai cliccato il bottone rosso di salvataggio** (verifica la presenza di eventuali triangolini rossi o scritte `None` nelle celle).
                  2. Le uscite non appaiono perché non sono state registrate nella sezione *Uscite* del Tab 6 con una data e una colonna di importo valide (`Importo_Netto`), ma sono state lasciate come semplici stime di budget.
                * **Come risolverlo:** Clicca sempre i bottoni di salvataggio dedicati e compila le date puntuali nel registro contabile.
                """)

            with st.expander("👷 3. Allarmi di Sovraccarico Risorse nel Radar"):
                st.markdown("""
                * **La Situazione:** Il Radar di sinistra urla che l'Impresa X è in sovraccarico.
                * **Perché accade:** Due lavorazioni distinte assegnate alla stessa risorsa hanno date d'inizio e fine sovrapposte nel Tab 1.
                * **La Soluzione:** Se la sovrapposizione è voluta (es. l'impresa ha più squadre che lavorano in contemporanea su stanze diverse), non devi modificare il calendario: ti basta cliccare il bottone **'👁️ Consenti Sovrapposizione'** sotto l'allarme per registrarla come eccezione autorizzata.
                """)

        # --- SEZIONE 5: ROADMAP VERSIONE 2.0 ---
        with t_sec5:
            st.subheader("🚀 Prossimi Sviluppi (Roadmap v2.0)")
            st.markdown("Ecco le funzioni di livello Enterprise che verranno integrate nelle prossime release autonome:")
            
            st.info("""
            * **📥 Importatore Nativo PriMus (.csv / .xls):** Integrazione diretta per mappare i computi metrici estimativi di ACCA software direttamente sull'albero WBS e sui budget di cantiere.
            * **🌐 Viewer BIM 4D / SLAM integrato:** Collegamento delle nuvole di punti 3D e dei virtual tour laser scanner direttamente alle singole voci di stato d'avanzamento.
            * **📄 Generatore Automatico del Giornale dei Lavori:** Esportazione in PDF impaginata con loghi, verbali di cantiere, firme e grafici di cash flow pronti per la DL.
            * **📊 Modulo Finanziario Immobiliare Avanzato:** Calcolo automatico del VAN (Valore Attivo Netto) e del TIR (Tasso Interno di Rendimento) per operazioni di sviluppo immobiliare.
            """)
