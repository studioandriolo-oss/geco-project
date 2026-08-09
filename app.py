import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import graphviz
from datetime import datetime, date

# Configurazione Pagina
st.set_page_config(page_title="WBS/OBS Manager & EVM", layout="wide")
st.title("🏗️ Project Workflow & EVM Controller")

# --- 1. INIZIALIZZAZIONE DATI (Session State) ---
# Creiamo dei dati di base se l'app viene aperta per la prima volta
if 'wbs_data' not in st.session_state:
    st.session_state.wbs_data = pd.DataFrame({
        'ID_WBS': ['2.1', '2.2', '2.3'],
        'Attività': ['Scavi', 'Rinforzo Strutturale P1', 'Getto Fondazioni'],
        'Inizio_Previsto': [date(2026, 9, 1), date(2026, 9, 15), date(2026, 10, 1)],
        'Fine_Prevista': [date(2026, 9, 14), date(2026, 10, 15), date(2026, 10, 10)],
        'Inizio_Effettivo': [date(2026, 9, 2), date(2026, 9, 18), None],
        'Fine_Effettiva': [date(2026, 9, 16), None, None],
        'BAC_Budget': [15000.0, 25000.0, 30000.0],
        '%_Completamento': [100, 40, 0],
        'AC_Costo_Reale': [15500.0, 12000.0, 0.0],
        'ID_OBS_Assegnato': ['1.1', '1.2', '1.1'] # Incrocio con OBS
    })

if 'obs_data' not in st.session_state:
    st.session_state.obs_data = pd.DataFrame({
        'ID_OBS': ['1.1', '1.2'],
        'Ruolo': ['Capo Cantiere', 'Strutturista'],
        'Risorsa': ['Mario Rossi', 'Studio Tecnico']
    })

# Calcoli EVM Dinamici sul DataFrame
def calcola_evm(df):
    # Calcolo PV (Planned Value) semplificato: proporzionale ai giorni trascorsi o 100% se data superata
    # Nota: In una versione avanzata qui si usa la data di controllo (status date)
    df['EV'] = df['BAC_Budget'] * (df['%_Completamento'] / 100)
    df['CV'] = df['EV'] - df['AC_Costo_Reale'] # Cost Variance
    
    # Prevenzione divisione per zero
    df['SPI'] = df.apply(lambda x: (x['EV'] / x['BAC_Budget']) if x['BAC_Budget'] > 0 else 1, axis=1) # Semplificazione per demo
    df['CPI'] = df.apply(lambda x: (x['EV'] / x['AC_Costo_Reale']) if x['AC_Costo_Reale'] > 0 else 1, axis=1)
    return df

st.session_state.wbs_data = calcola_evm(st.session_state.wbs_data)

# --- CREAZIONE TAB ---
tab1, tab2, tab3, tab4 = st.tabs(["🗂️ Setup WBS/OBS", "🕸️ Nodi & Matrice", "📅 Cronoprogramma", "📈 EVM & Cash Flow"])

# --- TAB 1: SETUP E INSERIMENTO DATI ---
with tab1:
    st.header("Compilazione Strutture")
    st.subheader("WBS (Work Breakdown Structure)")
    # Editor interattivo tipo foglio di calcolo
    edited_wbs = st.data_editor(st.session_state.wbs_data, num_rows="dynamic", use_container_width=True)
    st.session_state.wbs_data = edited_wbs

    st.subheader("OBS (Organization Breakdown Structure)")
    edited_obs = st.data_editor(st.session_state.obs_data, num_rows="dynamic", use_container_width=True)
    st.session_state.obs_data = edited_obs

# --- TAB 2: MATRICE E GRAFO A NODI ---
with tab2:
    st.header("Incrocio Logico (Work Packages)")
    st.markdown("Generazione automatica dei nodi di collegamento tra risorse (OBS) e attività (WBS).")
    
    # Creazione Grafo con Graphviz (simile alla logica nodale)
    graph = graphviz.Digraph(engine='dot')
    graph.attr(rankdir='LR') # Da sinistra a destra
    
    # Nodi OBS
    for _, row in st.session_state.obs_data.iterrows():
        graph.node(row['ID_OBS'], f"{row['Ruolo']}\n({row['Risorsa']})", shape='box', style='filled', fillcolor='lightblue')
        
    # Nodi WBS e Connessioni (I Work Packages)
    for _, row in st.session_state.wbs_data.iterrows():
        wp_label = f"WP: {row['Attività']}\nBudget: €{row['BAC_Budget']}"
        graph.node(
        row['ID_WBS'], 
        wp_label, 
        shape='Mrecord', # Puoi cambiarlo in 'box', 'rect', o 'ellipse' 
        style='filled', 
        fillcolor='lightgreen',
        width='2.5',     # <--- Larghezza minima (in pollici)
        height='1.2',    # <--- Altezza minima (in pollici)
        fontsize='12'    # <--- Dimensione del testo
    )
        
        # Connessione
        if pd.notna(row['ID_OBS_Assegnato']):
            obs_ids = str(row['ID_OBS_Assegnato']).split(',') # Permette assegnazioni multiple (es: "1.1, 1.2")
            for o_id in obs_ids:
                graph.edge(o_id.strip(), row['ID_WBS'])
                
    st.graphviz_chart(graph, use_container_width=True)

# --- TAB 3: CRONOPROGRAMMA (GANTT) ---
with tab3:
    st.header("Cronoprogramma Lavori")
    vista = st.selectbox("Seleziona Vista", ["Progetto (Baseline)", "Esecuzione (As-Built)", "Comparativa"])
    
    df_gantt = st.session_state.wbs_data.copy()
    
    # Assicuriamoci che le date siano datetime per Plotly
    df_gantt['Inizio_Previsto'] = pd.to_datetime(df_gantt['Inizio_Previsto'])
    df_gantt['Fine_Prevista'] = pd.to_datetime(df_gantt['Fine_Prevista'])
    df_gantt['Inizio_Effettivo'] = pd.to_datetime(df_gantt['Inizio_Effettivo'])
    df_gantt['Fine_Effettiva'] = pd.to_datetime(df_gantt['Fine_Effettiva']).fillna(pd.Timestamp.now())
    
    fig = go.Figure()
    
    if vista in ["Progetto (Baseline)", "Comparativa"]:
        # CORREZIONE: Convertiamo il Timedelta (durata) in millisecondi
        durata_prevista_ms = (df_gantt['Fine_Prevista'] - df_gantt['Inizio_Previsto']).dt.total_seconds() * 1000
        
        fig.add_trace(go.Bar(
            x=durata_prevista_ms,
            y=df_gantt['Attività'],
            base=df_gantt['Inizio_Previsto'],
            orientation='h',
            name='Baseline',
            width=0.4; # <--- spessore delle righe del cronoprogramma
            marker=dict(color='rgba(0, 0, 255, 0.4)') if vista == "Comparativa" else dict(color='blue')
        ))
        
    if vista in ["Esecuzione (As-Built)", "Comparativa"]:
        # Filtra solo i task iniziati e crea una copia esplicita per evitare warning
        df_esec = df_gantt.dropna(subset=['Inizio_Effettivo']).copy()
        
        # CORREZIONE: Convertiamo il Timedelta (durata effettiva) in millisecondi
        durata_effettiva_ms = (df_esec['Fine_Effettiva'] - df_esec['Inizio_Effettivo']).dt.total_seconds() * 1000
        
        fig.add_trace(go.Bar(
            x=durata_effettiva_ms,
            y=df_esec['Attività'],
            base=df_esec['Inizio_Effettivo'],
            orientation='h',
            name='As-Built',
            marker=dict(color='red')
        ))
        
    fig.update_layout(
        barmode='overlay', 
        height=600, # <--- altezza tabella del cronoprogramma
        bargap=0.3, # <--- Spazio vuoto tra le barre: 0.0 è tutto unito, 0.5 è metà vuoto
        xaxis_title="Linea Temporale", 
        yaxis_title="WBS", 
        yaxis={'autorange': 'reversed'},
        xaxis_type='date' # Forza Plotly a leggere l'asse X (millisecondi + base) come calendario
    )
    st.plotly_chart(fig, use_container_width=True)

# --- TAB 4: EVM E CASH FLOW ---
with tab4:
    st.header("Controllo Costi e Analisi EVM")
    
    df_evm = st.session_state.wbs_data.copy()
    
    # Metriche Globali di Progetto
    tot_bac = df_evm['BAC_Budget'].sum()
    tot_ev = df_evm['EV'].sum()
    tot_ac = df_evm['AC_Costo_Reale'].sum()
    
    cpi_globale = tot_ev / tot_ac if tot_ac > 0 else 1
    
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("Budget Totale (BAC)", f"€ {tot_bac:,.2f}")
    col_m2.metric("Lavoro Eseguito (EV)", f"€ {tot_ev:,.2f}")
    col_m3.metric("Costi Sostenuti (AC)", f"€ {tot_ac:,.2f}")
    col_m4.metric("CPI Globale", f"{cpi_globale:.2f}", 
                  delta="Over-budget" if cpi_globale < 1 else "Under-budget", 
                  delta_color="inverse")
    
    st.divider()
    
    col_chart, col_table = st.columns([5, 5])
    
    with col_chart:
        st.subheader("Raffronto Costi per Attività")
        # Grafico a barre raggruppate per visualizzare BAC, EV, AC
        fig_evm = go.Figure(data=[
            go.Bar(name='BAC (Budget)', x=df_evm['Attività'], y=df_evm['BAC_Budget'], marker_color='lightgray'),
            go.Bar(name='EV (Valore Guadagnato)', x=df_evm['Attività'], y=df_evm['EV'], marker_color='green'),
            go.Bar(name='AC (Costo Reale)', x=df_evm['Attività'], y=df_evm['AC_Costo_Reale'], marker_color='red')
        ])
        fig_evm.update_layout(barmode='group')
        st.plotly_chart(fig_evm, use_container_width=True)
        
    with col_table:
        st.subheader("Indicatori di Performance (KPI)")
        # Tabella per mostrare lo stato di salute di ogni WP
        df_kpi = df_evm[['Attività', '%_Completamento', 'CPI', 'SPI', 'CV']].copy()
        
        # Formattazione condizionale per evidenziare i problemi (Stile Pandas)
        def color_kpi(val):
            if isinstance(val, (int, float)):
                if val < 1.0: return 'color: red'
                elif val >= 1.0: return 'color: green'
            return ''
        
        st.dataframe(df_kpi.style.map(color_kpi, subset=['CPI', 'SPI'])
                            .format({'CPI': "{:.2f}", 'SPI': "{:.2f}", 'CV': "€ {:.2f}"}), 
                     use_container_width=True)
            
        st.dataframe(df_kpi.style.map(color_kpi, subset=['CPI', 'SPI'])
                            .format({'CPI': "{:.2f}", 'SPI': "{:.2f}", 'CV': "€ {:.2f}"}), 
                     use_container_width=True)
