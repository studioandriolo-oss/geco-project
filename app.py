import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import graphviz
import streamlit.components.v1 as components
from datetime import datetime, date

# Configurazione Pagina
st.set_page_config(page_title="WBS/OBS Manager & EVM", layout="wide")
st.title("🏗️ Project Workflow & EVM Controller")

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
    
# ---  INIZIALIZZAZIONE REGISTRO CONTABILE ---
if 'registro_data' not in st.session_state:
    st.session_state.registro_data = pd.DataFrame({
        'Data': [date(2026, 9, 5), date(2026, 9, 10)],
        'N_Doc': ['FATT-01', 'FATT-02'],
        'Fornitore': ['Mario Rossi', 'Nolo Scavi Srl'],
        'Voce_WBS': ['1.1 - Scavi con mezzi meccanici', '1.1 - Scavi con mezzi meccanici'], # Voci a tendina
        'Importo_Netto': [2000.0, 3200.0],
        'Descrizione': ['Acconto lavori', 'Nolo escavatore']
    })

# --- MOTORE AGGIORNAMENTO COSTI REALI (Da Tab 6 a Tab 1) ---
def aggiorna_costi_reali():
    df_reg = st.session_state.registro_data.copy()
    # Estraiamo l'ID (es. "1.1") dalla voce a tendina (es. "1.1 - Scavi con mezzi meccanici")
    df_reg['ID_WBS_calc'] = df_reg['Voce_WBS'].astype(str).apply(lambda x: x.split(' - ')[0] if ' - ' in x else None)
    
    # Sommiamo gli importi netti raggruppandoli per ID_WBS
    costi_raggruppati = df_reg.groupby('ID_WBS_calc')['Importo_Netto'].sum().reset_index()
    cost_map = dict(zip(costi_raggruppati['ID_WBS_calc'], costi_raggruppati['Importo_Netto']))
    
    # Applichiamo i costi calcolati alla tabella WBS
    wbs = st.session_state.wbs_data
    wbs['AC_Costo_Reale'] = wbs['ID_WBS'].apply(lambda x: cost_map.get(str(x), 0.0))
    st.session_state.wbs_data = wbs

# Eseguiamo i calcoli in sequenza prima di disegnare l'interfaccia
aggiorna_costi_reali()

# Calcoli EVM Dinamici sul DataFrame
def calcola_evm(df):
    df['EV'] = df['BAC_Budget'] * (df['%_Completamento'] / 100)
    df['CV'] = df['EV'] - df['AC_Costo_Reale']
    df['SPI'] = df.apply(lambda x: (x['EV'] / x['BAC_Budget']) if x['BAC_Budget'] > 0 else 1, axis=1)
    df['CPI'] = df.apply(lambda x: (x['EV'] / x['AC_Costo_Reale']) if x['AC_Costo_Reale'] > 0 else 1, axis=1)
    return df

st.session_state.wbs_data = calcola_evm(st.session_state.wbs_data)

# Calcoli EVM Dinamici sul DataFrame
def calcola_evm(df):
    df['EV'] = df['BAC_Budget'] * (df['%_Completamento'] / 100)
    df['CV'] = df['EV'] - df['AC_Costo_Reale']
    df['SPI'] = df.apply(lambda x: (x['EV'] / x['BAC_Budget']) if x['BAC_Budget'] > 0 else 1, axis=1)
    df['CPI'] = df.apply(lambda x: (x['EV'] / x['AC_Costo_Reale']) if x['AC_Costo_Reale'] > 0 else 1, axis=1)
    return df

st.session_state.wbs_data = calcola_evm(st.session_state.wbs_data)

# --- CREAZIONE TAB ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🗂️ WBS (Lavorazioni)", 
    "👥 OBS (Risorse)", 
    "🕸️ Nodi & Matrice", 
    "📅 Cronoprogramma", 
    "📈 EVM & Cash Flow",
    "🧾 Reg. Contabile"    
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
                    )
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
    st.header("Incrocio Logico (Work Packages)")
    
    mostra_relazioni = st.toggle("👁️ Mostra Relazioni tra WP (Interferenze)", value=True)
    
    graph = graphviz.Digraph(engine='dot')
    graph.attr(rankdir='LR', ranksep='1.5', nodesep='0.8', splines='spline')
    graph.attr('node', fontname='Helvetica', fontsize='10', margin='0.2')
    
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
        
    df_wp_reali = st.session_state.wbs_data[st.session_state.wbs_data['ID_WBS'].astype(str).str.contains('\.')]
    valid_wbs_ids = set(df_wp_reali['ID_WBS'].astype(str))
    
    for _, row in df_wp_reali.iterrows():
        attivita = str(row['Attività'])
        budget = row['BAC_Budget']
        
        wp_html = f"<<TABLE BORDER='0' CELLBORDER='0' CELLSPACING='4'>"
        wp_html += f"<TR><TD><B>WP: {attivita}</B></TD></TR>"
        wp_html += f"<TR><TD>Budget: &euro; {budget:,.2f}</TD></TR>"
        wp_html += "</TABLE>>"
        
        graph.node(
            f"WBS_{row['ID_WBS']}", 
            label=wp_html, 
            shape='rect', 
            style='rounded,filled', 
            fillcolor='#C8E6C9', 
            color='#388E3C',     
            penwidth='1.5'
        )
        
        if pd.notna(row['ID_OBS_Assegnato']):
            obs_ids = str(row['ID_OBS_Assegnato']).split(',')
            for o_id in obs_ids:
                if o_id.strip():
                    graph.edge(
                        f"OBS_{o_id.strip()}", 
                        f"WBS_{row['ID_WBS']}", 
                        color='#757575', 
                        penwidth='1.5',
                        arrowsize='0.8'
                    )
                    
        if mostra_relazioni and 'Predecessori' in row and pd.notna(row['Predecessori']):
            preds = str(row['Predecessori']).split(',')
            for p_id in preds:
                p_id = p_id.strip()
                if p_id in valid_wbs_ids:
                    graph.edge(
                        f"WBS_{p_id}", 
                        f"WBS_{row['ID_WBS']}", 
                        color='#FF9800',  
                        style='dashed',   
                        penwidth='1.0',   
                        arrowsize='0.6'
                    )

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
                        document.getElementById('svg-container').innerHTML = "Errore nel caricamento del grafico SVG.";
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

# --- TAB 4: CRONOPROGRAMMA (GANTT) ---
with tab4:
    st.header("Cronoprogramma Lavori")
    vista = st.selectbox("Seleziona Vista", ["Progetto (Baseline)", "Esecuzione (Esecutivo)", "Comparativa"])
    
    df_gantt = st.session_state.wbs_data.copy()
    df_gantt = df_gantt[df_gantt['ID_WBS'].astype(str).str.contains('\.')] 
    
    df_gantt['Inizio_Previsto'] = pd.to_datetime(df_gantt['Inizio_Previsto'])
    df_gantt['Fine_Prevista'] = pd.to_datetime(df_gantt['Fine_Prevista'])
    df_gantt['Inizio_Effettivo'] = pd.to_datetime(df_gantt['Inizio_Effettivo'])
    df_gantt['Fine_Effettiva'] = pd.to_datetime(df_gantt['Fine_Effettiva']).fillna(pd.Timestamp.now())
    
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
    
    df_evm = st.session_state.wbs_data.copy()
    
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
    
    st.subheader("Raffronto Costi per Attività")
        
    fig_evm = go.Figure(data=[
        go.Bar(
            name='BAC (Budget)', 
            x=df_evm['Attività'], 
            y=df_evm['BAC_Budget'], 
            marker_color='lightgray',
            text=df_evm['BAC_Budget'],         
            texttemplate='€ %{text:,.0f}',     
            textposition='outside',            
            textangle=-90                      
        ),
        go.Bar(
            name='EV (Valore Guadagnato)', 
            x=df_evm['Attività'], 
            y=df_evm['EV'], 
            marker_color='green',
            text=df_evm['EV'],
            texttemplate='€ %{text:,.0f}',
            textposition='outside',
            textangle=-90
        ),
        go.Bar(
            name='AC (Costo Reale)', 
            x=df_evm['Attività'], 
            y=df_evm['AC_Costo_Reale'], 
            marker_color='red',
            text=df_evm['AC_Costo_Reale'],
            texttemplate='€ %{text:,.0f}',
            textposition='outside',
            textangle=-90
        )
    ])
        
    fig_evm.update_layout(
        barmode='group',
        margin=dict(t=80),         
        uniformtext_minsize=9,     
        uniformtext_mode='hide'    
    )
    st.plotly_chart(fig_evm, use_container_width=True)
        
    col_KPI, col_LEGENDA = st.columns([7, 3]) 
    
    with col_KPI:
        st.subheader("Indicatori di Performance (KPI)")
        df_kpi = df_evm[['Attività', '%_Completamento', 'CPI', 'SPI', 'CV']].copy()
        
        def color_kpi(val):
            if isinstance(val, (int, float)):
                if val < 1.0: return 'color: red'
                elif val >= 1.0: return 'color: green'
            return ''
            
        st.dataframe(df_kpi.style.map(color_kpi, subset=['CPI', 'SPI'])
                            .format({'CPI': "{:.2f}", 'SPI': "{:.2f}", 'CV': "€ {:.2f}"}), 
                     use_container_width=True)

    with col_LEGENDA:
        st.subheader("Legenda")
        st.markdown("""
        * **CPI (Cost Performance Index):** Efficienza dei costi.  
        Se **< 1**, stai spendendo più del budget previsto per il lavoro svolto.
        * **SPI (Schedule Performance Index):** Efficienza temporale.  
        Se **< 1**, sei in ritardo rispetto alla programmazione.
        * **CV (Cost Variance):** Scostamento dei costi assoluto (EV - AC).  
        Un valore negativo indica una perdita monetaria sull'attività.
        """)
        
# --- TAB 6: REGISTRO CONTABILE ---
with tab6:
    st.header("Registro Contabile")
    st.markdown("Inserisci qui le fatture e i SAL. Gli importi netti si sommeranno automaticamente aggiornando la voce *AC_Costo_Reale* nella WBS.")
    
    # 1. Prepariamo dinamicamente le voci per il menu a tendina (Prendiamo solo i sotto-nodi WBS operativi)
    df_wbs = st.session_state.wbs_data
    leaf_wbs = df_wbs[df_wbs['ID_WBS'].astype(str).str.contains('\.')]
    # Creiamo una lista formattata "ID - Nome Attività" (es. "1.1 - Scavi con mezzi meccanici")
    wbs_options = [f"{row['ID_WBS']} - {row['Attività']}" for _, row in leaf_wbs.iterrows()]
    
    # 2. Creiamo la tabella di input
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
                options=wbs_options, # Passiamo la lista dinamica!
                required=True
            )
        }
    )
    
    # 3. Aggiorniamo i dati in tempo reale
    if not edited_registro.equals(st.session_state.registro_data):
        st.session_state.registro_data = edited_registro
        # Se c'è una modifica, riavvia l'app in modo che il motore in alto ricalcoli
        # i costi, li spari nel Tab 1 e ricalcoli l'EVM nel Tab 5.
        st.rerun()
