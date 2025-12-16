import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(
    page_title="Pianificazione Produzione Tessile",
    page_icon="🧵",
    layout="wide"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1E3A5F;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffc107;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .critical-box {
        background-color: #8B0000;
         color: #FFFFFF;
        border: 1px solid #dc3545;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #28a745;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    @media print {
        .stButton, .stFileUploader, .stTabs, .stSidebar {
            display: none !important;
        }
    }
</style>
""", unsafe_allow_html=True)

COLUMN_MAPPING = {
    1: 'ID_Cartellino',      # B: M25_NR_CARTELLINO
    2: 'Cliente',            # C: EW2_COD_CLIFOR
    3: 'Articolo',           # D: Z01_CD_ART
    6: 'Linea',              # G: Linea
    7: 'Macro_Fase',         # H: Macro_Fase
    8: 'M24_QT_SALDO',       # I: M24_QT_SALDO
    26: 'min_prd',           # Colonna 26: min_prd
    38: 'ritardo_cartellino' # Colonna 38: ritardo cartellino.1
}

MIN_REQUIRED_COLUMNS = 39

def get_priority_status(delay):
    if delay > 10:
        return "🔴 Critico"
    elif delay > 5:
        return "🟠 Alto"
    elif delay > 0:
        return "🟡 Medio"
    else:
        return "🟢 Normale"

def get_priority_color(delay):
    if delay > 10:
        return "#dc3545"
    elif delay > 5:
        return "#fd7e14"
    elif delay > 0:
        return "#ffc107"
    else:
        return "#28a745"

def calculate_completion_date(min_prd, delay, working_hours_per_day=8):
    days_needed = min_prd / (working_hours_per_day * 60)
    total_days = days_needed + max(0, delay)
    completion_date = datetime.now() + timedelta(days=total_days)
    return completion_date

def create_excel_download(df, sheet_name="Dati"):
    output = BytesIO()
    if df.empty:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        return output
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#1E3A5F',
            'font_color': 'white',
            'border': 1
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            col_max_len = df[value].astype(str).apply(len).max() if len(df) > 0 else 0
            max_len = max(col_max_len, len(str(value))) + 2
            worksheet.set_column(col_num, col_num, min(max_len, 30))
    output.seek(0)
    return output

def process_dataframe_by_position(df):
    if len(df.columns) < MIN_REQUIRED_COLUMNS:
        return None, f"Il file deve contenere almeno {MIN_REQUIRED_COLUMNS} colonne. Trovate: {len(df.columns)}"
    
    processed_df = pd.DataFrame()
    for col_index, col_name in COLUMN_MAPPING.items():
        processed_df[col_name] = df.iloc[:, col_index]
    
    for col in ['min_prd', 'ritardo_cartellino', 'M24_QT_SALDO']:
        processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce').fillna(0)
    
    return processed_df, "OK"

st.markdown('<p class="main-header">Pianificazione Produzione Tessile</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Carica i dati di produzione per generare ordini di lavoro, dashboard gestionale e stime di consegna</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Carica file Excel (.xlsx)",
    type=['xlsx'],
    help="Il file deve contenere almeno 8 colonne nell'ordine specificato (A-H)"
)

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        with st.expander("🔧 Dettagli tecnici colonne", expanded=False):
            st.markdown("**Colonne trovate nel file:**")
            col_info = []
            for i, col_name in enumerate(df_raw.columns):
                col_letter = chr(65 + i) if i < 26 else f"Col{i}"
                sample_values = df_raw.iloc[:3, i].tolist() if len(df_raw) > 0 else []
                col_info.append({
                    'Indice': i,
                    'Lettera Excel': col_letter,
                    'Nome Colonna': str(col_name),
                    'Esempio Valori': str(sample_values[:3])
                })
            st.dataframe(pd.DataFrame(col_info), use_container_width=True, hide_index=True)
            
            st.markdown("**Mapping attuale usato dal sistema:**")
            mapping_info = []
            for idx, internal_name in COLUMN_MAPPING.items():
                col_letter = chr(65 + idx) if idx < 26 else f"Col{idx}"
                actual_col = df_raw.columns[idx] if idx < len(df_raw.columns) else "N/A"
                mapping_info.append({
                    'Indice': idx,
                    'Lettera': col_letter,
                    'Colonna Reale nel File': str(actual_col),
                    'Nome Interno Sistema': internal_name
                })
            st.dataframe(pd.DataFrame(mapping_info), use_container_width=True, hide_index=True)
        
        df, message = process_dataframe_by_position(df_raw)
        
        if df is None:
            st.error(f"Errore nei dati: {message}")
        else:
            st.success(f"File caricato con successo! {len(df)} righe trovate.")
            
            tab1, tab2, tab3, tab4 = st.tabs([
                "📋 Ordini di Lavoro", 
                "📊 Dashboard Gestionale", 
                "📅 Stime Consegna",
                "🧮 Simula Nuovo Ordine"
            ])
            
            with tab1:
                st.markdown("### Ordini di Lavoro per Linea")
                st.markdown("*Ordinati per ritardo (priorità decrescente)*")
                
                work_orders_list = []
                
                for linea in sorted(df['Linea'].astype(str).unique()):
                    linea_df = df[df['Linea'].astype(str) == linea].copy()
                    linea_df = linea_df.sort_values('ritardo_cartellino', ascending=False)
                    
                    st.markdown(f"#### Linea: {linea}")
                    
                    display_df = linea_df[['ID_Cartellino', 'Cliente', 'Articolo', 'Macro_Fase', 'min_prd', 'ritardo_cartellino']].copy()
                    display_df['Priorità'] = display_df['ritardo_cartellino'].apply(get_priority_status)
                    display_df = display_df.rename(columns={
                        'ID_Cartellino': 'ID',
                        'Macro_Fase': 'Fase',
                        'min_prd': 'Minuti Produzione',
                        'ritardo_cartellino': 'Ritardo (giorni)'
                    })
                    
                    st.dataframe(
                        display_df,
                        use_container_width=True,
                        hide_index=True
                    )
                    
                    for _, row in linea_df.iterrows():
                        work_orders_list.append({
                            'Linea': linea,
                            'ID': row['ID_Cartellino'],
                            'Cliente': row['Cliente'],
                            'Articolo': row['Articolo'],
                            'Fase': row['Macro_Fase'],
                            'Minuti Produzione': row['min_prd'],
                            'Ritardo (giorni)': row['ritardo_cartellino'],
                            'Priorità': get_priority_status(row['ritardo_cartellino'])
                        })
                    
                    st.markdown("---")
                
                work_orders_df = pd.DataFrame(work_orders_list)
                excel_data = create_excel_download(work_orders_df, "Ordini_Lavoro")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 Scarica Ordini di Lavoro (Excel)",
                        data=excel_data,
                        file_name=f"ordini_lavoro_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col2:
                    st.button("🖨️ Stampa (Ctrl+P)", help="Usa Ctrl+P per stampare questa pagina")
            
            with tab2:
                st.markdown("### Dashboard Gestionale")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    total_orders = len(df)
                    st.metric("Ordini Totali", f"{total_orders:,}")
                
                with col2:
                    total_min = df['min_prd'].sum()
                    total_hours = total_min / 60
                    st.metric("Ore Produzione Totali", f"{total_hours:,.1f}")
                
                with col3:
                    avg_delay = df['ritardo_cartellino'].mean()
                    st.metric("Ritardo Medio (giorni)", f"{avg_delay:.1f}")
                
                with col4:
                    critical_orders = len(df[df['ritardo_cartellino'] > 10])
                    st.metric("Ordini Critici", f"{critical_orders}")
                
                st.markdown("---")
                
                st.markdown("#### Riepilogo per Linea")
                
                linea_summary = df.groupby(df['Linea'].astype(str)).agg({
                    'ID_Cartellino': 'count',
                    'min_prd': 'sum',
                    'ritardo_cartellino': 'mean',
                    'M24_QT_SALDO': 'sum'
                }).reset_index()
                
                linea_summary.columns = ['Linea', 'N. Ordini', 'Minuti Totali', 'Ritardo Medio', 'Quantità Saldo']
                linea_summary['Ore Totali'] = (linea_summary['Minuti Totali'] / 60).round(1)
                linea_summary['Ritardo Medio'] = linea_summary['Ritardo Medio'].round(1)
                
                st.dataframe(
                    linea_summary[['Linea', 'N. Ordini', 'Ore Totali', 'Ritardo Medio', 'Quantità Saldo']],
                    use_container_width=True,
                    hide_index=True
                )
                
                st.markdown("---")
                st.markdown("#### Analisi Colli di Bottiglia")
                
                bottleneck_threshold = linea_summary['Ore Totali'].mean() * 1.5
                bottlenecks = linea_summary[linea_summary['Ore Totali'] > bottleneck_threshold]
                
                if len(bottlenecks) > 0:
                    for _, row in bottlenecks.iterrows():
                        st.markdown(f"""
                        <div class="critical-box">
                            <strong>⚠️ Collo di Bottiglia: Linea {row['Linea']}</strong><br>
                            Carico: {row['Ore Totali']:.1f} ore ({row['N. Ordini']} ordini) - 
                            Ritardo medio: {row['Ritardo Medio']:.1f} giorni
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div class="success-box">
                        <strong>✅ Nessun collo di bottiglia rilevato</strong><br>
                        Il carico di lavoro è distribuito uniformemente tra le linee.
                    </div>
                    """, unsafe_allow_html=True)
                
                high_delay_linee = linea_summary[linea_summary['Ritardo Medio'] > 5]
                if len(high_delay_linee) > 0:
                    for _, row in high_delay_linee.iterrows():
                        st.markdown(f"""
                        <div class="warning-box">
                            <strong>⏰ Ritardo Elevato: Linea {row['Linea']}</strong><br>
                            Ritardo medio: {row['Ritardo Medio']:.1f} giorni - 
                            Necessaria attenzione prioritaria
                        </div>
                        """, unsafe_allow_html=True)
                
                dashboard_excel = create_excel_download(linea_summary, "Dashboard")
                st.download_button(
                    label="📥 Scarica Report Dashboard (Excel)",
                    data=dashboard_excel,
                    file_name=f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with tab3:
                st.markdown("### Stime di Consegna")
                
                st.markdown("#### Parametri di Calcolo")
                working_hours = st.slider("Ore lavorative giornaliere", 4, 12, 8)
                
                st.markdown("---")
                
                delivery_data = []
                
                for _, row in df.iterrows():
                    completion_date = calculate_completion_date(
                        row['min_prd'], 
                        row['ritardo_cartellino'],
                        working_hours
                    )
                    days_to_complete = (completion_date - datetime.now()).days
                    
                    delivery_data.append({
                        'ID': row['ID_Cartellino'],
                        'Cliente': row['Cliente'],
                        'Articolo': row['Articolo'],
                        'Linea': row['Linea'],
                        'Minuti Produzione': row['min_prd'],
                        'Ritardo Attuale': row['ritardo_cartellino'],
                        'Data Completamento Stimata': completion_date.strftime('%d/%m/%Y'),
                        'Giorni al Completamento': days_to_complete,
                        'Priorità': get_priority_status(row['ritardo_cartellino'])
                    })
                
                delivery_df = pd.DataFrame(delivery_data)
                delivery_df = delivery_df.sort_values('Giorni al Completamento', ascending=False)
                
                st.markdown("#### Riepilogo Consegne per Cliente")
                
                cliente_summary = delivery_df.groupby('Cliente').agg({
                    'ID': 'count',
                    'Minuti Produzione': 'sum',
                    'Giorni al Completamento': 'max'
                }).reset_index()
                cliente_summary.columns = ['Cliente', 'N. Ordini', 'Minuti Totali', 'Giorni Max Completamento']
                cliente_summary = cliente_summary.sort_values('Giorni Max Completamento', ascending=False)
                
                st.dataframe(cliente_summary, use_container_width=True, hide_index=True)
                
                st.markdown("---")
                st.markdown("#### Dettaglio Consegne")
                
                linea_options = ['Tutte'] + sorted(df['Linea'].astype(str).unique().tolist())
                selected_linea = st.selectbox(
                    "Filtra per Linea",
                    options=linea_options
                )
                
                if selected_linea != 'Tutte':
                    filtered_df = delivery_df[delivery_df['Linea'].astype(str) == selected_linea]
                else:
                    filtered_df = delivery_df
                
                st.dataframe(
                    filtered_df,
                    use_container_width=True,
                    hide_index=True
                )
                
                delivery_excel = create_excel_download(delivery_df, "Stime_Consegna")
                st.download_button(
                    label="📥 Scarica Stime Consegna (Excel)",
                    data=delivery_excel,
                    file_name=f"stime_consegna_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with tab4:
                st.markdown("### Simula Nuovo Ordine")
                st.markdown("*Calcola la data di consegna stimata per un nuovo ordine*")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    tipologia = st.selectbox(
                        "Tipologia Lavorazione",
                        options=["BIANCO", "TINTO", "FINISSAGGIO"],
                        help="Seleziona il tipo di lavorazione"
                    )
                
                with col2:
                    metri = st.number_input(
                        "Metri da produrre",
                        min_value=100,
                        max_value=1000000,
                        value=10000,
                        step=1000,
                        help="Inserisci la quantità in metri"
                    )
                
                st.markdown("---")
                
                METRI_GIORNO = 100000
                total_min_prd = df['min_prd'].sum()
                carico_attuale_giorni = total_min_prd / (8 * 60)
                
                giorni_nuovo_ordine = metri / METRI_GIORNO
                giorni_totali = carico_attuale_giorni + giorni_nuovo_ordine
                giorni_con_buffer = giorni_totali * 1.20
                
                data_consegna = datetime.now() + timedelta(days=giorni_con_buffer)
                
                st.markdown("#### Risultato Simulazione")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Carico Attuale", f"{carico_attuale_giorni:.1f} giorni")
                with col2:
                    st.metric("Tempo Nuovo Ordine", f"{giorni_nuovo_ordine:.1f} giorni")
                with col3:
                    st.metric("Buffer Prudenziale", "+20%")
                
                st.markdown(f"""
                <div class="success-box" style="text-align: center; font-size: 1.3rem;">
                    <strong>📅 Data Consegna Stimata: {data_consegna.strftime('%d/%m/%Y')}</strong><br>
                    <span style="font-size: 0.9rem;">Tipologia: {tipologia} | Metri: {metri:,} | Giorni totali: {giorni_con_buffer:.1f}</span>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("---")
                st.markdown("**Parametri di calcolo:**")
                st.markdown(f"""
                - Capacità produttiva giornaliera: **{METRI_GIORNO:,} metri/giorno**
                - Carico produzione attuale: **{total_min_prd:,.0f} minuti** ({carico_attuale_giorni:.1f} giorni)
                - Buffer prudenziale: **+20%** sul tempo totale
                """)
                
    except Exception as e:
        st.error(f"Errore durante l'elaborazione del file: {str(e)}")
        st.info("Verifica che il file sia un documento Excel valido (.xlsx)")

else:
    st.info("👆 Carica un file Excel per iniziare l'analisi della produzione")
    
    with st.expander("ℹ️ Formato file richiesto"):
        st.markdown("""
        Il file Excel deve contenere le colonne nell'ordine seguente (per posizione):
        
        | Posizione | Colonna Excel | Descrizione |
        |-----------|---------------|-------------|
        | **0** | A | ID Cartellino - Identificativo univoco dell'ordine |
        | **1** | B | Cliente - Nome del cliente |
        | **2** | C | Articolo - Codice o descrizione articolo |
        | **3** | D | Linea - Linea di produzione |
        | **4** | E | min_prd - Minuti di produzione necessari |
        | **5** | F | ritardo_cartellino - Giorni di ritardo (positivo = in ritardo) |
        | **6** | G | Macro_Fase - Fase di lavorazione |
        | **7** | H | M24_QT_SALDO - Quantità saldo |
        
        **Nota:** I nomi delle intestazioni nel file Excel possono essere qualsiasi. 
        Il sistema legge i dati in base alla posizione delle colonne (A, B, C, ecc.).
        """)

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #888; font-size: 0.9rem;'>"
    "Sistema di Pianificazione Produzione Tessile | Sviluppato da Riccardo Spinello | © 2025"
    "</p>",
    unsafe_allow_html=True
)
