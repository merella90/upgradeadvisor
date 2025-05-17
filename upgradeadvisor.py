import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# Configurazione pagina
st.set_page_config(
    page_title="Hotel Upgrade Advisor",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Funzioni di utilità per generare il file Excel
def create_excel_file(params, historical_data, otb_data, pickup_data):
    # Creazione del workbook
    wb = Workbook()
    
    # Definizione degli stili professionali
    color_primary = "4472C4"  # Blu aziendale
    color_header_bg = "4472C4"  # Sfondo intestazioni
    color_header_font = "FFFFFF"  # Testo intestazioni
    
    # Stili di bordo
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='medium')
    )
    
    # Stili di font
    header_font = Font(name='Calibri', size=11, bold=True, color=color_header_font)
    title_font = Font(name='Calibri', size=14, bold=True, color=color_primary)
    normal_font = Font(name='Calibri', size=11)
    
    # Stili di riempimento
    header_fill = PatternFill(start_color=color_header_bg, end_color=color_header_bg, fill_type="solid")
    input_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    alt_row_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    # Allineamenti
    center_align = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    
    # --------------------------
    # 1. Foglio PARAMETRI
    # --------------------------
    ws_param = wb.active
    ws_param.title = "Parametri"
    
    # Titolo del foglio
    ws_param.merge_cells('A1:C1')
    ws_param['A1'] = "PARAMETRI DI CONFIGURAZIONE"
    ws_param['A1'].font = title_font
    ws_param['A1'].alignment = center_align
    
    # Intestazioni
    ws_param.append(["Parametro", "Valore", "Note"])
    for col in range(1, 4):
        cell = ws_param.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = center_align
    
    # Dati parametri
    parameters = [
        ["ADR Medio", params['adr'], "Tariffa media giornaliera in €"],
        ["Soglia Upgrade (fissa)", params['upgrade_threshold'], "Non utilizzata nella versione corrente"],
        ["Margine minimo richiesto", params['min_margin'], "Espresso come percentuale (0.6 = 60%)"],
        ["Giorni di analisi", params['days'], "Numero di giorni inclusi nell'analisi"]
    ]
    
    for i, (param, value, note) in enumerate(parameters, start=3):
        ws_param.cell(row=i, column=1, value=param).font = normal_font
        ws_param.cell(row=i, column=1, value=param).border = thin_border
        ws_param.cell(row=i, column=1, value=param).alignment = left_align
        
        cell = ws_param.cell(row=i, column=2, value=value)
        cell.border = thin_border
        cell.font = normal_font
        cell.alignment = right_align
        
        # Formattazione numeri in stile italiano
        if param == "ADR Medio":
            cell.number_format = '€ #.##0,00'
        elif param == "Soglia Upgrade (fissa)":
            cell.number_format = '€ #.##0,00'
        elif param == "Margine minimo richiesto":
            cell.number_format = '0%'
        
        ws_param.cell(row=i, column=3, value=note).font = normal_font
        ws_param.cell(row=i, column=3, value=note).border = thin_border
        ws_param.cell(row=i, column=3, value=note).alignment = left_align
        
        # Righe alternate per migliorare la leggibilità
        if i % 2 == 1:
            for col in range(1, 4):
                ws_param.cell(row=i, column=col).fill = alt_row_fill
    
    # Regolazione larghezza colonne
    ws_param.column_dimensions['A'].width = 25
    ws_param.column_dimensions['B'].width = 15
    ws_param.column_dimensions['C'].width = 40
    
    # --------------------------
    # 2. Foglio STORICO
    # --------------------------
    ws_storico = wb.create_sheet("Storico 2024")
    
    # Titolo del foglio
    ws_storico.merge_cells('A1:E1')
    ws_storico['A1'] = "DATI STORICI GIUGNO 2024"
    ws_storico['A1'].font = title_font
    ws_storico['A1'].alignment = center_align
    
    # Intestazioni
    headers = ["Data", "Room Nights 2024", "Domanda", "Frequenza", "Probabilità"]
    ws_storico.append(headers)
    for col in range(1, 6):
        cell = ws_storico.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = center_align
    
    # Inserimento dati storici
    for i, (date, rn) in enumerate(historical_data.items()):
        row = i + 3
        
        # Data
        date_cell = ws_storico.cell(row=row, column=1, value=date)
        date_cell.number_format = 'DD/MM/YYYY'
        date_cell.alignment = center_align
        date_cell.border = thin_border
        
        # Room Nights
        rn_cell = ws_storico.cell(row=row, column=2, value=rn)
        rn_cell.fill = input_fill
        rn_cell.border = thin_border
        rn_cell.alignment = center_align
        
        # Domanda
        dom_cell = ws_storico.cell(row=row, column=3, value=f"=B{row}")
        dom_cell.border = thin_border
        dom_cell.alignment = center_align
        
        # Frequenza
        freq_cell = ws_storico.cell(row=row, column=4, value=f"=CONTA.SE(C$3:C$32;C{row})")
        freq_cell.border = thin_border
        freq_cell.alignment = center_align
        
        # Probabilità
        prob_cell = ws_storico.cell(row=row, column=5, value=f"=D{row}/CONTA.VALORI(C$3:C$32)")
        prob_cell.number_format = '0,00%'
        prob_cell.border = thin_border
        prob_cell.alignment = center_align
        
        # Righe alternate
        if i % 2 == 1:
            for col in range(1, 6):
                ws_storico.cell(row=row, column=col).fill = alt_row_fill
    
    # --------------------------
    # 3. Foglio OTB GIUGNO 2025
    # --------------------------
    ws_otb = wb.create_sheet("OTB Giugno 2025")
    
    # Titolo del foglio
    ws_otb.merge_cells('A1:C1')
    ws_otb['A1'] = "PRENOTAZIONI ACQUISITE (ON THE BOOKS) - GIUGNO 2025"
    ws_otb['A1'].font = title_font
    ws_otb['A1'].alignment = center_align
    
    # Intestazioni
    headers = ["Data", "Room Nights", "Note"]
    ws_otb.append(headers)
    for col in range(1, 4):
        cell = ws_otb.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = center_align
    
    # Inserimento dati OTB
    for i, (date, rn) in enumerate(otb_data.items()):
        row = i + 3
        
        # Data
        date_cell = ws_otb.cell(row=row, column=1, value=date)
        date_cell.number_format = 'DD/MM/YYYY'
        date_cell.alignment = center_align
        date_cell.border = thin_border
        
        # Room Nights
        rn_cell = ws_otb.cell(row=row, column=2, value=rn)
        rn_cell.fill = input_fill
        rn_cell.border = thin_border
        rn_cell.alignment = center_align
        
        # Note (cella vuota formattata)
        note_cell = ws_otb.cell(row=row, column=3)
        note_cell.border = thin_border
        
        # Righe alternate
        if i % 2 == 1:
            for col in range(1, 4):
                ws_otb.cell(row=row, column=col).fill = alt_row_fill
    
    # --------------------------
    # 4. Foglio PICKUP vs SPIT
    # --------------------------
    ws_pickup = wb.create_sheet("Pickup vs SPIT")
    
    # Titolo del foglio
    ws_pickup.merge_cells('A1:D1')
    ws_pickup['A1'] = "CONFRONTO PICKUP A 7 GIORNI: 2025 vs 2024"
    ws_pickup['A1'].font = title_font
    ws_pickup['A1'].alignment = center_align
    
    # Intestazioni
    headers = ["Data", "Pickup 7gg 2025", "Pickup 7gg 2024", "Delta"]
    ws_pickup.append(headers)
    for col in range(1, 5):
        cell = ws_pickup.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = center_align
    
    # Inserimento dati pickup
    for i, (date, data) in enumerate(pickup_data.items()):
        row = i + 3
        
        # Data
        date_cell = ws_pickup.cell(row=row, column=1, value=date)
        date_cell.number_format = 'DD/MM/YYYY'
        date_cell.alignment = center_align
        date_cell.border = thin_border
        
        # Pickup 2025
        pickup25_cell = ws_pickup.cell(row=row, column=2, value=data['2025'])
        pickup25_cell.fill = input_fill
        pickup25_cell.border = thin_border
        pickup25_cell.alignment = center_align
        
        # Pickup 2024
        pickup24_cell = ws_pickup.cell(row=row, column=3, value=data['2024'])
        pickup24_cell.fill = input_fill
        pickup24_cell.border = thin_border
        pickup24_cell.alignment = center_align
        
        # Delta
        delta_cell = ws_pickup.cell(row=row, column=4, value=f"=B{row}-C{row}")
        delta_cell.border = thin_border
        delta_cell.alignment = center_align
        
        # Righe alternate
        if i % 2 == 1:
            for col in range(1, 5):
                ws_pickup.cell(row=row, column=col).fill = alt_row_fill
    
    # Formattazione condizionale su Delta
    ws_pickup.conditional_formatting.add(
        f"D3:D{row}", CellIsRule(operator="lessThan", formula=["0"], 
                                stopIfTrue=False, 
                                fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), 
                                font=Font(color="9C0006"))
    )
    ws_pickup.conditional_formatting.add(
        f"D3:D{row}", CellIsRule(operator="greaterThan", formula=["0"], 
                                stopIfTrue=False, 
                                fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), 
                                font=Font(color="006100"))
    )
    
    # --------------------------
    # 5. Foglio DASHBOARD
    # --------------------------
    ws_dash = wb.create_sheet("Dashboard Upgrade")
    
    # Titolo del foglio
    ws_dash.merge_cells('A1:H1')
    ws_dash['A1'] = "DASHBOARD DECISIONALE UPGRADE - GIUGNO 2025"
    ws_dash['A1'].font = title_font
    ws_dash['A1'].alignment = center_align
    
    # Intestazioni
    headers = [
        "Data", "RN Attuali", "Prob. Domanda ≥ RN", "Expected Revenue",
        "Delta Pickup vs LY", "Soglia Upgrade Dinamica (€)",
        "Upgrade Consigliato", "Suggerimento Operativo"
    ]
    ws_dash.append(headers)
    for col in range(1, 9):
        cell = ws_dash.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = center_align
    
    # Dati dashboard
    dates = list(otb_data.keys())
    for i, date in enumerate(dates):
        row = i + 3
        
        # Data
        date_cell = ws_dash.cell(row=row, column=1, value=date)
        date_cell.number_format = 'DD/MM/YYYY'
        date_cell.alignment = center_align
        date_cell.border = thin_border
        
        # RN Attuali
        rn_cell = ws_dash.cell(row=row, column=2, value=f'=CERCA.X(A{row};\'OTB Giugno 2025\'!A:A;\'OTB Giugno 2025\'!B:B;"")')
        rn_cell.border = thin_border
        rn_cell.alignment = center_align
        
        # Probabilità
        prob_cell = ws_dash.cell(row=row, column=3, value=f'=SOMMA.SE(\'Storico 2024\'!C$3:C$32;">=" & B{row};\'Storico 2024\'!E$3:E$32)')
        prob_cell.number_format = '0,00%'
        prob_cell.border = thin_border
        prob_cell.alignment = center_align
        
        # Expected Revenue
        rev_cell = ws_dash.cell(row=row, column=4, value=f'=C{row}*Parametri!$B$2')
        rev_cell.number_format = '€ #.##0,00'
        rev_cell.border = thin_border
        rev_cell.alignment = center_align
        
        # Delta Pickup
        delta_cell = ws_dash.cell(row=row, column=5, value=f'=CERCA.X(A{row};\'Pickup vs SPIT\'!A:A;\'Pickup vs SPIT\'!D:D;"")')
        delta_cell.border = thin_border
        delta_cell.alignment = center_align
        
        # Soglia Upgrade
        soglia_cell = ws_dash.cell(row=row, column=6, value=f'=Parametri!$B$2*C{row}*Parametri!$B$3')
        soglia_cell.number_format = '€ #.##0,00'
        soglia_cell.border = thin_border
        soglia_cell.alignment = center_align
        
        # Upgrade Consigliato
        upgrade_cell = ws_dash.cell(row=row, column=7, value=f'=SE(D{row}<F{row};"Sì";"No")')
        upgrade_cell.border = thin_border
        upgrade_cell.alignment = center_align
        
        # Suggerimento
        sugg_cell = ws_dash.cell(
            row=row, column=8, 
            value=(f'=SE(G{row}="Sì";'
                f'"Bassa probabilità di vendita - considera upgrade gratuito";'
                f'SE(C{row}>0,7;"Alta probabilità di vendita - non fare upgrade";'
                f'"Probabilità media - valuta disponibilità e strategia"))')
        )
        sugg_cell.border = thin_border
        sugg_cell.alignment = left_align
        
        # Righe alternate
        if i % 2 == 1:
            for col in range(1, 9):
                ws_dash.cell(row=row, column=col).fill = alt_row_fill
    
    # Formattazioni condizionali
    ws_dash.conditional_formatting.add(
        f"G3:G{row}", CellIsRule(operator="equal", formula=['"Sì"'], 
                                stopIfTrue=False, 
                                fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), 
                                font=Font(color="006100"))
    )
    ws_dash.conditional_formatting.add(
        f"G3:G{row}", CellIsRule(operator="equal", formula=['"No"'], 
                                stopIfTrue=False, 
                                fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), 
                                font=Font(color="9C0006"))
    )
    
    # Imposta dashboard come foglio attivo
    wb.active = wb["Dashboard Upgrade"]
    
    # Salva in buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer

# Funzioni di utilità per calcoli e visualizzazioni
def calculate_probabilities(historical_data):
    # Calcola frequenze
    freq = {}
    for value in historical_data.values():
        if value in freq:
            freq[value] += 1
        else:
            freq[value] = 1
    
    # Calcola probabilità
    prob = {}
    total = len(historical_data)
    for value in sorted(freq.keys(), reverse=True):
        # Probabilità cumulativa di avere richiesta >= value
        prob[value] = sum(freq[v] for v in freq if v >= value) / total
    
    return prob

def calculate_dashboard_data(historical_data, otb_data, pickup_data, params):
    # Calcola probabilità dalla storico
    probabilities = calculate_probabilities(historical_data)
    
    dashboard_data = []
    for date, rn in otb_data.items():
        # Probabilità che la domanda sia >= rn
        prob = 0
        for hist_rn, hist_prob in probabilities.items():
            if hist_rn >= rn:
                prob = hist_prob
                break
        
        # Expected Revenue
        exp_revenue = prob * params['adr']
        
        # Delta pickup
        delta = 0
        if date in pickup_data:
            delta = pickup_data[date]['2025'] - pickup_data[date]['2024']
        
        # Soglia upgrade
        threshold = params['adr'] * prob * params['min_margin']
        
        # Decisione
        recommendation = "Sì" if exp_revenue < threshold else "No"
        
        # Suggerimento
        if recommendation == "Sì":
            suggestion = "Bassa probabilità di vendita - considera upgrade gratuito"
        elif prob > 0.7:
            suggestion = "Alta probabilità di vendita - non fare upgrade"
        else:
            suggestion = "Probabilità media - valuta disponibilità e strategia"
        
        dashboard_data.append({
            "Data": date,
            "RN Attuali": rn,
            "Probabilità": prob,
            "Expected Revenue": exp_revenue,
            "Delta Pickup": delta,
            "Soglia Upgrade": threshold,
            "Upgrade Consigliato": recommendation,
            "Suggerimento": suggestion
        })
    
    return dashboard_data

# Layout dell'interfaccia
st.title("🏨 Hotel Upgrade Advisor")
st.write("Strumento avanzato per decisioni di revenue management alberghiero")

# Sidebar per configurazione
st.sidebar.header("Configurazione")

# Parametri
with st.sidebar.expander("Parametri di Base", expanded=True):
    adr = st.number_input("ADR Medio (€)", min_value=50.0, max_value=1000.0, value=381.93, step=5.0)
    min_margin = st.slider("Margine minimo richiesto", min_value=0.1, max_value=0.9, value=0.6, step=0.05, format="%.2f")
    upgrade_threshold = st.number_input("Soglia Upgrade fissa (€)", min_value=0.0, max_value=500.0, value=100.0, step=10.0)
    days = st.slider("Giorni di analisi", min_value=7, max_value=60, value=30, step=1)

# Genera date
start_date_2024 = datetime(2024, 6, 1)
start_date_2025 = datetime(2025, 6, 1)
dates_2024 = [start_date_2024 + timedelta(days=i) for i in range(days)]
dates_2025 = [start_date_2025 + timedelta(days=i) for i in range(days)]

# Opzioni di caricamento dati
st.sidebar.header("Caricamento Dati")
data_option = st.sidebar.radio("Seleziona metodo di caricamento", 
                              ["Genera dati casuali", "Carica file Excel", "Inserisci manualmente"])

# Inizializza dati
historical_data = {}
otb_data = {}
pickup_data = {}

# Opzione 1: Genera dati casuali
if data_option == "Genera dati casuali":
    st.sidebar.subheader("Configurazione dati casuali")
    seed = st.sidebar.number_input("Seed casuale", value=42, min_value=1, max_value=9999)
    occupancy_mean = st.sidebar.slider("Occupazione media prevista (%)", min_value=30, max_value=95, value=75)
    seasonality = st.sidebar.checkbox("Applica stagionalità nel periodo", value=True)
    
    # Genera dati storici casuali
    np.random.seed(seed)
    
    for date in dates_2024:
        day_factor = 1.0
        # Applica stagionalità se selezionata
        if seasonality:
            # Weekend (Ven-Sab) ha occupazione più alta
            if date.weekday() in [4, 5]:  # 4=Ven, 5=Sab
                day_factor = 1.2
            # Metà mese ha picco di occupazione
            if 10 <= date.day <= 20:
                day_factor += 0.1
        
        # Calcola occupazione casuale con distribuzione normale
        base_rooms = np.random.normal(occupancy_mean, 15)
        rooms = max(0, min(100, int(base_rooms * day_factor)))
        historical_data[date] = rooms
    
    # Genera dati OTB e pickup casuali
    for date in dates_2025:
        # OTB: più basso dello storico (ancora da riempire)
        otb_factor = np.random.uniform(0.4, 0.7)
        hist_date = date.replace(year=2024)
        if hist_date in historical_data:
            otb_data[date] = int(historical_data[hist_date] * otb_factor)
        else:
            otb_data[date] = int(70 * otb_factor)
        
        # Pickup: confronto tra anni per 7 giorni
        p2025 = np.random.randint(5, 15)
        trend_factor = np.random.uniform(0.8, 1.2)  # Variazione YoY
        p2024 = int(p2025 / trend_factor)
        
        pickup_data[date] = {
            '2025': p2025,
            '2024': p2024
        }

# Opzione 2: Carica file Excel
elif data_option == "Carica file Excel":
    uploaded_file = st.sidebar.file_uploader("Carica file Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Carica dati da Excel
        try:
            xls = pd.ExcelFile(uploaded_file)
            
            # Controlla fogli necessari
            required_sheets = ["Storico 2024", "OTB Giugno 2025", "Pickup vs SPIT"]
            missing_sheets = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
            
            if missing_sheets:
                st.sidebar.error(f"Fogli mancanti: {', '.join(missing_sheets)}")
            else:
                # Carica dati storici
                df_hist = pd.read_excel(xls, "Storico 2024", skiprows=1)
                for _, row in df_hist.iterrows():
                    if isinstance(row["Data"], datetime):
                        historical_data[row["Data"]] = row["Room Nights 2024"]
                
                # Carica dati OTB
                df_otb = pd.read_excel(xls, "OTB Giugno 2025", skiprows=1)
                for _, row in df_otb.iterrows():
                    if isinstance(row["Data"], datetime):
                        otb_data[row["Data"]] = row["Room Nights"]
                
                # Carica dati Pickup
                df_pickup = pd.read_excel(xls, "Pickup vs SPIT", skiprows=1)
                for _, row in df_pickup.iterrows():
                    if isinstance(row["Data"], datetime):
                        pickup_data[row["Data"]] = {
                            '2025': row["Pickup 7gg 2025"], 
                            '2024': row["Pickup 7gg 2024"]
                        }
                
                st.sidebar.success("Dati caricati con successo!")
        
        except Exception as e:
            st.sidebar.error(f"Errore nel caricamento del file: {e}")

# Opzione 3: Inserimento manuale
elif data_option == "Inserisci manualmente":
    with st.sidebar.expander("Guida all'inserimento", expanded=False):
        st.write("""
        Utilizza le schede qui sotto per inserire manualmente i dati. 
        Per ogni giorno, potrai specificare:
        - Storico 2024: Room nights vendute l'anno scorso
        - OTB 2025: Room nights già in casa per quest'anno
        - Pickup: Prenotazioni ricevute negli ultimi 7 giorni
        """)
    
    # Crea tre schede per inserimento dati
    tab1, tab2, tab3 = st.tabs(["Storico 2024", "OTB 2025", "Pickup"])
    
    # Tab 1: Dati storici
    with tab1:
        st.subheader("Dati storici Giugno 2024")
        col1, col2 = st.columns(2)
        
        # Prima metà del mese
        with col1:
            st.write("Prima metà del mese")
            for i in range(15):
                date = dates_2024[i]
                key = f"hist_{date.strftime('%Y%m%d')}"
                room_nights = st.number_input(
                    f"{date.strftime('%d/%m/%Y')}",
                    min_value=0,
                    max_value=1000,
                    value=70,
                    key=key
                )
                historical_data[date] = room_nights
        
        # Seconda metà del mese
        with col2:
            st.write("Seconda metà del mese")
            for i in range(15, min(30, len(dates_2024))):
                date = dates_2024[i]
                key = f"hist_{date.strftime('%Y%m%d')}"
                room_nights = st.number_input(
                    f"{date.strftime('%d/%m/%Y')}",
                    min_value=0,
                    max_value=1000,
                    value=70,
                    key=key
                )
                historical_data[date] = room_nights
    
    # Tab 2: Dati OTB
    with tab2:
        st.subheader("OTB Giugno 2025")
        col1, col2 = st.columns(2)
        
        # Prima metà del mese
        with col1:
            st.write("Prima metà del mese")
            for i in range(15):
                date = dates_2025[i]
                key = f"otb_{date.strftime('%Y%m%d')}"
                room_nights = st.number_input(
                    f"{date.strftime('%d/%m/%Y')}",
                    min_value=0,
                    max_value=1000,
                    value=40,
                    key=key
                )
                otb_data[date] = room_nights
        
        # Seconda metà del mese
        with col2:
            st.write("Seconda metà del mese")
            for i in range(15, min(30, len(dates_2025))):
                date = dates_2025[i]
                key = f"otb_{date.strftime('%Y%m%d')}"
                room_nights = st.number_input(
                    f"{date.strftime('%d/%m/%Y')}",
                    min_value=0,
                    max_value=1000,
                    value=40,
                    key=key
                )
                otb_data[date] = room_nights
    
    # Tab 3: Dati Pickup
    with tab3:
        st.subheader("Pickup a 7 giorni")
        
        for i in range(min(10, len(dates_2025))):
            date = dates_2025[i]
            col1, col2 = st.columns(2)
            
            with col1:
                key_2025 = f"pickup_2025_{date.strftime('%Y%m%d')}"
                pickup_2025 = st.number_input(
                    f"{date.strftime('%d/%m/%Y')} - 2025",
                    min_value=0,
                    max_value=100,
                    value=8,
                    key=key_2025
                )
            
            with col2:
                key_2024 = f"pickup_2024_{date.strftime('%Y%m%d')}"
                pickup_2024 = st.number_input(
                    f"{date.strftime('%d/%m/%Y')} - 2024",
                    min_value=0,
                    max_value=100,
                    value=7,
                    key=key_2024
                )
            
            pickup_data[date] = {
                '2025': pickup_2025,
                '2024': pickup_2024
            }
        
        if st.checkbox("Mostra più date"):
            for i in range(10, min(30, len(dates_2025))):
                date = dates_2025[i]
                col1, col2 = st.columns(2)
                
                with col1:
                    key_2025 = f"pickup_2025_{date.strftime('%Y%m%d')}"
                    pickup_2025 = st.number_input(
                        f"{date.strftime('%d/%m/%Y')} - 2025",
                        min_value=0,
                        max_value=100,
                        value=8,
                        key=key_2025
                    )
                
                with col2:
                    key_2024 = f"pickup_2024_{date.strftime('%Y%m%d')}"
                    pickup_2024 = st.number_input(
                        f"{date.strftime('%d/%m/%Y')} - 2024",
                        min_value=0,
                        max_value=100,
                        value=7,
                        key=key_2024
                    )
                
                pickup_data[date] = {
                    '2025': pickup_2025,
                    '2024': pickup_2024
                }

# Assicurarsi che ci siano dati da mostrare
if not historical_data or not otb_data or not pickup_data:
    st.warning("Seleziona un'opzione di caricamento dati nella barra laterale")
else:
    # Prepara i parametri
    params = {
        'adr': adr,
        'min_margin': min_margin,
        'upgrade_threshold': upgrade_threshold,
        'days': days
    }
    
    # Calcola i dati della dashboard
    dashboard_data = calculate_dashboard_data(historical_data, otb_data, pickup_data, params)
    
    # Visualizza i risultati
    tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "Grafici", "Dati di Input", "Previsioni"])
    
    # Tab 1: Dashboard
    with tab1:
        st.subheader("Dashboard Decisionale Upgrade")
        
        # Metriche principali
        col1, col2, col3, col4 = st.columns(4)
        
        # Calcola metriche chiave
        total_rooms = sum(otb.get("RN Attuali", 0) for otb in dashboard_data)
        avg_probability = sum(d.get("Probabilità", 0) for d in dashboard_data) / len(dashboard_data) if dashboard_data else 0
        recommended_upgrades = sum(1 for d in dashboard_data if d.get("Upgrade Consigliato") == "Sì")
        total_revenue = sum(d.get("Expected Revenue", 0) for d in dashboard_data)
        
        with col1:
            st.metric(label="Room Nights Totali", value=f"{total_rooms}")
        with col2:
            st.metric(label="Probabilità Media", value=f"{avg_probability:.1%}")
        with col3:
            st.metric(label="Upgrade Consigliati", value=f"{recommended_upgrades}")
        with col4:
            st.metric(label="Revenue Stimato", value=f"€{total_revenue:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        # Tabella con i risultati della dashboard
        df_dashboard = pd.DataFrame(dashboard_data)
        
        # Formatta colonne
        df_dashboard["Data"] = df_dashboard["Data"].dt.strftime("%d/%m/%Y")
        df_dashboard["Probabilità"] = df_dashboard["Probabilità"].map("{:.1%}".format)
        df_dashboard["Expected Revenue"] = df_dashboard["Expected Revenue"].map("€{:,.2f}".format).str.replace(".", "X").str.replace(",", ".").str.replace("X", ",")
        df_dashboard["Soglia Upgrade"] = df_dashboard["Soglia Upgrade"].map("€{:,.2f}".format).str.replace(".", "X").str.replace(",", ".").str.replace("X", ",")
        
        # Crea dataframe formattato
        st.dataframe(df_dashboard, use_container_width=True)
        
        # Download del file Excel
        if st.button("Genera File Excel"):
            excel_buffer = create_excel_file(params, historical_data, otb_data, pickup_data)
            st.download_button(
                label="📥 Scarica Excel",
                data=excel_buffer,
                file_name="hotel_upgrade_advisor.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # Tab 2: Grafici
    with tab2:
        st.subheader("Grafici di Analisi")
        
        # Converti dati per grafici
        df_dash_raw = pd.DataFrame(dashboard_data)
        
        # Grafico 1: Probabilità di Vendita vs RN Attuali
        st.subheader("Probabilità di Vendita vs Room Nights")
        fig1 = px.scatter(
            df_dash_raw, 
            x="Data", 
            y="Probabilità", 
            size="RN Attuali",
            color="Upgrade Consigliato",
            color_discrete_map={"Sì": "green", "No": "red"},
            hover_data=["Suggerimento"]
        )
        fig1.update_layout(yaxis_tickformat=".0%")
        st.plotly_chart(fig1, use_container_width=True)
        
        # Grafico 2: Expected Revenue vs Soglia Upgrade
        st.subheader("Expected Revenue vs Soglia Upgrade")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=df_dash_raw["Data"],
            y=df_dash_raw["Expected Revenue"],
            name="Expected Revenue",
            marker_color='royalblue'
        ))
        fig2.add_trace(go.Bar(
            x=df_dash_raw["Data"],
            y=df_dash_raw["Soglia Upgrade"],
            name="Soglia Upgrade",
            marker_color='firebrick'
        ))
        fig2.update_layout(barmode='group', xaxis_tickangle=-45)
        st.plotly_chart(fig2, use_container_width=True)
        
        # Grafico 3: Delta Pickup 2025 vs 2024
        st.subheader("Delta Pickup: 2025 vs 2024")
        
        pickup_colors = ['green' if delta > 0 else 'red' for delta in df_dash_raw["Delta Pickup"]]
        
        fig3 = px.bar(
            df_dash_raw,
            x="Data",
            y="Delta Pickup",
            color_discrete_sequence=["#5D69B1"]
        )
        fig3.update_traces(marker_color=pickup_colors)
        st.plotly_chart(fig3, use_container_width=True)
    
    # Tab 3: Dati di Input
    with tab3:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Storico 2024")
            df_hist = pd.DataFrame({
                "Data": list(historical_data.keys()),
                "Room Nights": list(historical_data.values())
            })
            df_hist["Data"] = df_hist["Data"].dt.strftime("%d/%m/%Y")
            st.dataframe(df_hist, use_container_width=True)
            
            st.subheader("Analisi Statistica Storico")
            hist_values = list(historical_data.values())
            hist_stats = {
                "Min": min(hist_values),
                "Max": max(hist_values),
                "Media": sum(hist_values) / len(hist_values),
                "Mediana": sorted(hist_values)[len(hist_values) // 2]
            }
            st.write(hist_stats)
            
            # Istogramma distribuzione storica
            fig_hist = px.histogram(df_hist, x="Room Nights", nbins=10, title="Distribuzione Room Nights 2024")
            st.plotly_chart(fig_hist, use_container_width=True)
        
        with col2:
            st.subheader("OTB 2025")
            df_otb = pd.DataFrame({
                "Data": list(otb_data.keys()),
                "Room Nights": list(otb_data.values())
            })
            df_otb["Data"] = df_otb["Data"].dt.strftime("%d/%m/%Y")
            st.dataframe(df_otb, use_container_width=True)
            
            st.subheader("Pickup vs SPIT")
            df_pickup = pd.DataFrame([
                {
                    "Data": date.strftime("%d/%m/%Y"),
                    "Pickup 2025": data['2025'],
                    "Pickup 2024": data['2024'],
                    "Delta": data['2025'] - data['2024']
                }
                for date, data in pickup_data.items()
            ])
            st.dataframe(df_pickup, use_container_width=True)
    
    # Tab 4: Previsioni
    with tab4:
        st.subheader("Previsioni e Suggerimenti")
        
        # Analisi aggregata degli upgrade
        upgrade_counts = {"Sì": 0, "No": 0}
        for d in dashboard_data:
            upgrade_counts[d["Upgrade Consigliato"]] += 1
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Distribuzione Decisioni Upgrade")
            fig_pie = px.pie(values=list(upgrade_counts.values()), names=list(upgrade_counts.keys()),
                           color_discrete_map={"Sì": "#4CAF50", "No": "#F44336"})
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            st.subheader("Suggerimenti Operativi")
            
            # Raggruppa suggerimenti
            suggestions = {}
            for d in dashboard_data:
                sugg = d["Suggerimento"]
                if sugg in suggestions:
                    suggestions[sugg] += 1
                else:
                    suggestions[sugg] = 1
            
            # Crea dataframe per visualizzazione
            df_sugg = pd.DataFrame({
                "Suggerimento": list(suggestions.keys()),
                "Conteggio": list(suggestions.values())
            })
            
            fig_sugg = px.bar(df_sugg, x="Suggerimento", y="Conteggio", color="Suggerimento",
                             color_discrete_map={
                                 "Bassa probabilità di vendita - considera upgrade gratuito": "#4CAF50",
                                 "Alta probabilità di vendita - non fare upgrade": "#F44336",
                                 "Probabilità media - valuta disponibilità e strategia": "#FFC107"
                             })
            st.plotly_chart(fig_sugg, use_container_width=True)
        
        # Riepilogo e analisi complessiva
        st.subheader("Analisi Complessiva")
        
        # Identifica giorni critici (bassa probabilità)
        low_prob_days = [d for d in dashboard_data if d["Probabilità"] < 0.4]
        
        if low_prob_days:
            st.warning("Giorni con Bassa Probabilità di Vendita (< 40%)")
            low_prob_df = pd.DataFrame(low_prob_days)
            low_prob_df["Data"] = pd.to_datetime(low_prob_df["Data"]).dt.strftime("%d/%m/%Y")
            low_prob_df["Probabilità"] = low_prob_df["Probabilità"].map("{:.1%}".format)
            low_prob_df = low_prob_df[["Data", "RN Attuali", "Probabilità", "Suggerimento"]]
            st.dataframe(low_prob_df, use_container_width=True)
        
        # Identifica giorni di alta opportunità
        high_opp_days = [d for d in dashboard_data if d["Probabilità"] > 0.7 and d["RN Attuali"] < 60]
        
        if high_opp_days:
            st.success("Giorni con Alta Probabilità di Vendita ma Bassa Occupazione (Opportunità)")
            high_opp_df = pd.DataFrame(high_opp_days)
            high_opp_df["Data"] = pd.to_datetime(high_opp_df["Data"]).dt.strftime("%d/%m/%Y")
            high_opp_df["Probabilità"] = high_opp_df["Probabilità"].map("{:.1%}".format)
            high_opp_df = high_opp_df[["Data", "RN Attuali", "Probabilità", "Suggerimento"]]
            st.dataframe(high_opp_df, use_container_width=True)
        
        # Riepilogo generale
        st.info("""
        ### Riepilogo Strategico
        
        Questo strumento analizza i dati storici e le prenotazioni attuali per supportare decisioni di revenue management.
        
        **Interpretazione:**
        - Quando l'**Expected Revenue** è inferiore alla **Soglia Upgrade**, conviene offrire upgrade gratuiti per liberare camere di categoria inferiore che hanno maggiore probabilità di essere vendute.
        - Un **Delta Pickup** positivo indica una performance migliore rispetto all'anno precedente.
        - Utilizza i **Suggerimenti Operativi** per guidare le decisioni giornaliere.
        
        Per analisi più dettagliate, scarica il file Excel dalla Dashboard.
        """)

# Footer con informazioni
st.sidebar.markdown("---")
st.sidebar.info("""
### Hotel Upgrade Advisor v1.0
Sviluppato per ottimizzare le decisioni di revenue management alberghiero.

Per assistenza o personalizzazioni:
- Email: support@hotelupgradeadvisor.com
- GitHub: [github.com/hotelupgradeadvisor](https://github.com)
""")

# Istruzioni per deploy
with st.sidebar.expander("Istruzioni Deploy", expanded=False):
    st.markdown("""
    ### Per deployare su Streamlit Community Cloud:
    
    1. Carica questo codice su GitHub in un repository pubblico
    2. Aggiungi un file `requirements.txt` con le dipendenze:
       ```
       streamlit
       pandas
       numpy
       plotly
       openpyxl
       ```
    3. Visita [https://streamlit.io/cloud](https://streamlit.io/cloud)
    4. Accedi con il tuo account GitHub
    5. Seleziona il repository e clicca "Deploy"
    """)
