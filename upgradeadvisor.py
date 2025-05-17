import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
import networkx as nx
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta, date
import os
import re
import sqlite3
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# Configurazione pagina
st.set_page_config(
    page_title="Hotel Upgrade Advisor Pro",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Definizione CSS personalizzato
st.markdown("""
<style>
.section-title {
    color: #4472C4;
    padding-bottom: 10px;
    border-bottom: 1px solid #e0e0e0;
    margin-bottom: 20px;
}
.processing-step {
    padding: 10px;
    border-radius: 5px;
    margin-bottom: 15px;
}
.progress {
    background-color: #f0f7ff;
    border-left: 4px solid #4472C4;
}
.completed {
    background-color: #e6f7e6;
    border-left: 4px solid #5cb85c;
}
.error {
    background-color: #fff0f0;
    border-left: 4px solid #d9534f;
}
.data-warning {
    background-color: #fff9e6;
    padding: 10px;
    border-radius: 5px;
    margin-bottom: 15px;
    border-left: 4px solid #f0ad4e;
}
.data-info {
    background-color: #e6f7ff;
    padding: 10px;
    border-radius: 5px;
    margin-bottom: 15px;
    border-left: 4px solid #5bc0de;
}
.insights-card {
    padding: 15px;
    margin-bottom: 15px;
    border-radius: 5px;
    border-left: 4px solid #5cb85c;
}
.positivo {
    background-color: #e6f7e6;
    border-left: 4px solid #5cb85c;
}
.negativo {
    background-color: #fff0f0;
    border-left: 4px solid #d9534f;
}
.neutro {
    background-color: #f7f7f7;
    border-left: 4px solid #777777;
}
.attenzione {
    background-color: #fff9e6;
    border-left: 4px solid #f0ad4e;
}
</style>
""", unsafe_allow_html=True)

# Funzioni di utilità
def get_data_path(filename):
    """Restituisce il percorso corretto per il file dati"""
    if os.path.exists(filename):
        return filename
    elif os.path.exists(os.path.join("data", filename)):
        return os.path.join("data", filename)
    else:
        os.makedirs("data", exist_ok=True)
        return os.path.join("data", filename)

def get_connection():
    """Crea una connessione al database SQLite"""
    db_path = get_data_path("hotel_data.db")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def ensure_date(date_object):
    """Converte l'oggetto in date se è datetime, se è già date lo lascia invariato"""
    if hasattr(date_object, 'date'):
        return date_object.date()
    return date_object

def init_database():
    """Inizializza il database se non esiste"""
    conn = get_connection()
    cursor = conn.cursor()
    
    # Tabella per dati di pickup
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS pickup_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hotel TEXT NOT NULL,
        stay_date TEXT NOT NULL,
        roomnights INTEGER,
        vs_1day INTEGER,
        vs_2day INTEGER,
        vs_3day INTEGER,
        vs_7day INTEGER,
        vs_7day_spit INTEGER,
        insertion_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # Tabella per dati di produzione
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS production_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hotel TEXT NOT NULL,
        date TEXT NOT NULL,
        weekday TEXT,
        room_nights INTEGER,
        occupancy REAL,
        adr REAL,
        room_revenue REAL,
        vs_spit_rn INTEGER,
        vs_spit_occ REAL,
        vs_spit_adr REAL,
        vs_ap_rn INTEGER,
        vs_ap_occ REAL,
        vs_ap_adr REAL,
        insertion_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # Tabella per configurazione hotel
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS hotels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hotel_name TEXT UNIQUE NOT NULL,
        total_rooms INTEGER DEFAULT 0,
        address TEXT,
        city TEXT,
        stars INTEGER DEFAULT 4,
        seasonal BOOLEAN DEFAULT 0,
        creation_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # Tabella per tipologie camere
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS room_types (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hotel_id INTEGER NOT NULL,
        room_type_name TEXT NOT NULL,
        rooms_in_type INTEGER DEFAULT 0,
        is_entry_level BOOLEAN DEFAULT 0,
        adr REAL DEFAULT 0,
        min_margin REAL DEFAULT 0.6,
        upgrade_threshold REAL DEFAULT 100,
        upgrade_target_type TEXT,
        creation_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (hotel_id) REFERENCES hotels (id),
        UNIQUE (hotel_id, room_type_name)
    )
    ''')
    
    # Tabella per camere Out of Order
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS rooms_ooo (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hotel_id INTEGER NOT NULL,
        room_type_id INTEGER NOT NULL,
        date_from TEXT NOT NULL,
        date_to TEXT NOT NULL,
        rooms_count INTEGER DEFAULT 1,
        reason TEXT,
        creation_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (hotel_id) REFERENCES hotels (id),
        FOREIGN KEY (room_type_id) REFERENCES room_types (id)
    )
    ''')
    
    # Tabella per parametri hotel - da mantenere per compatibilità
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS hotel_settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hotel TEXT UNIQUE NOT NULL,
        adr REAL DEFAULT 381.93,
        min_margin REAL DEFAULT 0.6,
        upgrade_threshold REAL DEFAULT 100,
        days INTEGER DEFAULT 30,
        is_entry_level BOOLEAN DEFAULT 0,
        room_type_name TEXT,
        rooms_in_type INTEGER DEFAULT 0,
        total_rooms INTEGER DEFAULT 0,
        insertion_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        last_update TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    conn.commit()
    conn.close()

# Funzioni per la gestione degli hotel
def get_all_hotels():
    """Recupera tutti gli hotel dal database"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("SELECT * FROM hotels ORDER BY hotel_name")
        results = cursor.fetchall()
        return [dict(r) for r in results]
    except Exception as e:
        print(f"Errore nel recupero degli hotel: {e}")
        return []
    finally:
        conn.close()

def get_hotel_by_name(hotel_name):
    """Recupera un hotel dal database tramite nome"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("SELECT * FROM hotels WHERE hotel_name = ?", (hotel_name,))
        result = cursor.fetchone()
        return dict(result) if result else None
    except Exception as e:
        print(f"Errore nel recupero dell'hotel: {e}")
        return None
    finally:
        conn.close()

def add_hotel(hotel_data):
    """Aggiunge un nuovo hotel al database"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
        INSERT INTO hotels 
        (hotel_name, total_rooms, address, city, stars, seasonal) 
        VALUES (?, ?, ?, ?, ?, ?)
        """, (
            hotel_data.get('hotel_name', ''),
            hotel_data.get('total_rooms', 0),
            hotel_data.get('address', ''),
            hotel_data.get('city', ''),
            hotel_data.get('stars', 4),
            hotel_data.get('seasonal', False)
        ))
        conn.commit()
        
        # Restituisci l'ID dell'hotel appena inserito
        cursor.execute("SELECT last_insert_rowid()")
        result = cursor.fetchone()
        return result[0] if result else None
    except Exception as e:
        conn.rollback()
        print(f"Errore nell'inserimento dell'hotel: {e}")
        return None
    finally:
        conn.close()

def update_hotel(hotel_id, hotel_data):
    """Aggiorna i dati di un hotel esistente"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
        UPDATE hotels SET
        hotel_name = ?,
        total_rooms = ?,
        address = ?,
        city = ?,
        stars = ?,
        seasonal = ?
        WHERE id = ?
        """, (
            hotel_data.get('hotel_name'),
            hotel_data.get('total_rooms'),
            hotel_data.get('address'),
            hotel_data.get('city'),
            hotel_data.get('stars'),
            hotel_data.get('seasonal'),
            hotel_id
        ))
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Errore nell'aggiornamento dell'hotel: {e}")
        return False
    finally:
        conn.close()

def delete_hotel(hotel_id):
    """Elimina un hotel dal database"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        # Elimina prima tutte le tipologie associate
        cursor.execute("DELETE FROM room_types WHERE hotel_id = ?", (hotel_id,))
        # Elimina l'hotel
        cursor.execute("DELETE FROM hotels WHERE id = ?", (hotel_id,))
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Errore nell'eliminazione dell'hotel: {e}")
        return False
    finally:
        conn.close()

# Funzioni per la gestione delle tipologie di camera
def get_room_types_by_hotel(hotel_id):
    """Recupera tutte le tipologie di camera di un hotel"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
        SELECT * FROM room_types 
        WHERE hotel_id = ?
        ORDER BY adr
        """, (hotel_id,))
        results = cursor.fetchall()
        return [dict(r) for r in results]
    except Exception as e:
        print(f"Errore nel recupero delle tipologie: {e}")
        return []
    finally:
        conn.close()

def get_room_type_by_name(hotel_id, room_type_name):
    """Recupera una tipologia di camera tramite nome"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
        SELECT * FROM room_types 
        WHERE hotel_id = ? AND room_type_name = ?
        """, (hotel_id, room_type_name))
        result = cursor.fetchone()
        return dict(result) if result else None
    except Exception as e:
        print(f"Errore nel recupero della tipologia: {e}")
        return None
    finally:
        conn.close()

def add_room_type(hotel_id, room_type_data):
    """Aggiunge una nuova tipologia di camera"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
        INSERT INTO room_types 
        (hotel_id, room_type_name, rooms_in_type, is_entry_level, adr, min_margin, 
         upgrade_threshold, upgrade_target_type)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            hotel_id,
            room_type_data.get('room_type_name', ''),
            room_type_data.get('rooms_in_type', 0),
            room_type_data.get('is_entry_level', False),
            room_type_data.get('adr', 0),
            room_type_data.get('min_margin', 0.6),
            room_type_data.get('upgrade_threshold', 100),
            room_type_data.get('upgrade_target_type')
        ))
        conn.commit()
        
        # Restituisci l'ID della tipologia appena inserita
        cursor.execute("SELECT last_insert_rowid()")
        result = cursor.fetchone()
        return result[0] if result else None
    except Exception as e:
        conn.rollback()
        print(f"Errore nell'inserimento della tipologia: {e}")
        return None
    finally:
        conn.close()

def update_room_type(room_type_id, room_type_data):
    """Aggiorna i dati di una tipologia esistente"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
        UPDATE room_types SET
        room_type_name = ?,
        rooms_in_type = ?,
        is_entry_level = ?,
        adr = ?,
        min_margin = ?,
        upgrade_threshold = ?,
        upgrade_target_type = ?
        WHERE id = ?
        """, (
            room_type_data.get('room_type_name'),
            room_type_data.get('rooms_in_type'),
            room_type_data.get('is_entry_level'),
            room_type_data.get('adr'),
            room_type_data.get('min_margin'),
            room_type_data.get('upgrade_threshold'),
            room_type_data.get('upgrade_target_type'),
            room_type_id
        ))
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Errore nell'aggiornamento della tipologia: {e}")
        return False
    finally:
        conn.close()

# Funzioni per la gestione delle camere Out of Order (OOO)
def get_active_ooo_rooms(hotel_id):
    """Recupera le camere fuori servizio attive"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
        SELECT o.*, r.room_type_name 
        FROM rooms_ooo o
        JOIN room_types r ON o.room_type_id = r.id
        WHERE o.hotel_id = ? AND date_to >= date('now')
        ORDER BY date_from
        """, (hotel_id,))
        results = cursor.fetchall()
        return [dict(r) for r in results]
    except Exception as e:
        print(f"Errore nel recupero delle camere OOO: {e}")
        return []
    finally:
        conn.close()

def add_ooo_rooms(hotel_id, room_type_id, ooo_data):
    """Aggiunge camere fuori servizio"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
        INSERT INTO rooms_ooo
        (hotel_id, room_type_id, date_from, date_to, rooms_count, reason)
        VALUES (?, ?, ?, ?, ?, ?)
        """, (
            hotel_id,
            room_type_id,
            ooo_data.get('date_from'),
            ooo_data.get('date_to'),
            ooo_data.get('rooms_count', 1),
            ooo_data.get('reason', '')
        ))
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Errore nell'inserimento delle camere OOO: {e}")
        return False
    finally:
        conn.close()

def end_ooo(ooo_id):
    """Termina un periodo di OOO anticipatamente"""
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        # Imposta la data di fine a ieri
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')
        cursor.execute("""
        UPDATE rooms_ooo 
        SET date_to = ?
        WHERE id = ? AND date_to > date('now')
        """, (yesterday, ooo_id))
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Errore nella terminazione dell'OOO: {e}")
        return False
    finally:
        conn.close()

def get_effective_capacity(hotel_id, room_type_id, date):
    """
    Calcola la capacità effettiva di una tipologia per una data specifica,
    considerando le camere fuori servizio.
    """
    try:
        conn = get_connection()
        cursor = conn.cursor()
        
        # Ottieni capacità base
        cursor.execute("""
        SELECT rooms_in_type FROM room_types 
        WHERE id = ? AND hotel_id = ?
        """, (room_type_id, hotel_id))
        
        result = cursor.fetchone()
        if not result:
            return 0
        
        base_capacity = result['rooms_in_type']
        
        # Trova camere OOO per questa data
        date_str = date.strftime('%d/%m/%Y') if isinstance(date, datetime) else date
        
        cursor.execute("""
        SELECT SUM(rooms_count) as ooo_rooms FROM rooms_ooo
        WHERE hotel_id = ? AND room_type_id = ?
        AND date_from <= ? AND date_to >= ?
        """, (hotel_id, room_type_id, date_str, date_str))
        
        ooo_result = cursor.fetchone()
        ooo_count = ooo_result['ooo_rooms'] if ooo_result and ooo_result['ooo_rooms'] else 0
        
        # Capacità effettiva
        effective_capacity = max(0, base_capacity - ooo_count)
        
        return effective_capacity
    
    except Exception as e:
        print(f"Errore nel calcolo della capacità effettiva: {e}")
        return 0
    finally:
        conn.close()

# Funzioni per l'analisi dei dati
def analyze_vs_spit_trend(hotel_name, days=60):
    """
    Analizza il trend vs SPIT nelle ultime settimane per un hotel.
    
    Args:
        hotel_name: Nome dell'hotel
        days: Numero di giorni da analizzare
    
    Returns:
        Dizionario con statistiche sul trend vs SPIT
    """
    daily_data = get_latest_excel_data(hotel=hotel_name, section="Produzione Giornaliera")
    if not daily_data or "daily_data" not in daily_data:
        return {"error": "Dati non disponibili"}
    
    df = pd.DataFrame(daily_data["daily_data"])
    df["date_obj"] = pd.to_datetime(df["date"], format="%d/%m/%Y", errors="coerce")
    
    if "vs_spit_rn" not in df.columns or "room_nights" not in df.columns:
        return {"error": "Dati vs SPIT non disponibili"}
    
    # Converti a numerico
    df["vs_spit_rn"] = pd.to_numeric(df["vs_spit_rn"], errors="coerce")
    df["room_nights"] = pd.to_numeric(df["room_nights"], errors="coerce")
    
    today = datetime.now().date()
    cutoff_date = today - timedelta(days=days)
    
    recent_data = df[df["date_obj"].dt.date >= cutoff_date]
    
    # Totali per il periodo
    total_rn = recent_data["room_nights"].sum()
    total_vs_spit = recent_data["vs_spit_rn"].sum()
    
    # Calcola i roomnights dell'anno scorso
    spit_rn = total_rn - total_vs_spit
    
    # Calcola variazione percentuale
    vs_spit_pct = (total_vs_spit / spit_rn * 100) if spit_rn > 0 else 0
    
    return {
        "total_rn": total_rn,
        "total_vs_spit": total_vs_spit,
        "spit_rn": spit_rn,
        "vs_spit_pct": vs_spit_pct,
        "period_days": days
    }

# Funzioni per l'estrazione e l'elaborazione dei dati PowerBI
def excel_column_to_index(column_letter):
    """Converte lettera colonna Excel in indice numerico (base 0)."""
    result = 0
    for char in column_letter:
        result = result * 26 + (ord(char.upper()) - ord('A') + 1)
    return result - 1

def detect_report_type(df):
    """
    Identifica automaticamente il tipo di report basandosi sulla struttura delle colonne.
    Restituisce il tipo di report e un livello di confidenza.
    """
    # Converti i nomi delle colonne in lowercase per confronto case-insensitive
    cols_lower = [str(col).lower() for col in df.columns]
    
    # Check if it's a pickup report
    pickup_cols = ["vs 1gg", "vs 2gg", "vs 3gg", "vs 7gg", "vs 7gg SPIT"]
    pickup_cols_found = sum(1 for col in pickup_cols if any(col.lower() in str(c).lower() for c in df.columns))
    if ("soggiorno" in cols_lower or "data" in cols_lower) and pickup_cols_found >= 2:
        return "Pickup", pickup_cols_found
    
    # Schema colonne produzione portafoglio mensile
    if any("mese" in col for col in cols_lower) and any("occ.%" in col for col in cols_lower or "occupazione" in col for col in cols_lower):
        confidence = sum(1 for pattern in ["room revenue", "adr", "revpar", "mese"] 
                         if any(pattern in col for col in cols_lower))
        if confidence >= 3:
            return "Produzione Portafoglio", confidence
    
    # Schema colonne produzione giornaliera
    if any("data" in col for col in cols_lower) and any("giorno" in col for col in cols_lower):
        confidence = sum(1 for pattern in ["% occ", "room nights", "adr", "data"] 
                         if any(pattern in col for col in cols_lower))
        if confidence >= 3:
            return "Produzione Giornaliera", confidence
    
    # Schema colonne produzione segmenti
    if any("segment" in col for col in cols_lower) and any(("perc" in col or "%" in col) for col in cols_lower):
        return "Prod. per Segmento", 3
    
    # Non è stato possibile identificare il tipo
    return None, 0

def detect_report_metadata(df):
    """
    Rileva automaticamente metadati dal report come tipologie di camera,
    periodo temporale e altre informazioni rilevanti.
    """
    metadata = {
        "room_type": None,
        "date_period": None,
        "hotel": None
    }
    
    # Rileva le camere/tipologie incluse nel report
    for col in df.columns:
        col_str = str(col).lower()
        # Cerca riferimenti a tipologie di camera nelle intestazioni
        if "tipologia" in col_str or "camera" in col_str or "tipo camera" in col_str:
            for cell in df[col].dropna():
                cell_str = str(cell).lower()
                # Cerca riferimenti a tipologie del Cala Cuncheddi
                for room_type in ["classic garden", "classic sea view", "superior sea view", 
                                  "executive", "deluxe", "family", "junior suite", "suite"]:
                    if room_type in cell_str:
                        metadata["room_type"] = room_type.title()
                        break
                if metadata["room_type"]:
                    break
        if metadata["room_type"]:
            break
    
    # Rileva periodo e mese/anno
    for col in df.columns:
        col_str = str(col).lower()
        if "data" in col_str or "periodo" in col_str or "mese" in col_str:
            # Cerca nelle celle di questa colonna per trovare riferimenti al periodo
            for cell in df[col].dropna():
                cell_str = str(cell).lower()
                # Cerca mese e anno
                months = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", 
                         "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"]
                for month in months:
                    if month in cell_str:
                        # Cerca l'anno
                        year_match = re.search(r'20\d{2}', cell_str)
                        if year_match:
                            metadata["date_period"] = f"{month.title()} {year_match.group(0)}"
                            break
                if metadata["date_period"]:
                    break
        if metadata["date_period"]:
            break
    
    # Rileva nome hotel (se presente)
    hotel_keywords = ["cala cuncheddi", "hotel", "struttura", "albergo"]
    for col in df.columns:
        for cell in df[col].dropna():
            cell_str = str(cell).lower()
            for keyword in hotel_keywords:
                if keyword in cell_str:
                    # Cerca di estrarre nome completo hotel
                    if "cala cuncheddi" in cell_str:
                        metadata["hotel"] = "Cala Cuncheddi"
                        break
            if metadata["hotel"]:
                break
        if metadata["hotel"]:
            break
            
    return metadata

def process_pickup_excel(uploaded_file, hotel):
    """
    Elabora il file Excel di pickup con gestione migliorata delle date italiane.
    """
    try:
        df = pd.read_excel(uploaded_file)
        
        # Rileva metadati dal report
        metadata = detect_report_metadata(df)
        
        # Use the exact column names from your Excel file
        expected_columns = ['Soggiorno', 'Roomnights', 'vs 1gg', 'vs 2gg', 'vs 3gg', 'vs 7gg', 'vs 7gg SPIT']
        
        # Verify all expected columns exist
        missing_columns = [col for col in expected_columns if not any(col.lower() in str(c).lower() for c in df.columns)]
        if missing_columns:
            return None, f"Colonne mancanti: {', '.join(missing_columns)}. Verificare il formato."
        
        # Identify the actual column names in the DataFrame
        soggiorno_col = next((col for col in df.columns if 'soggiorno' in str(col).lower()), None)
        roomnights_col = next((col for col in df.columns if 'roomnights' in str(col).lower()), None)
        vs_1gg_col = next((col for col in df.columns if 'vs 1gg' in str(col).lower()), None)
        vs_2gg_col = next((col for col in df.columns if 'vs 2gg' in str(col).lower()), None)
        vs_3gg_col = next((col for col in df.columns if 'vs 3gg' in str(col).lower()), None)
        vs_7gg_col = next((col for col in df.columns if 'vs 7gg' in str(col).lower() and 'spit' not in str(col).lower()), None)
        vs_7gg_spit_col = next((col for col in df.columns if 'vs 7gg spit' in str(col).lower()), None)
        
        if not all([soggiorno_col, roomnights_col, vs_7gg_col, vs_7gg_spit_col]):
            return None, "Alcune colonne essenziali non sono state trovate. Verificare formato file."
        
        pickup_data = []
        
        for idx, row in df.iterrows():
            # Skip header rows or empty rows
            if pd.isna(row[soggiorno_col]) or (isinstance(row[soggiorno_col], str) and "soggiorno" in row[soggiorno_col].lower()):
                continue
                
            # Format date properly - con gestione migliorata per date italiane
            stay_date = row[soggiorno_col]
            formatted_date = ""
            
            if isinstance(stay_date, (datetime, date)):
                formatted_date = stay_date.strftime('%d/%m/%Y')
            elif isinstance(stay_date, str):
                # Controlla se c'è un prefisso del giorno della settimana (es. "Dom", "Lun", ecc.)
                date_parts = stay_date.split()
                if len(date_parts) > 1 and any(day in date_parts[0].lower() for day in 
                                           ["dom", "lun", "mar", "mer", "gio", "ven", "sab"]):
                    # Estrai solo la parte della data
                    date_only = date_parts[1] if len(date_parts) > 1 else stay_date
                    formatted_date = date_only
                else:
                    formatted_date = stay_date
            else:
                formatted_date = str(stay_date)
                
            # Create data entry with proper error handling
            try:
                pickup_entry = {
                    "hotel": hotel,
                    "stay_date": formatted_date,
                    "roomnights": int(float(row[roomnights_col])) if pd.notna(row[roomnights_col]) else 0,
                    "vs_1day": int(float(row[vs_1gg_col])) if vs_1gg_col and pd.notna(row[vs_1gg_col]) else 0,
                    "vs_2day": int(float(row[vs_2gg_col])) if vs_2gg_col and pd.notna(row[vs_2gg_col]) else 0,
                    "vs_3day": int(float(row[vs_3gg_col])) if vs_3gg_col and pd.notna(row[vs_3gg_col]) else 0,
                    "vs_7day": int(float(row[vs_7gg_col])) if pd.notna(row[vs_7gg_col]) else 0,
                    "vs_7day_spit": int(float(row[vs_7gg_spit_col])) if pd.notna(row[vs_7gg_spit_col]) else 0
                }
                pickup_data.append(pickup_entry)
            except Exception as e:
                st.warning(f"Errore riga {idx}: {str(e)}")
                continue
                
        if not pickup_data:
            return None, "Nessun dato valido trovato nel file"
        
        # Informazioni estratte
        extracted_info = {
            "room_type": metadata["room_type"],
            "date_period": metadata["date_period"],
            "pickup_data": pickup_data
        }
            
        return extracted_info, None
    except Exception as e:
        return None, f"Errore nell'elaborazione: {str(e)}"

def process_daily_production_excel(uploaded_file, hotel):
    """
    Elabora il file Excel di produzione giornaliera caricato con gestione migliorata
    delle date in formato italiano.
    """
    try:
        df = pd.read_excel(uploaded_file)
        
        # Rileva metadati dal report
        metadata = detect_report_metadata(df)
        
        # Mappatura colonne per il report di produzione giornaliera
        expected_structure = {
            'date_col': {'names': ['Data', 'data', 'date'], 'index': 0, 'letter': 'A'},
            'weekday_col': {'names': ['Giorno', 'giorno', 'day'], 'index': 1, 'letter': 'B'},
            'occ_col': {'names': ['% Occ.', '% occ', 'occupancy', 'occupazione'], 'index': 2, 'letter': 'C'},
            'occ_vs_spit_col': {'names': ['vs SPIT', 'vs spit'], 'index': 3, 'letter': 'D'},
            'occ_vs_ap_col': {'names': ['vs AP', 'vs ap'], 'index': 4, 'letter': 'E'},
            'room_nights_col': {'names': ['Room nights', 'room nights', 'roomnights'], 'index': 5, 'letter': 'F'},
            'rn_vs_spit_col': {'names': ['vs SPIT', 'vs spit'], 'index': 6, 'letter': 'G'},
            'rn_vs_ap_col': {'names': ['vs AP', 'vs ap'], 'index': 7, 'letter': 'H'},
            'adr_col': {'names': ['ADR Cam', 'adr cam', 'adr'], 'index': 11, 'letter': 'L'},
            'adr_vs_spit_col': {'names': ['vs SPIT', 'vs spit'], 'index': 12, 'letter': 'M'},
            'adr_vs_ap_col': {'names': ['vs AP', 'vs ap'], 'index': 13, 'letter': 'N'},
            'room_rev_col': {'names': ['Room Revenue', 'room revenue', 'revenue'], 'index': 17, 'letter': 'R'},
            'rev_vs_spit_col': {'names': ['vs SPIT', 'vs spit'], 'index': 18, 'letter': 'S'},
            'rev_vs_ap_col': {'names': ['vs AP', 'vs ap'], 'index': 19, 'letter': 'T'}
        }
        
        col_map = {}
        column_list = list(df.columns)
        num_columns = len(column_list)
        
        # Funzione per trovare colonne per nome o posizione
        def find_column(target_names, context="", letter=None):
            context_lower = context.lower()
            
            # Prova a trovare per lettera della colonna
            if letter:
                try:
                    idx = excel_column_to_index(letter)
                    if idx < num_columns:
                        return column_list[idx]
                except:
                    pass
            
            # Prova a trovare per nome esatto
            for name in target_names:
                if name in column_list:
                    return name
            
            # Prova a trovare per nome case-insensitive
            for col in column_list:
                col_str = str(col).lower()
                for name in target_names:
                    if name.lower() == col_str:
                        return col
            
            # Prova a trovare con contesto
            if context:
                for col in column_list:
                    col_str = str(col).lower()
                    for name in target_names:
                        if name.lower() in col_str and context_lower in col_str:
                            return col
            
            # Prova a trovare per contenuto parziale
            for col in column_list:
                col_str = str(col).lower()
                for name in target_names:
                    if name.lower() in col_str:
                        return col
            
            return None
        
        # Costruisci la mappatura delle colonne
        for key, config in expected_structure.items():
            context = ""
            if "adr" in key:
                context = "adr"
            elif "occ" in key:
                context = "occ"
            elif "room_nights" in key or "rn" in key:
                context = "room nights"
            elif "rev" in key:
                context = "revenue"
            
            letter = config.get('letter')
            index = config['index']
            
            # Prova prima per lettera, poi per indice, poi per nome
            if letter:
                try:
                    idx = excel_column_to_index(letter)
                    if idx < num_columns:
                        col_map[key] = column_list[idx]
                        continue
                except:
                    pass
            
            if index < num_columns:
                col_map[key] = column_list[index]
            else:
                found_col = find_column(config['names'], context)
                if found_col:
                    col_map[key] = found_col
        
        if 'date_col' not in col_map or col_map['date_col'] not in df.columns:
            return None, "Colonna data non trovata. Verificare formato file."
        
        # Pulizia dati: rimuovi righe vuote e intestazioni duplicate
        df = df.dropna(subset=[col_map['date_col']])
        
        # Normalizza percentuali
        if 'occ_col' in col_map and col_map['occ_col'] in df.columns:
            occ_col = col_map['occ_col']
            if df[occ_col].dtype == 'object':
                df[occ_col] = df[occ_col].astype(str).str.replace('%', '').str.replace(',', '.').astype(float)
            if df[occ_col].max() <= 1.0:
                df[occ_col] = df[occ_col] * 100
        
        # Estrai dati giornalieri
        daily_data = []
        
        for idx, row in df.iterrows():
            try:
                # Verifica che sia una riga valida con data
                if pd.isna(row[col_map['date_col']]):
                    continue
                
                date_val = row[col_map['date_col']]
                
                # Salta righe con filtri o intestazioni
                if isinstance(date_val, str) and ("filtri" in date_val.lower() or "data" in date_val.lower()):
                    continue
                
                # Gestione speciale per date in formato italiano con prefisso giorno
                # Ad esempio: "Dom 01/06/2025" -> "01/06/2025"
                formatted_date = ""
                
                if isinstance(date_val, (datetime, date)):
                    formatted_date = date_val.strftime('%d/%m/%Y')
                elif isinstance(date_val, str):
                    # Controlla se c'è un prefisso del giorno della settimana (es. "Dom", "Lun", ecc.)
                    date_parts = date_val.split()
                    if len(date_parts) > 1 and any(day in date_parts[0].lower() for day in 
                                                ["dom", "lun", "mar", "mer", "gio", "ven", "sab"]):
                        # Estrai solo la parte della data
                        date_only = date_parts[1] if len(date_parts) > 1 else date_val
                        formatted_date = date_only
                    else:
                        formatted_date = date_val
                else:
                    formatted_date = str(date_val)
                
                # Estrai i dati richiesti con gestione errori
                daily_entry = {
                    "date": formatted_date,
                    "hotel": hotel
                }
                
                # Giorno della settimana
                if 'weekday_col' in col_map and pd.notna(row[col_map['weekday_col']]):
                    daily_entry["weekday"] = str(row[col_map['weekday_col']])
                
                # Room nights
                if 'room_nights_col' in col_map and pd.notna(row[col_map['room_nights_col']]):
                    rn_val = row[col_map['room_nights_col']]
                    if isinstance(rn_val, str):
                        rn_val = rn_val.replace('.', '').replace(',', '.')
                    daily_entry["room_nights"] = int(float(rn_val))
                
                # Occupazione
                if 'occ_col' in col_map and pd.notna(row[col_map['occ_col']]):
                    occ_val = row[col_map['occ_col']]
                    if isinstance(occ_val, str):
                        occ_val = occ_val.replace('%', '').replace(',', '.')
                    daily_entry["occupancy"] = float(occ_val)
                
                # ADR
                if 'adr_col' in col_map and pd.notna(row[col_map['adr_col']]):
                    adr_val = row[col_map['adr_col']]
                    if isinstance(adr_val, str):
                        adr_val = adr_val.replace('.', '').replace(',', '.')
                    daily_entry["adr"] = float(adr_val)
                
                # Room Revenue
                if 'room_rev_col' in col_map and pd.notna(row[col_map['room_rev_col']]):
                    rev_val = row[col_map['room_rev_col']]
                    if isinstance(rev_val, str):
                        rev_val = rev_val.replace('.', '').replace(',', '.')
                    daily_entry["room_revenue"] = float(rev_val)
                
                # vs SPIT RN
                if 'rn_vs_spit_col' in col_map and pd.notna(row[col_map['rn_vs_spit_col']]):
                    val = row[col_map['rn_vs_spit_col']]
                    if isinstance(val, str):
                        val = val.replace('.', '').replace(',', '.')
                    daily_entry["vs_spit_rn"] = int(float(val))
                
                # vs AP RN
                if 'rn_vs_ap_col' in col_map and pd.notna(row[col_map['rn_vs_ap_col']]):
                    val = row[col_map['rn_vs_ap_col']]
                    if isinstance(val, str):
                        val = val.replace('.', '').replace(',', '.')
                    daily_entry["vs_ap_rn"] = int(float(val))
                
                # vs SPIT Occ
                if 'occ_vs_spit_col' in col_map and pd.notna(row[col_map['occ_vs_spit_col']]):
                    val = row[col_map['occ_vs_spit_col']]
                    if isinstance(val, str):
                        val = val.replace('%', '').replace(',', '.')
                    daily_entry["vs_spit_occ"] = float(val)
                
                # vs AP Occ
                if 'occ_vs_ap_col' in col_map and pd.notna(row[col_map['occ_vs_ap_col']]):
                    val = row[col_map['occ_vs_ap_col']]
                    if isinstance(val, str):
                        val = val.replace('%', '').replace(',', '.')
                    daily_entry["vs_ap_occ"] = float(val)
                
                # vs SPIT ADR
                if 'adr_vs_spit_col' in col_map and pd.notna(row[col_map['adr_vs_spit_col']]):
                    val = row[col_map['adr_vs_spit_col']]
                    if isinstance(val, str):
                        val = val.replace('.', '').replace(',', '.')
                    daily_entry["vs_spit_adr"] = float(val)
                
                # vs AP ADR
                if 'adr_vs_ap_col' in col_map and pd.notna(row[col_map['adr_vs_ap_col']]):
                    val = row[col_map['adr_vs_ap_col']]
                    if isinstance(val, str):
                        val = val.replace('.', '').replace(',', '.')
                    daily_entry["vs_ap_adr"] = float(val)
                
                daily_data.append(daily_entry)
            except Exception as e:
                st.warning(f"Errore processando riga {idx}: {e}")
                continue
        
        if not daily_data:
            return None, "Nessun dato valido trovato nel file di produzione"
        
        # Informazioni estratte
        extracted_info = {
            "room_type": metadata["room_type"],
            "date_period": metadata["date_period"],
            "days_count": len(daily_data),
            "daily_data": daily_data
        }
        
        return extracted_info, None
    except Exception as e:
        return None, f"Errore nell'elaborazione del file di produzione: {str(e)}"

def map_bi_data_to_model(daily_production=None, pickup_data=None):
    """
    Mappa i dati dalla BI al formato richiesto dal modello di upgrade.
    """
    historical_data = {}
    otb_data = {}
    pickup_mapped = {}
    
    # Mapping dati storici dalla produzione giornaliera
    if daily_production:
        # Converti date in oggetti datetime per ordinamento e confronto
        for entry in daily_production:
            try:
                date_str = entry["date"]
                date_parts = date_str.split('/')
                if len(date_parts) == 3:
                    day, month, year = map(int, date_parts)
                    date_obj = datetime(year, month, day)
                    
                    # I dati di produzione rappresentano lo storico
                    historical_data[date_obj] = entry["room_nights"]
            except Exception as e:
                st.warning(f"Errore nel mapping data: {date_str} - {str(e)}")
    
    # Mapping dati pickup
    if pickup_data:
        for entry in pickup_data:
            try:
                date_str = entry["stay_date"]
                date_parts = date_str.split('/')
                if len(date_parts) == 3:
                    day, month, year = map(int, date_parts)
                    date_obj = datetime(year, month, day)
                    
                    # Le room nights correnti dal pickup sono l'OTB
                    otb_data[date_obj] = entry["roomnights"]
                    
                    # Costruisci i dati di pickup
                    pickup_mapped[date_obj] = {
                        "2025": entry["vs_7day"],
                        "2024": entry["vs_7day_spit"]
                    }
            except Exception as e:
                st.warning(f"Errore nel mapping pickup: {date_str} - {str(e)}")
    
    return historical_data, otb_data, pickup_mapped

def save_pickup_data(hotel, pickup_data):
    """Salva i dati di pickup nel database"""
    conn = get_connection()
    cursor = conn.cursor()
    
    saved_count = 0
    try:
        for entry in pickup_data:
            # Verifica se esiste già un record per questa data
            cursor.execute("""
            SELECT id FROM pickup_data 
            WHERE hotel = ? AND stay_date = ?
            """, (hotel, entry['stay_date']))
            
            existing = cursor.fetchone()
            
            if existing:
                # Aggiorna record esistente
                cursor.execute("""
                UPDATE pickup_data 
                SET roomnights = ?, vs_1day = ?, vs_2day = ?, vs_3day = ?, 
                    vs_7day = ?, vs_7day_spit = ?, insertion_date = CURRENT_TIMESTAMP
                WHERE hotel = ? AND stay_date = ?
                """, (
                    entry['roomnights'], entry.get('vs_1day', 0), entry.get('vs_2day', 0), 
                    entry.get('vs_3day', 0), entry.get('vs_7day', 0), entry.get('vs_7day_spit', 0),
                    hotel, entry['stay_date']
                ))
            else:
                # Inserisci nuovo record
                cursor.execute("""
                INSERT INTO pickup_data 
                (hotel, stay_date, roomnights, vs_1day, vs_2day, vs_3day, vs_7day, vs_7day_spit)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    hotel, entry['stay_date'], entry['roomnights'], 
                    entry.get('vs_1day', 0), entry.get('vs_2day', 0), entry.get('vs_3day', 0),
                    entry.get('vs_7day', 0), entry.get('vs_7day_spit', 0)
                ))
            
            saved_count += 1
        
        conn.commit()
        return saved_count
    except Exception as e:
        conn.rollback()
        st.error(f"Errore nel salvataggio dei dati pickup: {e}")
        return 0
    finally:
        conn.close()

def save_production_data(hotel, production_data):
    """Salva i dati di produzione nel database"""
    conn = get_connection()
    cursor = conn.cursor()
    
    saved_count = 0
    try:
        for entry in production_data:
            # Verifica se esiste già un record per questa data
            cursor.execute("""
            SELECT id FROM production_data 
            WHERE hotel = ? AND date = ?
            """, (hotel, entry['date']))
            
            existing = cursor.fetchone()
            
            if existing:
                # Aggiorna record esistente
                cursor.execute("""
                UPDATE production_data 
                SET weekday = ?, room_nights = ?, occupancy = ?, adr = ?, room_revenue = ?,
                    vs_spit_rn = ?, vs_spit_occ = ?, vs_spit_adr = ?,
                    vs_ap_rn = ?, vs_ap_occ = ?, vs_ap_adr = ?,
                    insertion_date = CURRENT_TIMESTAMP
                WHERE hotel = ? AND date = ?
                """, (
                    entry.get('weekday', ''), entry.get('room_nights', 0), 
                    entry.get('occupancy', 0), entry.get('adr', 0), 
                    entry.get('room_revenue', 0),
                    entry.get('vs_spit_rn', 0), entry.get('vs_spit_occ', 0), 
                    entry.get('vs_spit_adr', 0),
                    entry.get('vs_ap_rn', 0), entry.get('vs_ap_occ', 0), 
                    entry.get('vs_ap_adr', 0),
                    hotel, entry['date']
                ))
            else:
                # Inserisci nuovo record
                cursor.execute("""
                INSERT INTO production_data 
                (hotel, date, weekday, room_nights, occupancy, adr, room_revenue,
                vs_spit_rn, vs_spit_occ, vs_spit_adr, vs_ap_rn, vs_ap_occ, vs_ap_adr)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    hotel, entry['date'], entry.get('weekday', ''), 
                    entry.get('room_nights', 0), entry.get('occupancy', 0), 
                    entry.get('adr', 0), entry.get('room_revenue', 0),
                    entry.get('vs_spit_rn', 0), entry.get('vs_spit_occ', 0), 
                    entry.get('vs_spit_adr', 0),
                    entry.get('vs_ap_rn', 0), entry.get('vs_ap_occ', 0), 
                    entry.get('vs_ap_adr', 0)
                ))
            
            saved_count += 1
        
        conn.commit()
        return saved_count
    except Exception as e:
        conn.rollback()
        st.error(f"Errore nel salvataggio dei dati produzione: {e}")
        return 0
    finally:
        conn.close()

def get_latest_excel_data(hotel, section="Produzione Giornaliera"):
    """Recupera i dati più recenti caricati per un hotel e una sezione"""
    if section == "Produzione Giornaliera":
        return {"daily_data": get_production_data(hotel)}
    elif section == "Pickup":
        return {"pickup_data": get_pickup_data(hotel)}
    else:
        return None

def get_pickup_data(hotel, date_from=None, date_to=None):
    """Retrieve pickup data from database with optional date filtering."""
    conn = get_connection()
    c = conn.cursor()
    
    query = "SELECT * FROM pickup_data WHERE hotel = ?"
    params = [hotel]
    
    # Ensure date parameters are properly formatted for comparison
    if date_from:
        # Convert to standard format if needed
        try:
            if isinstance(date_from, (datetime, date)):
                date_from = date_from.strftime('%d/%m/%Y')
            query += " AND stay_date >= ?"
            params.append(date_from)
        except Exception as e:
            print(f"Error formatting date_from parameter: {str(e)}")
    
    if date_to:
        # Convert to standard format if needed
        try:
            if isinstance(date_to, (datetime, date)):
                date_to = date_to.strftime('%d/%m/%Y')
            query += " AND stay_date <= ?"
            params.append(date_to)
        except Exception as e:
            print(f"Error formatting date_to parameter: {str(e)}")
    
    query += " ORDER BY stay_date ASC"
    
    try:
        c.execute(query, params)
        results = c.fetchall()
        
        # Process results for consistency
        processed_results = []
        for row in results:
            row_dict = dict(row)
            
            # Ensure stay_date is properly formatted
            if "stay_date" in row_dict and row_dict["stay_date"]:
                try:
                    # If it's not already a properly formatted string
                    if not isinstance(row_dict["stay_date"], str) or not re.match(r'\d{2}/\d{2}/\d{4}', row_dict["stay_date"]):
                        if isinstance(row_dict["stay_date"], (datetime, date)):
                            row_dict["stay_date"] = row_dict["stay_date"].strftime('%d/%m/%Y')
                except Exception:
                    # Keep as is if conversion fails
                    pass
                    
            processed_results.append(row_dict)
            
        return processed_results
    except Exception as e:
        print(f"Error retrieving pickup data: {str(e)}")
        return []
    finally:
        conn.close()

def get_production_data(hotel, date_from=None, date_to=None):
    """Recupera dati di produzione dal database con filtro opzionale per date"""
    conn = get_connection()
    c = conn.cursor()
    
    query = "SELECT * FROM production_data WHERE hotel = ?"
    params = [hotel]
    
    # Applica filtri per data
    if date_from:
        try:
            if isinstance(date_from, (datetime, date)):
                date_from = date_from.strftime('%d/%m/%Y')
            query += " AND date >= ?"
            params.append(date_from)
        except Exception as e:
            print(f"Errore formattazione date_from: {str(e)}")
    
    if date_to:
        try:
            if isinstance(date_to, (datetime, date)):
                date_to = date_to.strftime('%d/%m/%Y')
            query += " AND date <= ?"
            params.append(date_to)
        except Exception as e:
            print(f"Errore formattazione date_to: {str(e)}")
    
    query += " ORDER BY date ASC"
    
    try:
        c.execute(query, params)
        results = c.fetchall()
        
        processed_results = []
        for row in results:
            processed_results.append(dict(row))
            
        return processed_results
    except Exception as e:
        print(f"Errore recupero dati produzione: {str(e)}")
        return []
    finally:
        conn.close()

# Funzioni per calcoli e visualizzazioni
def calculate_probabilities(historical_data):
    """
    Calcola la probabilità di avere una domanda >= a un dato valore.
    Basato sui dati storici delle room nights.
    """
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

def calculate_dashboard_data(historical_data, otb_data, pickup_data, params, room_type_config=None):
    """
    Calcola i dati per la dashboard, includendo controlli di inventario e tipo di camera.
    """
    # Calcola probabilità dalla storico
    probabilities = calculate_probabilities(historical_data)
    
    # Estrai informazioni sulla tipologia di camera
    is_entry_level = params.get('is_entry_level', False)
    rooms_in_type = params.get('rooms_in_type', 0)
    if room_type_config:
        is_entry_level = room_type_config.get('is_entry_level', is_entry_level)
        rooms_in_type = room_type_config.get('rooms_in_type', rooms_in_type)
    
    dashboard_data = []
    for date, rn in otb_data.items():
        # Capacità effettiva - se disponibile
        effective_capacity = rooms_in_type
        if room_type_config and 'id' in room_type_config and 'hotel_id' in room_type_config:
            capacity = get_effective_capacity(
                room_type_config['hotel_id'],
                room_type_config['id'],
                date
            )
            if capacity > 0:
                effective_capacity = capacity
        
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
        
        # Check se siamo in overbooking per questa tipologia
        is_overbooking = rn > effective_capacity
        
        # Decisione considerando entry-level e overbooking
        if is_entry_level or is_overbooking:
            recommendation = "No"
        else:
            recommendation = "Sì" if exp_revenue < threshold else "No"
        
        # Suggerimento personalizzato
        if is_overbooking:
            suggestion = f"Attenzione: overbooking di {rn - effective_capacity} camere"
        elif is_entry_level:
            suggestion = "Tipologia entry-level: non applicabile per upgrade"
        elif recommendation == "Sì":
            suggestion = "Bassa probabilità di vendita - considera upgrade gratuito"
        elif prob > 0.7:
            suggestion = "Alta probabilità di vendita - non fare upgrade"
        else:
            suggestion = "Probabilità media - valuta disponibilità e strategia"
        
        # Percentuale occupazione tipologia
        occupancy_pct = (rn / effective_capacity) * 100 if effective_capacity > 0 else 0
        
        dashboard_data.append({
            "Data": date,
            "RN Attuali": rn,
            "Probabilità": prob,
            "Expected Revenue": exp_revenue,
            "Delta Pickup": delta,
            "Soglia Upgrade": threshold,
            "Upgrade Consigliato": recommendation,
            "Suggerimento": suggestion,
            "% Occupazione Tipologia": occupancy_pct,
            "È Overbooking": is_overbooking,
            "Capacità Effettiva": effective_capacity
        })
    
    return dashboard_data

def create_excel_file(params, historical_data, otb_data, pickup_data, room_type_name="Standard"):
    """
    Crea un file Excel con tutte le analisi e i suggerimenti.
    """
    # Crea nuovo workbook
    wb = Workbook()
    
    # Definizione degli stili professionali
    # Colori aziendali
    color_primary = "4472C4"  # Blu aziendale
    color_secondary = "ED7D31"  # Arancione complementare
    color_accent = "70AD47"  # Verde per accenti positivi
    color_negative = "C00000"  # Rosso per valori negativi
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
    accent_font = Font(name='Calibri', size=11, color=color_accent, bold=True)
    negative_font = Font(name='Calibri', size=11, color=color_negative, bold=True)

    # Stili di riempimento
    header_fill = PatternFill(start_color=color_header_bg, end_color=color_header_bg, fill_type="solid")
    input_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # Giallo chiaro più professionale
    alt_row_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Grigio chiaro per righe alternate

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

    # Dati AGGIORNATI con nuovi parametri
    parameters = [
        ["ADR Medio", params['adr'], "Tariffa media giornaliera in €"],
        ["Soglia Upgrade (fissa)", params['upgrade_threshold'], "Non utilizzata nella versione corrente"],
        ["Margine minimo richiesto", params['min_margin'], "Espresso come percentuale (0.6 = 60%)"],
        ["Giorni di analisi", params['days'], "Numero di giorni inclusi nell'analisi"],
        ["Tipologia", room_type_name, "Nome della tipologia analizzata"],
        ["È Entry-Level?", "Sì" if params['is_entry_level'] else "No", "Se Sì, questa tipologia non riceve suggerimenti di upgrade"],
        ["Camere in questa tipologia", params['rooms_in_type'], "Numero di camere disponibili per questa tipologia"],
        ["Camere totali struttura", params['total_rooms'], "Numero totale di camere nella struttura"]
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

    # Creazione altri fogli (storico, OTB, dashboard, ecc.)
    # Codice per altri fogli Excel...
    
    # --------------------------
    # 2. Foglio STORICO 2024 (unificato)
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

    # Dati
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

    # Dati
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

    # Dati
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
    ws_dash.merge_cells('A1:I1')
    ws_dash['A1'] = f"DASHBOARD DECISIONALE UPGRADE - {room_type_name} - GIUGNO 2025"
    ws_dash['A1'].font = title_font
    ws_dash['A1'].alignment = center_align

    # Intestazioni - AGGIUNTE COLONNE
    headers = [
        "Data", "RN Attuali", "Prob. Domanda ≥ RN", "Expected Revenue",
        "Delta Pickup vs LY", "Soglia Upgrade Dinamica (€)",
        "% Occupazione", "Upgrade Consigliato", "Suggerimento Operativo"
    ]
    ws_dash.append(headers)
    for col in range(1, 10):
        cell = ws_dash.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = center_align

    # Altri fogli e formattazioni...
    
    # Imposta dashboard come foglio attivo
    wb.active = wb["Dashboard Upgrade"]
    
    # Salva in buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer

# Funzione di configurazione delle tipologie del Cala Cuncheddi
def setup_cala_cuncheddi_presets():
    """
    Configura automaticamente le tipologie di camera del Cala Cuncheddi
    con le relative gerarchie di upgrade.
    """
    # Verifica se l'hotel esiste già
    hotel = get_hotel_by_name("Cala Cuncheddi")
    if not hotel:
        # Crea l'hotel se non esiste
        hotel_id = add_hotel({
            'hotel_name': "Cala Cuncheddi",
            'total_rooms': 85,
            'address': "Via Dei Ginepri 3, 07026 Pittulongu, Olbia",
            'city': "Olbia",
            'stars': 5,
            'seasonal': True  # Hotel stagionale
        })
    else:
        hotel_id = hotel['id']
    
    # Definisci le tipologie di camera con le quantità aggiornate
    room_types = [
        {
            'name': "Classic Garden",
            'count': 33,
            'is_entry_level': True,
            'adr': 300,
            'upgrade_to': "Classic Sea View"
        },
        {
            'name': "Classic Sea View",
            'count': 12,
            'is_entry_level': False,
            'adr': 350,
            'upgrade_to': "Superior Sea View"
        },
        {
            'name': "Family",
            'count': 2,
            'is_entry_level': False,
            'adr': 495,
            'upgrade_to': "Junior Suite Pool"
        },
        {
            'name': "Superior Sea View",
            'count': 18,
            'is_entry_level': False,
            'adr': 400,
            'upgrade_to': "Executive"
        },
        {
            'name': "Executive",
            'count': 9,
            'is_entry_level': False,
            'adr': 450,
            'upgrade_to': "Deluxe"
        },
        {
            'name': "Deluxe",
            'count': 4,
            'is_entry_level': False,
            'adr': 520,
            'upgrade_to': "Junior Suite Pool"
        },
        {
            'name': "Junior Suite Pool",
            'count': 4,
            'is_entry_level': False,
            'adr': 650,
            'upgrade_to': "Suite"
        },
        {
            'name': "Suite",
            'count': 3,
            'is_entry_level': False,
            'adr': 780,
            'upgrade_to': None
        }
    ]
    
    # Aggiungi/aggiorna tutte le tipologie
    for rt in room_types:
        room_type = get_room_type_by_name(hotel_id, rt['name'])
        if room_type:
            # Aggiorna la tipologia esistente
            update_room_type(room_type['id'], {
                'room_type_name': rt['name'],
                'rooms_in_type': rt['count'],
                'is_entry_level': rt['is_entry_level'],
                'adr': rt['adr'],
                'upgrade_target_type': rt['upgrade_to']
            })
        else:
            # Crea nuova tipologia
            add_room_type(hotel_id, {
                'room_type_name': rt['name'],
                'rooms_in_type': rt['count'],
                'is_entry_level': rt['is_entry_level'],
                'adr': rt['adr'],
                'upgrade_target_type': rt['upgrade_to'],
                'min_margin': 0.6  # Margine predefinito
            })
    
    return True

# Interfaccia di importazione Excel migliorata
def render_excel_import_section(selected_hotels):
    """
    Renderizza la sezione di importazione Excel con interfaccia migliorata.
    """
    st.markdown('<h2 class="section-title">Importa file Excel</h2>', unsafe_allow_html=True)
    
    excel_file = st.file_uploader("Carica file Excel esportato da Power BI", type=["xlsx", "xls"])
    
    if excel_file:
        st.success(f"File caricato: {excel_file.name}")
        
        # Anteprima e rilevamento automatico
        try:
            df_preview = pd.read_excel(excel_file)
            st.write("Anteprima dati Excel:")
            st.dataframe(df_preview.head())
            
            # Rilevamento automatico del tipo di report
            report_type, confidence = detect_report_type(df_preview)
            
            # Rileva metadati aggiuntivi
            metadata = detect_report_metadata(df_preview)
            
            if report_type:
                st.info(f"Tipo di report rilevato automaticamente: **{report_type}** (confidenza: {confidence}/4)")
                default_index = ["Produzione Giornaliera", "Prod. per Segmento", "Produzione Portafoglio", "Pickup"].index(report_type) if report_type in ["Produzione Giornaliera", "Prod. per Segmento", "Produzione Portafoglio", "Pickup"] else 0
            else:
                st.warning("Non è stato possibile rilevare automaticamente il tipo di report. Seleziona manualmente.")
                default_index = 0
            
            # Mostra informazioni rilevate
            info_items = []
            if metadata["room_type"]:
                info_items.append(f"**Tipologia camera: {metadata['room_type']}**")
            if metadata["date_period"]:
                info_items.append(f"**Periodo: {metadata['date_period']}**")
            if metadata["hotel"]:
                info_items.append(f"**Hotel: {metadata['hotel']}**")
                
            if info_items:
                st.markdown(f"Informazioni rilevate: {' | '.join(info_items)}")
                
        except Exception as e:
            st.error(f"Errore nell'anteprima del file: {str(e)}")
            default_index = 0
        
        # L'utente può confermare o modificare il tipo di report
        report_type = st.selectbox("Tipo di report Excel", 
                                ["Produzione Giornaliera", "Prod. per Segmento", "Produzione Portafoglio", "Pickup"],
                                index=default_index,
                                help="Conferma o modifica il tipo di report rilevato automaticamente")
        
        # Seleziona hotel, con preferenza per quello rilevato
        default_hotel = next((i for i, h in enumerate(selected_hotels) if metadata.get("hotel") and h.lower() == metadata["hotel"].lower()), 0)
        hotel = st.selectbox("Seleziona hotel:", selected_hotels, index=default_hotel)
        
        # Selezione tipologia camera (se rilevata)
        if metadata.get("room_type"):
            room_types = ["Classic Garden", "Classic Sea View", "Superior Sea View", 
                         "Executive", "Deluxe", "Family", "Junior Suite Pool", "Suite"]
            
            # Trova l'indice della tipologia rilevata
            room_type_index = next((i for i, rt in enumerate(room_types) if rt.lower() == metadata["room_type"].lower()), 0)
            
            selected_room_type = st.selectbox("Tipologia camera:", 
                                           room_types,
                                           index=room_type_index,
                                           help="Conferma o modifica la tipologia camera rilevata")
        
        if st.button("Elabora Excel"):
            with st.spinner("Elaborazione file Excel in corso..."):
                progress_container = st.empty()
                progress_container.markdown('<div class="processing-step progress">Elaborazione file Excel in corso...</div>', unsafe_allow_html=True)
                
                if report_type == "Pickup":
                    # Elaborazione file pickup
                    extracted_info, error = process_pickup_excel(excel_file, hotel)
                    
                    if error:
                        progress_container.markdown(f'<div class="processing-step error">Errore: {error}</div>', unsafe_allow_html=True)
                    elif extracted_info:
                        rows_saved = save_pickup_data(hotel, extracted_info["pickup_data"])
                        progress_container.markdown('<div class="processing-step completed">Elaborazione file Excel completata ✓</div>', unsafe_allow_html=True)
                        
                        # Mostra riepilogo
                        st.markdown("### Riepilogo Elaborazione")
                        st.markdown(f"**Tipo di Report:** {report_type}")
                        st.markdown(f"**Hotel:** {hotel}")
                        
                        if extracted_info.get("room_type"):
                            st.markdown(f"**Tipologia Camera:** {extracted_info['room_type']}")
                        
                        if extracted_info.get("date_period"):
                            st.markdown(f"**Periodo:** {extracted_info['date_period']}")
                        
                        st.success(f"File pickup elaborato con successo: **{rows_saved}** righe salvate")
                        
                        # Mostra sample dei dati elaborati
                        st.markdown("### Esempio dati pickup elaborati:")
                        display_data = pd.DataFrame(extracted_info["pickup_data"][:10])
                        st.dataframe(display_data)
                
                elif report_type == "Produzione Giornaliera":
                    # Produzione Giornaliera con gestione migliorata
                    extracted_info, error = process_daily_production_excel(excel_file, hotel)
                    
                    if error:
                        progress_container.markdown(f'<div class="processing-step error">Errore: {error}</div>', unsafe_allow_html=True)
                    elif extracted_info:
                        # Mostra risultati
                        progress_container.markdown('<div class="processing-step completed">Elaborazione file Excel completata ✓</div>', unsafe_allow_html=True)
                        
                        st.markdown("### Riepilogo Elaborazione")
                        st.markdown(f"**Tipo di Report:** {report_type}")
                        st.markdown(f"**Hotel:** {hotel}")
                        
                        if extracted_info.get("room_type"):
                            st.markdown(f"**Tipologia Camera:** {extracted_info['room_type']}")
                        
                        if extracted_info.get("date_period"):
                            st.markdown(f"**Periodo:** {extracted_info['date_period']}")
                        
                        st.success(f"File produzione elaborato: **{len(extracted_info['daily_data'])} giorni** trovati")
                        
                        # Mostra sample dei dati elaborati
                        st.markdown("### Esempio dati elaborati:")
                        if len(extracted_info['daily_data']) > 0:
                            display_data = pd.DataFrame(extracted_info['daily_data'][:10])
                            st.dataframe(display_data)
                            
                            # Salva i dati
                            saved_count = save_production_data(hotel, extracted_info['daily_data'])
                            st.success(f"✅ {saved_count} record salvati nel database")
                
                else:
                    # Altri tipi di report
                    st.info(f"Elaborazione per il tipo '{report_type}' non ancora implementata")

# Layout dell'interfaccia principale
def main():
    """
    Funzione principale dell'applicazione Streamlit.
    """
    st.title("🏨 Hotel Upgrade Advisor Pro")
    st.write("Sistema avanzato per decisioni di revenue management con controllo inventario")

    # Inizializza il database
    init_database()
    
    # Sidebar per configurazione
    st.sidebar.header("Configurazione")

    # Verifica se esiste una configurazione per Cala Cuncheddi
    hotel = get_hotel_by_name("Cala Cuncheddi")
    if not hotel:
        if st.sidebar.button("Configura Cala Cuncheddi"):
            if setup_cala_cuncheddi_presets():
                st.sidebar.success("Cala Cuncheddi configurato con successo!")
                hotel = get_hotel_by_name("Cala Cuncheddi")
            else:
                st.sidebar.error("Errore nella configurazione di Cala Cuncheddi")

    # Selezione hotel
    hotels = get_all_hotels()
    if hotels:
        hotel_names = [h['hotel_name'] for h in hotels]
        default_index = hotel_names.index("Cala Cuncheddi") if "Cala Cuncheddi" in hotel_names else 0
        selected_hotel_name = st.sidebar.selectbox("Seleziona Hotel", hotel_names, index=default_index)
        selected_hotel = get_hotel_by_name(selected_hotel_name)
    else:
        st.sidebar.warning("Nessun hotel configurato")
        selected_hotel_name = "Hotel Demo"
        selected_hotel = None

    # Selezione tipologia camera
    room_types = []
    if selected_hotel:
        room_types = get_room_types_by_hotel(selected_hotel['id'])
        if room_types:
            room_type_names = [rt['room_type_name'] for rt in room_types]
            selected_room_type_name = st.sidebar.selectbox("Seleziona Tipologia", room_type_names)
            selected_room_type = next((rt for rt in room_types if rt['room_type_name'] == selected_room_type_name), None)
        else:
            st.sidebar.warning(f"Nessuna tipologia configurata per {selected_hotel_name}")
            selected_room_type_name = "Standard"
            selected_room_type = None
    else:
        selected_room_type_name = "Standard"
        selected_room_type = None

    # Definisci total_rooms con valore predefinito
    total_rooms = 85  # Valore predefinito per Cala Cuncheddi

    # Parametri di base
    with st.sidebar.expander("Parametri di Base", expanded=True):
        # Se abbiamo un hotel e una tipologia selezionati, usa i valori configurati
        if selected_room_type:
            adr = st.number_input("ADR Medio (€)", min_value=50.0, max_value=1000.0, value=selected_room_type.get('adr', 300.0), step=5.0)
            min_margin = st.slider("Margine minimo richiesto", min_value=0.1, max_value=0.9, 
                                  value=selected_room_type.get('min_margin', 0.6), step=0.05, format="%.2f")
            upgrade_threshold = st.number_input("Soglia Upgrade fissa (€)", min_value=0.0, max_value=500.0, 
                                              value=selected_room_type.get('upgrade_threshold', 100.0), step=10.0)
            days = st.slider("Giorni di analisi", min_value=7, max_value=60, value=30, step=1)
            
            # Parametri inventario direttamente dalla configurazione salvata
            is_entry_level = selected_room_type.get('is_entry_level', False)
            rooms_in_type = selected_room_type.get('rooms_in_type', 20)
            # Ottieni total_rooms dall'hotel selezionato
            if selected_hotel:
                total_rooms = selected_hotel.get('total_rooms', 85)
                
            st.write(f"**Camere in questa tipologia:** {rooms_in_type}")
            st.write(f"**Tipologia entry-level:** {'Sì' if is_entry_level else 'No'}")
            st.write(f"**Totale camere hotel:** {total_rooms}")
        else:
            # Valori di default
            adr = st.number_input("ADR Medio (€)", min_value=50.0, max_value=1000.0, value=300.0, step=5.0)
            min_margin = st.slider("Margine minimo richiesto", min_value=0.1, max_value=0.9, value=0.6, step=0.05, format="%.2f")
            upgrade_threshold = st.number_input("Soglia Upgrade fissa (€)", min_value=0.0, max_value=500.0, value=100.0, step=10.0)
            days = st.slider("Giorni di analisi", min_value=7, max_value=60, value=30, step=1)
            
            # Parametri inventario manuali
            st.markdown("### Configurazione Inventario")
            is_entry_level = st.checkbox("È tipologia entry-level?", value=False)
            rooms_in_type = st.number_input("Camere in questa tipologia", min_value=1, max_value=1000, value=20, step=1)
            total_rooms = st.number_input("Camere totali struttura", min_value=1, max_value=5000, value=85, step=1)

    # Pulsante per salvare configurazione
    if selected_hotel and selected_room_type:
        if st.sidebar.button("Aggiorna Configurazione"):
            updated = update_room_type(selected_room_type['id'], {
                'room_type_name': selected_room_type_name,
                'rooms_in_type': rooms_in_type,
                'is_entry_level': is_entry_level,
                'adr': adr,
                'min_margin': min_margin,
                'upgrade_threshold': upgrade_threshold
            })
            if updated:
                st.sidebar.success(f"Configurazione per {selected_room_type_name} aggiornata!")
            else:
                st.sidebar.error("Errore nell'aggiornamento della configurazione")

    # Prepara parametri
    params = {
        'adr': adr,
        'min_margin': min_margin,
        'upgrade_threshold': upgrade_threshold,
        'days': days,
        'is_entry_level': is_entry_level,
        'room_type_name': selected_room_type_name,
        'rooms_in_type': rooms_in_type,
        'total_rooms': selected_hotel.get('total_rooms', total_rooms) if selected_hotel else total_rooms
    }

    # Menu principale
    st.sidebar.header("Menu")
    menu = st.sidebar.radio("Seleziona sezione", 
                           ["Dashboard", "Importa Excel", "Configurazione Hotel", "Analisi Dati", "Stato Inventario"])

    # Opzioni di caricamento dati
    st.sidebar.header("Caricamento Dati")
    data_option = st.sidebar.radio("Seleziona metodo di caricamento", 
                                  ["Carica dati da BI", "Carica file Excel", "Genera dati casuali", "Inserisci manualmente", "Usa dati salvati"])

    # Inizializza dati
    historical_data = {}
    otb_data = {}
    pickup_data = {}

    # Contenuto principale basato sul menu selezionato
    if menu == "Dashboard":
        st.header("🏨 Dashboard Upgrade Advisor")
        
        # Qui andrà il contenuto della dashboard...
        st.info("Seleziona una tipologia di camera e un metodo di caricamento dati per visualizzare la dashboard.")
        
    elif menu == "Importa Excel":
        # Utilizza la nuova funzione di importazione Excel migliorata
        render_excel_import_section([h['hotel_name'] for h in hotels] if hotels else ["Cala Cuncheddi", "Hotel Demo"])
        
    elif menu == "Configurazione Hotel":
        st.header("⚙️ Configurazione Hotel")
        
        # Tab per gestire hotel e tipologie
        tab1, tab2, tab3 = st.tabs(["Gestione Hotel", "Tipologie di Camera", "Camere Out of Order"])
        
        with tab1:
            st.subheader("Gestione Hotel")
            
            # Form per aggiungere/modificare hotel
            with st.form("hotel_form"):
                hotel_name = st.text_input("Nome Hotel", value=selected_hotel.get('hotel_name', '') if selected_hotel else '')
                col1, col2 = st.columns(2)
                total_rooms = col1.number_input("Camere Totali", min_value=1, max_value=5000, 
                                              value=selected_hotel.get('total_rooms', 85) if selected_hotel else 85)
                stars = col2.number_input("Stelle", min_value=1, max_value=5, 
                                        value=selected_hotel.get('stars', 4) if selected_hotel else 4)
                
                address = st.text_input("Indirizzo", value=selected_hotel.get('address', '') if selected_hotel else '')
                city = st.text_input("Città", value=selected_hotel.get('city', '') if selected_hotel else '')
                seasonal = st.checkbox("Hotel Stagionale", value=selected_hotel.get('seasonal', False) if selected_hotel else False)
                
                submit_button = st.form_submit_button("Salva Hotel")
                
                if submit_button:
                    hotel_data = {
                        'hotel_name': hotel_name,
                        'total_rooms': total_rooms,
                        'stars': stars,
                        'address': address,
                        'city': city,
                        'seasonal': seasonal
                    }
                    
                    if selected_hotel:
                        # Aggiorna hotel esistente
                        if update_hotel(selected_hotel['id'], hotel_data):
                            st.success(f"Hotel {hotel_name} aggiornato con successo!")
                        else:
                            st.error("Errore nell'aggiornamento dell'hotel")
                    else:
                        # Aggiungi nuovo hotel
                        hotel_id = add_hotel(hotel_data)
                        if hotel_id:
                            st.success(f"Hotel {hotel_name} aggiunto con successo!")
                        else:
                            st.error("Errore nell'aggiunta dell'hotel")
        
        with tab2:
            st.subheader("Gestione Tipologie di Camera")
            
            if not selected_hotel:
                st.warning("Seleziona prima un hotel dalla sidebar")
            else:
                # Mostra le tipologie esistenti
                room_types = get_room_types_by_hotel(selected_hotel['id'])
                
                if room_types:
                    st.markdown("### Tipologie Attuali")
                    room_df = pd.DataFrame(room_types)
                    room_df = room_df[['room_type_name', 'rooms_in_type', 'is_entry_level', 'adr', 'upgrade_target_type']]
                    room_df.columns = ['Tipologia', 'Camere', 'Entry Level', 'ADR', 'Upgrade a']
                    st.dataframe(room_df)
                
                # Form per aggiungere nuova tipologia
                st.markdown("### Aggiungi Nuova Tipologia")
                with st.form("room_type_form"):
                    room_type_name = st.text_input("Nome Tipologia")
                    col1, col2 = st.columns(2)
                    rooms_in_type = col1.number_input("Numero Camere", min_value=1, max_value=1000, value=10)
                    adr_type = col2.number_input("ADR", min_value=50.0, max_value=2000.0, value=300.0, step=10.0)
                    
                    col3, col4 = st.columns(2)
                    is_entry = col3.checkbox("È tipologia entry-level?")
                    
                    # Lista target upgrade (tutte le tipologie con ADR superiore)
                    upgrade_targets = ["Nessuno"] + [rt['room_type_name'] for rt in room_types if rt['adr'] > adr_type]
                    upgrade_to = col4.selectbox("Upgrade a", upgrade_targets)
                    
                    submit_button = st.form_submit_button("Salva Tipologia")
                    
                    if submit_button:
                        # Aggiungi nuova tipologia
                        room_type_data = {
                            'room_type_name': room_type_name,
                            'rooms_in_type': rooms_in_type,
                            'is_entry_level': is_entry,
                            'adr': adr_type,
                            'upgrade_target_type': upgrade_to if upgrade_to != "Nessuno" else None
                        }
                        
                        room_type_id = add_room_type(selected_hotel['id'], room_type_data)
                        if room_type_id:
                            st.success(f"Tipologia {room_type_name} aggiunta con successo!")
                        else:
                            st.error("Errore nell'aggiunta della tipologia")
        
        with tab3:
            st.subheader("Gestione Camere Out of Order")
            
            if not selected_hotel:
                st.warning("Seleziona prima un hotel dalla sidebar")
            else:
                # Mostra camere OOO attive
                ooo_rooms = get_active_ooo_rooms(selected_hotel['id'])
                
                if ooo_rooms:
                    st.markdown("### Camere Fuori Servizio Attive")
                    ooo_df = pd.DataFrame(ooo_rooms)
                    ooo_df = ooo_df[['room_type_name', 'date_from', 'date_to', 'rooms_count', 'reason']]
                    ooo_df.columns = ['Tipologia', 'Da', 'A', 'Numero Camere', 'Motivo']
                    st.dataframe(ooo_df)
                
                # Form per aggiungere nuove camere OOO
                st.markdown("### Dichiara Camere Fuori Servizio")
                
                room_types = get_room_types_by_hotel(selected_hotel['id'])
                if not room_types:
                    st.warning("Nessuna tipologia di camera configurata per questo hotel")
                else:
                    room_type_names = [rt['room_type_name'] for rt in room_types]
                    
                    with st.form("ooo_form"):
                        room_type_name = st.selectbox("Tipologia Camera", room_type_names)
                        col1, col2 = st.columns(2)
                        date_from = col1.date_input("Data Inizio")
                        date_to = col2.date_input("Data Fine")
                        
                        rooms_count = st.number_input("Numero Camere OOO", min_value=1, max_value=100, value=1)
                        reason = st.text_area("Motivo")
                        
                        submit_button = st.form_submit_button("Salva Camere OOO")
                        
                        if submit_button:
                            # Trova id tipologia
                            room_type = next((rt for rt in room_types if rt['room_type_name'] == room_type_name), None)
                            
                            if room_type:
                                # Converti date in formato italiano
                                date_from_str = date_from.strftime('%d/%m/%Y')
                                date_to_str = date_to.strftime('%d/%m/%Y')
                                
                                # Aggiungi camere OOO
                                ooo_data = {
                                    'date_from': date_from_str,
                                    'date_to': date_to_str,
                                    'rooms_count': rooms_count,
                                    'reason': reason
                                }
                                
                                if add_ooo_rooms(selected_hotel['id'], room_type['id'], ooo_data):
                                    st.success(f"{rooms_count} camere {room_type_name} dichiarate OOO con successo!")
                                else:
                                    st.error("Errore nella registrazione delle camere OOO")
                            else:
                                st.error("Tipologia camera non trovata")
    
    elif menu == "Analisi Dati":
        st.header("📊 Analisi Dati")
        
        # Opzioni di analisi
        analysis_type = st.selectbox("Seleziona analisi", 
                                    ["Trend vs SPIT", "Performance per Giorno", "Analisi Pickup", "Forecast Avanzato"])
        
        if analysis_type == "Trend vs SPIT":
            st.subheader("Trend vs SPIT")
            
            if not selected_hotel:
                st.warning("Seleziona un hotel dalla sidebar")
            else:
                period = st.slider("Periodo di analisi (giorni)", 7, 180, 60)
                trend_data = analyze_vs_spit_trend(selected_hotel_name, days=period)
                
                if "error" in trend_data:
                    st.error(trend_data["error"])
                else:
                    # Visualizza metriche
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("Room Nights Totali", f"{trend_data['total_rn']:.0f}")
                    with col2:
                        st.metric("vs SPIT", f"{trend_data['total_vs_spit']:.0f} ({trend_data['vs_spit_pct']:.1f}%)")
                    with col3:
                        st.metric("Room Nights Anno Precedente", f"{trend_data['spit_rn']:.0f}")
                    
                    # Visualizza grafico
                    if st.checkbox("Mostra grafico"):
                        pass  # Qui andrà il codice per il grafico
    
    elif menu == "Stato Inventario":
        st.header("🔍 Stato Inventario")
        
        if not selected_hotel:
            st.warning("Seleziona un hotel dalla sidebar")
        else:
            # Visualizza stato inventario
            st.subheader(f"Inventario {selected_hotel_name}")
            
            # Tab per diverse visualizzazioni
            tab1, tab2 = st.tabs(["Per Tipologia", "Camere OOO"])
            
            with tab1:
                st.markdown("### Stato per Tipologia di Camera")
                
                room_types = get_room_types_by_hotel(selected_hotel['id'])
                if not room_types:
                    st.warning("Nessuna tipologia configurata per questo hotel")
                else:
                    # Prepara dati
                    room_data = []
                    for rt in room_types:
                        # Controllo camere OOO attive per questa tipologia
                        today = datetime.now().strftime('%d/%m/%Y')
                        effective_capacity = get_effective_capacity(
                            selected_hotel['id'], 
                            rt['id'], 
                            today
                        )
                        
                        room_data.append({
                            'Tipologia': rt['room_type_name'],
                            'Capacità Totale': rt['rooms_in_type'],
                            'Camere OOO': rt['rooms_in_type'] - effective_capacity,
                            'Capacità Effettiva': effective_capacity,
                            'Entry Level': '✓' if rt['is_entry_level'] else '',
                            'ADR': f"€ {rt['adr']:.2f}".replace('.', ','),
                            'Upgrade a': rt.get('upgrade_target_type', '')
                        })
                    
                    # Visualizza tabella
                    room_df = pd.DataFrame(room_data)
                    st.dataframe(room_df, use_container_width=True)
                    
                    # Grafico a torta distribuzione camere
                    if st.checkbox("Mostra grafico distribuzione camere"):
                        room_counts = [rt['rooms_in_type'] for rt in room_types]
                        room_names = [rt['room_type_name'] for rt in room_types]
                        
                        fig = px.pie(
                            values=room_counts,
                            names=room_names,
                            title=f"Distribuzione Camere {selected_hotel_name}"
                        )
                        st.plotly_chart(fig, use_container_width=True)
            
            with tab2:
                st.markdown("### Camere Out of Order")
                
                ooo_rooms = get_active_ooo_rooms(selected_hotel['id'])
                if not ooo_rooms:
                    st.info("Nessuna camera fuori servizio attiva")
                else:
                    ooo_df = pd.DataFrame(ooo_rooms)
                    ooo_df = ooo_df[['room_type_name', 'date_from', 'date_to', 'rooms_count', 'reason']]
                    ooo_df.columns = ['Tipologia', 'Da', 'A', 'Numero Camere', 'Motivo']
                    st.dataframe(ooo_df, use_container_width=True)
                    
                    # Opzioni per terminare OOO
                    if st.button("Termina OOO Selezionata"):
                        st.warning("Funzionalità non disponibile in questa versione")

    # Footer con informazioni
    st.sidebar.markdown("---")
    st.sidebar.info("""
    ### Hotel Upgrade Advisor Pro v2.0
    Sviluppato per ottimizzare le decisioni di revenue management alberghiero.

    **Caratteristiche principali:**
    - Integrazione con Business Intelligence
    - Riconoscimento automatico delle tipologie
    - Controllo dell'inventario e overbooking
    - Gestione tipologie entry-level
    - Dashboard completa per il revenue management

    **© 2025 Hotel Solutions**
    """)

# Avvio dell'applicazione
if __name__ == "__main__":
    main()
