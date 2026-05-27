"""
==================================================================
VOI GROUP TOOLKIT  ·  v2
Valutazione e gestione gruppi leisure — VOI Alimini Resort
Metrica primaria: ADR bed  ·  displacement a due livelli (allotment / WEB)
==================================================================
"""

import io
import math
from datetime import date, timedelta

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ------------------------------------------------------------------
# CONFIG / STILE
# ------------------------------------------------------------------
st.set_page_config(page_title="VOI Group Toolkit", page_icon="🏖️", layout="wide")

PRIM, ACCENT = "#0F4C5C", "#E8833A"
VERDE, GIALLO, ROSSO, SAND = "#2E7D32", "#E0911A", "#C62828", "#F6F2EA"
COLOR = {"verde": VERDE, "giallo": GIALLO, "rosso": ROSSO}
ICON = {"verde": "✅", "giallo": "⚠️", "rosso": "⛔"}

SETS = ["Totale", "Individuali (no Alpitour)", "Alpitour individuali"]

st.markdown(f"""
<style>
  .main .block-container {{ padding-top: 1.3rem; max-width: 1260px; }}
  .vt-header {{ background: linear-gradient(120deg,{PRIM} 0%,#14657A 100%);
     color:#fff; padding:22px 28px; border-radius:14px; margin-bottom:18px; }}
  .vt-header h1 {{ margin:0; font-size:1.5rem; font-weight:700; }}
  .vt-header p {{ margin:4px 0 0; opacity:.85; font-size:.9rem; }}
  .vt-card {{ background:{SAND}; border:1px solid #E4DCC9; border-radius:12px;
     padding:14px 18px; margin-bottom:12px; }}
  .vt-check {{ border-radius:10px; padding:11px 16px; margin-bottom:9px;
     color:#fff; font-size:.9rem; }}
  .vt-verdict {{ border-radius:14px; padding:20px 26px; text-align:center;
     color:#fff; margin:8px 0 16px; }}
  .vt-verdict h2 {{ margin:0; font-size:1.45rem; letter-spacing:.5px; }}
  .vt-verdict p {{ margin:6px 0 0; opacity:.9; }}
  .vt-tag {{ display:inline-block; background:{ACCENT}; color:#fff; font-size:.7rem;
     padding:2px 9px; border-radius:20px; margin-left:8px; vertical-align:middle; }}
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------
def eur(x):
    try:
        return f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".") + " €"
    except Exception:
        return "—"


def eur2(x):
    try:
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " €"
    except Exception:
        return "—"


def to_excel_bytes(dfs: dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for sheet, d in dfs.items():
            d.to_excel(w, sheet_name=sheet[:31], index=False)
    return buf.getvalue()


def periodi_default():
    rows = [
        ("Apertura / Bassa", date(2026, 5, 23), date(2026, 6, 6),  3,  64,  52, 200, 35, 10),
        ("Bassa Giugno",     date(2026, 6, 7),  date(2026, 6, 27), 3,  86,  67, 200, 82, 47),
        ("Media Luglio",     date(2026, 6, 28), date(2026, 8, 1),  7, 105,  86, 200, 93, 78),
        ("Alta Agosto",      date(2026, 8, 2),  date(2026, 8, 22), 7, 150, 113, 200, 93, 80),
        ("Spalla Settembre", date(2026, 8, 23), date(2026, 9, 12), 5,  95,  78, 200, 75, 45),
        ("Chiusura",         date(2026, 9, 13), date(2026, 9, 27), 3,  70,  66, 200, 64, 42),
    ]
    cols = ["Periodo", "Data inizio", "Data fine", "Min stay",
            "ADR bed WEB", "ADR bed Alpitour", "Allotment ALPI",
            "Occupancy attesa %", "Utilizzo allotment %"]
    df = pd.DataFrame(rows, columns=cols)
    df["Data inizio"] = pd.to_datetime(df["Data inizio"])
    df["Data fine"] = pd.to_datetime(df["Data fine"])
    return df


def match_periodo(periodi, giorno):
    g = pd.Timestamp(giorno)
    for _, r in periodi.iterrows():
        di, dfi = r["Data inizio"], r["Data fine"]
        if pd.notna(di) and pd.notna(dfi) and pd.Timestamp(di) <= g <= pd.Timestamp(dfi):
            return r
    return None


def analizza_soggiorno(periodi, check_in, check_out):
    """Assegna ogni notte al periodo. Ritorna notti, segmenti pesati, notti orfane."""
    notti = (check_out - check_in).days
    seg, nomatch = {}, 0
    for n in range(notti):
        r = match_periodo(periodi, check_in + timedelta(days=n))
        if r is None:
            nomatch += 1
            continue
        nome = r["Periodo"]
        if nome not in seg:
            seg[nome] = {"notti": 0,
                         "web": float(r["ADR bed WEB"]),
                         "alpi": float(r["ADR bed Alpitour"]),
                         "allot": int(r["Allotment ALPI"]),
                         "occ": float(r["Occupancy attesa %"]),
                         "util": float(r["Utilizzo allotment %"]),
                         "min": int(r["Min stay"])}
        seg[nome]["notti"] += 1
    return notti, seg, nomatch


# --- storico ---
def leggi_file_storico(file):
    df = pd.read_excel(file, 0)
    seg_block = set(df["Segmento"].dropna().unique()) if "Segmento" in df.columns else set()
    df["dt"] = pd.to_datetime(df["Giorno"].astype(str).str.split(" ").str[-1],
                              format="%d/%m/%Y", errors="coerce")
    if "Segmento" in df.columns:
        daily = df[df["Segmento"] == "Total"].dropna(subset=["dt"]).copy()
    else:
        daily = df.dropna(subset=["dt"]).copy()
    daily = daily[(daily["dt"].dt.month >= 4) & (daily["dt"].dt.month <= 10)]
    anno = int(daily["dt"].dt.year.mode().iloc[0]) if len(daily) else None
    return daily, anno, seg_block


def indovina_set(seg_block):
    s = {str(x).upper() for x in seg_block}
    if any("GRUPPI" in x for x in s):
        return "Totale"
    if any("ALPITOUR INDIVIDUALI" in x for x in s) and not any("DIRETTI" in x for x in s):
        return "Alpitour individuali"
    if any("DIRETTI" in x for x in s) or any("WEB PORTALI" in x for x in s):
        return "Individuali (no Alpitour)"
    return "Totale"


def pulisci_storico(df):
    n0 = len(df)
    df = df[df["ADR Bed"].between(25, 260)]
    df = df[df["% Occ."].between(0, 1.05)]
    return df, n0 - len(df)


def righe_periodo(df, di, dfine):
    """Righe storiche che cadono nello stesso intervallo mese/giorno, per ogni anno."""
    if df is None or df.empty:
        return df
    mask = pd.Series(False, index=df.index)
    for anno in sorted(df["dt"].dt.year.unique()):
        try:
            s = pd.Timestamp(year=anno, month=di.month, day=di.day)
            e = pd.Timestamp(year=anno, month=dfine.month, day=dfine.day)
        except ValueError:
            continue
        if e >= s:
            mask |= df["dt"].between(s, e)
    return df[mask]


# ------------------------------------------------------------------
# SESSION STATE
# ------------------------------------------------------------------
if "periodi" not in st.session_state:
    st.session_state.periodi = periodi_default()
if "valutazioni" not in st.session_state:
    st.session_state.valutazioni = []
if "soglie" not in st.session_state:
    st.session_state.soglie = {"low": 0.70, "mid": 0.85, "high": 0.95, "auth": 35000}
if "storico" not in st.session_state:
    st.session_state.storico = {}        # set -> DataFrame concatenato pulito
if "storico_info" not in st.session_state:
    st.session_state.storico_info = ""

# ------------------------------------------------------------------
# HEADER + SIDEBAR
# ------------------------------------------------------------------
st.markdown(f"""
<div class="vt-header">
  <h1>🏖️ VOI Group Toolkit <span class="vt-tag">v2 · ADR bed</span></h1>
  <p>Valutazione richieste gruppi leisure · VOI Alimini Resort · ecosistema Alpitour</p>
</div>""", unsafe_allow_html=True)

pagina = st.sidebar.radio("Sezione",
                          ["🧮 Valutazione gruppo", "📂 Dati storici",
                           "⚙️ Setup periodi", "📋 Riepilogo"],
                          label_visibility="collapsed")
st.sidebar.divider()
st.sidebar.caption("Soglie ADR bed (% della tariffa WEB del periodo)")
s = st.session_state.soglie
s["low"] = st.sidebar.slider("Occupancy < 60%", 0.40, 1.10, s["low"], 0.01)
s["mid"] = st.sidebar.slider("Occupancy 60–80%", 0.40, 1.10, s["mid"], 0.01)
s["high"] = st.sidebar.slider("Occupancy > 80%", 0.40, 1.20, s["high"], 0.01)
s["auth"] = st.sidebar.number_input("Soglia autorizzazione direzione (€)",
                                    0, 1_000_000, int(s["auth"]), 5000)
if st.session_state.storico:
    st.sidebar.success(f"Storico caricato: {', '.join(st.session_state.storico.keys())}")


# ==================================================================
# PAGINA — DATI STORICI
# ==================================================================
if pagina == "📂 Dati storici":
    st.subheader("📂 Caricamento dati storici")
    st.caption("Carica gli export Scrigno dei tre set (consuntivi stagionali). "
               "Il toolkit ne ricava le tariffe WEB e Alpitour, l'occupancy e "
               "l'utilizzo dell'allotment per ciascun periodo del Setup.")

    files = st.file_uploader("Trascina qui i file .xlsx (anche tutti insieme)",
                             type=["xlsx"], accept_multiple_files=True)

    if files:
        meta = []
        cache = {}
        for f in files:
            daily, anno, block = leggi_file_storico(f)
            cache[f.name] = daily
            rng = (f"{daily['dt'].min().date()} → {daily['dt'].max().date()}"
                   if len(daily) else "—")
            meta.append({"File": f.name, "Anno": anno or 0,
                         "Periodo dati": rng, "Set": indovina_set(block)})
        meta_df = pd.DataFrame(meta)

        st.markdown("##### Assegna ogni file al suo set")
        st.caption("Il set è stato indovinato dai segmenti presenti; correggilo se serve.")
        edited = st.data_editor(
            meta_df, hide_index=True, use_container_width=True,
            disabled=["File", "Anno", "Periodo dati"],
            column_config={"Set": st.column_config.SelectboxColumn("Set", options=SETS)})

        if st.button("⚙️ Elabora e applica al Setup periodi", type="primary",
                     use_container_width=True):
            storico, glitch_tot = {}, 0
            for set_name in SETS:
                fs = edited[edited["Set"] == set_name]["File"].tolist()
                if not fs:
                    continue
                merged = pd.concat([cache[fn] for fn in fs], ignore_index=True)
                merged, glitch = pulisci_storico(merged)
                glitch_tot += glitch
                storico[set_name] = merged
            st.session_state.storico = storico

            # --- aggrega per periodo ---
            per = st.session_state.periodi.copy()
            tot = storico.get("Totale")
            ind = storico.get("Individuali (no Alpitour)")
            alp = storico.get("Alpitour individuali")
            applicati = 0
            for idx, r in per.iterrows():
                di, dfi = r["Data inizio"], r["Data fine"]
                if pd.isna(di) or pd.isna(dfi):
                    continue
                di, dfi = di.date(), dfi.date()
                if tot is not None:
                    rp = righe_periodo(tot, di, dfi)
                    if rp is not None and len(rp):
                        per.at[idx, "Occupancy attesa %"] = round(rp["% Occ."].mean() * 100, 1)
                if ind is not None:
                    rp = righe_periodo(ind, di, dfi)
                    if rp is not None and len(rp):
                        per.at[idx, "ADR bed WEB"] = round(rp["ADR Bed"].median(), 1)
                if alp is not None:
                    rp = righe_periodo(alp, di, dfi)
                    if rp is not None and len(rp):
                        per.at[idx, "ADR bed Alpitour"] = round(rp["ADR Bed"].median(), 1)
                        allot = r["Allotment ALPI"] or 200
                        per.at[idx, "Utilizzo allotment %"] = round(
                            rp["Room nights"].mean() / allot * 100, 1)
                applicati += 1
            st.session_state.periodi = per
            st.session_state.storico_info = (
                f"{len(files)} file · set: {', '.join(storico.keys())} · "
                f"{glitch_tot} righe anomale escluse")
            st.success(f"✅ Storico elaborato e applicato a {applicati} periodi. "
                       f"{glitch_tot} righe anomale (ADR bed fuori 25–260 €) escluse. "
                       f"Vai su «Setup periodi» per verificare.")

    if st.session_state.storico:
        st.divider()
        st.markdown("##### Quadro storico per set")
        rows = []
        for k, d in st.session_state.storico.items():
            rows.append({"Set": k, "Righe-giorno": len(d),
                         "Anni": ", ".join(map(str, sorted(d["dt"].dt.year.unique()))),
                         "Occ. media": f"{d['% Occ.'].mean()*100:.1f}%",
                         "ADR bed mediana": eur2(d["ADR Bed"].median())})
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)


# ==================================================================
# PAGINA — SETUP PERIODI
# ==================================================================
elif pagina == "⚙️ Setup periodi":
    st.subheader("⚙️ Setup periodi tariffari")
    st.caption("Anagrafica periodi. Le tariffe sono **ADR bed per pax/notte**. "
               "Se hai caricato lo storico, i valori sono pre-compilati dai consuntivi "
               "(restano modificabili).")
    if st.session_state.storico_info:
        st.info(f"📂 Pre-compilato da storico — {st.session_state.storico_info}")

    cfg = {
        "Periodo": st.column_config.TextColumn("Periodo", width="medium"),
        "Data inizio": st.column_config.DateColumn("Inizio", format="DD/MM/YYYY"),
        "Data fine": st.column_config.DateColumn("Fine", format="DD/MM/YYYY"),
        "Min stay": st.column_config.NumberColumn("MLOS", min_value=1, max_value=21, step=1),
        "ADR bed WEB": st.column_config.NumberColumn("ADR bed WEB", format="%.1f €",
                       help="Tariffa individuale dinamica (CRS / Vertical Booking / Blastness + OTA + diretto)."),
        "ADR bed Alpitour": st.column_config.NumberColumn("ADR bed Alpitour", format="%.1f €",
                            help="ADR bed degli individuali Alpitour in allotment."),
        "Allotment ALPI": st.column_config.NumberColumn("Allotment ALPI", min_value=0, step=1),
        "Occupancy attesa %": st.column_config.NumberColumn("Occ. attesa %", format="%.1f"),
        "Utilizzo allotment %": st.column_config.NumberColumn("Utilizzo allot. %", format="%.1f",
                                help="Quota dell'allotment tipicamente riempita dagli individuali Alpitour."),
    }
    edited = st.data_editor(st.session_state.periodi, column_config=cfg,
                            num_rows="dynamic", use_container_width=True, hide_index=True)
    st.session_state.periodi = edited

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("⬇️ Esporta periodi", to_excel_bytes({"Periodi": edited}),
                           "voi_periodi.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with c2:
        up = st.file_uploader("⬆️ Importa periodi", type=["xlsx"], label_visibility="collapsed")
        if up is not None:
            try:
                imp = pd.read_excel(up)
                imp["Data inizio"] = pd.to_datetime(imp["Data inizio"])
                imp["Data fine"] = pd.to_datetime(imp["Data fine"])
                st.session_state.periodi = imp
                st.success("Periodi importati.")
                st.rerun()
            except Exception as e:
                st.error(f"Import non riuscito: {e}")
    with c3:
        if st.button("↺ Ripristina periodi demo", use_container_width=True):
            st.session_state.periodi = periodi_default()
            st.session_state.storico_info = ""
            st.rerun()


# ==================================================================
# PAGINA — RIEPILOGO
# ==================================================================
elif pagina == "📋 Riepilogo":
    st.subheader("📋 Riepilogo valutazioni")
    if not st.session_state.valutazioni:
        st.info("Nessuna valutazione salvata in questa sessione.")
    else:
        df = pd.DataFrame(st.session_state.valutazioni)
        st.dataframe(df, use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Esporta riepilogo", to_excel_bytes({"Valutazioni": df}),
                               "voi_valutazioni_gruppi.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        with c2:
            if st.button("🗑️ Svuota riepilogo", use_container_width=True):
                st.session_state.valutazioni = []
                st.rerun()


# ==================================================================
# PAGINA — VALUTAZIONE GRUPPO
# ==================================================================
else:
    periodi = st.session_state.periodi
    st.subheader("🧮 Valutazione richiesta gruppo")
    if not st.session_state.storico:
        st.caption("💡 Suggerimento: carica i consuntivi in «Dati storici» per basare la "
                   "valutazione su tariffe e occupancy reali invece che sui valori demo.")

    c1, c2, c3 = st.columns(3)
    with c1:
        nome_gruppo = st.text_input("Nome / riferimento gruppo", "Gruppo senza nome")
        check_in = st.date_input("Check-in", date(2026, 7, 11), format="DD/MM/YYYY")
        check_out = st.date_input("Check-out", date(2026, 7, 14), format="DD/MM/YYYY")
    with c2:
        camere = st.number_input("Camere richieste", 1, 500, 30, 1)
        pax_cam = st.number_input("Pax / camera", 1.0, 4.0, 2.25, 0.05,
                                  help="Default gruppi leisure = 2,25.")
        meal = st.selectbox("Meal plan (riferimento)", ["BB", "HB", "FB"], index=1)
    with c3:
        tariffa = st.number_input("Tariffa proposta — ADR bed (€/pax/notte)",
                                  0.0, 1000.0, 95.0, 1.0)
        ancillare = st.number_input("Ricavo ancillare extra (€/pax/notte)",
                                    0.0, 500.0, 0.0, 1.0)
        allot_residuo = st.number_input("Allotment ALPI residuo (da Scrigno)",
                                        0, 500, 20, 1,
                                        help="Camere ancora libere nell'allotment Alpitour "
                                             "per le date. Verifica manualmente su Scrigno.")

    # --- pre-analisi periodi (per default override) ---
    notti, seg, nomatch = analizza_soggiorno(periodi, check_in, check_out) \
        if check_out > check_in else (0, {}, 0)
    if seg:
        nv = sum(v["notti"] for v in seg.values())
        occ_def = sum(v["notti"] * v["occ"] for v in seg.values()) / nv
        util_def = sum(v["notti"] * v["util"] for v in seg.values()) / nv
    else:
        occ_def, util_def = 75.0, 50.0

    with st.expander("⚙️ Parametri avanzati (default da periodi storici)"):
        a1, a2, a3 = st.columns(3)
        with a1:
            occupancy = st.number_input("Occupancy attesa (%)", 0.0, 100.0,
                                        round(float(occ_def), 1), 1.0)
        with a2:
            util_allot = st.number_input("Utilizzo allotment Alpitour (%)", 0.0, 100.0,
                                         round(float(util_def), 1), 1.0,
                                         help="Probabilità che lo slot di allotment venga "
                                              "comunque riempito dagli individuali Alpitour.")
        with a3:
            pickup_web = st.number_input("Pick-up casa / WEB (%)", 0.0, 100.0,
                                         round(float(occ_def), 1), 1.0,
                                         help="Probabilità che le camere oltre allotment "
                                              "si vendano comunque a tariffa WEB.")

    valuta = st.button("▶️  Valuta richiesta", type="primary", use_container_width=True)

    # ---------- ELABORAZIONE ----------
    if valuta:
        if check_out <= check_in:
            st.error("Il check-out deve essere successivo al check-in.")
            st.stop()
        if not seg:
            st.error("Le date non rientrano in alcun periodo configurato (vedi «Setup periodi»).")
            st.stop()
        if nomatch:
            st.warning(f"⚠️ {nomatch} notti su {notti} fuori da ogni periodo: escluse dal calcolo.")

        nv = sum(v["notti"] for v in seg.values())
        web_w = sum(v["notti"] * v["web"] for v in seg.values()) / nv
        alpi_w = sum(v["notti"] * v["alpi"] for v in seg.values()) / nv
        min_eff = max(v["min"] for v in seg.values())

        # volumi gruppo
        pax = camere * pax_cam
        bed_nights = pax * nv
        rev_camere = bed_nights * tariffa
        rev_anc = bed_nights * ancillare
        rev_totale = rev_camere + rev_anc
        adr_room = tariffa * pax_cam

        # displacement a due livelli
        camere_allot = min(camere, allot_residuo)
        camere_over = max(0, camere - allot_residuo)
        rev_alt_allot = camere_allot * pax_cam * alpi_w * nv * (util_allot / 100)
        rev_alt_web = camere_over * pax_cam * web_w * nv * (pickup_web / 100)
        rev_alt = rev_alt_allot + rev_alt_web
        displacement = rev_totale - rev_alt

        # soglia ADR bed
        pct = s["low"] if occupancy < 60 else s["mid"] if occupancy < 80 else s["high"]
        soglia_bed = web_w * pct

        # controproposta
        denom = camere * pax_cam * nv
        tariffa_be = (rev_alt - rev_anc) / denom if denom else 0
        controproposta = math.ceil(max(tariffa_be, soglia_bed))

        # ===== CHECK 1 — ALLOTMENT =====
        if camere_over == 0:
            c1r = ("verde", "Allotment ALPI",
                   f"Le {camere} camere rientrano nell'allotment residuo ({allot_residuo}). "
                   f"Nessuna erosione dell'inventario WEB.")
        elif camere_over <= max(2, 0.15 * camere):
            c1r = ("giallo", "Allotment ALPI",
                   f"{camere_over} camere oltre allotment ({allot_residuo} residue): "
                   f"erosione contenuta dell'inventario WEB, valutate a tariffa dinamica.")
        else:
            c1r = ("rosso", "Allotment ALPI",
                   f"{camere_over} camere oltre allotment ({allot_residuo} residue): "
                   f"erosione significativa dell'inventario WEB ad alto valore.")

        # ===== CHECK 2 — MIN STAY =====
        if notti >= min_eff:
            c2r = ("verde", "Minimum stay",
                   f"Soggiorno di {notti} notti ≥ MLOS del periodo ({min_eff}).")
        elif notti >= min_eff - 1:
            c2r = ("giallo", "Minimum stay",
                   f"{notti} notti contro MLOS {min_eff}: deroga lieve, da autorizzare.")
        else:
            c2r = ("rosso", "Minimum stay",
                   f"{notti} notti sotto il MLOS di {min_eff}: deroga importante.")

        # ===== CHECK 3 — ADR BED =====
        gap = tariffa - soglia_bed
        if tariffa >= soglia_bed:
            c3r = ("verde", "ADR bed vs soglia",
                   f"Tariffa {eur2(tariffa)} ≥ soglia {eur2(soglia_bed)} "
                   f"({pct*100:.0f}% della WEB con occupancy {occupancy:.0f}%).")
        elif tariffa >= soglia_bed * 0.92:
            c3r = ("giallo", "ADR bed vs soglia",
                   f"Tariffa {eur2(tariffa)} di poco sotto la soglia {eur2(soglia_bed)} "
                   f"(gap {eur2(gap)}/pax).")
        else:
            c3r = ("rosso", "ADR bed vs soglia",
                   f"Tariffa {eur2(tariffa)} sotto la soglia {eur2(soglia_bed)} "
                   f"(gap {eur2(gap)}/pax).")

        # ===== CHECK 4 — DISPLACEMENT =====
        if displacement > 0:
            c4r = ("verde", "Displacement netto",
                   f"Il gruppo genera {eur(displacement)} di valore incrementale "
                   f"rispetto alla vendita alternativa attesa.")
        elif displacement >= -0.05 * rev_alt:
            c4r = ("giallo", "Displacement netto",
                   f"Displacement marginalmente negativo ({eur(displacement)}): "
                   f"valore quasi equivalente all'alternativa.")
        else:
            c4r = ("rosso", "Displacement netto",
                   f"Il gruppo distrugge {eur(abs(displacement))} di valore "
                   f"rispetto alla vendita alternativa attesa.")

        checks = [c1r, c2r, c3r, c4r]
        stati = [c[0] for c in checks]
        if "rosso" in stati:
            verdetto, vcol = "RIFIUTARE O RINEGOZIARE", "rosso"
        elif "giallo" in stati:
            verdetto, vcol = "VALUTARE — CONTROPROPOSTA CONSIGLIATA", "giallo"
        else:
            verdetto, vcol = "ACCETTARE", "verde"

        # ---------- OUTPUT ----------
        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Pax totali", f"{pax:.0f}", help=f"{camere} camere × {pax_cam} pax/cam")
        m2.metric("Bed nights", f"{bed_nights:.0f}")
        m3.metric("ADR bed gruppo", eur2(tariffa))
        m4.metric("ADR room gruppo", eur2(adr_room))

        m5, m6, m7, m8 = st.columns(4)
        m5.metric("Valore totale gruppo", eur(rev_totale))
        m6.metric("Alternativa attesa", eur(rev_alt))
        m7.metric("Displacement netto", eur(displacement),
                  delta=f"{displacement/rev_alt*100:+.1f}%" if rev_alt else None)
        m8.metric("Camere oltre allotment", f"{camere_over}")

        st.markdown(f"""
        <div class="vt-verdict" style="background:{COLOR[vcol]}">
          <h2>{ICON[vcol]}  {verdetto}</h2>
          <p>{nome_gruppo} · {check_in.strftime('%d/%m/%Y')} → {check_out.strftime('%d/%m/%Y')}
             · {notti} notti · {camere} camere · meal {meal}</p>
        </div>""", unsafe_allow_html=True)

        cL, cR = st.columns([3, 2])
        with cL:
            st.markdown("##### Esito controlli")
            for stato, titolo, dett in checks:
                st.markdown(f"""<div class="vt-check" style="background:{COLOR[stato]}">
                  <b>{ICON[stato]} {titolo}</b><br>{dett}</div>""", unsafe_allow_html=True)
            if rev_totale > s["auth"]:
                st.warning(f"📨 Valore totale {eur(rev_totale)} oltre la soglia di "
                           f"{eur(s['auth'])}: **richiede autorizzazione direzione**.")
        with cR:
            st.markdown("##### Controproposta")
            st.markdown(f"""<div class="vt-card">
              <p style="margin:0 0 4px;font-size:.84rem;color:#555">Break-even bed (displ. = 0)</p>
              <p style="margin:0;font-size:1.3rem;font-weight:700;color:{PRIM}">{eur2(tariffa_be)}/pax</p>
              <hr style="margin:9px 0;border-color:#E4DCC9">
              <p style="margin:0 0 4px;font-size:.84rem;color:#555">Soglia ADR bed (occ {occupancy:.0f}%)</p>
              <p style="margin:0;font-size:1.3rem;font-weight:700;color:{PRIM}">{eur2(soglia_bed)}/pax</p>
              <hr style="margin:9px 0;border-color:#E4DCC9">
              <p style="margin:0 0 4px;font-size:.84rem;color:#555">✅ Tariffa bed da richiedere</p>
              <p style="margin:0;font-size:1.55rem;font-weight:800;color:{ACCENT}">{eur(controproposta)}/pax</p>
              <p style="margin:3px 0 0;font-size:.76rem;color:#777">≈ {eur(controproposta*pax_cam)}/camera</p>
            </div>""", unsafe_allow_html=True)

        g1, g2 = st.columns(2)
        with g1:
            fig = go.Figure()
            fig.add_bar(name="Ricavo camere", x=["Gruppo"], y=[rev_camere], marker_color=PRIM)
            fig.add_bar(name="Ricavo ancillare", x=["Gruppo"], y=[rev_anc], marker_color=ACCENT)
            fig.add_bar(name="Alt. — slot allotment", x=["Alternativa"],
                        y=[rev_alt_allot], marker_color="#7E9AA3")
            fig.add_bar(name="Alt. — inventario WEB", x=["Alternativa"],
                        y=[rev_alt_web], marker_color="#B9C5C9")
            fig.update_layout(barmode="stack", title="Valore gruppo vs alternativa attesa",
                              height=350, margin=dict(t=46, b=10, l=10, r=10),
                              legend=dict(orientation="h", y=-0.2))
            st.plotly_chart(fig, use_container_width=True)
        with g2:
            fig2 = go.Figure(go.Bar(
                x=["Proposta", "Soglia", "WEB", "Alpitour"],
                y=[tariffa, soglia_bed, web_w, alpi_w],
                marker_color=[ACCENT, GIALLO, PRIM, "#7E9AA3"],
                text=[eur2(v) for v in [tariffa, soglia_bed, web_w, alpi_w]],
                textposition="outside"))
            fig2.update_layout(title="ADR bed — confronto (€/pax/notte)",
                               height=350, margin=dict(t=46, b=10, l=10, r=10))
            st.plotly_chart(fig2, use_container_width=True)

        with st.expander("🔎 Dettaglio periodi del soggiorno"):
            det = pd.DataFrame([
                {"Periodo": k, "Notti": v["notti"], "ADR WEB": eur2(v["web"]),
                 "ADR Alpitour": eur2(v["alpi"]), "Occ. attesa": f"{v['occ']:.0f}%",
                 "Utilizzo allot.": f"{v['util']:.0f}%", "MLOS": v["min"]}
                for k, v in seg.items()])
            st.dataframe(det, hide_index=True, use_container_width=True)
            st.caption(f"Valori pesati sul soggiorno → ADR WEB {eur2(web_w)} · "
                       f"ADR Alpitour {eur2(alpi_w)} · MLOS effettivo {min_eff}.")

        record = {"Gruppo": nome_gruppo, "Check-in": check_in.strftime("%d/%m/%Y"),
                  "Check-out": check_out.strftime("%d/%m/%Y"), "Notti": notti,
                  "Camere": camere, "Pax": round(pax), "Meal": meal,
                  "ADR bed": round(tariffa, 2), "Valore totale": round(rev_totale),
                  "Displacement": round(displacement),
                  "Controproposta bed": controproposta, "Verdetto": verdetto}
        if st.button("💾 Salva valutazione nel riepilogo", use_container_width=True):
            st.session_state.valutazioni.append(record)
            st.success("Valutazione salvata.")

    with st.expander("ℹ️ Metodologia di calcolo"):
        st.markdown("""
**Displacement a due livelli.** Le camere del gruppo entro l'allotment ALPI residuo e quelle
che lo sforano hanno un costo-opportunità diverso:

- **Entro allotment** → l'alternativa è che Alpitour riempia quello slot con un individuale TO.
  Valore = *ADR bed Alpitour × utilizzo tipico dell'allotment* (più l'allotment si satura di
  norma, più il displacement è reale).
- **Oltre allotment** → si erode inventario di casa, che si venderebbe alla **tariffa WEB**
  (dinamica CRS / Vertical Booking / Blastness + OTA + diretto). Valore = *ADR bed WEB ×
  probabilità di pick-up*.

Il **displacement netto** è il valore totale del gruppo (camere + ancillare) meno la somma
delle due alternative attese.

**Soglia ADR bed** = percentuale della tariffa WEB del periodo, crescente con l'occupancy.
**Controproposta** = la più alta tra la tariffa di break-even (displacement nullo) e la soglia.

Tariffe, occupancy e utilizzo dell'allotment sono pre-compilati dai consuntivi caricati in
«Dati storici» (mediana per le ADR, robusta agli errori di export), e restano modificabili.
""")
