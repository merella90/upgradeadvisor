"""
==================================================================
VOI GROUP TOOLKIT — Valutazione e gestione gruppi leisure
VOI Alimini Resort  /  ecosistema Alpitour
Metrica primaria: ADR bed  (ADR room derivato da pax/cam)
==================================================================
"""

import io
import math
from datetime import date, timedelta

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ------------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------------
st.set_page_config(page_title="VOI Group Toolkit", page_icon="🏖️", layout="wide")

PRIM = "#0F4C5C"      # teal mediterraneo
ACCENT = "#E8833A"    # terracotta Puglia
VERDE = "#2E7D32"
GIALLO = "#E0911A"
ROSSO = "#C62828"
SAND = "#F6F2EA"

MEAL_PLANS = ["BB", "HB", "FB"]
MEAL_LABEL = {"BB": "Pernottamento + colazione", "HB": "Mezza pensione", "FB": "Pensione completa"}

st.markdown(f"""
<style>
  .main .block-container {{ padding-top: 1.4rem; max-width: 1250px; }}
  .vt-header {{
     background: linear-gradient(120deg, {PRIM} 0%, #14657A 100%);
     color: #fff; padding: 22px 28px; border-radius: 14px; margin-bottom: 18px;
  }}
  .vt-header h1 {{ margin: 0; font-size: 1.55rem; font-weight: 700; }}
  .vt-header p  {{ margin: 4px 0 0; opacity: .85; font-size: .92rem; }}
  .vt-card {{
     background: {SAND}; border: 1px solid #E4DCC9; border-radius: 12px;
     padding: 14px 18px; margin-bottom: 12px;
  }}
  .vt-check {{
     border-radius: 10px; padding: 12px 16px; margin-bottom: 10px;
     color: #fff; font-size: .92rem;
  }}
  .vt-check b {{ font-size: 1rem; }}
  .vt-verdict {{
     border-radius: 14px; padding: 20px 26px; text-align: center;
     color: #fff; margin: 8px 0 16px;
  }}
  .vt-verdict h2 {{ margin: 0; font-size: 1.5rem; letter-spacing: .5px; }}
  .vt-verdict p  {{ margin: 6px 0 0; opacity: .9; }}
  .vt-tag {{
     display:inline-block; background:{ACCENT}; color:#fff; font-size:.72rem;
     padding:2px 9px; border-radius:20px; margin-left:8px; vertical-align:middle;
  }}
</style>
""", unsafe_allow_html=True)

COLOR = {"verde": VERDE, "giallo": GIALLO, "rosso": ROSSO}
ICON = {"verde": "✅", "giallo": "⚠️", "rosso": "⛔"}


# ------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------
def eur(x):
    """Formato valuta italiano: 1.234 €"""
    try:
        return f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".") + " €"
    except Exception:
        return "—"


def eur2(x):
    try:
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " €"
    except Exception:
        return "—"


def periodi_default():
    rows = [
        ("Apertura / Bassa",  date(2026, 5, 23), date(2026, 6, 6),  3,  58,  72,  86,  42,  52,  62, 25),
        ("Bassa Giugno",      date(2026, 6, 7),  date(2026, 6, 27), 3,  68,  84, 100,  50,  62,  74, 30),
        ("Media Luglio",      date(2026, 6, 28), date(2026, 8, 1),  7,  92, 112, 132,  68,  82,  96, 35),
        ("Alta Agosto",       date(2026, 8, 2),  date(2026, 8, 22), 7, 138, 162, 188, 104, 122, 142, 20),
        ("Spalla Settembre",  date(2026, 8, 23), date(2026, 9, 12), 5,  84, 102, 120,  62,  76,  90, 30),
        ("Chiusura",          date(2026, 9, 13), date(2026, 9, 27), 3,  60,  74,  88,  44,  54,  64, 25),
    ]
    cols = ["Periodo", "Data inizio", "Data fine", "Min stay",
            "ADR bed FIT BB", "ADR bed FIT HB", "ADR bed FIT FB",
            "ADR bed TO BB", "ADR bed TO HB", "ADR bed TO FB", "Allotment ALPI"]
    df = pd.DataFrame(rows, columns=cols)
    df["Data inizio"] = pd.to_datetime(df["Data inizio"])
    df["Data fine"] = pd.to_datetime(df["Data fine"])
    return df


def match_periodo(periodi, giorno):
    g = pd.Timestamp(giorno)
    for _, r in periodi.iterrows():
        di, df_ = r["Data inizio"], r["Data fine"]
        if pd.notna(di) and pd.notna(df_) and pd.Timestamp(di) <= g <= pd.Timestamp(df_):
            return r
    return None


def analizza_soggiorno(periodi, check_in, check_out, meal):
    """Assegna ogni notte al suo periodo. Ritorna notti, segmenti, notti senza periodo."""
    notti = (check_out - check_in).days
    seg = {}
    nomatch = 0
    for n in range(notti):
        g = check_in + timedelta(days=n)
        r = match_periodo(periodi, g)
        if r is None:
            nomatch += 1
            continue
        nome = r["Periodo"]
        if nome not in seg:
            seg[nome] = {"notti": 0,
                         "fit": float(r[f"ADR bed FIT {meal}"]),
                         "to": float(r[f"ADR bed TO {meal}"]),
                         "min": int(r["Min stay"]),
                         "allot": int(r["Allotment ALPI"])}
        seg[nome]["notti"] += 1
    return notti, seg, nomatch


def pct_soglia(occ, low, mid, high):
    if occ < 60:
        return low
    elif occ < 80:
        return mid
    return high


def to_excel_bytes(dfs: dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for sheet, d in dfs.items():
            d.to_excel(w, sheet_name=sheet[:31], index=False)
    return buf.getvalue()


# ------------------------------------------------------------------
# SESSION STATE
# ------------------------------------------------------------------
if "periodi" not in st.session_state:
    st.session_state.periodi = periodi_default()
if "valutazioni" not in st.session_state:
    st.session_state.valutazioni = []
if "soglie" not in st.session_state:
    st.session_state.soglie = {"low": 0.70, "mid": 0.85, "high": 0.95, "auth": 35000}


# ------------------------------------------------------------------
# HEADER + NAV
# ------------------------------------------------------------------
st.markdown(f"""
<div class="vt-header">
  <h1>🏖️ VOI Group Toolkit <span class="vt-tag">ADR bed</span></h1>
  <p>Valutazione richieste gruppi leisure · VOI Alimini Resort · ecosistema Alpitour</p>
</div>
""", unsafe_allow_html=True)

pagina = st.sidebar.radio("Sezione",
                          ["🧮 Valutazione gruppo", "⚙️ Setup periodi", "📋 Riepilogo valutazioni"],
                          label_visibility="collapsed")
st.sidebar.divider()
st.sidebar.caption("Soglie ADR bed (% della tariffa FIT bed del periodo)")
s = st.session_state.soglie
s["low"] = st.sidebar.slider("Occupancy < 60%", 0.40, 1.10, s["low"], 0.01)
s["mid"] = st.sidebar.slider("Occupancy 60–80%", 0.40, 1.10, s["mid"], 0.01)
s["high"] = st.sidebar.slider("Occupancy > 80%", 0.40, 1.20, s["high"], 0.01)
s["auth"] = st.sidebar.number_input("Soglia autorizzazione direzione (€)",
                                    0, 1_000_000, int(s["auth"]), 5000)


# ==================================================================
# PAGINA — SETUP PERIODI
# ==================================================================
if pagina == "⚙️ Setup periodi":
    st.subheader("⚙️ Setup periodi tariffari")
    st.caption("Anagrafica periodi della stagione. Le tariffe sono **ADR bed per pax/notte** "
               "per ciascun meal plan. La valutazione gruppo legge automaticamente questi dati.")

    cfg = {
        "Periodo": st.column_config.TextColumn("Periodo", width="medium"),
        "Data inizio": st.column_config.DateColumn("Inizio", format="DD/MM/YYYY"),
        "Data fine": st.column_config.DateColumn("Fine", format="DD/MM/YYYY"),
        "Min stay": st.column_config.NumberColumn("MLOS", min_value=1, max_value=21, step=1),
        "ADR bed FIT BB": st.column_config.NumberColumn("FIT BB", format="%.0f €"),
        "ADR bed FIT HB": st.column_config.NumberColumn("FIT HB", format="%.0f €"),
        "ADR bed FIT FB": st.column_config.NumberColumn("FIT FB", format="%.0f €"),
        "ADR bed TO BB": st.column_config.NumberColumn("TO BB", format="%.0f €"),
        "ADR bed TO HB": st.column_config.NumberColumn("TO HB", format="%.0f €"),
        "ADR bed TO FB": st.column_config.NumberColumn("TO FB", format="%.0f €"),
        "Allotment ALPI": st.column_config.NumberColumn("Allot. ALPI", min_value=0, step=1),
    }
    edited = st.data_editor(st.session_state.periodi, column_config=cfg,
                            num_rows="dynamic", use_container_width=True, hide_index=True)
    st.session_state.periodi = edited

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("⬇️ Esporta periodi (Excel)",
                           to_excel_bytes({"Periodi": edited}),
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
            st.rerun()

    st.info("**FIT** = tariffa bed di riferimento per la vendita diretta/individuale. "
            "**TO** = tariffa bed netta contrattualizzata Alpitour. "
            "Il displacement usa la FIT per le camere oltre allotment e la TO per quelle entro allotment.")


# ==================================================================
# PAGINA — RIEPILOGO
# ==================================================================
elif pagina == "📋 Riepilogo valutazioni":
    st.subheader("📋 Riepilogo valutazioni")
    if not st.session_state.valutazioni:
        st.info("Nessuna valutazione salvata in questa sessione. "
                "Vai su «Valutazione gruppo» e usa **Salva valutazione**.")
    else:
        df = pd.DataFrame(st.session_state.valutazioni)
        st.dataframe(df, use_container_width=True, hide_index=True)
        c1, c2 = st.columns([1, 1])
        with c1:
            st.download_button("⬇️ Esporta riepilogo (Excel)",
                               to_excel_bytes({"Valutazioni": df}),
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

    # ---------- INPUT ----------
    with st.container():
        c1, c2, c3 = st.columns(3)
        with c1:
            nome_gruppo = st.text_input("Nome / riferimento gruppo", "Gruppo senza nome")
            check_in = st.date_input("Check-in", date(2026, 7, 11), format="DD/MM/YYYY")
            check_out = st.date_input("Check-out", date(2026, 7, 14), format="DD/MM/YYYY")
        with c2:
            camere = st.number_input("Camere richieste", 1, 500, 30, 1)
            pax_cam = st.number_input("Pax / camera", 1.0, 4.0, 2.25, 0.05,
                                      help="Default gruppi leisure = 2,25. Sovrascrivi con il dato reale "
                                           "della richiesta (colonna pax/cam).")
            meal = st.selectbox("Meal plan", MEAL_PLANS,
                                index=1, format_func=lambda m: f"{m} — {MEAL_LABEL[m]}")
        with c3:
            tariffa = st.number_input("Tariffa proposta — ADR bed (€/pax/notte)",
                                      0.0, 1000.0, 95.0, 1.0)
            ancillare = st.number_input("Ricavo ancillare extra (€/pax/notte)",
                                        0.0, 500.0, 0.0, 1.0,
                                        help="F&B extra, escursioni, spa ecc. non inclusi nel meal plan.")

    st.markdown('<div class="vt-card">', unsafe_allow_html=True)
    c4, c5, c6 = st.columns(3)
    with c4:
        allot_residuo = st.number_input("Allotment ALPI residuo sulle date (da Scrigno)",
                                        0, 500, 20, 1,
                                        help="Camere ancora disponibili nell'allotment Alpitour per il "
                                             "periodo. Verifica manualmente su Scrigno. Se il soggiorno "
                                             "scavalca più periodi, inserisci il valore del periodo più critico.")
    with c5:
        occupancy = st.slider("Occupancy attesa nel periodo (%)", 0, 100, 75, 1)
    with c6:
        pickup = st.slider("Probabilità pick-up alternativo (%)", 0, 100, 75, 1,
                           help="Probabilità che le camere vengano comunque vendute se NON si accetta "
                                "il gruppo. Default ≈ occupancy attesa; correggi in base alla forza "
                                "della domanda OTB.")
    st.markdown('</div>', unsafe_allow_html=True)

    valuta = st.button("▶️  Valuta richiesta", type="primary", use_container_width=True)

    # ---------- ELABORAZIONE ----------
    if valuta:
        if check_out <= check_in:
            st.error("Il check-out deve essere successivo al check-in.")
            st.stop()
        if periodi.empty:
            st.error("Nessun periodo configurato. Vai su «Setup periodi».")
            st.stop()

        notti, seg, nomatch = analizza_soggiorno(periodi, check_in, check_out, meal)
        if not seg:
            st.error("Le date selezionate non rientrano in nessun periodo configurato.")
            st.stop()
        if nomatch > 0:
            st.warning(f"⚠️ {nomatch} notti su {notti} non rientrano in alcun periodo configurato "
                       f"e sono escluse dal calcolo.")

        notti_valide = sum(v["notti"] for v in seg.values())
        fit_w = sum(v["notti"] * v["fit"] for v in seg.values()) / notti_valide
        to_w = sum(v["notti"] * v["to"] for v in seg.values()) / notti_valide
        min_stay_eff = max(v["min"] for v in seg.values())

        # --- volumi gruppo ---
        pax = camere * pax_cam
        bed_nights = pax * notti_valide
        room_nights = camere * notti_valide
        rev_camere = bed_nights * tariffa
        rev_anc = bed_nights * ancillare
        rev_totale = rev_camere + rev_anc
        adr_room_gruppo = tariffa * pax_cam

        # --- alternativa / displacement ---
        camere_allot = min(camere, allot_residuo)
        camere_over = max(0, camere - allot_residuo)
        rev_alt_lordo = (camere_allot * pax_cam * to_w +
                         camere_over * pax_cam * fit_w) * notti_valide
        rev_alt_atteso = rev_alt_lordo * (pickup / 100)
        displacement = rev_totale - rev_alt_atteso

        # --- soglia ADR bed ---
        pct = pct_soglia(occupancy, s["low"], s["mid"], s["high"])
        soglia_bed = fit_w * pct

        # --- controproposta ---
        denom = camere * pax_cam * notti_valide
        tariffa_be = (rev_alt_atteso - rev_anc) / denom if denom else 0
        tariffa_target = max(tariffa_be, soglia_bed)
        controproposta = math.ceil(tariffa_target)

        # =================== CHECK 1 — ALLOTMENT ===================
        if camere_over == 0:
            c1r = ("verde", "Allotment ALPI",
                   f"Le {camere} camere rientrano nell'allotment residuo ({allot_residuo}). "
                   f"Nessuna erosione dell'inventario FIT.")
        elif camere_over <= max(2, 0.15 * camere):
            c1r = ("giallo", "Allotment ALPI",
                   f"{camere_over} camere oltre allotment ({allot_residuo} residue): "
                   f"erosione contenuta dell'inventario diretto.")
        else:
            c1r = ("rosso", "Allotment ALPI",
                   f"{camere_over} camere oltre allotment ({allot_residuo} residue): "
                   f"erosione significativa dell'inventario FIT diretto.")

        # =================== CHECK 2 — MIN STAY ===================
        if notti >= min_stay_eff:
            c2r = ("verde", "Minimum stay",
                   f"Soggiorno di {notti} notti ≥ MLOS del periodo ({min_stay_eff}).")
        elif notti >= min_stay_eff - 1:
            c2r = ("giallo", "Minimum stay",
                   f"Soggiorno di {notti} notti contro MLOS {min_stay_eff}: "
                   f"deroga lieve, da autorizzare.")
        else:
            c2r = ("rosso", "Minimum stay",
                   f"Soggiorno di {notti} notti sotto il MLOS di {min_stay_eff}: "
                   f"deroga importante sulla durata minima.")

        # =================== CHECK 3 — ADR BED ===================
        gap = tariffa - soglia_bed
        if tariffa >= soglia_bed:
            c3r = ("verde", "ADR bed vs soglia",
                   f"Tariffa {eur2(tariffa)} ≥ soglia {eur2(soglia_bed)} "
                   f"({pct*100:.0f}% della FIT bed con occupancy {occupancy}%).")
        elif tariffa >= soglia_bed * 0.92:
            c3r = ("giallo", "ADR bed vs soglia",
                   f"Tariffa {eur2(tariffa)} di poco sotto la soglia {eur2(soglia_bed)} "
                   f"(gap {eur2(gap)}/pax).")
        else:
            c3r = ("rosso", "ADR bed vs soglia",
                   f"Tariffa {eur2(tariffa)} sotto la soglia {eur2(soglia_bed)} "
                   f"(gap {eur2(gap)}/pax).")

        # =================== CHECK 4 — DISPLACEMENT ===================
        if displacement > 0:
            c4r = ("verde", "Displacement netto",
                   f"Il gruppo genera {eur(displacement)} di valore incrementale "
                   f"rispetto alla vendita alternativa attesa.")
        elif displacement >= -0.05 * rev_alt_atteso:
            c4r = ("giallo", "Displacement netto",
                   f"Displacement marginalmente negativo ({eur(displacement)}): "
                   f"valore quasi equivalente all'alternativa.")
        else:
            c4r = ("rosso", "Displacement netto",
                   f"Il gruppo distrugge {eur(abs(displacement))} di valore "
                   f"rispetto alla vendita alternativa attesa.")

        checks = [c1r, c2r, c3r, c4r]
        stati = [c[0] for c in checks]

        # =================== VERDETTO ===================
        if "rosso" in stati:
            verdetto, vcol = "RIFIUTARE O RINEGOZIARE", "rosso"
        elif "giallo" in stati:
            verdetto, vcol = "VALUTARE — CONTROPROPOSTA CONSIGLIATA", "giallo"
        else:
            verdetto, vcol = "ACCETTARE", "verde"

        # ---------- OUTPUT ----------
        st.divider()

        # metriche gruppo
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Pax totali", f"{pax:.0f}", help=f"{camere} camere × {pax_cam} pax/cam")
        m2.metric("Bed nights", f"{bed_nights:.0f}")
        m3.metric("ADR bed gruppo", eur2(tariffa))
        m4.metric("ADR room gruppo", eur2(adr_room_gruppo))

        m5, m6, m7, m8 = st.columns(4)
        m5.metric("Ricavo camere", eur(rev_camere))
        m6.metric("Ricavo ancillare", eur(rev_anc))
        m7.metric("Valore totale gruppo", eur(rev_totale))
        m8.metric("Displacement netto", eur(displacement),
                  delta=f"{displacement/rev_alt_atteso*100:+.1f}% vs alternativa"
                  if rev_alt_atteso else None)

        # verdetto
        st.markdown(f"""
        <div class="vt-verdict" style="background:{COLOR[vcol]}">
          <h2>{ICON[vcol]}  {verdetto}</h2>
          <p>{nome_gruppo} · {check_in.strftime('%d/%m/%Y')} → {check_out.strftime('%d/%m/%Y')}
             · {notti} notti · {camere} camere · meal {meal}</p>
        </div>""", unsafe_allow_html=True)

        # check semaforo
        cL, cR = st.columns([3, 2])
        with cL:
            st.markdown("##### Esito controlli")
            for stato, titolo, dett in checks:
                st.markdown(f"""
                <div class="vt-check" style="background:{COLOR[stato]}">
                  <b>{ICON[stato]} {titolo}</b><br>{dett}
                </div>""", unsafe_allow_html=True)

            if rev_totale > s["auth"]:
                st.warning(f"📨 Valore totale {eur(rev_totale)} oltre la soglia di "
                           f"{eur(s['auth'])}: **richiede autorizzazione della direzione**.")

        with cR:
            st.markdown("##### Controproposta")
            st.markdown(f"""
            <div class="vt-card">
              <p style="margin:0 0 6px;font-size:.86rem;color:#555">
                 Tariffa bed di <b>break-even</b> (displacement = 0)</p>
              <p style="margin:0;font-size:1.4rem;font-weight:700;color:{PRIM}">{eur2(tariffa_be)}/pax</p>
              <hr style="margin:10px 0;border-color:#E4DCC9">
              <p style="margin:0 0 6px;font-size:.86rem;color:#555">
                 Soglia ADR bed (occupancy {occupancy}%)</p>
              <p style="margin:0;font-size:1.4rem;font-weight:700;color:{PRIM}">{eur2(soglia_bed)}/pax</p>
              <hr style="margin:10px 0;border-color:#E4DCC9">
              <p style="margin:0 0 6px;font-size:.86rem;color:#555">
                 ✅ Tariffa bed da richiedere</p>
              <p style="margin:0;font-size:1.6rem;font-weight:800;color:{ACCENT}">
                 {eur(controproposta)}/pax</p>
              <p style="margin:4px 0 0;font-size:.78rem;color:#777">
                 ≈ {eur(controproposta*pax_cam)}/camera ADR room</p>
            </div>""", unsafe_allow_html=True)

        # --- grafici ---
        g1, g2 = st.columns(2)
        with g1:
            fig = go.Figure()
            fig.add_bar(name="Ricavo camere", x=["Gruppo"], y=[rev_camere], marker_color=PRIM)
            fig.add_bar(name="Ricavo ancillare", x=["Gruppo"], y=[rev_anc], marker_color=ACCENT)
            fig.add_bar(name="Vendita alternativa attesa", x=["Alternativa"],
                        y=[rev_alt_atteso], marker_color="#9AA7AD")
            fig.update_layout(barmode="stack", title="Valore gruppo vs alternativa",
                              height=340, margin=dict(t=46, b=10, l=10, r=10),
                              legend=dict(orientation="h", y=-0.18))
            st.plotly_chart(fig, use_container_width=True)
        with g2:
            fig2 = go.Figure(go.Bar(
                x=["Proposta gruppo", "Soglia ADR", "FIT bed", "TO netto bed"],
                y=[tariffa, soglia_bed, fit_w, to_w],
                marker_color=[ACCENT, GIALLO, PRIM, "#9AA7AD"],
                text=[eur2(v) for v in [tariffa, soglia_bed, fit_w, to_w]],
                textposition="outside"))
            fig2.update_layout(title="ADR bed — confronto (€/pax/notte)",
                               height=340, margin=dict(t=46, b=10, l=10, r=10),
                               yaxis_title="€/pax/notte")
            st.plotly_chart(fig2, use_container_width=True)

        # --- dettaglio periodi ---
        with st.expander("🔎 Dettaglio periodi del soggiorno"):
            det = pd.DataFrame([
                {"Periodo": k, "Notti": v["notti"],
                 f"FIT bed {meal}": eur2(v["fit"]), f"TO bed {meal}": eur2(v["to"]),
                 "MLOS": v["min"], "Allotment periodo": v["allot"]}
                for k, v in seg.items()])
            st.dataframe(det, use_container_width=True, hide_index=True)
            st.caption(f"Tariffe pesate sul soggiorno → FIT bed {eur2(fit_w)} · "
                       f"TO bed {eur2(to_w)} · MLOS effettivo (più restrittivo) {min_stay_eff}.")

        # --- salva ---
        record = {
            "Gruppo": nome_gruppo, "Check-in": check_in.strftime("%d/%m/%Y"),
            "Check-out": check_out.strftime("%d/%m/%Y"), "Notti": notti,
            "Camere": camere, "Pax/cam": pax_cam, "Pax": round(pax),
            "Meal": meal, "ADR bed": round(tariffa, 2),
            "Valore totale": round(rev_totale), "Displacement": round(displacement),
            "Controproposta bed": controproposta, "Verdetto": verdetto,
        }
        if st.button("💾 Salva valutazione nel riepilogo", use_container_width=True):
            st.session_state.valutazioni.append(record)
            st.success("Valutazione salvata. La trovi nella sezione «Riepilogo valutazioni».")

    # ---------- METODOLOGIA ----------
    with st.expander("ℹ️ Metodologia di calcolo"):
        st.markdown("""
**Logica di valutazione**

- **ADR bed** è la metrica primaria: tariffa per pax/notte. L'**ADR room** è derivato moltiplicando per il rapporto pax/cam (default 2,25 per i gruppi leisure).
- I **soggiorni multi-periodo** vengono gestiti notte per notte: le tariffe FIT e TO sono medie ponderate sulle notti effettive in ciascun periodo.
- Il **displacement** confronta il valore totale del gruppo con la vendita alternativa attesa:
  - le camere entro l'allotment ALPI residuo sono valutate alla **tariffa netta TO**;
  - le camere oltre allotment erodono inventario diretto e sono valutate alla **tariffa FIT**;
  - il totale alternativo è ponderato per la **probabilità di pick-up** (quante di quelle camere si venderebbero davvero).
- La **soglia ADR bed** è una percentuale della tariffa FIT bed del periodo, crescente con l'occupancy attesa (configurabile nella barra laterale).
- La **controproposta** suggerita è la più alta tra la tariffa di break-even (displacement nullo) e la soglia ADR.
""")
