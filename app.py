# app.py — Dashboard Performance SDA — versione Streamlit web + Supabase
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import date, timedelta
import io

from core import (
    leggi_file_corrieri,
    aggrega_filiale,
    calcola_tariffa,
    FASCE_DEFAULT,
    ottieni_fasce_filiale,
)

st.set_page_config(
    page_title="Dashboard Performance SDA",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

COLORI_FILIALI = ["#3b82f6", "#22c55e", "#a855f7", "#f59e0b", "#14b8a6", "#ef4444", "#ec4899", "#f97316"]

LAYOUT_DARK = dict(
    plot_bgcolor="#181c24", paper_bgcolor="#0f1117", font_color="#f1f5f9",
    legend=dict(bgcolor="#1e2330", bordercolor="#2a3045"),
    margin=dict(l=0, r=0, t=30, b=0),
)

st.markdown("""
<style>
    .kpi-box {
        background: #1e2330;
        border: 1px solid #2a3045;
        border-radius: 10px;
        padding: 16px 18px 12px;
        text-align: center;
    }
    .kpi-val  { font-size: 2rem; font-weight: 700; margin: 0; line-height: 1.1; }
    .kpi-lbl  { font-size: 0.78rem; color: #94a3b8; margin-top: 4px; }
    div[data-testid="stDataFrame"] { font-size: 0.85rem; }
    section[data-testid="stSidebar"] { background: #181c24; }
</style>
""", unsafe_allow_html=True)

def kpi_card(label: str, value: str, color: str = "#3b82f6"):
    st.markdown(
        f'<div class="kpi-box">'
        f'<p class="kpi-val" style="color:{color}">{value}</p>'
        f'<p class="kpi-lbl">{label}</p>'
        f'</div>',
        unsafe_allow_html=True,
    )

def fmt_eur(v: float) -> str:
    return f"€ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_n(v: float) -> str:
    return f"{int(v):,}".replace(",", ".")

def colore_filiale(filiali: list, nome: str) -> str:
    try:
        return COLORI_FILIALI[filiali.index(nome) % len(COLORI_FILIALI)]
    except ValueError:
        return "#3b82f6"

# ── SUPABASE ──────────────────────────────────────────────────
from supabase import create_client

@st.cache_resource
def get_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

def get_progetto() -> str:
    """Legge il nome del progetto dai secrets Streamlit."""
    return st.secrets["PROGETTO"]

def carica_da_supabase() -> dict | None:
    """Legge i record del progetto corrente da Supabase."""
    sb       = get_supabase()
    progetto = get_progetto()
    rows     = []
    offset   = 0
    while True:
        chunk = (
            sb.table("performance_corrieri")
            .select("*")
            .eq("progetto", progetto)          # ← filtra per progetto
            .range(offset, offset + 999)
            .execute()
            .data
        )
        rows.extend(chunk)
        if len(chunk) < 1000:
            break
        offset += 1000

    if not rows:
        return None

    # Ricostruisce la struttura: {filiale: {data: {giro: {...}}}}
    dati: dict = {}
    for r in rows:
        fil  = r["filiale"]
        d    = date.fromisoformat(r["data"])
        giro = r["giro"]
        dati.setdefault(fil, {}).setdefault(d, {})[giro] = {
            "lv_af":   r["lv_af"],
            "lv_ok":   r["lv_ok"],
            "lv_rit":  r["lv_rit"],
            "stop_ok": r["stop_ok"],
            "stop_rit": r["stop_rit"],
            "ldv_tot": r["lv_ok"] + r["lv_rit"],
        }
    return dati

def importa_su_supabase(dati_nuovi: dict) -> int:
    """
    Fa upsert dei dati su Supabase (accoda senza duplicare).
    La tabella deve avere UNIQUE (progetto, filiale, data, giro).
    """
    sb       = get_supabase()
    progetto = get_progetto()
    records  = []
    for filiale, giorni in dati_nuovi.items():
        for d, giri in giorni.items():
            for giro, v in giri.items():
                records.append({
                    "progetto": progetto,          # ← aggiunto
                    "filiale":  filiale,
                    "data":     d.isoformat(),
                    "giro":     str(giro),
                    "lv_af":    int(v.get("lv_af",   0)),
                    "lv_ok":    int(v.get("lv_ok",   0)),
                    "lv_rit":   int(v.get("lv_rit",  0)),
                    "stop_ok":  int(v.get("stop_ok", 0)),
                    "stop_rit": int(v.get("stop_rit", 0)),
                })

    # Upsert a blocchi di 500 record
    for i in range(0, len(records), 500):
        sb.table("performance_corrieri").upsert(
            records[i:i + 500],
            on_conflict="progetto,filiale,data,giro"  # ← aggiunto progetto
        ).execute()

    return len(records)

# ── SESSION STATE ──────────────────────────────────────────────
if "dati"    not in st.session_state: st.session_state.dati    = None
if "date_da" not in st.session_state: st.session_state.date_da = None
if "date_a"  not in st.session_state: st.session_state.date_a  = None

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📦 Dashboard Performance")
    progetto_attivo = st.secrets.get("PROGETTO", "—")
    st.markdown(f"🗂 **Progetto:** `{progetto_attivo}`")
    st.markdown("---")

    # ── SEZIONE IMPORTA DATI ──
    st.markdown("### 📥 Importa Dati")
    st.caption("Carica un file Excel: i dati vengono accodati a quelli esistenti su Supabase.")

    uploaded = st.file_uploader("Seleziona file Excel", type=["xlsx", "xls"])
    if uploaded:
        if st.button("⬆ Importa su Supabase", use_container_width=True, type="primary"):
            with st.spinner("Lettura e importazione in corso..."):
                try:
                    dati_nuovi = leggi_file_corrieri(uploaded)
                    n_righe = importa_su_supabase(dati_nuovi)
                    st.success(f"✅ {n_righe} record importati!")
                    st.session_state.dati = None  # forza ricaricamento dal DB
                    st.rerun()
                except Exception as e:
                    st.error(f"Errore importazione: {e}")

    st.markdown("---")

    # ── CARICA DATI DA SUPABASE ──
    if st.session_state.dati is None:
        with st.spinner("Caricamento dati da Supabase..."):
            try:
                st.session_state.dati = carica_da_supabase()
                if st.session_state.dati:
                    st.success(f"✅ {len(st.session_state.dati)} filiali caricate")
                else:
                    st.warning("Nessun dato presente. Importa un file Excel.")
            except Exception as e:
                st.error(f"Errore connessione Supabase: {e}")

    if st.button("🔄 Ricarica da Supabase", use_container_width=True):
        st.session_state.dati = None
        st.rerun()

    st.markdown("---")

    # ── FILTRO DATE ──
    if st.session_state.dati:
        tutte_date = sorted({d for fil in st.session_state.dati.values() for d in fil})
        if tutte_date:
            st.markdown("### 📅 Periodo")
            col1, col2 = st.columns(2)
            with col1:
                date_da = st.date_input("Dal", value=tutte_date[0], min_value=tutte_date[0], max_value=tutte_date[-1])
            with col2:
                date_a = st.date_input("Al", value=tutte_date[-1], min_value=tutte_date[0], max_value=tutte_date[-1])
            st.session_state.date_da = date_da
            st.session_state.date_a  = date_a
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("Oggi"):
                    st.session_state.date_da = st.session_state.date_a = tutte_date[-1]
                    st.rerun()
            with c2:
                if st.button("7gg"):
                    st.session_state.date_a  = tutte_date[-1]
                    st.session_state.date_da = max(tutte_date[0], tutte_date[-1] - timedelta(days=6))
                    st.rerun()
            with c3:
                if st.button("Tutto"):
                    st.session_state.date_da = tutte_date[0]
                    st.session_state.date_a  = tutte_date[-1]
                    st.rerun()

if not st.session_state.dati:
    st.info("👈 Importa un file Excel dalla barra laterale, oppure attendi il caricamento da Supabase.")
    st.stop()

dati    = st.session_state.dati
filiali = sorted(dati.keys())
date_da = st.session_state.date_da
date_a  = st.session_state.date_a

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Panoramica",
    "🏢 Dettaglio Filiale",
    "📋 Tutti i Giri",
    "📅 Giornaliero",
    "💶 Tariffa",
])

# ══════════════════════════════════════════════════════════════
# TAB 1 — PANORAMICA
# ══════════════════════════════════════════════════════════════
with tab1:
    st.markdown("### Panoramica Multi-Filiale")
    riepilogo = []
    for fil in filiali:
        agg, _, _ = aggrega_filiale(dati[fil], date_da, date_a)
        if agg:
            riepilogo.append({"filiale": fil, **agg})

    if riepilogo:
        df_riep = pd.DataFrame(riepilogo)
        tot_af  = df_riep["tot_lv_af"].sum()
        tot_ok  = df_riep["tot_lv_ok"].sum()
        tot_rit = df_riep["tot_lv_rit"].sum()
        tot_giri_giorni = sum(
            sum(len(giri) for d, giri in dati[f].items()
                if (date_da is None or d >= date_da) and (date_a is None or d <= date_a))
            for f in filiali)
        prod_complessiva_media = (tot_ok + tot_rit) / tot_giri_giorni if tot_giri_giorni > 0 else 0.0

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: kpi_card("Filiali attive",       str(len(riepilogo)),             "#3b82f6")
        with c2: kpi_card("Tot LV Affidate",       fmt_n(tot_af),                  "#3b82f6")
        with c3: kpi_card("Tot LV Ok",             fmt_n(tot_ok),                  "#22c55e")
        with c4: kpi_card("Tot LV Ritiro",         fmt_n(tot_rit),                 "#a855f7")
        with c5: kpi_card("Prod. Media Corrieri",  f"{prod_complessiva_media:.1f}", "#f59e0b")

        st.markdown("---")

        # GRAFICO 1 — Produttività LDV OK+RIT per filiale
        st.markdown("#### Produttività Media Corrieri (LDV OK+RIT) per Filiale")
        fig_prod = go.Figure()
        for i, row in df_riep.iterrows():
            col = colore_filiale(filiali, row["filiale"])
            fig_prod.add_trace(go.Bar(
                x=[row["filiale"]], y=[round(row["media_prod"], 1)],
                marker_color=col, name=row["filiale"],
                text=[f"{row['media_prod']:.1f}"], textposition="outside",
                showlegend=False,
            ))
        fig_prod.update_layout(**LAYOUT_DARK, height=300, xaxis=dict(gridcolor="#2a3045"), yaxis=dict(gridcolor="#2a3045", title="Media Giornaliera Pezzi per Corriere"))
        st.plotly_chart(fig_prod, use_container_width=True)

        # GRAFICO 2 — LV Ok vs LV Ritiro per filiale
        st.markdown("#### LV Ok vs LV Ritiro per Filiale")
        fig_lv = go.Figure()
        fig_lv.add_trace(go.Bar(
            name="LV Ok",
            x=df_riep["filiale"], y=df_riep["tot_lv_ok"],
            marker_color="#22c55e",
            text=df_riep["tot_lv_ok"].apply(lambda v: fmt_n(v)),
            textposition="outside",
        ))
        fig_lv.add_trace(go.Bar(
            name="LV Ritiro",
            x=df_riep["filiale"], y=df_riep["tot_lv_rit"],
            marker_color="#a855f7",
            text=df_riep["tot_lv_rit"].apply(lambda v: fmt_n(v)),
            textposition="outside",
        ))
        fig_lv.update_layout(**LAYOUT_DARK, barmode="group", height=300, xaxis=dict(gridcolor="#2a3045"), yaxis=dict(gridcolor="#2a3045"))
        st.plotly_chart(fig_lv, use_container_width=True)

        # TABELLA RIEPILOGATIVA
        st.markdown("#### Tabella Riepilogativa Filiali")
        df_tab = df_riep[["filiale", "n_giorni", "tot_lv_af", "tot_lv_ok", "tot_lv_rit", "media_prod"]].copy()
        df_tab.columns = ["Filiale", "Giorni Periodo", "Tot LV AF", "Tot LV Ok", "Tot LV Rit", "Prod. Media Corriere"]
        df_tab["Tot LV AF"]  = df_tab["Tot LV AF"].apply(fmt_n)
        df_tab["Tot LV Ok"]  = df_tab["Tot LV Ok"].apply(fmt_n)
        df_tab["Tot LV Rit"] = df_tab["Tot LV Rit"].apply(fmt_n)
        df_tab["Prod. Media Corriere"] = df_tab["Prod. Media Corriere"].apply(lambda x: f"{x:.1f}")
        st.dataframe(df_tab, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
# TAB 2 — DETTAGLIO FILIALE
# ══════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### Dettaglio Singola Filiale")
    fil_sel = st.selectbox("Seleziona filiale", filiali, key="sel_fil")
    agg, giornate, per_giro = aggrega_filiale(dati[fil_sel], date_da, date_a)

    if agg:
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: kpi_card("Giorni Attivi",        str(agg["n_giorni"]),        "#94a3b8")
        with c2: kpi_card("Tot LV Affidate",       fmt_n(agg["tot_lv_af"]),    "#3b82f6")
        with c3: kpi_card("Tot LV Ok",             fmt_n(agg["tot_lv_ok"]),    "#22c55e")
        with c4: kpi_card("Tot LV Rit",            fmt_n(agg["tot_lv_rit"]),   "#a855f7")
        with c5: kpi_card("Prod. Media Corrieri",  f"{agg['media_prod']:.1f}", "#f59e0b")

        st.markdown("---")

        # GRAFICO — Andamento giornaliero LV Ok e LV Rit
        st.markdown("#### Andamento Giornaliero LV Ok e LV Ritiro")
        giorni_data = []
        for d in sorted(giornate):
            lv_ok_d  = sum(v["lv_ok"]  for v in giornate[d].values())
            lv_rit_d = sum(v["lv_rit"] for v in giornate[d].values())
            giorni_data.append({"data": d, "lv_ok": lv_ok_d, "lv_rit": lv_rit_d})
        df_giorni = pd.DataFrame(giorni_data)

        if not df_giorni.empty:
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Scatter(
                x=df_giorni["data"], y=df_giorni["lv_ok"],
                name="LV Ok", line=dict(color="#22c55e", width=2),
                fill="tozeroy", fillcolor="rgba(34,197,94,0.1)",
                mode="lines+markers", marker=dict(size=5),
            ))
            fig_trend.add_trace(go.Scatter(
                x=df_giorni["data"], y=df_giorni["lv_rit"],
                name="LV Rit", line=dict(color="#a855f7", width=2),
                mode="lines+markers", marker=dict(size=5),
            ))
            fig_trend.update_layout(**LAYOUT_DARK, height=280, xaxis=dict(gridcolor="#2a3045"), yaxis=dict(gridcolor="#2a3045"))
            st.plotly_chart(fig_trend, use_container_width=True)

        # GRAFICO — Produttività media per giro (barre orizzontali)
        st.markdown("#### Produttività Media per Giro (LDV OK+RIT)")
        if per_giro:
            giri_sorted = sorted(per_giro.items())
            fig_giri = go.Figure(go.Bar(
                y=[f"Giro {g}" for g, _ in giri_sorted],
                x=[round(v["ldv_tot"], 1) for _, v in giri_sorted],
                orientation="h",
                marker_color=[colore_filiale(filiali, fil_sel)] * len(giri_sorted),
                text=[f"{v['ldv_tot']:.1f}" for _, v in giri_sorted],
                textposition="outside",
            ))
            fig_giri.update_layout(**LAYOUT_DARK, height=max(250, len(giri_sorted) * 28), xaxis=dict(gridcolor="#2a3045", title="LDV OK+RIT medi/giorno"), yaxis=dict(gridcolor="#2a3045", autorange="reversed"))
            fig_giri.update_layout(margin=dict(l=0, r=60, t=20, b=0))
            st.plotly_chart(fig_giri, use_container_width=True)

        # TABELLA GIRI
        st.markdown("#### Rendimento dei singoli Giri (Medie Giornaliere del Periodo)")
        if per_giro:
            rows_giro = [{
                "Giro":                              g,
                "Giorni Presenza":                   v["n"],
                "LV AF (Media)":                     f"{v['lv_af']:.1f}",
                "LV Ok (Media)":                     f"{v['lv_ok']:.1f}",
                "LV Rit (Media)":                    f"{v['lv_rit']:.1f}",
                "Stop Ok (Media)":                   f"{v['stop_ok']:.1f}",
                "Stop Rit (Media)":                  f"{v['stop_rit']:.1f}",
                "Prod Media Giorno (LV OK + RIT)":   f"{v['ldv_tot']:.1f}",
            } for g, v in sorted(per_giro.items())]
            st.dataframe(pd.DataFrame(rows_giro), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
# TAB 3 — TUTTI I GIRI
# ══════════════════════════════════════════════════════════════
with tab3:
    st.markdown("### Elenco Completo Multi-Filiale di tutti i Giri")
    righe_tutti = []
    for fil in filiali:
        _, _, pg = aggrega_filiale(dati[fil], date_da, date_a)
        if pg:
            for giro, v in sorted(pg.items()):
                righe_tutti.append({
                    "Filiale":                           fil,
                    "Giro":                              giro,
                    "Giorni":                            v["n"],
                    "LV AF (Media)":                     round(v["lv_af"],   1),
                    "LV Ok (Media)":                     round(v["lv_ok"],   1),
                    "LV Rit (Media)":                    round(v["lv_rit"],  1),
                    "Stop Ok (Media)":                   round(v["stop_ok"], 1),
                    "Stop Rit (Media)":                  round(v["stop_rit"],1),
                    "Prod Media Giorno (LV OK + RIT)":   round(v["ldv_tot"], 1),
                })

    if righe_tutti:
        df_tutti = pd.DataFrame(righe_tutti)
        fil_filter = st.multiselect("Filtra per filiale", filiali, default=filiali, key="filter_tutti")
        df_tutti_f = df_tutti[df_tutti["Filiale"].isin(fil_filter)]

        # GRAFICO — Confronto produttività per giro tra filiali selezionate
        if not df_tutti_f.empty and len(fil_filter) > 0:
            st.markdown("#### Confronto Produttività per Giro tra Filiali")
            fig_conf = go.Figure()
            for fil in fil_filter:
                df_f = df_tutti_f[df_tutti_f["Filiale"] == fil]
                fig_conf.add_trace(go.Bar(
                    name=fil,
                    x=df_f["Giro"].astype(str),
                    y=df_f["Prod Media Giorno (LV OK + RIT)"],
                    marker_color=colore_filiale(filiali, fil),
                ))
            fig_conf.update_layout(**LAYOUT_DARK, barmode="group", height=320, xaxis=dict(gridcolor="#2a3045", title="Giro"), yaxis=dict(gridcolor="#2a3045", title="LDV OK+RIT medi/giorno"))
            st.plotly_chart(fig_conf, use_container_width=True)

        st.dataframe(df_tutti_f, use_container_width=True, hide_index=True)

        # Download CSV
        csv = df_tutti_f.to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button("⬇ Scarica CSV", data=csv,
                           file_name="tutti_giri.csv", mime="text/csv")

# ══════════════════════════════════════════════════════════════
# TAB 4 — GIORNALIERO
# ══════════════════════════════════════════════════════════════

with tab4:

    st.markdown("### Dettaglio Giornaliero per Filiale")

    fil_giorno = st.selectbox(
        "Seleziona filiale",
        filiali,
        key="fil_giorno"
    )

    _, giornate_g, _ = aggrega_filiale(
        dati[fil_giorno],
        date_da,
        date_a
    )

    if giornate_g:

        date_sel = st.selectbox(
            "Seleziona data",
            sorted(giornate_g.keys(), reverse=True),
            format_func=lambda d: d.strftime("%d/%m/%Y")
        )

        giri_day = giornate_g[date_sel]

        lv_af_g = sum(
            v.get("lv_af",0)
            for v in giri_day.values()
        )

        lv_ok_g = sum(
            v.get("lv_ok",0)
            for v in giri_day.values()
        )

        lv_rit_g = sum(
            v.get("lv_rit",0)
            for v in giri_day.values()
        )

        stop_ok = sum(
            v.get("stop_ok",0)
            for v in giri_day.values()
        )

        stop_rit = sum(
            v.get("stop_rit",0)
            for v in giri_day.values()
        )

        tot_ldv = lv_ok_g + lv_rit_g

        n_giri = len(giri_day)

        prod_media = (
            tot_ldv / n_giri
            if n_giri > 0
            else 0
        )

        rdc = (
            lv_ok_g / lv_af_g *100
            if lv_af_g >0
            else 0
        )

        perc_rit = (
            lv_rit_g / tot_ldv *100
            if tot_ldv >0
            else 0
        )

        # KPI CONSEGNE

        st.markdown("## 📦 KPI CONSEGNE")

        c1,c2,c3,c4,c5 = st.columns(5)

        with c1:
            kpi_card(
                "LV Affidate",
                fmt_n(lv_af_g),
                "#3b82f6"
            )

        with c2:
            kpi_card(
                "LV OK",
                fmt_n(lv_ok_g),
                "#22c55e"
            )

        with c3:
            kpi_card(
                "Volume Totale",
                fmt_n(tot_ldv),
                "#94a3b8"
            )

        with c4:
            kpi_card(
                "Prod.Media",
                f"{prod_media:.1f}",
                "#f59e0b"
            )

        with c5:
            kpi_card(
                "RDC",
                f"{rdc:.1f}%",
                "#ef4444"
            )

        st.markdown("---")

        # RDC PER GIRO

        st.markdown(
            "#### Classifica RDC per Giro"
        )

        righe_rdc=[]

        for g,v in giri_day.items():

            lv_af = int(v.get("lv_af",0))
            lv_ok = int(v.get("lv_ok",0))

            rdc_giro = (
                (lv_ok/lv_af)*100
                if lv_af>0
                else 0
            )

            righe_rdc.append({

                "Giro":g,
                "RDC":rdc_giro

            })

        df_rdc = pd.DataFrame(
            righe_rdc
        )

        df_rdc = df_rdc.sort_values(
            by="RDC",
            ascending=False
        )

        fig_day = go.Figure()

        fig_day.add_trace(

            go.Bar(

                y=[
                    f"Giro {g}"
                    for g in df_rdc["Giro"]
                ],

                x=df_rdc["RDC"],

                orientation="h",

                text=[
                    f"{x:.1f}%"
                    for x in df_rdc["RDC"]
                ],

                textposition="inside",

                insidetextanchor="middle",

                textfont=dict(
                    color="white",
                    size=12
                ),

                marker_color=[

                    "#00C853" if x>=97
                    else "#FFD600" if x>=94
                    else "#FF3D00"

                    for x in df_rdc["RDC"]

                ]

            )

        )

        fig_day.update_layout(

            **LAYOUT_DARK,

            height=max(
                250,
                len(df_rdc)*30
            ),

            xaxis=dict(
                title="RDC %",
                range=[0,100],
                gridcolor="#2a3045"
            ),

            yaxis=dict(
                autorange="reversed"
            )

        )

        st.plotly_chart(
            fig_day,
            use_container_width=True
        )

        st.markdown("---")

        # KPI RITIRI

        st.markdown("## 🚚 KPI RITIRI")

        c1,c2,c3 = st.columns(3)

        with c1:
            kpi_card(
                "LV Ritiro",
                fmt_n(lv_rit_g),
                "#a855f7"
            )

        with c2:
            kpi_card(
                "Stop Ritiro",
                fmt_n(stop_rit),
                "#14b8a6"
            )

        with c3:
            kpi_card(
                "% Ritiri",
                f"{perc_rit:.1f}%",
                "#f59e0b"
            )

        st.markdown("---")

        # KPI NPS

        st.markdown("## ⭐ KPI NPS")

        st.info(
            "In attesa collegamento file NPS"
        )

    else:

        st.warning(
            "Nessun dato disponibile nel periodo selezionato."
        )
# ══════════════════════════════════════════════════════════════
# TAB 5 — TARIFFA
# ══════════════════════════════════════════════════════════════
with tab5:
    st.markdown("### Calcolo Fatturato a Scaglioni Progressivi")
    fil_tar = st.selectbox("Seleziona filiale per tariffazione", filiali, key="fil_tar")
    _, giornate_t, _ = aggrega_filiale(dati[fil_tar], date_da, date_a)

    if giornate_t:
        fasce_attive = ottieni_fasce_filiale(fil_tar)

        with st.expander(f"Ispeziona scaglioni attivi per la filiale: {fil_tar}", expanded=False):
            for idx, f in enumerate(fasce_attive):
                st.write(f"Scaglione {idx+1} ➔ Da: {f['da']} a {f['a']} LDV | Tariffa: € {f['prezzo']:.3f}")

        righe_tar, tot_v, tot_f, med_f = calcola_tariffa(giornate_t, fasce_attive)

        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi_card("Giorni",                   str(len(righe_tar)),  "#94a3b8")
        with c2: kpi_card("Volume Totale (LV OK+RIT)", fmt_n(tot_v),        "#3b82f6")
        with c3: kpi_card("Fatturato Totale Stimato",  fmt_eur(tot_f),      "#22c55e")
        with c4: kpi_card("Fatturato Medio / Giorno",  fmt_eur(med_f),      "#a855f7")

        st.markdown("---")

        # GRAFICO — Fatturato giornaliero
        st.markdown("#### Fatturato Giornaliero")
        df_tar = pd.DataFrame(righe_tar)
        fig_tar = go.Figure(go.Bar(
            x=df_tar["data"],
            y=df_tar["fatturato"],
            marker_color="#22c55e",
            text=[fmt_eur(v) for v in df_tar["fatturato"]],
            textposition="outside",
        ))
        fig_tar.update_layout(**LAYOUT_DARK, height=300, xaxis=dict(gridcolor="#2a3045"), yaxis=dict(gridcolor="#2a3045", tickprefix="€ "))
        fig_tar.update_layout(margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig_tar, use_container_width=True)

        # GRAFICO — Volume LDV giornaliero
        st.markdown("#### Volume LDV (OK+RIT) Giornaliero")
        fig_vol = go.Figure(go.Bar(
            x=df_tar["data"],
            y=df_tar["volume"],
            marker_color="#3b82f6",
            text=[fmt_n(v) for v in df_tar["volume"]],
            textposition="outside",
        ))
        fig_vol.update_layout(**LAYOUT_DARK, height=260, xaxis=dict(gridcolor="#2a3045"), yaxis=dict(gridcolor="#2a3045", title="LDV"))
        fig_vol.update_layout(margin=dict(l=0, r=0, t=20, b=0))
        st.plotly_chart(fig_vol, use_container_width=True)

        st.markdown("#### Dettaglio giornaliero scaglioni")
        st.dataframe(df_tar, use_container_width=True, hide_index=True)
    else:
        st.warning("Nessun dato disponibile nel periodo selezionato.")
