import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os
import io
import requests

# ── Config ───────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Tableau de Bord LBFI", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
  .main-header {
    background: #1a3a6b; color: white; padding: 14px 24px;
    border-radius: 8px; margin-bottom: 12px;
    font-size: 22px; font-weight: 700; letter-spacing: 0.5px;
  }
  .kpi-card {
    background: white; border-radius: 10px; padding: 20px 16px;
    text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.10);
  }
  .kpi-label { font-size: 13px; color: #555; margin-bottom: 6px; font-weight: 600; text-transform: uppercase; }
  .kpi-value { font-size: 36px; font-weight: 800; color: #1a3a6b; }
  .kpi-warn  { font-size: 36px; font-weight: 800; color: #e07b00; }
  .kpi-ok    { font-size: 36px; font-weight: 800; color: #1a8c4e; }
  .block-container { padding-top: 1rem !important; }
  div[data-testid="stDateInput"] label { font-weight: 600; font-size: 13px; color: #444; }
  div[data-testid="stSelectbox"] label { font-weight: 600; font-size: 13px; color: #444; }
</style>
""", unsafe_allow_html=True)

C_PROD  = "#4db8e8"
C_DEVIS = "#1a3a6b"

# Helper pour nettoyer les IDs
def clean_id(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "nat", "none", ""]: return ""
    if val_str.endswith(".0"): return val_str[:-2]
    return val_str

# ── Chargement des données avec gestion multi-feuilles ────────────────────────
@st.cache_data(ttl=3600)
def load_all_data(path_or_url_or_bytes):
    # Gestion du contournement du blocage Microsoft 403
    if isinstance(path_or_url_or_bytes, str) and path_or_url_or_bytes.startswith(("http://", "https://")):
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        response = requests.get(path_or_url_or_bytes, headers=headers)
        response.raise_for_status()
        file_to_read = io.BytesIO(response.content)
    else:
        file_to_read = path_or_url_or_bytes

    # Lecture globale du fichier Excel pour scanner les onglets
    xls = pd.ExcelFile(file_to_read, engine="openpyxl")
    sheets = xls.sheet_names

    # 1. Chargement de la feuille "employes" (pour la correspondance id -> Nom)
    df_emp = pd.DataFrame()
    sheet_emp = [s for s in sheets if "employ" in s.lower()]
    if sheet_emp:
        df_emp = pd.read_excel(xls, sheet_name=sheet_emp[0])
        # Détection automatique des colonnes ID et Nom
        id_cols = [c for c in df_emp.columns if any(p in c.lower() for p in ["id", "code", "matr"])]
        name_cols = [c for c in df_emp.columns if any(p in c.lower() for p in ["nom", "prenom", "agent", "opérateur", "operateur"])]
        
        col_map = {}
        if id_cols: col_map[id_cols[0]] = "id_operateur"
        if name_cols: col_map[name_cols[0]] = "nom_operateur"
        df_emp = df_emp.rename(columns=col_map)
        
        if "id_operateur" in df_emp.columns:
            df_emp["id_operateur"] = df_emp["id_operateur"].apply(clean_id)
            df_emp = df_emp.drop_duplicates(subset=["id_operateur"])

    # 2. Chargement de la feuille "ordres_fabrication"
    sheet_of = [s for s in sheets if "ordre" in s.lower() or "fabrication" in s.lower()]
    sheet_of = sheet_of[0] if sheet_of else "ordres_fabrication"
    
    df_of = pd.read_excel(xls, sheet_name=sheet_of)
    df_of["tps_op_h"]    = df_of["Temps opérateur (h)"] / 3_600_000
    df_of["tps_devis_h"] = df_of["Temps devis (nombre)"] / 3_600_000
    df_of["date_cloture"] = pd.to_datetime(df_of["date_cloture"], errors="coerce")
    df_of["semaine_label"] = (df_of["date_cloture"].dt.isocalendar().year.astype(str) +
                           "-S" + df_of["date_cloture"].dt.isocalendar().week.astype(str).str.zfill(2))
    df_of["mois_label"]  = df_of["date_cloture"].dt.to_period("M").astype(str)
    df_of["annee_label"] = df_of["date_cloture"].dt.year.astype(str)

    df_of["numero_dossier"] = df_of["numero_dossier"].apply(clean_id)
    df_of["numero_devis"]   = df_of["numero_devis"].apply(clean_id)
    df_of["client"]         = df_of["client"].fillna("Inconnu").astype(str).str.strip()
    df_of["reference"]      = df_of["reference"].fillna("Inconnu").astype(str).str.strip()
    df_of = df_of[df_of["date_cloture"].notna()].copy()

    # Tentative de liaison de l'opérateur sur la feuille "ordres_fabrication" s'il y a un ID
    id_op_of_cols = [c for c in df_of.columns if any(p in c.lower() for p in ["id_operateur", "id_opérateur", "id_employe", "code_operateur"])]
    if id_op_of_cols:
        df_of = df_of.rename(columns={id_op_of_cols[0]: "id_operateur"})
        df_of["id_operateur"] = df_of["id_operateur"].apply(clean_id)
        if not df_emp.empty and "id_operateur" in df_emp.columns:
            df_of = df_of.merge(df_emp, on="id_operateur", how="left")
            df_of["nom_operateur"] = df_of["nom_operateur"].fillna(df_of["id_operateur"].apply(lambda x: f"Opérateur {x}" if x else "Inconnu"))
        else:
            df_of["nom_operateur"] = df_of["id_operateur"].apply(lambda x: f"Opérateur {x}" if x else "Inconnu")
    else:
        df_of["nom_operateur"] = "Inconnu"

    # 3. Chargement de la feuille "temps réel opérateur"
    df_tr = pd.DataFrame()
    sheet_tr_candidates = [s for s in sheets if any(p in s.lower() for p in ["reel", "réel", "temps_reel", "temps reel"])]
    
    if sheet_tr_candidates:
        df_tr = pd.read_excel(xls, sheet_name=sheet_tr_candidates[0])
        # Détection des colonnes Date, ID Opérateur, et Heures
        tr_id_cols = [c for c in df_tr.columns if any(p in c.lower() for p in ["id", "code", "opérateur", "operateur", "employe", "employé"])]
        tr_date_cols = [c for c in df_tr.columns if any(p in c.lower() for p in ["date", "jour", "debut", "début", "saisie"])]
        tr_hour_cols = [c for c in df_tr.columns if any(p in c.lower() for p in ["heure", "hour", "temps", "durée", "duree", "tps", "travail"])]
        
        col_map_tr = {}
        if tr_id_cols: col_map_tr[tr_id_cols[0]] = "id_operateur"
        if tr_date_cols: col_map_tr[tr_date_cols[0]] = "date_travail"
        if tr_hour_cols: col_map_tr[tr_hour_cols[0]] = "heures_travaillees"
        
        df_tr = df_tr.rename(columns=col_map_tr)
        
        # Valeurs par défaut si colonnes manquantes
        if "id_operateur" not in df_tr.columns: df_tr["id_operateur"] = "Inconnu"
        if "date_travail" not in df_tr.columns: df_tr["date_travail"] = pd.Timestamp.today()
        if "heures_travaillees" not in df_tr.columns: df_tr["heures_travaillees"] = 0.0
        
        # Nettoyage formats
        df_tr["date_travail"] = pd.to_datetime(df_tr["date_travail"], errors="coerce")
        df_tr["heures_travaillees"] = pd.to_numeric(df_tr["heures_travaillees"], errors="coerce").fillna(0.0)
        df_tr["id_operateur"] = df_tr["id_operateur"].apply(clean_id)
        
        # Conversion si les heures sont stockées en millisecondes
        if df_tr["heures_travaillees"].max() > 100_000:
            df_tr["heures_travaillees"] = df_tr["heures_travaillees"] / 3_600_000

        # Fusion avec les employés pour avoir le nom lisible
        if not df_emp.empty and "id_operateur" in df_emp.columns:
            df_tr = df_tr.merge(df_emp, on="id_operateur", how="left")
            df_tr["nom_operateur"] = df_tr["nom_operateur"].fillna(df_tr["id_operateur"].apply(lambda x: f"Opérateur {x}" if x else "Inconnu"))
        else:
            df_tr["nom_operateur"] = df_tr["id_operateur"].apply(lambda x: f"Opérateur {x}" if x else "Inconnu")
            
        df_tr = df_tr[df_tr["date_travail"].notna()].copy()
        
        # Mailles temporelles pour le Temps Réel
        df_tr["semaine_label"] = (df_tr["date_travail"].dt.isocalendar().year.astype(str) +
                               "-S" + df_tr["date_travail"].dt.isocalendar().week.astype(str).str.zfill(2))
        df_tr["mois_label"]  = df_tr["date_travail"].dt.to_period("M").astype(str)
        df_tr["annee_label"] = df_tr["date_travail"].dt.year.astype(str)
        df_tr["jour_label"]  = df_tr["date_travail"].dt.strftime("%Y-%m-%d")
        df_tr["heure_label"] = df_tr["date_travail"].dt.hour.astype(str).str.zfill(2) + "h"

    return df_of, df_tr, df_emp

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="main-header">📊  Tableau de Bord LBFI – Suivi des Temps de Production</div>',
            unsafe_allow_html=True)

# ── Résolution du fichier (Cloud / Secrets) ───────────────────────────────────
excel_url = st.secrets.get("EXCEL_DATA_URL", None)
df_all = None
df_tr = pd.DataFrame()
df_emp = pd.DataFrame()
src_label = ""

if excel_url:
    try:
        df_all, df_tr, df_emp = load_all_data(excel_url)
        src_label = "Lien Cloud (OneDrive/SharePoint)"
    except Exception as e:
        st.error(f"❌ Erreur lors de l'accès automatique au fichier Cloud : {e}")

if df_all is None:
    st.warning("⚠️ Fichier source cloud non configuré ou inaccessible. Déposez le fichier Excel ci-dessous.")
    uploaded = st.file_uploader("Fichier **Suivi Activité LBFI DATA SOURCE.xlsm / .xlsx**", type=["xlsx","xlsm"])
    if uploaded is None:
        st.info("Pour automatiser cette étape, configurez la variable `EXCEL_DATA_URL` dans les secrets Streamlit.")
        st.stop()
    df_all, df_tr, df_emp = load_all_data(uploaded)
    src_label = "Fichier téléversé manuellement"

# ── Plage globale des données étendue aux deux tables ────────────────────────
today     = pd.Timestamp.today().normalize()
data_min  = df_all["date_cloture"].min().date()
data_max  = df_all["date_cloture"].max().date()

if not df_tr.empty:
    data_min = min(data_min, df_tr["date_travail"].min().date())
    data_max = max(data_max, df_tr["date_travail"].max().date())

ytd_start = max(pd.Timestamp(today.year, 1, 1).date(), data_min)
ytd_end   = min(today.date(), data_max)

# ── Session state ─────────────────────────────────────────────────────────────
def _init(key, val):
    if key not in st.session_state:
        st.session_state[key] = val

_init("date_debut",      ytd_start)
_init("date_fin",        ytd_end)
_init("sel_statut",      "Tous")
_init("sel_poste",       "Tous")
_init("sel_op",          "Tous")
_init("sel_operateur",   "Tous")
_init("_dossier_filtre", "Tous")
_init("_devis_filtre",   "Tous")
_init("_last_tbl_sel",   False)

def reset_filters():
    st.session_state["date_debut"]       = ytd_start
    st.session_state["date_fin"]         = ytd_end
    st.session_state["sel_statut"]       = "Tous"
    st.session_state["sel_poste"]        = "Tous"
    st.session_state["sel_op"]           = "Tous"
    st.session_state["sel_operateur"]    = "Tous"
    st.session_state["_dossier_filtre"]  = "Tous"
    st.session_state["_devis_filtre"]    = "Tous"
    st.session_state["maille"]           = "Mois"
    st.session_state["_last_tbl_sel"]    = False

# ── Ligne filtres ─────────────────────────────────────────────────────────────
row1 = st.columns([1.1, 1.1, 1, 1, 1, 0.8])
row2 = st.columns([1.1, 1.1, 1.1, 2.9])          

date_debut = row1[0].date_input("📅 Date début", min_value=data_min, max_value=data_max, key="date_debut")
date_fin   = row1[1].date_input("📅 Date fin", min_value=data_min, max_value=data_max, key="date_fin")

statuts    = ["Tous"] + sorted(df_all["statut_production"].dropna().unique().tolist())
sel_statut = row1[2].selectbox("Statut production", statuts, key="sel_statut")

postes     = ["Tous"] + sorted(df_all["poste"].dropna().unique().tolist())
sel_poste  = row1[3].selectbox("Poste", postes, key="sel_poste")

operations = ["Tous"] + sorted(df_all["operation"].dropna().unique().tolist())
sel_op     = row1[4].selectbox("Opération", operations, key="sel_op")

row1[5].markdown("<br>", unsafe_allow_html=True)
row1[5].button("🔄 Réinitialiser", on_click=reset_filters, use_container_width=True)

# Filtres dossier et devis
dossiers = ["Tous"] + sorted([d for d in df_all["numero_dossier"].unique() if d])
try:
    idx_dossier = dossiers.index(st.session_state["_dossier_filtre"])
except ValueError:
    idx_dossier = 0

def on_dossier_change():
    st.session_state["_dossier_filtre"] = st.session_state["widget_dossier"]

row2[0].selectbox("N° dossier fabrication", dossiers, index=idx_dossier, key="widget_dossier", on_change=on_dossier_change)

devis_list = ["Tous"] + sorted([v for v in df_all["numero_devis"].unique() if v])
try:
    idx_devis = devis_list.index(st.session_state["_devis_filtre"])
except ValueError:
    idx_devis = 0

def on_devis_change():
    st.session_state["_devis_filtre"] = st.session_state["widget_devis"]

row2[1].selectbox("N° de devis", devis_list, index=idx_devis, key="widget_devis", on_change=on_devis_change)

# Nouveau Filtre par Opérateur (depuis les feuilles TR ou OF)
if not df_tr.empty:
    operateurs = ["Tous"] + sorted(df_tr["nom_operateur"].dropna().unique().tolist())
else:
    operateurs = ["Tous"]
    
try:
    idx_op = operateurs.index(st.session_state["sel_operateur"])
except ValueError:
    idx_op = 0

sel_operateur = row2[2].selectbox("Opérateur", operateurs, index=idx_op, key="sel_operateur")
row2[3].empty()

# ── Appliquer tous les filtres ────────────────────────────────────────────────
dossier_actif = st.session_state["_dossier_filtre"]
devis_actif   = st.session_state["_devis_filtre"]

df = df_all.copy()
df = df[(df["date_cloture"].dt.date >= date_debut) & (df["date_cloture"].dt.date <= date_fin)]
if sel_statut    != "Tous": df = df[df["statut_production"] == sel_statut]
if sel_poste     != "Tous": df = df[df["poste"]             == sel_poste]
if sel_op        != "Tous": df = df[df["operation"]         == sel_op]
if dossier_actif != "Tous": df = df[df["numero_dossier"]    == dossier_actif]
if devis_actif   != "Tous": df = df[df["numero_devis"]      == devis_actif]

# Filtrer aussi la table de production classique par opérateur s'il est mappé
if sel_operateur != "Tous" and "nom_operateur" in df.columns:
    if not (df["nom_operateur"] == "Inconnu").all():
        df = df[df["nom_operateur"] == sel_operateur]

# ── KPIs Globaux ──────────────────────────────────────────────────────────────
total_op    = df["tps_op_h"].sum()
total_devis = df["tps_devis_h"].sum()
efficacite  = total_op / total_devis if total_devis > 0 else 0
eff_cls     = "kpi-ok" if efficacite <= 1 else "kpi-warn"

k1, k2, k3 = st.columns(3)
k1.markdown(f'<div class="kpi-card"><div class="kpi-label">⏱ Temps de production total</div><div class="kpi-value">{total_op:,.2f} h</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi-card"><div class="kpi-label">📋 Temps devis total</div><div class="kpi-value">{total_devis:,.2f} h</div></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi-card"><div class="kpi-label">⚡ Efficacité (Prod / Devis)</div><div class="{eff_cls}">{efficacite:.2f}</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Graphique temporel ────────────────────────────────────────────────────────
maille     = st.radio("Maille temporelle", ["Semaine","Mois","Année"], index=1, horizontal=True, key="maille")
maille_col = {"Semaine":"semaine_label","Mois":"mois_label","Année":"annee_label"}[maille]

by_maille  = df.groupby(maille_col)[["tps_op_h","tps_devis_h"]].sum().reset_index()
if maille == "Mois":
    by_maille["_s"] = pd.PeriodIndex(by_maille[maille_col], freq="M")
    by_maille = by_maille.sort_values("_s").drop(columns="_s")
else:
    by_maille = by_maille.sort_values(maille_col)

fig_maille = go.Figure([
    go.Bar(x=by_maille[maille_col], y=by_maille["tps_op_h"], name="Temps de production", marker_color=C_PROD),
    go.Bar(x=by_maille[maille_col], y=by_maille["tps_devis_h"], name="Temps devis", marker_color=C_DEVIS),
])
fig_maille.update_layout(
    title=f"Temps de production vs devis par {maille.lower()}",
    barmode="group", xaxis_title=maille, yaxis_title="Heures",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(t=50, b=40, l=40, r=20), height=340,
    xaxis_tickangle=-35 if maille == "Semaine" else 0,
)
st.plotly_chart(fig_maille, use_container_width=True)

# ── Graphique par poste + Tableau Dossiers ────────────────────────────────────
by_poste = (df.groupby("poste")[["tps_op_h","tps_devis_h"]].sum().reset_index().sort_values("tps_devis_h", ascending=True))
fig_poste = go.Figure([
    go.Bar(y=by_poste["poste"], x=by_poste["tps_op_h"], name="Temps de production", marker_color=C_PROD, orientation="h"),
    go.Bar(y=by_poste["poste"], x=by_poste["tps_devis_h"], name="Temps devis", marker_color=C_DEVIS, orientation="h"),
])
fig_poste.update_layout(
    title="Temps par poste", barmode="group", xaxis_title="Heures",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(t=50, b=40, l=10, r=20), height=320,
)

by_dossier = (df.groupby(["numero_dossier","numero_devis","client","reference"])[["tps_op_h","tps_devis_h"]].sum().reset_index())
by_dossier["ratio"] = (by_dossier["tps_op_h"] / by_dossier["tps_devis_h"].replace(0, float("nan"))).round(2)
by_dossier = by_dossier.sort_values("tps_devis_h", ascending=False).reset_index(drop=True)
by_dossier.columns = ["N° dossier","N° devis","Client","Référence","Tps prod (h)","Tps devis (h)","Ratio"]
by_dossier["Tps prod (h)"]  = by_dossier["Tps prod (h)"].round(2)
by_dossier["Tps devis (h)"] = by_dossier["Tps devis (h)"].round(2)

col_poste, col_tbl = st.columns([1, 1])
col_poste.plotly_chart(fig_poste, use_container_width=True)

with col_tbl:
    if dossier_actif != "Tous" or devis_actif != "Tous":
        label = []
        if dossier_actif != "Tous": label.append(f"dossier <strong>{dossier_actif}</strong>")
        if devis_actif   != "Tous": label.append(f"devis <strong>{devis_actif}</strong>")
        st.markdown(f'<div style="background:#e8f4fd;border-left:4px solid #1a3a6b;padding:6px 12px;border-radius:4px;margin-bottom:6px;font-size:13px;">🔍 Filtré sur : {" · ".join(label)}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="background:#f0f2f6;padding:6px 12px;border-radius:4px;margin-bottom:6px;font-size:13px;color:#666;">👆 Cliquez sur une ligne pour filtrer tout le dashboard sur ce dossier</div>', unsafe_allow_html=True)

    tbl_event = st.dataframe(by_dossier, use_container_width=True, height=270, hide_index=True, on_select="rerun", selection_mode="single-row", key="tbl_dossier")
    selected_rows = tbl_event.selection.rows if tbl_event.selection else []
    had_sel = st.session_state.get("_last_tbl_sel", False)

    if selected_rows:
        row_index = selected_rows[0]
        if row_index < len(by_dossier):
            row_data = by_dossier.iloc[row_index]
            clicked_dossier = row_data["N° dossier"]
            clicked_devis   = row_data["N° devis"]
            if dossier_actif != clicked_dossier or devis_actif != clicked_devis:
                st.session_state["_dossier_filtre"] = clicked_dossier
                st.session_state["_devis_filtre"]   = clicked_devis
                st.session_state["_last_tbl_sel"]   = True
                st.rerun()
    elif had_sel and not selected_rows:
        st.session_state["_dossier_filtre"] = "Tous"
        st.session_state["_devis_filtre"]   = "Tous"
        st.session_state["_last_tbl_sel"] = False
        st.rerun()

    st.session_state["_last_tbl_sel"] = bool(selected_rows) and (selected_rows[0] < len(by_dossier))

# ── Tableau par opération + graphique ────────────────────────────────────────
by_op_tbl = (df.groupby("operation")[["tps_op_h","tps_devis_h"]].sum().reset_index())
by_op_tbl["ratio"] = (by_op_tbl["tps_op_h"] / by_op_tbl["tps_devis_h"].replace(0, float("nan"))).round(2)
by_op_tbl = by_op_tbl.sort_values("tps_devis_h", ascending=False).reset_index(drop=True)
by_op_tbl.columns = ["Opération","Tps prod (h)","Tps devis (h)","Ratio"]
by_op_tbl["Tps prod (h)"]  = by_op_tbl["Tps prod (h)"].round(2)
by_op_tbl["Tps devis (h)"] = by_op_tbl["Tps devis (h)"].round(2)

by_op_chart = by_op_tbl.head(20)
fig_op = go.Figure([
    go.Bar(x=by_op_chart["Opération"], y=by_op_chart["Tps prod (h)"], name="Temps de production", marker_color=C_PROD),
    go.Bar(x=by_op_chart["Opération"], y=by_op_chart["Tps devis (h)"], name="Temps devis", marker_color=C_DEVIS),
])
fig_op.update_layout(
    title="Temps par opération (Top 20)", barmode="group",
    xaxis_title="", yaxis_title="Heures", xaxis_tickangle=-35,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(t=50, b=150, l=40, r=20), height=420,
)

col_op_chart, col_op_tbl = st.columns([1.3, 1])
col_op_chart.plotly_chart(fig_op, use_container_width=True)
with col_op_tbl:
    st.markdown("**Détail par opération**")
    st.dataframe(by_op_tbl, use_container_width=True, height=390, hide_index=True)

# ── NOUVELLE SECTION : TEMPS RÉEL OPÉRATEURS ──────────────────────────────────
if not df_tr.empty:
    st.markdown("---")
    st.markdown("### 🕒 Suivi du Temps Réel des Opérateurs")

    # Filtrer df_tr selon les dates globales et l'opérateur sélectionné
    df_tr_filtered = df_tr[(df_tr["date_travail"].dt.date >= date_debut) & (df_tr["date_travail"].dt.date <= date_fin)]
    if sel_operateur != "Tous":
        df_tr_filtered = df_tr_filtered[df_tr_filtered["nom_operateur"] == sel_operateur]

    # KPIs Temps Réel
    tr_k1, tr_k2, tr_k3 = st.columns(3)
    total_hours_tr = df_tr_filtered["heures_travaillees"].sum()
    active_ops_count = df_tr_filtered["nom_operateur"].nunique()
    avg_hours_tr = total_hours_tr / active_ops_count if active_ops_count > 0 else 0

    tr_k1.markdown(f'<div class="kpi-card"><div class="kpi-label">⏳ Total heures réelles travaillées</div><div class="kpi-value">{total_hours_tr:,.2f} h</div></div>', unsafe_allow_html=True)
    tr_k2.markdown(f'<div class="kpi-card"><div class="kpi-label">👥 Opérateurs Actifs</div><div class="kpi-value">{active_ops_count}</div></div>', unsafe_allow_html=True)
    tr_k3.markdown(f'<div class="kpi-card"><div class="kpi-label">📊 Moyenne d\'Heures / Opérateur</div><div class="kpi-value">{avg_hours_tr:,.2f} h</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Sélecteur de maille temporelle spécifique au temps réel (Jour / Semaine / Heure)
    col_tr_radio, col_tr_empty = st.columns([1.5, 2.5])
    maille_tr = col_tr_radio.radio(
        "Maille temporelle (Suivi Opérateur)", 
        ["Jour", "Semaine", "Heure"], 
        index=0, 
        horizontal=True, 
        key="maille_tr"
    )
    
    maille_tr_col = {
        "Jour": "jour_label",
        "Semaine": "semaine_label",
        "Heure": "heure_label"
    }[maille_tr]

    # Agrégation des données
    by_maille_tr = df_tr_filtered.groupby([maille_tr_col, "nom_operateur"])["heures_travaillees"].sum().reset_index()
    by_maille_tr = by_maille_tr.sort_values(maille_tr_col)

    # Génération du graphique Plotly (Barres empilées pour voir le total cumulé et la part de chacun)
    fig_tr = go.Figure()
    for op in by_maille_tr["nom_operateur"].unique():
        op_sub = by_maille_tr[by_maille_tr["nom_operateur"] == op]
        fig_tr.add_trace(go.Bar(
            x=op_sub[maille_tr_col],
            y=op_sub["heures_travaillees"],
            name=op
        ))

    fig_tr.update_layout(
        title=f"Heures travaillées par {maille_tr.lower()} et opérateur",
        barmode="stack",
        xaxis_title=maille_tr,
        yaxis_title="Heures",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50, b=40, l=40, r=20), height=380,
    )

    # Tableau de synthèse des opérateurs
    by_op_tbl = df_tr_filtered.groupby("nom_operateur")["heures_travaillees"].sum().reset_index()
    by_op_tbl = by_op_tbl.sort_values("heures_travaillees", ascending=False).reset_index(drop=True)
    by_op_tbl.columns = ["Opérateur", "Total heures (h)"]
    by_op_tbl["Total heures (h)"] = by_op_tbl["Total heures (h)"].round(2)

    col_tr_chart, col_tr_tbl = st.columns([1.3, 1])
    col_tr_chart.plotly_chart(fig_tr, use_container_width=True)
    with col_tr_tbl:
        st.markdown("**Synthèse des heures réelles par opérateur**")
        st.dataframe(by_op_tbl, use_container_width=True, height=330, hide_index=True)
else:
    st.markdown("---")
    st.warning("⚠️ Onglet 'Temps Réel Opérateur' absent ou illisible dans le fichier source.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(f"📁 Source : `{src_label}` · {len(df_all):,} lignes OF · {len(df_tr) if not df_tr.empty else 0:,} lignes Temps Réel · Période : {date_debut} → {date_fin}")
