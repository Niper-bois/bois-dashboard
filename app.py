from __future__ import annotations

import io
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

st.set_page_config(
    page_title="BOIS Dashboard V5",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL_PATH = APP_DIR / "data" / "BOIS_Excel_Master_V5.xlsx"
if not DEFAULT_EXCEL_PATH.exists():
    fallback_excel = APP_DIR / "BOIS_Excel_Master_V5.xlsx"
    if fallback_excel.exists():
        DEFAULT_EXCEL_PATH = fallback_excel

SHEETS_STANDARD = {
    "base_clientes": "Base de Clientes",
    "modulos_cliente": "Módulos por Cliente",
    "scorecard": "Scorecard Diagnóstico",
    "problemas": "Matriz de Problemas",
    "acciones": "Plan de Acciones",
    "finanzas": "Análisis Financiero",
    "supply": "Supply Chain",
    "comercial": "Estructura Comercial",
    "catalogos": "Catálogos",
}
MODULE_SHEETS = [f"M{i:02d}" for i in range(1, 21)]


# ---------- UTILIDADES ----------
def fmt_money(value) -> str:
    if pd.isna(value):
        return "—"
    return f"€ {value:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(value, digits: int = 1) -> str:
    if pd.isna(value):
        return "—"
    return f"{value*100:.{digits}f}%"


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() if c is not None else f"col_{i}" for i, c in enumerate(df.columns)]
    return df


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    df = df.dropna(how="all")
    return df


def radar_color(semaforo: str) -> str:
    semaforo = (semaforo or "").lower()
    if "verde" in semaforo:
        return "#16a34a"
    if "rojo" in semaforo:
        return "#dc2626"
    return "#f59e0b"


def metric_card(label: str, value: str, delta: str | None = None):
    st.markdown(
        f"""
        <div style="padding:16px;border:1px solid rgba(128,128,128,.18);border-radius:18px;background:linear-gradient(180deg,rgba(255,255,255,.02),rgba(255,255,255,.01));">
            <div style="font-size:0.9rem;color:#94a3b8;margin-bottom:6px;">{label}</div>
            <div style="font-size:1.8rem;font-weight:700;line-height:1.1;">{value}</div>
            {f'<div style="margin-top:6px;color:#60a5fa;font-size:0.9rem;">{delta}</div>' if delta else ''}
        </div>
        """,
        unsafe_allow_html=True,
    )


def safe_table(rows_slice: list, columns: list[str]) -> pd.DataFrame:
    width = len(columns)
    normalized_rows = []
    for row in rows_slice:
        row = list(row) if isinstance(row, (list, tuple)) else [row]
        if len(row) < width:
            row = row + [None] * (width - len(row))
        else:
            row = row[:width]
        normalized_rows.append(row)
    return pd.DataFrame(normalized_rows, columns=columns).dropna(how="all")


@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes: bytes | None) -> dict:
    source = io.BytesIO(file_bytes) if file_bytes else DEFAULT_EXCEL_PATH
    xls = pd.ExcelFile(source, engine="openpyxl")

    data: dict[str, pd.DataFrame | dict | list] = {}
    for key, sheet_name in SHEETS_STANDARD.items():
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df = clean_df(df)
        data[key] = df

    data["dashboard_raw"] = clean_df(pd.read_excel(xls, sheet_name="Dashboard Ejecutivo", header=None))

    for sheet_name in ["Radar Cliente", "Informe Cliente", "Informe Inversor", "Reporte por Módulo", "Reporte Completo", "Gráficos y Reportes"]:
        data[sheet_name] = clean_df(pd.read_excel(xls, sheet_name=sheet_name, header=None))

    module_detail_rows = []
    module_names = {}
    module_tables = {}
    for sheet_name in MODULE_SHEETS:
        df_raw = clean_df(pd.read_excel(xls, sheet_name=sheet_name, header=None))
        rows = df_raw.where(pd.notna(df_raw), None).values.tolist()
        title = rows[0][0] if rows and rows[0] and rows[0][0] else sheet_name
        module_names[sheet_name] = str(title)

        kpi = safe_table(rows[4:10], ["Indicador", "Valor", "Fuente", "Comentario"])
        kpi = kpi[kpi["Indicador"].notna()] if not kpi.empty else kpi

        activation = safe_table(rows[12:], ["Cliente", "Activado", "Semáforo", "Acciones", "Problemas", "Ahorro"])
        activation = activation[activation["Cliente"].notna()] if not activation.empty else activation

        module_tables[sheet_name] = {
            "title": title,
            "kpi": kpi,
            "activation": activation,
            "cliente_base": rows[2][1] if len(rows) > 2 and len(rows[2]) > 1 else None,
        }

        for _, row in activation.iterrows():
            module_detail_rows.append(
                {
                    "Módulo": sheet_name,
                    "Nombre módulo": str(title).replace(f"{sheet_name} — ", ""),
                    "Cliente": row.get("Cliente"),
                    "Activado": row.get("Activado"),
                    "Semáforo": row.get("Semáforo"),
                    "Acciones": row.get("Acciones", 0) or 0,
                    "Problemas": row.get("Problemas", 0) or 0,
                    "Ahorro": row.get("Ahorro", 0) or 0,
                }
            )

    data["module_names"] = module_names
    data["module_tables"] = module_tables
    data["module_long"] = pd.DataFrame(module_detail_rows)
    data["all_sheet_names"] = xls.sheet_names
    return data


# ---------- INTERFAZ ----------
CUSTOM_CSS = """
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
[data-testid="stSidebar"] {border-right: 1px solid rgba(128,128,128,.14);}
.stMetric {border: 1px solid rgba(128,128,128,.18); border-radius: 18px; padding: 12px;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

st.title("BOIS — Sistema Operativo en Dashboard")
st.caption("Versión preparada para despliegue con link. Carga el Excel actual o sustituye el archivo por una nueva versión del BOIS.")

with st.sidebar:
    st.header("Control")
    uploaded = st.file_uploader("Sustituir Excel backend", type=["xlsx"], help="Opcional. Si no subes nada, la app usa el Excel incluido en el paquete.")
    file_bytes = uploaded.getvalue() if uploaded else None
    data = load_excel_data(file_bytes)

    base = data["base_clientes"].copy()
    available_states = sorted(base["Estado proyecto"].dropna().astype(str).unique().tolist())
    available_countries = sorted(base["País"].dropna().astype(str).unique().tolist())
    available_sectors = sorted(base["Sector"].dropna().astype(str).unique().tolist())

    page = st.radio(
        "Secciones",
        [
            "Resumen ejecutivo",
            "Clientes",
            "Finanzas",
            "Problemas y acciones",
            "Módulos",
            "Informe por cliente",
            "Explorador Excel",
        ],
    )

    st.divider()
    state_filter = st.multiselect("Estado proyecto", available_states, default=available_states)
    country_filter = st.multiselect("País", available_countries, default=available_countries)
    sector_filter = st.multiselect("Sector", available_sectors, default=available_sectors)


def filter_clients(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Estado proyecto" in out:
        out = out[out["Estado proyecto"].astype(str).isin(state_filter)]
    if "País" in out:
        out = out[out["País"].astype(str).isin(country_filter)]
    if "Sector" in out:
        out = out[out["Sector"].astype(str).isin(sector_filter)]
    return out


base = filter_clients(data["base_clientes"])
selected_clients = base["Marca/Nombre comercial"].dropna().astype(str).tolist()

modules_client = data["modulos_cliente"].copy()
scorecard = data["scorecard"].copy()
problemas = data["problemas"].copy()
acciones = data["acciones"].copy()
finanzas = data["finanzas"].copy()
supply = data["supply"].copy()
comercial = data["comercial"].copy()
module_long = data["module_long"].copy()

if selected_clients:
    modules_client = modules_client[modules_client["Cliente"].astype(str).isin(selected_clients)]
    scorecard = scorecard[scorecard["Cliente"].astype(str).isin(selected_clients)]
    problemas = problemas[problemas["Cliente"].astype(str).isin(selected_clients)]
    acciones = acciones[acciones["Cliente"].astype(str).isin(selected_clients)]
    finanzas = finanzas[finanzas["Cliente"].astype(str).isin(selected_clients)]
    supply = supply[supply["Cliente"].astype(str).isin(selected_clients)]
    comercial = comercial[comercial["Cliente"].astype(str).isin(selected_clients)]
    module_long = module_long[module_long["Cliente"].astype(str).isin(selected_clients)]


if page == "Resumen ejecutivo":
    total_clientes = len(base)
    activos = int((base["Estado proyecto"].astype(str) == "Activo").sum()) if not base.empty else 0
    total_fact = pd.to_numeric(base["Facturación anual (€)"], errors="coerce").sum()
    total_ebitda = pd.to_numeric(base["EBITDA (€)"], errors="coerce").sum()
    score_prom = pd.to_numeric(scorecard["Score total"], errors="coerce").mean()
    ahorro_total = pd.to_numeric(acciones["Ahorro anual esperado (€)"], errors="coerce").sum()
    inversion_total = pd.to_numeric(acciones["Coste implementación (€)"], errors="coerce").sum()
    roi_prom = pd.to_numeric(finanzas["ROI cartera (%)"], errors="coerce").mean() / 100 if not finanzas.empty else np.nan

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Clientes visibles", str(total_clientes), f"{activos} activos")
    with c2:
        metric_card("Facturación cartera", fmt_money(total_fact), f"EBITDA: {fmt_money(total_ebitda)}")
    with c3:
        metric_card("Score operativo medio", f"{score_prom:.2f}" if pd.notna(score_prom) else "—", f"ROI medio: {fmt_pct(roi_prom)}")
    with c4:
        metric_card("Ahorro anual identificado", fmt_money(ahorro_total), f"Inversión: {fmt_money(inversion_total)}")

    st.divider()
    left, right = st.columns((1.35, 1))

    with left:
        ranking = finanzas[["Cliente", "EBITDA actual (€)", "Mejora EBITDA (€)", "Payback period (meses)"]].copy()
        ranking = ranking.sort_values("Mejora EBITDA (€)", ascending=False)
        fig = px.bar(ranking, x="Cliente", y="Mejora EBITDA (€)", color="Payback period (meses)", title="Impacto EBITDA por cliente", text_auto=".2s")
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

    with right:
        state_counts = base["Estado proyecto"].fillna("Sin estado").value_counts().reset_index()
        state_counts.columns = ["Estado", "Proyectos"]
        fig = px.pie(state_counts, names="Estado", values="Proyectos", title="Mix de estados de proyecto")
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

    left, right = st.columns(2)
    with left:
        sector_counts = base.groupby("Sector", dropna=False).size().reset_index(name="Clientes")
        fig = px.bar(sector_counts, x="Sector", y="Clientes", title="Clientes por sector", text_auto=True)
        fig.update_layout(height=360)
        st.plotly_chart(fig, use_container_width=True)
    with right:
        cat_counts = acciones.groupby("Categoría", dropna=False).size().reset_index(name="Acciones")
        fig = px.bar(cat_counts, x="Categoría", y="Acciones", title="Acciones por categoría", text_auto=True)
        fig.update_layout(height=360)
        st.plotly_chart(fig, use_container_width=True)

    pipeline_cols = ["Marca/Nombre comercial", "País", "Sector", "Facturación anual (€)", "EBITDA (€)", "Estado proyecto", "Módulos contratados"]
    st.subheader("Pipeline operativo")
    st.dataframe(base[pipeline_cols], use_container_width=True, hide_index=True)

elif page == "Clientes":
    st.subheader("Base de clientes")
    top_cols = st.columns([1.2, 1, 1, 1])
    with top_cols[0]:
        client_selected = st.selectbox("Cliente", ["Todos"] + selected_clients)
    with top_cols[1]:
        size_selected = st.selectbox("Tamaño", ["Todos"] + sorted(base["Tamaño empresa"].dropna().astype(str).unique().tolist()))
    with top_cols[2]:
        state_selected = st.selectbox("Estado", ["Todos"] + sorted(base["Estado proyecto"].dropna().astype(str).unique().tolist()))
    with top_cols[3]:
        sort_col = st.selectbox("Ordenar por", ["Facturación anual (€)", "EBITDA (€)", "Fecha última actualización"])

    client_df = base.copy()
    if client_selected != "Todos":
        client_df = client_df[client_df["Marca/Nombre comercial"] == client_selected]
    if size_selected != "Todos":
        client_df = client_df[client_df["Tamaño empresa"] == size_selected]
    if state_selected != "Todos":
        client_df = client_df[client_df["Estado proyecto"] == state_selected]
    client_df = client_df.sort_values(sort_col, ascending=False)

    st.dataframe(client_df, use_container_width=True, hide_index=True)
    st.download_button("Descargar tabla clientes CSV", client_df.to_csv(index=False).encode("utf-8-sig"), file_name="clientes_filtrados.csv", mime="text/csv")

elif page == "Finanzas":
    st.subheader("Motor financiero")
    c1, c2 = st.columns((1.15, 1))
    with c1:
        fig = px.scatter(finanzas, x="Payback period (meses)", y="ROI cartera (%)", size="Mejora EBITDA (€)", color="Cliente", hover_name="Cliente", title="Payback vs ROI")
        fig.update_layout(height=450)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        summary = finanzas[["Cliente", "Facturación actual (€)", "EBITDA actual (€)", "Mejora EBITDA (€)", "ROI cartera (%)"]].copy()
        summary = summary.sort_values("Mejora EBITDA (€)", ascending=False)
        st.dataframe(summary, use_container_width=True, hide_index=True)

elif page == "Problemas y acciones":
    st.subheader("Backlog operativo")
    tab1, tab2 = st.tabs(["Problemas", "Acciones"])

    with tab1:
        cols = st.columns(4)
        urgency = cols[0].multiselect("Urgencia", sorted(problemas["Urgencia"].dropna().unique().tolist()), default=sorted(problemas["Urgencia"].dropna().unique().tolist()))
        estado_prob = cols[1].multiselect("Estado problema", sorted(problemas["Estado"].dropna().astype(str).unique().tolist()), default=sorted(problemas["Estado"].dropna().astype(str).unique().tolist()))
        modulo_prob = cols[2].multiselect("Módulo", sorted(problemas["Módulo"].dropna().astype(str).unique().tolist()), default=sorted(problemas["Módulo"].dropna().astype(str).unique().tolist()))
        responsable_prob = cols[3].multiselect("Responsable", sorted(problemas["Responsable asignado"].dropna().astype(str).unique().tolist()), default=sorted(problemas["Responsable asignado"].dropna().astype(str).unique().tolist()))
        prob_df = problemas.copy()
        prob_df = prob_df[prob_df["Urgencia"].isin(urgency) & prob_df["Estado"].astype(str).isin(estado_prob) & prob_df["Módulo"].astype(str).isin(modulo_prob) & prob_df["Responsable asignado"].astype(str).isin(responsable_prob)]
        st.dataframe(prob_df, use_container_width=True, hide_index=True)

    with tab2:
        cols = st.columns(4)
        estado_acc = cols[0].multiselect("Estado acción", sorted(acciones["Estado"].dropna().astype(str).unique().tolist()), default=sorted(acciones["Estado"].dropna().astype(str).unique().tolist()))
        cat_acc = cols[1].multiselect("Categoría", sorted(acciones["Categoría"].dropna().astype(str).unique().tolist()), default=sorted(acciones["Categoría"].dropna().astype(str).unique().tolist()))
        modulo_acc = cols[2].multiselect("Módulo relacionado", sorted(acciones["Módulo relacionado"].dropna().astype(str).unique().tolist()), default=sorted(acciones["Módulo relacionado"].dropna().astype(str).unique().tolist()))
        responsable_acc = cols[3].multiselect("Responsable", sorted(acciones["Responsable"].dropna().astype(str).unique().tolist()), default=sorted(acciones["Responsable"].dropna().astype(str).unique().tolist()))
        act_df = acciones.copy()
        act_df = act_df[act_df["Estado"].astype(str).isin(estado_acc) & act_df["Categoría"].astype(str).isin(cat_acc) & act_df["Módulo relacionado"].astype(str).isin(modulo_acc) & act_df["Responsable"].astype(str).isin(responsable_acc)]
        st.dataframe(act_df, use_container_width=True, hide_index=True)

elif page == "Módulos":
    st.subheader("Sistema modular")
    tab1, tab2 = st.tabs(["Matriz de activación", "Detalle de módulo"])

    with tab1:
        heatmap_df = modules_client.copy()
        module_cols = [c for c in heatmap_df.columns if c.startswith("M") and len(c) == 3]
        if not heatmap_df.empty and module_cols:
            matrix = heatmap_df.set_index("Cliente")[module_cols].replace({"Sí": 1, "No": 0})
            fig = px.imshow(matrix, aspect="auto", text_auto=True, title="Módulos activos por cliente", color_continuous_scale=[[0, "#0f172a"], [1, "#38bdf8"]])
            fig.update_layout(height=520)
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(modules_client, use_container_width=True, hide_index=True)

    with tab2:
        module_choice = st.selectbox("Selecciona módulo", MODULE_SHEETS, format_func=lambda x: data["module_names"].get(x, x))
        module_meta = data["module_tables"][module_choice]
        c1, c2 = st.columns((1, 1.2))
        with c1:
            st.markdown(f"### {module_meta['title']}")
            st.caption(f"Cliente base de la hoja: {module_meta['cliente_base']}")
            st.dataframe(module_meta["kpi"], use_container_width=True, hide_index=True)
        with c2:
            activation = module_meta["activation"].copy()
            st.dataframe(activation, use_container_width=True, hide_index=True)
            if not activation.empty:
                fig = px.bar(activation, x="Cliente", y="Ahorro", color="Semáforo", title="Ahorro potencial del módulo", text_auto=".2s")
                fig.update_layout(height=360)
                st.plotly_chart(fig, use_container_width=True)

elif page == "Explorador Excel":
    st.subheader("Explorador del workbook")
    sheet_choice = st.selectbox("Hoja", data["all_sheet_names"])
    if sheet_choice in data and isinstance(data[sheet_choice], pd.DataFrame):
        df_show = data[sheet_choice]
    elif sheet_choice in SHEETS_STANDARD.values():
        key = [k for k, v in SHEETS_STANDARD.items() if v == sheet_choice][0]
        df_show = data[key]
    elif sheet_choice in MODULE_SHEETS:
        df_show = data["module_tables"][sheet_choice]["activation"]
    else:
        df_show = clean_df(pd.read_excel(io.BytesIO(file_bytes) if file_bytes else DEFAULT_EXCEL_PATH, sheet_name=sheet_choice, header=None))
    st.dataframe(df_show, use_container_width=True, hide_index=True)
    st.info("Esta sección es la capa de auditoría: sirve para validar que el dashboard está leyendo el Excel correcto.")

st.divider()
with st.expander("Notas operativas"):
    st.markdown(
        """
        - Esta app usa el **Excel como backend**. Si sustituyes el fichero por una nueva versión guardada desde Excel, el dashboard se actualiza.
        - Si un Excel nuevo viene con fórmulas sin recalcular, los resultados pueden quedarse con valores antiguos. Solución: abrir, recalcular y guardar antes de subirlo.
        - La app está preparada para **despliegue por URL** en Streamlit Community Cloud o Render.
        """
    )
