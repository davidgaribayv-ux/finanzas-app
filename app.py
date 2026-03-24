import streamlit as st
import gspread
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import json

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Finanzas Personales · David & Lau",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CONSTANTES FINANCIERAS (Línea Base Abril)
# ─────────────────────────────────────────────
INGRESOS_FIJOS = 7_100.00
GASTOS_FIJOS   = 7_708.43
DEUDA_TC       = 2_997.74
INGRESO_EXT = [
    {"fecha": "27 Mar", "monto": 700.00,   "moneda": "EUR", "descripcion": "Pago freelance"},
    {"fecha": "15 Abr", "monto": 2_590.00, "moneda": "EUR", "descripcion": "Pago freelance"},
]

CATEGORIAS = [
    "Comida para casa",
    "Comida personal",
    "Deliverys",
    "Cosas de la casa",
    "Taxis / Transporte",
    "Pascal",
    "Gasto hormiga",
    "Otros",
]
METODOS   = ["Tarjeta de Crédito", "Efectivo / Débito"]
PERSONAS  = ["David", "Lau"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ─────────────────────────────────────────────
# CONEXIÓN A GOOGLE SHEETS
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner="Conectando a Google Sheets…")
def get_gspread_client():
    """Retorna un cliente gspread autenticado usando st.secrets."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def get_worksheet():
    """Abre la hoja de cálculo y devuelve el primer worksheet."""
    client = get_gspread_client()
    spreadsheet = client.open(st.secrets["spreadsheet"]["name"])
    return spreadsheet.sheet1


@st.cache_data(ttl=60, show_spinner="Cargando datos…")
def load_data() -> pd.DataFrame:
    """Lee todos los registros del Sheet y los devuelve como DataFrame."""
    ws = get_worksheet()
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=["Marca temporal", "FECHA", "¿Quién pagó?",
                                     "Monto", "Categoría", "Método", "Descripción"])
    df = pd.DataFrame(records)
    # Normalizar tipos
    df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
    df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce").fillna(0)
    df.sort_values("FECHA", ascending=False, inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def append_row(row: list):
    """Agrega una fila al final del Sheet."""
    ws = get_worksheet()
    ws.append_row(row, value_input_option="USER_ENTERED")
    # Invalidar caché para que el dashboard se refresque
    load_data.clear()


# ─────────────────────────────────────────────
# ESTILOS CSS PERSONALIZADOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* Sidebar */
    [data-testid="stSidebar"] { background-color: #0f172a; }
    [data-testid="stSidebar"] * { color: #e2e8f0 !important; }

    /* Tarjetas de métricas */
    [data-testid="metric-container"] {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 16px 20px;
    }
    [data-testid="stMetricValue"] { font-size: 1.6rem !important; font-weight: 700; }

    /* Encabezados de sección */
    .section-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin: 1.5rem 0 0.5rem;
        padding-bottom: 4px;
        border-bottom: 1px solid #334155;
    }

    /* Banner ejecutivo */
    .exec-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%);
        border: 1px solid #1d4ed8;
        border-radius: 14px;
        padding: 20px 28px;
        margin-bottom: 1.5rem;
    }
    .exec-banner h2 { color: #93c5fd; margin: 0 0 4px; font-size: 1.2rem; }
    .exec-banner p  { color: #cbd5e1; margin: 0; font-size: 0.85rem; }

    /* Alerta de superávit/déficit */
    .deficit  { color: #f87171; font-weight: 700; }
    .superavit{ color: #4ade80; font-weight: 700; }

    /* Tabla */
    [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SIDEBAR — NAVEGACIÓN
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💰 Finanzas")
    st.markdown("### David & Lau")
    st.divider()
    pagina = st.radio(
        "Navegar a",
        ["📝 Ingresar Gasto", "📊 Dashboard"],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown(
        "<p style='font-size:0.75rem; color:#64748b;'>Datos sincronizados con<br>Google Sheets · Actualización cada 60s</p>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# BANNER RESUMEN EJECUTIVO (siempre visible)
# ─────────────────────────────────────────────
balance = INGRESOS_FIJOS - GASTOS_FIJOS
balance_class = "superavit" if balance >= 0 else "deficit"
balance_symbol = "+" if balance >= 0 else ""

st.markdown(f"""
<div class="exec-banner">
    <h2>📋 Resumen Ejecutivo — Abril 2025</h2>
    <p>Línea base mensual fija · Actualizar manualmente cuando cambie</p>
</div>
""", unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)
col1.metric("Ingresos Fijos", f"S/ {INGRESOS_FIJOS:,.2f}", help="Salarios y pagos recurrentes de abril")
col2.metric(
    "Gastos Fijos + Deudas",
    f"S/ {GASTOS_FIJOS:,.2f}",
    delta=f"{balance_symbol}S/ {abs(balance):,.2f}",
    delta_color="normal" if balance >= 0 else "inverse",
    help="Alquiler Av. La Paz + préstamo + nueva deuda",
)
col3.metric("Deuda Tarjeta Actual", f"S/ {DEUDA_TC:,.2f}", help="Se paga con ingreso del 27 de marzo")

with col4:
    st.markdown("**Ingresos Ext. Pendientes**")
    for ie in INGRESO_EXT:
        st.markdown(
            f"&nbsp;&nbsp;`{ie['fecha']}` — **€ {ie['monto']:,.2f}**  <small style='color:#64748b'>{ie['descripcion']}</small>",
            unsafe_allow_html=True,
        )

st.divider()


# ═════════════════════════════════════════════
# PÁGINA 1 — INGRESO DE GASTOS
# ═════════════════════════════════════════════
if pagina == "📝 Ingresar Gasto":
    st.header("📝 Registrar Nuevo Gasto")
    st.markdown("Completa el formulario y presiona **Guardar** para enviarlo a Google Sheets.")

    with st.form("form_gasto", clear_on_submit=True):
        c1, c2 = st.columns(2)
        fecha      = c1.date_input("Fecha", value=date.today())
        quien      = c2.selectbox("¿Quién pagó?", PERSONAS)

        c3, c4 = st.columns(2)
        monto      = c3.number_input("Monto (S/)", min_value=0.01, step=0.50, format="%.2f")
        metodo     = c4.selectbox("Método de pago", METODOS)

        categoria  = st.selectbox("Categoría", CATEGORIAS)
        descripcion = st.text_input("Descripción", placeholder="Ej: Mercado Metro, Uber a Miraflores…")

        submitted = st.form_submit_button("💾 Guardar Gasto", use_container_width=True, type="primary")

    if submitted:
        if monto <= 0:
            st.error("El monto debe ser mayor a 0.")
        else:
            timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            fecha_str = fecha.strftime("%d/%m/%Y")
            nueva_fila = [timestamp, fecha_str, quien, monto, categoria, metodo, descripcion]
            try:
                with st.spinner("Guardando en Google Sheets…"):
                    append_row(nueva_fila)
                st.success(f"✅ Gasto de **S/ {monto:.2f}** registrado correctamente por **{quien}**.")
                st.balloons()
            except Exception as e:
                st.error(f"Error al guardar: {e}")


# ═════════════════════════════════════════════
# PÁGINA 2 — DASHBOARD INTERACTIVO
# ═════════════════════════════════════════════
elif pagina == "📊 Dashboard":
    st.header("📊 Dashboard de Gastos")

    # Botón de recarga manual
    if st.button("🔄 Recargar datos", type="secondary"):
        load_data.clear()
        st.rerun()

    df = load_data()

    if df.empty:
        st.info("No hay datos registrados aún. Ingresa el primer gasto desde el formulario.")
        st.stop()

    # ── Filtro de Mes ──────────────────────────────────────────────────
    st.markdown('<p class="section-header">Filtros</p>', unsafe_allow_html=True)
    df_valid = df.dropna(subset=["FECHA"])
    meses_disponibles = (
        df_valid["FECHA"]
        .dt.to_period("M")
        .drop_duplicates()
        .sort_values(ascending=False)
    )
    opciones_mes = ["Todos"] + [str(m) for m in meses_disponibles]
    mes_sel = st.selectbox("Mes", opciones_mes, index=0)

    if mes_sel == "Todos":
        df_filtrado = df_valid.copy()
    else:
        df_filtrado = df_valid[df_valid["FECHA"].dt.to_period("M").astype(str) == mes_sel]

    # ── KPIs ───────────────────────────────────────────────────────────
    st.markdown('<p class="section-header">KPIs del Período</p>', unsafe_allow_html=True)

    mask_tc      = df_filtrado["Método"].str.contains("Tarjeta", case=False, na=False)
    mask_efectivo = ~mask_tc
    total_tc      = df_filtrado.loc[mask_tc, "Monto"].sum()
    total_efectivo = df_filtrado.loc[mask_efectivo, "Monto"].sum()
    total_general  = df_filtrado["Monto"].sum()

    gasto_david = df_filtrado.loc[df_filtrado["¿Quién pagó?"] == "David", "Monto"].sum()
    gasto_lau   = df_filtrado.loc[df_filtrado["¿Quién pagó?"] == "Lau",   "Monto"].sum()

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total Gastado",          f"S/ {total_general:,.2f}")
    k2.metric("Deuda Tarjeta de Crédito", f"S/ {total_tc:,.2f}",
              help="Suma de gastos pagados con Tarjeta de Crédito en el período seleccionado")
    k3.metric("Gasto Efectivo / Débito", f"S/ {total_efectivo:,.2f}")
    k4.metric("Gasto David",             f"S/ {gasto_david:,.2f}")
    k5.metric("Gasto Lau",               f"S/ {gasto_lau:,.2f}")

    # ── Gráficos ───────────────────────────────────────────────────────
    st.markdown('<p class="section-header">Análisis Visual</p>', unsafe_allow_html=True)
    gc1, gc2 = st.columns([1.2, 1])

    # Donut — Gastos por Categoría
    with gc1:
        cat_group = (
            df_filtrado.groupby("Categoría")["Monto"]
            .sum()
            .reset_index()
            .sort_values("Monto", ascending=False)
        )
        fig_donut = px.pie(
            cat_group,
            names="Categoría",
            values="Monto",
            hole=0.55,
            title="Distribución por Categoría",
            color_discrete_sequence=px.colors.qualitative.Bold,
        )
        fig_donut.update_traces(
            textposition="outside",
            textinfo="percent+label",
            hovertemplate="<b>%{label}</b><br>S/ %{value:,.2f}<br>%{percent}<extra></extra>",
        )
        fig_donut.update_layout(
            showlegend=False,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color="#e2e8f0",
            title_font_size=15,
            margin=dict(t=50, b=10, l=10, r=10),
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    # Barras — David vs Lau por Categoría
    with gc2:
        person_cat = (
            df_filtrado.groupby(["Categoría", "¿Quién pagó?"])["Monto"]
            .sum()
            .reset_index()
        )
        fig_bar = px.bar(
            person_cat,
            x="Monto",
            y="Categoría",
            color="¿Quién pagó?",
            orientation="h",
            title="David vs Lau por Categoría",
            barmode="group",
            color_discrete_map={"David": "#3b82f6", "Lau": "#ec4899"},
            labels={"Monto": "S/", "Categoría": ""},
        )
        fig_bar.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color="#e2e8f0",
            title_font_size=15,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=50, b=10, l=10, r=10),
            xaxis=dict(gridcolor="#1e293b"),
            yaxis=dict(gridcolor="rgba(0,0,0,0)"),
        )
        fig_bar.update_traces(
            hovertemplate="<b>%{y}</b><br>S/ %{x:,.2f}<extra></extra>"
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # Línea de tiempo de gastos acumulados
    if len(df_filtrado) > 1:
        st.markdown('<p class="section-header">Evolución Temporal</p>', unsafe_allow_html=True)
        timeline = (
            df_filtrado.sort_values("FECHA")
            .groupby("FECHA")["Monto"]
            .sum()
            .cumsum()
            .reset_index()
        )
        timeline.columns = ["Fecha", "Gasto Acumulado"]
        fig_line = px.area(
            timeline,
            x="Fecha",
            y="Gasto Acumulado",
            title="Gasto Acumulado en el Período",
            labels={"Gasto Acumulado": "S/"},
            color_discrete_sequence=["#3b82f6"],
        )
        fig_line.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color="#e2e8f0",
            title_font_size=15,
            xaxis=dict(gridcolor="#1e293b"),
            yaxis=dict(gridcolor="#1e293b"),
            margin=dict(t=40, b=10, l=10, r=10),
        )
        fig_line.update_traces(
            fillcolor="rgba(59,130,246,0.15)",
            hovertemplate="<b>%{x|%d %b}</b><br>Acumulado: S/ %{y:,.2f}<extra></extra>",
        )
        st.plotly_chart(fig_line, use_container_width=True)

    # ── Tabla de Historial ─────────────────────────────────────────────
    st.markdown('<p class="section-header">Historial de Gastos</p>', unsafe_allow_html=True)

    df_tabla = df_filtrado[[
        "FECHA", "¿Quién pagó?", "Monto", "Categoría", "Método", "Descripción"
    ]].copy()
    df_tabla["FECHA"] = df_tabla["FECHA"].dt.strftime("%d/%m/%Y")
    df_tabla["Monto"] = df_tabla["Monto"].apply(lambda x: f"S/ {x:,.2f}")

    st.dataframe(
        df_tabla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "FECHA":         st.column_config.TextColumn("Fecha",    width="small"),
            "¿Quién pagó?":  st.column_config.TextColumn("Quién",    width="small"),
            "Monto":         st.column_config.TextColumn("Monto",    width="small"),
            "Categoría":     st.column_config.TextColumn("Categoría"),
            "Método":        st.column_config.TextColumn("Método"),
            "Descripción":   st.column_config.TextColumn("Descripción", width="large"),
        },
        height=420,
    )

    # Botón de descarga CSV
    csv = df_filtrado.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Descargar CSV",
        data=csv,
        file_name=f"gastos_{mes_sel.replace(' ', '_')}.csv",
        mime="text/csv",
    )
