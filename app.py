import streamlit as st
import gspread
import pandas as pd
import plotly.express as px
from google.oauth2.service_account import Credentials
from datetime import datetime, date

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
# CONSTANTES
# ─────────────────────────────────────────────
CATEGORIAS = [
    "Comida para casa", "Comida personal", "Deliverys",
    "Cosas de la casa", "Taxis / Transporte", "Pascal",
    "Gasto hormiga", "Otros",
]
METODOS  = ["Tarjeta de Crédito", "Efectivo / Débito"]
PERSONAS = ["David", "Lau"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ─────────────────────────────────────────────
# CONEXIÓN A GOOGLE SHEETS
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner="Conectando a Google Sheets…")
def get_gspread_client():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    return gspread.authorize(creds)


def get_spreadsheet():
    return get_gspread_client().open(st.secrets["spreadsheet"]["name"])


@st.cache_data(ttl=60, show_spinner="Cargando gastos…")
def load_data() -> pd.DataFrame:
    ws = get_spreadsheet().sheet1
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=["Marca temporal","FECHA","¿Quién pagó?","Monto","Categoría","Método","Descripción"])
    df = pd.DataFrame(records)
    df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
    df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce").fillna(0)
    df.sort_values("FECHA", ascending=False, inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


@st.cache_data(ttl=30, show_spinner="Cargando configuración…")
def load_config() -> dict:
    """Lee la pestaña Configuracion y devuelve un dict con todos los valores."""
    try:
        ws = get_spreadsheet().worksheet("Configuracion")
        rows = ws.get_all_values()
    except Exception:
        return _default_config()

    config = {
        "mes_referencia": "Abril 2025",
        "INGRESOS_FIJOS": 7100.00,
        "GASTOS_FIJOS":   7708.43,
        "DEUDA_TC":       2997.74,
        "INGRESO_EXT":    [],
    }
    for row in rows[1:]:  # saltar encabezado
        if len(row) < 2 or not row[0]:
            continue
        clave = row[0].strip()
        valor = row[1].strip() if row[1] else ""
        extra1 = row[2].strip() if len(row) > 2 else ""
        extra2 = row[3].strip() if len(row) > 3 else ""

        if clave == "mes_referencia":
            config["mes_referencia"] = valor
        elif clave == "INGRESOS_FIJOS":
            config["INGRESOS_FIJOS"] = float(valor or 0)
        elif clave == "GASTOS_FIJOS":
            config["GASTOS_FIJOS"] = float(valor or 0)
        elif clave == "DEUDA_TC":
            config["DEUDA_TC"] = float(valor or 0)
        elif clave == "INGRESO_EXT" and valor:
            config["INGRESO_EXT"].append({
                "monto": float(valor),
                "moneda": extra1 or "EUR",
                "fecha": extra2 or "",
            })
    return config


def _default_config() -> dict:
    return {
        "mes_referencia": "Abril 2025",
        "INGRESOS_FIJOS": 7100.00,
        "GASTOS_FIJOS":   7708.43,
        "DEUDA_TC":       2997.74,
        "INGRESO_EXT": [
            {"monto": 700.00,   "moneda": "EUR", "fecha": "27 Mar"},
            {"monto": 2590.00,  "moneda": "EUR", "fecha": "15 Abr"},
        ],
    }


def save_config(mes, ingresos, gastos, deuda, ingresos_ext: list):
    """Sobreescribe la pestaña Configuracion con los nuevos valores."""
    ws = get_spreadsheet().worksheet("Configuracion")
    ws.clear()
    data = [["clave","valor","extra1","extra2","descripcion"]]
    data.append(["mes_referencia", mes,          "", "", "Mes del presupuesto"])
    data.append(["INGRESOS_FIJOS", str(ingresos),"", "", "Salarios y pagos fijos"])
    data.append(["GASTOS_FIJOS",   str(gastos),  "", "", "Alquiler + préstamo + deudas fijas"])
    data.append(["DEUDA_TC",       str(deuda),   "", "", "Deuda tarjeta de crédito actual"])
    for ie in ingresos_ext:
        data.append(["INGRESO_EXT", str(ie["monto"]), ie["moneda"], ie["fecha"], "Ingreso extraordinario"])
    ws.update(values=data, range_name="A1")
    ws.format("A1:E1", {
        "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
        "backgroundColor": {"red":0.07,"green":0.14,"blue":0.35},
    })
    load_config.clear()


def append_gasto(row: list):
    get_spreadsheet().sheet1.append_row(row, value_input_option="USER_ENTERED")
    load_data.clear()


# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #0f172a; }
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
[data-testid="metric-container"] {
    background: #1e293b; border: 1px solid #334155;
    border-radius: 12px; padding: 16px 20px;
}
[data-testid="stMetricValue"] { font-size: 1.6rem !important; font-weight: 700; }
.section-header {
    font-size: 1.05rem; font-weight: 600; color: #94a3b8;
    text-transform: uppercase; letter-spacing: 0.08em;
    margin: 1.5rem 0 0.5rem; padding-bottom: 4px;
    border-bottom: 1px solid #334155;
}
.exec-banner {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%);
    border: 1px solid #1d4ed8; border-radius: 14px;
    padding: 20px 28px; margin-bottom: 1.5rem;
}
.exec-banner h2 { color: #93c5fd; margin: 0 0 4px; font-size: 1.2rem; }
.exec-banner p  { color: #cbd5e1; margin: 0; font-size: 0.85rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💰 Finanzas")
    st.markdown("### David & Lau")
    st.divider()
    pagina = st.radio(
        "Navegar a",
        ["📝 Ingresar Gasto", "📊 Dashboard", "⚙️ Configuración"],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown(
        "<p style='font-size:0.75rem;color:#64748b;'>Datos en Google Sheets<br>Actualización cada 60s</p>",
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────
# CARGAR CONFIGURACIÓN
# ─────────────────────────────────────────────
cfg = load_config()
INGRESOS_FIJOS = cfg["INGRESOS_FIJOS"]
GASTOS_FIJOS   = cfg["GASTOS_FIJOS"]
DEUDA_TC       = cfg["DEUDA_TC"]
INGRESO_EXT    = cfg["INGRESO_EXT"]
MES_REF        = cfg["mes_referencia"]

# ─────────────────────────────────────────────
# BANNER RESUMEN EJECUTIVO
# ─────────────────────────────────────────────
balance = INGRESOS_FIJOS - GASTOS_FIJOS
balance_symbol = "+" if balance >= 0 else ""

st.markdown(f"""
<div class="exec-banner">
    <h2>📋 Resumen Ejecutivo — {MES_REF}</h2>
    <p>Línea base mensual · Se actualiza desde ⚙️ Configuración</p>
</div>
""", unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)
col1.metric("Ingresos Fijos",        f"S/ {INGRESOS_FIJOS:,.2f}")
col2.metric("Gastos Fijos + Deudas", f"S/ {GASTOS_FIJOS:,.2f}",
    delta=f"{balance_symbol}S/ {abs(balance):,.2f}",
    delta_color="normal" if balance >= 0 else "inverse",
)
col3.metric("Deuda Tarjeta Actual",  f"S/ {DEUDA_TC:,.2f}")
with col4:
    st.markdown("**Ingresos Ext. Pendientes**")
    if INGRESO_EXT:
        for ie in INGRESO_EXT:
            st.markdown(
                f"&nbsp;&nbsp;`{ie['fecha']}` — **{ie['moneda']} {ie['monto']:,.2f}**",
                unsafe_allow_html=True,
            )
    else:
        st.markdown("<small style='color:#64748b'>Sin ingresos extraordinarios</small>", unsafe_allow_html=True)

st.divider()


# ═════════════════════════════════════════════
# PÁGINA 1 — INGRESO DE GASTOS
# ═════════════════════════════════════════════
if pagina == "📝 Ingresar Gasto":
    st.header("📝 Registrar Nuevo Gasto")

    with st.form("form_gasto", clear_on_submit=True):
        c1, c2 = st.columns(2)
        fecha  = c1.date_input("Fecha", value=date.today())
        quien  = c2.selectbox("¿Quién pagó?", PERSONAS)
        c3, c4 = st.columns(2)
        monto  = c3.number_input("Monto (S/)", min_value=0.01, step=0.50, format="%.2f")
        metodo = c4.selectbox("Método de pago", METODOS)
        categoria   = st.selectbox("Categoría", CATEGORIAS)
        descripcion = st.text_input("Descripción", placeholder="Ej: Mercado Metro, Uber a Miraflores…")
        submitted   = st.form_submit_button("💾 Guardar Gasto", use_container_width=True, type="primary")

    if submitted:
        if monto <= 0:
            st.error("El monto debe ser mayor a 0.")
        else:
            nueva_fila = [
                datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                fecha.strftime("%d/%m/%Y"),
                quien, monto, categoria, metodo, descripcion,
            ]
            try:
                with st.spinner("Guardando…"):
                    append_gasto(nueva_fila)
                st.success(f"✅ Gasto de **S/ {monto:.2f}** registrado por **{quien}**.")
                st.balloons()
            except Exception as e:
                st.error(f"Error: {e}")


# ═════════════════════════════════════════════
# PÁGINA 2 — DASHBOARD
# ═════════════════════════════════════════════
elif pagina == "📊 Dashboard":
    st.header("📊 Dashboard de Gastos")

    if st.button("🔄 Recargar datos"):
        load_data.clear()
        st.rerun()

    df = load_data()
    if df.empty:
        st.info("No hay datos aún. Ingresa el primer gasto.")
        st.stop()

    # Filtro mes
    st.markdown('<p class="section-header">Filtros</p>', unsafe_allow_html=True)
    df_valid = df.dropna(subset=["FECHA"])
    meses = ["Todos"] + [str(m) for m in sorted(df_valid["FECHA"].dt.to_period("M").unique(), reverse=True)]
    mes_sel = st.selectbox("Mes", meses)
    df_f = df_valid if mes_sel == "Todos" else df_valid[df_valid["FECHA"].dt.to_period("M").astype(str) == mes_sel]

    # KPIs
    st.markdown('<p class="section-header">KPIs del Período</p>', unsafe_allow_html=True)
    mask_tc = df_f["Método"].str.contains("Tarjeta", case=False, na=False)
    total_tc      = df_f.loc[mask_tc,  "Monto"].sum()
    total_efectivo = df_f.loc[~mask_tc, "Monto"].sum()
    total_general  = df_f["Monto"].sum()
    gasto_david    = df_f.loc[df_f["¿Quién pagó?"] == "David", "Monto"].sum()
    gasto_lau      = df_f.loc[df_f["¿Quién pagó?"] == "Lau",   "Monto"].sum()

    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Total Gastado",            f"S/ {total_general:,.2f}")
    k2.metric("Deuda Tarjeta de Crédito", f"S/ {total_tc:,.2f}")
    k3.metric("Efectivo / Débito",        f"S/ {total_efectivo:,.2f}")
    k4.metric("Gasto David",              f"S/ {gasto_david:,.2f}")
    k5.metric("Gasto Lau",                f"S/ {gasto_lau:,.2f}")

    # Gráficos
    st.markdown('<p class="section-header">Análisis Visual</p>', unsafe_allow_html=True)
    gc1, gc2 = st.columns([1.2, 1])

    with gc1:
        cat_group = df_f.groupby("Categoría")["Monto"].sum().reset_index().sort_values("Monto", ascending=False)
        fig_donut = px.pie(cat_group, names="Categoría", values="Monto", hole=0.55,
                           title="Distribución por Categoría",
                           color_discrete_sequence=px.colors.qualitative.Bold)
        fig_donut.update_traces(textposition="outside", textinfo="percent+label",
                                hovertemplate="<b>%{label}</b><br>S/ %{value:,.2f}<extra></extra>")
        fig_donut.update_layout(showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
                                plot_bgcolor="rgba(0,0,0,0)", font_color="#e2e8f0",
                                title_font_size=15, margin=dict(t=50,b=10,l=10,r=10))
        st.plotly_chart(fig_donut, use_container_width=True)

    with gc2:
        person_cat = df_f.groupby(["Categoría","¿Quién pagó?"])["Monto"].sum().reset_index()
        fig_bar = px.bar(person_cat, x="Monto", y="Categoría", color="¿Quién pagó?",
                         orientation="h", title="David vs Lau por Categoría", barmode="group",
                         color_discrete_map={"David":"#3b82f6","Lau":"#ec4899"},
                         labels={"Monto":"S/","Categoría":""})
        fig_bar.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                              font_color="#e2e8f0", title_font_size=15,
                              legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1),
                              margin=dict(t=50,b=10,l=10,r=10), xaxis=dict(gridcolor="#1e293b"))
        st.plotly_chart(fig_bar, use_container_width=True)

    if len(df_f) > 1:
        st.markdown('<p class="section-header">Evolución Temporal</p>', unsafe_allow_html=True)
        timeline = df_f.sort_values("FECHA").groupby("FECHA")["Monto"].sum().cumsum().reset_index()
        timeline.columns = ["Fecha","Gasto Acumulado"]
        fig_line = px.area(timeline, x="Fecha", y="Gasto Acumulado",
                           title="Gasto Acumulado en el Período",
                           color_discrete_sequence=["#3b82f6"])
        fig_line.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                               font_color="#e2e8f0", title_font_size=15,
                               xaxis=dict(gridcolor="#1e293b"), yaxis=dict(gridcolor="#1e293b"),
                               margin=dict(t=40,b=10,l=10,r=10))
        fig_line.update_traces(fillcolor="rgba(59,130,246,0.15)")
        st.plotly_chart(fig_line, use_container_width=True)

    # Tabla
    st.markdown('<p class="section-header">Historial de Gastos</p>', unsafe_allow_html=True)
    df_tabla = df_f[["FECHA","¿Quién pagó?","Monto","Categoría","Método","Descripción"]].copy()
    df_tabla["FECHA"] = df_tabla["FECHA"].dt.strftime("%d/%m/%Y")
    df_tabla["Monto"] = df_tabla["Monto"].apply(lambda x: f"S/ {x:,.2f}")
    st.dataframe(df_tabla, use_container_width=True, hide_index=True, height=420,
                 column_config={
                     "FECHA":        st.column_config.TextColumn("Fecha",   width="small"),
                     "¿Quién pagó?": st.column_config.TextColumn("Quién",   width="small"),
                     "Monto":        st.column_config.TextColumn("Monto",   width="small"),
                     "Descripción":  st.column_config.TextColumn("Descripción", width="large"),
                 })
    csv = df_f.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Descargar CSV", csv, f"gastos_{mes_sel}.csv", "text/csv")


# ═════════════════════════════════════════════
# PÁGINA 3 — CONFIGURACIÓN
# ═════════════════════════════════════════════
elif pagina == "⚙️ Configuración":
    st.header("⚙️ Configuración Financiera")
    st.markdown("Actualiza tus ingresos, gastos fijos y deudas. Los cambios se reflejan al instante en el banner y el dashboard.")

    # ── Protección con contraseña ──────────────────────────────────────
    PASSWORD = st.secrets.get("app_password", "finanzas2025")

    if "config_auth" not in st.session_state:
        st.session_state.config_auth = False

    if not st.session_state.config_auth:
        st.markdown('<p class="section-header">🔐 Acceso restringido</p>', unsafe_allow_html=True)
        with st.form("form_login"):
            pwd = st.text_input("Contraseña", type="password")
            login = st.form_submit_button("Entrar", use_container_width=True, type="primary")
        if login:
            if pwd == PASSWORD:
                st.session_state.config_auth = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

    # ── Formulario de configuración ────────────────────────────────────
    st.success("✅ Acceso concedido")
    if st.button("🔒 Cerrar sesión"):
        st.session_state.config_auth = False
        st.rerun()

    st.markdown('<p class="section-header">📅 Período de Referencia</p>', unsafe_allow_html=True)
    mes_input = st.text_input("Mes de referencia", value=MES_REF, placeholder="Ej: Mayo 2025")

    st.markdown('<p class="section-header">💵 Ingresos y Gastos Fijos</p>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    ing_input    = c1.number_input("Ingresos Fijos (S/)",        value=float(INGRESOS_FIJOS), min_value=0.0, step=100.0, format="%.2f")
    gastos_input = c2.number_input("Gastos Fijos + Deudas (S/)", value=float(GASTOS_FIJOS),   min_value=0.0, step=100.0, format="%.2f")

    st.markdown('<p class="section-header">💳 Deuda Tarjeta de Crédito</p>', unsafe_allow_html=True)
    deuda_input = st.number_input("Deuda Tarjeta Actual (S/)", value=float(DEUDA_TC), min_value=0.0, step=10.0, format="%.2f")

    st.markdown('<p class="section-header">🌍 Ingresos Extraordinarios</p>', unsafe_allow_html=True)
    st.caption("Agrega o elimina pagos extraordinarios pendientes (freelance, bonos, etc.)")

    # Estado de ingresos extraordinarios en sesión
    if "ingresos_ext_edit" not in st.session_state:
        st.session_state.ingresos_ext_edit = [ie.copy() for ie in INGRESO_EXT] if INGRESO_EXT else []

    # Mostrar los existentes
    to_delete = []
    for i, ie in enumerate(st.session_state.ingresos_ext_edit):
        ce1, ce2, ce3, ce4 = st.columns([2, 1.5, 1.5, 0.5])
        ie["fecha"]  = ce1.text_input(f"Fecha #{i+1}",   value=ie.get("fecha",""),  key=f"fecha_{i}")
        ie["monto"]  = ce2.number_input(f"Monto #{i+1}", value=float(ie.get("monto",0)), min_value=0.0, step=50.0, format="%.2f", key=f"monto_{i}")
        ie["moneda"] = ce3.selectbox(f"Moneda #{i+1}", ["EUR","USD","S/"], index=["EUR","USD","S/"].index(ie.get("moneda","EUR")), key=f"moneda_{i}")
        if ce4.button("🗑️", key=f"del_{i}", help="Eliminar"):
            to_delete.append(i)

    for i in reversed(to_delete):
        st.session_state.ingresos_ext_edit.pop(i)
        st.rerun()

    if st.button("➕ Agregar ingreso extraordinario"):
        st.session_state.ingresos_ext_edit.append({"fecha":"","monto":0.0,"moneda":"EUR"})
        st.rerun()

    st.divider()

    # Preview del nuevo balance
    nuevo_balance = ing_input - gastos_input
    color = "#4ade80" if nuevo_balance >= 0 else "#f87171"
    simbolo = "+" if nuevo_balance >= 0 else ""
    st.markdown(f"""
    **Preview del balance:** <span style='color:{color}; font-weight:700; font-size:1.1rem'>
    {simbolo}S/ {nuevo_balance:,.2f}</span>
    """, unsafe_allow_html=True)

    st.divider()

    if st.button("💾 Guardar Configuración", type="primary", use_container_width=True):
        ext_validos = [ie for ie in st.session_state.ingresos_ext_edit if ie["monto"] > 0]
        try:
            with st.spinner("Guardando en Google Sheets…"):
                save_config(
                    mes=mes_input,
                    ingresos=ing_input,
                    gastos=gastos_input,
                    deuda=deuda_input,
                    ingresos_ext=ext_validos,
                )
            del st.session_state["ingresos_ext_edit"]
            st.success("✅ Configuración guardada. El banner se actualizará en segundos.")
            st.rerun()
        except Exception as e:
            st.error(f"Error al guardar: {e}")
