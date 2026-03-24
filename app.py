import streamlit as st
import gspread
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import calendar

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Finanzas · David & Lau",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

CATEGORIAS = ["Comida para casa","Comida personal","Deliverys","Cosas de la casa",
              "Taxis / Transporte","Pascal","Gasto hormiga","Otros"]
METODOS    = ["Tarjeta de Crédito","Efectivo / Débito"]
PERSONAS   = ["David","Lau"]
MONEDAS    = ["PEN","USD","EUR"]
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

# ─────────────────────────────────────────────
# CONEXIÓN
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_client():
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    return gspread.authorize(creds)

def get_sh():
    return get_client().open(st.secrets["spreadsheet"]["name"])

# ─────────────────────────────────────────────
# LOADERS
# ─────────────────────────────────────────────
@st.cache_data(ttl=30)
def load_config():
    try:
        rows = get_sh().worksheet("Configuracion").get_all_values()
        cfg = {"tipo_cambio_usd": 3.47, "tipo_cambio_eur": 3.95, "mes_activo": datetime.now().strftime("%Y-%m")}
        for r in rows[1:]:
            if len(r) >= 2 and r[0]:
                k, v = r[0].strip(), r[1].strip()
                if k in ("tipo_cambio_usd","tipo_cambio_eur"):
                    cfg[k] = float(v or 3.47)
                elif k == "mes_activo":
                    cfg[k] = v
        return cfg
    except:
        return {"tipo_cambio_usd":3.47,"tipo_cambio_eur":3.95,"mes_activo":datetime.now().strftime("%Y-%m")}

@st.cache_data(ttl=60)
def load_ingresos_fijos():
    try:
        return get_sh().worksheet("Ingresos_Fijos").get_all_records()
    except: return []

@st.cache_data(ttl=60)
def load_gastos_fijos():
    try:
        return get_sh().worksheet("Gastos_Fijos").get_all_records()
    except: return []

@st.cache_data(ttl=60)
def load_deudas():
    try:
        return get_sh().worksheet("Deudas").get_all_records()
    except: return []

@st.cache_data(ttl=60)
def load_ahorros():
    try:
        return get_sh().worksheet("Ahorros").get_all_records()
    except: return []

@st.cache_data(ttl=60)
def load_gastos_diarios():
    try:
        records = get_sh().sheet1.get_all_records()
        if not records: return pd.DataFrame()
        df = pd.DataFrame(records)
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce").fillna(0)
        return df.sort_values("FECHA", ascending=False).reset_index(drop=True)
    except: return pd.DataFrame()

@st.cache_data(ttl=30)
def load_checklist(mes: str):
    try:
        records = get_sh().worksheet("Checklist").get_all_records()
        return [r for r in records if r.get("mes","") == mes]
    except: return []

@st.cache_data(ttl=30)
def load_ingresos_variables():
    try:
        return get_sh().worksheet("Ingresos_Variables").get_all_records()
    except: return []

# ─────────────────────────────────────────────
# HELPERS DE CONVERSIÓN
# ─────────────────────────────────────────────
def to_pen(monto, moneda, usd, eur):
    if moneda == "USD": return float(monto) * usd
    if moneda == "EUR": return float(monto) * eur
    return float(monto)

# ─────────────────────────────────────────────
# ESCRITURAS EN SHEETS
# ─────────────────────────────────────────────
def guardar_config(usd, eur, mes):
    ws = get_sh().worksheet("Configuracion")
    ws.clear()
    ws.update(values=[
        ["clave","valor","descripcion"],
        ["tipo_cambio_usd", str(usd), "Tipo de cambio USD a PEN"],
        ["tipo_cambio_eur", str(eur), "Tipo de cambio EUR a PEN"],
        ["mes_activo", mes, "Mes activo YYYY-MM"],
    ], range_name="A1")
    load_config.clear()

def guardar_checklist_generado(mes, items):
    ws = get_sh().worksheet("Checklist")
    for item in items:
        ws.append_row([mes, item["nombre"], item["tipo"], item["monto"],
                       item["moneda"], "NO", "", ""], value_input_option="USER_ENTERED")
    load_checklist.clear()

def marcar_pago_checklist(mes, nombre, pagado: bool):
    ws = get_sh().worksheet("Checklist")
    records = ws.get_all_records()
    for i, r in enumerate(records, start=2):
        if r.get("mes") == mes and r.get("nombre") == nombre:
            estado = "SI" if pagado else "NO"
            fecha  = datetime.now().strftime("%d/%m/%Y") if pagado else ""
            ws.update(values=[[estado, fecha]], range_name=f"F{i}:G{i}")
            break
    load_checklist.clear()

def append_gasto(row):
    get_sh().sheet1.append_row(row, value_input_option="USER_ENTERED")
    load_gastos_diarios.clear()

def append_ingreso_variable(row):
    ws = get_sh().worksheet("Ingresos_Variables")
    ws.append_row(row, value_input_option="USER_ENTERED")
    load_ingresos_variables.clear()

def actualizar_cuotas_deuda(nombre, nuevas_cuotas):
    ws = get_sh().worksheet("Deudas")
    records = ws.get_all_records()
    for i, r in enumerate(records, start=2):
        if r.get("nombre") == nombre:
            ws.update(values=[[str(nuevas_cuotas)]], range_name=f"E{i}")
            break
    load_deudas.clear()

def guardar_sheet_completo(nombre_ws, headers, filas):
    ws = get_sh().worksheet(nombre_ws)
    ws.clear()
    ws.update(values=[headers] + filas, range_name="A1")
    ws.format(f"A1:{chr(64+len(headers))}1", {
        "textFormat": {"bold":True,"foregroundColor":{"red":1,"green":1,"blue":1}},
        "backgroundColor": {"red":0.07,"green":0.14,"blue":0.35},
    })

def actualizar_ahorros(soles, usd_amt, eur_amt):
    hoy = datetime.now().strftime("%d/%m/%Y")
    guardar_sheet_completo("Ahorros",
        ["tipo","monto","moneda","fecha_actualizacion","descripcion"],
        [["Efectivo/Cuenta", str(soles), "PEN", hoy, "Ahorros en soles"],
         ["Efectivo/Cuenta", str(usd_amt), "USD", hoy, "Ahorros en dólares"],
         ["Efectivo/Cuenta", str(eur_amt), "EUR", hoy, "Ahorros en euros"]])
    load_ahorros.clear()

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"]   { background-color:#0f172a; }
[data-testid="stSidebar"] * { color:#e2e8f0 !important; }
[data-testid="metric-container"] {
    background:#1e293b; border:1px solid #334155;
    border-radius:12px; padding:16px 20px;
}
[data-testid="stMetricValue"] { font-size:1.5rem !important; font-weight:700; }
.sec { font-size:1rem; font-weight:600; color:#94a3b8; text-transform:uppercase;
       letter-spacing:.08em; margin:1.5rem 0 .5rem; padding-bottom:4px;
       border-bottom:1px solid #334155; }
.card { background:#1e293b; border:1px solid #334155; border-radius:12px; padding:16px 20px; margin-bottom:12px; }
.pill-verde  { background:#14532d; color:#4ade80; border-radius:999px; padding:2px 10px; font-size:.8rem; font-weight:600; }
.pill-rojo   { background:#7f1d1d; color:#f87171; border-radius:999px; padding:2px 10px; font-size:.8rem; font-weight:600; }
.pill-gris   { background:#1e293b; color:#94a3b8; border-radius:999px; padding:2px 10px; font-size:.8rem; font-weight:600; }
.balance-pos { color:#4ade80; font-weight:700; font-size:1.4rem; }
.balance-neg { color:#f87171; font-weight:700; font-size:1.4rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💰 Finanzas")
    st.markdown("### David & Lau")
    st.divider()
    pagina = st.radio("", [
        "🏥 Situación Actual",
        "✅ Checklist del Mes",
        "📝 Registrar Gasto",
        "💰 Registrar Ingreso",
        "📊 Análisis",
        "⚙️ Configuración",
    ], label_visibility="collapsed")
    st.divider()
    if st.button("🔄 Recargar todo", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("<p style='font-size:.72rem;color:#475569;margin-top:8px'>Sincronizado con Google Sheets</p>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CARGA DE DATOS GLOBALES
# ─────────────────────────────────────────────
cfg   = load_config()
USD   = cfg["tipo_cambio_usd"]
EUR   = cfg["tipo_cambio_eur"]
MES   = cfg["mes_activo"]         # formato YYYY-MM
try:
    MES_LABEL = datetime.strptime(MES + "-01", "%Y-%m-%d").strftime("%B %Y").capitalize()
except:
    MES_LABEL = MES


# ══════════════════════════════════════════════════════
# PÁGINA 1 — SITUACIÓN ACTUAL
# ══════════════════════════════════════════════════════
if pagina == "🏥 Situación Actual":
    st.header(f"🏥 Situación Actual — {MES_LABEL}")

    ingresos_fijos   = load_ingresos_fijos()
    gastos_fijos     = load_gastos_fijos()
    deudas           = load_deudas()
    ahorros          = load_ahorros()
    gastos_df        = load_gastos_diarios()
    ingresos_var     = load_ingresos_variables()

    # ── Cálculos ──
    total_ing_fijos = sum(to_pen(r["monto"], r["moneda"], USD, EUR)
                          for r in ingresos_fijos if str(r.get("activo","")).upper()=="SI")

    total_gastos_fijos = sum(to_pen(r["monto"], r["moneda"], USD, EUR)
                             for r in gastos_fijos if str(r.get("activo","")).upper()=="SI")

    total_deudas_mes = sum(to_pen(r["cuota_mensual"], r["moneda"], USD, EUR)
                           for r in deudas if int(r.get("cuotas_restantes",0)) > 0)

    # Gastos variables del mes activo
    if not gastos_df.empty and "FECHA" in gastos_df.columns:
        mask_mes = gastos_df["FECHA"].dt.to_period("M").astype(str) == MES
        gastos_variables_mes = gastos_df.loc[mask_mes, "Monto"].sum()
    else:
        gastos_variables_mes = 0.0

    # Ingresos variables recibidos este mes
    ing_var_recibidos = sum(
        to_pen(r["monto"], r["moneda"], USD, EUR)
        for r in ingresos_var
        if str(r.get("recibido","")).upper()=="SI"
    )
    ing_var_pendientes = [r for r in ingresos_var if str(r.get("recibido","")).upper()=="NO"]

    total_ingresos   = total_ing_fijos + ing_var_recibidos
    total_egresos    = total_gastos_fijos + total_deudas_mes + gastos_variables_mes
    balance          = total_ingresos - total_egresos

    # Ahorros en PEN
    ahorros_pen = sum(to_pen(r["monto"], r["moneda"], USD, EUR) for r in ahorros)

    # ── Tipo de cambio banner ──
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("💵 USD → PEN", f"{USD:.2f}", help="Actualizable en ⚙️ Configuración")
    col_b.metric("💶 EUR → PEN", f"{EUR:.2f}", help="Actualizable en ⚙️ Configuración")
    col_c.metric("📅 Mes Activo", MES_LABEL)
    st.divider()

    # ── KPIs principales ──
    st.markdown('<p class="sec">📊 Resumen del Mes</p>', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Total Ingresos",         f"S/ {total_ingresos:,.2f}",
              delta=f"+S/ {ing_var_recibidos:,.2f} variable" if ing_var_recibidos else None)
    k2.metric("Total Egresos",          f"S/ {total_egresos:,.2f}")
    bal_delta = f"+S/ {balance:,.2f}" if balance >= 0 else f"S/ {balance:,.2f}"
    k3.metric("Balance Proyectado",     f"S/ {balance:,.2f}",
              delta=bal_delta, delta_color="normal" if balance >= 0 else "inverse")
    k4.metric("Ahorros Totales (PEN)",  f"S/ {ahorros_pen:,.2f}")

    # ── Desglose lado a lado ──
    st.markdown('<p class="sec">💵 Ingresos vs 💸 Egresos</p>', unsafe_allow_html=True)
    left, right = st.columns(2)

    with left:
        st.markdown("**Ingresos Fijos**")
        for r in ingresos_fijos:
            if str(r.get("activo","")).upper() == "SI":
                pen = to_pen(r["monto"], r["moneda"], USD, EUR)
                st.markdown(f"""<div class="card" style="padding:10px 16px">
                    <span style="color:#e2e8f0">{r['nombre']}</span>
                    <span style="float:right;color:#4ade80;font-weight:600">S/ {pen:,.2f}</span>
                </div>""", unsafe_allow_html=True)
        if ing_var_recibidos > 0:
            st.markdown(f"""<div class="card" style="padding:10px 16px;border-color:#3b82f6">
                <span style="color:#93c5fd">Ingresos Variables Recibidos</span>
                <span style="float:right;color:#4ade80;font-weight:600">S/ {ing_var_recibidos:,.2f}</span>
            </div>""", unsafe_allow_html=True)
        st.markdown(f"<div style='text-align:right;font-weight:700;color:#4ade80;font-size:1.1rem'>Total: S/ {total_ingresos:,.2f}</div>", unsafe_allow_html=True)

        if ing_var_pendientes:
            st.markdown('<p class="sec" style="margin-top:1.2rem">⏳ Ingresos Pendientes de Recibir</p>', unsafe_allow_html=True)
            for r in ing_var_pendientes:
                pen = to_pen(r["monto"], r["moneda"], USD, EUR)
                st.markdown(f"""<div class="card" style="padding:10px 16px;border-color:#f59e0b">
                    <span style="color:#fbbf24">{r['fuente']} · {r.get('fecha','')}</span>
                    <span style="float:right;color:#fbbf24;font-weight:600">{r['moneda']} {float(r['monto']):,.2f} ≈ S/ {pen:,.2f}</span>
                </div>""", unsafe_allow_html=True)

    with right:
        st.markdown("**Gastos Fijos**")
        for r in gastos_fijos:
            if str(r.get("activo","")).upper() == "SI":
                pen = to_pen(r["monto"], r["moneda"], USD, EUR)
                orig = f"{r['moneda']} {float(r['monto']):,.2f}" if r["moneda"] != "PEN" else ""
                st.markdown(f"""<div class="card" style="padding:10px 16px">
                    <span style="color:#e2e8f0">{r['nombre']}</span>
                    <span style="float:right;color:#f87171;font-weight:600">S/ {pen:,.2f}
                    {"<br><small style='color:#64748b'>"+orig+"</small>" if orig else ""}</span>
                </div>""", unsafe_allow_html=True)

        if total_deudas_mes > 0:
            st.markdown("**Cuotas de Deudas del Mes**")
            for r in deudas:
                cuotas = int(r.get("cuotas_restantes",0))
                if cuotas > 0:
                    pen = to_pen(r["cuota_mensual"], r["moneda"], USD, EUR)
                    orig = f"{r['moneda']} {float(r['cuota_mensual']):,.2f}" if r["moneda"] != "PEN" else ""
                    st.markdown(f"""<div class="card" style="padding:10px 16px;border-color:#f97316">
                        <span style="color:#fdba74">{r['nombre']}</span>
                        <span class="pill-gris" style="margin-left:6px">{cuotas} cuota{'s' if cuotas>1 else ''}</span>
                        <span style="float:right;color:#f97316;font-weight:600">S/ {pen:,.2f}
                        {"<br><small style='color:#64748b'>"+orig+"</small>" if orig else ""}</span>
                    </div>""", unsafe_allow_html=True)

        # Gastos Variables — siempre visible aunque sea 0
        st.markdown(f"""<div class="card" style="padding:10px 16px;border-color:#8b5cf6">
            <span style="color:#c4b5fd">🛒 Gastos Variables del Mes <small style="color:#64748b;font-size:.75rem">(registrados día a día)</small></span>
            <span style="float:right;color:#a78bfa;font-weight:700;font-size:1.05rem">S/ {gastos_variables_mes:,.2f}</span>
        </div>""", unsafe_allow_html=True)

        st.markdown(f"<div style='text-align:right;font-weight:700;color:#f87171;font-size:1.1rem'>Total: S/ {total_egresos:,.2f}</div>", unsafe_allow_html=True)

    # ── Gastos Variables del Mes — Desglose detallado ──
    st.markdown('<p class="sec">🛒 Desglose de Gastos Variables del Mes</p>', unsafe_allow_html=True)

    if not gastos_df.empty and "FECHA" in gastos_df.columns:
        mask_mes = gastos_df["FECHA"].dt.to_period("M").astype(str) == MES
        df_mes = gastos_df[mask_mes].copy()
    else:
        df_mes = pd.DataFrame()

    if df_mes.empty:
        st.info(f"📭 Aún no hay gastos registrados para **{MES_LABEL}**. Usa '📝 Registrar Gasto' para agregar.")
    else:
        gv1, gv2, gv3, gv4 = st.columns(4)
        gv1.metric("Total Variables", f"S/ {gastos_variables_mes:,.2f}")
        gv2.metric("N° de Gastos",    f"{len(df_mes)}")
        gv3.metric("Promedio Diario", f"S/ {gastos_variables_mes / max(df_mes['FECHA'].dt.day.max(),1):,.2f}")
        tc_var = df_mes.loc[df_mes["Método"].str.contains("Tarjeta", case=False, na=False), "Monto"].sum()
        gv4.metric("En Tarjeta",      f"S/ {tc_var:,.2f}")

        # Tabla por categoría
        by_cat = (df_mes.groupby("Categoría")["Monto"].sum()
                  .reset_index().sort_values("Monto", ascending=False))
        by_cat["% del Total"] = (by_cat["Monto"] / gastos_variables_mes * 100).round(1)
        by_cat["Monto"] = by_cat["Monto"].apply(lambda x: f"S/ {x:,.2f}")
        by_cat["% del Total"] = by_cat["% del Total"].apply(lambda x: f"{x}%")

        cv1, cv2 = st.columns([1, 1.4])
        with cv1:
            st.markdown("**Por Categoría**")
            st.dataframe(by_cat, use_container_width=True, hide_index=True,
                         column_config={
                             "Categoría": st.column_config.TextColumn("Categoría"),
                             "Monto":     st.column_config.TextColumn("Monto"),
                             "% del Total": st.column_config.TextColumn("%"),
                         })
        with cv2:
            st.markdown("**Últimos 10 gastos del mes**")
            df_tabla = df_mes[["FECHA","¿Quién pagó?","Monto","Categoría","Descripción"]].head(10).copy()
            df_tabla["FECHA"] = df_tabla["FECHA"].dt.strftime("%d/%m")
            df_tabla["Monto"] = df_tabla["Monto"].apply(lambda x: f"S/ {x:,.2f}")
            st.dataframe(df_tabla, use_container_width=True, hide_index=True,
                         column_config={
                             "FECHA":        st.column_config.TextColumn("Fecha", width="small"),
                             "¿Quién pagó?": st.column_config.TextColumn("Quién", width="small"),
                             "Monto":        st.column_config.TextColumn("Monto", width="small"),
                             "Descripción":  st.column_config.TextColumn("Descripción"),
                         })

    # ── Balance final ──
    st.divider()
    color_cls = "balance-pos" if balance >= 0 else "balance-neg"
    icono = "✅" if balance >= 0 else "🔴"
    st.markdown(f"""
    <div class="card" style="text-align:center;padding:20px;border-color:{'#4ade80' if balance>=0 else '#f87171'}">
        <p style="color:#94a3b8;margin:0;font-size:.9rem">BALANCE FINAL DEL MES</p>
        <p class="{color_cls}" style="margin:8px 0;font-size:2rem">{icono} S/ {balance:,.2f}</p>
        <p style="color:#64748b;margin:0;font-size:.8rem">
        Ingresos S/ {total_ingresos:,.2f} &nbsp;−&nbsp; Egresos S/ {total_egresos:,.2f}
        &nbsp;·&nbsp; Variables S/ {gastos_variables_mes:,.2f}</p>
    </div>""", unsafe_allow_html=True)

    # ── Ahorros ──
    st.markdown('<p class="sec">🏦 Ahorros Actuales</p>', unsafe_allow_html=True)
    sa1,sa2,sa3,sa4 = st.columns(4)
    soles_ah = next((to_pen(r["monto"],"PEN",USD,EUR) for r in ahorros if r.get("moneda")=="PEN"),0)
    usd_ah   = next((float(r["monto"]) for r in ahorros if r.get("moneda")=="USD"),0)
    eur_ah   = next((float(r["monto"]) for r in ahorros if r.get("moneda")=="EUR"),0)
    sa1.metric("Soles",   f"S/ {soles_ah:,.2f}")
    sa2.metric("Dólares", f"$ {usd_ah:,.2f}", f"≈ S/ {usd_ah*USD:,.2f}")
    sa3.metric("Euros",   f"€ {eur_ah:,.2f}", f"≈ S/ {eur_ah*EUR:,.2f}")
    sa4.metric("Total en PEN", f"S/ {ahorros_pen:,.2f}")


# ══════════════════════════════════════════════════════
# PÁGINA 2 — CHECKLIST DEL MES
# ══════════════════════════════════════════════════════
elif pagina == "✅ Checklist del Mes":
    st.header("✅ Checklist de Pagos del Mes")
    st.caption("Marca cada pago cuando lo hayas realizado. Se guarda en Google Sheets.")

    gastos_fijos = load_gastos_fijos()
    deudas       = load_deudas()

    checklist = load_checklist(MES)

    # Auto-generar si está vacío
    if not checklist:
        items_nuevos = []
        for r in gastos_fijos:
            if str(r.get("activo","")).upper() == "SI":
                items_nuevos.append({"nombre":r["nombre"],"tipo":"Gasto Fijo",
                                     "monto":r["monto"],"moneda":r["moneda"]})
        for r in deudas:
            if int(r.get("cuotas_restantes",0)) > 0:
                items_nuevos.append({"nombre":r["nombre"],"tipo":"Deuda",
                                     "monto":r["cuota_mensual"],"moneda":r["moneda"]})
        if items_nuevos:
            with st.spinner("Generando checklist del mes…"):
                guardar_checklist_generado(MES, items_nuevos)
            st.success(f"✅ Checklist de {MES_LABEL} generado con {len(items_nuevos)} ítems.")
            st.rerun()
        else:
            st.info("No hay gastos fijos ni deudas configuradas. Ve a ⚙️ Configuración.")
            st.stop()

    # Calcular progreso
    total_items   = len(checklist)
    pagados       = sum(1 for r in checklist if str(r.get("pagado","")).upper()=="SI")
    pendientes    = total_items - pagados
    monto_total   = sum(to_pen(r["monto"], r["moneda"], USD, EUR) for r in checklist)
    monto_pagado  = sum(to_pen(r["monto"], r["moneda"], USD, EUR)
                        for r in checklist if str(r.get("pagado","")).upper()=="SI")
    monto_pendiente = monto_total - monto_pagado

    # KPIs
    ck1,ck2,ck3,ck4 = st.columns(4)
    ck1.metric("Total a Pagar",    f"S/ {monto_total:,.2f}")
    ck2.metric("Ya Pagado",        f"S/ {monto_pagado:,.2f}", f"{pagados}/{total_items} ítems")
    ck3.metric("Pendiente",        f"S/ {monto_pendiente:,.2f}", f"{pendientes} ítems")
    pct = int((pagados/total_items)*100) if total_items else 0
    ck4.metric("Progreso",         f"{pct}%")

    st.progress(pct/100)
    st.divider()

    # Tabla de ítems
    st.markdown('<p class="sec">📋 Ítems del Mes</p>', unsafe_allow_html=True)

    for r in checklist:
        pagado = str(r.get("pagado","")).upper() == "SI"
        pen    = to_pen(r["monto"], r["moneda"], USD, EUR)
        orig   = f"({r['moneda']} {float(r['monto']):,.2f})" if r["moneda"] != "PEN" else ""
        tipo_color = "#3b82f6" if r["tipo"]=="Gasto Fijo" else "#f97316"
        icono = "✅" if pagado else "⬜"
        fecha_pago = r.get("fecha_pago","")

        col_icon, col_info, col_monto, col_btn = st.columns([0.5, 3.5, 1.5, 1.5])

        with col_icon:
            st.markdown(f"<div style='font-size:1.4rem;margin-top:8px'>{icono}</div>", unsafe_allow_html=True)
        with col_info:
            st.markdown(f"""
            <div style='margin-top:6px'>
                <span style='color:#e2e8f0;font-weight:600'>{r['nombre']}</span>
                <span style='color:{tipo_color};font-size:.75rem;margin-left:8px;
                background:{tipo_color}22;padding:2px 8px;border-radius:99px'>{r['tipo']}</span>
                {"<span style='color:#64748b;font-size:.75rem;margin-left:8px'>Pagado el "+fecha_pago+"</span>" if pagado and fecha_pago else ""}
            </div>""", unsafe_allow_html=True)
        with col_monto:
            st.markdown(f"<div style='margin-top:8px;text-align:right;color:{'#4ade80' if pagado else '#f87171'};font-weight:700'>S/ {pen:,.2f}<br><small style='color:#64748b'>{orig}</small></div>", unsafe_allow_html=True)
        with col_btn:
            if pagado:
                if st.button("↩️ Desmarcar", key=f"des_{r['nombre']}", use_container_width=True):
                    with st.spinner("Actualizando…"):
                        marcar_pago_checklist(MES, r["nombre"], False)
                    st.rerun()
            else:
                if st.button("✅ Pagado", key=f"pag_{r['nombre']}", type="primary", use_container_width=True):
                    with st.spinner("Guardando…"):
                        marcar_pago_checklist(MES, r["nombre"], True)
                    st.rerun()
        st.divider()


# ══════════════════════════════════════════════════════
# PÁGINA 3 — REGISTRAR GASTO
# ══════════════════════════════════════════════════════
elif pagina == "📝 Registrar Gasto":
    st.header("📝 Registrar Gasto Diario")

    with st.form("form_gasto", clear_on_submit=True):
        c1,c2 = st.columns(2)
        fecha  = c1.date_input("Fecha", value=date.today())
        quien  = c2.selectbox("¿Quién pagó?", PERSONAS)
        c3,c4 = st.columns(2)
        monto  = c3.number_input("Monto (S/)", min_value=0.01, step=0.50, format="%.2f")
        metodo = c4.selectbox("Método de pago", METODOS)
        categoria   = st.selectbox("Categoría", CATEGORIAS)
        descripcion = st.text_input("Descripción", placeholder="Ej: Mercado Metro, Uber…")
        submitted   = st.form_submit_button("💾 Guardar Gasto", use_container_width=True, type="primary")

    if submitted:
        if monto <= 0:
            st.error("El monto debe ser mayor a 0.")
        else:
            row = [datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                   fecha.strftime("%d/%m/%Y"), quien, monto, categoria, metodo, descripcion]
            try:
                with st.spinner("Guardando…"):
                    append_gasto(row)
                st.success(f"✅ Gasto de **S/ {monto:.2f}** registrado por **{quien}**.")
                st.balloons()
            except Exception as e:
                st.error(f"Error: {e}")


# ══════════════════════════════════════════════════════
# PÁGINA 4 — REGISTRAR INGRESO
# ══════════════════════════════════════════════════════
elif pagina == "💰 Registrar Ingreso":
    st.header("💰 Registrar Ingreso Recibido")
    st.caption("Usa esta página cuando recibas un pago variable o quieras confirmar un ingreso pendiente.")

    ingresos_var = load_ingresos_variables()
    pendientes   = [r for r in ingresos_var if str(r.get("recibido","")).upper()=="NO"]

    if pendientes:
        st.markdown('<p class="sec">⏳ Ingresos Pendientes — Marcar como recibido</p>', unsafe_allow_html=True)
        for r in pendientes:
            pen = to_pen(r["monto"], r["moneda"], USD, EUR)
            col1, col2, col3 = st.columns([3,2,1.5])
            col1.markdown(f"**{r['fuente']}** · {r.get('fecha','')} · {r.get('descripcion','')}")
            col2.markdown(f"**{r['moneda']} {float(r['monto']):,.2f}** ≈ S/ {pen:,.2f}")
            if col3.button("✅ Recibido", key=f"recv_{r['fuente']}_{r.get('fecha','')}", type="primary"):
                ws = get_sh().worksheet("Ingresos_Variables")
                all_records = ws.get_all_records()
                for i, rec in enumerate(all_records, start=2):
                    if rec.get("fuente")==r["fuente"] and rec.get("fecha")==r.get("fecha","") and str(rec.get("recibido","")).upper()=="NO":
                        ws.update(values=[["SI"]], range_name=f"G{i}")
                        break
                load_ingresos_variables.clear()
                st.success(f"✅ Ingreso de {r['moneda']} {float(r['monto']):,.2f} marcado como recibido.")
                st.rerun()
        st.divider()

    st.markdown('<p class="sec">➕ Registrar Nuevo Ingreso Variable</p>', unsafe_allow_html=True)
    with st.form("form_ingreso", clear_on_submit=True):
        c1,c2 = st.columns(2)
        fecha_i  = c1.date_input("Fecha de recepción", value=date.today())
        fuente   = c2.text_input("Fuente / Descripción", placeholder="Freelance, Bono, Venta…")
        c3,c4 = st.columns(2)
        monto_i  = c3.number_input("Monto", min_value=0.01, step=10.0, format="%.2f")
        moneda_i = c4.selectbox("Moneda", MONEDAS)
        ya_recibido = st.checkbox("Ya lo recibí (marcar como recibido ahora)", value=True)
        sub_ing  = st.form_submit_button("💾 Guardar Ingreso", use_container_width=True, type="primary")

    if sub_ing:
        if not fuente:
            st.error("Indica la fuente del ingreso.")
        elif monto_i <= 0:
            st.error("El monto debe ser mayor a 0.")
        else:
            row_i = [datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                     fecha_i.strftime("%d/%m/%Y"), fuente, monto_i, moneda_i,
                     fuente, "SI" if ya_recibido else "NO"]
            with st.spinner("Guardando…"):
                append_ingreso_variable(row_i)
            pen_eq = to_pen(monto_i, moneda_i, USD, EUR)
            st.success(f"✅ Ingreso de **{moneda_i} {monto_i:,.2f}** (≈ S/ {pen_eq:,.2f}) registrado.")
            st.balloons()


# ══════════════════════════════════════════════════════
# PÁGINA 5 — ANÁLISIS
# ══════════════════════════════════════════════════════
elif pagina == "📊 Análisis":
    st.header("📊 Análisis de Gastos")

    df = load_gastos_diarios()
    if df.empty:
        st.info("No hay gastos registrados aún.")
        st.stop()

    df_valid = df.dropna(subset=["FECHA"])
    meses    = ["Todos"] + [str(m) for m in sorted(df_valid["FECHA"].dt.to_period("M").unique(), reverse=True)]
    mes_sel  = st.selectbox("Mes", meses)
    df_f     = df_valid if mes_sel=="Todos" else df_valid[df_valid["FECHA"].dt.to_period("M").astype(str)==mes_sel]

    mask_tc = df_f["Método"].str.contains("Tarjeta", case=False, na=False)
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Total",            f"S/ {df_f['Monto'].sum():,.2f}")
    k2.metric("Tarjeta Crédito",  f"S/ {df_f.loc[mask_tc,'Monto'].sum():,.2f}")
    k3.metric("Efectivo/Débito",  f"S/ {df_f.loc[~mask_tc,'Monto'].sum():,.2f}")
    k4.metric("David",            f"S/ {df_f.loc[df_f['¿Quién pagó?']=='David','Monto'].sum():,.2f}")
    k5.metric("Lau",              f"S/ {df_f.loc[df_f['¿Quién pagó?']=='Lau','Monto'].sum():,.2f}")

    gc1,gc2 = st.columns([1.2,1])
    with gc1:
        cat_g = df_f.groupby("Categoría")["Monto"].sum().reset_index().sort_values("Monto",ascending=False)
        fig = px.pie(cat_g, names="Categoría", values="Monto", hole=0.55,
                     title="Por Categoría", color_discrete_sequence=px.colors.qualitative.Bold)
        fig.update_traces(textposition="outside", textinfo="percent+label")
        fig.update_layout(showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
                          plot_bgcolor="rgba(0,0,0,0)", font_color="#e2e8f0",
                          margin=dict(t=50,b=10,l=10,r=10))
        st.plotly_chart(fig, use_container_width=True)
    with gc2:
        pc = df_f.groupby(["Categoría","¿Quién pagó?"])["Monto"].sum().reset_index()
        fig2 = px.bar(pc, x="Monto", y="Categoría", color="¿Quién pagó?",
                      orientation="h", title="David vs Lau", barmode="group",
                      color_discrete_map={"David":"#3b82f6","Lau":"#ec4899"})
        fig2.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                           font_color="#e2e8f0", xaxis=dict(gridcolor="#1e293b"),
                           legend=dict(orientation="h",y=1.05), margin=dict(t=50,b=10,l=10,r=10))
        st.plotly_chart(fig2, use_container_width=True)

    if len(df_f) > 1:
        tl = df_f.sort_values("FECHA").groupby("FECHA")["Monto"].sum().cumsum().reset_index()
        tl.columns = ["Fecha","Acumulado"]
        fig3 = px.area(tl, x="Fecha", y="Acumulado", title="Gasto Acumulado",
                       color_discrete_sequence=["#3b82f6"])
        fig3.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                           font_color="#e2e8f0", xaxis=dict(gridcolor="#1e293b"),
                           yaxis=dict(gridcolor="#1e293b"), margin=dict(t=40,b=10,l=10,r=10))
        fig3.update_traces(fillcolor="rgba(59,130,246,0.15)")
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown('<p class="sec">Historial</p>', unsafe_allow_html=True)
    df_t = df_f[["FECHA","¿Quién pagó?","Monto","Categoría","Método","Descripción"]].copy()
    df_t["FECHA"] = df_t["FECHA"].dt.strftime("%d/%m/%Y")
    df_t["Monto"] = df_t["Monto"].apply(lambda x: f"S/ {x:,.2f}")
    st.dataframe(df_t, use_container_width=True, hide_index=True, height=400)
    st.download_button("⬇️ CSV", df_f.to_csv(index=False).encode(), f"gastos_{mes_sel}.csv","text/csv")


# ══════════════════════════════════════════════════════
# PÁGINA 6 — CONFIGURACIÓN
# ══════════════════════════════════════════════════════
elif pagina == "⚙️ Configuración":
    st.header("⚙️ Configuración del Sistema")

    PASSWORD = st.secrets.get("app_password","finanzas2025")
    if "auth" not in st.session_state:
        st.session_state.auth = False
    if not st.session_state.auth:
        with st.form("login"):
            pwd = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Entrar", use_container_width=True, type="primary"):
                if pwd == PASSWORD:
                    st.session_state.auth = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
        st.stop()

    st.success("✅ Acceso concedido")
    if st.button("🔒 Cerrar sesión"): st.session_state.auth=False; st.rerun()

    tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs([
        "💱 Tipo de Cambio","💵 Ingresos Fijos","💸 Gastos Fijos",
        "💳 Deudas","🏦 Ahorros","📅 Cerrar Mes"
    ])

    # ── Tab 1: Tipo de Cambio ──
    with tab1:
        st.markdown("Actualiza el tipo de cambio cuando cambie. Se usa para convertir USD y EUR a PEN en toda la app.")
        with st.form("fx_form"):
            c1,c2,c3 = st.columns(3)
            new_usd = c1.number_input("USD → PEN", value=float(USD), min_value=1.0, step=0.01, format="%.4f")
            new_eur = c2.number_input("EUR → PEN", value=float(EUR), min_value=1.0, step=0.01, format="%.4f")
            new_mes = c3.text_input("Mes Activo (YYYY-MM)", value=MES)
            if st.form_submit_button("💾 Guardar Tipo de Cambio", type="primary", use_container_width=True):
                guardar_config(new_usd, new_eur, new_mes)
                st.success("✅ Tipo de cambio actualizado.")
                st.rerun()

    # ── Tab 2: Ingresos Fijos ──
    with tab2:
        st.caption("Agrega o edita tus fuentes de ingreso fijo mensual.")
        ingresos_fijos = load_ingresos_fijos()
        if "if_edit" not in st.session_state:
            st.session_state.if_edit = [r.copy() for r in ingresos_fijos]
        to_del = []
        for i,r in enumerate(st.session_state.if_edit):
            c1,c2,c3,c4,c5 = st.columns([3,1.5,1.2,1.2,0.5])
            r["nombre"] = c1.text_input("Fuente", value=r.get("nombre",""), key=f"if_n{i}")
            r["monto"]  = c2.number_input("Monto", value=float(r.get("monto",0)), min_value=0.0, step=100.0, key=f"if_m{i}")
            r["moneda"] = c3.selectbox("Moneda", MONEDAS, index=MONEDAS.index(r.get("moneda","PEN")) if r.get("moneda","PEN") in MONEDAS else 0, key=f"if_mo{i}")
            r["activo"] = c4.selectbox("Activo", ["SI","NO"], index=0 if str(r.get("activo","SI")).upper()=="SI" else 1, key=f"if_a{i}")
            if c5.button("🗑️", key=f"if_del{i}"): to_del.append(i)
        for i in reversed(to_del): st.session_state.if_edit.pop(i); st.rerun()
        if st.button("➕ Agregar Ingreso"): st.session_state.if_edit.append({"nombre":"","monto":0,"moneda":"PEN","activo":"SI","descripcion":""}); st.rerun()
        if st.button("💾 Guardar Ingresos Fijos", type="primary", use_container_width=True):
            headers = ["nombre","monto","moneda","activo","descripcion"]
            filas = [[r["nombre"],r["monto"],r["moneda"],r["activo"],r.get("descripcion","")] for r in st.session_state.if_edit if r["nombre"]]
            guardar_sheet_completo("Ingresos_Fijos", headers, filas)
            load_ingresos_fijos.clear()
            del st.session_state["if_edit"]
            st.success("✅ Ingresos fijos guardados.")
            st.rerun()

    # ── Tab 3: Gastos Fijos ──
    with tab3:
        st.caption("Administra tus gastos fijos mensuales. Puedes desactivar temporalmente sin eliminar.")
        gastos_fijos = load_gastos_fijos()
        if "gf_edit" not in st.session_state:
            st.session_state.gf_edit = [r.copy() for r in gastos_fijos]
        to_del2 = []
        for i,r in enumerate(st.session_state.gf_edit):
            c1,c2,c3,c4,c5,c6 = st.columns([2.5,1.5,1.2,1.5,1.2,0.5])
            r["nombre"]   = c1.text_input("Nombre", value=r.get("nombre",""), key=f"gf_n{i}")
            r["monto"]    = c2.number_input("Monto", value=float(r.get("monto",0)), min_value=0.0, step=10.0, key=f"gf_m{i}")
            r["moneda"]   = c3.selectbox("Moneda", MONEDAS, index=MONEDAS.index(r.get("moneda","PEN")) if r.get("moneda","PEN") in MONEDAS else 0, key=f"gf_mo{i}")
            r["categoria"]= c4.text_input("Cat.", value=r.get("categoria",""), key=f"gf_c{i}")
            r["activo"]   = c5.selectbox("Activo", ["SI","NO"], index=0 if str(r.get("activo","SI")).upper()=="SI" else 1, key=f"gf_a{i}")
            if c6.button("🗑️", key=f"gf_del{i}"): to_del2.append(i)
        for i in reversed(to_del2): st.session_state.gf_edit.pop(i); st.rerun()
        if st.button("➕ Agregar Gasto Fijo"): st.session_state.gf_edit.append({"nombre":"","monto":0,"moneda":"PEN","categoria":"","activo":"SI","descripcion":""}); st.rerun()
        if st.button("💾 Guardar Gastos Fijos", type="primary", use_container_width=True):
            headers = ["nombre","monto","moneda","categoria","activo","descripcion"]
            filas = [[r["nombre"],r["monto"],r["moneda"],r.get("categoria",""),r["activo"],r.get("descripcion","")] for r in st.session_state.gf_edit if r["nombre"]]
            guardar_sheet_completo("Gastos_Fijos", headers, filas)
            load_gastos_fijos.clear()
            del st.session_state["gf_edit"]
            st.success("✅ Gastos fijos guardados.")
            st.rerun()

    # ── Tab 4: Deudas ──
    with tab4:
        st.caption("Administra tus deudas. Reduce 'Cuotas Restantes' cada mes que pagues.")
        deudas = load_deudas()
        if "deu_edit" not in st.session_state:
            st.session_state.deu_edit = [r.copy() for r in deudas]
        to_del3 = []
        for i,r in enumerate(st.session_state.deu_edit):
            pen = to_pen(r.get("cuota_mensual",0), r.get("moneda","PEN"), USD, EUR)
            st.markdown(f"**{r.get('nombre','Nueva Deuda')}** · Cuota: S/ {pen:,.2f} · Cuotas restantes: {r.get('cuotas_restantes',0)}")
            c1,c2,c3,c4,c5,c6 = st.columns([2.5,1.5,1.5,1.2,1.5,0.5])
            r["nombre"]          = c1.text_input("Nombre", value=r.get("nombre",""), key=f"d_n{i}")
            r["monto_total"]     = c2.number_input("Total", value=float(r.get("monto_total",0)), min_value=0.0, step=100.0, key=f"d_t{i}")
            r["cuota_mensual"]   = c3.number_input("Cuota", value=float(r.get("cuota_mensual",0)), min_value=0.0, step=50.0, key=f"d_c{i}")
            r["moneda"]          = c4.selectbox("Moneda", MONEDAS, index=MONEDAS.index(r.get("moneda","PEN")) if r.get("moneda","PEN") in MONEDAS else 0, key=f"d_mo{i}")
            r["cuotas_restantes"]= c5.number_input("Cuotas Rest.", value=int(r.get("cuotas_restantes",0)), min_value=0, step=1, key=f"d_r{i}")
            if c6.button("🗑️", key=f"d_del{i}"): to_del3.append(i)
            st.divider()
        for i in reversed(to_del3): st.session_state.deu_edit.pop(i); st.rerun()
        if st.button("➕ Agregar Deuda"): st.session_state.deu_edit.append({"nombre":"","monto_total":0,"cuota_mensual":0,"moneda":"PEN","cuotas_restantes":1,"descripcion":""}); st.rerun()
        if st.button("💾 Guardar Deudas", type="primary", use_container_width=True):
            headers = ["nombre","monto_total","cuota_mensual","moneda","cuotas_restantes","descripcion"]
            filas = [[r["nombre"],r["monto_total"],r["cuota_mensual"],r["moneda"],r["cuotas_restantes"],r.get("descripcion","")] for r in st.session_state.deu_edit if r["nombre"]]
            guardar_sheet_completo("Deudas", headers, filas)
            load_deudas.clear()
            del st.session_state["deu_edit"]
            st.success("✅ Deudas guardadas.")
            st.rerun()

    # ── Tab 5: Ahorros ──
    with tab5:
        st.caption("Actualiza tus ahorros actuales. Hazlo cada vez que cambien.")
        ahorros = load_ahorros()
        soles_a = next((float(r["monto"]) for r in ahorros if r.get("moneda")=="PEN"), 0.0)
        usd_a   = next((float(r["monto"]) for r in ahorros if r.get("moneda")=="USD"), 0.0)
        eur_a   = next((float(r["monto"]) for r in ahorros if r.get("moneda")=="EUR"), 0.0)
        with st.form("ahorros_form"):
            c1,c2,c3 = st.columns(3)
            new_soles = c1.number_input("Soles (PEN)", value=soles_a, min_value=0.0, step=50.0, format="%.2f")
            new_usd   = c2.number_input("Dólares (USD)", value=usd_a, min_value=0.0, step=10.0, format="%.2f")
            new_eur   = c3.number_input("Euros (EUR)", value=eur_a, min_value=0.0, step=10.0, format="%.2f")
            total_pen_preview = new_soles + new_usd*USD + new_eur*EUR
            st.info(f"💡 Total equivalente en PEN: **S/ {total_pen_preview:,.2f}**")
            if st.form_submit_button("💾 Actualizar Ahorros", type="primary", use_container_width=True):
                actualizar_ahorros(new_soles, new_usd, new_eur)
                st.success("✅ Ahorros actualizados.")
                st.rerun()

    # ── Tab 6: Cerrar Mes ──
    with tab6:
        st.markdown("### 📅 Cerrar Mes y Avanzar al Siguiente")
        st.warning("""
        **¿Qué hace esta acción?**
        - Reduce en 1 las cuotas restantes de cada deuda activa
        - Elimina del sistema las deudas que llegan a 0 cuotas
        - Avanza el Mes Activo al siguiente mes
        - El checklist del nuevo mes se generará automáticamente
        """)
        nuevo_mes_dt = None
        try:
            dt = datetime.strptime(MES + "-01", "%Y-%m-%d")
            if dt.month == 12:
                nuevo_mes_dt = f"{dt.year+1}-01"
            else:
                nuevo_mes_dt = f"{dt.year}-{dt.month+1:02d}"
        except:
            nuevo_mes_dt = MES
        st.info(f"El mes activo pasará de **{MES_LABEL}** → **{nuevo_mes_dt}**")

        if st.button("🔒 Confirmar Cierre de Mes", type="primary", use_container_width=True):
            deudas_act = load_deudas()
            nuevas_filas = []
            for r in deudas_act:
                cuotas = int(r.get("cuotas_restantes",0))
                if cuotas > 1:
                    r["cuotas_restantes"] = cuotas - 1
                    nuevas_filas.append([r["nombre"],r["monto_total"],r["cuota_mensual"],
                                         r["moneda"],r["cuotas_restantes"],r.get("descripcion","")])
                elif cuotas == 1:
                    st.info(f"✅ Deuda **{r['nombre']}** finalizada y eliminada.")
            guardar_sheet_completo("Deudas",
                ["nombre","monto_total","cuota_mensual","moneda","cuotas_restantes","descripcion"],
                nuevas_filas)
            guardar_config(USD, EUR, nuevo_mes_dt)
            load_deudas.clear()
            load_config.clear()
            st.success(f"✅ Mes cerrado. Nuevo mes activo: {nuevo_mes_dt}")
            st.rerun()
