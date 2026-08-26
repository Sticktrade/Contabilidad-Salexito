import streamlit as st
import pandas as pd
import openpyxl
import re
import json
import base64
from datetime import datetime, date
import plotly.express as px
import io
from sqlalchemy import create_engine, text

# ReportLab para exportación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ESTILOS VISUALES
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control de Caja y Presupuestos - Papelería",
    page_icon="📚",
    layout="wide"
)

st.markdown("""
<style>
    :root, body, html { color-scheme: light !important; }
    html, body, [data-testid="stAppViewContainer"], .main, .stApp { background-color: #ffffff !important; color: #0f172a !important; }
    div[data-baseweb="popover"], div[data-baseweb="calendar"], div[data-baseweb="menu"], div[role="dialog"] {
        background-color: #ffffff !important; color: #0f172a !important; border: 1px solid #cbd5e1 !important;
    }
    .stTabs [data-baseweb="tab-list"] { background-color: #f1f5f9 !important; border-radius: 10px !important; padding: 6px !important; gap: 6px !important; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; border-radius: 8px !important; padding: 8px 16px !important; }
    .stTabs [data-baseweb="tab"][aria-selected="true"] { background-color: #0284c7 !important; border-color: #0284c7 !important; }
    .stTabs [data-baseweb="tab"][aria-selected="true"] [data-testid="stMarkdownContainer"] p { color: #ffffff !important; }
    label, [data-testid="stWidgetLabel"], label p { color: #0f172a !important; font-weight: 600 !important; }
    div[data-baseweb="select"] > div, div[data-baseweb="input"], input, textarea { background-color: #ffffff !important; color: #0f172a !important; border: 1px solid #cbd5e1 !important; border-radius: 8px !important; }
    button[kind="primary"], div.stButton > button, button[data-testid="stFormSubmitButton"] > button { background-color: #10b981 !important; color: #ffffff !important; border: none !important; border-radius: 8px !important; font-weight: 600 !important; }
    .header-banner { background: linear-gradient(135deg, #e0f2fe 0%, #dcfce7 100%); padding: 20px; border-radius: 12px; border: 1px solid #bae6fd; margin-bottom: 20px; }
    .header-banner h1 { color: #0369a1 !important; margin: 0; font-size: 24px; font-weight: 700; }
    .header-banner p { color: #047857 !important; margin: 4px 0 0 0; font-size: 14px; }
    .metric-card-green { background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px; padding: 12px; text-align: center; }
    .metric-card-blue { background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 10px; padding: 12px; text-align: center; }
    .metric-card-amber { background-color: #fffbebf1; border: 1px solid #fde68a; border-radius: 10px; padding: 12px; text-align: center; }
    .metric-card-title { font-size: 13px; color: #475569 !important; font-weight: 600; }
    .metric-card-val { font-size: 20px; font-weight: 700; color: #0f172a !important; }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# CONEXIÓN A BASE DE DATOS (SUPABASE O SQLITE)
# -----------------------------------------------------------------------------
@st.cache_resource
def get_db_engine():
    if "SUPABASE_URL" in st.secrets and st.secrets["SUPABASE_URL"].strip():
        db_url = st.secrets["SUPABASE_URL"].strip()
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)
        try:
            eng = create_engine(db_url, pool_pre_ping=True, pool_size=5, max_overflow=10)
            with eng.connect() as conn:
                conn.execute(text("SELECT 1;"))
            return eng, "postgres", None
        except Exception as e:
            return create_engine("sqlite:///papeleria.db"), "sqlite", str(e)
    else:
        return create_engine("sqlite:///papeleria.db"), "sqlite", "Sin configurar SUPABASE_URL en Secrets."

engine, db_type, conn_error = get_db_engine()

if conn_error and "SUPABASE_URL" in st.secrets:
    st.error(f"⚠️ Error al conectar a Supabase: {conn_error}")

@st.cache_resource
def init_db():
    try:
        with engine.begin() as conn:
            if db_type == "postgres":
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS cierres_diarios (
                        fecha VARCHAR(20) PRIMARY KEY,
                        efectivo FLOAT,
                        nequi FLOAT,
                        daviplata FLOAT,
                        banco FLOAT,
                        total FLOAT,
                        observaciones TEXT
                    );
                    CREATE TABLE IF NOT EXISTS transacciones_diarias (
                        id SERIAL PRIMARY KEY,
                        fecha VARCHAR(20),
                        concepto TEXT,
                        cantidad INT,
                        codigo INT,
                        ingreso FLOAT,
                        gastos FLOAT,
                        saldo FLOAT
                    );
                    CREATE TABLE IF NOT EXISTS gastos_mensuales (
                        id SERIAL PRIMARY KEY,
                        fecha VARCHAR(20),
                        concepto TEXT,
                        monto FLOAT,
                        categoria VARCHAR(100),
                        comprobante TEXT
                    );
                    CREATE TABLE IF NOT EXISTS presupuestos_mensuales (
                        id SERIAL PRIMARY KEY,
                        mes_año VARCHAR(20),
                        concepto TEXT,
                        monto_inicial FLOAT,
                        monto_final FLOAT,
                        categoria VARCHAR(100)
                    );
                """))
            else:
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS cierres_diarios (
                        fecha TEXT PRIMARY KEY,
                        efectivo REAL,
                        nequi REAL,
                        daviplata REAL,
                        banco REAL,
                        total REAL,
                        observaciones TEXT
                    );
                    CREATE TABLE IF NOT EXISTS transacciones_diarias (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        fecha TEXT,
                        concepto TEXT,
                        cantidad INTEGER,
                        codigo INTEGER,
                        ingreso REAL,
                        gastos REAL,
                        saldo REAL
                    );
                    CREATE TABLE IF NOT EXISTS gastos_mensuales (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        fecha TEXT,
                        concepto TEXT,
                        monto REAL,
                        categoria TEXT,
                        comprobante TEXT
                    );
                    CREATE TABLE IF NOT EXISTS presupuestos_mensuales (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        mes_año TEXT,
                        concepto TEXT,
                        monto_inicial REAL,
                        monto_final REAL,
                        categoria TEXT
                    );
                """))
    except Exception as ex:
        st.warning(f"Error inicializando las tablas de base de datos: {ex}")

init_db()

# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES Y CONSULTAS CON CACHE
# -----------------------------------------------------------------------------
@st.cache_data(ttl=60)
def load_cierres_diarios():
    with engine.connect() as conn:
        return pd.read_sql_query(text("SELECT * FROM cierres_diarios ORDER BY fecha ASC"), conn)

@st.cache_data(ttl=60)
def load_transacciones_diarias(fecha=None):
    with engine.connect() as conn:
        if fecha:
            return pd.read_sql_query(text("SELECT * FROM transacciones_diarias WHERE fecha = :f ORDER BY id ASC"), conn, params={"f": fecha})
        return pd.read_sql_query(text("SELECT * FROM transacciones_diarias ORDER BY fecha ASC, id ASC"), conn)

@st.cache_data(ttl=60)
def load_gastos_ligeros():
    with engine.connect() as conn:
        return pd.read_sql_query(text("""
            SELECT id, fecha, concepto, monto, categoria, 
                   CASE WHEN comprobante IS NOT NULL AND length(comprobante) > 5 THEN 1 ELSE 0 END as tiene_comprobante 
            FROM gastos_mensuales ORDER BY fecha ASC, id ASC
        """), conn)

@st.cache_data(ttl=60)
def load_comprobante_gasto(gasto_id):
    with engine.connect() as conn:
        df = pd.read_sql_query(text("SELECT comprobante FROM gastos_mensuales WHERE id = :gid"), conn, params={"gid": gasto_id})
        if not df.empty and df['comprobante'].iloc[0]:
            return df['comprobante'].iloc[0]
        return None

@st.cache_data(ttl=60)
def load_presupuestos_mensuales():
    with engine.connect() as conn:
        return pd.read_sql_query(text("SELECT id, mes_año, concepto, monto_inicial, monto_final, categoria FROM presupuestos_mensuales ORDER BY mes_año DESC, id ASC"), conn)

def parse_excel_control_caja(file_bytes):
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    
    fechas_encontradas = []
    for r in range(1, 10):
        for c in range(1, 10):
            val = ws.cell(row=r, column=c).value
            if val:
                val_str = str(val).strip()
                match = re.search(r'(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})', val_str)
                if match:
                    d, m, y = match.groups()
                    if len(y) == 2: y = "20" + y
                    fechas_encontradas.append(f"{int(y):04d}-{int(m):02d}-{int(d):02d}")
    
    fecha_cierre = fechas_encontradas[0] if fechas_encontradas else date.today().strftime('%Y-%m-%d')
    
    data_rows = []
    for r in range(1, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in row_vals):
            data_rows.append(row_vals)
            
    df_raw = pd.DataFrame(data_rows)
    
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v).upper() for v in row if v is not None])
        if "CONCEPTO" in row_str and ("INGRESO" in row_str or "INGRESOS" in row_str):
            header_idx = idx
            break
            
    items = []
    if header_idx is not None:
        df_items = df_raw.iloc[header_idx + 1:].copy()
        for _, row in df_items.iterrows():
            concepto = row.iloc[0] if len(row) > 0 else None
            if pd.notna(concepto) and str(concepto).strip():
                c_str = str(concepto).strip()
                if any(k in c_str.upper() for k in ["TOTAL", "SALDO FINAL", "RESUMEN", "EFECTIVO EN CAJA"]):
                    continue
                cant = row.iloc[1] if len(row) > 1 and pd.notna(row.iloc[1]) else 1
                cod = row.iloc[2] if len(row) > 2 and pd.notna(row.iloc[2]) else 0
                ing = row.iloc[3] if len(row) > 3 and pd.notna(row.iloc[3]) else 0.0
                gas = row.iloc[4] if len(row) > 4 and pd.notna(row.iloc[4]) else 0.0
                sal = row.iloc[5] if len(row) > 5 and pd.notna(row.iloc[5]) else 0.0
                
                try: ing = float(ing) if ing else 0.0
                except: ing = 0.0
                try: gas = float(gas) if gas else 0.0
                except: gas = 0.0
                try: sal = float(sal) if sal else 0.0
                except: sal = 0.0
                
                items.append({
                    "fecha": fecha_cierre,
                    "concepto": c_str,
                    "cantidad": int(cant) if str(cant).isdigit() else 1,
                    "codigo": int(cod) if str(cod).isdigit() else 0,
                    "ingreso": ing,
                    "gastos": gas,
                    "saldo": sal
                })
                
    efectivo, nequi, daviplata, banco = 0.0, 0.0, 0.0, 0.0
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str):
                v_upper = val.upper().strip()
                neighbor_val = ws.cell(row=r, column=c+1).value or ws.cell(row=r+1, column=c).value
                try: num_val = float(neighbor_val)
                except: num_val = 0.0
                
                if "EFECTIVO" in v_upper and num_val > 0: efectivo = num_val
                elif "NEQUI" in v_upper and num_val > 0: nequi = num_val
                elif "DAVIPLATA" in v_upper and num_val > 0: daviplata = num_val
                elif ("BANCO" in v_upper or "BANCOLOMBIA" in v_upper) and num_val > 0: banco = num_val

    return fecha_cierre, items, efectivo, nequi, daviplata, banco

# -----------------------------------------------------------------------------
# GENERACIÓN DE PDF PROFESIONAL DE PRESUPUESTO
# -----------------------------------------------------------------------------
@st.cache_data(ttl=300)
def generate_presupuesto_pdf(mes_nombre, df_pres):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#0369a1'), spaceAfter=4)
    subtitle_style = ParagraphStyle('DocSubtitle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#047857'), spaceAfter=14)
    h2_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor('#0f172a'), spaceBefore=10, spaceAfter=6)
    cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#1e293b'))
    cell_bold = ParagraphStyle('CellB', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#0f172a'))
    header_cell = ParagraphStyle('HCell', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.white)

    story.append(Paragraph(f"Formulario de Presupuesto Mensual — {mes_nombre}", title_style))
    story.append(Paragraph("Papelería — Planificación Financiera y Control de Gastos", subtitle_style))
    story.append(Spacer(1, 4))

    tot_ini = df_pres['monto_inicial'].sum() if not df_pres.empty else 0.0
    tot_fin = df_pres['monto_final'].sum() if not df_pres.empty else 0.0
    dif = tot_fin - tot_ini

    summary_data = [
        [Paragraph("Presupuestado (Valor Inicial)", header_cell), Paragraph("Ejecutado (Valor Final)", header_cell), Paragraph("Variación / Diferencia", header_cell)],
        [Paragraph(f"${tot_ini:,.0f}", cell_bold), Paragraph(f"${tot_fin:,.0f}", cell_bold), Paragraph(f"${dif:,.0f}", cell_bold)]
    ]
    t_summary = Table(summary_data, colWidths=[180, 180, 180])
    t_summary.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0284c7')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#f0f9ff')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#bae6fd')),
    ]))
    story.append(t_summary)
    story.append(Spacer(1, 12))

    story.append(Paragraph("Desglose de Campos y Rubros del Mes", h2_style))
    
    table_rows = [[
        Paragraph("Concepto / Rubro", header_cell),
        Paragraph("Categoría", header_cell),
        Paragraph("Valor Inicial ($)", header_cell),
        Paragraph("Valor Final ($)", header_cell),
        Paragraph("Diferencia ($)", header_cell)
    ]]

    for idx, r in df_pres.iterrows():
        mi = float(r['monto_inicial']) if pd.notna(r['monto_inicial']) else 0.0
        mf = float(r['monto_final']) if pd.notna(r['monto_final']) else 0.0
        d = mf - mi
        d_str = f"+${d:,.0f}" if d > 0 else (f"-${abs(d):,.0f}" if d < 0 else "$0")
        
        table_rows.append([
            Paragraph(str(r['concepto']), cell_style),
            Paragraph(str(r.get('categoria', 'General')), cell_style),
            Paragraph(f"${mi:,.0f}", cell_style),
            Paragraph(f"${mf:,.0f}", cell_style),
            Paragraph(d_str, cell_bold)
        ])

    table_rows.append([
        Paragraph("TOTAL", header_cell),
        Paragraph("", header_cell),
        Paragraph(f"${tot_ini:,.0f}", header_cell),
        Paragraph(f"${tot_fin:,.0f}", header_cell),
        Paragraph(f"${dif:,.0f}", header_cell)
    ])

    t_pres = Table(table_rows, colWidths=[160, 100, 95, 95, 90])
    t_pres.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#10b981')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#f0fdf4')]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#047857')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
    ]))
    story.append(t_pres)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# ESTRUCTURA PRINCIPAL DE LA APLICACIÓN DE STREAMLIT
# -----------------------------------------------------------------------------
st.markdown("""
<div class="header-banner">
    <h1>📚 Sistema de Control de Caja y Presupuestos Mensuales</h1>
    <p>Gestión y Administración Financiera Remota</p>
</div>
""", unsafe_allow_html=True)

tabs = st.tabs([
    "📥 Cargar Caja Diaria", 
    "📊 Recuento Mensual", 
    "💸 Gastos Mensuales", 
    "📋 Presupuestos Mensuales", 
    "📈 Métricas", 
    "⚙️ Edición / Histórico"
])

# -----------------------------------------------------------------------------
# PESTAÑA 1: CARGAR CAJA DIARIA
# -----------------------------------------------------------------------------
with tabs[0]:
    st.subheader("📥 Registro o Carga de Cierre Diario de Caja")
    st.write("Puedes subir directamente el archivo Excel del día o llenar el formulario manualmente.")
    
    col_u1, col_u2 = st.columns([1, 1])
    with col_u1:
        uploaded_file = st.file_uploader("📂 Subir Excel de Control de Caja Diaria (.xlsx)", type=["xlsx"])
        if uploaded_file is not None:
            file_bytes = uploaded_file.read()
            f_fecha, f_items, f_efec, f_neq, f_davi, f_banco = parse_excel_control_caja(file_bytes)
            f_total = f_efec + f_neq + f_davi + f_banco
            
            st.success(f"✅ Excel procesado correctamente para la fecha: **{f_fecha}**")
            st.info(f"📊 Totales extraídos: Efectivo: ${f_efec:,.0f} | Nequi: ${f_neq:,.0f} | Daviplata: ${f_davi:,.0f} | Banco: ${f_banco:,.0f} (Total: ${f_total:,.0f})")
            
            if st.button("🚀 Guardar Cierre de Excel en Base de Datos", type="primary"):
                with engine.begin() as conn:
                    conn.execute(text("DELETE FROM transacciones_diarias WHERE fecha = :f"), {"f": f_fecha})
                    for item in f_items:
                        conn.execute(text("""
                            INSERT INTO transacciones_diarias (fecha, concepto, cantidad, codigo, ingreso, gastos, saldo)
                            VALUES (:fecha, :concepto, :cantidad, :codigo, :ingreso, :gastos, :saldo)
                        """), item)
                    
                    conn.execute(text("""
                        INSERT INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                        VALUES (:f, :e, :n, :d, :b, :t, :o)
                        ON CONFLICT (fecha) DO UPDATE SET
                            efectivo = EXCLUDED.efectivo,
                            nequi = EXCLUDED.nequi,
                            daviplata = EXCLUDED.daviplata,
                            banco = EXCLUDED.banco,
                            total = EXCLUDED.total,
                            observaciones = EXCLUDED.observaciones;
                    """ if db_type == "postgres" else """
                        INSERT OR REPLACE INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                        VALUES (:f, :e, :n, :d, :b, :t, :o);
                    """), {"f": f_fecha, "e": f_efec, "n": f_neq, "d": f_davi, "b": f_banco, "t": f_total, "o": f"Cargado vía Excel {uploaded_file.name}"})
                
                st.cache_data.clear()
                st.success("🎉 Cierre cargado y guardado con éxito en la base de datos.")
                st.rerun()

    with col_u2:
        st.markdown("##### 📝 Formulario Manual de Cierre Diario")
        with st.form("form_cierre_manual", clear_on_submit=False):
            m_fecha = st.date_input("Fecha de Cierre:", date.today())
            c1, c2 = st.columns(2)
            with c1:
                m_efec = st.number_input("Efectivo ($):", min_value=0.0, step=1000.0)
                m_neq = st.number_input("Nequi ($):", min_value=0.0, step=1000.0)
            with c2:
                m_davi = st.number_input("Daviplata ($):", min_value=0.0, step=1000.0)
                m_banco = st.number_input("Banco ($):", min_value=0.0, step=1000.0)
            m_obs = st.text_area("Observaciones del día:")
            
            m_tot = m_efec + m_neq + m_davi + m_banco
            st.markdown(f"**Total del Cierre:** `${m_tot:,.0f}`")
            
            btn_manual = st.form_submit_button("💾 Guardar Cierre Manual", type="primary")
            if btn_manual:
                str_f = m_fecha.strftime('%Y-%m-%d')
                with engine.begin() as conn:
                    conn.execute(text("""
                        INSERT INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                        VALUES (:f, :e, :n, :d, :b, :t, :o)
                        ON CONFLICT (fecha) DO UPDATE SET
                            efectivo = EXCLUDED.efectivo,
                            nequi = EXCLUDED.nequi,
                            daviplata = EXCLUDED.daviplata,
                            banco = EXCLUDED.banco,
                            total = EXCLUDED.total,
                            observaciones = EXCLUDED.observaciones;
                    """ if db_type == "postgres" else """
                        INSERT OR REPLACE INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                        VALUES (:f, :e, :n, :d, :b, :t, :o);
                    """), {"f": str_f, "e": m_efec, "n": m_neq, "d": m_davi, "b": m_banco, "t": m_tot, "o": m_obs})
                st.cache_data.clear()
                st.success(f"✅ Cierre para el día {str_f} registrado correctamente.")
                st.rerun()

# -----------------------------------------------------------------------------
# PESTAÑA 2: RECUENTO MENSUAL DE CAJA
# -----------------------------------------------------------------------------
with tabs[1]:
    st.subheader("📊 Recuento y Consolidado Mensual de Caja")
    df_cierres = load_cierres_diarios()
    
    if df_cierres.empty:
        st.info("No hay datos de cierres registrados en el sistema.")
    else:
        df_cierres['fecha_dt'] = pd.to_datetime(df_cierres['fecha'])
        df_cierres['mes_año'] = df_cierres['fecha_dt'].dt.strftime('%Y-%m')
        
        meses_disp = sorted(df_cierres['mes_año'].unique().tolist(), reverse=True)
        mes_sel = st.selectbox("Selecciona el Mes a Consultar:", meses_disp, key="select_mes_recuento")
        
        df_m = df_cierres[df_cierres['mes_año'] == mes_sel].sort_values("fecha", ascending=True).copy()
        
        tot_efec = df_m['efectivo'].sum()
        tot_neq = df_m['nequi'].sum()
        tot_davi = df_m['daviplata'].sum()
        tot_banco = df_m['banco'].sum()
        tot_gen = df_m['total'].sum()
        
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💵 Efectivo</div><div class="metric-card-val">${tot_efec:,.0f}</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">📱 Nequi</div><div class="metric-card-val">${tot_neq:,.0f}</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="metric-card-amber"><div class="metric-card-title">🔴 Daviplata</div><div class="metric-card-val">${tot_davi:,.0f}</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">🏦 Banco</div><div class="metric-card-val">${tot_banco:,.0f}</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💰 Total General</div><div class="metric-card-val">${tot_gen:,.0f}</div></div>', unsafe_allow_html=True)
        
        st.write("")
        st.dataframe(
            df_m[['fecha', 'efectivo', 'nequi', 'daviplata', 'banco', 'total', 'observaciones']],
            use_container_width=True,
            column_config={
                "fecha": "Fecha",
                "efectivo": st.column_config.NumberColumn("Efectivo", format="$%d"),
                "nequi": st.column_config.NumberColumn("Nequi", format="$%d"),
                "daviplata": st.column_config.NumberColumn("Daviplata", format="$%d"),
                "banco": st.column_config.NumberColumn("Banco", format="$%d"),
                "total": st.column_config.NumberColumn("Total Día", format="$%d"),
                "observaciones": "Observaciones"
            }
        )

# -----------------------------------------------------------------------------
# PESTAÑA 3: GASTOS MENSUALES (ORDENADOS DE PRINCIPIO A FIN DE MES)
# -----------------------------------------------------------------------------
with tabs[2]:
    st.subheader("💸 Registro y Control de Gastos Mensuales")
    
    col_g1, col_g2 = st.columns([1, 1])
    with col_g1:
        st.markdown("##### ➕ Registrar Nuevo Gasto")
        with st.form("form_nuevo_gasto", clear_on_submit=True):
            g_fecha = st.date_input("Fecha del Gasto:", date.today())
            g_concepto = st.text_input("Concepto / Proveedor / Descripción:")
            g_monto = st.number_input("Monto ($):", min_value=0.0, step=1000.0)
            g_cat = st.selectbox("Categoría:", ["Mercancía / Inventario", "Servicios Públicos", "Nómina / Pagos", "Arriendo", "Mantenimiento / Equipos", "Impuestos", "Otros Gastos"])
            g_file = st.file_uploader("📷 Comprobante / Recibo (Imagen/PDF opcional)", type=["png", "jpg", "jpeg", "pdf"])
            
            btn_gasto = st.form_submit_button("💾 Guardar Gasto", type="primary")
            if btn_gasto and g_concepto.strip() and g_monto > 0:
                b64_comp = None
                if g_file is not None:
                    b64_comp = base64.b64encode(g_file.read()).decode('utf-8')
                    
                with engine.begin() as conn:
                    conn.execute(text("""
                        INSERT INTO gastos_mensuales (fecha, concepto, monto, categoria, comprobante)
                        VALUES (:f, :c, :m, :cat, :comp)
                    """), {"f": g_fecha.strftime('%Y-%m-%d'), "c": g_concepto.strip(), "m": g_monto, "cat": g_cat, "comp": b64_comp})
                st.cache_data.clear()
                st.success("✅ Gasto registrado correctamente.")
                st.rerun()

    with col_g2:
        st.markdown("##### 📋 Listado de Gastos Ordenados Chronológicamente")
        df_gastos_lig = load_gastos_ligeros()
        
        if df_gastos_lig.empty:
            st.info("No hay gastos registrados en el sistema.")
        else:
            # ORDEN DE PRINCIPIO A FIN DE MES (ASCENDENTE)
            df_gastos_lig['fecha_dt'] = pd.to_datetime(df_gastos_lig['fecha'])
            df_gastos_lig = df_gastos_lig.sort_values("fecha_dt", ascending=True).reset_index(drop=True)
            
            df_gastos_lig['mes_año'] = df_gastos_lig['fecha_dt'].dt.strftime('%Y-%m')
            meses_g = sorted(df_gastos_lig['mes_año'].unique().tolist(), reverse=True)
            mes_g_sel = st.selectbox("Filtrar Gastos por Mes:", meses_g, key="select_mes_gastos")
            
            df_g_m = df_gastos_lig[df_gastos_lig['mes_año'] == mes_g_sel].sort_values("fecha_dt", ascending=True)
            tot_g_m = df_g_m['monto'].sum()
            
            st.markdown(f"**Total de Gastos en {mes_g_sel}:** `${tot_g_m:,.0f}`")
            
            st.dataframe(
                df_g_m[['id', 'fecha', 'concepto', 'monto', 'categoria', 'tiene_comprobante']],
                use_container_width=True,
                column_config={
                    "id": "ID",
                    "fecha": "Fecha",
                    "concepto": "Concepto",
                    "monto": st.column_config.NumberColumn("Monto ($)", format="$%d"),
                    "categoria": "Categoría",
                    "tiene_comprobante": st.column_config.CheckboxColumn("¿Comprobante?")
                }
            )
            
            # Visor de comprobantes bajo demanda
            gastos_con_comp = df_g_m[df_g_m['tiene_comprobante'] == 1]['id'].tolist()
            if gastos_con_comp:
                st.markdown("##### 👁️ Ver Comprobante Adjunto")
                gasto_v_id = st.selectbox("Selecciona ID de Gasto para ver comprobante:", gastos_con_comp)
                if gasto_v_id:
                    comp_data = load_comprobante_gasto(gasto_v_id)
                    if comp_data:
                        try:
                            img_bytes = base64.b64decode(comp_data)
                            st.image(img_bytes, caption=f"Comprobante para Gasto ID #{gasto_v_id}", use_column_width=True)
                        except Exception as e:
                            st.error(f"Error cargando imagen de comprobante: {e}")

# -----------------------------------------------------------------------------
# PESTAÑA 4: PRESUPUESTOS MENSUALES (FORMULARIO, TABLA Y PDF)
# -----------------------------------------------------------------------------
with tabs[3]:
    st.subheader("📋 Presupuesto Mensual (Formulario y Control)")
    st.write("A principio de mes, llena el formulario con los conceptos y su **Valor Inicial**. Al final del mes, actualiza el **Valor Final** en la tabla interactiva para obtener el cálculo automático y descargar el reporte en PDF.")
    
    df_pres_all = load_presupuestos_mensuales()
    df_cierres_aux = load_cierres_diarios()
    meses_cierres = pd.to_datetime(df_cierres_aux['fecha']).dt.strftime('%Y-%m').unique() if not df_cierres_aux.empty else []
    meses_pres = df_pres_all['mes_año'].unique() if not df_pres_all.empty else []
    
    all_months = sorted(list(set(list(meses_pres) + list(meses_cierres) + [date.today().strftime('%Y-%m')])), reverse=True)
    
    col_p1, col_p2 = st.columns([1, 1])
    with col_p1:
        mes_sel_pres = st.selectbox("Selecciona o Crea el Mes del Presupuesto:", all_months, key="select_mes_presupuesto")
        
    df_p_mes = df_pres_all[df_pres_all['mes_año'] == mes_sel_pres].copy()
    
    # Formulario para ingresar nuevos rubros a principio de mes
    st.markdown("##### ➕ Formulario: Agregar Rubro al Presupuesto")
    with st.form("form_nuevo_rubro_presupuesto", clear_on_submit=True):
        col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
        with col_f1:
            input_concepto = st.text_input("Concepto / Rubro (ej. Renta, Servicios, Mercancía):")
        with col_f2:
            input_categoria = st.selectbox("Categoría:", ["Gastos Fijos", "Servicios", "Impuestos", "Deudas", "Personal", "Mercancía", "Inversión/Equipo", "Otro"])
        with col_f3:
            input_monto_ini = st.number_input("Valor Inicial ($):", min_value=0.0, step=10000.0)
            
        btn_add_rubro = st.form_submit_button("➕ Añadir al Presupuesto", type="primary")
        if btn_add_rubro and input_concepto.strip():
            with engine.begin() as conn:
                conn.execute(text("""
                    INSERT INTO presupuestos_mensuales (mes_año, concepto, monto_inicial, monto_final, categoria)
                    VALUES (:mes, :c, :mi, :mf, :cat);
                """), {
                    "mes": mes_sel_pres,
                    "c": input_concepto.strip().upper(),
                    "mi": float(input_monto_ini),
                    "mf": float(input_monto_ini),
                    "cat": input_categoria
                })
            st.cache_data.clear()
            st.success(f"✅ Rubro '{input_concepto.strip()}' añadido correctamente al presupuesto de {mes_sel_pres}.")
            st.rerun()

    st.divider()

    # Tabla interactiva y cálculo automático
    if df_p_mes.empty:
        st.info(f"Aún no hay conceptos registrados para el presupuesto de **{mes_sel_pres}**.")
        col_seed1, _ = st.columns([1, 1])
        with col_seed1:
            if st.button("✨ Cargar Plantilla Estándar para este Mes", type="primary"):
                template_items = [
                    ("RENTA", 1100000.0, 1100000.0, "Gastos Fijos"),
                    ("SERVICIOS (AGUA/LUZ/GAS)", 280000.0, 280000.0, "Servicios"),
                    ("INTERNET / TELÉFONO", 75000.0, 75000.0, "Servicios"),
                    ("PAGO DE IMPUESTOS", 250000.0, 250000.0, "Impuestos"),
                    ("ABONO DEUDA ANA", 160000.0, 160000.0, "Deudas"),
                    ("PAGO JUANCHO", 120000.0, 120000.0, "Personal"),
                    ("PAGO ANA", 80000.0, 80000.0, "Personal"),
                    ("MERCANCÍA / PROVEEDORES", 800000.0, 800000.0, "Mercancía"),
                    ("ABONO FOTOCOPIADORA", 150000.0, 150000.0, "Inversión/Equipo")
                ]
                with engine.begin() as conn:
                    for c_nom, m_ini, m_fin, c_cat in template_items:
                        conn.execute(text("""
                            INSERT INTO presupuestos_mensuales (mes_año, concepto, monto_inicial, monto_final, categoria)
                            VALUES (:mes, :c, :mi, :mf, :cat);
                        """), {"mes": mes_sel_pres, "c": c_nom, "mi": m_ini, "mf": m_fin, "cat": c_cat})
                st.cache_data.clear()
                st.success("✅ Plantilla cargada con éxito.")
                st.rerun()

    if not df_p_mes.empty:
        tot_ini = df_p_mes['monto_inicial'].sum()
        tot_fin = df_p_mes['monto_final'].sum()
        dif_tot = tot_fin - tot_ini
        
        pdf_pres_bytes = generate_presupuesto_pdf(mes_sel_pres, df_p_mes)
        with col_p2:
            st.write("")
            st.download_button(
                label="📄 Descargar Formulario de Presupuesto en PDF",
                data=pdf_pres_bytes,
                file_name=f"Presupuesto_{mes_sel_pres}.pdf",
                mime="application/pdf",
                type="primary"
            )

        st.markdown("##### 💰 Resumen Automático del Presupuesto Mensual:")
        kp1, kp2, kp3 = st.columns(3)
        kp1.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">📌 Total Presupuestado (Valor Inicial)</div><div class="metric-card-val">${tot_ini:,.0f}</div></div>', unsafe_allow_html=True)
        kp2.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💸 Total Ejecutado (Valor Final)</div><div class="metric-card-val">${tot_fin:,.0f}</div></div>', unsafe_allow_html=True)
        
        diff_card_class = "metric-card-green" if dif_tot <= 0 else "metric-card-amber"
        diff_prefix = f"-${abs(dif_tot):,.0f} (Ahorro)" if dif_tot < 0 else (f"+${dif_tot:,.0f} (Exceso)" if dif_tot > 0 else "$0 (En meta)")
        kp3.markdown(f'<div class="{diff_card_class}"><div class="metric-card-title">⚖️ Diferencia / Variación</div><div class="metric-card-val">{diff_prefix}</div></div>', unsafe_allow_html=True)

        st.write("")
        st.markdown("##### 📝 Tabla de Presupuesto (Llena el Valor Final al Cierre del Mes):")
        
        df_p_mes['diferencia'] = df_p_mes['monto_final'] - df_p_mes['monto_inicial']
        
        df_pres_edited = st.data_editor(
            df_p_mes[['id', 'concepto', 'categoria', 'monto_inicial', 'monto_final', 'diferencia']],
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "concepto": st.column_config.TextColumn("Concepto / Rubro", required=True),
                "categoria": st.column_config.SelectboxColumn("Categoría", options=["Gastos Fijos", "Servicios", "Impuestos", "Deudas", "Personal", "Mercancía", "Inversión/Equipo", "Otro"]),
                "monto_inicial": st.column_config.NumberColumn("Valor Inicial ($)", format="$%d", required=True),
                "monto_final": st.column_config.NumberColumn("Valor Final ($)", format="$%d", required=True),
                "diferencia": st.column_config.NumberColumn("Diferencia ($)", format="$%d", disabled=True)
            },
            key=f"presupuesto_editor_{mes_sel_pres}"
        )

        col_psave, _ = st.columns([1, 2])
        with col_psave:
            if st.button("💾 Guardar Cambios en Presupuesto", type="primary"):
                with engine.begin() as conn:
                    conn.execute(text("DELETE FROM presupuestos_mensuales WHERE mes_año = :mes"), {"mes": mes_sel_pres})
                    for _, r in df_pres_edited.iterrows():
                        c_text = str(r['concepto']).strip() if pd.notna(r['concepto']) else ''
                        if c_text:
                            m_ini = float(r['monto_inicial']) if pd.notna(r['monto_inicial']) else 0.0
                            m_fin = float(r['monto_final']) if pd.notna(r['monto_final']) else 0.0
                            cat_val = str(r['categoria']) if pd.notna(r['categoria']) else 'General'
                            conn.execute(text("""
                                INSERT INTO presupuestos_mensuales (mes_año, concepto, monto_inicial, monto_final, categoria)
                                VALUES (:mes, :c, :mi, :mf, :cat);
                            """), {"mes": mes_sel_pres, "c": c_text, "mi": m_ini, "mf": m_fin, "cat": cat_val})
                st.cache_data.clear()
                st.success("✅ ¡Presupuesto actualizado correctamente!")
                st.rerun()

        st.divider()
        st.markdown("##### 📊 Comparativo Gráfico: Presupuestado vs Real")
        df_melt = df_p_mes.melt(id_vars=['concepto'], value_vars=['monto_inicial', 'monto_final'], var_name='Tipo', value_name='Monto')
        df_melt['Tipo'] = df_melt['Tipo'].map({'monto_inicial': 'Valor Inicial (Proyectado)', 'monto_final': 'Valor Final (Ejecutado)'})
        
        fig_pres = px.bar(df_melt, x='concepto', y='Monto', color='Tipo', barmode='group', text_auto='.2s',
                          labels={'concepto': 'Concepto', 'Monto': 'Monto ($)'},
                          color_discrete_map={'Valor Inicial (Proyectado)': '#0284c7', 'Valor Final (Ejecutado)': '#10b981'})
        fig_pres.update_layout(plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#0f172a'), yaxis=dict(showgrid=True, gridcolor='#f1f5f9'))
        st.plotly_chart(fig_pres, use_container_width=True)

# -----------------------------------------------------------------------------
# PESTAÑA 5: MÉTRICAS Y GRÁFICOS
# -----------------------------------------------------------------------------
with tabs[4]:
    st.subheader("📈 Análisis de Tendencias y Métricas")
    df_cierres = load_cierres_diarios()
    df_gastos_lig = load_gastos_ligeros()
    
    if not df_cierres.empty:
        df_cierres['fecha_dt'] = pd.to_datetime(df_cierres['fecha'])
        df_cierres = df_cierres.sort_values("fecha_dt", ascending=True)
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown("##### 💵 Tendencia Diaria de Ingresos")
            fig_ing = px.line(df_cierres, x='fecha', y='total', title="Evolución de Ventas Diarias ($)", markers=True, color_discrete_sequence=['#10b981'])
            fig_ing.update_layout(plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#0f172a'))
            st.plotly_chart(fig_ing, use_container_width=True)
            
        with col_m2:
            st.markdown("##### 💳 Distribución de Canales de Pago")
            totales_canal = {
                'Efectivo': df_cierres['efectivo'].sum(),
                'Nequi': df_cierres['nequi'].sum(),
                'Daviplata': df_cierres['daviplata'].sum(),
                'Banco': df_cierres['banco'].sum()
            }
            df_canales = pd.DataFrame(list(totales_canal.items()), columns=['Canal', 'Monto'])
            fig_pie = px.pie(df_canales, values='Monto', names='Canal', title="Participación por Medio de Pago", color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_pie, use_container_width=True)

# -----------------------------------------------------------------------------
# PESTAÑA 6: EDICIÓN E HISTÓRICO
# -----------------------------------------------------------------------------
with tabs[5]:
    st.subheader("⚙️ Edición de Registros Históricos")
    df_cierres = load_cierres_diarios()
    
    if not df_cierres.empty:
        st.write("Modifica o elimina cierres ingresados anteriormente:")
        
        df_edited = st.data_editor(
            df_cierres,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "fecha": st.column_config.TextColumn("Fecha (YYYY-MM-DD)", disabled=True),
                "efectivo": st.column_config.NumberColumn("Efectivo", format="$%d"),
                "nequi": st.column_config.NumberColumn("Nequi", format="$%d"),
                "daviplata": st.column_config.NumberColumn("Daviplata", format="$%d"),
                "banco": st.column_config.NumberColumn("Banco", format="$%d"),
                "total": st.column_config.NumberColumn("Total", format="$%d"),
                "observaciones": "Observaciones"
            },
            key="editor_cierres_historico"
        )
        
        if st.button("💾 Guardar Cambios en Histórico", type="primary"):
            with engine.begin() as conn:
                for _, r in df_edited.iterrows():
                    tot_r = (r['efectivo'] or 0) + (r['nequi'] or 0) + (r['daviplata'] or 0) + (r['banco'] or 0)
                    conn.execute(text("""
                        UPDATE cierres_diarios
                        SET efectivo = :e, nequi = :n, daviplata = :d, banco = :b, total = :t, observaciones = :o
                        WHERE fecha = :f
                    """), {"f": r['fecha'], "e": r['efectivo'], "n": r['nequi'], "d": r['daviplata'], "b": r['banco'], "t": tot_r, "o": r['observaciones']})
            st.cache_data.clear()
            st.success("✅ Histórico de cierres actualizado.")
            st.rerun()
