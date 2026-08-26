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

# Generación de PDF profesional
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y FORZADO GLOBAL DE TEMA CLARO
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control de Caja y Presupuesto - Papelería",
    page_icon="📚",
    layout="wide"
)

st.markdown("""
<style>
    :root, body, html {
        color-scheme: light !important;
    }
    
    html, body, [data-testid="stAppViewContainer"], .main, .stApp {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }

    div[data-baseweb="popover"],
    div[data-baseweb="calendar"],
    div[data-baseweb="menu"],
    div[role="dialog"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1) !important;
    }

    div[data-baseweb="calendar"] * {
        color: #0f172a !important;
        background-color: transparent !important;
    }

    div[data-baseweb="calendar"] button {
        color: #0f172a !important;
        background-color: #ffffff !important;
    }

    div[data-baseweb="calendar"] button:hover {
        background-color: #e0f2fe !important;
        color: #0284c7 !important;
    }

    div[data-baseweb="calendar"] [aria-selected="true"] {
        background-color: #0284c7 !important;
        color: #ffffff !important;
    }

    div[data-testid="stDateInput"] input,
    div[data-testid="stDateInput"] > div {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        background-color: #f1f5f9 !important;
        border-radius: 10px !important;
        padding: 6px !important;
        gap: 6px !important;
        border: 1px solid #e2e8f0 !important;
    }

    .stTabs [data-baseweb="tab"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
    }

    .stTabs [data-baseweb="tab"] [data-testid="stMarkdownContainer"] p {
        color: #334155 !important;
        font-weight: 700 !important;
        font-size: 14px !important;
    }

    .stTabs [data-baseweb="tab"]:hover {
        background-color: #e0f2fe !important;
        border-color: #0284c7 !important;
    }

    .stTabs [data-baseweb="tab"]:hover [data-testid="stMarkdownContainer"] p {
        color: #0284c7 !important;
    }

    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background-color: #0284c7 !important;
        border-color: #0284c7 !important;
    }

    .stTabs [data-baseweb="tab"][aria-selected="true"] [data-testid="stMarkdownContainer"] p {
        color: #ffffff !important;
    }

    label, [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p, label p {
        color: #0f172a !important;
        font-weight: 600 !important;
        font-size: 14px !important;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="input"],
    input, textarea {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }

    div[data-baseweb="select"] span, div[data-baseweb="select"] div {
        color: #0f172a !important;
    }

    div[data-testid="stNumberInput"] button {
        background-color: #f1f5f9 !important;
        border: 1px solid #cbd5e1 !important;
    }

    div[data-testid="stNumberInput"] button * {
        color: #0f172a !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        background-color: #f8fafc !important;
        border: 2px dashed #0284c7 !important;
        border-radius: 12px !important;
    }

    [data-testid="stFileUploaderDropzone"] * {
        color: #334155 !important;
    }

    button[kind="primary"], 
    div.stButton > button, 
    button[data-testid="stFormSubmitButton"] > button,
    [data-testid="stFormSubmitButton"] button {
        background-color: #10b981 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }

    button[kind="primary"] *, 
    div.stButton > button *, 
    [data-testid="stFormSubmitButton"] button * {
        color: #ffffff !important;
    }

    .header-banner {
        background: linear-gradient(135deg, #e0f2fe 0%, #dcfce7 100%);
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #bae6fd;
        margin-bottom: 20px;
    }
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
# CONEXIÓN SEGURA A BASE DE DATOS Y CACHÉ ALTO RENDIMIENTO
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
    st.error(f"⚠️ Error al conectar a Supabase:\n\n{conn_error}")

@st.cache_resource
def init_db():
    try:
        with engine.begin() as conn:
            if db_type == "postgres":
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS cierres_diarios (
                        fecha VARCHAR(20) PRIMARY KEY,
                        efectivo FLOAT, nequi FLOAT, daviplata FLOAT, banco FLOAT, total FLOAT, observaciones TEXT
                    );
                """))
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS transacciones_diarias (
                        id SERIAL PRIMARY KEY,
                        fecha VARCHAR(20), concepto TEXT, cantidad INT, codigo INT, ingreso FLOAT, gastos FLOAT, saldo FLOAT
                    );
                """))
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS gastos_mensuales (
                        id SERIAL PRIMARY KEY,
                        fecha VARCHAR(20), concepto TEXT, monto FLOAT, categoria VARCHAR(100), comprobante TEXT
                    );
                """))
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS presupuestos_mensuales (
                        id SERIAL PRIMARY KEY,
                        mes_año VARCHAR(20), concepto TEXT, monto_inicial FLOAT, monto_final FLOAT, categoria VARCHAR(100)
                    );
                """))
            else:
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS cierres_diarios (
                        fecha TEXT PRIMARY KEY,
                        efectivo REAL, nequi REAL, daviplata REAL, banco REAL, total REAL, observaciones TEXT
                    );
                """))
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS transacciones_diarias (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        fecha TEXT, concepto TEXT, cantidad INTEGER, codigo INTEGER, ingreso REAL, gastos REAL, saldo REAL
                    );
                """))
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS gastos_mensuales (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        fecha TEXT, concepto TEXT, monto REAL, categoria TEXT, comprobante TEXT
                    );
                """))
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS presupuestos_mensuales (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        mes_año TEXT, concepto TEXT, monto_inicial REAL, monto_final REAL, categoria TEXT
                    );
                """))

            # Precarga inicial de datos de presupuestos históricos si la tabla está vacía
            res_pres = conn.execute(text("SELECT COUNT(*) FROM presupuestos_mensuales;")).fetchone()
            if res_pres and res_pres[0] == 0:
                hist_pres = [
                    # Febrero 2026
                    ("2026-02", "Renta", 1100000.0, 1100000.0, "Gastos Fijos"),
                    ("2026-02", "Servicios", 250000.0, 250000.0, "Servicios"),
                    ("2026-02", "Internet", 75000.0, 75000.0, "Servicios"),
                    ("2026-02", "Pago de impuestos cuota 2", 250000.0, 250000.0, "Impuestos"),
                    ("2026-02", "Abono deuda 5,5M Ana", 160000.0, 160000.0, "Deudas"),
                    ("2026-02", "Pago Juancho", 120000.0, 120000.0, "Personal"),
                    ("2026-02", "Pago Ana", 80000.0, 80000.0, "Personal"),
                    ("2026-02", "Mercancia", 1032585.0, 1032585.0, "Mercancía"),
                    ("2026-02", "Abono deuda Israel (Señora)", 100000.0, 100000.0, "Deudas"),
                    ("2026-02", "Fotocopiadora abono", 150000.0, 150000.0, "Inversión/Equipo"),
                    # Marzo 2026
                    ("2026-03", "Renta", 1100000.0, 1100000.0, "Gastos Fijos"),
                    ("2026-03", "Servicios", 280000.0, 280000.0, "Servicios"),
                    ("2026-03", "Internet", 75000.0, 75000.0, "Servicios"),
                    ("2026-03", "Pago de impuestos cuota 3", 250000.0, 250000.0, "Impuestos"),
                    ("2026-03", "Abono deuda 5,5M Ana", 160000.0, 160000.0, "Deudas"),
                    ("2026-03", "Deuda Alexis", 100000.0, 0.0, "Deudas"),
                    ("2026-03", "Deuda Salomon", 100000.0, 0.0, "Deudas"),
                    ("2026-03", "Pago Juancho", 120000.0, 120000.0, "Personal"),
                    ("2026-03", "Pago Ana", 80000.0, 80000.0, "Personal"),
                    ("2026-03", "Mercancia", 1000000.0, 1481400.0, "Mercancía"),
                    ("2026-03", "Abono deuda Israel (Señora)", 100000.0, 100000.0, "Deudas"),
                    ("2026-03", "Fotocopiadora abono", 200000.0, 200000.0, "Inversión/Equipo"),
                    ("2026-03", "Pago Servicios Israel", 0.0, 1632000.0, "Servicios"),
                    # Junio 2026
                    ("2026-06", "Renta", 1100000.0, 1100000.0, "Gastos Fijos"),
                    ("2026-06", "Servicios", 280000.0, 280000.0, "Servicios"),
                    ("2026-06", "Internet", 75000.0, 75000.0, "Servicios"),
                    ("2026-06", "Abono deuda 5,5M Ana", 160000.0, 160000.0, "Deudas"),
                    ("2026-06", "Pago Juancho", 150000.0, 150000.0, "Personal"),
                    ("2026-06", "Pago Ana", 100000.0, 80000.0, "Personal"),
                    ("2026-06", "Mercancia", 532344.0, 532344.0, "Mercancía"),
                    ("2026-06", "Fotocopiadora abono", 200000.0, 150000.0, "Inversión/Equipo")
                ]
                for m, c, mi, mf, cat in hist_pres:
                    conn.execute(text("""
                        INSERT INTO presupuestos_mensuales (mes_año, concepto, monto_inicial, monto_final, categoria)
                        VALUES (:m, :c, :mi, :mf, :cat);
                    """), {"m": m, "c": c, "mi": mi, "mf": mf, "cat": cat})
    except Exception as ex:
        st.warning(f"Error inicializando tablas: {ex}")

init_db()

# -----------------------------------------------------------------------------
# CONSULTAS OPTIMIZADAS CON CACHÉ
# -----------------------------------------------------------------------------
@st.cache_data(ttl=60)
def load_cierres_diarios():
    with engine.connect() as conn:
        return pd.read_sql_query(text("SELECT * FROM cierres_diarios ORDER BY fecha ASC"), conn)

@st.cache_data(ttl=60)
def load_transacciones_diarias():
    with engine.connect() as conn:
        return pd.read_sql_query(text("SELECT * FROM transacciones_diarias"), conn)

@st.cache_data(ttl=60)
def load_gastos_ligeros():
    """Carga los gastos ordenados cronológicamente (fecha ASC, id ASC) de principio a fin de mes"""
    with engine.connect() as conn:
        query = """
            SELECT id, fecha, concepto, monto, categoria, 
                   CASE WHEN comprobante IS NOT NULL AND length(comprobante) > 5 THEN 1 ELSE 0 END as tiene_comprobante
            FROM gastos_mensuales ORDER BY fecha ASC, id ASC
        """
        return pd.read_sql_query(text(query), conn)

def load_single_comprobante(gasto_id):
    with engine.connect() as conn:
        res = conn.execute(text("SELECT comprobante FROM gastos_mensuales WHERE id = :id"), {"id": gasto_id}).fetchone()
        return res[0] if res else None

@st.cache_data(ttl=60)
def load_presupuestos_mensuales():
    with engine.connect() as conn:
        return pd.read_sql_query(text("SELECT id, mes_año, concepto, monto_inicial, monto_final, categoria FROM presupuestos_mensuales ORDER BY mes_año DESC, id ASC"), conn)

# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES PARA COMPROBANTES Y MANEJO DE ARCHIVOS
# -----------------------------------------------------------------------------
def encode_file_to_json(uploaded_file):
    if uploaded_file is None:
        return ""
    file_bytes = uploaded_file.getvalue()
    b64_str = base64.b64encode(file_bytes).decode('utf-8')
    return json.dumps({
        "name": uploaded_file.name,
        "type": uploaded_file.type,
        "data": b64_str
    })

def decode_comprobante_json(comp_str):
    if not comp_str or not str(comp_str).startswith("{"):
        return None
    try:
        data = json.loads(comp_str)
        return {
            "name": data.get("name", "comprobante"),
            "type": data.get("type", "application/octet-stream"),
            "bytes": base64.b64decode(data.get("data", ""))
        }
    except Exception:
        return None

# -----------------------------------------------------------------------------
# GENERACIÓN DE REPORTES PDF
# -----------------------------------------------------------------------------
@st.cache_data(ttl=300)
def generate_pdf_report(mes_nombre, df_cierres_mes, df_gastos_mes, tot_efectivo, tot_nequi, tot_davi, tot_caja, tot_gastos, liquidez_neta):
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

    story.append(Paragraph(f"Recuento de Caja Mensual — {mes_nombre}", title_style))
    story.append(Paragraph("Papelería — Reporte Consolidado de Administración Remota", subtitle_style))
    story.append(Spacer(1, 4))

    summary_data = [
        [Paragraph("Efectivo", header_cell), Paragraph("Nequi", header_cell), Paragraph("Daviplata/Banco", header_cell), Paragraph("Total Caja", header_cell), Paragraph("Liquidez Neto", header_cell)],
        [Paragraph(f"${tot_efectivo:,.0f}", cell_bold), Paragraph(f"${tot_nequi:,.0f}", cell_bold), Paragraph(f"${tot_davi:,.0f}", cell_bold), Paragraph(f"${tot_caja:,.0f}", cell_bold), Paragraph(f"${liquidez_neta:,.0f}", cell_bold)]
    ]
    t_summary = Table(summary_data, colWidths=[100, 100, 110, 110, 120])
    t_summary.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#10b981')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#f0fdf4')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#bbf7d0')),
    ]))
    story.append(t_summary)
    story.append(Spacer(1, 10))

    story.append(Paragraph("Tabla Recuento Diario de Caja", h2_style))
    daily_rows = [[Paragraph("Fecha", header_cell), Paragraph("Efectivo", header_cell), Paragraph("Nequi", header_cell), Paragraph("Daviplata / Banco", header_cell), Paragraph("Total Día", header_cell)]]
    for idx, r in df_cierres_mes.iterrows():
        daily_rows.append([
            Paragraph(str(r['fecha']), cell_style), Paragraph(f"${r['efectivo']:,.0f}", cell_style),
            Paragraph(f"${r['nequi']:,.0f}", cell_style), Paragraph(f"${r['daviplata']:,.0f}", cell_style),
            Paragraph(f"${r['total']:,.0f}", cell_bold)
        ])
    daily_rows.append([
        Paragraph("TOTAL", header_cell), Paragraph(f"${tot_efectivo:,.0f}", header_cell),
        Paragraph(f"${tot_nequi:,.0f}", header_cell), Paragraph(f"${tot_davi:,.0f}", header_cell),
        Paragraph(f"${tot_caja:,.0f}", header_cell)
    ])

    t_daily = Table(daily_rows, colWidths=[110, 110, 110, 110, 100])
    t_daily.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0284c7')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#f8fafc')]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#0369a1')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0')),
    ]))
    story.append(t_daily)

    if not df_gastos_mes.empty:
        story.append(Spacer(1, 10))
        story.append(Paragraph("Gastos Mensuales y Compras", h2_style))
        gastos_rows = [[Paragraph("Fecha", header_cell), Paragraph("Concepto", header_cell), Paragraph("Monto ($)", header_cell)]]
        for idx, r in df_gastos_mes.iterrows():
            gastos_rows.append([Paragraph(str(r['fecha']), cell_style), Paragraph(str(r['concepto']), cell_style), Paragraph(f"${r['monto']:,.0f}", cell_bold)])
        gastos_rows.append([Paragraph("TOTAL GASTOS", header_cell), Paragraph("", header_cell), Paragraph(f"${tot_gastos:,.0f}", header_cell)])
        t_gastos = Table(gastos_rows, colWidths=[120, 270, 150])
        t_gastos.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e11d48')),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#fff1f2')]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#be123c')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#fecdd3')),
        ]))
        story.append(t_gastos)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

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

    story.append(Paragraph(f"Presupuesto Mensual — {mes_nombre}", title_style))
    story.append(Paragraph("Papelería — Planificación Financiera y Ejecución Real", subtitle_style))
    story.append(Spacer(1, 4))

    tot_ini = df_pres['monto_inicial'].sum() if not df_pres.empty else 0.0
    tot_fin = df_pres['monto_final'].sum() if not df_pres.empty else 0.0
    dif = tot_fin - tot_ini

    summary_data = [
        [Paragraph("Presupuestado (Valor Inicial)", header_cell), Paragraph("Ejecutado (Valor Final)", header_cell), Paragraph("Diferencia / Variación", header_cell)],
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

    story.append(Paragraph("Desglose Detallado por Conceptos", h2_style))
    
    table_rows = [[
        Paragraph("Concepto", header_cell),
        Paragraph("Categoría", header_cell),
        Paragraph("Valor Inicial ($)", header_cell),
        Paragraph("Valor Final ($)", header_cell),
        Paragraph("Diferencia ($)", header_cell)
    ]]

    for idx, r in df_pres.iterrows():
        mi = float(r['monto_inicial'])
        mf = float(r['monto_final'])
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
# PARSER EXCEL
# -----------------------------------------------------------------------------
def clean_date(val, filename_date=None):
    if isinstance(val, (datetime, date)): return val.strftime("%Y-%m-%d")
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try: return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError: pass
    match = re.search(r'(\d{1,2})[/\-](\d{1,2})\d*[/\-](\d{4})', val_str)
    if match:
        d, m, y = match.groups()
        try: return datetime(int(y), int(m), int(d)).strftime("%Y-%m-%d")
        except ValueError: pass
    return filename_date if filename_date else date.today().strftime("%Y-%m-%d")

def parse_daily_excel(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet = wb.active
    data = [[sheet.cell(row=r, column=c).value for c in range(1, sheet.max_column + 1)] for r in range(1, sheet.max_row + 1)]
    
    header_idx = next((idx for idx, row in enumerate(data) if 'FECHA' in [str(x).upper() if x else '' for x in row] and 'CONCEPTO' in [str(x).upper() if x else '' for x in row]), None)
    if header_idx is None: return None, "No se encontró el formato estándar en este Excel."
        
    header = [str(x).strip().upper() if x else '' for x in data[header_idx]]
    col_fecha = header.index('FECHA') if 'FECHA' in header else 1
    col_concepto = header.index('CONCEPTO') if 'CONCEPTO' in header else 2
    col_cantidad = header.index('CANTIDAD') if 'CANTIDAD' in header else 3
    col_codigo = next((header.index(kw) for kw in ['CODIGO', 'CÓDIGO', 'NUMERO', 'NÚMERO'] if kw in header), 4)
    col_ingreso = header.index('INGRESO') if 'INGRESO' in header else 5
    col_gastos = header.index('GASTOS') if 'GASTOS' in header else 6
    col_saldo = header.index('SALDO') if 'SALDO' in header else 7
    
    filename_date_match = re.search(r'(\d{2}-\d{2}-\d{4})', uploaded_file.name)
    filename_date = datetime.strptime(filename_date_match.group(1), "%d-%m-%Y").strftime("%Y-%m-%d") if filename_date_match else None

    extracted_date = None
    transactions = []
    
    for row in data[header_idx+1:]:
        if row[col_concepto] is None and row[col_ingreso] is None and row[col_gastos] is None: continue
        code = row[col_codigo]
        if code is None or str(code).strip() == '': continue
            
        if row[col_fecha] and not extracted_date: extracted_date = clean_date(row[col_fecha], filename_date)
        concepto = str(row[col_concepto]).strip() if row[col_concepto] else ''
        try: cant = int(row[col_cantidad]) if row[col_cantidad] else 1
        except: cant = 1
            
        ingreso = float(row[col_ingreso]) if row[col_ingreso] else 0.0
        gastos = float(row[col_gastos]) if row[col_gastos] else 0.0
        saldo = float(row[col_saldo]) if row[col_saldo] else 0.0
        
        transactions.append({'concepto': concepto, 'cantidad': cant, 'codigo': int(code) if str(code).isdigit() else 1, 'ingreso': ingreso, 'gastos': gastos, 'saldo': saldo})
        
    if not extracted_date: extracted_date = filename_date if filename_date else date.today().strftime("%Y-%m-%d")
        
    saldo_efectivo = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 1)
    saldo_daviplata = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 2)
    saldo_nequi = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 3)
    saldo_banco = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 4)
    total_caja = saldo_efectivo + saldo_daviplata + saldo_nequi + saldo_banco
    
    summary = {'fecha': extracted_date, 'efectivo': saldo_efectivo, 'daviplata': saldo_daviplata, 'nequi': saldo_nequi, 'banco': saldo_banco, 'total': total_caja}
    return summary, pd.DataFrame(transactions)

# -----------------------------------------------------------------------------
# ESTRUCTURA DE LA APLICACIÓN
# -----------------------------------------------------------------------------
st.markdown(f"""
<div class="header-banner">
    <h1>📚 Sistema de Control de Caja y Presupuesto</h1>
    <p>Gestión remota de papelería — Motor: <b>{db_type.upper()}</b></p>
</div>
""", unsafe_allow_html=True)

tabs = st.tabs(["📥 Cargar Caja Diaria", "📊 Recuento Mensual", "💸 Gastos Mensuales", "📋 Presupuestos Mensuales", "📈 Métricas", "⚙️ Edición / Histórico"])

with tabs[0]:
    st.subheader("📥 Cargar reporte diario enviado desde el local")
    uploaded_file = st.file_uploader("Adjunta el archivo Excel diario (ej. 28-07-2026.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        summary, df_trans = parse_daily_excel(uploaded_file)
        if isinstance(df_trans, str): st.error(df_trans)
        else:
            st.success("✅ Archivo procesado correctamente.")
            col_f1, _ = st.columns([1, 2])
            with col_f1: fecha_final = st.date_input("Fecha del Cierre:", value=datetime.strptime(summary['fecha'], "%Y-%m-%d").date())
            fecha_str = fecha_final.strftime("%Y-%m-%d")
            
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💵 Efectivo</div><div class="metric-card-val">${summary["efectivo"]:,.0f}</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">🟣 Nequi</div><div class="metric-card-val">${summary["nequi"]:,.0f}</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="metric-card-green"><div class="metric-card-title">🔴 Daviplata</div><div class="metric-card-val">${summary["daviplata"]:,.0f}</div></div>', unsafe_allow_html=True)
            c4.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">🏦 Banco</div><div class="metric-card-val">${summary["banco"]:,.0f}</div></div>', unsafe_allow_html=True)
            c5.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💰 Total Día</div><div class="metric-card-val">${summary["total"]:,.0f}</div></div>', unsafe_allow_html=True)
            
            st.write("")
            df_edited = st.data_editor(
                df_trans, num_rows="dynamic", use_container_width=True,
                column_config={"codigo": st.column_config.SelectboxColumn("Código Pago", options=[1, 2, 3, 4])}
            )
            
            if st.button("💾 Guardar e Integrar en la Caja Mensual", type="primary"):
                eff = sum(row['ingreso'] - row['gastos'] for _, row in df_edited.iterrows() if row['codigo'] == 1)
                dav = sum(row['ingreso'] - row['gastos'] for _, row in df_edited.iterrows() if row['codigo'] == 2)
                neq = sum(row['ingreso'] - row['gastos'] for _, row in df_edited.iterrows() if row['codigo'] == 3)
                ban = sum(row['ingreso'] - row['gastos'] for _, row in df_edited.iterrows() if row['codigo'] == 4)
                tot = eff + dav + neq + ban
                
                with engine.begin() as conn:
                    if db_type == "postgres":
                        conn.execute(text("""
                            INSERT INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                            VALUES (:fecha, :efectivo, :nequi, :daviplata, :banco, :total, :obs)
                            ON CONFLICT (fecha) DO UPDATE SET
                                efectivo = EXCLUDED.efectivo,
                                nequi = EXCLUDED.nequi,
                                daviplata = EXCLUDED.daviplata,
                                banco = EXCLUDED.banco,
                                total = EXCLUDED.total,
                                observaciones = EXCLUDED.observaciones;
                        """), {"fecha": fecha_str, "efectivo": eff, "nequi": neq, "daviplata": dav, "banco": ban, "total": tot, "obs": f"Cargado desde {uploaded_file.name}"})
                        
                        conn.execute(text("DELETE FROM transacciones_diarias WHERE fecha = :fecha"), {"fecha": fecha_str})
                        for _, row in df_edited.iterrows():
                            conn.execute(text("""
                                INSERT INTO transacciones_diarias (fecha, concepto, cantidad, codigo, ingreso, gastos, saldo)
                                VALUES (:fecha, :concepto, :cantidad, :codigo, :ingreso, :gastos, :saldo);
                            """), {"fecha": fecha_str, "concepto": row['concepto'], "cantidad": row['cantidad'], "codigo": row['codigo'], "ingreso": row['ingreso'], "gastos": row['gastos'], "saldo": row['saldo']})
                    else:
                        conn.execute(text("INSERT OR REPLACE INTO cierres_diarios VALUES (:fecha, :efectivo, :nequi, :daviplata, :banco, :total, :obs)"),
                                     {"fecha": fecha_str, "efectivo": eff, "nequi": neq, "daviplata": dav, "banco": ban, "total": tot, "obs": f"Cargado desde {uploaded_file.name}"})
                        conn.execute(text("DELETE FROM transacciones_diarias WHERE fecha = :fecha"), {"fecha": fecha_str})
                        for _, row in df_edited.iterrows():
                            conn.execute(text("INSERT INTO transacciones_diarias (fecha, concepto, cantidad, codigo, ingreso, gastos, saldo) VALUES (:fecha, :concepto, :cantidad, :codigo, :ingreso, :gastos, :saldo)"),
                                         {"fecha": fecha_str, "concepto": row['concepto'], "cantidad": row['cantidad'], "codigo": row['codigo'], "ingreso": row['ingreso'], "gastos": row['gastos'], "saldo": row['saldo']})
                
                st.cache_data.clear()
                st.success(f"🎉 ¡Cierre del día {fecha_str} guardado exitosamente!")

with tabs[1]:
    st.subheader("📊 Recuento de Caja Mensual")
    df_cierres = load_cierres_diarios()
    df_gastos = load_gastos_ligeros()
    
    if df_cierres.empty:
        st.info("Aún no hay cierres cargados en la base de datos.")
    else:
        df_cierres['mes_año'] = pd.to_datetime(df_cierres['fecha']).dt.strftime('%Y-%m')
        meses_disponibles = df_cierres['mes_año'].unique()
        
        col_m1, col_m2 = st.columns([1, 1])
        with col_m1: mes_sel = st.selectbox("Selecciona el Mes:", meses_disponibles, index=len(meses_disponibles)-1)
            
        df_c_mes = df_cierres[df_cierres['mes_año'] == mes_sel].copy()
        tot_efectivo, tot_nequi, tot_davi, tot_banco, tot_caja = df_c_mes['efectivo'].sum(), df_c_mes['nequi'].sum(), df_c_mes['daviplata'].sum(), df_c_mes['banco'].sum(), df_c_mes['total'].sum()
        
        if not df_gastos.empty:
            df_gastos['mes_año'] = pd.to_datetime(df_gastos['fecha']).dt.strftime('%Y-%m')
            df_g_mes = df_gastos[df_gastos['mes_año'] == mes_sel].copy()
            tot_gastos_grandes = df_g_mes['monto'].sum()
        else:
            df_g_mes = pd.DataFrame()
            tot_gastos_grandes = 0.0
            
        liquidez_neta = tot_caja - abs(tot_gastos_grandes)

        pdf_bytes = generate_pdf_report(mes_sel, df_c_mes, df_g_mes, tot_efectivo, tot_nequi, tot_davi, tot_caja, tot_gastos_grandes, liquidez_neta)
        with col_m2:
            st.write("")
            st.download_button(label="📄 Descargar Recuento Mensual en PDF", data=pdf_bytes, file_name=f"Recuento_Caja_{mes_sel}.pdf", mime="application/pdf", type="primary")
        
        st.markdown("##### Totales Consolidados del Mes:")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💵 Total Efectivo</div><div class="metric-card-val">${tot_efectivo:,.0f}</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">🟣 Total Nequi</div><div class="metric-card-val">${tot_nequi:,.0f}</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="metric-card-green"><div class="metric-card-title">🔴 Total Daviplata</div><div class="metric-card-val">${tot_davi:,.0f}</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">📦 Total Caja Recuento</div><div class="metric-card-val">${tot_caja:,.0f}</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="metric-card-green"><div class="metric-card-title">📊 Liquidez Neto</div><div class="metric-card-val">${liquidez_neta:,.0f}</div></div>', unsafe_allow_html=True)
        
        st.write("")
        col_t1, col_t2 = st.columns([3, 2])
        
        with col_t1:
            st.markdown("##### Tabla Recuento Diario:")
            html_daily = """<div style="overflow-x:auto;"><table style="width:100%; border-collapse:collapse; background-color:#ffffff; color:#0f172a; border:1px solid #cbd5e1; font-family:sans-serif; border-radius:8px;">
              <thead><tr style="background-color:#0284c7; color:#ffffff; font-weight:bold; text-align:center;">
                <th style="padding:10px; border:1px solid #cbd5e1;">Fecha</th>
                <th style="padding:10px; border:1px solid #cbd5e1;">Efectivo</th>
                <th style="padding:10px; border:1px solid #cbd5e1;">Nequi</th>
                <th style="padding:10px; border:1px solid #cbd5e1;">Daviplata / Banco</th>
                <th style="padding:10px; border:1px solid #cbd5e1;">Total</th>
              </tr></thead><tbody>"""
            for idx, r in df_c_mes.reset_index().iterrows():
                bg = "#ffffff" if idx % 2 == 0 else "#f8fafc"
                html_daily += f"""<tr style="background-color:{bg}; text-align:center; color:#0f172a;">
                  <td style="padding:8px; border:1px solid #e2e8f0; font-weight:500;">{r['fecha']}</td>
                  <td style="padding:8px; border:1px solid #e2e8f0;">${r['efectivo']:,.0f}</td>
                  <td style="padding:8px; border:1px solid #e2e8f0;">${r['nequi']:,.0f}</td>
                  <td style="padding:8px; border:1px solid #e2e8f0;">${r['daviplata']:,.0f}</td>
                  <td style="padding:8px; border:1px solid #e2e8f0; font-weight:bold; color:#0369a1;">${r['total']:,.0f}</td>
                </tr>"""
            html_daily += f"""<tr style="background-color:#e0f2fe; color:#0369a1; font-weight:bold; text-align:center;">
                  <td style="padding:10px; border:1px solid #bae6fd;">TOTAL</td>
                  <td style="padding:10px; border:1px solid #bae6fd;">${tot_efectivo:,.0f}</td>
                  <td style="padding:10px; border:1px solid #bae6fd;">${tot_nequi:,.0f}</td>
                  <td style="padding:10px; border:1px solid #bae6fd;">${tot_davi:,.0f}</td>
                  <td style="padding:10px; border:1px solid #bae6fd; font-size:16px;">${tot_caja:,.0f}</td>
                </tr></tbody></table></div>"""
            st.markdown(html_daily, unsafe_allow_html=True)
            
        with col_t2:
            st.markdown("##### Gastos Mensuales y Mercancía (Cronológico: Inicio a Fin de Mes):")
            if df_g_mes.empty: st.info("No hay gastos mayores ingresados en este mes.")
            else:
                html_gastos = """<div style="overflow-x:auto;"><table style="width:100%; border-collapse:collapse; background-color:#ffffff; color:#0f172a; border:1px solid #cbd5e1; font-family:sans-serif; border-radius:8px;">
                  <thead><tr style="background-color:#e11d48; color:#ffffff; font-weight:bold; text-align:left;">
                    <th style="padding:10px; border:1px solid #cbd5e1;">Fecha</th>
                    <th style="padding:10px; border:1px solid #cbd5e1;">Concepto</th>
                    <th style="padding:10px; border:1px solid #cbd5e1; text-align:right;">Monto ($)</th>
                  </tr></thead><tbody>"""
                for idx, r in df_g_mes.reset_index().iterrows():
                    bg = "#ffffff" if idx % 2 == 0 else "#fff1f2"
                    has_comp = " 📎" if r['tiene_comprobante'] == 1 else ""
                    html_gastos += f"""<tr style="background-color:{bg}; color:#0f172a;">
                      <td style="padding:8px; border:1px solid #fecdd3; font-weight:500;">{r['fecha']}</td>
                      <td style="padding:8px; border:1px solid #fecdd3;">{r['concepto']}{has_comp}</td>
                      <td style="padding:8px; border:1px solid #fecdd3; text-align:right; font-weight:bold; color:#be123c;">${r['monto']:,.0f}</td>
                    </tr>"""
                html_gastos += f"""<tr style="background-color:#ffe4e6; color:#be123c; font-weight:bold;">
                      <td style="padding:10px; border:1px solid #fecdd3;">TOTAL GASTOS</td>
                      <td style="padding:10px; border:1px solid #fecdd3;"></td>
                      <td style="padding:10px; border:1px solid #fecdd3; text-align:right; font-size:15px;">${tot_gastos_grandes:,.0f}</td>
                    </tr></tbody></table></div>"""
                st.markdown(html_gastos, unsafe_allow_html=True)

with tabs[2]:
    st.subheader("💸 Registrar Gasto Mensual o Compra Grande")
    with st.form("form_gastos_clean", clear_on_submit=True):
        col_g1, col_g2, col_g3 = st.columns(3)
        with col_g1: fecha_gasto = st.date_input("Fecha:", value=date.today())
        with col_g2: categoria_gasto = st.selectbox("Categoría:", ["Arriendo", "Servicios (Agua/Luz/Gas)", "Internet/Teléfono", "Mercancía/Proveedores", "Pago Deuda/Terceros", "Otro"])
        with col_g3: monto_gasto = st.number_input("Monto en Pesos ($):", min_value=0.0, step=1000.0)
        
        col_g4, col_g5 = st.columns([2, 2])
        with col_g4: concepto_gasto = st.text_input("Concepto / Proveedor:")
        with col_g5: uploaded_comprobante = st.file_uploader("📎 Adjuntar Comprobante (Opcional - Imagen o PDF):", type=["pdf", "png", "jpg", "jpeg", "webp"])
        
        submitted = st.form_submit_button("➕ Registrar Gasto Mensual", type="primary")
        if submitted and monto_gasto > 0 and concepto_gasto.strip():
            comprobante_json = encode_file_to_json(uploaded_comprobante)
            with engine.begin() as conn:
                conn.execute(text("""
                    INSERT INTO gastos_mensuales (fecha, concepto, monto, categoria, comprobante)
                    VALUES (:fecha, :concepto, :monto, :categoria, :comprobante);
                """), {
                    "fecha": fecha_gasto.strftime("%Y-%m-%d"),
                    "concepto": concepto_gasto.upper().strip(),
                    "monto": -abs(monto_gasto),
                    "categoria": categoria_gasto,
                    "comprobante": comprobante_json
                })
            st.cache_data.clear()
            st.success(f"✅ Gasto '{concepto_gasto}' registrado exitosamente.")
            st.rerun()

    st.divider()
    st.subheader("🖼️ Visor y Descarga de Comprobantes Adjuntos")
    df_gastos_all = load_gastos_ligeros()
    
    if df_gastos_all.empty:
        st.info("Aún no hay gastos registrados.")
    else:
        df_con_comprobante = df_gastos_all[df_gastos_all['tiene_comprobante'] == 1].copy()
        
        if df_con_comprobante.empty:
            st.info("Aún no hay comprobantes o facturas adjuntadas a los gastos.")
        else:
            options_dict = {f"[{r['fecha']}] {r['concepto']} (${abs(r['monto']):,.0f})": r['id'] for _, r in df_con_comprobante.iterrows()}
            selected_label = st.selectbox("Selecciona un gasto para ver o descargar su comprobante:", list(options_dict.keys()))
            
            selected_id = options_dict[selected_label]
            selected_row = df_con_comprobante[df_con_comprobante['id'] == selected_id].iloc[0]
            
            raw_comp_str = load_single_comprobante(selected_id)
            file_info = decode_comprobante_json(raw_comp_str)
            
            if file_info:
                col_v1, col_v2 = st.columns([1, 2])
                with col_v1:
                    st.write(f"**Archivo:** `{file_info['name']}`")
                    st.write(f"**Categoría:** {selected_row['categoria']}")
                    st.download_button(
                        label=f"⬇️ Descargar {file_info['name']}",
                        data=file_info['bytes'],
                        file_name=file_info['name'],
                        mime=file_info['type'],
                        type="primary"
                    )
                with col_v2:
                    if file_info['type'].startswith("image/"):
                        st.image(file_info['bytes'], caption=f"Vista previa: {file_info['name']}", use_column_width=True)
                    elif file_info['type'] == "application/pdf":
                        st.info("📄 Archivo PDF adjunto. Haz clic a la izquierda para descargarlo y abrirlo.")

    st.divider()
    st.subheader("✏️ Edición Rápida y Actualización de Gastos")
    st.write("Modifica datos de los gastos (ordenados de inicio a fin de mes) o adjunta/reemplaza comprobantes.")
    
    df_gastos_display = df_gastos_all.copy()
    df_gastos_display['tiene_comprobante'] = df_gastos_display['tiene_comprobante'].map({1: "📎 Sí", 0: "❌ No"})
    
    df_gastos_edited = st.data_editor(
        df_gastos_display[['id', 'fecha', 'concepto', 'monto', 'categoria', 'tiene_comprobante']],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "id": st.column_config.NumberColumn("ID", disabled=True),
            "fecha": st.column_config.TextColumn("Fecha (AAAA-MM-DD)", required=True),
            "concepto": st.column_config.TextColumn("Concepto / Proveedor", required=True),
            "monto": st.column_config.NumberColumn("Monto ($)", format="$%d", required=True),
            "categoria": st.column_config.SelectboxColumn("Categoría", options=["Arriendo", "Servicios (Agua/Luz/Gas)", "Internet/Teléfono", "Mercancía/Proveedores", "Pago Deuda/Terceros", "Otro"]),
            "tiene_comprobante": st.column_config.TextColumn("Comprobante", disabled=True)
        },
        key="gastos_editor_table"
    )
    
    col_e1, col_e2 = st.columns([2, 2])
    with col_e1:
        if st.button("💾 Guardar Cambios de Texto / Montos", type="primary"):
            with engine.begin() as conn:
                for _, r in df_gastos_edited.iterrows():
                    g_id = r['id']
                    if pd.notna(g_id):
                        m_val = float(r['monto'])
                        if m_val > 0: m_val = -m_val
                        conn.execute(text("""
                            UPDATE gastos_mensuales 
                            SET fecha = :fecha, concepto = :concepto, monto = :monto, categoria = :categoria
                            WHERE id = :id;
                        """), {
                            "fecha": str(r['fecha']),
                            "concepto": str(r['concepto']).upper().strip(),
                            "monto": m_val,
                            "categoria": str(r['categoria']),
                            "id": int(g_id)
                        })
            st.cache_data.clear()
            st.success("✅ ¡Gastos modificados correctamente!")
            st.rerun()

    with col_e2:
        with st.expander("📎 Adjuntar o Reemplazar Comprobante en un Gasto Existente"):
            options_edit = {f"[{r['fecha']}] {r['concepto']} (${abs(r['monto']):,.0f})": r['id'] for _, r in df_gastos_all.iterrows()}
            if options_edit:
                selected_edit_label = st.selectbox("Selecciona el gasto al cual agregarle el comprobante:", list(options_edit.keys()))
                edit_id = options_edit[selected_edit_label]
                new_file = st.file_uploader("Selecciona el archivo comprobante:", type=["pdf", "png", "jpg", "jpeg", "webp"], key="edit_uploader")
                if st.button("💾 Adjuntar Comprobante al Gasto"):
                    if new_file:
                        new_json = encode_file_to_json(new_file)
                        with engine.begin() as conn:
                            conn.execute(text("UPDATE gastos_mensuales SET comprobante = :comp WHERE id = :id;"), {"comp": new_json, "id": edit_id})
                        st.cache_data.clear()
                        st.success("✅ Comprobante adjuntado con éxito.")
                        st.rerun()

with tabs[3]:
    st.subheader("📋 Presupuestos Mensuales (Proyectado vs Real)")
    st.write("Planifica y compara el presupuesto proyectado (Valor Inicial) contra lo ejecutado (Valor Final) de cada mes.")
    
    df_pres_all = load_presupuestos_mensuales()
    
    df_cierres_aux = load_cierres_diarios()
    meses_cierres = pd.to_datetime(df_cierres_aux['fecha']).dt.strftime('%Y-%m').unique() if not df_cierres_aux.empty else []
    meses_pres = df_pres_all['mes_año'].unique() if not df_pres_all.empty else []
    
    all_months = sorted(list(set(list(meses_pres) + list(meses_cierres) + [date.today().strftime('%Y-%m')])), reverse=True)
    
    col_p1, col_p2 = st.columns([1, 1])
    with col_p1:
        mes_sel_pres = st.selectbox("Selecciona el Mes para Planificar Presupuesto:", all_months, key="select_mes_presupuesto")
        
    df_p_mes = df_pres_all[df_pres_all['mes_año'] == mes_sel_pres].copy()
    
    if df_p_mes.empty:
        st.info(f"Aún no hay conceptos de presupuesto registrados para el mes **{mes_sel_pres}**.")
        col_seed1, _ = st.columns([1, 1])
        with col_seed1:
            if st.button("✨ Cargar Plantilla Estándar de Presupuesto", type="primary"):
                template_items = [
                    ("Renta", 1100000.0, 1100000.0, "Gastos Fijos"),
                    ("Servicios (Agua/Luz/Gas)", 280000.0, 280000.0, "Servicios"),
                    ("Internet / Teléfono", 75000.0, 75000.0, "Servicios"),
                    ("Pago de impuestos", 250000.0, 250000.0, "Impuestos"),
                    ("Abono deuda 5,5M Ana", 160000.0, 160000.0, "Deudas"),
                    ("Pago Juancho", 120000.0, 120000.0, "Personal"),
                    ("Pago Ana", 80000.0, 80000.0, "Personal"),
                    ("Mercancía / Proveedores", 800000.0, 800000.0, "Mercancía"),
                    ("Abono Fotocopiadora / Inversión", 150000.0, 150000.0, "Inversión/Equipo")
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
            st.download_button(label="📄 Descargar Presupuesto Mensual en PDF", data=pdf_pres_bytes, file_name=f"Presupuesto_{mes_sel_pres}.pdf", mime="application/pdf", type="primary")

        st.markdown("##### Resumen del Presupuesto:")
        kp1, kp2, kp3 = st.columns(3)
        kp1.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">📌 Total Presupuestado (Valor Inicial)</div><div class="metric-card-val">${tot_ini:,.0f}</div></div>', unsafe_allow_html=True)
        kp2.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💸 Total Ejecutado (Valor Final)</div><div class="metric-card-val">${tot_fin:,.0f}</div></div>', unsafe_allow_html=True)
        
        diff_card_class = "metric-card-green" if dif_tot <= 0 else "metric-card-amber"
        diff_prefix = f"-${abs(dif_tot):,.0f} (Ahorro)" if dif_tot < 0 else (f"+${dif_tot:,.0f} (Exceso)" if dif_tot > 0 else "$0 (En meta)")
        kp3.markdown(f'<div class="{diff_card_class}"><div class="metric-card-title">⚖️ Diferencia / Variación</div><div class="metric-card-val">{diff_prefix}</div></div>', unsafe_allow_html=True)

        st.write("")
        st.markdown("##### Tabla Interactiva de Presupuesto (Modificable en vivo):")
        
        df_p_mes['diferencia'] = df_p_mes['monto_final'] - df_p_mes['monto_inicial']
        
        df_pres_edited = st.data_editor(
            df_p_mes[['id', 'concepto', 'categoria', 'monto_inicial', 'monto_final', 'diferencia']],
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "concepto": st.column_config.TextColumn("Concepto / Rubro", required=True),
                "categoria": st.column_config.SelectboxColumn("Categoría", options=["Gastos Fijos", "Servicios", "Impuestos", "Deudas", "Personal", "Mercancía", "Inversión/Equipo", "Otro"]),
                "monto_inicial": st.column_config.NumberColumn("Valor Inicial (Proyectado $)", format="$%d", required=True),
                "monto_final": st.column_config.NumberColumn("Valor Final (Ejecutado $)", format="$%d", required=True),
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

with tabs[4]:
    st.subheader("📈 Análisis de Facturación por Días y Curva Mensual")
    df_cierres = load_cierres_diarios()
    df_trans = load_transacciones_diarias()
    
    if not df_cierres.empty:
        df_cierres['mes_año'] = pd.to_datetime(df_cierres['fecha']).dt.strftime('%Y-%m')
        meses_disponibles = df_cierres['mes_año'].unique()
        col_m1, _ = st.columns([1, 1])
        with col_m1: mes_sel_graf = st.selectbox("Selecciona Mes para ver Facturación Diaria:", meses_disponibles, index=len(meses_disponibles)-1)
        df_c_graf = df_cierres[df_cierres['mes_año'] == mes_sel_graf].copy()
        
        st.markdown("#### 1. Facturación Día a Día")
        fig_bar = px.bar(df_c_graf, x="fecha", y="total", text_auto='.2s', labels={"fecha": "Día", "total": "Facturación ($)"}, color_discrete_sequence=['#10b981'])
        fig_bar.update_layout(plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#0f172a'), yaxis=dict(showgrid=True, gridcolor='#f1f5f9'))
        st.plotly_chart(fig_bar, use_container_width=True)
        
        st.divider()
        st.markdown("#### 2. Curva Evolutiva por Meses")
        df_mensual_sum = df_cierres.groupby('mes_año')['total'].sum().reset_index()
        fig_curve = px.line(df_mensual_sum, x="mes_año", y="total", markers=True, labels={"mes_año": "Mes", "total": "Facturación Acumulada ($)"}, color_discrete_sequence=['#0284c7'])
        fig_curve.update_layout(plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#0f172a'), yaxis=dict(showgrid=True, gridcolor='#f1f5f9'))
        st.plotly_chart(fig_curve, use_container_width=True)

with tabs[5]:
    st.subheader("⚙️ Edición Rápida de Cierres Diarios e Ingresos")
    st.write("Modifica directamente cualquier cierre diario o saldo si hubo un error en fecha, efectivo, Nequi, Daviplata o Banco.")
    
    df_all = load_cierres_diarios()
    
    if not df_all.empty:
        df_edit_hist = st.data_editor(
            df_all[['fecha', 'efectivo', 'nequi', 'daviplata', 'banco', 'total', 'observaciones']],
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "fecha": st.column_config.TextColumn("Fecha (AAAA-MM-DD)", required=True),
                "efectivo": st.column_config.NumberColumn("Efectivo ($)", format="$%d"),
                "nequi": st.column_config.NumberColumn("Nequi ($)", format="$%d"),
                "daviplata": st.column_config.NumberColumn("Daviplata ($)", format="$%d"),
                "banco": st.column_config.NumberColumn("Banco ($)", format="$%d"),
                "total": st.column_config.NumberColumn("Total ($)", format="$%d"),
                "observaciones": st.column_config.TextColumn("Observaciones")
            },
            key="historico_editor"
        )
        
        if st.button("💾 Guardar Cambios en Cierres Diarios / Ingresos", type="primary"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM cierres_diarios;"))
                for _, r in df_edit_hist.iterrows():
                    eff = float(r['efectivo']) if r['efectivo'] is not None else 0.0
                    neq = float(r['nequi']) if r['nequi'] is not None else 0.0
                    dav = float(r['daviplata']) if r['daviplata'] is not None else 0.0
                    ban = float(r['banco']) if r['banco'] is not None else 0.0
                    tot = eff + neq + dav + ban
                    
                    if db_type == "postgres":
                        conn.execute(text("""
                            INSERT INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                            VALUES (:fecha, :efectivo, :nequi, :daviplata, :banco, :total, :obs);
                        """), {"fecha": str(r['fecha']), "efectivo": eff, "nequi": neq, "daviplata": dav, "banco": ban, "total": tot, "obs": str(r['observaciones']) if r['observaciones'] else ''})
                    else:
                        conn.execute(text("INSERT INTO cierres_diarios VALUES (:fecha, :efectivo, :nequi, :daviplata, :banco, :total, :obs);"),
                                     {"fecha": str(r['fecha']), "efectivo": float(r['efectivo']), "nequi": float(r['nequi']), "daviplata": float(r['daviplata']), "banco": float(r['banco']), "total": float(r['total']), "obs": str(r['observaciones']) if r['observaciones'] else ''})
            st.cache_data.clear()
            st.success("✅ ¡Cierres de caja e ingresos actualizados correctamente!")
            st.rerun()
