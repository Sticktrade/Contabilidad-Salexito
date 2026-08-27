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

# Generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y FORZADO GLOBAL DE TEMA CLARO
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control de Caja - Papelería",
    page_icon="📚",
    layout="wide"
)

st.markdown("""
<style>
    /* 1. Forzar esquema claro global en navegador y portales de Streamlit */
    :root, body, html {
        color-scheme: light !important;
    }
    
    html, body, [data-testid="stAppViewContainer"], .main, .stApp {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }

    /* 2. Reparar Calendario Emergente y Menús Desplegables (Portales de BaseWeb) */
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

    /* 3. Campo de Entrada de Fecha (Date Input) */
    div[data-testid="stDateInput"] input,
    div[data-testid="stDateInput"] > div {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }

    /* 4. Pestañas (Tabs) con Títulos Visibles */
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

    /* 5. Labels y Cajas de Texto / Selección */
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

    /* 6. Uploader y Botones */
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

    /* Banners y Métricas */
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
    .metric-card-amber { background-color: #fffbeb; border: 1px solid #fde68a; border-radius: 10px; padding: 12px; text-align: center; }
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
                        mes VARCHAR(20), concepto TEXT, valor_inicial FLOAT, valor_final FLOAT
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
                        mes TEXT, concepto TEXT, valor_inicial REAL, valor_final REAL
                    );
                """))
    except Exception as ex:
        st.warning(f"Error inicializando tablas: {ex}")

init_db()

# -----------------------------------------------------------------------------
# CONSULTAS OPTIMIZADAS CON CACHÉ (Evita descargas pesadas innecesarias)
# -----------------------------------------------------------------------------
@st.cache_data(ttl=60)
def load_cierres_diarios():
    try:
        with engine.connect() as conn:
            return pd.read_sql_query(text("SELECT * FROM cierres_diarios ORDER BY fecha ASC"), conn)
    except Exception:
        return pd.DataFrame(columns=['fecha', 'efectivo', 'nequi', 'daviplata', 'banco', 'total', 'observaciones'])

@st.cache_data(ttl=60)
def load_transacciones_diarias():
    try:
        with engine.connect() as conn:
            return pd.read_sql_query(text("SELECT * FROM transacciones_diarias"), conn)
    except Exception:
        return pd.DataFrame(columns=['id', 'fecha', 'concepto', 'cantidad', 'codigo', 'ingreso', 'gastos', 'saldo'])

@st.cache_data(ttl=60)
def load_gastos_ligeros():
    """Carga los gastos SIN traer el texto pesadísimo del archivo comprobante en base64 (Orden cronológico por defecto)"""
    try:
        with engine.connect() as conn:
            query = """
                SELECT id, fecha, concepto, monto, categoria, 
                       CASE WHEN comprobante IS NOT NULL AND length(comprobante) > 5 THEN 1 ELSE 0 END as tiene_comprobante
                FROM gastos_mensuales ORDER BY fecha ASC
            """
            return pd.read_sql_query(text(query), conn)
    except Exception:
        return pd.DataFrame(columns=['id', 'fecha', 'concepto', 'monto', 'categoria', 'tiene_comprobante'])

@st.cache_data(ttl=60)
def load_presupuestos_all():
    try:
        with engine.connect() as conn:
            return pd.read_sql_query(text("SELECT * FROM presupuestos_mensuales ORDER BY id ASC"), conn)
    except Exception:
        return pd.DataFrame(columns=['id', 'mes', 'concepto', 'valor_inicial', 'valor_final'])

def load_single_comprobante(gasto_id):
    """Obtiene el comprobante solo cuando el usuario selecciona ver o descargar ese gasto específico"""
    try:
        with engine.connect() as conn:
            res = conn.execute(text("SELECT comprobante FROM gastos_mensuales WHERE id = :id"), {"id": gasto_id}).fetchone()
            return res[0] if res else None
    except Exception:
        return None

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
def generate_pdf_presupuesto(mes_nombre, df_presupuesto, total_inicial, total_final, diferencia):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#ea580c'), spaceAfter=4)
    subtitle_style = ParagraphStyle('DocSubtitle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#c2410c'), spaceAfter=14)
    h2_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor('#0f172a'), spaceBefore=10, spaceAfter=6)
    cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#1e293b'))
    cell_bold = ParagraphStyle('CellB', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#0f172a'))
    header_cell = ParagraphStyle('HCell', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.white)

    story.append(Paragraph(f"Presupuesto Mensual de Gastos — {mes_nombre}", title_style))
    story.append(Paragraph("Papelería — Planeación y Ejecución Presupuestal Remota", subtitle_style))
    story.append(Spacer(1, 4))

    summary_data = [
        [Paragraph("Valor Inicial (Presupuestado)", header_cell), Paragraph("Valor Final (Ejecutado)", header_cell), Paragraph("Diferencia / Variación", header_cell)],
        [Paragraph(f"${total_inicial:,.0f}", cell_bold), Paragraph(f"${total_final:,.0f}", cell_bold), Paragraph(f"${diferencia:,.0f}", cell_bold)]
    ]
    t_summary = Table(summary_data, colWidths=[180, 180, 180])
    t_summary.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ea580c')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#fff7ed')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#fed7aa')),
    ]))
    story.append(t_summary)
    story.append(Spacer(1, 12))

    story.append(Paragraph("Detalle de Conceptos Presupuestados", h2_style))
    budget_rows = [[
        Paragraph("Concepto / Rubro", header_cell), 
        Paragraph("Valor Inicial ($)", header_cell), 
        Paragraph("Valor Final ($)", header_cell), 
        Paragraph("Diferencia ($)", header_cell)
    ]]
    
    for idx, r in df_presupuesto.iterrows():
        v_ini = float(r['valor_inicial']) if pd.notna(r['valor_inicial']) else 0.0
        v_fin = float(r['valor_final']) if pd.notna(r['valor_final']) else 0.0
        diff = v_fin - v_ini
        budget_rows.append([
            Paragraph(str(r['concepto']), cell_style), 
            Paragraph(f"${v_ini:,.0f}", cell_style),
            Paragraph(f"${v_fin:,.0f}", cell_style), 
            Paragraph(f"${diff:,.0f}", cell_bold)
        ])
        
    budget_rows.append([
        Paragraph("TOTAL PRESUPUESTO", header_cell), 
        Paragraph(f"${total_inicial:,.0f}", header_cell),
        Paragraph(f"${total_final:,.0f}", header_cell), 
        Paragraph(f"${diferencia:,.0f}", header_cell)
    ])

    t_budget = Table(budget_rows, colWidths=[210, 110, 110, 110])
    t_budget.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ea580c')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#fff7ed')]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#c2410c')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#fed7aa')),
    ]))
    story.append(t_budget)

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
    <h1>📚 Sistema de Control de Caja y Contabilidad</h1>
    <p>Gestión remota de papelería — Motor: <b>{db_type.upper()}</b></p>
</div>
""", unsafe_allow_html=True)

tabs = st.tabs(["📥 Cargar Caja Diaria", "📊 Recuento Mensual", "💸 Gastos Mensuales", "📋 Presupuestos", "📈 Métricas", "⚙️ Edición / Histórico"])

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
            df_g_mes = df_gastos[df_gastos['mes_año'] == mes_sel].sort_values('fecha', ascending=True).copy()
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
            st.markdown("##### Gastos Mensuales y Mercancía:")
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
            
            # Carga del archivo solo bajo demanda
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
    st.write("Modifica datos de los gastos o adjunta/reemplaza un comprobante si olvidaste subirlo.")
    
    df_gastos_display = df_gastos_all.copy()
    if 'tiene_comprobante' in df_gastos_display.columns:
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
    st.subheader("📋 Presupuesto de Gastos Mensuales")
    st.write("Planifica y realiza el seguimiento de tus gastos proyectados (Valor Inicial) vs. los gastos reales (Valor Final).")
    
    df_pres_all = load_presupuestos_all()
    
    # Selección de Mes
    df_cierres_aux = load_cierres_diarios()
    meses_disponibles = []
    if not df_cierres_aux.empty and 'fecha' in df_cierres_aux.columns:
        df_cierres_aux['mes_año'] = pd.to_datetime(df_cierres_aux['fecha']).dt.strftime('%Y-%m')
        meses_disponibles = list(df_cierres_aux['mes_año'].unique())
    
    mes_actual_str = date.today().strftime('%Y-%m')
    if mes_actual_str not in meses_disponibles:
        meses_disponibles.append(mes_actual_str)
        
    if not df_pres_all.empty and 'mes' in df_pres_all.columns:
        for m in df_pres_all['mes'].unique():
            if m and str(m) not in meses_disponibles:
                meses_disponibles.append(str(m))
                
    col_p1, col_p2 = st.columns([1, 1])
    with col_p1:
        mes_pres_sel = st.selectbox("Selecciona el Mes del Presupuesto:", sorted(meses_disponibles, reverse=True), key="sb_mes_presupuesto")
    
    df_pres_mes = df_pres_all[df_pres_all['mes'] == mes_pres_sel].copy() if (not df_pres_all.empty and 'mes' in df_pres_all.columns) else pd.DataFrame()
    
    if df_pres_mes.empty or 'concepto' not in df_pres_mes.columns:
        default_concepts = ["Renta", "Servicios", "Internet", "Abono deuda 5,5M Ana", "Pago Juancho", "Pago Ana", "Mercancia", "Fotocopiadora abono"]
        df_pres_input = pd.DataFrame({
            "concepto": default_concepts,
            "valor_inicial": [0.0] * len(default_concepts),
            "valor_final": [0.0] * len(default_concepts)
        })
    else:
        df_pres_input = df_pres_mes[['concepto', 'valor_inicial', 'valor_final']].copy()
        
    st.markdown("##### ✏️ Formulario Interactivo de Presupuesto (Conceptos y Valores):")
    df_pres_edited = st.data_editor(
        df_pres_input,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "concepto": st.column_config.TextColumn("Concepto / Rubro", required=True),
            "valor_inicial": st.column_config.NumberColumn("Valor Inicial ($)", format="$%d", min_value=0.0, default=0.0),
            "valor_final": st.column_config.NumberColumn("Valor Final ($)", format="$%d", min_value=0.0, default=0.0)
        },
        key="presupuesto_data_editor"
    )
    
    tot_pres_inicial = df_pres_edited['valor_inicial'].sum() if not df_pres_edited.empty else 0.0
    tot_pres_final = df_pres_edited['valor_final'].sum() if not df_pres_edited.empty else 0.0
    tot_pres_diff = tot_pres_final - tot_pres_inicial
    
    col_btn1, _ = st.columns([1, 1])
    with col_btn1:
        if st.button("💾 Guardar Presupuesto del Mes", type="primary"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM presupuestos_mensuales WHERE mes = :mes"), {"mes": mes_pres_sel})
                for _, r in df_pres_edited.iterrows():
                    c_nom = str(r['concepto']).strip()
                    if c_nom:
                        v_i = float(r['valor_inicial']) if pd.notna(r['valor_inicial']) else 0.0
                        v_f = float(r['valor_final']) if pd.notna(r['valor_final']) else 0.0
                        conn.execute(text("""
                            INSERT INTO presupuestos_mensuales (mes, concepto, valor_inicial, valor_final)
                            VALUES (:mes, :concepto, :valor_inicial, :valor_final);
                        """), {"mes": mes_pres_sel, "concepto": c_nom, "valor_inicial": v_i, "valor_final": v_f})
            st.cache_data.clear()
            st.success(f"✅ ¡Presupuesto para {mes_pres_sel} guardado exitosamente!")
            st.rerun()
            
    pdf_pres_bytes = generate_pdf_presupuesto(mes_pres_sel, df_pres_edited, tot_pres_inicial, tot_pres_final, tot_pres_diff)
    with col_p2:
        st.write("")
        st.download_button(
            label="📄 Descargar Presupuesto Mensual en PDF",
            data=pdf_pres_bytes,
            file_name=f"Presupuesto_Mensual_{mes_pres_sel}.pdf",
            mime="application/pdf",
            type="primary"
        )
        
    st.divider()
    
    p1, p2, p3 = st.columns(3)
    p1.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">📌 Total Valor Inicial</div><div class="metric-card-val">${tot_pres_inicial:,.0f}</div></div>', unsafe_allow_html=True)
    p2.markdown(f'<div class="metric-card-green"><div class="metric-card-title">🎯 Total Valor Final</div><div class="metric-card-val">${tot_pres_final:,.0f}</div></div>', unsafe_allow_html=True)
    p3.markdown(f'<div class="metric-card-amber"><div class="metric-card-title">📊 Diferencia / Variación</div><div class="metric-card-val">${tot_pres_diff:,.0f}</div></div>', unsafe_allow_html=True)
    
    st.write("")
    st.markdown("##### 📊 Vista Consolidada del Presupuesto (Estilo Recuento Mensual):")
    
    if df_pres_edited.empty:
        st.info("No hay datos cargados en el presupuesto.")
    else:
        html_presupuesto = """<div style="overflow-x:auto;"><table style="width:100%; border-collapse:collapse; background-color:#ffffff; color:#0f172a; border:1px solid #fed7aa; font-family:sans-serif; border-radius:8px;">
        <thead><tr style="background-color:#ea580c; color:#ffffff; font-weight:bold; text-align:left;">
          <th style="padding:10px; border:1px solid #fed7aa;">Concepto / Rubro</th>
          <th style="padding:10px; border:1px solid #fed7aa; text-align:right;">Valor Inicial ($)</th>
          <th style="padding:10px; border:1px solid #fed7aa; text-align:right;">Valor Final ($)</th>
          <th style="padding:10px; border:1px solid #fed7aa; text-align:right;">Diferencia ($)</th>
        </tr></thead><tbody>"""
        
        for idx, r in df_pres_edited.reset_index().iterrows():
            bg = "#ffffff" if idx % 2 == 0 else "#fff7ed"
            v_ini = float(r['valor_inicial']) if pd.notna(r['valor_inicial']) else 0.0
            v_fin = float(r['valor_final']) if pd.notna(r['valor_final']) else 0.0
            diff = v_fin - v_ini
            
            highlight_style = "background-color:#fef9c3; font-weight:bold; color:#854d0e;" if (v_fin > 0 and v_fin != v_ini) else ""
            
            html_presupuesto += f"""<tr style="background-color:{bg}; color:#0f172a;">
              <td style="padding:8px; border:1px solid #fed7aa; font-weight:500;">{r['concepto']}</td>
              <td style="padding:8px; border:1px solid #fed7aa; text-align:right;">${v_ini:,.0f}</td>
              <td style="padding:8px; border:1px solid #fed7aa; text-align:right; {highlight_style}">${v_fin:,.0f}</td>
              <td style="padding:8px; border:1px solid #fed7aa; text-align:right; font-weight:bold; color:#c2410c;">${diff:,.0f}</td>
            </tr>"""
            
        html_presupuesto += f"""<tr style="background-color:#ffedd5; color:#c2410c; font-weight:bold;">
          <td style="padding:10px; border:1px solid #fed7aa;">TOTAL PRESUPUESTO</td>
          <td style="padding:10px; border:1px solid #fed7aa; text-align:right; font-size:15px;">${tot_pres_inicial:,.0f}</td>
          <td style="padding:10px; border:1px solid #fed7aa; text-align:right; font-size:15px;">${tot_pres_final:,.0f}</td>
          <td style="padding:10px; border:1px solid #fed7aa; text-align:right; font-size:15px;">${tot_pres_diff:,.0f}</td>
        </tr></tbody></table></div>"""
        
        st.markdown(html_presupuesto, unsafe_allow_html=True)

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
