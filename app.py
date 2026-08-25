import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import re
from datetime import datetime, date
import plotly.express as px
import io
from sqlalchemy import create_engine, text

# PDF Generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y TEMA CLARO
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control de Caja - Papelería",
    page_icon="📚",
    layout="wide"
)

st.markdown("""
<style>
    /* Fondo claro absoluto */
    html, body, [data-testid="stAppViewContainer"], .main, .stApp {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }

    /* Uploader */
    [data-testid="stFileUploader"], [data-testid="stFileUploaderDropzone"] {
        background-color: #f8fafc !important;
        border: 2px dashed #cbd5e1 !important;
        border-radius: 12px !important;
    }
    [data-testid="stFileUploaderDropzone"] * {
        color: #334155 !important;
    }

    /* Pestañas (Tabs) */
    button[data-baseweb="tab"] {
        background-color: #f1f5f9 !important;
        border: 1px solid #cbd5e1 !important;
        color: #334155 !important;
        border-radius: 8px !important;
        margin-right: 6px !important;
        padding: 8px 16px !important;
        font-weight: 600 !important;
    }
    button[data-baseweb="tab"]:hover {
        background-color: #e0f2fe !important;
        color: #0284c7 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #0284c7 !important;
        color: #ffffff !important;
        border-color: #0284c7 !important;
    }

    /* Banner */
    .header-banner {
        background: linear-gradient(135deg, #e0f2fe 0%, #dcfce7 100%);
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #bae6fd;
        margin-bottom: 20px;
    }
    .header-banner h1 { color: #0369a1 !important; margin: 0; font-size: 24px; font-weight: 700; }
    .header-banner p { color: #047857 !important; margin: 4px 0 0 0; font-size: 14px; }

    /* Tarjetas de Métricas */
    .metric-card-green { background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px; padding: 12px; text-align: center; }
    .metric-card-blue { background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 10px; padding: 12px; text-align: center; }
    .metric-card-title { font-size: 13px; color: #475569 !important; font-weight: 600; }
    .metric-card-val { font-size: 20px; font-weight: 700; color: #0f172a !important; }

    /* Botones */
    div.stButton > button {
        background-color: #10b981 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# CONEXIÓN SEGURA A BASE DE DATOS
# -----------------------------------------------------------------------------
@st.cache_resource
def get_db_engine():
    if "SUPABASE_URL" in st.secrets and st.secrets["SUPABASE_URL"].strip():
        db_url = st.secrets["SUPABASE_URL"].strip()
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)
        try:
            eng = create_engine(db_url, pool_pre_ping=True, pool_size=5, max_overflow=10)
            # Probar conexión
            with eng.connect() as conn:
                conn.execute(text("SELECT 1;"))
            return eng, "postgres", None
        except Exception as e:
            return create_engine("sqlite:///papeleria.db"), "sqlite", str(e)
    else:
        return create_engine("sqlite:///papeleria.db"), "sqlite", "Sin configurar SUPABASE_URL en Secrets."

engine, db_type, conn_error = get_db_engine()

if conn_error and "SUPABASE_URL" in st.secrets:
    st.error(f"⚠️ **Error al conectar a Supabase:** No se pudo establecer conexión con la URL ingresada.\n\n"
             f"**Detalle técnico:** `{conn_error}`\n\n"
             f"**Sugerencias para resolverlo:**\n"
             f"1. Verifica que reemplazaste `[YOUR-PASSWORD]` con la clave del proyecto.\n"
             f"2. Si la clave tiene `@`, cámbialo por `%40`.\n"
             f"3. Usa la URL del Pooler (`.pooler.supabase.com:6543`) y añade `?sslmode=require` al final.")

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
    except Exception as ex:
        st.warning(f"Error inicializando tablas de base de datos: {ex}")

init_db()

# -----------------------------------------------------------------------------
# GENERACIÓN DE REPORTE PDF
# -----------------------------------------------------------------------------
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

tabs = st.tabs(["📥 Cargar Caja Diaria", "📊 Recuento Mensual", "💸 Gastos Mensuales", "📈 Métricas", "⚙️ Histórico"])

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
                
                st.success(f"🎉 ¡Cierre del día {fecha_str} guardado exitosamente!")

with tabs[1]:
    st.subheader("📊 Recuento de Caja Mensual")
    with engine.connect() as conn:
        df_cierres = pd.read_sql_query(text("SELECT * FROM cierres_diarios ORDER BY fecha ASC"), conn)
        df_gastos = pd.read_sql_query(text("SELECT * FROM gastos_mensuales ORDER BY fecha ASC"), conn)
    
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
                    html_gastos += f"""<tr style="background-color:{bg}; color:#0f172a;">
                      <td style="padding:8px; border:1px solid #fecdd3; font-weight:500;">{r['fecha']}</td>
                      <td style="padding:8px; border:1px solid #fecdd3;">{r['concepto']}</td>
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
        concepto_gasto = st.text_input("Concepto:")
        submitted = st.form_submit_button("➕ Registrar Gasto Mensual", type="primary")
        if submitted and monto_gasto > 0 and concepto_gasto.strip():
            with engine.begin() as conn:
                conn.execute(text("""
                    INSERT INTO gastos_mensuales (fecha, concepto, monto, categoria, comprobante)
                    VALUES (:fecha, :concepto, :monto, :categoria, :comprobante);
                """), {"fecha": fecha_gasto.strftime("%Y-%m-%d"), "concepto": concepto_gasto.upper().strip(), "monto": -abs(monto_gasto), "categoria": categoria_gasto, "comprobante": ""})
            st.success(f"✅ Gasto '{concepto_gasto}' registrado.")
            st.rerun()

with tabs[3]:
    st.subheader("📈 Análisis de Facturación por Días y Curva Mensual")
    with engine.connect() as conn:
        df_cierres = pd.read_sql_query(text("SELECT * FROM cierres_diarios ORDER BY fecha ASC"), conn)
        df_trans = pd.read_sql_query(text("SELECT * FROM transacciones_diarias"), conn)
    
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

with tabs[4]:
    st.subheader("⚙️ Histórico Completo de Cierres Diarios")
    with engine.connect() as conn:
        df_all = pd.read_sql_query(text("SELECT * FROM cierres_diarios ORDER BY fecha DESC"), conn)
    
    if not df_all.empty:
        df_edit_hist = st.data_editor(df_all, num_rows="dynamic", use_container_width=True, key="historico_editor")
        if st.button("💾 Guardar Cambios en Histórico"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM cierres_diarios;"))
                for _, r in df_edit_hist.iterrows():
                    if db_type == "postgres":
                        conn.execute(text("""
                            INSERT INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                            VALUES (:fecha, :efectivo, :nequi, :daviplata, :banco, :total, :obs);
                        """), {"fecha": str(r['fecha']), "efectivo": float(r['efectivo']), "nequi": float(r['nequi']), "daviplata": float(r['daviplata']), "banco": float(r['banco']), "total": float(r['total']), "obs": str(r['observaciones']) if r['observaciones'] else ''})
                    else:
                        conn.execute(text("INSERT INTO cierres_diarios VALUES (:fecha, :efectivo, :nequi, :daviplata, :banco, :total, :obs);"),
                                     {"fecha": str(r['fecha']), "efectivo": float(r['efectivo']), "nequi": float(r['nequi']), "daviplata": float(r['daviplata']), "banco": float(r['banco']), "total": float(r['total']), "obs": str(r['observaciones']) if r['observaciones'] else ''})
            st.success("✅ Cambios guardados en la base de datos.")
