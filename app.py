import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import re
from datetime import datetime, date
import plotly.express as px
import plotly.graph_objects as go
import io

# Generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y FORZADO DE TEMA CLARO ABSOLUTO
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control de Caja - Papelería",
    page_icon="📚",
    layout="wide"
)

# Estilos CSS inyectados para forzar el tema claro en todos los componentes
st.markdown("""
<style>
    /* Fondo blanco absoluto */
    html, body, [data-testid="stAppViewContainer"], .main, .stApp {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }

    /* Visibilidad de Pestañas (Tabs) superiores */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f5f9 !important;
        padding: 8px;
        border-radius: 12px;
        border: 1px solid #cbd5e1;
    }

    .stTabs [data-baseweb="tab"] {
        height: 42px;
        border-radius: 8px;
        color: #334155 !important;
        font-weight: 600 !important;
        background-color: transparent !important;
    }

    .stTabs [data-baseweb="tab"]:hover {
        color: #0284c7 !important;
        background-color: #e0f2fe !important;
    }

    .stTabs [aria-selected="true"] {
        background-color: #0284c7 !important;
        color: #ffffff !important;
    }

    .stTabs [data-baseweb="tab"] div p {
        color: inherit !important;
        font-size: 15px !important;
        font-weight: 600 !important;
    }

    /* Selectbox y campos de entrada en fondo claro */
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    .stSelectbox div[data-baseweb="select"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }

    div[data-baseweb="popover"], div[role="listbox"], ul[role="listbox"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }

    li[role="option"] {
        color: #0f172a !important;
        background-color: #ffffff !important;
    }

    li[role="option"]:hover {
        background-color: #e0f2fe !important;
        color: #0369a1 !important;
    }

    /* Banner Superior */
    .header-banner {
        background: linear-gradient(135deg, #e0f2fe 0%, #dcfce7 100%);
        padding: 22px;
        border-radius: 14px;
        border: 1px solid #bae6fd;
        margin-bottom: 20px;
    }
    
    .header-banner h1 {
        color: #0369a1 !important;
        margin: 0;
        font-size: 26px;
        font-weight: 700;
    }
    
    .header-banner p {
        color: #047857 !important;
        margin: 4px 0 0 0;
        font-size: 14px;
    }

    /* Cajas de Métricas Claras */
    .metric-card-green {
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 10px;
        padding: 14px;
        text-align: center;
    }
    .metric-card-blue {
        background-color: #f0f9ff;
        border: 1px solid #bae6fd;
        border-radius: 10px;
        padding: 14px;
        text-align: center;
    }
    .metric-card-title {
        font-size: 13px;
        color: #475569 !important;
        font-weight: 600;
        margin-bottom: 4px;
    }
    .metric-card-val {
        font-size: 20px;
        font-weight: 700;
        color: #0f172a !important;
    }

    /* Botones Verde/Azul */
    div.stButton > button {
        background-color: #10b981 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 8px 18px !important;
    }
    div.stButton > button:hover {
        background-color: #059669 !important;
    }

    /* Textos generales */
    h1, h2, h3, h4, h5, h6, p, label, span {
        color: #0f172a !important;
    }
</style>
""", unsafe_allow_html=True)

DB_NAME = "papeleria.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS cierres_diarios (
            fecha TEXT PRIMARY KEY,
            efectivo REAL,
            nequi REAL,
            daviplata REAL,
            banco REAL,
            total REAL,
            observaciones TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS transacciones_diarias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            concepto TEXT,
            cantidad INTEGER,
            codigo INTEGER,
            ingreso REAL,
            gastos REAL,
            saldo REAL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS gastos_mensuales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            concepto TEXT,
            monto REAL,
            categoria TEXT,
            comprobante TEXT
        )
    """)
    conn.commit()
    conn.close()

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

    # Resumen
    summary_data = [
        [Paragraph("Efectivo", header_cell), Paragraph("Nequi", header_cell), Paragraph("Daviplata/Banco", header_cell), Paragraph("Total Caja", header_cell), Paragraph("Liquidez Neto", header_cell)],
        [Paragraph(f"${tot_efectivo:,.0f}", cell_bold), Paragraph(f"${tot_nequi:,.0f}", cell_bold), Paragraph(f"${tot_davi:,.0f}", cell_bold), Paragraph(f"${tot_caja:,.0f}", cell_bold), Paragraph(f"${liquidez_neta:,.0f}", cell_bold)]
    ]
    t_summary = Table(summary_data, colWidths=[100, 100, 110, 110, 120])
    t_summary.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#10b981')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#f0fdf4')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#bbf7d0')),
    ]))
    story.append(t_summary)
    story.append(Spacer(1, 10))

    # Recuento Diario
    story.append(Paragraph("Tabla Recuento Diario de Caja", h2_style))
    daily_rows = [[Paragraph("Fecha", header_cell), Paragraph("Efectivo", header_cell), Paragraph("Nequi", header_cell), Paragraph("Daviplata / Banco", header_cell), Paragraph("Total Día", header_cell)]]
    for idx, r in df_cierres_mes.iterrows():
        daily_rows.append([
            Paragraph(str(r['fecha']), cell_style),
            Paragraph(f"${r['efectivo']:,.0f}", cell_style),
            Paragraph(f"${r['nequi']:,.0f}", cell_style),
            Paragraph(f"${r['daviplata']:,.0f}", cell_style),
            Paragraph(f"${r['total']:,.0f}", cell_bold)
        ])
    daily_rows.append([
        Paragraph("TOTAL", header_cell),
        Paragraph(f"${tot_efectivo:,.0f}", header_cell),
        Paragraph(f"${tot_nequi:,.0f}", header_cell),
        Paragraph(f"${tot_davi:,.0f}", header_cell),
        Paragraph(f"${tot_caja:,.0f}", header_cell)
    ])

    t_daily = Table(daily_rows, colWidths=[110, 110, 110, 110, 100])
    t_daily.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0284c7')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#f8fafc')]),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#0369a1')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0')),
    ]))
    story.append(t_daily)

    # Gastos
    if not df_gastos_mes.empty:
        story.append(Spacer(1, 10))
        story.append(Paragraph("Gastos Mensuales y Compras (No Diarios)", h2_style))
        gastos_rows = [[Paragraph("Fecha", header_cell), Paragraph("Concepto", header_cell), Paragraph("Monto ($)", header_cell)]]
        for idx, r in df_gastos_mes.iterrows():
            gastos_rows.append([
                Paragraph(str(r['fecha']), cell_style),
                Paragraph(str(r['concepto']), cell_style),
                Paragraph(f"${r['monto']:,.0f}", cell_bold)
            ])
        gastos_rows.append([
            Paragraph("TOTAL GASTOS", header_cell),
            Paragraph("", header_cell),
            Paragraph(f"${tot_gastos:,.0f}", header_cell)
        ])
        t_gastos = Table(gastos_rows, colWidths=[120, 270, 150])
        t_gastos.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e11d48')),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('ALIGN', (2,0), (2,-1), 'RIGHT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#fff1f2')]),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#be123c')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#fecdd3')),
        ]))
        story.append(t_gastos)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# PARSER DE EXCEL
# -----------------------------------------------------------------------------
def clean_date(val, filename_date=None):
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
            
    match = re.search(r'(\d{1,2})[/\-](\d{1,2})\d*[/\-](\d{4})', val_str)
    if match:
        d, m, y = match.groups()
        try:
            return datetime(int(y), int(m), int(d)).strftime("%Y-%m-%d")
        except ValueError:
            pass
            
    if filename_date:
        return filename_date
        
    return date.today().strftime("%Y-%m-%d")

def parse_daily_excel(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet = wb.active
    
    data = []
    for r in range(1, sheet.max_row + 1):
        row_vals = [sheet.cell(row=r, column=c).value for c in range(1, sheet.max_column + 1)]
        data.append(row_vals)
        
    header_idx = None
    for idx, row in enumerate(data):
        row_str = [str(x).upper() if x is not None else '' for x in row]
        if 'FECHA' in row_str and 'CONCEPTO' in row_str:
            header_idx = idx
            break
            
    if header_idx is None:
        return None, "No se encontró el formato estándar de transacciones en este Excel."
        
    header = [str(x).strip().upper() if x is not None else '' for x in data[header_idx]]
    col_fecha = header.index('FECHA') if 'FECHA' in header else 1
    col_concepto = header.index('CONCEPTO') if 'CONCEPTO' in header else 2
    col_cantidad = header.index('CANTIDAD') if 'CANTIDAD' in header else 3
    
    col_codigo = 4
    for kw in ['CODIGO', 'CÓDIGO', 'NUMERO', 'NÚMERO']:
        if kw in header:
            col_codigo = header.index(kw)
            break
            
    col_ingreso = header.index('INGRESO') if 'INGRESO' in header else 5
    col_gastos = header.index('GASTOS') if 'GASTOS' in header else 6
    col_saldo = header.index('SALDO') if 'SALDO' in header else 7
    
    filename = uploaded_file.name
    filename_date_match = re.search(r'(\d{2}-\d{2}-\d{4})', filename)
    filename_date = None
    if filename_date_match:
        try:
            filename_date = datetime.strptime(filename_date_match.group(1), "%d-%m-%Y").strftime("%Y-%m-%d")
        except:
            pass

    extracted_date = None
    transactions = []
    
    for row in data[header_idx+1:]:
        if row[col_concepto] is None and row[col_ingreso] is None and row[col_gastos] is None:
            continue
        code = row[col_codigo]
        if code is None or str(code).strip() == '':
            continue
            
        fecha_val = row[col_fecha]
        if fecha_val is not None and str(fecha_val).strip() != '' and extracted_date is None:
            extracted_date = clean_date(fecha_val, filename_date)
            
        concepto = str(row[col_concepto]).strip() if row[col_concepto] else ''
        try:
            cant = int(row[col_cantidad]) if row[col_cantidad] is not None else 1
        except:
            cant = 1
            
        ingreso = float(row[col_ingreso]) if row[col_ingreso] is not None else 0.0
        gastos = float(row[col_gastos]) if row[col_gastos] is not None else 0.0
        saldo = float(row[col_saldo]) if row[col_saldo] is not None else 0.0
        
        transactions.append({
            'concepto': concepto,
            'cantidad': cant,
            'codigo': int(code) if str(code).isdigit() else 1,
            'ingreso': ingreso,
            'gastos': gastos,
            'saldo': saldo
        })
        
    if not extracted_date:
        extracted_date = filename_date if filename_date else date.today().strftime("%Y-%m-%d")
        
    saldo_efectivo = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 1)
    saldo_daviplata = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 2)
    saldo_nequi = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 3)
    saldo_banco = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 4)
    total_caja = saldo_efectivo + saldo_daviplata + saldo_nequi + saldo_banco
    
    summary = {
        'fecha': extracted_date,
        'efectivo': saldo_efectivo,
        'daviplata': saldo_daviplata,
        'nequi': saldo_nequi,
        'banco': saldo_banco,
        'total': total_caja
    }
    
    return summary, pd.DataFrame(transactions)

# -----------------------------------------------------------------------------
# INTERFAZ PRINCIPAL
# -----------------------------------------------------------------------------
st.markdown("""
<div class="header-banner">
    <h1>📚 Sistema de Control de Caja y Contabilidad</h1>
    <p>Gestión remota de papelería — Interfaz clara, limpia e intuitiva</p>
</div>
""", unsafe_allow_html=True)

tabs = st.tabs([
    "📥 Cargar Caja Diaria", 
    "📊 Recuento Mensual", 
    "💸 Gastos Mensuales / Compras", 
    "📈 Métricas de Facturación",
    "⚙️ Histórico y Edición"
])

# -----------------------------------------------------------------------------
# PESTAÑA 1: CARGAR CAJA DIARIA
# -----------------------------------------------------------------------------
with tabs[0]:
    st.subheader("📥 Cargar reporte diario enviado desde el local")
    uploaded_file = st.file_uploader("Adjunta el archivo Excel diario (ej. 28-07-2026.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        summary, df_trans = parse_daily_excel(uploaded_file)
        
        if isinstance(df_trans, str):
            st.error(df_trans)
        else:
            st.success("✅ Archivo procesado correctamente.")
            
            col_f1, _ = st.columns([1, 2])
            with col_f1:
                fecha_final = st.date_input(
                    "Fecha del Cierre:", 
                    value=datetime.strptime(summary['fecha'], "%Y-%m-%d").date()
                )
            fecha_str = fecha_final.strftime("%Y-%m-%d")
            
            st.markdown("##### Saldos Extraídos del Día:")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💵 Efectivo</div><div class="metric-card-val">${summary["efectivo"]:,.0f}</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">🟣 Nequi</div><div class="metric-card-val">${summary["nequi"]:,.0f}</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="metric-card-green"><div class="metric-card-title">🔴 Daviplata</div><div class="metric-card-val">${summary["daviplata"]:,.0f}</div></div>', unsafe_allow_html=True)
            c4.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">🏦 Banco</div><div class="metric-card-val">${summary["banco"]:,.0f}</div></div>', unsafe_allow_html=True)
            c5.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💰 Total Día</div><div class="metric-card-val">${summary["total"]:,.0f}</div></div>', unsafe_allow_html=True)
            
            st.write("")
            st.markdown("##### Detalle de Transacciones (Edición manual habilitada):")
            df_edited = st.data_editor(
                df_trans,
                num_rows="dynamic",
                column_config={
                    "codigo": st.column_config.SelectboxColumn(
                        "Código Pago", 
                        options=[1, 2, 3, 4], 
                        help="1: Efectivo | 2: Daviplata | 3: Nequi | 4: Banco"
                    ),
                    "ingreso": st.column_config.NumberColumn("Ingreso ($)", format="$%d"),
                    "gastos": st.column_config.NumberColumn("Gastos ($)", format="$%d"),
                    "saldo": st.column_config.NumberColumn("Saldo Acumulado ($)", format="$%d")
                },
                use_container_width=True
            )
            
            if st.button("💾 Guardar e Integrar en la Caja Mensual", type="primary"):
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                
                eff = sum(row['ingreso'] - row['gastos'] for _, row in df_edited.iterrows() if row['codigo'] == 1)
                dav = sum(row['ingreso'] - row['gastos'] for _, row in df_edited.iterrows() if row['codigo'] == 2)
                neq = sum(row['ingreso'] - row['gastos'] for _, row in df_edited.iterrows() if row['codigo'] == 3)
                ban = sum(row['ingreso'] - row['gastos'] for _, row in df_edited.iterrows() if row['codigo'] == 4)
                tot = eff + dav + neq + ban
                
                c.execute("""
                    INSERT OR REPLACE INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (fecha_str, eff, neq, dav, ban, tot, f"Cargado desde {uploaded_file.name}"))
                
                c.execute("DELETE FROM transacciones_diarias WHERE fecha = ?", (fecha_str,))
                for _, row in df_edited.iterrows():
                    c.execute("""
                        INSERT INTO transacciones_diarias (fecha, concepto, cantidad, codigo, ingreso, gastos, saldo)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (fecha_str, row['concepto'], row['cantidad'], row['codigo'], row['ingreso'], row['gastos'], row['saldo']))
                    
                conn.commit()
                conn.close()
                st.success(f"🎉 ¡Cierre del día {fecha_str} guardado exitosamente!")

# -----------------------------------------------------------------------------
# PESTAÑA 2: RECUENTO MENSUAL (TABLAS EN HTML BLANCO PURO)
# -----------------------------------------------------------------------------
with tabs[1]:
    st.subheader("📊 Recuento de Caja Mensual")
    
    conn = sqlite3.connect(DB_NAME)
    df_cierres = pd.read_sql_query("SELECT * FROM cierres_diarios ORDER BY fecha ASC", conn)
    df_gastos = pd.read_sql_query("SELECT * FROM gastos_mensuales ORDER BY fecha ASC", conn)
    conn.close()
    
    if df_cierres.empty:
        st.info("Aún no hay cierres cargados. Sube un archivo en la primera pestaña.")
    else:
        df_cierres['mes_año'] = pd.to_datetime(df_cierres['fecha']).dt.strftime('%Y-%m')
        meses_disponibles = df_cierres['mes_año'].unique()
        
        col_m1, col_m2 = st.columns([1, 1])
        with col_m1:
            mes_sel = st.selectbox("Selecciona el Mes:", meses_disponibles, index=len(meses_disponibles)-1)
            
        df_c_mes = df_cierres[df_cierres['mes_año'] == mes_sel].copy()
        
        tot_efectivo = df_c_mes['efectivo'].sum()
        tot_nequi = df_c_mes['nequi'].sum()
        tot_davi = df_c_mes['daviplata'].sum()
        tot_banco = df_c_mes['banco'].sum()
        tot_caja = df_c_mes['total'].sum()
        
        if not df_gastos.empty:
            df_gastos['mes_año'] = pd.to_datetime(df_gastos['fecha']).dt.strftime('%Y-%m')
            df_g_mes = df_gastos[df_gastos['mes_año'] == mes_sel].copy()
            tot_gastos_grandes = df_g_mes['monto'].sum()
        else:
            df_g_mes = pd.DataFrame()
            tot_gastos_grandes = 0.0
            
        liquidez_neta = tot_caja - abs(tot_gastos_grandes)

        # Generar archivo PDF para descarga
        pdf_bytes = generate_pdf_report(
            mes_sel, df_c_mes, df_g_mes, 
            tot_efectivo, tot_nequi, tot_davi, tot_caja, tot_gastos_grandes, liquidez_neta
        )
        
        with col_m2:
            st.write("")
            st.download_button(
                label="📄 Descargar Recuento Mensual en PDF",
                data=pdf_bytes,
                file_name=f"Recuento_Caja_{mes_sel}.pdf",
                mime="application/pdf",
                type="primary"
            )
        
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
            st.markdown("##### Tabla Recuento Diario (Estructura Excel):")
            
            # Construcción de la Tabla HTML con Fondo Blanco Garantizado
            html_daily = """
            <div style="overflow-x:auto; margin-top:8px;">
            <table style="width:100%; border-collapse:collapse; background-color:#ffffff; color:#0f172a; border:1px solid #cbd5e1; font-family:sans-serif; border-radius:8px;">
              <thead>
                <tr style="background-color:#0284c7; color:#ffffff; font-weight:bold; text-align:center;">
                  <th style="padding:10px; border:1px solid #cbd5e1;">Fecha</th>
                  <th style="padding:10px; border:1px solid #cbd5e1;">Efectivo</th>
                  <th style="padding:10px; border:1px solid #cbd5e1;">Nequi</th>
                  <th style="padding:10px; border:1px solid #cbd5e1;">Daviplata / Banco</th>
                  <th style="padding:10px; border:1px solid #cbd5e1;">Total</th>
                </tr>
              </thead>
              <tbody>
            """
            for idx, r in df_c_mes.reset_index().iterrows():
                bg = "#ffffff" if idx % 2 == 0 else "#f8fafc"
                html_daily += f"""
                <tr style="background-color:{bg}; text-align:center; color:#0f172a;">
                  <td style="padding:8px; border:1px solid #e2e8f0; font-weight:500;">{r['fecha']}</td>
                  <td style="padding:8px; border:1px solid #e2e8f0;">${r['efectivo']:,.0f}</td>
                  <td style="padding:8px; border:1px solid #e2e8f0;">${r['nequi']:,.0f}</td>
                  <td style="padding:8px; border:1px solid #e2e8f0;">${r['daviplata']:,.0f}</td>
                  <td style="padding:8px; border:1px solid #e2e8f0; font-weight:bold; color:#0369a1;">${r['total']:,.0f}</td>
                </tr>
                """
            html_daily += f"""
                <tr style="background-color:#e0f2fe; color:#0369a1; font-weight:bold; text-align:center;">
                  <td style="padding:10px; border:1px solid #bae6fd;">TOTAL</td>
                  <td style="padding:10px; border:1px solid #bae6fd;">${tot_efectivo:,.0f}</td>
                  <td style="padding:10px; border:1px solid #bae6fd;">${tot_nequi:,.0f}</td>
                  <td style="padding:10px; border:1px solid #bae6fd;">${tot_davi:,.0f}</td>
                  <td style="padding:10px; border:1px solid #bae6fd; font-size:16px;">${tot_caja:,.0f}</td>
                </tr>
              </tbody>
            </table>
            </div>
            """
            st.markdown(html_daily, unsafe_allow_html=True)
            
        with col_t2:
            st.markdown("##### Gastos Mensuales y Mercancía:")
            if df_g_mes.empty:
                st.info("No hay gastos mayores ingresados en este mes.")
            else:
                html_gastos = """
                <div style="overflow-x:auto; margin-top:8px;">
                <table style="width:100%; border-collapse:collapse; background-color:#ffffff; color:#0f172a; border:1px solid #cbd5e1; font-family:sans-serif; border-radius:8px;">
                  <thead>
                    <tr style="background-color:#e11d48; color:#ffffff; font-weight:bold; text-align:left;">
                      <th style="padding:10px; border:1px solid #cbd5e1;">Fecha</th>
                      <th style="padding:10px; border:1px solid #cbd5e1;">Concepto</th>
                      <th style="padding:10px; border:1px solid #cbd5e1; text-align:right;">Monto ($)</th>
                    </tr>
                  </thead>
                  <tbody>
                """
                for idx, r in df_g_mes.reset_index().iterrows():
                    bg = "#ffffff" if idx % 2 == 0 else "#fff1f2"
                    html_gastos += f"""
                    <tr style="background-color:{bg}; color:#0f172a;">
                      <td style="padding:8px; border:1px solid #fecdd3; font-weight:500;">{r['fecha']}</td>
                      <td style="padding:8px; border:1px solid #fecdd3;">{r['concepto']}</td>
                      <td style="padding:8px; border:1px solid #fecdd3; text-align:right; font-weight:bold; color:#be123c;">${r['monto']:,.0f}</td>
                    </tr>
                    """
                html_gastos += f"""
                    <tr style="background-color:#ffe4e6; color:#be123c; font-weight:bold;">
                      <td style="padding:10px; border:1px solid #fecdd3;">TOTAL GASTOS</td>
                      <td style="padding:10px; border:1px solid #fecdd3;"></td>
                      <td style="padding:10px; border:1px solid #fecdd3; text-align:right; font-size:15px;">${tot_gastos_grandes:,.0f}</td>
                    </tr>
                  </tbody>
                </table>
                </div>
                """
                st.markdown(html_gastos, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# PESTAÑA 3: GASTOS MENSUALES
# -----------------------------------------------------------------------------
with tabs[2]:
    st.subheader("💸 Registrar Gasto Mensual o Compra Grande")
    st.write("Agrega aquí arriendo, servicios, proveedores o pagos mayores.")
    
    with st.form("form_gastos_clean", clear_on_submit=True):
        col_g1, col_g2, col_g3 = st.columns(3)
        with col_g1:
            fecha_gasto = st.date_input("Fecha:", value=date.today())
        with col_g2:
            categoria_gasto = st.selectbox("Categoría:", [
                "Arriendo", "Servicios (Agua/Luz/Gas)", "Internet/Teléfono", 
                "Mercancía/Proveedores", "Pago Deuda/Terceros", "Otro"
            ])
        with col_g3:
            monto_gasto = st.number_input("Monto en Pesos ($):", min_value=0.0, step=1000.0)
            
        concepto_gasto = st.text_input("Concepto (ej. RESMAS Y TINTA EPSON, ARRIENDO LOCAL):")
        
        submitted = st.form_submit_button("➕ Registrar Gasto Mensual", type="primary")
        
        if submitted:
            if monto_gasto > 0 and concepto_gasto.strip():
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                c.execute("""
                    INSERT INTO gastos_mensuales (fecha, concepto, monto, categoria, comprobante)
                    VALUES (?, ?, ?, ?, ?)
                """, (
                    fecha_gasto.strftime("%Y-%m-%d"), 
                    concepto_gasto.upper().strip(), 
                    -abs(monto_gasto), 
                    categoria_gasto, 
                    ""
                ))
                conn.commit()
                conn.close()
                st.success(f"✅ Gasto '{concepto_gasto}' por ${monto_gasto:,.0f} registrado.")
                st.rerun()

# -----------------------------------------------------------------------------
# PESTAÑA 4: MÉTRICAS DE FACTURACIÓN
# -----------------------------------------------------------------------------
with tabs[3]:
    st.subheader("📈 Análisis de Facturación por Días y Curva Mensual")
    
    conn = sqlite3.connect(DB_NAME)
    df_cierres = pd.read_sql_query("SELECT * FROM cierres_diarios ORDER BY fecha ASC", conn)
    df_trans = pd.read_sql_query("SELECT * FROM transacciones_diarias", conn)
    conn.close()
    
    if df_cierres.empty:
        st.info("Aún no hay suficientes cierres cargados para generar las gráficas.")
    else:
        df_cierres['mes_año'] = pd.to_datetime(df_cierres['fecha']).dt.strftime('%Y-%m')
        meses_disponibles = df_cierres['mes_año'].unique()
        
        col_m1, _ = st.columns([1, 1])
        with col_m1:
            mes_sel_graf = st.selectbox("Selecciona Mes para ver Facturación Diaria:", meses_disponibles, index=len(meses_disponibles)-1)
            
        df_c_graf = df_cierres[df_cierres['mes_año'] == mes_sel_graf].copy()
        
        # Gráfico 1: Barras por días
        st.markdown("#### 1. Facturación Día a Día")
        fig_bar = px.bar(
            df_c_graf,
            x="fecha",
            y="total",
            text_auto='.2s',
            labels={"fecha": "Día / Fecha", "total": "Facturación Total ($)"},
            title=f"Facturación Diaria en {mes_sel_graf}",
            color_discrete_sequence=['#10b981']
        )
        fig_bar.update_layout(
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=dict(color='#0f172a'),
            xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=True, gridcolor='#f1f5f9')
        )
        st.plotly_chart(fig_bar, use_container_width=True)
        
        max_day = df_c_graf.loc[df_c_graf['total'].idxmax()]
        min_day = df_c_graf.loc[df_c_graf['total'].idxmin()]
        
        c_max, c_min = st.columns(2)
        c_max.success(f"🏆 **Mayor facturación:** {max_day['fecha']} con **${max_day['total']:,.0f}**")
        c_min.warning(f"📉 **Menor facturación:** {min_day['fecha']} con **${min_day['total']:,.0f}**")
        
        st.divider()
        
        # Gráfico 2: Detalle por día
        st.markdown("#### 2. Inspeccionar Detalle de un Día")
        dias_del_mes = df_c_graf['fecha'].tolist()
        dia_inspeccionar = st.selectbox("Selecciona un día específico:", dias_del_mes, index=len(dias_del_mes)-1)
        
        df_t_dia = df_trans[df_trans['fecha'] == dia_inspeccionar]
        
        if df_t_dia.empty:
            st.write("No hay detalle de ítems registrado para este día.")
        else:
            col_d1, col_d2 = st.columns([2, 1])
            with col_d1:
                st.markdown(f"**Movimientos del {dia_inspeccionar}:**")
                st.dataframe(
                    df_t_dia[['concepto', 'cantidad', 'codigo', 'ingreso', 'gastos']],
                    column_config={
                        "concepto": "Concepto / Ítem",
                        "cantidad": "Cant.",
                        "codigo": "Cód. Caja",
                        "ingreso": st.column_config.NumberColumn("Ingreso ($)", format="$%d"),
                        "gastos": st.column_config.NumberColumn("Gasto ($)", format="$%d")
                    },
                    use_container_width=True,
                    hide_index=True
                )
            with col_d2:
                eff_d = df_t_dia[df_t_dia['codigo']==1]['ingreso'].sum() - df_t_dia[df_t_dia['codigo']==1]['gastos'].sum()
                neq_d = df_t_dia[df_t_dia['codigo']==3]['ingreso'].sum() - df_t_dia[df_t_dia['codigo']==3]['gastos'].sum()
                
                st.markdown("**Composición del Cobro:**")
                st.write(f"💵 **Efectivo:** ${eff_d:,.0f}")
                st.write(f"🟣 **Nequi:** ${neq_d:,.0f}")
                st.write(f"📊 **Total Día:** ${(eff_d+neq_d):,.0f}")
                
        st.divider()
        
        # Gráfico 3: Curva por Meses
        st.markdown("#### 3. Curva Evolutiva por Meses")
        df_mensual_sum = df_cierres.groupby('mes_año')['total'].sum().reset_index()
        
        fig_curve = px.line(
            df_mensual_sum,
            x="mes_año",
            y="total",
            markers=True,
            labels={"mes_año": "Mes", "total": "Facturación Acumulada ($)"},
            title="Evolución de Facturación Mes a Mes",
            color_discrete_sequence=['#0284c7']
        )
        fig_curve.update_layout(
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=dict(color='#0f172a'),
            xaxis=dict(showgrid=True, gridcolor='#f1f5f9'),
            yaxis=dict(showgrid=True, gridcolor='#f1f5f9')
        )
        st.plotly_chart(fig_curve, use_container_width=True)

# -----------------------------------------------------------------------------
# PESTAÑA 5: HISTÓRICO
# -----------------------------------------------------------------------------
with tabs[4]:
    st.subheader("⚙️ Histórico Completo de Cierres Diarios")
    
    conn = sqlite3.connect(DB_NAME)
    df_all = pd.read_sql_query("SELECT * FROM cierres_diarios ORDER BY fecha DESC", conn)
    conn.close()
    
    if not df_all.empty:
        st.write("Edita directamente cualquier valor histórico si necesitas corregir datos.")
        df_edit_hist = st.data_editor(
            df_all,
            num_rows="dynamic",
            use_container_width=True,
            key="historico_editor"
        )
        
        if st.button("💾 Guardar Cambios en Histórico"):
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("DELETE FROM cierres_diarios")
            for _, r in df_edit_hist.iterrows():
                c.execute("""
                    INSERT INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (r['fecha'], r['efectivo'], r['nequi'], r['daviplata'], r['banco'], r['total'], r['observaciones']))
            conn.commit()
            conn.close()
            st.success("✅ Cambios guardados en la base de datos.")
