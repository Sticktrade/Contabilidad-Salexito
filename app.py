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
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
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
# CONEXIÓN A BASE DE DATOS
# -----------------------------------------------------------------------------
@st.cache_resource
def get_db_engine():
    if "SUPABASE_URL" in st.secrets and st.secrets["SUPABASE_URL"].strip():
        db_url = st.secrets["SUPABASE_URL"].strip()
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)
        try:
            eng = create_engine(db_url, pool_pre_ping=True, pool_size=5, max_overflow=10)
            with eng.connect() as conn: conn.execute(text("SELECT 1;"))
            return eng, "postgres", None
        except Exception as e:
            return create_engine("sqlite:///papeleria.db"), "sqlite", str(e)
    else:
        return create_engine("sqlite:///papeleria.db"), "sqlite", "Sin configurar SUPABASE_URL en Secrets."

engine, db_type, conn_error = get_db_engine()

@st.cache_resource
def init_db():
    try:
        with engine.begin() as conn:
            if db_type == "postgres":
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS cierres_diarios (fecha VARCHAR(20) PRIMARY KEY, efectivo FLOAT, nequi FLOAT, daviplata FLOAT, banco FLOAT, total FLOAT, observaciones TEXT);
                    CREATE TABLE IF NOT EXISTS transacciones_diarias (id SERIAL PRIMARY KEY, fecha VARCHAR(20), concepto TEXT, cantidad INT, codigo INT, ingreso FLOAT, gastos FLOAT, saldo FLOAT);
                    CREATE TABLE IF NOT EXISTS gastos_mensuales (id SERIAL PRIMARY KEY, fecha VARCHAR(20), concepto TEXT, monto FLOAT, categoria VARCHAR(100), comprobante TEXT);
                    CREATE TABLE IF NOT EXISTS presupuestos_mensuales (id SERIAL PRIMARY KEY, mes_año VARCHAR(20), concepto TEXT, monto_inicial FLOAT, monto_final FLOAT, categoria VARCHAR(100));
                """))
            else:
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS cierres_diarios (fecha TEXT PRIMARY KEY, efectivo REAL, nequi REAL, daviplata REAL, banco REAL, total REAL, observaciones TEXT);
                    CREATE TABLE IF NOT EXISTS transacciones_diarias (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, concepto TEXT, cantidad INTEGER, codigo INTEGER, ingreso REAL, gastos REAL, saldo REAL);
                    CREATE TABLE IF NOT EXISTS gastos_mensuales (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, concepto TEXT, monto REAL, categoria TEXT, comprobante TEXT);
                    CREATE TABLE IF NOT EXISTS presupuestos_mensuales (id INTEGER PRIMARY KEY AUTOINCREMENT, mes_año TEXT, concepto TEXT, monto_inicial REAL, monto_final REAL, categoria TEXT);
                """))
    except Exception as ex: st.warning(f"Error inicializando tablas: {ex}")

init_db()

# -----------------------------------------------------------------------------
# CONSULTAS DE DATOS
# -----------------------------------------------------------------------------
@st.cache_data(ttl=60)
def load_cierres_diarios():
    with engine.connect() as conn: return pd.read_sql_query(text("SELECT * FROM cierres_diarios ORDER BY fecha ASC"), conn)

@st.cache_data(ttl=60)
def load_gastos_ligeros():
    with engine.connect() as conn:
        return pd.read_sql_query(text("SELECT id, fecha, concepto, monto, categoria, CASE WHEN comprobante IS NOT NULL AND length(comprobante) > 5 THEN 1 ELSE 0 END as tiene_comprobante FROM gastos_mensuales ORDER BY fecha ASC, id ASC"), conn)

@st.cache_data(ttl=60)
def load_presupuestos_mensuales():
    with engine.connect() as conn:
        return pd.read_sql_query(text("SELECT id, mes_año, concepto, monto_inicial, monto_final, categoria FROM presupuestos_mensuales ORDER BY mes_año DESC, id ASC"), conn)

# -----------------------------------------------------------------------------
# GENERACIÓN DE PDF DE PRESUPUESTO
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
# ESTRUCTURA DE LA APLICACIÓN
# -----------------------------------------------------------------------------
st.markdown("""
<div class="header-banner">
    <h1>📚 Sistema de Control de Caja y Presupuestos Mensuales</h1>
    <p>Gestión y Administración Financiera Remota</p>
</div>
""", unsafe_allow_html=True)

tabs = st.tabs(["📋 Presupuesto Mensual (Formulario)", "📊 Recuento Mensual", "💸 Gastos Mensuales", "📈 Métricas", "⚙️ Histórico"])

# PESTAÑA PRINCIPAL DE PRESUPUESTOS (FORMULARIO Y TABLA DESCARGABLE)
with tabs[0]:
    st.subheader("📋 Formulario y Tabla de Presupuesto Mensual")
    st.write("Ingresa los conceptos y su **Valor Inicial** a principio de mes. Al cierre del mes, completa el **Valor Final** para comparar lo presupuestado vs lo ejecutado.")
    
    df_pres_all = load_presupuestos_mensuales()
    meses_existentes = sorted(list(set(df_pres_all['mes_año'].unique().tolist() + [date.today().strftime('%Y-%m')])), reverse=True)
    
    col_p1, col_p2 = st.columns([1, 1])
    with col_p1:
        mes_sel = st.selectbox("Selecciona o Crea el Mes de Trabajo:", meses_existentes, key="select_mes_form")
    
    df_mes = df_pres_all[df_pres_all['mes_año'] == mes_sel].copy()
    
    # 1. FORMULARIO PARA AÑADIR NUEVOS CAMPOS AL PRESUPUESTO
    st.divider()
    st.markdown("##### ➕ Formulario: Agregar Nuevo Rubro al Presupuesto")
    with st.form("form_add_budget_item", clear_on_submit=True):
        col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
        with col_f1: new_concepto = st.text_input("Concepto / Rubro (ej. Renta, Luz, Mercancía):")
        with col_f2: new_categoria = st.selectbox("Categoría:", ["Gastos Fijos", "Servicios", "Impuestos", "Deudas", "Personal", "Mercancía", "Inversión/Equipo", "Otro"])
        with col_f3: new_val_init = st.number_input("Valor Inicial ($):", min_value=0.0, step=5000.0)
        
        btn_add = st.form_submit_button("➕ Añadir Campo al Presupuesto", type="primary")
        if btn_add and new_concepto.strip():
            with engine.begin() as conn:
                conn.execute(text("""
                    INSERT INTO presupuestos_mensuales (mes_año, concepto, monto_inicial, monto_final, categoria)
                    VALUES (:mes, :c, :mi, :mf, :cat);
                """), {"mes": mes_sel, "c": new_concepto.strip().upper(), "mi": new_val_init, "mf": 0.0, "cat": new_categoria})
            st.cache_data.clear()
            st.success(f"✅ Rubro '{new_concepto}' añadido al presupuesto del mes {mes_sel}.")
            st.rerun()

    # 2. TABLA INTERACTIVA DE PRESUPUESTO
    st.divider()
    st.markdown(f"##### 📝 Tabla de Presupuesto del Mes — {mes_sel}")
    
    if df_mes.empty:
        st.info("Aún no has agregado rubros para este mes. Usa el formulario de arriba o carga la plantilla base.")
        if st.button("✨ Cargar Plantilla Estándar", type="primary"):
            plantilla = [
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
                for c_nom, m_ini, m_fin, c_cat in plantilla:
                    conn.execute(text("""
                        INSERT INTO presupuestos_mensuales (mes_año, concepto, monto_inicial, monto_final, categoria)
                        VALUES (:mes, :c, :mi, :mf, :cat);
                    """), {"mes": mes_sel, "c": c_nom, "mi": m_ini, "mf": m_fin, "cat": c_cat})
            st.cache_data.clear()
            st.rerun()
    else:
        tot_ini = df_mes['monto_inicial'].sum()
        tot_fin = df_mes['monto_final'].sum()
        dif_tot = tot_fin - tot_ini

        # Métricas automáticas
        k1, k2, k3 = st.columns(3)
        k1.markdown(f'<div class="metric-card-blue"><div class="metric-card-title">📌 Total Presupuestado (Valor Inicial)</div><div class="metric-card-val">${tot_ini:,.0f}</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="metric-card-green"><div class="metric-card-title">💸 Total Ejecutado (Valor Final)</div><div class="metric-card-val">${tot_fin:,.0f}</div></div>', unsafe_allow_html=True)
        
        diff_class = "metric-card-green" if dif_tot <= 0 else "metric-card-amber"
        diff_txt = f"-${abs(dif_tot):,.0f} (Ahorro)" if dif_tot < 0 else (f"+${dif_tot:,.0f} (Exceso)" if dif_tot > 0 else "$0")
        k3.markdown(f'<div class="{diff_class}"><div class="metric-card-title">⚖️ Diferencia / Variación</div><div class="metric-card-val">{diff_txt}</div></div>', unsafe_allow_html=True)

        st.write("")
        df_mes['diferencia'] = df_mes['monto_final'] - df_mes['monto_inicial']
        
        df_edited = st.data_editor(
            df_mes[['id', 'concepto', 'categoria', 'monto_inicial', 'monto_final', 'diferencia']],
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
            key=f"editor_presupuesto_{mes_sel}"
        )

        col_s1, col_s2 = st.columns([1, 1])
        with col_s1:
            if st.button("💾 Guardar Cambios en la Tabla", type="primary"):
                with engine.begin() as conn:
                    conn.execute(text("DELETE FROM presupuestos_mensuales WHERE mes_año = :mes"), {"mes": mes_sel})
                    for _, r in df_edited.iterrows():
                        c_text = str(r['concepto']).strip() if pd.notna(r['concepto']) else ''
                        if c_text:
                            m_ini = float(r['monto_inicial']) if pd.notna(r['monto_inicial']) else 0.0
                            m_fin = float(r['monto_final']) if pd.notna(r['monto_final']) else 0.0
                            cat_val = str(r['categoria']) if pd.notna(r['categoria']) else 'General'
                            conn.execute(text("""
                                INSERT INTO presupuestos_mensuales (mes_año, concepto, monto_inicial, monto_final, categoria)
                                VALUES (:mes, :c, :mi, :mf, :cat);
                            """), {"mes": mes_sel, "c": c_text, "mi": m_ini, "mf": m_fin, "cat": cat_val})
                st.cache_data.clear()
                st.success("✅ ¡Presupuesto actualizado correctamente!")
                st.rerun()

        # BOTÓN DE DESCARGA EN PDF
        pdf_bytes = generate_presupuesto_pdf(mes_sel, df_mes)
        with col_s2:
            st.download_button(
                label="📄 Descargar Formulario de Presupuesto en PDF",
                data=pdf_bytes,
                file_name=f"Presupuesto_Mensual_{mes_sel}.pdf",
                mime="application/pdf",
                type="primary"
            )

# Resto de pestañas informativas...
with tabs[1]:
    st.subheader("📊 Recuento Mensual de Caja")
    df_cierres = load_cierres_diarios()
    if not df_cierres.empty: st.dataframe(df_cierres, use_container_width=True)

with tabs[2]:
    st.subheader("💸 Gastos Mensuales y Compras")
    df_gastos = load_gastos_ligeros()
    if not df_gastos.empty: st.dataframe(df_gastos, use_container_width=True)

with tabs[3]:
    st.subheader("📈 Métricas")
    st.info("Métricas de ventas e ingresos.")

with tabs[4]:
    st.subheader("⚙️ Histórico")
    st.info("Edición de cierres diarios.")
