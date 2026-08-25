import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import re
from datetime import datetime, date
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA Y TEMA CLARO
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control de Caja - Papelería",
    page_icon="📚",
    layout="wide"
)

# Estilos CSS personalizados para forzar un diseño claro (Fondo Blanco y Acentos Verde/Azul)
st.markdown("""
<style>
    /* Fondo principal blanco */
    .stApp {
        background-color: #ffffff;
        color: #1e293b;
    }
    
    /* Banners y contenedores destacados */
    .header-banner {
        background: linear-gradient(135deg, #e0f2fe 0%, #dcfce7 100%);
        padding: 24px;
        border-radius: 14px;
        border: 1px solid #bae6fd;
        margin-bottom: 20px;
    }
    
    .header-banner h1 {
        color: #0369a1;
        margin: 0;
        font-size: 28px;
        font-weight: 700;
    }
    
    .header-banner p {
        color: #047857;
        margin: 5px 0 0 0;
        font-size: 15px;
    }

    /* Targetas de métricas claras */
    .metric-box-green {
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .metric-box-blue {
        background-color: #f0f9ff;
        border: 1px solid #bae6fd;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .metric-title {
        font-size: 14px;
        color: #64748b;
        font-weight: 600;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 22px;
        font-weight: 700;
        color: #0f172a;
    }

    /* Botones verde claro / azul claro */
    div.stButton > button {
        background-color: #10b981 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 10px 20px !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:hover {
        background-color: #059669 !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1) !important;
    }

    /* Ajuste de solapas (Tabs) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f8fafc;
        padding: 8px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        border-radius: 8px;
        color: #475569;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #e0f2fe !important;
        color: #0369a1 !important;
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
# PARSER DE ARCHIVOS EXCEL DIARIOS
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
        return None, "No se encontró la cabecera estándar (FECHA, CONCEPTO) en el Excel."
        
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
# ENCABEZADO PRINCIPAL
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
            c1.markdown(f'<div class="metric-box-green"><div class="metric-title">💵 Efectivo</div><div class="metric-value">${summary["efectivo"]:,.0f}</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="metric-box-blue"><div class="metric-title">🟣 Nequi</div><div class="metric-value">${summary["nequi"]:,.0f}</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="metric-box-green"><div class="metric-title">🔴 Daviplata</div><div class="metric-value">${summary["daviplata"]:,.0f}</div></div>', unsafe_allow_html=True)
            c4.markdown(f'<div class="metric-box-blue"><div class="metric-title">🏦 Banco</div><div class="metric-value">${summary["banco"]:,.0f}</div></div>', unsafe_allow_html=True)
            c5.markdown(f'<div class="metric-box-green"><div class="metric-title">💰 Total Día</div><div class="metric-value">${summary["total"]:,.0f}</div></div>', unsafe_allow_html=True)
            
            st.write("")
            st.markdown("##### Detalle de Transacciones (Edición manual habilitada por si deseas ajustar valores):")
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
# PESTAÑA 2: RECUENTO MENSUAL (CON FORMATO EXACTO AL EXCEL ORIGINAL)
# -----------------------------------------------------------------------------
with tabs[1]:
    st.subheader("📊 Recuento de Caja Mensual")
    
    conn = sqlite3.connect(DB_NAME)
    df_cierres = pd.read_sql_query("SELECT * FROM cierres_diarios ORDER BY fecha ASC", conn)
    df_gastos = pd.read_sql_query("SELECT * FROM gastos_mensuales ORDER BY fecha ASC", conn)
    conn.close()
    
    if df_cierres.empty:
        st.info("Aún no hay registros de cierres diarios. Carga un archivo en la primera pestaña.")
    else:
        df_cierres['mes_año'] = pd.to_datetime(df_cierres['fecha']).dt.strftime('%Y-%m')
        meses_disponibles = df_cierres['mes_año'].unique()
        
        col_m, _ = st.columns([1, 2])
        with col_m:
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
        
        # Tarjetas de Totales como en el Excel Original
        st.markdown("##### Totales Consolidados del Mes:")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.markdown(f'<div class="metric-box-green"><div class="metric-title">💵 Total Efectivo</div><div class="metric-value">${tot_efectivo:,.0f}</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="metric-box-blue"><div class="metric-title">🟣 Total Nequi</div><div class="metric-value">${tot_nequi:,.0f}</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="metric-box-green"><div class="metric-title">🔴 Total Daviplata</div><div class="metric-value">${tot_davi:,.0f}</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="metric-box-blue"><div class="metric-title">📦 Total Caja Recuento</div><div class="metric-value">${tot_caja:,.0f}</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="metric-box-green"><div class="metric-title">📊 Liquidez Neto</div><div class="metric-value">${liquidez_neta:,.0f}</div></div>', unsafe_allow_html=True)
        
        st.write("")
        col_t1, col_t2 = st.columns([3, 2])
        
        with col_t1:
            st.markdown("##### Tabla Recuento Diario (Estructura Excel):")
            
            # Fila de totales tipo Excel
            df_display = df_c_mes[['fecha', 'efectivo', 'nequi', 'daviplata', 'total']].copy()
            
            # Añadir fila de 'Total' al final
            df_tot_row = pd.DataFrame([{
                'fecha': 'TOTAL',
                'efectivo': tot_efectivo,
                'nequi': tot_nequi,
                'daviplata': tot_davi,
                'total': tot_caja
            }])
            df_table_final = pd.concat([df_display, df_tot_row], ignore_index=True)
            
            st.dataframe(
                df_table_final,
                column_config={
                    "fecha": "Fecha",
                    "efectivo": st.column_config.NumberColumn("Efectivo", format="$%d"),
                    "nequi": st.column_config.NumberColumn("Nequi", format="$%d"),
                    "daviplata": st.column_config.NumberColumn("Daviplata / Banco", format="$%d"),
                    "total": st.column_config.NumberColumn("Total", format="$%d"),
                },
                use_container_width=True,
                hide_index=True
            )
            
        with col_t2:
            st.markdown("##### Gastos Mensuales y Mercancía (No diarios):")
            if df_g_mes.empty:
                st.info("No hay gastos mayores ingresados en este mes.")
            else:
                st.dataframe(
                    df_g_mes[['fecha', 'concepto', 'monto']],
                    column_config={
                        "fecha": "Fecha",
                        "concepto": "Concepto",
                        "monto": st.column_config.NumberColumn("Monto ($)", format="$%d"),
                    },
                    use_container_width=True,
                    hide_index=True
                )

# -----------------------------------------------------------------------------
# PESTAÑA 3: GASTOS MENSUALES / COMPRAS
# -----------------------------------------------------------------------------
with tabs[2]:
    st.subheader("💸 Registrar Gasto Mensual, Factura o Compra Grande")
    st.write("Agrega aquí arriendo, servicios, proveedores o pagos a terceros para mantener limpia la caja diaria.")
    
    with st.form("form_gastos_claros", clear_on_submit=True):
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
# PESTAÑA 4: MÉTRICAS DE FACTURACIÓN Y DETALLES
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
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            mes_sel_graf = st.selectbox("Selecciona Mes para ver Facturación Diaria:", meses_disponibles, index=len(meses_disponibles)-1)
            
        df_c_graf = df_cierres[df_cierres['mes_año'] == mes_sel_graf].copy()
        
        # 1. Gráfico de Barras por Días
        st.markdown("#### 1. Facturación Día a Día (Días de Mayor y Menor Venta)")
        
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
            font=dict(color='#1e293b'),
            xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=True, gridcolor='#f1f5f9')
        )
        st.plotly_chart(fig_bar, use_container_width=True)
        
        # Días destacados
        max_day = df_c_graf.loc[df_c_graf['total'].idxmax()]
        min_day = df_c_graf.loc[df_c_graf['total'].idxmin()]
        
        c_max, c_min = st.columns(2)
        c_max.success(f"🏆 **Día con MAYOR facturación:** {max_day['fecha']} con **${max_day['total']:,.0f}**")
        c_min.warning(f"📉 **Día con MENOR facturación:** {min_day['fecha']} con **${min_day['total']:,.0f}**")
        
        st.divider()
        
        # 2. Inspector / Detalle de "Por qué" facturó más o menos un día específico
        st.markdown("#### 2. Inspeccionar Detalle de un Día (¿Por qué se facturó más o menos?)")
        
        dias_del_mes = df_c_graf['fecha'].tolist()
        dia_inspeccionar = st.selectbox("Selecciona un día específico para ver sus movimientos:", dias_del_mes, index=len(dias_del_mes)-1)
        
        df_t_dia = df_trans[df_trans['fecha'] == dia_inspeccionar]
        
        if df_t_dia.empty:
            st.write("No hay detalle de ítems registrado para este día en la base de datos.")
        else:
            col_d1, col_d2 = st.columns([2, 1])
            with col_d1:
                st.markdown(f"**Movimientos registrados el {dia_inspeccionar}:**")
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
        
        # 3. Curva Evolutiva por Meses
        st.markdown("#### 3. Curva Evolutiva por Meses (Tendencia de Crecimiento)")
        
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
            font=dict(color='#1e293b'),
            xaxis=dict(showgrid=True, gridcolor='#f1f5f9'),
            yaxis=dict(showgrid=True, gridcolor='#f1f5f9')
        )
        st.plotly_chart(fig_curve, use_container_width=True)

# -----------------------------------------------------------------------------
# PESTAÑA 5: HISTÓRICO Y EDICIÓN MANUAL
# -----------------------------------------------------------------------------
with tabs[4]:
    st.subheader("⚙️ Histórico Completo de Cierres Diarios")
    
    conn = sqlite3.connect(DB_NAME)
    df_all = pd.read_sql_query("SELECT * FROM cierres_diarios ORDER BY fecha DESC", conn)
    conn.close()
    
    if not df_all.empty:
        st.write("Edita directamente cualquier valor histórico si necesitas hacer alguna corrección manual.")
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
            st.success("✅ Cambios guardados correctamente en la base de datos.")
