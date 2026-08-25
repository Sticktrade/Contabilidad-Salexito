import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import re
from datetime import datetime, date
import plotly.express as px

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Control de Caja - Papelería",
    page_icon="📚",
    layout="wide"
)

DB_NAME = "papeleria.db"

def init_db():
    """Inicializa la base de datos SQLite con las tablas necesarias."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    # Tabla para el consolidado diario de la caja mensual
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
    # Tabla para el detalle de transacciones de cada día
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
    # Tabla para gastos grandes, mensuales o compras de mercancía
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
# FUNCIONES DE PARSEO E INTELIGENCIA DE DATOS
# -----------------------------------------------------------------------------
def clean_date(val, filename_date=None):
    """Limpia y convalida la fecha del archivo o usa la del nombre si hay errores de tipeo."""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
            
    # Intentar corregir con expresión regular errores como "29/0772026"
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
    """Extrae automáticamente los ingresos, gastos y tipos de caja del archivo del local."""
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
    
    # Extraer fecha desde el nombre del archivo si viene corrupta dentro de las celdas
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
        
    # Calcular saldos netos diarios por método de pago
    # Codigo 1: Efectivo | Codigo 2: Daviplata | Codigo 3: Nequi | Codigo 4: Banco
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
# INTERFAZ WEB
# -----------------------------------------------------------------------------
st.title("📚 Sistema de Gestión de Caja y Contabilidad - Papelería")
st.caption("Administración remota eficiente, automatizada e intuitiva.")

tabs = st.tabs([
    "📥 Cargar Caja Diaria", 
    "📊 Recuento Mensual", 
    "💸 Gastos Mensuales / Compras", 
    "📈 Métricas y Visuales",
    "⚙️ Histórico y Edición Manual"
])

# -----------------------------------------------------------------------------
# PESTAÑA 1: CARGAR CAJA DIARIA
# -----------------------------------------------------------------------------
with tabs[0]:
    st.header("Cargar reporte diario enviado desde el local")
    uploaded_file = st.file_uploader("Adjunta el archivo Excel del día (ej. 28-07-2026.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        summary, df_trans = parse_daily_excel(uploaded_file)
        
        if isinstance(df_trans, str):
            st.error(df_trans)
        else:
            st.success("✅ Archivo procesado con éxito. Revisa la información antes de guardar:")
            
            col_f1, col_f2 = st.columns([1, 2])
            with col_f1:
                fecha_final = st.date_input(
                    "Fecha asignada al cierre:", 
                    value=datetime.strptime(summary['fecha'], "%Y-%m-%d").date()
                )
            fecha_str = fecha_final.strftime("%Y-%m-%d")
            
            st.subheader("Resumen de Saldos Calculados:")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("💵 Efectivo", f"${summary['efectivo']:,.0f}")
            c2.metric("🟣 Nequi", f"${summary['nequi']:,.0f}")
            c3.metric("🔴 Daviplata", f"${summary['daviplata']:,.0f}")
            c4.metric("🏦 Banco", f"${summary['banco']:,.0f}")
            c5.metric("💰 Total Caja", f"${summary['total']:,.0f}")
            
            st.subheader("Detalle de Transacciones (Edita si requieres corregir alguna cifra):")
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
                st.success(f"🎉 ¡Cierre del día {fecha_str} registrado exitosamente en la caja mensual!")

# -----------------------------------------------------------------------------
# PESTAÑA 2: RECUENTO MENSUAL (CONSOLIDADO)
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("📊 Recuento de Caja Mensual")
    
    conn = sqlite3.connect(DB_NAME)
    df_cierres = pd.read_sql_query("SELECT * FROM cierres_diarios ORDER BY fecha ASC", conn)
    df_gastos = pd.read_sql_query("SELECT * FROM gastos_mensuales ORDER BY fecha ASC", conn)
    conn.close()
    
    if df_cierres.empty:
        st.info("Aún no hay cierres diarios registrados. Carga tu primer archivo en la pestaña 'Cargar Caja Diaria'.")
    else:
        df_cierres['mes_año'] = pd.to_datetime(df_cierres['fecha']).dt.strftime('%Y-%m')
        meses_disponibles = df_cierres['mes_año'].unique()
        
        mes_sel = st.selectbox("Selecciona el Mes a consultar:", meses_disponibles, index=len(meses_disponibles)-1)
        
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
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("💰 Total Ingresos de Caja", f"${tot_caja:,.0f}")
        kpi2.metric("💸 Total Gastos Grandes/Mensuales", f"${tot_gastos_grandes:,.0f}")
        kpi3.metric("📊 Liquidez de Caja Neto", f"${liquidez_neta:,.0f}", delta=f"${liquidez_neta:,.0f}")
        kpi4.metric("💵 Efectivo Acumulado", f"${tot_efectivo:,.0f}")
        
        st.divider()
        col_t1, col_t2 = st.columns([3, 2])
        
        with col_t1:
            st.subheader("Caja Diaria Consolidada")
            st.dataframe(
                df_c_mes[['fecha', 'efectivo', 'nequi', 'daviplata', 'banco', 'total']],
                column_config={
                    "fecha": "Fecha",
                    "efectivo": st.column_config.NumberColumn("Efectivo", format="$%d"),
                    "nequi": st.column_config.NumberColumn("Nequi", format="$%d"),
                    "daviplata": st.column_config.NumberColumn("Daviplata", format="$%d"),
                    "banco": st.column_config.NumberColumn("Banco", format="$%d"),
                    "total": st.column_config.NumberColumn("Total Día", format="$%d"),
                },
                use_container_width=True,
                hide_index=True
            )
            
        with col_t2:
            st.subheader("Gastos Mensuales y Mercancía")
            if df_g_mes.empty:
                st.write("No hay gastos grandes registrados en este mes.")
            else:
                st.dataframe(
                    df_g_mes[['fecha', 'concepto', 'categoria', 'monto']],
                    column_config={
                        "fecha": "Fecha",
                        "concepto": "Concepto",
                        "categoria": "Categoría",
                        "monto": st.column_config.NumberColumn("Monto ($)", format="$%d"),
                    },
                    use_container_width=True,
                    hide_index=True
                )

# -----------------------------------------------------------------------------
# PESTAÑA 3: GASTOS MENSUALES / COMPRAS DE MERCANCÍA
# -----------------------------------------------------------------------------
with tabs[2]:
    st.header("💸 Registrar Gasto Mensual, Compra o Factura")
    st.write("Ingresa los gastos mayores (Arriendo, Servicios, Proveedores) que no deben contaminar la caja diaria.")
    
    with st.form("form_gastos", clear_on_submit=True):
        col_g1, col_g2, col_g3 = st.columns(3)
        with col_g1:
            fecha_gasto = st.date_input("Fecha del Gasto:", value=date.today())
        with col_g2:
            categoria_gasto = st.selectbox("Categoría:", [
                "Arriendo", "Servicios (Agua/Luz/Gas)", "Internet/Teléfono", 
                "Mercancía/Proveedores", "Nómina/Pagos", "Plataformas/Recargas", "Otro"
            ])
        with col_g3:
            monto_gasto = st.number_input("Monto en Pesos ($):", min_value=0.0, step=1000.0)
            
        concepto_gasto = st.text_input("Concepto o Descripción (ej. RESMAS Y TINTA EPSON, ARRIENDO LOCAL):")
        
        submitted = st.form_submit_button("➕ Registrar Gasto", type="primary")
        
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
                    monto_gasto, 
                    categoria_gasto, 
                    ""
                ))
                conn.commit()
                conn.close()
                st.success(f"✅ Gasto '{concepto_gasto}' por ${monto_gasto:,.0f} registrado con éxito.")
                st.rerun()
            else:
                st.warning("Ingresa un concepto y un monto válido.")

# -----------------------------------------------------------------------------
# PESTAÑA 4: MÉTRICAS Y VISUALES (DASHBOARD)
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("📈 Análisis Visual de Métricas")
    
    conn = sqlite3.connect(DB_NAME)
    df_cierres = pd.read_sql_query("SELECT * FROM cierres_diarios ORDER BY fecha ASC", conn)
    df_gastos = pd.read_sql_query("SELECT * FROM gastos_mensuales ORDER BY fecha ASC", conn)
    conn.close()
    
    if df_cierres.empty:
        st.info("No hay suficiente información para generar gráficos. Registra algunos días primero.")
    else:
        st.subheader("Evolución del Ingreso Diario por Método de Pago")
        fig_line = px.line(
            df_cierres, 
            x="fecha", 
            y=["efectivo", "nequi", "daviplata", "total"],
            labels={"value": "Pesos ($)", "fecha": "Fecha", "variable": "Canal"},
            title="Evolución de Flujo de Caja",
            markers=True
        )
        st.plotly_chart(fig_line, use_container_width=True)
        
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.subheader("Distribución de Ingresos por Canal")
            tot_ef = df_cierres['efectivo'].sum()
            tot_neq = df_cierres['nequi'].sum()
            tot_dav = df_cierres['daviplata'].sum()
            tot_ban = df_cierres['banco'].sum()
            
            df_pie = pd.DataFrame({
                'Canal': ['Efectivo', 'Nequi', 'Daviplata', 'Banco'],
                'Monto': [tot_ef, tot_neq, tot_dav, tot_ban]
            })
            fig_pie = px.pie(df_pie, values='Monto', names='Canal', hole=0.4, title="Participación en Caja")
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_g2:
            if not df_gastos.empty:
                st.subheader("Gastos Mayores por Categoría")
                fig_bar = px.bar(
                    df_gastos, 
                    x="categoria", 
                    y="monto", 
                    color="categoria",
                    title="Gastos Acumulados por Tipo",
                    labels={"monto": "Monto ($)", "categoria": "Categoría"}
                )
                st.plotly_chart(fig_bar, use_container_width=True)

# -----------------------------------------------------------------------------
# PESTAÑA 5: HISTÓRICO Y AJUSTES MANUALES
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("⚙️ Histórico Completo y Corrección Manual")
    st.write("Modifica o elimina registros si necesitas corregir datos históricos.")
    
    conn = sqlite3.connect(DB_NAME)
    df_all_cierres = pd.read_sql_query("SELECT * FROM cierres_diarios ORDER BY fecha DESC", conn)
    conn.close()
    
    if not df_all_cierres.empty:
        st.subheader("Cierres Diarios Registrados")
        edited_historico = st.data_editor(
            df_all_cierres,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_historico"
        )
        
        if st.button("💾 Actualizar Cambios en Histórico"):
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("DELETE FROM cierres_diarios")
            for _, r in edited_historico.iterrows():
                c.execute("""
                    INSERT INTO cierres_diarios (fecha, efectivo, nequi, daviplata, banco, total, observaciones)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (r['fecha'], r['efectivo'], r['nequi'], r['daviplata'], r['banco'], r['total'], r['observaciones']))
            conn.commit()
            conn.close()
            st.success("Histórico actualizado correctamente.")
