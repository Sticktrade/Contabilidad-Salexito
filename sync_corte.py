import os
import imaplib
import email
from email.header import decode_header
import re
import openpyxl
from supabase import create_client, Client

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE BÚSQUEDA Y CREDENCIALES
# -----------------------------------------------------------------------------
GMAIL_USER = os.environ.get("GMAIL_USER", "").strip()
GMAIL_PASS = os.environ.get("GMAIL_PASS", "").strip()
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "").strip()

# Agrega aquí cualquier palabra clave que pueda venir en el asunto (no distingue mayúsculas)
ASUNTOS_BUSQUEDA = ["Corte", "CUENTAS"]
DIAS_ATRAS = 15

# Opcional: si quieres filtrar también por quien envía el correo, ponlo aquí (ej. ["tienda@gmail.com"])
# Déjalo vacío [] para aceptar correos de cualquier remitente que coincida con el asunto
REMITENTES_PERMITIDOS = []

# -----------------------------------------------------------------------------
# UTILIDADES DE LIMPIEZA Y DECODIFICACIÓN
# -----------------------------------------------------------------------------
def decode_mime_words(s):
    if not s:
        return ""
    decoded_parts = decode_header(s)
    result = []
    for content, charset in decoded_parts:
        if isinstance(content, bytes):
            result.append(content.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(content)
    return "".join(result)

def parse_float(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    cleaned = str(val).strip().replace("$", "").replace(",", "")
    if not cleaned or cleaned == "-":
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

def parse_code(val):
    if val is None:
        return 1
    try:
        return int(float(str(val).strip()))
    except ValueError:
        return 1

# -----------------------------------------------------------------------------
# CONEXIÓN Y BÚSQUEDA EN GMAIL
# -----------------------------------------------------------------------------
def connect_gmail():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_PASS)
    return mail

def fetch_corte_attachments():
    mail = connect_gmail()
    # "[Gmail]/All Mail" busca en recibidos y archivados; si prefieres solo recibidos, usa "inbox"
    mail.select('"[Gmail]/All Mail"')

    # Construir consulta nativa de Gmail: subject:(Corte OR CUENTAS) has:attachment newer_than:15d
    asuntos_query = " OR ".join(ASUNTOS_BUSQUEDA)
    gmail_query = f"subject:({asuntos_query}) has:attachment newer_than:{DIAS_ATRAS}d"

    if REMITENTES_PERMITIDOS:
        from_query = " OR ".join(REMITENTES_PERMITIDOS)
        gmail_query += f" from:({from_query})"

    status, messages = mail.search(None, "X-GM-RAW", f'"{gmail_query}"')
    email_ids = messages[0].split()

    if not email_ids:
        print(f"No se encontraron correos para la búsqueda: {gmail_query}")
        mail.logout()
        return

    print(f"Encontrados {len(email_ids)} correos en los últimos {DIAS_ATRAS} días. Procesando...")

    for email_id in reversed(email_ids):
        status, data = mail.fetch(email_id, "(RFC822)")

        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject = decode_mime_words(msg.get("Subject", "Sin asunto"))

                for part in msg.walk():
                    if part.get_content_maintype() == "multipart" or part.get("Content-Disposition") is None:
                        continue

                    raw_filename = part.get_filename()
                    filename = decode_mime_words(raw_filename).strip()

                    if filename and filename.lower().endswith((".xlsx", ".xls")):
                        print(f"📄 Adjunto encontrado en '{subject}': {filename}")
                        file_data = part.get_payload(decode=True)
                        os.makedirs("downloads", exist_ok=True)
                        filepath = os.path.join("downloads", filename)

                        with open(filepath, "wb") as f:
                            f.write(file_data)

                        try:
                            process_and_upload(filepath, filename)
                        except Exception as e:
                            print(f"❌ Error procesando '{filename}': {e}")

    mail.logout()

# -----------------------------------------------------------------------------
# PROCESAMIENTO DEL EXCEL E INSERCIÓN EN SUPABASE
# -----------------------------------------------------------------------------
def process_and_upload(filepath, filename):
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    fecha_cierre = None

    match_iso = re.search(r'(\d{4})[-_.](\d{1,2})[-_.](\d{1,2})', filename)
    if match_iso:
        ano, mes, dia = match_iso.groups()
        fecha_cierre = f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"
    else:
        match_dmy = re.search(r'(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})', filename)
        if match_dmy:
            dia, mes, ano = match_dmy.groups()
            fecha_cierre = f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"

    if not fecha_cierre:
        print(f"⚠️ No se identificó fecha en el nombre del archivo: {filename}")
        return

    response = supabase.table("cierres_diarios").select("fecha").eq("fecha", fecha_cierre).execute()
    if response.data:
        print(f"ℹ️ El cierre del {fecha_cierre} ya está registrado. Se omite.")
        return

    print(f"🚀 Leyendo Excel para registrar el cierre del {fecha_cierre}...")

    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb.active
    data = [[sheet.cell(row=r, column=c).value for c in range(1, sheet.max_column + 1)] for r in range(1, sheet.max_row + 1)]

    header_idx = next(
        (idx for idx, row in enumerate(data)
         if 'FECHA' in [str(x).strip().upper() if x else '' for x in row]
         and 'CONCEPTO' in [str(x).strip().upper() if x else '' for x in row]),
        None
    )

    if header_idx is None:
        print(f"❌ No se encontró la estructura estándar en el Excel: {filename}")
        return

    header = [str(x).strip().upper() if x else '' for x in data[header_idx]]
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

    transactions = []
    for row in data[header_idx + 1:]:
        if row[col_concepto] is None and row[col_ingreso] is None and row[col_gastos] is None:
            continue

        code = row[col_codigo]
        if code is None or str(code).strip() == '':
            continue

        concepto = str(row[col_concepto]).strip() if row[col_concepto] else ''
        try:
            cant = int(float(str(row[col_cantidad]).strip())) if row[col_cantidad] else 1
        except ValueError:
            cant = 1

        ingreso = parse_float(row[col_ingreso])
        gastos = parse_float(row[col_gastos])
        saldo = parse_float(row[col_saldo])

        transactions.append({
            'concepto': concepto,
            'cantidad': cant,
            'codigo': parse_code(code),
            'ingreso': ingreso,
            'gastos': gastos,
            'saldo': saldo
        })

    saldo_efectivo = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 1)
    saldo_daviplata = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 2)
    saldo_nequi = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 3)
    saldo_banco = sum(t['ingreso'] - t['gastos'] for t in transactions if t['codigo'] == 4)
    total_caja = saldo_efectivo + saldo_daviplata + saldo_nequi + saldo_banco

    supabase.table("cierres_diarios").insert({
        "fecha": fecha_cierre,
        "efectivo": saldo_efectivo,
        "nequi": saldo_nequi,
        "daviplata": saldo_daviplata,
        "banco": saldo_banco,
        "total": total_caja,
        "observaciones": f"Cargado automáticamente desde {filename}"
    }).execute()

    transacciones_db = [
        {
            "fecha": fecha_cierre,
            "concepto": t['concepto'],
            "cantidad": t['cantidad'],
            "codigo": t['codigo'],
            "ingreso": t['ingreso'],
            "gastos": t['gastos'],
            "saldo": t['saldo']
        }
        for t in transactions
    ]

    if transacciones_db:
        supabase.table("transacciones_diarias").insert(transacciones_db).execute()

    print(f"✅ Registrado con éxito el día {fecha_cierre} ({len(transactions)} transacciones).")

if __name__ == "__main__":
    fetch_corte_attachments()
