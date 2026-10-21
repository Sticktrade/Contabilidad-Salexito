import os
import imaplib
import email
from email.header import decode_header
import re
from datetime import datetime, timedelta
import pandas as pd
from supabase import create_client, Client

# Variables de entorno obtenidas desde los Secretos de GitHub
GMAIL_USER = os.environ.get("GMAIL_USER", "").strip()
GMAIL_PASS = os.environ.get("GMAIL_PASS", "").strip()

# Sanitización de la URL de Supabase para evitar errores de formato (ej. PGRST125)
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "").strip()

def connect_gmail():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_PASS)
    return mail

def fetch_corte_attachments():
    mail = connect_gmail()
    mail.select("inbox")

    # Calcular la fecha de hace 15 días en formato IMAP (ej. "06-Oct-2026")
    since_date = (datetime.now() - timedelta(days=15)).strftime("%d-%b-%Y")

    # Buscar correos con asunto "Corte" recibidos en los últimos 15 días (leídos o no leídos)
    search_criterion = f'(SUBJECT "Corte" SINCE "{since_date}")'
    status, messages = mail.search(None, search_criterion)
    email_ids = messages[0].split()

    if not email_ids:
        print(f"No se encontraron correos con el asunto 'Corte' en los últimos 15 días (desde {since_date}).")
        mail.logout()
        return

    print(f"Encontrados {len(email_ids)} correos en los últimos 15 días. Procesando...")

    for email_id in reversed(email_ids):
        status, data = mail.fetch(email_id, "(RFC822)")

        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart" or part.get("Content-Disposition") is None:
                        continue

                    filename = part.get_filename()
                    if filename and (filename.endswith(".xlsx") or filename.endswith(".xls")):
                        print(f"📄 Archivo adjunto encontrado: {filename}")
                        
                        file_data = part.get_payload(decode=True)
                        os.makedirs("downloads", exist_ok=True)
                        filepath = os.path.join("downloads", filename)
                        
                        with open(filepath, "wb") as f:
                            f.write(file_data)
                        
                        process_and_upload(filepath, filename)

    mail.logout()

def process_and_upload(filepath, filename):
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

    fecha_cierre = None

    # 1. Buscar formato AAAA-MM-DD o AAAA-M-D (ej: 2026-4-23 o 2026-04-23)
    match_iso = re.search(r'(\d{4})[-_.](\d{1,2})[-_.](\d{1,2})', filename)
    if match_iso:
        ano, mes, dia = match_iso.groups()
        fecha_cierre = f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"
    else:
        # 2. Buscar formato DD-MM-AAAA o D-M-AAAA (ej: 23-4-2026 o 5-4-2026)
        match_dmy = re.search(r'(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})', filename)
        if match_dmy:
            dia, mes, ano = match_dmy.groups()
            fecha_cierre = f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"

    if not fecha_cierre:
        print(f"⚠️ No se identificó fecha en el nombre del archivo: {filename}")
        return

    # 3. Consultar si la fecha ya existe en la tabla 'cierres_diarios'
    response = supabase.table("cierres_diarios").select("fecha").eq("fecha", fecha_cierre).execute()

    if response.data:
        print(f"ℹ️ El cierre del {fecha_cierre} ya está registrado. Se omite.")
    else:
        print(f"🚀 Registrando cierre del {fecha_cierre} en Supabase...")

        df = pd.read_excel(filepath)

        # Totales por defecto
        efectivo_val = 0.0
        nequi_val = 0.0
        daviplata_val = 0.0
        banco_val = 0.0
        total_val = efectivo_val + nequi_val + daviplata_val + banco_val

        supabase.table("cierres_diarios").insert({
            "fecha": fecha_cierre,
            "efectivo": efectivo_val,
            "nequi": nequi_val,
            "daviplata": daviplata_val,
            "banco": banco_val,
            "total": total_val,
            "observaciones": f"Cargado automáticamente desde {filename}"
        }).execute()

        print(f"✅ Registrado con éxito el día {fecha_cierre}.")

if __name__ == "__main__":
    fetch_corte_attachments()
