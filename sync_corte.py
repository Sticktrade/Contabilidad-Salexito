import os
import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
from supabase import create_client, Client

# Variables de entorno obtenidas desde los Secretos de GitHub
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASS = os.environ.get("GMAIL_PASS")
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

def connect_gmail():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_PASS)
    return mail

def fetch_corte_attachments():
    mail = connect_gmail()
    mail.select("inbox")

    # Buscar correos con el asunto "Corte"
    status, messages = mail.search(None, '(SUBJECT "Corte")')
    email_ids = messages[0].split()

    if not email_ids:
        print("No se encontraron correos con el asunto 'Corte'.")
        return

    # Procesar los correos encontrados (empezando por el más reciente)
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
                        
                        # Procesar e intentar subir a Supabase
                        process_and_upload(filepath, filename)

    mail.logout()

def process_and_upload(filepath, filename):
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

    # 1. Extraer la fecha del nombre del archivo (formato AAAA-MM-DD o DD-MM-AAAA)
    match = re.search(r'\d{4}-\d{2}-\d{2}', filename)
    if match:
        fecha_cierre = match.group(0)
    else:
        match_alt = re.search(r'(\d{2})-(\d{2})-(\d{4})', filename)
        if match_alt:
            dia, mes, ano = match_alt.groups()
            fecha_cierre = f"{ano}-{mes}-{dia}"
        else:
            print(f"⚠️ No se identificó fecha en el nombre del archivo: {filename}")
            return

    # 2. Verificar si la fecha ya existe en 'cierres_diarios'
    response = supabase.table("cierres_diarios").select("fecha").eq("fecha", fecha_cierre).execute()

    if response.data:
        print(f"ℹ️ El cierre del {fecha_cierre} ya está en la base de datos. Se omite.")
    else:
        print(f"🚀 Registrando cierre del {fecha_cierre} en Supabase...")

        # 3. Leer el Excel descargado
        # NOTA: Si los valores de efectivo, nequi, etc., están en celdas específicas, 
        # puedes adaptar la lectura con pandas.
        df = pd.read_excel(filepath)

        efectivo_val = 0.0
        nequi_val = 0.0
        daviplata_val = 0.0
        banco_val = 0.0
        total_val = efectivo_val + nequi_val + daviplata_val + banco_val

        # 4. Insertar fila en tu tabla cierres_diarios
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
