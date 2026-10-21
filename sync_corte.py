import os
import imaplib
import email
from email.header import decode_header
import re

GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASS = os.environ.get("GMAIL_PASS")

def connect_gmail():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_PASS)
    return mail

def fetch_corte_attachments():
    mail = connect_gmail()
    mail.select("inbox")

    # Buscar correos no leídos o recientes con asunto "Corte"
    status, messages = mail.search(None, '(SUBJECT "Corte")')
    email_ids = messages[0].split()

    if not email_ids:
        print("No se encontraron correos nuevos con el asunto 'Corte'.")
        return

    # Tomar el último correo recibido
    latest_email_id = email_ids[-1]
    status, data = mail.fetch(latest_email_id, "(RFC822)")

    for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8")
            
            print(f"Procesando correo: {subject}")

            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue

                filename = part.get_filename()
                if filename and (filename.endswith(".xlsx") or filename.endswith(".xls")):
                    # Extraer fecha del nombre del archivo (ejemplo: Cierre_2026-10-20.xlsx)
                    print(f"Archivo adjunto encontrado: {filename}")
                    
                    file_data = part.get_payload(decode=True)
                    
                    # Guardar localmente para procesar o subir
                    os.makedirs("downloads", exist_ok=True)
                    filepath = os.path.join("downloads", filename)
                    with open(filepath, "wb") as f:
                        f.write(file_data)
                    
                    print(f"Archivo descargado correctamente en: {filepath}")
                    # Aquí se ejecuta la integración con la base de datos o API de tu app
                    process_and_upload(filepath, filename)

    mail.logout()

def process_and_upload(filepath, filename):
    # Lógica de conexión a tu base de datos (Supabase / Google Drive / API Streamlit)
    # 1. Comprobar si la fecha del archivo ya existe en los registros.
    # 2. Si no existe, insertar o subir el archivo.
    print(f"Verificando y sincronizando {filename} con la plataforma...")

if __name__ == "__main__":
    fetch_corte_attachments()
