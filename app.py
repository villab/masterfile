import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# --- Cargar credenciales desde secrets ---
sharepoint_user = st.secrets["sharepoint_user"]
app_password = st.secrets["app_password"]

smtp_server = st.secrets["smtp_server"]
smtp_port = st.secrets["smtp_port"]
smtp_user = st.secrets["smtp_user"]
smtp_pass = st.secrets["smtp_pass"]
email_from = st.secrets["email_from"]
email_to = st.secrets["email_to"]

# --- Configuración de SharePoint ---
SITE_URL = "https://tusitio.sharepoint.com/sites/TuSitio"
FOLDER_URL = "/sites/TuSitio/Documentos compartidos"
BACKUP_FOLDER_URL = "/sites/TuSitio/Documentos compartidos/Backups"

# --- Función para enviar correo ---
def enviar_correo_con_adjunto(asunto, cuerpo, archivo_bytes, nombre_archivo):
    mensaje = MIMEMultipart()
    mensaje["From"] = email_from
    mensaje["To"] = email_to
    mensaje["Subject"] = asunto

    mensaje.attach(MIMEText(cuerpo, "plain"))

    parte = MIMEBase("application", "octet-stream")
    parte.set_payload(archivo_bytes.read())
    encoders.encode_base64(parte)
    parte.add_header("Content-Disposition", f"attachment; filename={nombre_archivo}")
    mensaje.attach(parte)

    with smtplib.SMTP(smtp_server, smtp_port) as servidor:
        servidor.starttls()
        servidor.login(smtp_user, smtp_pass)
        servidor.sendmail(email_from, email_to.split(","), mensaje.as_string())

# --- Conexión a SharePoint ---
ctx = ClientContext(SITE_URL).with_credentials(UserCredential(sharepoint_user, app_password))

# --- Descargar archivo original ---
file_url = f"{FOLDER_URL}/MasterfileSutel.xlsx"
response = ctx.web.get_file_by_server_relative_url(file_url).download().execute_query()
df_original = pd.read_excel(BytesIO(response.content))

# --- Mostrar en Streamlit ---
st.title("Editor de Masterfile")
df_modificado = st.data_editor(df_original, num_rows="dynamic")

if st.button("Guardar cambios"):
    cambios = []
    for i in range(len(df_modificado)):
        if not df_modificado.iloc[i].equals(df_original.iloc[i]):
            cambios.append(str(df_modificado.iloc[i, 1]))  # Columna 2 (índice 1)

    if cambios:
        filas_cambiadas = "\n".join([f"• {c}" for c in cambios])
    else:
        filas_cambiadas = "Ninguna fila detectada"

    timestamp = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%Y%m%d_%H%M%S")
    nuevo_nombre = f"MasterfileSutel_{timestamp}.xlsx"

    output_stream = BytesIO()
    df_modificado.to_excel(output_stream, index=False)
    output_stream.seek(0)

    try:
        ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL).expand(["Files"]).get().execute_query()
    except:
        ctx.web.folders.add(BACKUP_FOLDER_URL).execute_query()

    backup_folder = ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL)
    backup_folder.upload_file(nuevo_nombre, output_stream).execute_query()

    output_stream.seek(0)

    main_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
    main_folder.upload_file("MasterfileSutel.xlsx", output_stream).execute_query()

    st.success(f"✅ Cambios guardados y copia creada en 'Backups' como {nuevo_nombre}")

    # --- Enviar correo ---
    try:
        cuerpo = (
            f"Se ha guardado una nueva versión del Masterfile: {nuevo_nombre}\n\n"
            f"Filas modificadas (columna 2):\n{filas_cambiadas}"
        )

        enviar_correo_con_adjunto(
            asunto="Nueva versión del Masterfile guardada",
            cuerpo=cuerpo,
            archivo_bytes=output_stream,
            nombre_archivo=nuevo_nombre
        )
        st.success("📧 Correo enviado notificando la nueva versión.")
    except Exception as e:
        st.error(f"Error al enviar correo: {e}")
