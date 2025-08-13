#VERSION OK CON AJUSTE HORA DE COPIA Y EDICION y envio correo y descripcion de stms modificados


import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import os
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from zoneinfo import ZoneInfo
import smtplib
from email.message import EmailMessage

#------ Configuración de vista de la pagina----------
st.set_page_config(layout="wide")
st.title("📋 Masterfile Sutel")
# ================== CONFIGURACIÓN ==================
USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
FILE_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion/MasterfileSutel.xlsx"
FOLDER_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion"
BACKUP_FOLDER_URL = f"{FOLDER_URL}/Backups"

# ================== CONFIG SMTP ==================
SMTP_SERVER = st.secrets["smtp_server"]
SMTP_PORT = st.secrets["smtp_port"]
SMTP_USER = st.secrets["smtp_user"]
SMTP_PASS = st.secrets["smtp_pass"]
EMAIL_FROM = st.secrets["email_from"]
EMAIL_TO = st.secrets["email_to"]  # Puede ser "correo1,correo2" o uno solo

def enviar_correo_con_adjunto(asunto, cuerpo, archivo_bytes, nombre_archivo):
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg.set_content(cuerpo)

    msg.add_attachment(
        archivo_bytes.getvalue(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=nombre_archivo
    )

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg)

try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))

    nombre_archivo = os.path.basename(FILE_URL)

    # Descargar archivo original
    file = ctx.web.get_file_by_server_relative_url(FILE_URL)
    file_stream = BytesIO()
    file.download(file_stream).execute_query()
    file_stream.seek(0)

    # Leer Excel original
    df_original = pd.read_excel(file_stream)
    st.success(f"📂 Cargado {nombre_archivo} ✅")

    # Mostrar tabla editable
    gb = GridOptionsBuilder.from_dataframe(df_original)
    gb.configure_default_column(editable=True, resizable=True, filter=True, sortable=True)
    gb.configure_pagination(enabled=False)
    grid_options = gb.build()

    grid_response = AgGrid(
        df_original,
        gridOptions=grid_options,
        height=500,
        fit_columns_on_grid_load=False,
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        allow_unsafe_jscode=True,
        theme="balham",
        reload_data=False
    )

    df_modificado = pd.DataFrame(grid_response["data"])

    # Guardar cambios
    if st.button("💾 Guardar nueva versión de Masterfile"):
        # Detectar cambios y obtener filas modificadas
        cambios = []
        for i in range(len(df_modificado)):
            if not df_modificado.iloc[i].equals(df_original.iloc[i]):
                cambios.append(str(df_modificado.iloc[i, 1]))  # Columna 2 (índice 1)

        # 👉 Convertir a viñetas para el correo
        if cambios:
            filas_cambiadas = "\n" + "\n".join([f"• {c}" for c in cambios])
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

        # Enviar correo con detalle de fila cambiada (en viñetas)
        try:
            enviar_correo_con_adjunto(
                asunto="Nueva versión del Masterfile Sutel Fijo",
                cuerpo=(
                    f"Buen día,\n\n"
                    f"Se ha guardado una nueva versión del Masterfile: \n\n {nuevo_nombre}\n\n"
                    f"STM actualizados:{filas_cambiadas}\n\n"
                    f"Un saludo"
                ),
                archivo_bytes=output_stream,
                nombre_archivo=nuevo_nombre
            )
            st.success("📧 Correo enviado notificando la nueva versión.")
        except Exception as e:
            st.error(f"Error al enviar correo: {e}")

except Exception as e:
    st.error(f"Error: {e}")


