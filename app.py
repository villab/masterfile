# VERSION CON SUBCARPETAS DE BACKUP SEPARADAS PARA FIJO Y MOVILIDAD

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

# ------ Configuración de vista de la pagina ----------
st.set_page_config(layout="wide")
st.title("📋 Masterfile Entorno de medición Fijo y Movilidad")

# ================== CONFIGURACIÓN ==================
USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
FOLDER_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion"

ARCHIVOS = {
    "Fijo": "MasterfileSutel.xlsx",
    "Movilidad": "MasterfileSutel_Movilidad.xlsx"
}

# ================== CONFIG SMTP ==================
SMTP_SERVER = st.secrets["smtp_server"]
SMTP_PORT = st.secrets["smtp_port"]
SMTP_USER = st.secrets["smtp_user"]
SMTP_PASS = st.secrets["smtp_pass"]
EMAIL_FROM = st.secrets["email_from"]
EMAIL_TO = st.secrets["email_to"]


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


def manejar_archivo(nombre_modo, nombre_archivo):
    """Carga, muestra, permite editar y guardar un archivo específico"""
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))

    FILE_URL = f"{FOLDER_URL}/{nombre_archivo}"

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
    if st.button(f"💾 Guardar nueva versión ({nombre_modo})"):
        # Detectar cambios y obtener filas modificadas
        cambios = []
        for i in range(len(df_modificado)):
            if not df_modificado.iloc[i].equals(df_original.iloc[i]):
                cambios.append(str(df_modificado.iloc[i, 1]))  # Columna 2 (índice 1)

        if cambios:
            filas_cambiadas = "\n" + "\n".join([f"• {c}" for c in cambios])
        else:
            filas_cambiadas = "Ninguna fila detectada"

        timestamp = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%Y%m%d_%H%M%S")
        nuevo_nombre = f"{nombre_archivo.replace('.xlsx', '')}_{timestamp}.xlsx"

        output_stream = BytesIO()
        df_modificado.to_excel(output_stream, index=False)
        output_stream.seek(0)

        # ====== Crear carpeta de backup para este tipo ======
        backup_folder_url = f"{FOLDER_URL}/Backups/{nombre_modo}"
        try:
            ctx.web.get_folder_by_server_relative_url(backup_folder_url).expand(["Files"]).get().execute_query()
        except:
            ctx.web.folders.add(backup_folder_url).execute_query()

        # Subir copia a Backups/{Fijo|Movilidad}
        backup_folder = ctx.web.get_folder_by_server_relative_url(backup_folder_url)
        backup_folder.upload_file(nuevo_nombre, output_stream).execute_query()

        # Subir archivo principal actualizado
        output_stream.seek(0)
        main_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
        main_folder.upload_file(nombre_archivo, output_stream).execute_query()

        st.success(f"✅ Cambios guardados y copia creada en 'Backups/{nombre_modo}' como {nuevo_nombre}")

        # Enviar correo
        try:
            enviar_correo_con_adjunto(
                asunto=f"Nueva versión del Masterfile Sutel {nombre_modo}",
                cuerpo=(
                    f"Buen día,\n\n"
                    f"Se ha guardado una nueva versión del Masterfile {nombre_modo}: \n\n {nuevo_nombre}\n\n"
                    f"STM actualizados:{filas_cambiadas}\n\n"
                    f"Un saludo"
                ),
                archivo_bytes=output_stream,
                nombre_archivo=nuevo_nombre
            )
            st.success("📧 Correo enviado notificando la nueva versión.")
        except Exception as e:
            st.error(f"Error al enviar correo: {e}")


# ================== INTERFAZ CON PESTAÑAS ==================
try:
    tab_fijo, tab_movilidad = st.tabs(["📄 Masterfile Fijo", "📄 Masterfile Movilidad"])

    with tab_fijo:
        manejar_archivo("Fijo", ARCHIVOS["Fijo"])

    with tab_movilidad:
        manejar_archivo("Movilidad", ARCHIVOS["Movilidad"])

except Exception as e:
    st.error(f"Error: {e}")
