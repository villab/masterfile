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
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#------ Configuración de vista de la pagina----------
st.set_page_config(layout="wide") 

# ================== CONFIGURACIÓN ==================
USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SMTP_SERVER = st.secrets["smtp_server"]
SMTP_PORT = st.secrets["smtp_port"]
SMTP_USER = st.secrets["smtp_user"]
SMTP_PASS = st.secrets["smtp_pass"]
EMAIL_FROM = st.secrets["email_from"]
EMAIL_TO = st.secrets["email_to"].split(",")

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
FILE_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion/MasterfileSutel.xlsx"
FOLDER_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion"
BACKUP_FOLDER_URL = f"{FOLDER_URL}/Backups"

try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))

    # Obtener solo el nombre del archivo
    nombre_archivo = os.path.basename(FILE_URL)

    # Descargar archivo original
    file = ctx.web.get_file_by_server_relative_url(FILE_URL)
    file_stream = BytesIO()
    file.download(file_stream).execute_query()
    file_stream.seek(0)

    # ================== LECTURA DEL EXCEL ==================
    df_original = pd.read_excel(file_stream)
    df = df_original.copy()
    st.success(f"📂 Cargado  {nombre_archivo} ✅") 

    # ================== Mostrar tabla editable ==================
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        editable=True,
        resizable=True,
        filter=True,
        sortable=True
    )
    gb.configure_pagination(enabled=False)
    grid_options = gb.build()

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        height=500,
        fit_columns_on_grid_load=False,
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        allow_unsafe_jscode=True,
        theme="balham",
        reload_data=False
    )

    df_editado = pd.DataFrame(grid_response["data"])

    # ================== GUARDAR CAMBIOS ==================
    if st.button("💾 Guardar nueva versión de Masterfile"):
        cambios = []

        # Detectar cambios en celdas
        for i in range(min(len(df_original), len(df_editado))):
            for col in df_original.columns:
                if df_original.at[i, col] != df_editado.at[i, col]:
                    celda_identificadora = df_original.iloc[i, 1]  # Columna 2
                    cambios.append(
                        f"<li><b>{celda_identificadora}</b> → Columna '<i>{col}</i>' cambiado de '<b>{df_original.at[i, col]}</b>' a '<b>{df_editado.at[i, col]}</b>'</li>"
                    )

        # Detectar filas nuevas
        if len(df_editado) > len(df_original):
            nuevas_filas = df_editado.iloc[len(df_original):]
            for _, fila in nuevas_filas.iterrows():
                cambios.append(f"<li><b>Nueva fila añadida:</b> {fila.to_dict()}</li>")

        # Guardar si hubo cambios o nuevas filas
        if cambios:
            timestamp = datetime.now(ZoneInfo("America/Costa_Rica")).strftime("%Y%m%d_%H%M%S")
            nuevo_nombre = f"MasterfileSutel_{timestamp}.xlsx"

            # Guardar DataFrame en memoria
            output_stream = BytesIO()
            df_editado.to_excel(output_stream, index=False)
            output_stream.seek(0)

            # Verificar o crear carpeta Backups
            try:
                ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL).expand(["Files"]).get().execute_query()
            except:
                ctx.web.folders.add(BACKUP_FOLDER_URL).execute_query()

            # Subir copia con fecha a Backups
            backup_folder = ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL)
            backup_folder.upload_file(nuevo_nombre, output_stream).execute_query()

            # Sobrescribir el archivo original
            output_stream.seek(0)
            main_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
            main_folder.upload_file("MasterfileSutel.xlsx", output_stream).execute_query()

            st.success(f"✅ Cambios guardados y copia creada en 'Backups' como {nuevo_nombre}")

            # ================== ENVIAR CORREO ==================
            cuerpo_html = f"""
            <html>
            <body>
                <p>Se han realizado los siguientes cambios en el Masterfile:</p>
                <ul>
                    {''.join(cambios)}
                </ul>
            </body>
            </html>
            """
            msg = MIMEMultipart()
            msg["From"] = EMAIL_FROM
            msg["To"] = ", ".join(EMAIL_TO)
            msg["Subject"] = "Cambios en Masterfile Sutel"
            msg.attach(MIMEText(cuerpo_html, "html"))

            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())

        else:
            st.info("No se detectaron cambios en el archivo.")

except Exception as e:
    st.error(f"Error: {e}")
