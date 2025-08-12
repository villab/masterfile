import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import os

# ================== CONFIGURACIÓN DE PÁGINA ==================
st.set_page_config(layout="wide")  # Hace que la app use todo el ancho

# ================== CONFIGURACIÓN ==================
USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

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
    df = pd.read_excel(file_stream)
    st.success(f"📂 Cargado masterfile del día: {nombre_archivo} ✅") 

    # Mostrar y permitir edición (más grande)
    edited_df = st.data_editor(
        df, 
        num_rows="dynamic", 
        height=600,  # Más alto
        use_container_width=True  # Ocupar todo el ancho
    )

    # ================== GUARDAR CAMBIOS ==================
    if st.button("💾 Guardar nueva versión de Masterfile"):
        # Nombre con fecha y hora (YYYYMMDD_HHMMSS)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nuevo_nombre = f"MasterfileSutel_{timestamp}.xlsx"

        # Guardar DataFrame en memoria
        output_stream = BytesIO()
        edited_df.to_excel(output_stream, index=False)
        output_stream.seek(0)

        # Verificar o crear carpeta Backups
        try:
            ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL).expand(["Files"]).get().execute_query()
        except:
            ctx.web.folders.add(BACKUP_FOLDER_URL).execute_query()

        # Subir copia con fecha a Backups
        backup_folder = ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL)
        backup_folder.upload_file(nuevo_nombre, output_stream).execute_query()

        # Volver a poner el puntero al inicio
        output_stream.seek(0)

        # Sobrescribir el archivo original
        main_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
        main_folder.upload_file("MasterfileSutel.xlsx", output_stream).execute_query()

        st.success(f"✅ Cambios guardados y copia creada en 'Backups' como {nuevo_nombre}")

except Exception as e:
    st.error(f"Error: {e}")
