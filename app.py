import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ================== CONFIGURACIÓN ==================
USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
FILE_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion/MasterfileSutel.xlsx"

# ================== CONEXIÓN A SHAREPOINT ==================
try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))
    
    # Descargar archivo
    file = ctx.web.get_file_by_server_relative_url(FILE_URL)
    file_stream = BytesIO()
    file.download(file_stream).execute_query()
    file_stream.seek(0)

    # ================== LECTURA DEL EXCEL ==================
    df = pd.read_excel(file_stream)
    st.success("Archivo cargado correctamente desde SharePoint ✅")

    # Mostrar y permitir edición
    edited_df = st.data_editor(df, num_rows="dynamic")

    # ================== GUARDAR CAMBIOS ==================
    if st.button("Guardar copia con fecha"):
        # Nombre con fecha
        fecha_hoy = datetime.now().strftime("%Y%m%d")
        nuevo_nombre = f"MasterfileSutel_{fecha_hoy}.xlsx"

        # Guardar a memoria
        output_stream = BytesIO()
        edited_df.to_excel(output_stream, index=False)
        output_stream.seek(0)

        # Subir a SharePoint
        target_folder_url = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion"
        target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
        target_folder.upload_file(nuevo_nombre, output_stream).execute_query()

        st.success(f"Copia guardada como {nuevo_nombre} en SharePoint ✅")

except Exception as e:
    st.error(f"Error: {e}")
