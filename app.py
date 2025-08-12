import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import os

# ================== CONFIGURACIÓN ==================
st.set_page_config(layout="wide")

USERNAME = st.secrets["sharepoint_user"]
APP_PASSWORD = st.secrets["app_password"]

SITE_URL = "https://caseonit.sharepoint.com/sites/Sutel"
FILE_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion/MasterfileSutel.xlsx"
FOLDER_URL = "/sites/Sutel/Documentos compartidos/01. Documentos MedUX/Automatizacion"
BACKUP_FOLDER_URL = f"{FOLDER_URL}/Backups"

try:
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, APP_PASSWORD))

    # Obtener nombre del archivo
    nombre_archivo = os.path.basename(FILE_URL)

    # Descargar archivo
    file = ctx.web.get_file_by_server_relative_url(FILE_URL)
    file_stream = BytesIO()
    file.download(file_stream).execute_query()
    file_stream.seek(0)

    # ================== LECTURA ==================
    df = pd.read_excel(file_stream)
    st.success(f"📂 Cargado masterfile del día: {nombre_archivo} ✅") 
    st.write(f"**Total registros:** {len(df):,}")

    # ================== TABLA CON FILTROS ==================
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_pagination(enabled=True, paginationAutoPageSize=False, paginationPageSize=20)
    gb.configure_default_column(resizable=True, filter=True, sortable=True)
    gb.configure_grid_options(domLayout='normal')
    grid_options = gb.build()

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        theme="balham",
        height=600,
        fit_columns_on_grid_load=True
    )

    edited_df = pd.DataFrame(grid_response['data'])

    # ================== GUARDAR CAMBIOS ==================
    if st.button("💾 Guardar nueva versión de Masterfile"):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nuevo_nombre = f"MasterfileSutel_{timestamp}.xlsx"

        output_stream = BytesIO()
        edited_df.to_excel(output_stream, index=False)
        output_stream.seek(0)

        # Crear carpeta Backups si no existe
        try:
            ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL).expand(["Files"]).get().execute_query()
        except:
            ctx.web.folders.add(BACKUP_FOLDER_URL).execute_query()

        # Subir copia con fecha
        backup_folder = ctx.web.get_folder_by_server_relative_url(BACKUP_FOLDER_URL)
        backup_folder.upload_file(nuevo_nombre, output_stream).execute_query()

        # Sobrescribir original
        output_stream.seek(0)
        main_folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
        main_folder.upload_file("MasterfileSutel.xlsx", output_stream).execute_query()

        st.success(f"✅ Cambios guardados y copia creada como {nuevo_nombre}")

except Exception as e:
    st.error(f"Error: {e}")
