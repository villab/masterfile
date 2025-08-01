import streamlit as st
import pandas as pd
from io import BytesIO
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# -------------------- CONFIG --------------------
USERNAME = st.secrets["sharepoint_user"]
PASSWORD = st.secrets["sharepoint_pass"]
ACCESS_KEY = st.secrets["app_password"]  # Contraseña de acceso a la app

SITE_URL = "https://tuempresa.sharepoint.com/sites/MiSitio"
LIBRARY_NAME = "Documentos compartidos"

# -------------------- SHAREPOINT FUNCTIONS --------------------
@st.cache_resource
def connect_sharepoint():
    return ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))

def list_files(ctx):
    library = ctx.web.lists.get_by_title(LIBRARY_NAME)
    files = library.root_folder.files
    ctx.load(files)
    ctx.execute_query()
    return [f.properties["Name"] for f in files if f.properties["Name"].endswith(".xlsx")]

def download_file(ctx, file_name):
    file_url = f"/sites/MiSitio/{LIBRARY_NAME}/{file_name}"
    file = ctx.web.get_file_by_server_relative_url(file_url).download()
    ctx.execute_query()
    return file.content

def upload_file(ctx, file_name, content):
    target_folder = ctx.web.lists.get_by_title(LIBRARY_NAME).root_folder
    target_folder.upload_file(file_name, content)
    ctx.execute_query()

# -------------------- LOGIN --------------------
st.set_page_config(page_title="Gestor Multiusuario XLSX", layout="wide")
st.title("🔐 Gestor Multiusuario de Datos en SharePoint")

password_input = st.text_input("Ingresa la clave de acceso", type="password")
if password_input != ACCESS_KEY:
    st.warning("Introduce la clave correcta para acceder.")
    st.stop()

st.success("✅ Acceso concedido")

# -------------------- APP --------------------
ctx = connect_sharepoint()

# 1. Seleccionar archivo existente o cargar uno nuevo
files = list_files(ctx)
file_choice = st.selectbox("Selecciona un archivo de SharePoint", [""] + files)

uploaded_file = st.file_uploader("O carga un archivo nuevo", type=["xlsx"])

if uploaded_file:
    # Si el usuario carga un archivo, lo subimos a SharePoint
    upload_file(ctx, uploaded_file.name, uploaded_file.getvalue())
    st.success(f"Archivo '{uploaded_file.name}' cargado a SharePoint ✅")
    file_choice = uploaded_file.name

if file_choice:
    # 2. Descargar y mostrar
    file_bytes = download_file(ctx, file_choice)
    df = pd.read_excel(BytesIO(file_bytes))
    st.success(f"Archivo '{file_choice}' cargado desde SharePoint ✅")

    # 3. Filtros
    st.subheader("Filtros dinámicos")
    filter_cols = st.multiselect("Selecciona columnas para filtrar", df.columns)
    filtered_df = df.copy()
    for col in filter_cols:
        valores = st.multiselect(f"Filtrar {col}", df[col].unique())
        if valores:
            filtered_df = filtered_df[filtered_df[col].isin(valores)]

    # 4. Edición
    st.subheader("Editar datos")
    edited_df = st.data_editor(filtered_df, num_rows="dynamic", use_container_width=True)

    # 5. Exportar a XLSX
    st.subheader("Exportar datos editados")

    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Datos Editados")
        return output.getvalue()

    excel_data = to_excel(edited_df)

    st.download_button(
        label="📥 Descargar Excel editado",
        data=excel_data,
        file_name=f"editado_{file_choice}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("💾 Guardar cambios en SharePoint"):
        upload_file(ctx, file_choice, excel_data)
        st.success("Archivo actualizado en SharePoint ✅")
