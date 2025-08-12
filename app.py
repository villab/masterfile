import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from datetime import datetime

# --- Configuración de página ---
st.set_page_config(page_title="Masterfile Viewer", layout="wide")

# --- Cargar el Excel ---
ruta_excel = "masterfile.xlsx"  # Cambia esto por tu ruta o lógica de carga
df = pd.read_excel(ruta_excel)

nombre_archivo = ruta_excel.split("/")[-1]
st.success(f"Cargado masterfile del día: **{nombre_archivo}**")

# --- Configurar tabla grande y ancha ---
gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_pagination(enabled=False)  # Sin paginación
gb.configure_default_column(resizable=True, filter=True, sortable=True)  # Columnas filtrables y ordenables
gb.configure_grid_options(
    domLayout='autoHeight',             # Ajuste de altura automática
    suppressSizeToFit=True,             # No comprimir columnas
    alwaysShowHorizontalScroll=True     # Scroll horizontal
)
grid_options = gb.build()

# --- CSS para expandir ancho ---
st.markdown("""
    <style>
    .ag-theme-balham {
        width: 100% !important;
        height: 700px !important; /* Ajusta alto si lo deseas */
    }
    </style>
    """, unsafe_allow_html=True)

# --- Mostrar la tabla ---
AgGrid(
    df,
    gridOptions=grid_options,
    fit_columns_on_grid_load=False,
    enable_enterprise_modules=False,
    update_mode=GridUpdateMode.NO_UPDATE,
    allow_unsafe_jscode=True,
    theme="balham"
)
