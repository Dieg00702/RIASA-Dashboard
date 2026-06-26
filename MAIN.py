import streamlit as st

# 1. Configurar la pantalla ancha
st.set_page_config(layout="wide", page_title="Dashboard RIASA")

st.markdown("""
    <style>
    /* 1. Aprovecha el ancho de la pantalla reduciendo márgenes externos globales */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1.5rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        max-width: 100% !important;
    }
    
    /* 2. Controla la separación entre las 3 columnas principales */
    div[data-testid="stHorizontalBlock"] {
        gap: 1.2rem !important;
    }
    
    /* 3. Devuelve el margen interno (padding) correcto a las tarjetas */
    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        padding: 20px !important;
    }
    
    /* 4. Asegura que los textos internos no queden pegados a los extremos */
    div[data-testid="element-container"] {
        padding-left: 5px !important;
        padding-right: 5px !important;
    }
    </style>
""", unsafe_allow_html=True)

# 2. Importar las funciones de tus 3 archivos independientes
from FACT import mostrar_facturacion
from IND import mostrar_indicadores
from GRAFICA import mostrar_graficas

# Título global
st.title("Panel de Control Operativo — RIASA")

# 3. Crear las 3 columnas principales en la pantalla
col_fact, col_ind, col_graf = st.columns(3)
# 4. Asignar cada función a su respectiva columna
with col_fact:
    st.subheader("📊 Producción")
    mostrar_graficas()

with col_ind:
    st.subheader("📈 Facturacion")
    mostrar_facturacion()
    
with col_graf:
    st.subheader("🏭 Indicadores")
    mostrar_indicadores()

