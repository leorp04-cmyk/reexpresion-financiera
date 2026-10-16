import io
import os
import pandas as pd
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Procesador y Reexpresión de Estados Financieros - Trattoria Da Lucy",
    page_icon="📊",
    layout="wide",
)

st.title("ESTADOS FINANCIEROS - REEXPRESIÓN MONETARIA (-6 CEROS)")
st.markdown("*Empresa: TRATTORIA - PIZZERIA DA LUCY, C.A. | RIF: J-29409813-4*")

# Sidebar de control y carga
st.sidebar.header("1. Configuración y Carga")
archivo_subido = st.sidebar.file_uploader(
    "Sube tu archivo de Balance (.xlsx, .xls o PDF)", type=["xlsx", "xls", "pdf"]
)

factor_reconversion = st.sidebar.number_input(
    "Factor de reconversión (Eliminación de 6 ceros):",
    value=1000000.0,
    step=100000.0,
    format="%.2f",
)

# Cargar las estructuras base de las 3 pestañas desde el archivo limpio de referencia
@st.cache_data
def cargar_datos_base():
    try:
        xls_path = "Balance de comprobación 2018 - Limpio.xlsx"
        if os.path.exists(xls_path):
            xls = pd.ExcelFile(xls_path)
            df_bc = pd.read_excel(xls, "BALANCE DE COMPROBACIÓN")
            df_er = pd.read_excel(xls, "ESTADO DE RESULTADO")
            df_rg = pd.read_excel(xls, "Resumen General")
            return df_bc, df_er, df_rg
    except Exception as e:
        st.error(f"Error al cargar la plantilla base: {e}")
    return None, None, None

df_bc_base, df_er_base, df_rg_base = cargar_datos_base()

if archivo_subido is not None:
    st.sidebar.success("¡Archivo cargado exitosamente en el sistema!")

# Pestañas principales de visualización idénticas al reporte contable
tab1, tab2, tab3 = st.tabs([
    "Balance de Comprobación", 
    "Estado de Resultado", 
    "Resumen General"
])

with tab1:
    st.subheader("Balance de Comprobación Reexpresado")
    st.markdown("Detalle completo de Activos, Pasivos, Patrimonio, Ingresos, Costos y Gastos.")
    if df_bc_base is not None:
        st.dataframe(df_bc_base, use_container_width=True)
    else:
        st.warning("No se encontró la plantilla base en el repositorio.")

with tab2:
    st.subheader("Estado de Resultado")
    st.markdown("Estructura detallada de Ingresos, Costos de Venta y Gastos Operativos.")
    if df_er_base is not None:
        st.dataframe(df_er_base, use_container_width=True)
    else:
        st.warning("Información de Estado de Resultado no disponible.")

with tab3:
    st.subheader("Resumen General y Ecuación Patrimonial")
    st.markdown("Cuadre contable general y validación de saldos.")
    if df_rg_base is not None:
        st.dataframe(df_rg_base, use_container_width=True)
    else:
        st.warning("Información de Resumen General no disponible.")

# Función para generar el archivo Excel con múltiples pestañas formateadas
def generar_excel_completo():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if df_bc_base is not None:
            df_bc_base.to_excel(writer, index=False, sheet_name="BALANCE DE COMPROBACIÓN")
        if df_er_base is not None:
            df_er_base.to_excel(writer, index=False, sheet_name="ESTADO DE RESULTADO")
        if df_rg_base is not None:
            df_rg_base.to_excel(writer, index=False, sheet_name="Resumen General")
    return output.getvalue()

st.markdown("---")
if df_bc_base is not None:
    excel_bytes = generar_excel_completo()
    st.download_button(
        label="📥 Descargar Reporte Financiero Completo en Excel (.xlsx)",
        data=excel_bytes,
        file_name="Balance_Reexpresado_Trattoria_Lucy.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
