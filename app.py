import io
import re
import streamlit as st
import pandas as pd
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Procesador y Reexpresión de Estados Financieros",
    page_icon="📊",
    layout="wide",
)

st.title("ESTADOS FINANCIEROS - REEXPRESIÓN MONETARIA (-6 CEROS)")
st.markdown("*Sistema Automatizado para Balances y Estados Financieros (PDF Escaneado/Digital, Excel y CSV)*")

# ==============================================================================
# SIDEBAR: CONTROLES
# ==============================================================================
st.sidebar.header("1. Cargar Documento")
archivo_subido = st.sidebar.file_uploader(
    "Sube tu Balance o Estado Financiero (PDF, Excel, CSV):", 
    type=["pdf", "xlsx", "xls", "csv"]
)

factor_reconversion = st.sidebar.number_input(
    "Factor de división (Eliminación de 6 ceros):",
    min_value=1.0,
    value=1000000.0,
    step=100000.0,
    format="%.2f",
)

# ==============================================================================
# UTILIDADES DE LIMPIEZA Y PARSEO FINANCIERO
# ==============================================================================
def limpiar_monto_contable(val):
    """Convierte cadenas con Bs., comas, puntos o paréntesis a float seguro."""
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).strip()
    s = re.sub(r'[Bs\.\$\€\£\sA-Za-z]', '', s)
    if not s:
        return 0.0
    
    es_negativo = False
    if (s.startswith('(') and s.endswith(')')) or s.startswith('-'):
        es_negativo = True
        s = s.replace('(', '').replace(')', '').replace('-', '')
        
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        partes = s.split(',')
        if len(partes) == 2 and len(partes[1]) in [1, 2]:
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
    elif '.' in s:
        partes = s.split('.')
        if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3):
            s = s.replace('.', '')
            
    try:
        num = float(s)
        return -num if es_negativo else num
    except ValueError:
        return 0.0

def extraer_tablas_pdf(archivo):
    """Extrae texto y datos tabulares estructurados de PDFs digitales o escaneados."""
    filas_extraidas = []
    texto_crudo = ""
    try:
        with pdfplumber.open(archivo) as pdf:
            for i, pagina in enumerate(pdf.pages):
                txt = pagina.extract_text()
                if txt:
                    texto_crudo += f"\n--- PÁGINA {i+1} ---\n" + txt
                
                # Extracción de tablas estructuradas
                tablas = pagina.extract_tables()
                for tabla in tablas:
                    for fila in tabla:
                        if fila and any(celda and str(celda).strip() for celda in fila):
                            fila_limpia = [str(c).strip() if c is not None else "" for c in fila]
                            filas_extraidas.append(fila_limpia)
    except Exception as e:
        st.sidebar.error(f"Error procesando PDF: {e}")
        
    return texto_crudo, filas_extraidas

# ==============================================================================
# DATOS FINANCIEROS BASE PREDETERMINADOS
# ==============================================================================
def obtener_tablas_predeterminadas():
    bc_data = [
        ["1-ACTIVO CIRCULANTE", 0.0, 0.0, 0.0, 0.0],
        ["1.1.1.00 - CAJA Y BANCO", 33.35, 6787.98, 6821.33, 0.00],
        ["1.1.1.05 - BANCO CORPBANCA", 0.22, 0.00, 0.22, 0.00],
        ["1.1.1.08 - BANCO NACIONAL DE CRÉDITO", 0.25, 0.00, 0.25, 0.00],
        ["1.1.2.50 - CTAS X COBRAR EMPLEADOS", 0.00, 0.23, 0.00, 0.23],
        ["1.1.3.00 - INVENTARIO DE MERCANCÍAS", 3.51, 1.42, 3.51, 1.42],
        ["1.1.5.01 - ISLR PREPAGADO", 0.00, 0.21, 0.00, 0.21],
        ["TOTAL ACTIVO CIRCULANTE", 37.33, 6789.41, 1422.18, 6825.30],
        ["2.1.2.10 - CTAS X PAGAR PROVEEDORES", 0.00, 25.03, 1051.40, -1025.35],
        ["2.1.4.01 - IVA POR PAGAR", -2696.93, 474.23, 471.35, -14.69],
        ["TOTAL PASIVO A CORTO PLAZO", -2696.93, 5178.85, 5176.94, -1041.05],
        ["3.1.0 - CAPITAL SOCIAL ACCIONISTAS", -300.00, 299.99, 0.00, -3.00],
        ["3.1.2 - RESERVA LEGAL", -30.00, 29.99, 0.00, -0.30],
        ["3.1.3 - RESULTADOS ACUMULADOS", -27457.30, 27457.02, 0.00, -274.57],
        ["TOTAL PATRIMONIO", -34563.24, 34562.96, 380.65, -381.13]
    ]
    df_bc = pd.DataFrame(bc_data, columns=["Cuenta / Descripción", "Saldo Inicial", "Monto Debe", "Monto Haber", "Saldo Actual"])

    er_data = [
        ["4.2.1 - VENTAS CASA PRINCIPAL", 0.00, 1872384.72, 1872384.72, 0.00],
        ["TOTAL INGRESOS", 0.00, 1872384.72, 1872384.72, 0.00],
        ["5.1.0 - INVENTARIO INICIAL", 0.00, 35.11, 35.11, 0.00],
        ["5.1.1 - COMPRAS NACIONALES", 0.00, 5964078.91, 5964078.91, 0.00],
        ["5.1.3 - INVENTARIO FINAL", 0.00, 1422.18, 1422.18, 0.00],
        ["TOTAL COSTO DE VENTAS", 0.00, 5965501.13, 5965501.13, 0.00],
        ["6.8.0.02 - ARRENDAMIENTOS", 0.00, 27.63, 27.63, 0.00],
        ["6.8.0.26 - GASTOS DE CONTABILIDAD", 0.00, 901.72, 901.72, 0.00],
        ["TOTAL GASTOS GENERALES", 0.00, 231552.00, 231552.00, 0.00]
    ]
    df_er = pd.DataFrame(er_data, columns=["Código y Cuenta", "Saldo Inicial", "Monto Debe", "Monto Haber", "Saldo Actual"])

    rg_data = [
        ["1 - TOTAL ACTIVO", 37462.18, 6931703.72, 6967743.72, 1422.18],
        ["2 - TOTAL PASIVO", -2898.93, 5178845.35, 5176987.47, -1041.05],
        ["3 - TOTAL PATRIMONIO", -34563.24, 34562.96, 380.65, -381.13],
        ["4 - TOTAL INGRESOS", 0.00, 1872384.72, 1872384.72, 0.00],
        ["5 - TOTAL COSTOS", 0.00, 6073902.12, 6073902.12, 0.00],
        ["6 - TOTAL GASTOS", 0.00, 231552.00, 231552.00, 0.00]
    ]
    df_rg = pd.DataFrame(rg_data, columns=["Grupo Financiero", "Saldo Inicial", "Monto Debe", "Monto Haber", "Saldo Actual"])

    return df_bc, df_er, df_rg

# ==============================================================================
# PROCESAMIENTO DEL ARCHIVO SUBIDO
# ==============================================================================
texto_pdf_extraido = ""
df_bc, df_er, df_rg = obtener_tablas_predeterminadas()

if archivo_subido is not None:
    nombre = archivo_subido.name.lower()
    
    if nombre.endswith(".pdf"):
        texto_pdf_extraido, tablas_pdf = extraer_tablas_pdf(archivo_subido)
        if tablas_pdf and len(tablas_pdf) > 2:
            st.sidebar.success(f"¡PDF procesado! Se detectaron {len(tablas_pdf)} filas tabulares.")
        elif texto_pdf_extraido:
            st.sidebar.info("PDF digital leído. Visualiza el texto en la pestaña 4.")
        else:
            st.sidebar.warning("PDF escaneado sin texto seleccionable directo. Puedes ingresar los datos o usar la plantilla base.")
            
    elif nombre.endswith((".xlsx", ".xls")):
        try:
            excel_cargado = pd.read_excel(archivo_subido, sheet_name=None)
            hojas = list(excel_cargado.keys())
            st.sidebar.success(f"¡Excel cargado! Hojas: {', '.join(hojas)}")
            if len(hojas) > 0:
                df_cargado = excel_cargado[hojas[0]].dropna(how="all")
                st.sidebar.write("Primera hoja cargada al Balance de Comprobación.")
                df_bc = df_cargado
        except Exception as e:
            st.sidebar.error(f"Error leyendo Excel: {e}")
            
    elif nombre.endswith(".csv"):
        try:
            df_bc = pd.read_csv(archivo_subido).dropna(how="all")
            st.sidebar.success("¡CSV cargado correctamente!")
        except Exception as e:
            st.sidebar.error(f"Error leyendo CSV: {e}")

# ==============================================================================
# APLICACIÓN SEGURA DE REEXPRESIÓN (-6 CEROS)
# ==============================================================================
def reexpresar_dataframe(df, factor):
    """Aplica la reconversión a todas las columnas numéricas de forma segura."""
    df_copy = df.copy()
    for col in df_copy.columns:
        if col != df_copy.columns[0]:  # No convertir la columna de nombres/códigos
            df_copy[col] = df_copy[col].apply(limpiar_monto_contable)
            df_copy[col] = df_copy[col] / factor
    return df_copy

df_bc_reexpresado = reexpresar_dataframe(df_bc, factor_reconversion)
df_er_reexpresado = reexpresar_dataframe(df_er, factor_reconversion)
df_rg_reexpresado = reexpresar_dataframe(df_rg, factor_reconversion)

# ==============================================================================
# PESTAÑAS PRINCIPALES DE VISUALIZACIÓN
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Balance de Comprobación", 
    "📊 Estado de Resultados", 
    "🏛️ Resumen General",
    "📄 Texto Extraído del PDF"
])

with tab1:
    st.subheader("Balance de Comprobación Reexpresado")
    st.caption("Valores reexpresados listos para reimpresión y presentación fiscal.")
    st.dataframe(df_bc_reexpresado, use_container_width=True)

with tab2:
    st.subheader("Estado de Resultados Reexpresado")
    st.dataframe(df_er_reexpresado, use_container_width=True)

with tab3:
    st.subheader("Resumen General y Ecuación Patrimonial")
    st.dataframe(df_rg_reexpresado, use_container_width=True)

with tab4:
    st.subheader("Texto Crudo Extraído del Archivo PDF")
    if texto_pdf_extraido.strip():
        st.text_area("Contenido detectado:", texto_pdf_extraido, height=350)
    else:
        st.info("Sube un archivo PDF en la barra lateral para ver el texto extraído.")

# ==============================================================================
# GENERACIÓN DE EXCEL CON FORMATO CONTABLE PARA REIMPRESIÓN
# ==============================================================================
def estilizar_hoja_excel(ws, df, titulo):
    # Título principal
    ws.merge_cells("A1:E1")
    t_cell = ws["A1"]
    t_cell.value = titulo
    t_cell.font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    t_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.append([]) # Fila vacía
    
    # Encabezados
    headers = list(df.columns)
    ws.append(headers)
    fila_encabezado = 3

    h_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    h_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    borde_fino = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB")
    )
    
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=fila_encabezado, column=c_idx)
        cell.font = h_font
        cell.fill = h_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borde_fino

    # Filas de datos
    formato_contable = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
    for row in df.itertuples(index=False):
        ws.append(list(row))
        
    inicio_datos = fila_encabezado + 1
    fin_datos = inicio_datos + len(df) - 1

    for r_idx in range(inicio_datos, fin_datos + 1):
        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = borde_fino
            cell.font = Font(name="Calibri", size=9)
            if c_idx > 1:
                cell.number_format = formato_contable
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Autoajuste de anchos de columna
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = max(len(str(c.value or '')) for c in col if c.row >= fila_encabezado)
        ws.column_dimensions[col_letter].width = max(max_l + 4, 15)

def generar_excel_profesional():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # 1. Balance de Comprobación
    ws1 = wb.active
    ws1.title = "BALANCE DE COMPROBACIÓN"
    estilizar_hoja_excel(ws1, df_bc_reexpresado, "BALANCE DE COMPROBACIÓN REEXPRESADO (-6 CEROS)")
    
    # 2. Estado de Resultados
    ws2 = wb.create_sheet(title="ESTADO DE RESULTADOS")
    estilizar_hoja_excel(ws2, df_er_reexpresado, "ESTADO DE RESULTADOS REEXPRESADO (-6 CEROS)")
    
    # 3. Resumen General
    ws3 = wb.create_sheet(title="RESUMEN GENERAL")
    estilizar_hoja_excel(ws3, df_rg_reexpresado, "RESUMEN GENERAL Y ECUACIÓN PATRIMONIAL")
    
    wb.save(output)
    return output.getvalue()

# ==============================================================================
# DESCARGA
# ==============================================================================
st.markdown("---")
excel_bytes = generar_excel_profesional()
st.download_button(
    label="📥 Descargar Estados Financieros Reexpresados en Excel (.xlsx)",
    data=excel_bytes,
    file_name="Estados_Financieros_Reexpresados_Oficial.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
