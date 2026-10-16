import io
import json
import base64
import re
import requests
import pandas as pd
import streamlit as st
import pypdfium2 as pdfium
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Balance de Comprobación - Reexpresión Monetaria",
    page_icon="📑",
    layout="wide",
)

st.title("ESTADOS FINANCIEROS - REEXPRESIÓN MONETARIA (-6 CEROS)")
st.markdown("*Conversión y Maquetación Contable Profesional idéntica al Formato Oficial*")

# ==============================================================================
# GESTIÓN DE CREDENCIALES Y PARÁMETROS
# ==============================================================================
gemini_api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        gemini_api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

st.sidebar.header("⚙️ Configuración")
if not gemini_api_key:
    gemini_api_key = st.sidebar.text_input(
        "Gemini API Key:", 
        type="password",
        help="Clave gratuita de https://aistudio.google.com para leer el PDF escaneado con IA"
    )

st.sidebar.markdown("---")
st.sidebar.subheader("📄 Cargar Documento")
archivo_subido = st.sidebar.file_uploader(
    "Sube el Balance de Comprobación (PDF Escaneado):", 
    type=["pdf"]
)

factor_reconversion = st.sidebar.number_input(
    "Factor de Reexpresión (÷ 1,000,000):",
    min_value=1.0,
    value=1000000.0,
    step=100000.0,
    format="%.2f",
)

# Datos de membrete editables
st.sidebar.markdown("---")
st.sidebar.subheader("🏢 Datos del Encabezado")
empresa_nombre = st.sidebar.text_input("Razón Social:", value="TRATTORIA - PIZZERIA DA LUCY, C.A.")
empresa_rif = st.sidebar.text_input("R.I.F.:", value="J-29495819-4")
periodo_texto = st.sidebar.text_input("Período:", value="DEL 01/01/2018 AL 31/12/2018 (Cifras en Bolívares)")

# ==============================================================================
# CONVERTIDOR NUMÉRICO SEGURO
# ==============================================================================
def a_numero(val):
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r'[Bs\.\$\€\£\sA-Za-z]', '', s)
    if not s:
        return 0.0
    es_neg = s.startswith('-') or (s.startswith('(') and s.endswith(')'))
    s = s.replace('(', '').replace(')', '').replace('-', '')
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        n = float(s)
        return -n if es_neg else n
    except ValueError:
        return 0.0

# ==============================================================================
# MOTOR OCR CON GEMINI VISION
# ==============================================================================
def extraer_datos_con_gemini(pdf_bytes, api_key):
    try:
        pdf = pdfium.PdfDocument(pdf_bytes)
        images_parts = []
        total_paginas = min(len(pdf), 8)
        
        for i in range(total_paginas):
            page = pdf[i]
            image = page.render(scale=1.5).to_pil()
            buffered = io.BytesIO()
            image.save(buffered, format="JPEG", quality=75)
            img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
            images_parts.append({
                "inline_data": {
                    "mime_type": "image/jpeg",
                    "data": img_str
                }
            })

        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
        
        prompt_text = """
        Eres un contador público colegiado y auditor. Analiza todas las páginas escaneadas del Balance de Comprobación adjunto.
        Extrae TODAS las cuentas, subtotales y totales en orden jerárquico contable exacto:
        - 1-ACTIVO, 1.1-ACTIVO CIRCULANTE, TOTAL EFECTIVO, TOTAL ACTIVO
        - 2-PASIVO, PASIVO A CORTO PLAZO, TOTAL PASIVO
        - 3-PATRIMONIO, TOTAL PATRIMONIO
        - 4-INGRESOS, 5-COSTOS, 6-GASTOS, TOTALES

        Devuelve ÚNICAMENTE un objeto JSON con la llave "cuentas", donde cada elemento tenga:
        - "cuenta": Código y Nombre de la cuenta (ej: "1.1.1.00-CAJA Y BANCO" o "TOTAL EFECTIVO")
        - "saldo_inicial": número float (o 0 si está vacío)
        - "monto_debe": número float (o 0 si está vacío)
        - "monto_haber": número float (o 0 si está vacío)
        - "saldo_actual": número float (saldo final)
        
        Respeta los montos exactos originales escaneados sin redondear. No agregues comentarios fuera del JSON.
        """

        contents = [{"parts": [{"text": prompt_text}] + images_parts}]
        payload = {
            "contents": contents,
            "generationConfig": {
                "response_mime_type": "application/json",
                "temperature": 0.1
            }
        }

        headers = {"Content-Type": "application/json"}
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        
        if response.status_code == 200:
            res_json = response.json()
            texto_respuesta = res_json["candidates"][0]["content"]["parts"][0]["text"].strip()
            match = re.search(r'\{.*\}', texto_respuesta, re.DOTALL)
            json_str = match.group(0) if match else texto_respuesta
            datos = json.loads(json_str)
            return pd.DataFrame(datos.get("cuentas", []))
        else:
            st.error(f"Error en la API de Gemini ({response.status_code}): {response.text}")
            return None
    except Exception as e:
        st.error(f"Error procesando el documento con IA: {e}")
        return None

# ==============================================================================
# PLANTILLA BASE FIEL A LA ESTRUCTURA DE LA EMPRESA (TRATTORIA DA LUCY)
# ==============================================================================
def obtener_estructura_oficial():
    data = [
        {"cuenta": "1-ACTIVO", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.1-ACTIVO CIRCULANTE", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.1-EFECTIVO", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.1.00-CAJA Y BANCO", "saldo_inicial": 33250000.0, "monto_debe": 6787980000.0, "monto_haber": 6821330000.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.1.05-BANCO: CORPBANCA", "saldo_inicial": 220000.0, "monto_debe": 0.0, "monto_haber": 220000.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.1.08-BANCO: NACIONAL DE CREDITO", "saldo_inicial": 250000.0, "monto_debe": 0.0, "monto_haber": 250000.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL EFECTIVO", "saldo_inicial": 33720000.0, "monto_debe": 6787980000.0, "monto_haber": 6821800000.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.2-CUENTAS POR COBRAR", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.2.50-CTAS X COBRAR Empleados", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL CUENTAS POR COBRAR", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.3-INVENTARIOS", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.3.00-INVENTARIO DE MERCANCIAS", "saldo_inicial": 3510000.0, "monto_debe": 1420000.0, "monto_haber": 3510000.0, "saldo_actual": 1420000.0},
        {"cuenta": "1.1.3.01-REVAL. INVENTARIO DE MERCANCIAS", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL INVENTARIOS", "saldo_inicial": 3510000.0, "monto_debe": 1420000.0, "monto_haber": 3510000.0, "saldo_actual": 1420000.0},
        {"cuenta": "1.1.5-PREPAGADOS", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.1.5.01-I.S.L.R. PREPAGADO", "saldo_inicial": 0.0, "monto_debe": 210000.0, "monto_haber": 210000.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL PREPAGADOS", "saldo_inicial": 0.0, "monto_debe": 210000.0, "monto_haber": 210000.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL ACTIVO CIRCULANTE", "saldo_inicial": 37230000.0, "monto_debe": 6789610000.0, "monto_haber": 6825520000.0, "saldo_actual": 1420000.0},
        {"cuenta": "1.2-ACTIVO FIJO", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.2.1-MAQUINARIAS Y EQUIPOS", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.2.1.01-COSTO MAQUINARIAS Y EQUIPOS", "saldo_inicial": 200000.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 200000.0},
        {"cuenta": "1.2.1.02-D.A. MAQUINARIAS Y EQUIPOS", "saldo_inicial": -10000.0, "monto_debe": 10000.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "1.2.1.03-Revaluacion MAQUINARIAS Y EQUIPOS", "saldo_inicial": 0.0, "monto_debe": 100000.0, "monto_haber": 100000.0, "saldo_actual": 0.0},
        {"cuenta": "1.2.1.04-D.A. REVALUACION", "saldo_inicial": 0.0, "monto_debe": 10000.0, "monto_haber": 10000.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL MAQUINARIAS Y EQUIPOS", "saldo_inicial": 190000.0, "monto_debe": 120000.0, "monto_haber": 110000.0, "saldo_actual": 200000.0},
        {"cuenta": "TOTAL ACTIVO", "saldo_inicial": 37420000.0, "monto_debe": 6931700000.0, "monto_haber": 6967740000.0, "saldo_actual": 1420000.0},
        {"cuenta": "2-PASIVO", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "2.1-PASIVO A CORTO PLAZO", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "2.1.2-CUENTAS POR PAGAR", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "2.1.2.01-CTAS X PAGAR Proveedores", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 1070000.0, "saldo_actual": -1070000.0},
        {"cuenta": "2.1.2.10-CTAS X PAGAR ACCIONISTAS", "saldo_inicial": 0.0, "monto_debe": 4497670000.0, "monto_haber": 4497670000.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL CUENTAS POR PAGAR", "saldo_inicial": 0.0, "monto_debe": 4497670000.0, "monto_haber": 4498740000.0, "saldo_actual": -1070000.0},
        {"cuenta": "TOTAL PASIVO", "saldo_inicial": -2890000.0, "monto_debe": 5178850000.0, "monto_haber": 5176940000.0, "saldo_actual": -1040000.0},
        {"cuenta": "3-PATRIMONIO", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL PATRIMONIO", "saldo_inicial": -34530000.0, "monto_debe": 34500000.0, "monto_haber": 580000.0, "saldo_actual": -380000.0},
        {"cuenta": "4-INGRESOS", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "4.2.1-VENTAS CASA PRINCIPAL", "saldo_inicial": 0.0, "monto_debe": 1872380000.0, "monto_haber": 1872380000.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL INGRESOS", "saldo_inicial": 0.0, "monto_debe": 1872380000.0, "monto_haber": 1872380000.0, "saldo_actual": 0.0},
        {"cuenta": "5-COSTOS", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "5.1.1-COMPRAS NACIONALES", "saldo_inicial": 0.0, "monto_debe": 5964080000.0, "monto_haber": 5964080000.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL COSTOS", "saldo_inicial": 0.0, "monto_debe": 6074000000.0, "monto_haber": 6074000000.0, "saldo_actual": 0.0},
        {"cuenta": "6-GASTOS", "saldo_inicial": 0.0, "monto_debe": 0.0, "monto_haber": 0.0, "saldo_actual": 0.0},
        {"cuenta": "TOTAL GASTOS", "saldo_inicial": 0.0, "monto_debe": 231550000.0, "monto_haber": 231550000.0, "saldo_actual": 0.0},
        {"cuenta": "TOTALES", "saldo_inicial": 0.0, "monto_debe": 20122950000.0, "monto_haber": 20122950000.0, "saldo_actual": 0.0}
    ]
    return pd.DataFrame(data)

# ==============================================================================
# CARGA Y EXTRACCIÓN
# ==============================================================================
df_datos = None

if archivo_subido is not None:
    if not gemini_api_key:
        st.warning("⚠️ Ingresa tu `GEMINI_API_KEY` en la barra lateral para procesar el escaneo con Inteligencia Artificial.")
    else:
        with st.spinner("🤖 Digitalizando escaneo con Gemini Vision OCR..."):
            pdf_bytes = archivo_subido.read()
            df_ia = extraer_datos_con_gemini(pdf_bytes, gemini_api_key)
            if df_ia is not None and not df_ia.empty:
                st.success(f"✨ ¡Digitalizado con éxito! {len(df_ia)} cuentas procesadas.")
                df_datos = df_ia

if df_datos is None:
    df_datos = obtener_estructura_oficial()

for col in ["saldo_inicial", "monto_debe", "monto_haber", "saldo_actual"]:
    if col in df_datos.columns:
        df_datos[col] = df_datos[col].apply(a_numero)

st.subheader("📋 Validación de Cuentas")
st.caption("Puedes corregir o agregar cualquier cuenta en la tabla interactiva:")
df_editado = st.data_editor(df_datos, num_rows="dynamic", use_container_width=True)

# Reexpresión matemática (-6 ceros)
df_reexpresado = pd.DataFrame()
df_reexpresado["Nombre de la Cuenta"] = df_editado["cuenta"]
df_reexpresado["Saldo Inicial"] = df_editado["saldo_inicial"].apply(a_numero) / factor_reconversion
df_reexpresado["Monto Debe"] = df_editado["monto_debe"].apply(a_numero) / factor_reconversion
df_reexpresado["Monto Haber"] = df_editado["monto_haber"].apply(a_numero) / factor_reconversion
df_reexpresado["Saldo Actual"] = df_editado["saldo_actual"].apply(a_numero) / factor_reconversion

# ==============================================================================
# GENERACIÓN DE EXCEL CON DISEÑO CONTABLE EXACTO
# ==============================================================================
def generar_excel_impresion_oficial(df, factor, empresa, rif, periodo):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Balance de Comprobacion"

    # Fuentes y Colores Corporativos
    font_empresa = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    font_sub = Font(name="Calibri", size=8, color="4B5563")
    font_title = Font(name="Calibri", size=9, bold=True, color="1E3A8A")
    
    font_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    
    font_nivel1 = Font(name="Calibri", size=9, bold=True, color="1E3A8A")
    fill_nivel1 = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
    
    font_total = Font(name="Calibri", size=9, bold=True, color="000000")
    fill_total = PatternFill(start_color="E5EDF6", end_color="E5EDF6", fill_type="solid")
    
    font_gran_total = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    fill_gran_total = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    
    font_regular = Font(name="Calibri", size=8, color="000000")

    # Bordes
    border_celda = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    border_total = Border(
        left=Side(style='thin', color='B0C4DE'),
        right=Side(style='thin', color='B0C4DE'),
        top=Side(style='thin', color='1E3A8A'),
        bottom=Side(style='thin', color='1E3A8A')
    )
    border_gran_total = Border(
        left=Side(style='thin', color='1E3A8A'),
        right=Side(style='thin', color='1E3A8A'),
        top=Side(style='thin', color='1E3A8A'),
        bottom=Side(style='double', color='1E3A8A')
    )

    # Formato Contable: comas para miles, punto para decimales, guión para cero, paréntesis para negativos
    formato_contable = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

    # Membrete institucional
    ws["A1"] = empresa.upper()
    ws["A1"].font = font_empresa

    ws["A2"] = f"RIF: {rif.upper()}"
    ws["A2"].font = font_sub

    ws["A3"] = "BALANCE DE COMPROBACIÓN"
    ws["A3"].font = font_title

    ws["A4"] = f"{periodo} - Factor Reexpresión: ÷ {factor:,.2f}"
    ws["A4"].font = font_sub

    # Encabezados de Tabla (Fila 6)
    fila_h = 6
    headers = ["Nombre de la Cuenta", "Saldo Inicial", "Monto Debe", "Monto Haber", "Saldo Actual"]
    ws.append([]) # Fila 5 vacía
    ws.append(headers)

    for col_idx in range(1, 6):
        c = ws.cell(row=fila_h, column=col_idx)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_celda

    # Escribir filas con formato condicional jerárquico
    for _, row in df.iterrows():
        nombre = str(row["Nombre de la Cuenta"]).strip()
        valores = [
            nombre,
            row["Saldo Inicial"],
            row["Monto Debe"],
            row["Monto Haber"],
            row["Saldo Actual"]
        ]
        ws.append(valores)
        r_actual = ws.max_row
        nombre_mayus = nombre.upper()

        es_gran_total = "TOTALES" in nombre_mayus or "TOTAL GENERAL" in nombre_mayus
        es_total = nombre_mayus.startswith("TOTAL") and not es_gran_total
        es_grupo_mayor = bool(re.match(r'^[1-6]\s*-|^[1-6]\.[0-9]\s*-', nombre)) and not es_total

        # Aplicar estilos por tipo de fila
        for c_idx in range(1, 6):
            cell = ws.cell(row=r_actual, column=c_idx)

            if es_gran_total:
                cell.font = font_gran_total
                cell.fill = fill_gran_total
                cell.border = border_gran_total
            elif es_total:
                cell.font = font_total
                cell.fill = fill_total
                cell.border = border_total
            elif es_grupo_mayor:
                cell.font = font_nivel1
                cell.fill = fill_nivel1
                cell.border = border_celda
            else:
                cell.font = font_regular
                cell.border = border_celda

            # Formato numérico para columnas de cifras
            if c_idx > 1:
                cell.number_format = formato_contable
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                if "." in nombre and not es_total and not es_grupo_mayor:
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

    # Configuración de anchos de columna para visualización e impresión perfecta
    ws.column_dimensions['A'].width = 44
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 16

    # Opciones de página para imprimir
    ws.sheet_view.showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.markdown("---")
st.subheader("Vista Previa del Balance Reexpresado")
st.dataframe(df_reexpresado, use_container_width=True)

excel_oficial = generar_excel_impresion_oficial(
    df_reexpresado, 
    factor_reconversion, 
    empresa_nombre, 
    empresa_rif, 
    periodo_texto
)

st.download_button(
    label="📥 Descargar Balance de Comprobación Oficial (.xlsx)",
    data=excel_oficial,
    file_name="Balance_De_Comprobacion_Reexpresado_Oficial.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
