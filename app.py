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

st.set_page_config(
    page_title="Procesador y Reexpresión de Estados Financieros",
    page_icon="📊",
    layout="wide",
)

st.title("ESTADOS FINANCIEROS - REEXPRESIÓN MONETARIA (-6 CEROS)")
st.markdown("*Sistema de Digitalización Inteligente por IA y Generación de Excel Contable*")

# 1. Obtener la API Key desde st.secrets o input manual en la barra lateral
gemini_api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        gemini_api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

st.sidebar.header("1. Configuración de API y Documento")
if not gemini_api_key:
    gemini_api_key = st.sidebar.text_input(
        "Ingresa tu Gemini API Key:", 
        type="password",
        help="Obtén tu clave gratis en https://aistudio.google.com"
    )

archivo_subido = st.sidebar.file_uploader(
    "Sube tu Balance de Comprobación en PDF", type=["pdf"]
)

factor_reconversion = st.sidebar.number_input(
    "Factor de división (Eliminación de 6 ceros):",
    min_value=1.0,
    value=1000000.0,
    step=100000.0,
    format="%.2f",
)

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
        
        prompt_text = (
            "Eres un contador experto y auditor contable. Analiza estas páginas escaneadas del Balance de Comprobación. "
            "Extrae todas las cuentas contables con sus códigos, nombres y cifras numéricas. "
            "Devuelve un objeto JSON con una lista 'cuentas', donde cada elemento tenga: "
            "'cuenta' (string con código y nombre), 'saldo_inicial' (número o 0), 'monto_debe' (número o 0), "
            "'monto_haber' (número o 0) y 'saldo_actual' (número con saldo final). "
            "Devuelve ÚNICAMENTE el JSON sin explicaciones ni markdown."
        )

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

df_resultado = None

if archivo_subido is not None:
    if not gemini_api_key:
        st.warning("⚠️ Ingresa tu `GEMINI_API_KEY` en la barra lateral o en los Secrets de Streamlit para procesar el escaneo.")
    else:
        with st.spinner("🤖 Procesando PDF escaneado con Gemini Vision OCR..."):
            pdf_bytes = archivo_subido.read()
            df_ia = extraer_datos_con_gemini(pdf_bytes, gemini_api_key)
            if df_ia is not None and not df_ia.empty:
                st.success(f"✨ ¡OCR completado! Se digitalizaron {len(df_ia)} cuentas del documento escaneado.")
                df_resultado = df_ia
            else:
                st.info("Usando plantilla de respaldo interactiva.")

if df_resultado is None:
    datos_respaldo = [
        {"cuenta": "1.1.1.00 - CAJA Y BANCO", "saldo_inicial": 33348659.44, "monto_debe": 6787963434.55, "monto_haber": 6821332093.99, "saldo_actual": 0.00},
        {"cuenta": "1.1.1.05 - BANCO CORPBANCA", "saldo_inicial": 220000.00, "monto_debe": 0.00, "monto_haber": 220000.00, "saldo_actual": 0.00},
        {"cuenta": "1.1.3.00 - INVENTARIO DE MERCANCÍAS", "saldo_inicial": 3510464.00, "monto_debe": 1422181.81, "monto_haber": 1422181.80, "saldo_actual": 3510454.01},
        {"cuenta": "2.1.2.10 - CTAS X PAGAR PROVEEDORES", "saldo_inicial": 0.00, "monto_debe": 25032.48, "monto_haber": 1051396.76, "saldo_actual": -1025354.25},
        {"cuenta": "4.2.1 - VENTAS CASA PRINCIPAL", "saldo_inicial": 0.00, "monto_debe": 1872384718.59, "monto_haber": 1872384718.59, "saldo_actual": 0.00}
    ]
    df_resultado = pd.DataFrame(datos_respaldo)

for col in ["saldo_inicial", "monto_debe", "monto_haber", "saldo_actual"]:
    if col in df_resultado.columns:
        df_resultado[col] = df_resultado[col].apply(a_numero)

st.subheader("2. Validación y Edición de Cuentas")
st.caption("Puedes editar cualquier valor haciendo doble clic sobre la celda:")
df_editado = st.data_editor(df_resultado, num_rows="dynamic", use_container_width=True)

df_reexpresado = pd.DataFrame()
df_reexpresado["Cuenta / Descripción"] = df_editado["cuenta"]
df_reexpresado["Saldo Inicial"] = df_editado["saldo_inicial"].apply(a_numero) / factor_reconversion
df_reexpresado["Monto Debe"] = df_editado["monto_debe"].apply(a_numero) / factor_reconversion
df_reexpresado["Monto Haber"] = df_editado["monto_haber"].apply(a_numero) / factor_reconversion
df_reexpresado["Saldo Actual"] = df_editado["saldo_actual"].apply(a_numero) / factor_reconversion

st.markdown("---")
st.subheader("3. Vista Previa del Balance Reexpresado")
st.dataframe(df_reexpresado, use_container_width=True)

def generar_excel_profesional(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Balance Reexpresado"

    headers = ["Cuenta / Descripción", "Saldo Inicial", "Monto Debe", "Monto Haber", "Saldo Actual"]
    ws.append(headers)

    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center")
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = align_center

    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )

    for _, row in df.iterrows():
        ws.append(list(row))
        current_row = ws.max_row
        for col_idx in range(1, len(row) + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = Font(name="Arial", size=10)
            cell.border = thin_border
            if col_idx > 1:
                cell.number_format = '#,##0.00;(#,##0.00);"-"'
                cell.alignment = Alignment(horizontal="right")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.markdown("---")
excel_bytes = generar_excel_profesional(df_reexpresado)

st.download_button(
    label="📥 Descargar Excel Oficial para Reimpresión (.xlsx)",
    data=excel_bytes,
    file_name="Balance_Reexpresado_Oficial.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
