import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from copy import copy
from collections import defaultdict
from openpyxl import load_workbook

# ============================================================
# 1. CONFIGURACIÓN Y UTILIDADES FINANCIERAS
# ============================================================
st.set_page_config(page_title="Consolidador Confidelis", layout="wide")

MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}
MESES_INV = {v: k for k, v in MESES.items()}

def normalizar(t):
    return str(t).upper().strip() if t else ""

def extraer_numeros(texto):
    return [float(n.replace(",", "")) for n in re.findall(r"[\d,]+\.\d+", texto)]

def extraer_todos_numeros(texto):
    return [float(n.replace(",", "")) for n in re.findall(r"[\d,]+\.?\d*", texto) if n]

# ============================================================
# 2. MOTOR DE EXTRACCIÓN (Basado en extractor_gbm.py)
# ============================================================
def detectar_plataforma(texto):
    return "Prestadero" if "PRESTADERO" in texto.upper() else "GBM"

def extraer_nombre_cliente(pdf, plataforma):
    texto = pdf.pages[0].extract_text() or ""
    lineas = texto.split("\n")
    if plataforma == "Prestadero":
        for l in lineas:
            if "Periodo:" in l and "Estado de Cuenta" not in l:
                return re.split(r"\s+Periodo:", l)[0].strip().upper()
    else:
        for l in lineas:
            if "Contrato:" in l:
                p = re.split(r"\s+Contrato:", l)[0].strip()
                return p.replace("PUBLICO EN GENERAL - ", "").upper()
    return "DESCONOCIDO"

def extraer_portafolio_gbm(pdf):
    portafolio = []
    en_acciones = False
    for pag in pdf.pages:
        texto = pag.extract_text() or ""
        for l in texto.split("\n"):
            lu = l.upper()
            if "ACCIONES" in lu: en_acciones = True; continue
            if en_acciones and (lu.startswith("TOTAL") or "RENDIMIENTO" in lu): 
                en_acciones = False; continue
            if not en_acciones: continue
            
            m = re.match(r"^([A-Z]+(?:\s+\d+)?)\s+", l.strip())
            if m:
                nums = extraer_todos_numeros(l[m.end():])
                if len(nums) >= 8:
                    portafolio.append({"emisora": m.group(1).strip(), "valor": nums[7]})
    return portafolio

# ============================================================
# 3. LÓGICA DE CONSOLIDACIÓN (Sumas, Restas e Inserción)
# ============================================================
def actualizar_hoja_maestra(ws, datos_pdf):
    # Localizar filas clave
    fila_header = 23
    fila_totales = None
    for r in range(1, 100):
        val = str(ws.cell(r, 1).value).upper()
        if "INSTRUMENTO" in val: fila_header = r
        if "TOTALES" in val: 
            fila_totales = r
            break
    
    if not fila_totales: return
    
    # 1. Mapear lo que ya existe en el Excel
    existentes = {}
    for r in range(fila_header + 1, fila_totales):
        nom = normalizar(ws.cell(r, 1).value)
        if nom and nom != "-": existentes[nom] = r

    # 2. Actualizar o Identificar nuevos
    port_pdf = {normalizar(i["emisora"]): i["valor"] for i in datos_pdf.get("portafolio", [])}
    
    for emisora, valor_nuevo in port_pdf.items():
        if emisora in existentes:
            row = existentes[emisora]
            old_b = ws.cell(row, 2).value or 0
            ws.cell(row, 3).value = valor_nuevo  # C: Saldo Actual
            ws.cell(row, 5).value = valor_nuevo - old_b # E: Ganancia
            ws.cell(row, 14).value = valor_nuevo # N: Total
        else:
            # LÓGICA DE INSTRUMENTO NUEVO: Insertar antes de TOTALES
            ws.insert_rows(fila_totales)
            ws.cell(fila_totales, 1).value = emisora
            ws.cell(fila_totales, 2).value = valor_nuevo # B: Inversión inicial
            ws.cell(fila_totales, 3).value = valor_nuevo # C: Saldo
            ws.cell(fila_totales, 14).value = valor_nuevo
            ws.cell(fila_totales, 15).value = "GBM"
            fila_totales += 1 # Desplazar el marcador de totales

# ============================================================
# 4. INTERFAZ STREAMLIT
# ============================================================
def main():
    st.title("💰 Consolidador Financiero Unificado")
    st.markdown("Carga tus archivos para procesar la rentabilidad del mes.")

    with st.sidebar:
        st.header("Entrada de Archivos")
        maestro_file = st.file_uploader("Excel Maestro (.xlsx)", type="xlsx")
        pdf_files = st.file_uploader("PDFs GBM/Prestadero", type="pdf", accept_multiple_files=True)

    if st.button("🚀 Iniciar Proceso"):
        if maestro_file and pdf_files:
            try:
                # Procesar PDFs
                clientes_dict = defaultdict(lambda: {"portafolio": []})
                for f in pdf_files:
                    with pdfplumber.open(f) as pdf:
                        plat = detectar_plataforma(pdf.pages[0].extract_text() or "")
                        nombre = extraer_nombre_cliente(pdf, plat)
                        if plat == "GBM":
                            clientes_dict[nombre]["portafolio"].extend(extraer_portafolio_gbm(pdf))

                # Procesar Excel
                wb = load_workbook(maestro_file)
                for nombre, data in clientes_dict.items():
                    # Buscar hoja por nombre de cliente
                    sheet_name = next((s for s in wb.sheetnames if nombre in s.upper() or s.upper() in nombre), None)
                    if sheet_name:
                        actualizar_hoja_maestra(wb[sheet_name], data)
                        st.success(f"✅ Hoja '{sheet_name}' actualizada.")
                
                # Descarga
                out = io.BytesIO()
                wb.save(out)
                st.download_button(
                    label="📥 Descargar Reporte Consolidado",
                    data=out.getvalue(),
                    file_name="Consolidado_Final.xlsx",
                    mime="application/octet-stream"
                )
            except Exception as e:
                st.error(f"Error técnico: {str(e)}")
        else:
            st.warning("Por favor carga el Maestro y los PDFs.")

if __name__ == "__main__":
    main()
