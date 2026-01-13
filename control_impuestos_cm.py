import pandas as pd
import glob
import pdfplumber
import re
from datetime import datetime
from openpyxl.styles import numbers
import os

# --- version 01.2026 ---

# --- Configuración ---
CUIT_BUSCADO = "30505112578"
salida_excel = r"Z:\IMPUESTOS\Control_impuestos\control_retper_activo\retenciones_percepciones_08.2025.xlsx"
nombre_hoja_alicuotas = "Alícuotas 08.2025"
nombre_hoja_retper = "Retenciones y percepciones"

# Archivos de padrones
padrones = {
    "CABA": r"Z:\IMPUESTOS\Control_impuestos\control_retper_activo\padrones\ARDJU008082025.txt",
    "BA_Percepciones": r"Z:\IMPUESTOS\Control_impuestos\control_retper_activo\padrones\PadronRGSPer082025.txt",
    "BA_Retenciones": r"Z:\IMPUESTOS\Control_impuestos\control_retper_activo\padrones\PadronRGSRet082025.txt",
    "ER": r"Z:\IMPUESTOS\Control_impuestos\control_retper_activo\padrones\PadronRetPer202508"
}

pdf_retenciones = r"Z:\IMPUESTOS\Control_impuestos\control_retper_activo\ret_per\ret08.pdf"
pdf_percepciones = r"Z:\IMPUESTOS\Control_impuestos\control_retper_activo\ret_per\perc08.pdf"

# --- Funciones de parseo padrones ---
def leer_padron_caba(ruta):
    registros = []
    with open(ruta, "r", encoding="latin1") as f:
        for linea in f:
            partes = linea.strip().split(";")
            if len(partes) >= 11 and partes[3] == CUIT_BUSCADO:
                registros.append({
                    "Jurisdicción": "CABA",
                    "Tipo": "Percepción",
                    "Fecha Desde": partes[1],
                    "Fecha Hasta": partes[2],
                    "CUIT": partes[3],
                    "Alícuota": partes[7]
                })
                registros.append({
                    "Jurisdicción": "CABA",
                    "Tipo": "Retención",
                    "Fecha Desde": partes[1],
                    "Fecha Hasta": partes[2],
                    "CUIT": partes[3],
                    "Alícuota": partes[8]
                })
    return registros

def leer_padron_ba(ruta, tipo):
    registros = []
    with open(ruta, "r", encoding="latin1") as f:
        for linea in f:
            partes = linea.strip().split(";")
            if len(partes) >= 9 and partes[4] == CUIT_BUSCADO:
                try:
                    idx_n = max(i for i, val in enumerate(partes) if val == "N")
                    alicuota = partes[idx_n + 1]
                except Exception:
                    alicuota = ""
                registros.append({
                    "Jurisdicción": "Buenos Aires",
                    "Tipo": tipo,
                    "Fecha Desde": partes[2],
                    "Fecha Hasta": partes[3],
                    "CUIT": partes[4],
                    "Alícuota": alicuota
                })
    return registros

def leer_padron_entrerios(file_base):
    archivos_posibles = glob.glob(file_base + ".xls") + glob.glob(file_base + ".xlsx")
    if not archivos_posibles:
        raise FileNotFoundError(f"No se encontró el padrón de Entre Ríos en {file_base}")
    registros = []
    df = pd.read_excel(archivos_posibles[0], header=None, dtype=str)
    for _, row in df.iterrows():
        if row[3] == CUIT_BUSCADO:
            registros.append({
                "Jurisdicción": "Entre Ríos",
                "Tipo": "Percepción",
                "Fecha Desde": row[1],
                "Fecha Hasta": row[2],
                "CUIT": row[3],
                "Alícuota": row[7]
            })
            registros.append({
                "Jurisdicción": "Entre Ríos",
                "Tipo": "Retención",
                "Fecha Desde": row[1],
                "Fecha Hasta": row[2],
                "CUIT": row[3],
                "Alícuota": row[8]
            })
    return registros

# --- Construir tabla de alícuotas ---
registros_totales = []
registros_totales.extend(leer_padron_caba(padrones["CABA"]))
registros_totales.extend(leer_padron_ba(padrones["BA_Percepciones"], "Percepción"))
registros_totales.extend(leer_padron_ba(padrones["BA_Retenciones"], "Retención"))
registros_totales.extend(leer_padron_entrerios(padrones["ER"]))

# Agregar Santa Fe manualmente
registros_totales.append({
    "Jurisdicción": "Santa Fe",
    "Tipo": "Retención",
    "Fecha Desde": "Res. 029/22",
    "Fecha Hasta": "",
    "CUIT": "",
    "Alícuota": "0,60"
})
registros_totales.append({
    "Jurisdicción": "Santa Fe",
    "Tipo": "Percepción",
    "Fecha Desde": "Res. 13/2024",
    "Fecha Hasta": "",
    "CUIT": "",
    "Alícuota": "2,50"
})

df_padrones = pd.DataFrame(registros_totales)

# --- Función general para leer PDF (retenciones o percepciones) ---
def leer_pdf(ruta_pdf, tipo_mov):
    datos = []
    proveedor = None
    cuit_prov = None
    jurisdiccion = None
    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for linea in text.split("\n"):
                if linea.startswith("Régimen:"):
                    # Jurisdicción al inicio del bloque
                    if "CABA" in linea.upper():
                        jurisdiccion = "CABA"
                    elif "BS. AS" in linea.upper() or "BUENOS AIRES" in linea.upper():
                        jurisdiccion = "Buenos Aires"
                    elif "ER" in linea.upper() or "ENTRE RÍOS" in linea.upper():
                        jurisdiccion = "Entre Ríos"
                    elif "SFE" in linea.upper() or "SANTA FE" in linea.upper():
                        jurisdiccion = "Santa Fe"
                    else:
                        jurisdiccion = "Otra"
                    continue
                match_prov = re.match(r"\d+\s+(.+?)\s+(\d{2}-\d{8}-\d)", linea)
                if match_prov:
                    proveedor = match_prov.group(1).strip()
                    cuit_prov = match_prov.group(2).replace("-", "")
                    continue
                match_mov = re.match(r"(\d{2}/\d{2}/\d{2}).+?([\d\.\,]+)\s+([\d\.\,]+)$", linea)
                if match_mov and proveedor and jurisdiccion:
                    fecha = datetime.strptime(match_mov.group(1), "%d/%m/%y")
                    base = float(match_mov.group(2).replace(".", "").replace(",", "."))
                    monto = float(match_mov.group(3).replace(".", "").replace(",", "."))
                    datos.append({
                        "Fecha": fecha,
                        "Proveedor": proveedor,
                        "CUIT Proveedor": cuit_prov,
                        "Jurisdicción": jurisdiccion,
                        "Base Imponible": base,
                        "Retención" if tipo_mov=="Retención" else "Percepción": monto,
                        "Tipo": tipo_mov
                    })
    return datos

# --- Leer PDFs ---
df_retenciones = pd.DataFrame(leer_pdf(pdf_retenciones, "Retención"))
df_percepciones = pd.DataFrame(leer_pdf(pdf_percepciones, "Percepción"))

# Para percepciones, renombrar la columna a "Retención" para merge uniforme
if "Percepción" in df_percepciones.columns:
    df_percepciones = df_percepciones.rename(columns={"Percepción": "Retención"})
df_retper = pd.concat([df_retenciones, df_percepciones], ignore_index=True)
df_retper["Alícuota aplicada"] = (df_retper["Retención"] / df_retper["Base Imponible"]) * 100

# --- Merge con padrones ---
df_padrones["Alícuota Padrón"] = pd.to_numeric(df_padrones["Alícuota"].astype(str).str.replace(",", "."), errors="coerce")
df_merge = pd.merge(
    df_retper,
    df_padrones,
    left_on=["Jurisdicción", "Tipo"],
    right_on=["Jurisdicción", "Tipo"],
    how="left",
    suffixes=("", "_Padrón")
)

def comparar(row):
    if pd.isna(row["Alícuota Padrón"]):
        return "a completar"
    elif abs(row["Alícuota aplicada"] - row["Alícuota Padrón"]) < 0.1:
        return "OK"
    else:
        return f"Dif: {row['Alícuota aplicada'] - row['Alícuota Padrón']:.2f}"

df_merge["Control"] = df_merge.apply(comparar, axis=1)

# --- Exportar a Excel ---
with pd.ExcelWriter(salida_excel, engine="openpyxl") as writer:
    df_padrones.to_excel(writer, sheet_name=nombre_hoja_alicuotas, index=False)
    df_merge.to_excel(writer, sheet_name=nombre_hoja_retper, index=False)

    # Formato numérico 2 decimales en Retenciones y Percepciones
    ws = writer.sheets[nombre_hoja_retper]
    col_mapping = {col: idx+1 for idx, col in enumerate(df_merge.columns)}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[col_mapping["Base Imponible"] - 1].number_format = '#,##0.00'
        row[col_mapping["Retención"] - 1].number_format = '#,##0.00'
        row[col_mapping["Alícuota aplicada"] - 1].number_format = '#,##0.00'

print(f"✅ Archivo generado: {salida_excel}")
