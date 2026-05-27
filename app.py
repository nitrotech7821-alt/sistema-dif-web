import imaplib
import email
from email.header import decode_header
import xml.etree.ElementTree as ET
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import io

# --- CONFIGURACIÓN ---
# Usamos secretos de Streamlit para seguridad
EMAIL = st.secrets["EMAIL"]
PASSWORD = st.secrets["PASSWORD"]

# --- FUNCIONES ---
def limpiar_nombre(nombre):
    import re
    return re.sub(r'[<>:"/\\|?*]', '', nombre).strip()

def escribir_en_celda(ws, coord, valor):
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
    cell.value = valor

def leer_xml(xml_content):
    root = ET.fromstring(xml_content)
    ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
    emisor = root.find('cfdi:Emisor', ns)
    conceptos_node = root.find('cfdi:Conceptos', ns)
    lista = []
    if conceptos_node is not None:
        for c in conceptos_node.findall('cfdi:Concepto', ns):
            lista.append({
                "Descripcion": c.attrib.get("Descripcion", ""),
                "Cantidad": float(c.attrib.get("Cantidad", 0)),
                "Importe": float(c.attrib.get("Importe", 0)),
                "ValorUnitario": float(c.attrib.get("ValorUnitario", 0))
            })
    return {
        "Fecha": root.attrib.get("Fecha", ""),
        "Total": float(root.attrib.get("Total", 0)),
        "RFC": emisor.attrib.get("Rfc", "") if emisor is not None else "",
        "Emisor": emisor.attrib.get("Nombre", "") if emisor is not None else "",
        "Conceptos": lista
    }

def generar_excel_memoria(template_path, datos, folio, tipo):
    wb = load_workbook(template_path)
    ws = wb.active
    # Nota: Asegúrate de tener los archivos .xlsx en tu repositorio de GitHub
    if tipo == "requisicion":
        escribir_en_celda(ws, "L6", folio)
        escribir_en_celda(ws, "D5", datos.get("Emisor", ""))
        # ... (agrega aquí el resto de tu lógica de llenado)
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ ---
st.title("🏛 SISTEMA DIF")

if st.button("📩 Procesar Facturas de Gmail"):
    try:
        with st.spinner("Conectando a Gmail..."):
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(EMAIL, PASSWORD)
            mail.select("inbox")
            _, mensajes = mail.search(None, "ALL")
            ids = mensajes[0].split()
            
            for num in ids[-5:]: # Procesamos los últimos 5 para evitar bloqueos
                _, data = mail.fetch(num, "(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                
                for part in msg.walk():
                    if part.get_content_type() == "application/xml":
                        xml_data = part.get_payload(decode=True)
                        datos = leer_xml(xml_data)
                        st.success(f"Procesado: {datos['Emisor']}")
                        st.json(datos)
            
            mail.logout()
    except Exception as e:
        st.error(f"Error: {e}")
