import streamlit as st
import imaplib
import email
from email.header import decode_header
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import io

# --- CONFIGURACIÓN ---
# Estos valores se leen desde la pestaña "Secrets" de Streamlit
EMAIL_USER = st.secrets["EMAIL"]
EMAIL_PASS = st.secrets["PASSWORD"]

def procesar_xml(xml_content):
    root = ET.fromstring(xml_content)
    ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
    
    # Extraer datos (ejemplo simplificado)
    datos = {
        "RFC": root.find('.//cfdi:Emisor', ns).get('Rfc'),
        "Nombre": root.find('.//cfdi:Emisor', ns).get('Nombre'),
        "Total": root.find('.//cfdi:Comprobante', ns).get('Total')
    }
    return datos

# --- INTERFAZ ---
st.title("🏛️ SISTEMA DIF")

if st.button("📥 Procesar Facturas de Gmail"):
    try:
        # Conexión a Gmail
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")
        
        # Buscar correos (ejemplo: buscar todos)
        status, messages = mail.search(None, 'ALL')
        
        datos_procesados = []
        for num in messages[0].split():
            res, msg = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(msg[0][1])
            
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart': continue
                if part.get_filename() and part.get_filename().endswith('.xml'):
                    xml_data = part.get_payload(decode=True)
                    datos = procesar_xml(xml_data)
                    datos_procesados.append(datos)
        
        mail.close()
        mail.logout()
        
        # --- GENERACIÓN DE EXCEL EN MEMORIA ---
        wb = load_workbook("requisicion.xlsx") # Asegúrate de tener este archivo en tu repo
        ws = wb.active
        
        # Escribir los datos en celdas (ejemplo: fila 5, columna B)
        for i, d in enumerate(datos_procesados):
            ws[f'B{5+i}'] = d['Nombre']
            ws[f'C{5+i}'] = d['Total']
        
        # Guardar en buffer
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.success("¡Procesamiento exitoso!")
        
        # Botón de descarga para el usuario
        st.download_button(
            label="💾 Descargar Excel con Facturas",
            data=output,
            file_name="Reporte_Facturas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error: {e}")
