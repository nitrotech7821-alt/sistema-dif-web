import streamlit as st
import imaplib
import email
import xml.etree.ElementTree as ET
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURACIÓN ---
# Estos datos ahora se leen desde la sección Secrets de tu app
EMAIL_USER = st.secrets["EMAIL"]
EMAIL_PASS = st.secrets["PASSWORD"]

# --- FUNCIÓN PARA GUARDAR EN GOOGLE SHEETS ---
def guardar_en_sheets(datos):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    # Ahora leemos el diccionario [gcp] directamente desde los Secrets
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp"], scope)
    client = gspread.authorize(creds)
    
    # ID de tu hoja (el código largo de la URL)
    hoja = client.open_by_key('1rwUk0h9Yx8BA8jmVHHHtS6Etj7HqOfHsSzHrpVanqro')
    worksheet = hoja.sheet1
    
    # Agrega una nueva fila al final de la hoja
    worksheet.append_row([datos['RFC'], datos['Nombre'], datos['Total']])

# --- PROCESAMIENTO XML ---
def procesar_xml(xml_content):
    root = ET.fromstring(xml_content)
    # Namespace para CFDI 4.0
    ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
    
    emisor = root.find('.//cfdi:Emisor', ns)
    comprobante = root.find('.//cfdi:Comprobante', ns)
    
    return {
        "RFC": emisor.get('Rfc') if emisor is not None else "N/A",
        "Nombre": emisor.get('Nombre') if emisor is not None else "N/A",
        "Total": comprobante.get('Total') if comprobante is not None else "0"
    }

# --- INTERFAZ ---
st.title("🏛️ SISTEMA DIF")

if st.button("📥 Procesar Facturas y Enviar a Hoja de Cálculo"):
    try:
        # Conexión a Gmail
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")
        
        # Buscar todos los correos
        status, messages = mail.search(None, 'ALL')
        for num in messages[0].split():
            res, msg = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(msg[0][1])
            
            # Buscar adjuntos XML
            for part in msg.walk():
                if part.get_filename() and part.get_filename().endswith('.xml'):
                    xml_data = part.get_payload(decode=True)
                    datos = procesar_xml(xml_data)
                    guardar_en_sheets(datos)
        
        mail.close()
        mail.logout()
        st.success("✅ ¡Procesamiento finalizado! Datos guardados en Google Sheets.")
        
    except Exception as e:
        st.error(f"Error: {e}")
