import streamlit as st
import imaplib
import email
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import io
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURACIÓN ---
EMAIL_USER = st.secrets["EMAIL"]
EMAIL_PASS = st.secrets["PASSWORD"]

# --- FUNCIÓN PARA GUARDAR EN GOOGLE SHEETS ---
def guardar_en_sheets(datos):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # Asegúrate de que el archivo 'credenciales.json' esté en tu GitHub
    creds = ServiceAccountCredentials.from_json_keyfile_name('credenciales.json', scope)
    client = gspread.authorize(creds)
    
    # ID de tu hoja (el código largo de la URL)
    hoja = client.open_by_key('1rwUk0h9Yx8BA8jmVHHHtS6Etj7HqOfHsSzHrpVanqro')
    worksheet = hoja.sheet1
    
    # Escribir fila
    worksheet.append_row([datos['RFC'], datos['Nombre'], datos['Total']])

# --- PROCESAMIENTO XML ---
def procesar_xml(xml_content):
    root = ET.fromstring(xml_content)
    ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
    return {
        "RFC": root.find('.//cfdi:Emisor', ns).get('Rfc'),
        "Nombre": root.find('.//cfdi:Emisor', ns).get('Nombre'),
        "Total": root.find('.//cfdi:Comprobante', ns).get('Total')
    }

# --- INTERFAZ ---
st.title("🏛️ SISTEMA DIF")

if st.button("📥 Procesar Facturas y Enviar a Drive"):
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")
        
        status, messages = mail.search(None, 'ALL')
        for num in messages[0].split():
            res, msg = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(msg[0][1])
            for part in msg.walk():
                if part.get_filename() and part.get_filename().endswith('.xml'):
                    xml_data = part.get_payload(decode=True)
                    datos = procesar_xml(xml_data)
                    # Guardar en Sheets
                    guardar_en_sheets(datos)
        
        mail.close()
        mail.logout()
        st.success("¡Información enviada con éxito a tu hoja de cálculo!")
    except Exception as e:
        st.error(f"Error: {e}")
