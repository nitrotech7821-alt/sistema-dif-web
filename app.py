import streamlit as st
import imaplib
import email
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET

# Función para extraer datos del XML
def extraer_datos_xml(xml_bytes):
    try:
        root = ET.fromstring(xml_bytes)
        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
        emisor = root.find('.//cfdi:Emisor', ns).get('Nombre')
        total = root.find('.//cfdi:Comprobante', ns).get('Total')
        folio = root.find('.//cfdi:Comprobante', ns).get('Folio')
        return {"nombre": emisor, "total": total, "folio": folio}
    except:
        return {"nombre": "N/A", "total": "0", "folio": "N/A"}

# Función para rellenar los Excel (Basado en tus imágenes)
def llenar_archivo(plantilla_path, datos):
    wb = openpyxl.load_workbook(plantilla_path)
    ws = wb.active
    
    # Mapeo según tus imágenes:
    if "cotizacion" in plantilla_path:
        ws['B2'] = datos['folio']
        ws['B3'] = datos['nombre']
        ws['B4'] = datos['total']
    elif "requisicion" in plantilla_path:
        ws['B2'] = datos['folio']
        ws['B3'] = datos['nombre']
        ws['B4'] = datos['total']
    
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

st.title("🤖 Robot Administrativo DIF")

if st.button("Procesar Facturas del Correo"):
    # Conexión al correo
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"])
    mail.select("inbox")
    
    _, mensajes = mail.search(None, '(FROM "facturacion")') 
    
    for num in mensajes[0].split()[-3:]:
        _, data = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(data[0][1])
        
        for part in msg.walk():
            if part.get_filename() and part.get_filename().endswith(".xml"):
                xml_data = part.get_payload(decode=True)
                datos = extraer_datos_xml(xml_data)
                
                # Generar archivos
                cot_bytes = llenar_archivo("cotizacion.xlsx", datos)
                req_bytes = llenar_archivo("requisicion.xlsx", datos)
                
                # Crear ZIP (Carpeta comprimida)
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w') as zf:
                    zf.writestr(f"Factura_{datos['folio']}.xml", xml_data)
                    zf.writestr(f"Cotizacion_{datos['folio']}.xlsx", cot_bytes)
                    zf.writestr(f"Requisicion_{datos['folio']}.xlsx", req_bytes)
                
                st.download_button(
                    label=f"📥 Descargar Expediente {datos['folio']}", 
                    data=zip_buf.getvalue(), 
                    file_name=f"Expediente_{datos['folio']}.zip"
                )
    mail.logout()
