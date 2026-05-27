import streamlit as st
import imaplib
import email
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET

st.title("🤖 Robot Administrativo DIF")

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

def procesar():
    # 1. Conexión
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"])
    mail.select("inbox")
    
    _, mensajes = mail.search(None, '(FROM "facturacion")') 
    ids = mensajes[0].split()[-1:] # Tomamos solo la última factura para probar
    
    for num in ids:
        _, data = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(data[0][1])
        
        for part in msg.walk():
            if part.get_filename() and part.get_filename().endswith(".xml"):
                xml_data = part.get_payload(decode=True)
                datos = extraer_datos_xml(xml_data)
                
                # 2. Llenar Excel con las celdas correctas según tus archivos
                wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
                ws_cot = wb_cot.active
                ws_cot['H2'] = datos['folio'] # Ajuste basado en tu archivo
                
                wb_req = openpyxl.load_workbook("requisicion.xlsx")
                ws_req = wb_req.active
                ws_req['I9'] = datos['folio'] # Ajuste basado en tu archivo
                
                # 3. Guardar en memoria
                buf_cot = io.BytesIO()
                wb_cot.save(buf_cot)
                
                buf_req = io.BytesIO()
                wb_req.save(buf_req)
                
                # 4. Crear ZIP
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w') as zf:
                    zf.writestr(f"Factura_{datos['folio']}.xml", xml_data)
                    zf.writestr(f"Cotizacion_{datos['folio']}.xlsx", buf_cot.getvalue())
                    zf.writestr(f"Requisicion_{datos['folio']}.xlsx", buf_req.getvalue())
                
                # 5. Botón de descarga
                st.download_button(
                    label=f"📥 DESCARGAR: Factura {datos['folio']}",
                    data=zip_buf.getvalue(),
                    file_name=f"Expediente_{datos['folio']}.zip",
                    mime="application/zip"
                )
    mail.logout()

if st.button("Buscar Facturas y Preparar Descarga"):
    procesar()
