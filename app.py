import streamlit as st
import imaplib
import email
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET

st.title("🤖 Robot Administrativo DIF")

# Función para manejar celdas combinadas
def escribir_en_celda(ws, celda, valor):
    for range_ in ws.merged_cells.ranges:
        if celda in range_:
            ws.unmerge_cells(str(range_))
            break
    ws[celda] = valor

def extraer_datos_xml(xml_bytes):
    try:
        root = ET.fromstring(xml_bytes)
        # Ajusta el namespace según tu XML (usualmente cfdi:v4.0)
        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
        folio = root.find('.//cfdi:Comprobante', ns).get('Folio', 'SN')
        return folio
    except:
        return "ERROR_FOLIO"

if st.button("🚀 PROCESAR FACTURAS DEL CORREO"):
    try:
        # 1. Conexión
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"])
        mail.select("inbox")
        
        # 2. Buscar correos
        _, mensajes = mail.search(None, 'UNSEEN') # Busca correos no leídos
        
        for num in mensajes[0].split():
            _, data = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(data[0][1])
            
            for part in msg.walk():
                if part.get_filename() and part.get_filename().endswith(".xml"):
                    xml_data = part.get_payload(decode=True)
                    folio = extraer_datos_xml(xml_data)
                    
                    # 3. Procesar archivos
                    wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
                    escribir_en_celda(wb_cot.active, 'H2', folio)
                    
                    wb_req = openpyxl.load_workbook("requisicion.xlsx")
                    escribir_en_celda(wb_req.active, 'I7', folio)
                    
                    # 4. Crear ZIP en memoria
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w') as zf:
                        zf.writestr(f"Factura_{folio}.xml", xml_data)
                        
                        buf_cot = io.BytesIO()
                        wb_cot.save(buf_cot)
                        zf.writestr(f"Cotizacion_{folio}.xlsx", buf_cot.getvalue())
                        
                        buf_req = io.BytesIO()
                        wb_req.save(buf_req)
                        zf.writestr(f"Requisicion_{folio}.xlsx", buf_req.getvalue())
                    
                    # 5. Botón de descarga individual
                    st.download_button(
                        label=f"📥 Descargar Expediente: {folio}",
                        data=zip_buf.getvalue(),
                        file_name=f"Expediente_{folio}.zip",
                        mime="application/zip"
                    )
        mail.logout()
        st.success("Proceso terminado. Revisa los botones de descarga arriba.")
    except Exception as e:
        st.error(f"Error al conectar con correo: {e}")
