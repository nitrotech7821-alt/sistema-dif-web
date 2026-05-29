import os
import re
import io
import email
import shutil
import zipfile
import imaplib
import xml.etree.ElementTree as ET
from datetime import datetime
from email.header import decode_header

import streamlit as st
from openpyxl import load_workbook

# =====================================================
# CONFIGURACIÓN
# =====================================================
EMAIL_USER = "difhermosillomunicipal@gmail.com"
EMAIL_PASS = "qcec ftus qhbw ckdi"

CARPETA_BASE = "datos_sistema"
CARPETA_FACTURAS = os.path.join(CARPETA_BASE, "facturas")
CARPETA_GENERADOS = os.path.join(CARPETA_BASE, "generados")
CARPETA_FORMATOS = "formatos"

TEMPLATE_REQUISICION = os.path.join(CARPETA_FORMATOS, "requisicion.xlsx")
TEMPLATE_COTIZACION = os.path.join(CARPETA_FORMATOS, "cotizacion.xlsx")
ARCHIVO_FOLIO = os.path.join(CARPETA_BASE, "folio.txt")

RFC_CALOTE = "AALK801205TH8"

os.makedirs(CARPETA_FACTURAS, exist_ok=True)
os.makedirs(CARPETA_GENERADOS, exist_ok=True)
os.makedirs(CARPETA_BASE, exist_ok=True)

# =====================================================
# DISEÑO STREAMLIT
# =====================================================
st.set_page_config(
    page_title="Facturas CALOTE - DIF Hermosillo",
    page_icon="📄",
    layout="centered"
)

st.markdown("""
<style>
.stApp {
    background: linear-gradient(180deg, #f7f9fb 0%, #eef3f5 100%);
}

.block-container {
    padding-top: 25px;
}

.header-card {
    background: white;
    padding: 28px;
    border-radius: 20px;
    box-shadow: 0px 6px 18px rgba(0,0,0,0.10);
    text-align: center;
    margin-bottom: 25px;
}

.header-title {
    color: #187C7A;
    font-size: 32px;
    font-weight: 800;
    margin-top: 10px;
}

.header-subtitle {
    color: #555;
    font-size: 17px;
}

.info-card {
    background: white;
    padding: 22px;
    border-radius: 16px;
    box-shadow: 0px 4px 14px rgba(0,0,0,0.08);
    margin-bottom: 22px;
    border-left: 7px solid #E87522;
}

.result-card {
    background: white;
    padding: 22px;
    border-radius: 16px;
    box-shadow: 0px 4px 14px rgba(0,0,0,0.08);
    margin-top: 18px;
    margin-bottom: 12px;
    border-left: 7px solid #187C7A;
}

.stButton > button {
    background: linear-gradient(90deg, #E94E1B, #F2B233);
    color: white;
    border: none;
    border-radius: 14px;
    padding: 14px 20px;
    font-size: 18px;
    font-weight: 800;
    width: 100%;
    box-shadow: 0px 4px 12px rgba(0,0,0,0.18);
}

.stButton > button:hover {
    background: linear-gradient(90deg, #D94316, #E6A222);
    transform: scale(1.01);
}

.stDownloadButton > button {
    background: #187C7A;
    color: white;
    border: none;
    border-radius: 12px;
    padding: 12px 18px;
    font-size: 16px;
    font-weight: 700;
    width: 100%;
}

.stDownloadButton > button:hover {
    background: #0F6664;
}

.footer {
    text-align: center;
    color: #777;
    font-size: 13px;
    margin-top: 35px;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# FUNCIONES
# =====================================================
def limpiar_nombre(nombre):
    return re.sub(r'[\\/*?:"<>|]', "_", nombre)

def money(valor):
    try:
        return float(valor)
    except:
        return 0.0

def escribir(ws, celda, valor):
    for rango in ws.merged_cells.ranges:
        if celda in rango:
            ws.cell(rango.min_row, rango.min_col).value = valor
            return
    ws[celda] = valor

def obtener_folio():
    if not os.path.exists(ARCHIVO_FOLIO):
        with open(ARCHIVO_FOLIO, "w") as f:
            f.write("1721")

    with open(ARCHIVO_FOLIO, "r") as f:
        folio = int(f.read().strip())

    with open(ARCHIVO_FOLIO, "w") as f:
        f.write(str(folio + 1))

    return folio

def decodificar(texto):
    if not texto:
        return ""

    partes = decode_header(texto)
    resultado = ""

    for parte, codificacion in partes:
        if isinstance(parte, bytes):
            resultado += parte.decode(codificacion or "utf-8", errors="ignore")
        else:
            resultado += parte

    return resultado

# =====================================================
# LEER XML CFDI
# =====================================================
def leer_xml_cfdi(ruta_xml):
    ns = {
        "cfdi": "http://www.sat.gob.mx/cfd/4",
        "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital"
    }

    tree = ET.parse(ruta_xml)
    root = tree.getroot()

    emisor = root.find("cfdi:Emisor", ns)
    receptor = root.find("cfdi:Receptor", ns)
    conceptos = root.find("cfdi:Conceptos", ns)
    timbre = root.find(".//tfd:TimbreFiscalDigital", ns)

    datos = {
        "uuid": timbre.attrib.get("UUID", "") if timbre is not None else "",
        "serie": root.attrib.get("Serie", ""),
        "folio_factura": root.attrib.get("Folio", ""),
        "fecha": root.attrib.get("Fecha", ""),
        "subtotal": money(root.attrib.get("SubTotal", 0)),
        "total": money(root.attrib.get("Total", 0)),
        "moneda": root.attrib.get("Moneda", "MXN"),
        "proveedor": emisor.attrib.get("Nombre", "") if emisor is not None else "",
        "rfc_proveedor": emisor.attrib.get("Rfc", "") if emisor is not None else "",
        "receptor": receptor.attrib.get("Nombre", "") if receptor is not None else "",
        "rfc_receptor": receptor.attrib.get("Rfc", "") if receptor is not None else "",
        "conceptos": []
    }

    datos["iva"] = datos["total"] - datos["subtotal"]

    if conceptos is not None:
        for c in conceptos.findall("cfdi:Concepto", ns):
            datos["conceptos"].append({
                "cantidad": money(c.attrib.get("Cantidad", 0)),
                "unidad": c.attrib.get("Unidad", c.attrib.get("ClaveUnidad", "Pieza")),
                "codigo": c.attrib.get("NoIdentificacion", c.attrib.get("ClaveProdServ", "")),
                "descripcion": c.attrib.get("Descripcion", ""),
                "precio": money(c.attrib.get("ValorUnitario", 0)),
                "importe": money(c.attrib.get("Importe", 0))
            })

    return datos

# =====================================================
# REQUISICIÓN
# =====================================================
def llenar_requisicion(datos, folio, carpeta_salida):
    wb = load_workbook(TEMPLATE_REQUISICION)
    ws = wb.active

    fecha = datetime.now()

    escribir(ws, "J6", folio)
    escribir(ws, "J10", fecha.day)
    escribir(ws, "K10", fecha.month)
    escribir(ws, "M10", fecha.year)

    escribir(ws, "D9", "ADMINISTRATIVA")
    escribir(ws, "D10", "GENERALES")
    escribir(ws, "D11", "INMEDIATA")

    fila = 15

    for concepto in datos["conceptos"]:
        escribir(ws, f"B{fila}", concepto["cantidad"])
        escribir(ws, f"D{fila}", concepto["unidad"])
        escribir(ws, f"E{fila}", concepto["descripcion"])
        escribir(ws, f"H{fila}", f"FACTURA {datos['folio_factura']} / {datos['proveedor']}")
        fila += 1

    escribir(
        ws,
        "B51",
        f"COMPRA SEGÚN FACTURA {datos['folio_factura']} DE {datos['proveedor']} TOTAL ${datos['total']:,.2f}"
    )

    salida = os.path.join(carpeta_salida, f"REQUISICION_{folio}.xlsx")
    wb.save(salida)
    return salida

# =====================================================
# COTIZACIÓN
# =====================================================
def llenar_cotizacion(datos, folio, carpeta_salida):
    wb = load_workbook(TEMPLATE_COTIZACION)
    ws = wb.active

    fecha = datetime.now().strftime("%d/%m/%Y")
    escribir(ws, "I1", f"Cotización #{folio} {fecha}")

    fila = 14
    total_articulos = 0

    for concepto in datos["conceptos"]:
        escribir(ws, f"A{fila}", concepto["cantidad"])
        escribir(ws, f"B{fila}", concepto["codigo"])
        escribir(ws, f"C{fila}", concepto["descripcion"])
        escribir(ws, f"F{fila}", concepto["precio"])
        escribir(ws, f"G{fila}", concepto["importe"])
        escribir(ws, f"I{fila}", concepto["importe"])

        total_articulos += concepto["cantidad"]
        fila += 1

    escribir(ws, "A35", f"Total de artículos:\n\n{int(total_articulos)}")
    escribir(ws, "I35", datos["subtotal"])
    escribir(ws, "I36", datos["iva"])
    escribir(ws, "I37", datos["total"])
    escribir(ws, "A37", f"CANTIDAD CON LETRA: TOTAL ${datos['total']:,.2f} M.N.")

    salida = os.path.join(carpeta_salida, f"COTIZACION_{folio}.xlsx")
    wb.save(salida)
    return salida

# =====================================================
# DESCARGAR FACTURAS
# =====================================================
def descargar_facturas_calote():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    status, data = mail.search(None, '(OR TEXT "CALOTE" TEXT "KARLA GUADALUPE AMAYA")')

    if status != "OK":
        mail.logout()
        return []

    ids = data[0].split()
    resultados = []

    for num in ids:
        status, msg_data = mail.fetch(num, "(RFC822)")

        if status != "OK":
            continue

        msg = email.message_from_bytes(msg_data[0][1])
        asunto = decodificar(msg.get("Subject"))
        fecha_correo = datetime.now().strftime("%Y%m%d_%H%M%S")

        carpeta_correo = os.path.join(
            CARPETA_FACTURAS,
            limpiar_nombre(f"{fecha_correo}_{asunto[:40]}")
        )

        os.makedirs(carpeta_correo, exist_ok=True)

        archivos_xml = []
        archivos_pdf = []

        for parte in msg.walk():
            if parte.get_content_disposition() == "attachment":
                nombre = decodificar(parte.get_filename())

                if not nombre:
                    continue

                nombre_limpio = limpiar_nombre(nombre)
                ruta = os.path.join(carpeta_correo, nombre_limpio)

                with open(ruta, "wb") as f:
                    f.write(parte.get_payload(decode=True))

                if nombre_limpio.lower().endswith(".xml"):
                    archivos_xml.append(ruta)

                if nombre_limpio.lower().endswith(".pdf"):
                    archivos_pdf.append(ruta)

        for xml in archivos_xml:
            try:
                datos = leer_xml_cfdi(xml)

                if datos["rfc_proveedor"] != RFC_CALOTE:
                    continue

                folio = obtener_folio()

                carpeta_salida = os.path.join(
                    CARPETA_GENERADOS,
                    f"FOLIO_{folio}_{datos['folio_factura']}"
                )

                os.makedirs(carpeta_salida, exist_ok=True)

                xml_copiado = os.path.join(carpeta_salida, os.path.basename(xml))
                shutil.copy(xml, xml_copiado)

                pdf_copiado = None

                for pdf in archivos_pdf:
                    pdf_copiado = os.path.join(carpeta_salida, os.path.basename(pdf))
                    shutil.copy(pdf, pdf_copiado)

                req = llenar_requisicion(datos, folio, carpeta_salida)
                cot = llenar_cotizacion(datos, folio, carpeta_salida)

                resultados.append({
                    "folio": folio,
                    "factura": datos["folio_factura"],
                    "proveedor": datos["proveedor"],
                    "total": datos["total"],
                    "requisicion": req,
                    "cotizacion": cot,
                    "xml": xml_copiado,
                    "pdf": pdf_copiado
                })

            except Exception as e:
                resultados.append({
                    "error": str(e),
                    "xml": xml
                })

    mail.logout()
    return resultados

# =====================================================
# CREAR ZIP
# =====================================================
def crear_zip_expediente(r):
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(r["requisicion"], os.path.basename(r["requisicion"]))
        zipf.write(r["cotizacion"], os.path.basename(r["cotizacion"]))
        zipf.write(r["xml"], os.path.basename(r["xml"]))

        if r["pdf"] and os.path.exists(r["pdf"]):
            zipf.write(r["pdf"], os.path.basename(r["pdf"]))

    zip_buffer.seek(0)
    return zip_buffer

# =====================================================
# PANTALLA PRINCIPAL
# =====================================================
st.markdown('<div class="header-card">', unsafe_allow_html=True)

if os.path.exists("logo_dif.jpeg"):
    st.image("logo_dif.jpeg", width=260)
else:
    st.warning("No se encontró el logo. Guarda la imagen como logo_dif.jpeg")

st.markdown("""
<div class="header-title">Sistema de Facturas CALOTE</div>
<div class="header-subtitle">
DIF Hermosillo | Generación automática de requisiciones y cotizaciones
</div>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-card">
<b>📌 Función del sistema</b><br><br>
Este sistema descarga facturas de CALOTE desde el correo institucional,
lee el XML, genera automáticamente la requisición, cotización, XML y PDF,
y entrega todo en un expediente ZIP.
</div>
""", unsafe_allow_html=True)

if st.button("📥 Descargar facturas y generar documentos"):
    with st.spinner("Procesando correos y generando documentos..."):
        try:
            resultados = descargar_facturas_calote()
        except Exception as e:
            st.error(f"Error: {e}")
            resultados = []

    if not resultados:
        st.warning("No se encontraron facturas válidas.")
    else:
        st.success("Proceso terminado correctamente.")

        for r in resultados:
            if "error" in r:
                st.error(f"Error XML: {r['xml']} - {r['error']}")
                continue

            st.markdown(f"""
            <div class="result-card">
                <h3 style="color:#187C7A;">📄 Folio generado: {r['folio']}</h3>
                <p><b>Factura:</b> {r['factura']}</p>
                <p><b>Proveedor:</b> {r['proveedor']}</p>
                <p><b>Total:</b> ${r['total']:,.2f}</p>
            </div>
            """, unsafe_allow_html=True)

            zip_buffer = crear_zip_expediente(r)

            st.download_button(
                label="📦 Descargar expediente completo ZIP",
                data=zip_buffer,
                file_name=f"EXPEDIENTE_{r['folio']}.zip",
                mime="application/zip",
                key=f"zip_{r['folio']}"
            )

st.markdown("""
<div class="footer">
Sistema interno DIF Hermosillo · Facturas CALOTE
</div>
""", unsafe_allow_html=True)
