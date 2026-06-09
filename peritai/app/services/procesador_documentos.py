"""
Servicio de ingesta y procesamiento de documentos.
Maneja PDF, Word, Excel, imágenes con OCR.
"""
import os
import json
import hashlib
from pathlib import Path
from loguru import logger
import fitz  # pymupdf
from docx import Document
from openpyxl import load_workbook
from PIL import Image
import pytesseract
from app.config import settings
from app.agents.orquestador import extraer_datos_documento


TIPOS_SOPORTADOS = {
    "application/pdf": "pdf",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "word",
    "application/msword": "word",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "excel",
    "application/vnd.ms-excel": "excel",
    "image/jpeg": "imagen",
    "image/png": "imagen",
    "image/tiff": "imagen",
    "image/webp": "imagen",
}


def calcular_hash(ruta: str) -> str:
    sha256 = hashlib.sha256()
    with open(ruta, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            sha256.update(chunk)
    return sha256.hexdigest()


def extraer_texto_pdf(ruta: str) -> tuple[str, float]:
    """
    Extrae texto de un PDF. Si tiene poco texto (escaneado),
    aplica OCR página por página.
    Retorna (texto, confianza)
    """
    texto = ""
    confianza = 1.0

    try:
        doc = fitz.open(ruta)
        for pagina in doc:
            texto_pagina = pagina.get_text()
            texto += texto_pagina + "\n"
        doc.close()

        # Si hay muy poco texto, probablemente está escaneado
        palabras = len(texto.split())
        if palabras < 50:
            logger.info(f"PDF escaneado detectado en {ruta}, aplicando OCR...")
            texto, confianza = aplicar_ocr_pdf(ruta)
    except Exception as e:
        logger.error(f"Error extrayendo texto de PDF {ruta}: {e}")
        texto = ""
        confianza = 0.0

    return texto.strip(), confianza


def aplicar_ocr_pdf(ruta: str) -> tuple[str, float]:
    """Convierte páginas del PDF a imagen y aplica OCR."""
    texto_total = ""
    try:
        doc = fitz.open(ruta)
        for i, pagina in enumerate(doc):
            # Renderizar página como imagen de alta resolución
            mat = fitz.Matrix(2, 2)  # 2x zoom para mejor OCR
            pix = pagina.get_pixmap(matrix=mat)
            img_ruta = f"/tmp/peritai_ocr_p{i}.png"
            pix.save(img_ruta)

            img = Image.open(img_ruta)
            texto = pytesseract.image_to_string(img, lang="spa+eng")
            texto_total += texto + "\n"

            os.remove(img_ruta)
        doc.close()
        return texto_total.strip(), 0.75  # OCR tiene menos confianza
    except Exception as e:
        logger.error(f"Error en OCR de PDF: {e}")
        return "", 0.0


def extraer_texto_word(ruta: str) -> tuple[str, float]:
    try:
        doc = Document(ruta)
        parrafos = [p.text for p in doc.paragraphs if p.text.strip()]
        # También extraer texto de tablas
        for tabla in doc.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    if celda.text.strip():
                        parrafos.append(celda.text.strip())
        return "\n".join(parrafos), 1.0
    except Exception as e:
        logger.error(f"Error extrayendo texto de Word {ruta}: {e}")
        return "", 0.0


def extraer_texto_excel(ruta: str) -> tuple[str, float]:
    try:
        wb = load_workbook(ruta, read_only=True, data_only=True)
        lineas = []
        for nombre_hoja in wb.sheetnames:
            hoja = wb[nombre_hoja]
            lineas.append(f"=== Hoja: {nombre_hoja} ===")
            for fila in hoja.iter_rows(values_only=True):
                celdas_no_vacias = [str(c) for c in fila if c is not None]
                if celdas_no_vacias:
                    lineas.append(" | ".join(celdas_no_vacias))
        return "\n".join(lineas), 1.0
    except Exception as e:
        logger.error(f"Error extrayendo texto de Excel {ruta}: {e}")
        return "", 0.0


def extraer_texto_imagen(ruta: str) -> tuple[str, float]:
    try:
        img = Image.open(ruta)
        # Pre-procesar para mejor OCR
        img = img.convert("L")  # escala de grises
        texto = pytesseract.image_to_string(img, lang="spa+eng")
        confianza = 0.70 if len(texto.split()) > 10 else 0.40
        return texto.strip(), confianza
    except Exception as e:
        logger.error(f"Error en OCR de imagen {ruta}: {e}")
        return "", 0.0


def detectar_tipo_documento(nombre: str, texto: str) -> str:
    """
    Clasifica el documento por su nombre y contenido.
    Retorna: factura | cotizacion | foto | contrato | expediente | otro
    """
    nombre_lower = nombre.lower()
    texto_lower = texto.lower()[:500]

    if any(p in nombre_lower for p in ["factura", "invoice", "fact"]):
        return "factura"
    if any(p in nombre_lower for p in ["cotiz", "presupuesto", "quote", "proforma"]):
        return "cotizacion"
    if any(p in nombre_lower for p in ["foto", "img", "image", "photo", ".jpg", ".png", ".jpeg"]):
        return "foto"
    if any(p in nombre_lower for p in ["contrato", "contrac", "poliza", "poliz"]):
        return "contrato"

    # Detección por contenido
    if any(p in texto_lower for p in ["total a pagar", "subtotal", "iva", "factura"]):
        return "factura"
    if any(p in texto_lower for p in ["cotización", "presupuesto", "precio unitario", "validez"]):
        return "cotizacion"
    if any(p in texto_lower for p in ["contrato", "cláusula", "partes contratantes"]):
        return "contrato"

    return "otro"


async def procesar_documento(ruta_archivo: str, nombre_original: str) -> dict:
    """
    Pipeline completo de procesamiento de un documento.
    Retorna diccionario con todos los datos extraídos.
    """
    resultado = {
        "nombre_original": nombre_original,
        "ruta_storage": ruta_archivo,
        "hash": calcular_hash(ruta_archivo),
        "texto_extraido": "",
        "tipo_doc": "otro",
        "tipo_mime": "",
        "datos_extraidos": {},
        "estado": "error",
        "confianza_ocr": 0.0,
        "notas": "",
    }

    extension = Path(nombre_original).suffix.lower()

    try:
        # Extraer texto según el tipo de archivo
        if extension == ".pdf":
            resultado["tipo_mime"] = "application/pdf"
            texto, confianza = extraer_texto_pdf(ruta_archivo)
        elif extension in [".docx", ".doc"]:
            resultado["tipo_mime"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            texto, confianza = extraer_texto_word(ruta_archivo)
        elif extension in [".xlsx", ".xls"]:
            resultado["tipo_mime"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            texto, confianza = extraer_texto_excel(ruta_archivo)
        elif extension in [".jpg", ".jpeg", ".png", ".tiff", ".webp"]:
            resultado["tipo_mime"] = f"image/{extension[1:]}"
            texto, confianza = extraer_texto_imagen(ruta_archivo)
        else:
            resultado["notas"] = f"Tipo de archivo no soportado: {extension}"
            return resultado

        resultado["texto_extraido"] = texto
        resultado["confianza_ocr"] = confianza

        # Detectar tipo de documento
        tipo_doc = detectar_tipo_documento(nombre_original, texto)
        resultado["tipo_doc"] = tipo_doc

        # Extraer datos estructurados con IA (solo si hay texto suficiente)
        if len(texto.split()) > 10:
            datos = await extraer_datos_documento(texto, tipo_doc)
            resultado["datos_extraidos"] = datos

        resultado["estado"] = "procesado"
        logger.info(f"Documento procesado: {nombre_original} | tipo: {tipo_doc} | confianza: {confianza:.2f}")

    except Exception as e:
        resultado["notas"] = f"Error en procesamiento: {str(e)}"
        logger.error(f"Error procesando {nombre_original}: {e}")

    return resultado
