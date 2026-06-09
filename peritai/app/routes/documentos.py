"""Rutas para subida y procesamiento de documentos."""
import os
import uuid
import json
from pathlib import Path
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, BackgroundTasks
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from app.models.database import get_db
from app.models.schemas import Documento, Expediente
from app.services.procesador_documentos import procesar_documento
from app.config import settings
from loguru import logger

router = APIRouter(prefix="/documentos", tags=["documentos"])

EXTENSIONES_PERMITIDAS = {".pdf", ".docx", ".doc", ".xlsx", ".xls", ".jpg", ".jpeg", ".png", ".tiff", ".webp"}


@router.post("/subir/{expediente_id}")
async def subir_documento(
    expediente_id: str,
    background_tasks: BackgroundTasks,
    archivo: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
):
    # Verificar expediente existe
    result = await db.execute(select(Expediente).where(Expediente.id == expediente_id))
    exp = result.scalar_one_or_none()
    if not exp:
        raise HTTPException(status_code=404, detail="Expediente no encontrado")

    # Validar extensión
    ext = Path(archivo.filename).suffix.lower()
    if ext not in EXTENSIONES_PERMITIDAS:
        raise HTTPException(
            status_code=400,
            detail=f"Tipo de archivo no soportado. Permitidos: {', '.join(EXTENSIONES_PERMITIDAS)}"
        )

    # Guardar archivo en disco
    carpeta = Path(settings.upload_dir) / expediente_id
    carpeta.mkdir(parents=True, exist_ok=True)
    doc_id = str(uuid.uuid4())
    ruta_archivo = carpeta / f"{doc_id}{ext}"

    contenido = await archivo.read()
    with open(ruta_archivo, "wb") as f:
        f.write(contenido)

    # Crear registro en BD con estado pendiente
    doc = Documento(
        id=doc_id,
        expediente_id=expediente_id,
        nombre_original=archivo.filename,
        ruta_storage=str(ruta_archivo),
        estado="pendiente",
    )
    db.add(doc)
    await db.commit()

    # Procesar en segundo plano
    background_tasks.add_task(
        _procesar_en_background, doc_id, str(ruta_archivo), archivo.filename
    )

    return {
        "doc_id": doc_id,
        "nombre": archivo.filename,
        "estado": "pendiente",
        "mensaje": "Documento recibido. Procesando en segundo plano..."
    }


async def _procesar_en_background(doc_id: str, ruta: str, nombre: str):
    """Procesa el documento y actualiza la BD."""
    from app.models.database import AsyncSessionLocal
    async with AsyncSessionLocal() as db:
        try:
            resultado = await procesar_documento(ruta, nombre)

            result = await db.execute(select(Documento).where(Documento.id == doc_id))
            doc = result.scalar_one_or_none()
            if doc:
                doc.tipo_mime = resultado.get("tipo_mime", "")
                doc.tipo_doc = resultado.get("tipo_doc", "otro")
                doc.texto_extraido = resultado.get("texto_extraido", "")
                doc.datos_extraidos = json.dumps(resultado.get("datos_extraidos", {}), ensure_ascii=False)
                doc.estado = resultado.get("estado", "error")
                doc.confianza_ocr = resultado.get("confianza_ocr", 0.0)
                doc.notas = resultado.get("notas", "")
                await db.commit()
                logger.info(f"Documento {doc_id} procesado y actualizado en BD")
        except Exception as e:
            logger.error(f"Error en procesamiento background de {doc_id}: {e}")


@router.get("/{expediente_id}")
async def listar_documentos(expediente_id: str, db: AsyncSession = Depends(get_db)):
    result = await db.execute(
        select(Documento).where(Documento.expediente_id == expediente_id)
    )
    docs = result.scalars().all()
    return [
        {
            "id": d.id,
            "nombre_original": d.nombre_original,
            "tipo_doc": d.tipo_doc,
            "estado": d.estado,
            "confianza_ocr": d.confianza_ocr,
            "datos_extraidos": json.loads(d.datos_extraidos) if d.datos_extraidos else {},
            "created_at": str(d.created_at),
        }
        for d in docs
    ]


@router.get("/detalle/{doc_id}")
async def obtener_documento(doc_id: str, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Documento).where(Documento.id == doc_id))
    doc = result.scalar_one_or_none()
    if not doc:
        raise HTTPException(status_code=404, detail="Documento no encontrado")
    return {
        "id": doc.id,
        "nombre_original": doc.nombre_original,
        "tipo_doc": doc.tipo_doc,
        "estado": doc.estado,
        "confianza_ocr": doc.confianza_ocr,
        "texto_extraido": doc.texto_extraido,
        "datos_extraidos": json.loads(doc.datos_extraidos) if doc.datos_extraidos else {},
        "notas": doc.notas,
    }
