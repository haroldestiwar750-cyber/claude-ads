"""Rutas para interacción con el agente IA."""
import json
import uuid
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from pydantic import BaseModel
from typing import Optional
from app.models.database import get_db
from app.models.schemas import Expediente, Documento, Conversacion, MemoriaAgente
from app.agents.orquestador import (
    consultar_agente,
    generar_seccion_informe,
    detectar_datos_faltantes,
)

router = APIRouter(prefix="/agente", tags=["agente"])


class MensajeRequest(BaseModel):
    mensaje: str
    expediente_id: Optional[str] = None


class GenerarSeccionRequest(BaseModel):
    expediente_id: str
    seccion: str  # antecedentes | objeto | metodologia | hallazgos | valoracion | conclusiones


async def _construir_contexto_expediente(expediente_id: str, db: AsyncSession) -> tuple[dict, str]:
    """Construye el contexto completo de un expediente para el agente."""
    result = await db.execute(select(Expediente).where(Expediente.id == expediente_id))
    exp = result.scalar_one_or_none()
    if not exp:
        return {}, ""

    # Obtener documentos procesados
    result_docs = await db.execute(
        select(Documento).where(
            Documento.expediente_id == expediente_id,
            Documento.estado == "procesado"
        )
    )
    docs = result_docs.scalars().all()

    datos_expediente = {
        "codigo": exp.codigo,
        "titulo": exp.titulo,
        "tipo": exp.tipo,
        "cliente": exp.cliente_nombre,
        "fecha_siniestro": exp.fecha_siniestro,
        "ubicacion": exp.ubicacion,
        "descripcion": exp.descripcion,
    }

    docs_resumen = []
    for d in docs:
        datos = json.loads(d.datos_extraidos) if d.datos_extraidos else {}
        docs_resumen.append({
            "id": d.id,
            "nombre": d.nombre_original,
            "tipo": d.tipo_doc,
            "resumen": str(datos)[:300] if datos else d.texto_extraido[:300] if d.texto_extraido else "",
        })

    contexto_texto = f"""Expediente: {exp.codigo} — {exp.titulo}
Tipo de pericia: {exp.tipo}
Cliente: {exp.cliente_nombre}
Fecha del siniestro/evento: {exp.fecha_siniestro or 'No especificada'}
Ubicación: {exp.ubicacion or 'No especificada'}
Descripción: {exp.descripcion or 'No especificada'}

Documentos disponibles ({len(docs_resumen)}):
""" + "\n".join([f"  - [{d['nombre']}] tipo: {d['tipo']}" for d in docs_resumen])

    return {"expediente": datos_expediente, "documentos": docs_resumen}, contexto_texto


@router.post("/chat")
async def chat_con_agente(request: MensajeRequest, db: AsyncSession = Depends(get_db)):
    """Conversación libre con el agente."""
    contexto_texto = ""
    if request.expediente_id:
        _, contexto_texto = await _construir_contexto_expediente(request.expediente_id, db)

    # Recuperar historial de conversación
    historial = []
    if request.expediente_id:
        result = await db.execute(
            select(Conversacion)
            .where(Conversacion.expediente_id == request.expediente_id)
            .order_by(Conversacion.created_at)
            .limit(20)
        )
        convs = result.scalars().all()
        historial = [{"role": c.rol, "content": c.contenido} for c in convs]

    # Obtener respuesta del agente
    respuesta = await consultar_agente(
        mensaje=request.mensaje,
        historial=historial,
        contexto_expediente=contexto_texto,
    )

    # Guardar en historial
    id_exp = request.expediente_id
    conv_user = Conversacion(
        id=str(uuid.uuid4()),
        expediente_id=id_exp,
        rol="user",
        contenido=request.mensaje,
    )
    conv_asst = Conversacion(
        id=str(uuid.uuid4()),
        expediente_id=id_exp,
        rol="assistant",
        contenido=respuesta,
    )
    db.add(conv_user)
    db.add(conv_asst)
    await db.commit()

    return {"respuesta": respuesta, "expediente_id": request.expediente_id}


@router.post("/generar-seccion")
async def generar_seccion(request: GenerarSeccionRequest, db: AsyncSession = Depends(get_db)):
    """Genera una sección específica del informe pericial."""
    datos, _ = await _construir_contexto_expediente(request.expediente_id, db)

    if not datos:
        raise HTTPException(status_code=404, detail="Expediente no encontrado")

    result = await db.execute(select(Expediente).where(Expediente.id == request.expediente_id))
    exp = result.scalar_one_or_none()

    texto = await generar_seccion_informe(
        seccion=request.seccion,
        datos_expediente=datos["expediente"],
        documentos_resumen=datos["documentos"],
        tipo_pericia=exp.tipo or "danos_siniestros",
    )

    return {"seccion": request.seccion, "contenido": texto}


@router.get("/datos-faltantes/{expediente_id}")
async def verificar_datos_faltantes(expediente_id: str, db: AsyncSession = Depends(get_db)):
    """Detecta qué información falta en el expediente para el informe."""
    datos, _ = await _construir_contexto_expediente(expediente_id, db)
    if not datos:
        raise HTTPException(status_code=404, detail="Expediente no encontrado")

    result = await db.execute(select(Expediente).where(Expediente.id == expediente_id))
    exp = result.scalar_one_or_none()

    faltantes = await detectar_datos_faltantes(
        datos_actuales=datos["expediente"],
        tipo_pericia=exp.tipo or "danos_siniestros",
    )

    return {
        "expediente_id": expediente_id,
        "datos_faltantes": faltantes,
        "total_faltantes": len(faltantes),
    }


@router.get("/historial/{expediente_id}")
async def obtener_historial(expediente_id: str, db: AsyncSession = Depends(get_db)):
    """Recupera el historial de conversación de un expediente."""
    result = await db.execute(
        select(Conversacion)
        .where(Conversacion.expediente_id == expediente_id)
        .order_by(Conversacion.created_at)
    )
    convs = result.scalars().all()
    return [
        {"rol": c.rol, "contenido": c.contenido, "fecha": str(c.created_at)}
        for c in convs
    ]
