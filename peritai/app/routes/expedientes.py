"""Rutas CRUD para gestión de expedientes."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, desc
from pydantic import BaseModel
from typing import Optional
import uuid
from datetime import datetime
from app.models.database import get_db
from app.models.schemas import Expediente

router = APIRouter(prefix="/expedientes", tags=["expedientes"])


class ExpedienteCreate(BaseModel):
    titulo: str
    tipo: str
    cliente_nombre: str
    cliente_contacto: Optional[str] = None
    descripcion: Optional[str] = None
    fecha_siniestro: Optional[str] = None
    ubicacion: Optional[str] = None


class ExpedienteUpdate(BaseModel):
    titulo: Optional[str] = None
    estado: Optional[str] = None
    cliente_nombre: Optional[str] = None
    cliente_contacto: Optional[str] = None
    descripcion: Optional[str] = None
    fecha_siniestro: Optional[str] = None
    ubicacion: Optional[str] = None


def generar_codigo() -> str:
    año = datetime.now().year
    uid = str(uuid.uuid4())[:4].upper()
    return f"EXP-{año}-{uid}"


@router.get("/")
async def listar_expedientes(db: AsyncSession = Depends(get_db)):
    result = await db.execute(
        select(Expediente).order_by(desc(Expediente.created_at))
    )
    expedientes = result.scalars().all()
    return [
        {
            "id": e.id,
            "codigo": e.codigo,
            "titulo": e.titulo,
            "tipo": e.tipo,
            "estado": e.estado,
            "cliente_nombre": e.cliente_nombre,
            "created_at": str(e.created_at),
        }
        for e in expedientes
    ]


@router.post("/")
async def crear_expediente(data: ExpedienteCreate, db: AsyncSession = Depends(get_db)):
    expediente = Expediente(
        id=str(uuid.uuid4()),
        codigo=generar_codigo(),
        titulo=data.titulo,
        tipo=data.tipo,
        cliente_nombre=data.cliente_nombre,
        cliente_contacto=data.cliente_contacto,
        descripcion=data.descripcion,
        fecha_siniestro=data.fecha_siniestro,
        ubicacion=data.ubicacion,
        estado="activo",
    )
    db.add(expediente)
    await db.commit()
    await db.refresh(expediente)
    return {"id": expediente.id, "codigo": expediente.codigo, "mensaje": "Expediente creado"}


@router.get("/{expediente_id}")
async def obtener_expediente(expediente_id: str, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Expediente).where(Expediente.id == expediente_id))
    exp = result.scalar_one_or_none()
    if not exp:
        raise HTTPException(status_code=404, detail="Expediente no encontrado")
    return {
        "id": exp.id,
        "codigo": exp.codigo,
        "titulo": exp.titulo,
        "tipo": exp.tipo,
        "estado": exp.estado,
        "cliente_nombre": exp.cliente_nombre,
        "cliente_contacto": exp.cliente_contacto,
        "descripcion": exp.descripcion,
        "fecha_siniestro": exp.fecha_siniestro,
        "ubicacion": exp.ubicacion,
        "created_at": str(exp.created_at),
    }


@router.patch("/{expediente_id}")
async def actualizar_expediente(
    expediente_id: str, data: ExpedienteUpdate, db: AsyncSession = Depends(get_db)
):
    result = await db.execute(select(Expediente).where(Expediente.id == expediente_id))
    exp = result.scalar_one_or_none()
    if not exp:
        raise HTTPException(status_code=404, detail="Expediente no encontrado")

    for campo, valor in data.model_dump(exclude_none=True).items():
        setattr(exp, campo, valor)

    await db.commit()
    return {"mensaje": "Expediente actualizado"}
