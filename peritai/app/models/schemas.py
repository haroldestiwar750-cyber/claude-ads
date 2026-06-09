from sqlalchemy import Column, String, Text, Integer, Float, Boolean, DateTime, ForeignKey, Enum
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
import uuid
import enum
from app.models.database import Base


def gen_uuid():
    return str(uuid.uuid4())


class TipoPericia(str, enum.Enum):
    danos_siniestros = "danos_siniestros"
    valoracion_inmuebles = "valoracion_inmuebles"
    accidentes_viales = "accidentes_viales"
    danos_patrimoniales = "danos_patrimoniales"


class EstadoExpediente(str, enum.Enum):
    activo = "activo"
    en_proceso = "en_proceso"
    revision = "revision"
    cerrado = "cerrado"


class EstadoDocumento(str, enum.Enum):
    pendiente = "pendiente"
    procesado = "procesado"
    error = "error"


class EstadoInforme(str, enum.Enum):
    borrador = "borrador"
    revision = "revision"
    aprobado = "aprobado"
    final = "final"


class Expediente(Base):
    __tablename__ = "expedientes"

    id = Column(String, primary_key=True, default=gen_uuid)
    codigo = Column(String(50), unique=True, nullable=False)
    titulo = Column(Text, nullable=False)
    tipo = Column(String(50))
    estado = Column(String(20), default="activo")
    cliente_nombre = Column(String(200))
    cliente_contacto = Column(String(200))
    descripcion = Column(Text)
    fecha_siniestro = Column(String(20))
    ubicacion = Column(Text)
    created_at = Column(DateTime, server_default=func.now())
    updated_at = Column(DateTime, server_default=func.now(), onupdate=func.now())

    documentos = relationship("Documento", back_populates="expediente", cascade="all, delete")
    informes = relationship("Informe", back_populates="expediente", cascade="all, delete")
    memoria = relationship("MemoriaAgente", back_populates="expediente", cascade="all, delete")


class Documento(Base):
    __tablename__ = "documentos"

    id = Column(String, primary_key=True, default=gen_uuid)
    expediente_id = Column(String, ForeignKey("expedientes.id"), nullable=False)
    nombre_original = Column(Text, nullable=False)
    tipo_mime = Column(String(100))
    tipo_doc = Column(String(50))  # factura, foto, cotizacion, contrato, otro
    ruta_storage = Column(Text, nullable=False)
    texto_extraido = Column(Text)
    datos_extraidos = Column(Text)  # JSON como texto
    estado = Column(String(20), default="pendiente")
    confianza_ocr = Column(Float)
    notas = Column(Text)
    created_at = Column(DateTime, server_default=func.now())

    expediente = relationship("Expediente", back_populates="documentos")


class Informe(Base):
    __tablename__ = "informes"

    id = Column(String, primary_key=True, default=gen_uuid)
    expediente_id = Column(String, ForeignKey("expedientes.id"), nullable=False)
    version = Column(Integer, default=1)
    estado = Column(String(20), default="borrador")
    contenido = Column(Text)  # JSON como texto con secciones
    fuentes = Column(Text)    # JSON lista de documentos usados
    aprobado = Column(Boolean, default=False)
    notas_revision = Column(Text)
    ruta_word = Column(Text)
    ruta_pdf = Column(Text)
    created_at = Column(DateTime, server_default=func.now())
    updated_at = Column(DateTime, server_default=func.now(), onupdate=func.now())

    expediente = relationship("Expediente", back_populates="informes")


class MemoriaAgente(Base):
    __tablename__ = "memoria_agente"

    id = Column(String, primary_key=True, default=gen_uuid)
    expediente_id = Column(String, ForeignKey("expedientes.id"), nullable=False)
    tipo = Column(String(30))  # fact, decision, clarification, correction
    contenido = Column(Text)
    fuente_doc_id = Column(String)
    confianza = Column(Float, default=1.0)
    created_at = Column(DateTime, server_default=func.now())

    expediente = relationship("Expediente", back_populates="memoria")


class Conversacion(Base):
    __tablename__ = "conversaciones"

    id = Column(String, primary_key=True, default=gen_uuid)
    expediente_id = Column(String, ForeignKey("expedientes.id"), nullable=True)
    rol = Column(String(10))   # user | assistant
    contenido = Column(Text)
    created_at = Column(DateTime, server_default=func.now())
