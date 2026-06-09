"""Servidor principal de PeritAI."""
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
from fastapi.requests import Request
from pathlib import Path
from loguru import logger
from app.models.database import init_db
from app.routes import expedientes, documentos, agente
from app.config import settings

BASE_DIR = Path(__file__).parent.parent

app = FastAPI(
    title="PeritAI",
    description="Sistema de elaboración de informes periciales con IA",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "frontend" / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "frontend" / "templates"))

# Registrar rutas de la API
app.include_router(expedientes.router, prefix="/api")
app.include_router(documentos.router, prefix="/api")
app.include_router(agente.router, prefix="/api")


@app.on_event("startup")
async def startup():
    Path(settings.upload_dir).mkdir(parents=True, exist_ok=True)
    await init_db()
    logger.info("PeritAI iniciado correctamente")


@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse("index.html", {
        "request": request,
        "nombre_despacho": settings.nombre_despacho,
    })


@app.get("/expediente/{expediente_id}")
async def vista_expediente(request: Request, expediente_id: str):
    return templates.TemplateResponse("expediente.html", {
        "request": request,
        "expediente_id": expediente_id,
        "nombre_despacho": settings.nombre_despacho,
    })


@app.get("/health")
async def health():
    return {"status": "ok", "sistema": "PeritAI v1.0"}
