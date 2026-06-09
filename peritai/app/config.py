from pydantic_settings import BaseSettings
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent


class Settings(BaseSettings):
    anthropic_api_key: str
    host: str = "127.0.0.1"
    port: int = 8000
    debug: bool = True
    database_url: str = f"sqlite+aiosqlite:///{BASE_DIR}/data/peritai.db"
    upload_dir: str = str(BASE_DIR / "data" / "documentos")
    nombre_despacho: str = "Despacho Pericial"
    nombre_perito: str = ""
    email_perito: str = ""
    telefono_perito: str = ""

    model_config = {"env_file": str(BASE_DIR / ".env")}


settings = Settings()
