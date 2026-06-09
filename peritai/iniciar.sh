#!/bin/bash
# PeritAI — Script de inicio
# Ejecutar con: bash iniciar.sh

set -e

echo ""
echo "╔═══════════════════════════════════════╗"
echo "║         PeritAI — Iniciando...        ║"
echo "╚═══════════════════════════════════════╝"
echo ""

# Verificar Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: Python 3 no está instalado."
    echo "   Descárgalo en: https://www.python.org/downloads/"
    exit 1
fi

# Verificar archivo .env
if [ ! -f ".env" ]; then
    if [ -f ".env.example" ]; then
        echo "⚠️  No se encontró el archivo .env"
        echo "   Copiando .env.example como .env..."
        cp .env.example .env
        echo ""
        echo "📝 IMPORTANTE: Edita el archivo .env y agrega tu clave de Anthropic:"
        echo "   ANTHROPIC_API_KEY=sk-ant-XXXXXXXXXX"
        echo ""
        echo "   Puedes obtener tu clave en: https://console.anthropic.com/"
        echo ""
        read -p "¿Ya tienes tu clave de API? Presiona Enter cuando hayas editado .env..."
    else
        echo "❌ Error: No se encontró .env ni .env.example"
        exit 1
    fi
fi

# Verificar que la clave de API está configurada
if grep -q "sk-ant-XXXXXXXX" .env; then
    echo "❌ Error: Debes configurar tu ANTHROPIC_API_KEY en el archivo .env"
    echo "   Edita el archivo .env y reemplaza sk-ant-XXXXXXXX con tu clave real"
    exit 1
fi

# Crear entorno virtual si no existe
if [ ! -d "venv" ]; then
    echo "📦 Creando entorno virtual de Python..."
    python3 -m venv venv
fi

# Activar entorno virtual
source venv/bin/activate

# Instalar dependencias
echo "📦 Verificando dependencias..."
pip install -q -r requirements.txt

# Verificar Tesseract para OCR
if ! command -v tesseract &> /dev/null; then
    echo ""
    echo "⚠️  Tesseract OCR no está instalado."
    echo "   El sistema funcionará, pero no podrá leer documentos escaneados."
    echo "   Para instalarlo:"
    echo "     Mac:   brew install tesseract tesseract-lang"
    echo "     Ubuntu: sudo apt install tesseract-ocr tesseract-ocr-spa"
    echo "     Windows: https://github.com/UB-Mannheim/tesseract/wiki"
    echo ""
fi

# Crear carpetas necesarias
mkdir -p data/expedientes data/documentos data/informes

echo ""
echo "✅ Todo listo. Iniciando PeritAI..."
echo ""
echo "┌────────────────────────────────────────┐"
echo "│  Abre tu navegador en:                 │"
echo "│  http://localhost:8000                 │"
echo "│                                        │"
echo "│  Presiona Ctrl+C para detener          │"
echo "└────────────────────────────────────────┘"
echo ""

python3 -m uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload
