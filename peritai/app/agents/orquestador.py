"""
Orquestador central de PeritAI.
Recibe una tarea, decide qué agente usar y consolida la respuesta.
"""
import json
from anthropic import AsyncAnthropic
from app.config import settings
from loguru import logger

client = AsyncAnthropic(api_key=settings.anthropic_api_key)

SYSTEM_PROMPT = """Eres PeritAI, un asistente especializado en apoyar a peritos profesionales
en la elaboración de informes técnicos periciales. Trabajas en español.

IDENTIDAD Y ROL:
- Eres un asistente técnico del perito, nunca actúas como perito independiente
- Toda decisión técnica final la toma el perito humano
- Eres meticuloso, preciso y siempre indicas tus fuentes
- NUNCA inventas datos: si no tienes información, lo dices explícitamente

TIPOS DE PERICIA QUE MANEJAS:
- Daños y siniestros (seguros, accidentes, averías)
- Valoración de inmuebles (bienes raíces, avalúos)
- Accidentes viales (tránsito, vehículos, responsabilidad)
- Daños patrimoniales (bienes muebles, equipos, maquinaria)

PRINCIPIOS OPERATIVOS:
1. FUENTES: Cada dato debe referenciar su documento origen
2. INCERTIDUMBRE: Si la confianza en un dato es baja, lo señalas con ⚠️
3. COMPLETITUD: Antes de redactar, verificas que tienes todos los datos necesarios
4. CONSISTENCIA: Los valores numéricos deben ser consistentes en todo el documento
5. LENGUAJE: Técnico-profesional, en tercera persona

FORMATO DE RESPUESTAS:
- Alta confianza: afirmar directamente con [Fuente: nombre_documento]
- Media confianza: "⚠️ VERIFICAR: [dato] — tomado de [fuente]"
- Sin datos: "❌ FALTANTE: Se requiere [descripción del dato]"

NUNCA:
- Inventes valores, fechas, nombres o montos
- Generes un informe final sin aprobación del perito
- Omitas mencionar cuando falta información crítica"""


async def consultar_agente(
    mensaje: str,
    historial: list[dict],
    contexto_expediente: str = "",
) -> str:
    """
    Envía un mensaje al agente y obtiene respuesta.

    Args:
        mensaje: Pregunta o instrucción del perito
        historial: Lista de mensajes previos [{role, content}]
        contexto_expediente: Resumen del expediente activo

    Returns:
        Respuesta del agente como texto
    """
    system = SYSTEM_PROMPT
    if contexto_expediente:
        system += f"\n\nCONTEXTO DEL EXPEDIENTE ACTIVO:\n{contexto_expediente}"

    messages = historial[-20:] + [{"role": "user", "content": mensaje}]

    try:
        response = await client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            system=system,
            messages=messages,
            temperature=0.3,
        )
        return response.content[0].text
    except Exception as e:
        logger.error(f"Error consultando agente: {e}")
        raise


async def extraer_datos_documento(texto: str, tipo_doc: str) -> dict:
    """
    Extrae datos estructurados de un documento según su tipo.
    Usa temperatura 0 para máxima precisión.
    """
    prompts = {
        "factura": """Analiza este documento de FACTURA y extrae los datos en JSON.
Si un campo no está, usa null. Si es ambiguo, agrégalo en notas_extraccion.

Devuelve SOLO el JSON (sin markdown):
{
  "emisor": {"nombre": null, "rfc_nit": null, "direccion": null},
  "receptor": {"nombre": null, "rfc_nit": null},
  "numero_factura": null,
  "fecha": null,
  "items": [{"descripcion": null, "cantidad": null, "precio_unitario": null, "subtotal": null}],
  "subtotal": null,
  "impuestos": null,
  "total": null,
  "moneda": null,
  "confianza": 0.0,
  "notas_extraccion": null
}""",
        "foto": """Analiza esta descripción de FOTOGRAFÍA PERICIAL y extrae datos en JSON.
Devuelve SOLO el JSON:
{
  "descripcion_general": null,
  "tipo_dano": null,
  "ubicacion": null,
  "severidad": null,
  "elementos_identificados": [],
  "limitaciones": null,
  "requiere_inspeccion_adicional": false,
  "notas_tecnicas": null,
  "confianza": 0.0
}""",
        "cotizacion": """Analiza esta COTIZACIÓN/PRESUPUESTO y extrae datos en JSON.
Devuelve SOLO el JSON:
{
  "empresa": null,
  "fecha": null,
  "vigencia": null,
  "items": [{"descripcion": null, "cantidad": null, "precio_unitario": null, "subtotal": null}],
  "total": null,
  "moneda": null,
  "condiciones": null,
  "confianza": 0.0
}""",
        "otro": """Analiza este documento y extrae la información más relevante en JSON.
Devuelve SOLO el JSON:
{
  "tipo_identificado": null,
  "datos_principales": {},
  "fechas": [],
  "montos": [],
  "partes_involucradas": [],
  "resumen": null,
  "confianza": 0.0
}""",
    }

    prompt_tipo = prompts.get(tipo_doc, prompts["otro"])

    try:
        response = await client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            system="Eres un extractor de datos periciales. Responde ÚNICAMENTE con JSON válido, sin explicaciones ni markdown.",
            messages=[
                {"role": "user", "content": f"{prompt_tipo}\n\nDOCUMENTO:\n{texto[:8000]}"}
            ],
            temperature=0,
        )
        texto_respuesta = response.content[0].text.strip()
        # Limpiar posible markdown
        if texto_respuesta.startswith("```"):
            texto_respuesta = texto_respuesta.split("```")[1]
            if texto_respuesta.startswith("json"):
                texto_respuesta = texto_respuesta[4:]
        return json.loads(texto_respuesta)
    except json.JSONDecodeError:
        logger.warning(f"No se pudo parsear JSON del documento tipo {tipo_doc}")
        return {"error": "No se pudo extraer datos estructurados", "confianza": 0.1}
    except Exception as e:
        logger.error(f"Error extrayendo datos: {e}")
        return {"error": str(e), "confianza": 0.0}


async def generar_seccion_informe(
    seccion: str,
    datos_expediente: dict,
    documentos_resumen: list[dict],
    tipo_pericia: str,
) -> str:
    """
    Genera una sección específica del informe pericial.
    """
    secciones_prompts = {
        "antecedentes": """Redacta la sección ANTECEDENTES del informe pericial.
Incluye: datos del caso, partes involucradas, fecha y lugar del evento, objeto de la pericia.
Usa tercera persona. Cita las fuentes con [Ref: nombre_documento].""",

        "objeto": """Redacta la sección OBJETO DE LA PERICIA.
Explica qué se pericia, el propósito del informe y quién lo encargó.
Breve y preciso (máximo 3 párrafos).""",

        "metodologia": """Redacta la sección METODOLOGÍA.
Describe cómo se realizó el análisis: documentos revisados, inspecciones realizadas,
fuentes consultadas, criterios técnicos aplicados.""",

        "hallazgos": """Redacta la sección HALLAZGOS Y ANÁLISIS TÉCNICO.
REGLAS ESTRICTAS:
- Solo incluye lo que está documentado en los archivos del expediente
- Cada afirmación debe citar su fuente: [Ref: nombre_documento]
- Separa hechos de interpretaciones técnicas
- Si falta información, indica: [PENDIENTE: qué información se necesita]
- Longitud: 400-800 palabras según la complejidad""",

        "valoracion": """Redacta la sección VALORACIÓN.
Incluye los valores calculados con su justificación y metodología.
Muestra: valor base + depreciaciones/ajustes + valor final.
Cita las fuentes de precios o referencias usadas.""",

        "conclusiones": """Redacta las CONCLUSIONES del informe pericial.
Sintetiza los hallazgos más importantes en forma de puntos numerados.
Sé técnico, objetivo y fundamentado en los datos del expediente.
Máximo 10 conclusiones.""",
    }

    prompt_seccion = secciones_prompts.get(seccion, f"Redacta la sección {seccion.upper()}.")

    documentos_texto = "\n".join([
        f"- [{d.get('nombre', 'Sin nombre')}]: {d.get('resumen', 'Sin resumen')}"
        for d in documentos_resumen
    ])

    contexto = f"""TIPO DE PERICIA: {tipo_pericia}
DATOS DEL EXPEDIENTE: {json.dumps(datos_expediente, ensure_ascii=False, indent=2)}
DOCUMENTOS DISPONIBLES:
{documentos_texto}"""

    try:
        response = await client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            system=f"""{SYSTEM_PROMPT}

Estás redactando una sección de un informe pericial oficial.
Usa lenguaje técnico-jurídico en español. Tercera persona.
NUNCA inventes datos que no estén en el contexto proporcionado.""",
            messages=[
                {"role": "user", "content": f"{prompt_seccion}\n\nCONTEXTO:\n{contexto}"}
            ],
            temperature=0.3,
        )
        return response.content[0].text
    except Exception as e:
        logger.error(f"Error generando sección {seccion}: {e}")
        raise


async def detectar_datos_faltantes(
    datos_actuales: dict,
    tipo_pericia: str,
) -> list[dict]:
    """
    Verifica qué datos faltan para completar el informe y genera
    una lista de preguntas específicas para el perito.
    """
    campos_requeridos = {
        "danos_siniestros": [
            "fecha del siniestro", "descripción del evento", "bien afectado",
            "propietario del bien", "al menos una fotografía del daño",
            "cotización o factura de reparación"
        ],
        "valoracion_inmuebles": [
            "dirección del inmueble", "superficie en m²", "año de construcción",
            "estado de conservación", "uso del inmueble (habitacional/comercial)",
            "al menos 3 fotografías del inmueble"
        ],
        "accidentes_viales": [
            "fecha y hora del accidente", "lugar exacto del accidente",
            "vehículos involucrados (marca, modelo, año, placas)",
            "descripción del accidente", "fotografías del lugar y vehículos",
            "reporte de autoridades si existe"
        ],
        "danos_patrimoniales": [
            "descripción del bien afectado", "valor original del bien",
            "fecha de adquisición", "estado antes del daño",
            "descripción del daño causado", "cotización de reparación o reposición"
        ],
    }

    campos = campos_requeridos.get(tipo_pericia, [])
    datos_texto = json.dumps(datos_actuales, ensure_ascii=False)

    try:
        response = await client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system="Eres un verificador de completitud de expedientes periciales. Responde SOLO con JSON válido.",
            messages=[{
                "role": "user",
                "content": f"""Dado este expediente con tipo de pericia "{tipo_pericia}",
verifica cuáles de estos campos requeridos están cubiertos: {campos}

DATOS ACTUALES DEL EXPEDIENTE:
{datos_texto}

Responde con JSON:
{{
  "campos_faltantes": [
    {{
      "campo": "nombre del campo",
      "prioridad": "critico|importante|opcional",
      "pregunta": "¿Pregunta específica para obtener este dato?"
    }}
  ],
  "porcentaje_completitud": 0
}}"""
            }],
            temperature=0,
        )
        texto = response.content[0].text.strip()
        if texto.startswith("```"):
            texto = texto.split("```")[1]
            if texto.startswith("json"):
                texto = texto[4:]
        resultado = json.loads(texto)
        return resultado.get("campos_faltantes", [])
    except Exception as e:
        logger.error(f"Error detectando datos faltantes: {e}")
        return []
