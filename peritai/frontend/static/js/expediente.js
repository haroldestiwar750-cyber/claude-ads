const API = '/api';
const TIPOS = {
    danos_siniestros: 'Daños y Siniestros',
    valoracion_inmuebles: 'Valoración de Inmuebles',
    accidentes_viales: 'Accidentes Viales',
    danos_patrimoniales: 'Daños Patrimoniales',
};
const ICONOS_DOC = {
    factura: '🧾', cotizacion: '💰', foto: '📷',
    contrato: '📋', otro: '📄',
};

let informeSecciones = {};

// ── Carga inicial ─────────────────────────────────────────────
async function init() {
    await Promise.all([cargarExpediente(), cargarDocumentos()]);
}

async function cargarExpediente() {
    try {
        const resp = await fetch(`${API}/expedientes/${EXPEDIENTE_ID}`);
        const exp = await resp.json();
        document.getElementById('exp-titulo-header').textContent = exp.titulo;
        document.getElementById('info-expediente').innerHTML = `
            <div class="codigo-exp">${exp.codigo}</div>
            <div class="titulo-exp">${exp.titulo}</div>
            <span class="tipo-badge">${TIPOS[exp.tipo] || exp.tipo}</span>
            <div class="info-row"><strong>Cliente:</strong> ${exp.cliente_nombre || '—'}</div>
            <div class="info-row"><strong>Contacto:</strong> ${exp.cliente_contacto || '—'}</div>
            <div class="info-row"><strong>Fecha siniestro:</strong> ${exp.fecha_siniestro || '—'}</div>
            <div class="info-row"><strong>Ubicación:</strong> ${exp.ubicacion || '—'}</div>
            ${exp.descripcion ? `<div class="info-row" style="margin-top:8px">${exp.descripcion}</div>` : ''}
        `;
    } catch (err) {
        console.error(err);
    }
}

async function cargarDocumentos() {
    try {
        const resp = await fetch(`${API}/documentos/${EXPEDIENTE_ID}`);
        const docs = await resp.json();
        renderDocumentos(docs);
    } catch (err) {
        console.error(err);
    }
}

function renderDocumentos(docs) {
    const lista = document.getElementById('lista-docs');
    if (docs.length === 0) {
        lista.innerHTML = '<p style="font-size:0.82rem;color:#aaa;text-align:center;padding:10px">Sin documentos aún</p>';
        return;
    }
    lista.innerHTML = docs.map(d => `
        <div class="doc-item" title="${d.nombre_original}">
            <span class="doc-icon">${ICONOS_DOC[d.tipo_doc] || '📄'}</span>
            <span class="doc-nombre">${d.nombre_original}</span>
            <span class="doc-estado estado-${d.estado}">${d.estado}</span>
        </div>
    `).join('');
}

// ── Subida de archivos ────────────────────────────────────────
const zonaElem = document.getElementById('zona-drop');
const inputArchivo = document.getElementById('input-archivo');

zonaElem.addEventListener('dragover', e => { e.preventDefault(); zonaElem.classList.add('dragover'); });
zonaElem.addEventListener('dragleave', () => zonaElem.classList.remove('dragover'));
zonaElem.addEventListener('drop', async e => {
    e.preventDefault();
    zonaElem.classList.remove('dragover');
    await subirArchivos(e.dataTransfer.files);
});
inputArchivo.addEventListener('change', async e => { await subirArchivos(e.target.files); });

async function subirArchivos(files) {
    for (const archivo of files) {
        const fd = new FormData();
        fd.append('archivo', archivo);
        try {
            const resp = await fetch(`${API}/documentos/subir/${EXPEDIENTE_ID}`, {
                method: 'POST', body: fd,
            });
            if (!resp.ok) {
                const err = await resp.json();
                alert(`Error con "${archivo.name}": ${err.detail}`);
            }
        } catch (err) {
            alert(`Error subiendo "${archivo.name}": ${err.message}`);
        }
    }
    // Recargar lista después de 2 segundos (tiempo para procesar)
    setTimeout(cargarDocumentos, 2000);
    setTimeout(cargarDocumentos, 6000); // segunda recarga por si tarda más
}

// ── Tabs ──────────────────────────────────────────────────────
function cambiarTab(nombre) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('activo'));
    document.querySelectorAll('.tab-contenido').forEach(c => c.classList.add('oculto'));
    event.target.classList.add('activo');
    document.getElementById(`tab-${nombre}`).classList.remove('oculto');
}

// ── Chat ──────────────────────────────────────────────────────
function sugerir(texto) {
    document.getElementById('input-chat').value = texto;
    document.getElementById('input-chat').focus();
}

async function enviarMensaje() {
    const inputElem = document.getElementById('input-chat');
    const texto = inputElem.value.trim();
    if (!texto) return;

    inputElem.value = '';
    const chat = document.getElementById('mensajes-chat');

    // Mensaje del usuario
    chat.innerHTML += `<div class="mensaje usuario">${escapeHtml(texto)}</div>`;

    // Indicador de escritura
    const pensandoId = 'pensando-' + Date.now();
    chat.innerHTML += `
        <div class="pensando" id="${pensandoId}">
            <span></span><span></span><span></span>
        </div>`;
    chat.scrollTop = chat.scrollHeight;

    try {
        const resp = await fetch(`${API}/agente/chat`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ mensaje: texto, expediente_id: EXPEDIENTE_ID }),
        });
        const data = await resp.json();

        document.getElementById(pensandoId)?.remove();
        chat.innerHTML += `<div class="mensaje asistente">${formatearRespuesta(data.respuesta)}</div>`;
    } catch (err) {
        document.getElementById(pensandoId)?.remove();
        chat.innerHTML += `<div class="mensaje asistente" style="color:#dc2626">Error: ${err.message}</div>`;
    }
    chat.scrollTop = chat.scrollHeight;
}

// ── Generación de informe ─────────────────────────────────────
const NOMBRES_SECCIONES = {
    antecedentes: 'I. ANTECEDENTES',
    objeto: 'II. OBJETO DE LA PERICIA',
    metodologia: 'III. METODOLOGÍA',
    hallazgos: 'IV. HALLAZGOS Y ANÁLISIS TÉCNICO',
    valoracion: 'V. VALORACIÓN',
    conclusiones: 'VI. CONCLUSIONES',
};

async function generarSeccion(seccion) {
    const editor = document.getElementById('editor-informe');

    // Remover vacío placeholder
    const vacio = editor.querySelector('.editor-vacio');
    if (vacio) vacio.remove();

    // Crear o actualizar sección
    let seccionElem = document.getElementById(`seccion-${seccion}`);
    if (!seccionElem) {
        seccionElem = document.createElement('div');
        seccionElem.className = 'seccion-informe';
        seccionElem.id = `seccion-${seccion}`;
        editor.appendChild(seccionElem);
    }
    seccionElem.innerHTML = `
        <h3>${NOMBRES_SECCIONES[seccion]} <span style="color:#d97706;font-weight:400">(generando...)</span></h3>
        <div class="texto-seccion" style="color:#aaa">Por favor espera...</div>`;

    try {
        const resp = await fetch(`${API}/agente/generar-seccion`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ expediente_id: EXPEDIENTE_ID, seccion }),
        });
        const data = await resp.json();

        informeSecciones[seccion] = data.contenido;

        seccionElem.innerHTML = `
            <h3>${NOMBRES_SECCIONES[seccion]}</h3>
            <div class="texto-seccion" contenteditable="true"
                 onblur="informeSecciones['${seccion}'] = this.textContent">
                ${escapeHtml(data.contenido)}
            </div>`;

        // Cambiar al tab del informe si no está visible
        document.querySelectorAll('.tab')[1].click();
        editor.scrollTop = seccionElem.offsetTop;
    } catch (err) {
        seccionElem.innerHTML = `
            <h3>${NOMBRES_SECCIONES[seccion]}</h3>
            <div class="texto-seccion" style="color:#dc2626">Error: ${err.message}</div>`;
    }
}

// ── Verificación ──────────────────────────────────────────────
async function verificarDatos() {
    const resultado = document.getElementById('resultado-verificacion');
    resultado.innerHTML = '<p class="info-text">Analizando expediente...</p>';

    try {
        const resp = await fetch(`${API}/agente/datos-faltantes/${EXPEDIENTE_ID}`);
        const data = await resp.json();

        if (data.datos_faltantes.length === 0) {
            resultado.innerHTML = `
                <div style="text-align:center;padding:40px">
                    <div style="font-size:2rem">✅</div>
                    <p style="color:#16a34a;font-weight:600;margin-top:8px">El expediente está completo</p>
                    <p style="color:#888;font-size:0.85rem">Puedes proceder a generar el informe</p>
                </div>`;
            return;
        }

        const html = data.datos_faltantes.map(f => `
            <div class="campo-faltante">
                <div class="prioridad prioridad-${f.prioridad}">● ${f.prioridad.toUpperCase()}</div>
                <div class="pregunta">${f.pregunta}</div>
            </div>
        `).join('');

        resultado.innerHTML = `
            <p style="font-size:0.85rem;color:#888;margin-bottom:12px">
                Se encontraron ${data.datos_faltantes.length} elemento(s) pendiente(s):
            </p>
            ${html}`;
    } catch (err) {
        resultado.innerHTML = `<p style="color:#dc2626">Error: ${err.message}</p>`;
    }
}

// ── Utilidades ────────────────────────────────────────────────
function escapeHtml(texto) {
    return texto.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function formatearRespuesta(texto) {
    return escapeHtml(texto)
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/`(.*?)`/g, '<code>$1</code>')
        .replace(/\n/g, '<br>');
}

init();
