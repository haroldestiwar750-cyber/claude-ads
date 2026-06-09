const API = '/api';

// Tipos de pericia legibles
const TIPOS = {
    danos_siniestros: 'Daños y Siniestros',
    valoracion_inmuebles: 'Valoración de Inmuebles',
    accidentes_viales: 'Accidentes Viales',
    danos_patrimoniales: 'Daños Patrimoniales',
};

const ESTADOS_BADGE = {
    activo: 'badge-activo',
    en_proceso: 'badge-activo',
    revision: 'badge-revision',
    cerrado: 'badge-cerrado',
};

async function cargarExpedientes() {
    const lista = document.getElementById('lista-exp');
    try {
        const resp = await fetch(`${API}/expedientes/`);
        const exps = await resp.json();

        const totales = { activo: 0, revision: 0 };
        exps.forEach(e => {
            if (e.estado === 'activo' || e.estado === 'en_proceso') totales.activo++;
            if (e.estado === 'revision') totales.revision++;
        });

        document.getElementById('total-exp').textContent = exps.length;
        document.getElementById('exp-activos').textContent = totales.activo;
        document.getElementById('exp-revision').textContent = totales.revision;

        if (exps.length === 0) {
            lista.innerHTML = '<p style="color:#888;text-align:center;padding:40px">No hay expedientes aún. Crea el primero.</p>';
            return;
        }

        lista.innerHTML = exps.map(e => `
            <a href="/expediente/${e.id}" class="exp-card">
                <div class="exp-info">
                    <div class="codigo">${e.codigo}</div>
                    <div class="titulo">${e.titulo}</div>
                    <div class="cliente">👤 ${e.cliente_nombre || 'Sin cliente'} &nbsp;·&nbsp; ${TIPOS[e.tipo] || e.tipo}</div>
                </div>
                <span class="badge ${ESTADOS_BADGE[e.estado] || 'badge-cerrado'}">${e.estado}</span>
            </a>
        `).join('');
    } catch (err) {
        lista.innerHTML = `<p style="color:#dc2626">Error cargando expedientes: ${err.message}</p>`;
    }
}

function mostrarModalNuevo() {
    document.getElementById('modal-nuevo').classList.remove('oculto');
}

function cerrarModal() {
    document.getElementById('modal-nuevo').classList.add('oculto');
}

document.getElementById('form-nuevo-exp').addEventListener('submit', async (e) => {
    e.preventDefault();
    const btn = e.target.querySelector('[type="submit"]');
    btn.disabled = true;
    btn.textContent = 'Creando...';

    try {
        const resp = await fetch(`${API}/expedientes/`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                titulo: document.getElementById('titulo').value,
                tipo: document.getElementById('tipo').value,
                cliente_nombre: document.getElementById('cliente').value,
                cliente_contacto: document.getElementById('contacto').value || null,
                fecha_siniestro: document.getElementById('fecha-siniestro').value || null,
                ubicacion: document.getElementById('ubicacion').value || null,
                descripcion: document.getElementById('descripcion').value || null,
            }),
        });

        if (!resp.ok) throw new Error('Error al crear expediente');
        const data = await resp.json();
        window.location.href = `/expediente/${data.id}`;
    } catch (err) {
        alert('Error: ' + err.message);
        btn.disabled = false;
        btn.textContent = 'Crear Expediente';
    }
});

// Cerrar modal al hacer clic fuera
document.getElementById('modal-nuevo').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) cerrarModal();
});

cargarExpedientes();
