with open('static/dashboard.js', 'a', encoding='utf-8') as f:
    f.write('''
// --- CONFIG: Puestos Operativos ---
const formPuestos = document.getElementById('formPuestos');
const listaPuestos = document.getElementById('listaPuestos');

async function loadPuestos() {
    if (!listaPuestos) return;
    try {
        const res = await fetch('/api/config/puestos');
        const data = await res.json();
        if (data.success) {
            if (data.data.length === 0) {
                listaPuestos.innerHTML = '<li style="color: var(--dash-text-muted); text-align: center; font-size: 0.85rem;">No hay puestos registrados</li>';
            } else {
                listaPuestos.innerHTML = data.data.map(p => 
                    '<li style="display: flex; justify-content: space-between; align-items: center; padding: 0.5rem; background: var(--dash-surface); border-radius: var(--dash-radius-sm);">' +
                        '<span>' + p.nombre + '</span>' +
                        '<button type="button" class="dash-btn dash-btn-danger" style="padding: 0.25rem 0.5rem; font-size: 0.75rem;" onclick="deletePuesto(' + p.id + ')">' +
                            '<span class="material-icons" style="font-size: 1rem;">delete</span>' +
                        '</button>' +
                    '</li>'
                ).join('');
            }
        }
    } catch (e) {
        console.error('Error loading puestos:', e);
    }
}

if (formPuestos) {
    formPuestos.onsubmit = async (e) => {
        e.preventDefault();
        const input = document.getElementById('nuevoPuestoNombre');
        const nombre = input.value.trim();
        if (!nombre) return;
        
        try {
            const res = await fetch('/api/config/puestos', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ nombre })
            });
            const data = await res.json();
            if (data.success) {
                input.value = '';
                loadPuestos();
            } else {
                alert('Error al crear puesto: ' + data.message);
            }
        } catch (e) {
            alert('Error de red');
        }
    };
}

window.deletePuesto = async function(id) {
    if (!confirm('¿Seguro que desea eliminar este puesto?')) return;
    try {
        const res = await fetch('/api/config/puestos/' + id, { method: 'DELETE' });
        const data = await res.json();
        if (data.success) {
            loadPuestos();
        } else {
            alert('Error al eliminar puesto: ' + data.message);
        }
    } catch (e) {
        alert('Error de red');
    }
};

// Cargar puestos cuando se abre la tab de config
document.addEventListener('DOMContentLoaded', () => {
    // Si estamos en la página del dashboard, hook al cambio de tab
    const configTabBtn = document.querySelector('button[data-tab="config"]');
    if (configTabBtn) {
        configTabBtn.addEventListener('click', loadPuestos);
    }
    // Opcionalmente cargarlos de entrada
    if(document.getElementById('panel-config') && !document.getElementById('panel-config').classList.contains('hidden')){
        loadPuestos();
    }
});
''')
