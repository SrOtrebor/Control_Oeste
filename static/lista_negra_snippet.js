// --- LISTA NEGRA (PERSONAS DE INTERÉS) ---
let listaNegraData = [];

async function loadListaNegra() {
    try {
        const res = await fetch('/api/dashboard/lista_negra');
        const data = await res.json();
        if (data.success) {
            listaNegraData = data.records;
            renderListaNegra();
        }
    } catch (e) {
        console.error('Error loading lista negra:', e);
    }
}

function renderListaNegra() {
    const tbody = document.getElementById('bodyListaNegra');
    if (!tbody) return;
    
    if (listaNegraData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="empty-state">No hay personas en la lista negra</td></tr>';
        return;
    }

    tbody.innerHTML = listaNegraData.map(r => `
        <tr>
            <td><strong>${r.dni}</strong></td>
            <td>${r.nombre || '-'}</td>
            <td><span style="color: #ff3b30; font-weight: 500;">${r.motivo}</span></td>
            <td>
                <button class="dash-btn dash-btn-danger" style="padding: 0.25rem 0.5rem; font-size: 0.8rem;" onclick="deleteListaNegra(${r.id})">Quitar</button>
            </td>
        </tr>
    `).join('');
}

async function deleteListaNegra(id) {
    if(!confirm("¿Quitar a esta persona de la Lista Negra? Ya no generará alertas silenciosas.")) return;
    try {
        const res = await fetch(`/api/dashboard/lista_negra/${id}`, { method: 'DELETE' });
        const data = await res.json();
        if(data.success) loadListaNegra();
    } catch(e) {
        console.error(e);
        alert("Error de red");
    }
}

const formListaNegra = document.getElementById('formListaNegra');
if (formListaNegra) {
    formListaNegra.onsubmit = async (e) => {
        e.preventDefault();
        const payload = {
            dni: document.getElementById('lnDni').value.trim(),
            nombre: document.getElementById('lnNombre').value.trim(),
            motivo: document.getElementById('lnMotivo').value.trim()
        };

        try {
            const res = await fetch('/api/dashboard/lista_negra', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const data = await res.json();
            if(data.success) {
                formListaNegra.reset();
                loadListaNegra();
            } else {
                alert("Error al guardar: " + data.message);
            }
        } catch (e) {
            console.error(e);
            alert("Error de red");
        }
    };
}
