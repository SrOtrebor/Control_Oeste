// --- NÓMINAS (RRHH) MANAGEMENT ---
let nominasData = [];

async function loadNominas() {
    try {
        const res = await fetch('/api/dashboard/nominas');
        const data = await res.json();
        if (data.success) {
            nominasData = data.records;
            renderNominas();
        }
    } catch (e) {
        console.error('Error loading nominas:', e);
    }
}

function renderNominas() {
    const tbody = document.getElementById('bodyNominas');
    if (nominasData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-state">No hay empleados registrados en la nómina</td></tr>';
        return;
    }

    tbody.innerHTML = nominasData.map(r => `
        <tr>
            <td><strong>${r.dni}</strong></td>
            <td>${r.nombre}</td>
            <td><span class="badge badge-info">${r.categoria || 'N/A'}</span></td>
            <td style="display:none;">${r.puesto_especifico || '-'}</td>
            <td>${r.activo ? '<span class="badge badge-success">Activo</span>' : '<span class="badge badge-warning">Bloqueado</span>'}</td>
            <td>
                <button class="dash-btn dash-btn-primary" style="padding: 0.25rem 0.5rem; font-size: 0.8rem;" onclick='editNomina(${JSON.stringify(r)})'>Editar</button>
                <button class="dash-btn dash-btn-danger" style="padding: 0.25rem 0.5rem; font-size: 0.8rem;" onclick="deleteNomina(${r.id})">Borrar</button>
            </td>
        </tr>
    `).join('');
}

// Modal logic for ABM Personal
const modalNomina = document.getElementById('modalNomina');
const closeNominaModal = document.getElementById('closeNominaModal');
const btnNuevaNomina = document.getElementById('btnNuevaNomina');
const btnCancelarNomina = document.getElementById('btnCancelarNomina');
const formNomina = document.getElementById('formNomina');
const selectPuestoFisico = document.getElementById('nominaPuestoFisico');
const selectEmpresa = document.getElementById('nominaEmpresa');
const btnCapturarHuella = document.getElementById('btnCapturarHuella');

async function cargarPuestosEnModal() {
    try {
        const res = await fetch('/api/config/puestos_fisicos');
        const data = await res.json();
        if (data.success) {
            selectPuestoFisico.innerHTML = '<option value="">Seleccione un Puesto...</option>' + 
                data.data.map(p => `<option value="${p.id}">${p.nombre}</option>`).join('');
        }
    } catch (e) {
        console.error('Error cargando puestos:', e);
    }
}

async function cargarEmpresasEnModal() {
    try {
        const res = await fetch('/api/config/empresas');
        const data = await res.json();
        if (data.success) {
            selectEmpresa.innerHTML = '<option value="">Seleccione una Empresa...</option>' + 
                data.data.map(p => `<option value="${p.id}">${p.nombre}</option>`).join('');
        }
    } catch (e) {
        console.error('Error cargando empresas:', e);
    }
}

function resetBioUI() {
    document.getElementById('nominaHuellaTemplate').value = '';
    document.getElementById('bioStatusIcon').textContent = 'fingerprint';
    document.getElementById('bioStatusTitle').textContent = 'Sin huella registrada';
    document.getElementById('bioStatusDesc').textContent = 'Conecte el lector SecuGen Hamster Plus y presione capturar.';
    document.querySelector('.bio-status-card').classList.remove('success');
}

function resetPhotoUI() {
    document.getElementById('nominaFotoPreview').src = '/static/assets/default-avatar.png';
    document.getElementById('nominaFotoB64').value = '';
    if (window.localStream) {
        window.localStream.getTracks().forEach(track => track.stop());
    }
    document.getElementById('webcamVideo').style.display = 'none';
    document.getElementById('btnTomarFoto').style.display = 'none';
}

if(btnNuevaNomina) {
    btnNuevaNomina.onclick = () => {
        document.getElementById('modalNominaTitle').textContent = 'Agregar Personal';
        formNomina.reset();
        document.getElementById('nominaId').value = '';
        resetBioUI();
        resetPhotoUI();
        cargarPuestosEnModal();
        cargarEmpresasEnModal();
        document.querySelector('button[data-tab="config"]').click(); setTimeout(() => document.getElementById('formNomina').scrollIntoView({behavior: 'smooth', block: 'start'}), 100);
    };
}

function closeNomina() {
    resetPhotoUI();
    // no modal to remove
}

if(closeNominaModal) closeNominaModal.onclick = closeNomina;
if(btnCancelarNomina) btnCancelarNomina.onclick = closeNomina;



// Lógica de Foto
const fotoUpload = document.getElementById('nominaFotoUpload');
if (fotoUpload) {
    fotoUpload.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                document.getElementById('nominaFotoPreview').src = e.target.result;
                document.getElementById('nominaFotoB64').value = e.target.result;
            }
            reader.readAsDataURL(file);
        }
    });
}

const btnActivarCamara = document.getElementById('btnActivarCamara');
const btnTomarFoto = document.getElementById('btnTomarFoto');
const video = document.getElementById('webcamVideo');
const canvas = document.getElementById('photoCanvas');

if (btnActivarCamara) {
    btnActivarCamara.onclick = async () => {
        try {
            const stream = await navigator.mediaDevices.getUserMedia({ video: true });
            window.localStream = stream;
            video.srcObject = stream;
            video.style.display = 'block';
            btnTomarFoto.style.display = 'inline-flex';
        } catch (err) {
            alert('No se pudo acceder a la cámara: ' + err);
        }
    };
}

if (btnTomarFoto) {
    btnTomarFoto.onclick = () => {
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        canvas.getContext('2d').drawImage(video, 0, 0);
        const dataUrl = canvas.toDataURL('image/jpeg');
        document.getElementById('nominaFotoPreview').src = dataUrl;
        document.getElementById('nominaFotoB64').value = dataUrl;
        
        // Stop stream
        if (window.localStream) {
            window.localStream.getTracks().forEach(track => track.stop());
        }
        video.style.display = 'none';
        btnTomarFoto.style.display = 'none';
    };
}

async function editNomina(r) {
    document.getElementById('modalNominaTitle').textContent = 'Editar Personal';
    formNomina.reset();
    resetBioUI();
    resetPhotoUI();
    await cargarPuestosEnModal();
    await cargarEmpresasEnModal();
    
    document.getElementById('nominaId').value = r.id;
    document.getElementById('nominaDni').value = r.dni;
    document.getElementById('nominaNombre').value = r.nombre;
    
    if (r.id_empresa) document.getElementById('nominaEmpresa').value = r.id_empresa;
    if (r.tipo_empleado) document.getElementById('nominaTipoEmpleado').value = r.tipo_empleado;
    
    // El puesto físico base (categoria original)
    if (r.categoria) {
        let optionFound = false;
        for (let i = 0; i < selectPuestoFisico.options.length; i++) {
            if (selectPuestoFisico.options[i].text === r.categoria) {
                selectPuestoFisico.selectedIndex = i;
                optionFound = true;
                break;
            }
        }
    }
    
    document.getElementById('nominaActivo').value = r.activo ? '1' : '0';
    
    if (r.huella_template) {
        document.getElementById('nominaHuellaTemplate').value = 'EXISTE';
        document.getElementById('bioStatusIcon').textContent = 'check_circle';
        document.getElementById('bioStatusTitle').textContent = 'Huella Registrada';
        document.getElementById('bioStatusDesc').textContent = 'El empleado ya cuenta con datos biométricos.';
        document.querySelector('.bio-status-card').classList.add('success');
    }

    if (r.foto_path) {
        // En un caso real acá cargaríamos r.foto_path si es URL, o el b64.
        // Si el backend devuelve b64 directo, lo usamos
        if (r.foto_path.startsWith('data:image')) {
            document.getElementById('nominaFotoPreview').src = r.foto_path;
            document.getElementById('nominaFotoB64').value = r.foto_path;
        }
    }
    
    document.querySelector('button[data-tab="config"]').click(); setTimeout(() => document.getElementById('formNomina').scrollIntoView({behavior: 'smooth', block: 'start'}), 100);
}

async function deleteNomina(id) {
    if(!confirm("¿Está seguro de eliminar este empleado del sistema?")) return;
    try {
        const res = await fetch(`/api/dashboard/nomina/${id}`, { method: 'DELETE' });
        const data = await res.json();
        if(data.success) {
            loadNominas();
        } else {
            alert("Error al borrar: " + data.message);
        }
    } catch(e) {
        console.error(e);
        alert("Error de red");
    }
}

if(formNomina) {
    formNomina.onsubmit = async (e) => {
        e.preventDefault();
        const puestoOption = selectPuestoFisico.options[selectPuestoFisico.selectedIndex];
        const payload = {
            id: document.getElementById('nominaId').value,
            dni: document.getElementById('nominaDni').value.trim(),
            nombre: document.getElementById('nominaNombre').value.trim(),
            categoria: puestoOption && puestoOption.value ? puestoOption.text : '', // Guardamos el texto como "puesto historico"
            id_empresa: document.getElementById('nominaEmpresa').value,
            tipo_empleado: document.getElementById('nominaTipoEmpleado').value,
            foto_b64: document.getElementById('nominaFotoB64').value,
            activo: document.getElementById('nominaActivo').value === '1',
            huella_template: document.getElementById('nominaHuellaTemplate').value
        };

        try {
            const res = await fetch('/api/dashboard/nomina', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const data = await res.json();
            if(data.success) {
                closeNomina();
                loadNominas();
            } else {
                alert("Error al guardar: " + data.message);
            }
        } catch (e) {
            console.error(e);
            alert("Error de red");
        }
    };
}

// Simulador de SecuGen (Phase 3 Stub)
if (btnCapturarHuella) {
    btnCapturarHuella.onclick = async () => {
        btnCapturarHuella.innerHTML = '<span class="material-icons">hourglass_empty</span> Inicializando SecuGen...';
        btnCapturarHuella.disabled = true;
        
        try {
            const res = await fetch('/api/biometria/capturar', { method: 'POST' });
            const data = await res.json();
            
            if (data.success) {
                document.getElementById('nominaHuellaTemplate').value = data.template;
                document.getElementById('bioStatusIcon').textContent = 'check_circle';
                document.getElementById('bioStatusTitle').textContent = 'Huella Capturada Exitosamente';
                document.getElementById('bioStatusDesc').textContent = 'Calidad: Alta (500 DPI). Lista para guardar.';
                document.querySelector('.bio-status-card').classList.add('success');
            } else {
                alert("Error del sensor: " + data.message);
            }
        } catch (e) {
            alert("Error conectando con el servicio biométrico local.");
        } finally {
            btnCapturarHuella.innerHTML = '<span class="material-icons">scanner</span> Capturar Huella';
            btnCapturarHuella.disabled = false;
        }
    };
}
// --- FIN ABM PERSONAL ---

