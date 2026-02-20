/**
 * CITRO — Panel de Administración (SharePoint + Graph)
 */
let allSolicitudes = [];
let currentEditId  = null;

async function loadAdminData() {
    document.getElementById('loading-row').style.display = '';
    try {
        allSolicitudes = await getAllSolicitudes();
        updateStats(allSolicitudes);
        renderTable(allSolicitudes);
    } catch(e) {
        document.getElementById('admin-table-body').innerHTML =
            `<tr><td colspan="9" style="text-align:center;padding:40px;color:#A80000">Error cargando SharePoint: ${e.message}</td></tr>`;
    }
}

function updateStats(data) {
    const aut = data.filter(s=>s.Estado==='Aprobado').reduce((s,r)=>s+parseFloat(r.MontoAutorizado||0),0);
    document.getElementById('total-autorizado').textContent  = `$${aut.toLocaleString('es-MX')} MXN`;
    document.getElementById('total-solicitudes').textContent = data.length;
    document.getElementById('total-pendientes').textContent  = data.filter(s=>!s.Estado||['Pendiente','En Revisión'].includes(s.Estado)).length;
    document.getElementById('total-aprobadas').textContent   = data.filter(s=>s.Estado==='Aprobado').length;
}

function renderTable(data) {
    const tbody = document.getElementById('admin-table-body');
    if (!data.length) { tbody.innerHTML='<tr><td colspan="9" style="text-align:center;padding:40px;color:#888">Sin solicitudes</td></tr>'; return; }
    tbody.innerHTML = data.map(s => {
        const mS = s.MontoSolicitado>0  ? `$${parseFloat(s.MontoSolicitado).toLocaleString('es-MX')}`:'—';
        const mA = s.MontoAutorizado>0 && s.Estado==='Aprobado' ? `$${parseFloat(s.MontoAutorizado).toLocaleString('es-MX')}`:'—';
        const nEsc = (s.NotasCT||'').replace(/'/g,"\\'");
        return `<tr>
            <td><span class="folio-badge">${s.Folio||'—'}</span></td>
            <td>${s.FechaSolicitud?new Date(s.FechaSolicitud).toLocaleDateString('es-MX'):'—'}</td>
            <td><span class="tipo-badge ${tipoCls(s.TipoTramite)}">${s.TipoTramite||'—'}</span></td>
            <td>${s.NombreSolicitante||'—'}</td>
            <td><a href="mailto:${s.EmailSolicitante}">${s.EmailSolicitante||'—'}</a></td>
            <td>${mS}</td>
            <td><span class="estado-badge ${estadoCls(s.Estado)}">${s.Estado||'Pendiente'}</span></td>
            <td>${mA}</td>
            <td class="actions-cell">
                <button class="btn-action btn-edit" onclick="openEditModal('${s.spId}','${s.Estado||'Pendiente'}',${s.MontoAutorizado||0},'${nEsc}')" title="Editar">✏️</button>
                ${s.URLPdf?`<button class="btn-action" onclick="window.open('${s.URLPdf}','_blank')" title="Ver PDF">📄</button>`:''}
            </td></tr>`;
    }).join('');
}

function filterAdmin() {
    const q = document.getElementById('search-filter').value.toLowerCase();
    const st = document.getElementById('status-filter').value;
    const tp = document.getElementById('type-filter').value;
    const filtered = allSolicitudes.filter(s =>
        (!q  || `${s.Folio}${s.NombreSolicitante}${s.EmailSolicitante}`.toLowerCase().includes(q)) &&
        (!st || s.Estado===st) && (!tp || s.TipoTramite===tp)
    );
    updateStats(filtered);
    renderTable(filtered);
}

function openEditModal(spId, estado, monto, notas) {
    currentEditId = spId;
    document.getElementById('edit-sharepoint-id').value = spId;
    document.getElementById('edit-status').value         = estado;
    document.getElementById('edit-monto').value          = monto;
    document.getElementById('edit-notas').value          = notas;
    document.getElementById('edit-modal').classList.add('show');
}
function closeEditModal() { document.getElementById('edit-modal').classList.remove('show'); currentEditId=null; }

async function saveEdit() {
    if (!currentEditId) return;
    showLoading(true);
    try {
        const estado    = document.getElementById('edit-status').value;
        const monto     = document.getElementById('edit-monto').value;
        const notas     = document.getElementById('edit-notas').value;
        const notificar = document.getElementById('edit-notify').value;
        await updateSolicitudEnSharePoint(currentEditId, { estado, montoAutorizado:monto, notasCT:notas });
        if (notificar==='si') {
            const sol = allSolicitudes.find(s=>s.spId===currentEditId);
            if (sol) await notifyEstadoChange(sol, estado, monto, notas);
        }
        const i = allSolicitudes.findIndex(s=>s.spId===currentEditId);
        if (i>=0) Object.assign(allSolicitudes[i],{Estado:estado,MontoAutorizado:parseFloat(monto)||0,NotasCT:notas});
        updateStats(allSolicitudes); renderTable(allSolicitudes);
        closeEditModal(); showToast('✅ Guardado en SharePoint');
    } catch(e) { alert('Error: '+e.message); }
    finally { showLoading(false); }
}

async function notifyEstadoChange(sol, estado, monto, notas) {
    if (!sol.EmailSolicitante) return;
    const emo = {Aprobado:'✅',Rechazado:'❌','En Revisión':'🔍',Pendiente:'⏳'}[estado]||'📋';
    const html = `<div style="font-family:Segoe UI,sans-serif;max-width:580px;margin:0 auto;border:1px solid #EDEBE9;border-radius:8px;overflow:hidden">
        <div style="background:#0078D4;padding:22px"><h2 style="color:#fff;margin:0">${emo} Actualización de Solicitud</h2></div>
        <div style="padding:24px">
        <p>Estimado/a <strong>${sol.NombreSolicitante}</strong>,</p>
        <p>El <strong>H. Consejo Técnico del CITRO</strong> ha actualizado el estado de su solicitud:</p>
        <table style="width:100%;border-collapse:collapse;background:#FAF9F8;border-radius:6px;margin:16px 0;font-size:14px">
        <tr><td style="padding:9px 12px;font-weight:600;width:40%">Folio:</td><td style="padding:9px 12px">${sol.Folio}</td></tr>
        <tr style="background:#F3F2F1"><td style="padding:9px 12px;font-weight:600">Nuevo Estado:</td><td style="padding:9px 12px"><strong>${emo} ${estado}</strong></td></tr>
        ${estado==='Aprobado'&&monto>0?`<tr><td style="padding:9px 12px;font-weight:600">Monto Autorizado:</td><td style="padding:9px 12px;color:#107C10;font-weight:700">$${parseFloat(monto).toLocaleString('es-MX')} MXN</td></tr>`:''}
        ${notas?`<tr style="background:#F3F2F1"><td style="padding:9px 12px;font-weight:600">Notas del CT:</td><td style="padding:9px 12px">${notas}</td></tr>`:''}
        </table></div>
        <div style="background:#FAF9F8;padding:14px;text-align:center;border-top:1px solid #EDEBE9"><p style="margin:0;font-size:12px;color:#888">CITRO · Universidad Veracruzana</p></div></div>`;
    await sendEmailViaGraph(sol.EmailSolicitante, `CITRO — Actualización de solicitud (${sol.Folio})`, html);
}

async function loadUserSolicitudes() {
    const c = document.getElementById('solicitudes-list');
    c.innerHTML = '<div style="text-align:center;padding:40px"><div class="spinner-small"></div> Cargando desde SharePoint...</div>';
    try {
        const data = await getSolicitudesUsuario();
        if (!data.length) {
            c.innerHTML = `<div class="empty-state"><div style="font-size:48px;margin-bottom:16px">📋</div><h3>Sin solicitudes</h3><p>Envía tu primera solicitud al Consejo Técnico</p><button class="btn-primary-large" onclick="goToHome()">Nueva Solicitud</button></div>`;
            return;
        }
        c.innerHTML = data.map(s => {
            const mS = s.MontoSolicitado>0 ? `$${parseFloat(s.MontoSolicitado).toLocaleString('es-MX')} MXN`:null;
            const mA = s.MontoAutorizado>0&&s.Estado==='Aprobado' ? `<div class="card-monto-aut">💰 <strong>Autorizado:</strong> $${parseFloat(s.MontoAutorizado).toLocaleString('es-MX')} MXN</div>`:'';
            return `<div class="solicitud-card">
                <div class="solicitud-header">
                    <div><span class="folio-badge">${s.Folio}</span> <span class="tipo-badge ${tipoCls(s.TipoTramite)}">${s.TipoTramite}</span></div>
                    <span class="estado-badge ${estadoCls(s.Estado)}">${s.Estado||'Pendiente'}</span>
                </div>
                <div class="solicitud-body">
                    <div>📅 ${s.FechaSolicitud?new Date(s.FechaSolicitud).toLocaleDateString('es-MX',{dateStyle:'long'}):'—'}</div>
                    ${mS?`<div>💵 Solicitado: <strong>${mS}</strong></div>`:''}
                    ${mA}
                    ${s.NotasCT?`<div class="card-notas"><strong>Notas del CT:</strong> ${s.NotasCT}</div>`:''}
                </div>
                ${s.URLPdf?`<div class="solicitud-footer"><a href="${s.URLPdf}" target="_blank" class="btn-nav-secondary">📄 Ver PDF en SharePoint</a></div>`:''}
            </div>`;
        }).join('');
    } catch(e) { c.innerHTML = `<div style="padding:20px;color:#A80000">Error: ${e.message}</div>`; }
}

function estadoCls(e) { return {Aprobado:'estado-aprobado',Rechazado:'estado-rechazado','En Revisión':'estado-revision'}[e]||'estado-pendiente'; }
function tipoCls(t)   { return {'Apoyo Académico':'tipo-green','Aval Institucional':'tipo-blue','Apoyo a Terceros':'tipo-purple','Comité Tutorial':'tipo-teal','Solicitud Libre':'tipo-orange'}[t]||'tipo-green'; }
function showToast(msg) {
    const d = Object.assign(document.createElement('div'),{ textContent:msg });
    d.style.cssText='position:fixed;bottom:24px;right:24px;background:#107C10;color:#fff;padding:13px 22px;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.2);z-index:9999;font-weight:600;font-family:Segoe UI,sans-serif';
    document.body.appendChild(d); setTimeout(()=>d.remove(),3500);
}
document.addEventListener('click', e => { if (e.target===document.getElementById('edit-modal')) closeEditModal(); });
