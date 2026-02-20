/**
 * CITRO — Outlook Calendar vía Microsoft Graph
 */
async function addToOutlookCalendar() {
    if (!userState.isLoggedIn) { alert('Inicia sesión primero'); return; }
    const { formData:fd, currentTramite:tipo, folio } = appState;
    const cfg = FORMS_CONFIG[tipo];
    const fechaI = fd.fecha_inicio || fd.fecha_actividad;
    if (!fechaI) { addToOutlookManual(); return; }
    showLoading(true);
    try {
        const tok = await getAccessToken();
        const fechaF = fd.fecha_termino || fd.fecha_fin || fechaI;
        const event = {
            subject: `CITRO: ${cfg.title}`,
            body: { contentType:'HTML', content: buildEventBody(fd, folio, tipo) },
            start: { dateTime:`${fechaI}T09:00:00`, timeZone:'America/Mexico_City' },
            end:   { dateTime:`${fechaF}T18:00:00`, timeZone:'America/Mexico_City' },
            location: { displayName: fd.destino||fd.lugar||'CITRO - Universidad Veracruzana, Xalapa, Ver.' },
            categories: ['CITRO'],
            isReminderOn: true,
            reminderMinutesBeforeStart: 1440
        };
        const r = await fetch('https://graph.microsoft.com/v1.0/me/events', {
            method:'POST',
            headers:{ Authorization:`Bearer ${tok}`, 'Content-Type':'application/json' },
            body: JSON.stringify(event)
        });
        if (r.ok) {
            const ev = await r.json();
            const btn = document.querySelector('.btn-calendar-outlook');
            if (btn) { btn.textContent='✅ Evento creado en Outlook'; btn.style.background='#107C10'; btn.onclick=()=>window.open(ev.webLink,'_blank'); }
        } else { throw new Error('Error Graph Calendar'); }
    } catch(e) { console.error(e); addToOutlookManual(); }
    finally { showLoading(false); }
}

function buildEventBody(fd, folio, tipo) {
    const tipoNom = FORMS_CONFIG[tipo]?.title||tipo;
    const nombre  = fd.nombre_completo||fd.nombre_estudiante||fd.nombre_solicitante||'—';
    const monto   = fd.monto_total?`$${parseFloat(fd.monto_total).toLocaleString('es-MX')} MXN`:'Sin monto';
    return `<div style="font-family:Segoe UI,sans-serif">
        <h3 style="color:#0078D4">Solicitud CITRO — ${tipoNom}</h3>
        <table border="1" cellpadding="8" style="border-collapse:collapse;width:100%;font-size:14px">
        <tr style="background:#0078D4;color:white"><td colspan="2"><b>Folio: ${folio}</b></td></tr>
        <tr><td><b>Tipo</b></td><td>${tipoNom}</td></tr>
        <tr style="background:#F3F2F1"><td><b>Solicitante</b></td><td>${nombre}</td></tr>
        ${fd.titulo_actividad?`<tr><td><b>Actividad</b></td><td>${fd.titulo_actividad}</td></tr>`:''}
        ${fd.destino?`<tr style="background:#F3F2F1"><td><b>Destino</b></td><td>${fd.destino}</td></tr>`:''}
        <tr><td><b>Monto</b></td><td>${monto}</td></tr>
        <tr style="background:#FFF4CE"><td><b>Estado</b></td><td>⏳ Pendiente</td></tr>
        </table>
        <p style="color:#888;font-size:11px;margin-top:12px">Generado por Sistema CITRO — ${CONFIG.institucion.nombre}</p>
    </div>`;
}

function addToOutlookManual() {
    const { formData:fd, currentTramite:tipo, folio } = appState;
    const cfg = FORMS_CONFIG[tipo];
    const fi  = fd.fecha_inicio||new Date().toISOString().split('T')[0];
    const url = `https://outlook.office.com/calendar/0/deeplink/compose?`+
        `subject=${encodeURIComponent('CITRO: '+cfg.title)}`+
        `&body=${encodeURIComponent('Folio: '+folio+'\nTipo: '+cfg.title)}`+
        `&location=${encodeURIComponent(fd.destino||'CITRO - UV')}`+
        `&startdt=${fi}T09:00:00&enddt=${fi}T18:00:00`;
    window.open(url,'_blank');
}

// ── EXPORTAR CSV (Admin) ──────────────────────────────────────
async function exportToExcel() {
    if (!userState.isAdmin) return;
    showLoading(true);
    try {
        const data = await getAllSolicitudes();
        const headers = ['Folio','Fecha','Tipo','Nombre','Email','$ Solicitado','Estado','$ Autorizado','Notas CT'];
        const rows = data.map(s => [
            s.Folio||'',
            s.FechaSolicitud ? new Date(s.FechaSolicitud).toLocaleDateString('es-MX') : '',
            s.TipoTramite||'', s.NombreSolicitante||'', s.EmailSolicitante||'',
            s.MontoSolicitado||0, s.Estado||'Pendiente', s.MontoAutorizado||0, s.NotasCT||''
        ]);
        const csv = [headers,...rows]
            .map(r => r.map(c=>`"${String(c).replace(/"/g,'""')}"`).join(',')).join('\n');
        const blob = new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8;'});
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `CITRO_${new Date().toISOString().split('T')[0]}.csv`;
        a.click(); URL.revokeObjectURL(a.href);
    } catch(e) { alert('Error exportar: '+e.message); }
    finally { showLoading(false); }
}
