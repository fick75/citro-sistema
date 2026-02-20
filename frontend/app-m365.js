/**
 * CITRO — Lógica Principal M365
 */
const appState = { currentTramite:null, formData:{}, folio:null, lastPdfUrl:null };

// ── NAVEGACIÓN ────────────────────────────────────────────────
function showSection(name) {
    document.querySelectorAll('.landing-page,.form-section,.mis-solicitudes-section,.admin-panel-section,.success-section')
        .forEach(s => s.classList.remove('active'));
    const map = { landing:'landing-page', form:'form-section', 'mis-solicitudes':'mis-solicitudes-section',
                  'admin-panel':'admin-panel-section', success:'success-section' };
    document.getElementById(map[name])?.classList.add('active');
    window.scrollTo({ top:0, behavior:'smooth' });
}
function goToHome()      { showSection('landing'); resetForm(); }
function backToLanding() { showSection('landing'); resetForm(); }
function showLoading(v)  { document.getElementById('loading-overlay').classList[v?'add':'remove']('active'); }

// ── SELECCIÓN TRÁMITE ─────────────────────────────────────────
function selectTramite(tipo) {
    if (!userState.isLoggedIn) { signInWithMicrosoft(); return; }
    appState.currentTramite = tipo;
    loadForm(tipo);
    showSection('form');
}

// ── GENERACIÓN DE FORMULARIO ──────────────────────────────────
function loadForm(tipo) {
    const cfg = FORMS_CONFIG[tipo]; if (!cfg) return;
    document.getElementById('form-title').textContent    = cfg.title;
    document.getElementById('form-subtitle').textContent = cfg.subtitle||'Complete todos los campos marcados con *';
    const form = document.getElementById('dynamic-form');
    form.innerHTML = '';
    cfg.fields.forEach(f => {
        const g = document.createElement('div'); g.className = 'form-group';
        const lbl = document.createElement('label');
        lbl.className = 'form-label'+(f.required?' required':'');
        lbl.textContent = f.label; lbl.setAttribute('for', f.name);
        g.appendChild(lbl);
        let inp;
        if (f.type==='select') {
            inp = document.createElement('select'); inp.className='form-select';
            inp.appendChild(Object.assign(document.createElement('option'),{value:'',textContent:'Seleccione...'}));
            (f.options||[]).forEach(o => inp.appendChild(Object.assign(document.createElement('option'),{value:o,textContent:o})));
        } else if (f.type==='textarea') {
            inp = document.createElement('textarea'); inp.className='form-textarea'; inp.rows=f.rows||4; inp.placeholder=f.placeholder||'';
        } else {
            inp = document.createElement('input'); inp.type=f.type; inp.className='form-input'; inp.placeholder=f.placeholder||'';
        }
        inp.id=f.name; inp.name=f.name; if (f.required) inp.required=true;
        g.appendChild(inp);
        if (f.help) g.appendChild(Object.assign(document.createElement('div'),{className:'form-help-text',textContent:f.help}));
        form.appendChild(g);
    });
    // Pre-llenar con datos del usuario M365
    const emailF  = form.querySelector('[name="correo"],[name="correo_solicitante"]');
    const nombreF = form.querySelector('[name="nombre_completo"],[name="nombre_estudiante"],[name="nombre_solicitante"]');
    if (emailF  && userState.profile) emailF.value  = userState.profile.email;
    if (nombreF && userState.profile) nombreF.value = userState.profile.nombre;
}

// ── ENVIAR FORMULARIO ─────────────────────────────────────────
async function submitForm() {
    if (!userState.isLoggedIn) { signInWithMicrosoft(); return; }
    const form = document.getElementById('dynamic-form');
    if (!form.checkValidity()) { form.reportValidity(); return; }
    appState.formData = {};
    new FormData(form).forEach((v,k) => appState.formData[k]=v);
    appState.folio = generateFolio(appState.currentTramite);
    showLoading(true);
    try {
        // 1. PDF
        const pdfBlob = generatePDF(appState.currentTramite, appState.formData, appState.folio);
        // 2. Subir PDF a SharePoint
        const pdfUrl  = await uploadPDFToSharePoint(pdfBlob, appState.folio, appState.currentTramite);
        appState.lastPdfUrl = pdfUrl;
        // 3. Crear item en lista SharePoint
        await createSolicitudEnSharePoint({ folio:appState.folio, tipo:FORMS_CONFIG[appState.currentTramite].title, formData:appState.formData });
        // 4. Email al usuario
        if (CONFIG.email.enviarConfirmacion) {
            const dest = appState.formData.correo||appState.formData.correo_solicitante||userState.profile.email;
            await sendEmailViaGraph(dest, `CITRO — Solicitud recibida (Folio: ${appState.folio})`,
                buildConfirmationEmailHTML(appState.formData, appState.folio, appState.currentTramite, pdfUrl));
        }
        // 5. Notificación al CT
        await sendEmailViaGraph(CONFIG.email.adminEmail, `CITRO — Nueva solicitud: ${appState.folio}`,
            buildAdminNotificationHTML(appState.formData, appState.folio, appState.currentTramite, pdfUrl));
        // 6. Power Automate (opcional)
        if (CONFIG.powerAutomate.flowUrl) {
            fetch(CONFIG.powerAutomate.flowUrl, { method:'POST', headers:{'Content-Type':'application/json'},
                body:JSON.stringify({ folio:appState.folio, tipo:FORMS_CONFIG[appState.currentTramite].title,
                    solicitante:appState.formData.nombre_completo||'', email:userState.profile.email, pdfUrl:pdfUrl||'' })
            }).catch(()=>{});
        }
        showSuccess();
    } catch(e) {
        console.error(e);
        alert('Error al enviar la solicitud:\n'+e.message+'\n\nRevise la consola (F12) para más detalles.');
    } finally { showLoading(false); }
}

// ── PDF (jsPDF) ───────────────────────────────────────────────
function generatePDF(tipo, fd, folio) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF(); const m=20; let y=m;
    const cfg = FORMS_CONFIG[tipo];
    doc.setFontSize(11); doc.setFont(undefined,'bold');
    doc.text('CENTRO DE INVESTIGACIONES TROPICALES (CITRO)',105,y,{align:'center'}); y+=6;
    doc.setFontSize(9); doc.setFont(undefined,'normal');
    doc.text('Universidad Veracruzana',105,y,{align:'center'}); y+=7;
    doc.setDrawColor(0,120,212); doc.setLineWidth(0.7);
    doc.line(m,y,210-m,y); y+=7;
    doc.setTextColor(0,120,212); doc.setFontSize(8); doc.setFont(undefined,'bold');
    doc.text(`Folio: ${folio}`,105,y,{align:'center'}); y+=9;
    doc.setTextColor(0,0,0); doc.setFontSize(13); doc.setFont(undefined,'bold');
    doc.splitTextToSize((cfg.title||'').toUpperCase(),170).forEach(l=>{doc.text(l,105,y,{align:'center'});y+=7;});
    const mes=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
    const hoy=new Date();
    const fStr=`Xalapa, Ver., ${hoy.getDate()} de ${mes[hoy.getMonth()]} de ${hoy.getFullYear()}`;
    y+=4; doc.setFontSize(9); doc.setFont(undefined,'normal'); doc.text(fStr,190,y,{align:'right'}); y+=12;
    ['H. Consejo Técnico','Centro de Investigaciones Tropicales (CITRO)','Universidad Veracruzana'].forEach(l=>{doc.text(l,m,y);y+=5;});
    doc.setFont(undefined,'bold'); doc.text('Presente.',m,y); doc.setFont(undefined,'normal'); y+=12;
    const nombre = fd.nombre_completo||fd.nombre_estudiante||fd.nombre_solicitante||'[Nombre]';
    const mat    = fd.matricula||'[Matrícula]';
    doc.splitTextToSize(`Por medio del presente, ${nombre}, con número de identificación ${mat}, me dirijo respetuosamente a este H. Consejo Técnico para solicitar el apoyo correspondiente según los datos del documento con folio ${folio}.`,170)
        .forEach(l=>{doc.text(l,m,y);y+=5.5;}); y+=8;
    doc.setFontSize(10); doc.setFont(undefined,'bold'); doc.text('RESUMEN DE LA SOLICITUD',m,y); y+=7;
    doc.setFontSize(8); doc.setFont(undefined,'normal');
    let ri=0;
    Object.entries(fd).forEach(([k,v])=>{
        if (!v) return; if (y>252){doc.addPage();y=m;}
        if (ri%2===0){doc.setFillColor(235,243,251);doc.rect(m,y-3.5,170,6,'F');}
        doc.setFont(undefined,'bold'); doc.text(k.replace(/_/g,' ').toUpperCase(),m+1,y);
        doc.setFont(undefined,'normal'); doc.text(String(v).substring(0,80),m+65,y);
        y+=6; ri++;
    });
    y+=14; if(y>242){doc.addPage();y=m;}
    doc.text('Atentamente,',m,y); y+=20;
    doc.line(m,y,m+60,y); y+=5;
    doc.setFont(undefined,'bold'); doc.text(nombre,m,y); y+=5;
    doc.setFont(undefined,'normal'); doc.setFontSize(7); doc.text(mat,m,y);
    const pages=doc.internal.getNumberOfPages();
    for(let i=1;i<=pages;i++){
        doc.setPage(i); doc.setFontSize(7); doc.setTextColor(120,120,120);
        doc.text(`Folio: ${folio}  ·  ${fStr}  ·  CITRO / Universidad Veracruzana  ·  Microsoft 365`,105,287,{align:'center'});
    }
    return doc.output('blob');
}

// ── ÉXITO ─────────────────────────────────────────────────────
function showSuccess() {
    document.getElementById('success-folio').textContent = appState.folio;
    document.getElementById('success-date').textContent  = new Date().toLocaleDateString('es-MX');
    document.getElementById('success-type').textContent  = FORMS_CONFIG[appState.currentTramite]?.title||'';
    showSection('success');
}
function openSharePoint() { window.open(appState.lastPdfUrl||CONFIG.sharepoint.siteUrl,'_blank'); }

// ── UTILIDADES ────────────────────────────────────────────────
function generateFolio(tipo) {
    const p={apoyo_academico:'AAC',aval_institucional:'AVI',apoyo_terceros:'TER',comite_tutorial:'CMT',solicitud_libre:'LIB'};
    const n=new Date();
    return `${p[tipo]||'DOC'}-${n.getFullYear()}${String(n.getMonth()+1).padStart(2,'0')}${String(n.getDate()).padStart(2,'0')}-${String(n.getHours()).padStart(2,'0')}${String(n.getMinutes()).padStart(2,'0')}${String(n.getSeconds()).padStart(2,'0')}`;
}
function resetForm() {
    appState.currentTramite=null; appState.formData={}; appState.folio=null;
    const f=document.getElementById('dynamic-form'); if(f) f.innerHTML='';
}
function goToPanel() { if (userState.isAdmin) goToAdminPanel(); else goToMisSolicitudes(); }
