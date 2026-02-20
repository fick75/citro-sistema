/**
 * ═══════════════════════════════════════════════════════════════
 * CITRO — Integración SharePoint (OPTIMIZADO)
 * Microsoft Graph API + SharePoint REST API
 * Site: https://uvmx.sharepoint.com/sites/CtTramites2026
 * ═══════════════════════════════════════════════════════════════
 */

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
let _siteId = null;
let _siteIdExpiry = null;
const CACHE_DURATION = 30 * 60 * 1000; // 30 minutos

/**
 * Llamada a Microsoft Graph API con reintentos
 */
async function callGraph(endpoint, method = 'GET', body = null, retries = 3) {
    for (let attempt = 1; attempt <= retries; attempt++) {
        try {
            const token = await getAccessToken();
            
            const options = {
                method,
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            };

            if (body) {
                options.body = JSON.stringify(body);
            }

            const response = await fetch(`${GRAPH_BASE}${endpoint}`, options);

            // Manejar respuestas sin contenido
            if (response.status === 204) {
                return null;
            }

            // Manejar errores HTTP
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                const errorMsg = errorData.error?.message || `Error ${response.status}`;
                
                // Si es 401, el token expiró
                if (response.status === 401 && attempt < retries) {
                    if (CONFIG.options.debug) {
                        console.log(`⚠️ Token expirado, reintentando (${attempt}/${retries})...`);
                    }
                    await new Promise(resolve => setTimeout(resolve, 1000 * attempt));
                    continue;
                }
                
                throw new Error(errorMsg);
            }

            return response.json();

        } catch (error) {
            if (attempt === retries) {
                console.error(`❌ Error en Graph API después de ${retries} intentos:`, error);
                throw error;
            }
            
            if (CONFIG.options.debug) {
                console.log(`⚠️ Reintentando (${attempt}/${retries})...`);
            }
            
            await new Promise(resolve => setTimeout(resolve, 1000 * attempt));
        }
    }
}

/**
 * Obtener ID del sitio SharePoint (con cache)
 */
async function getSiteId() {
    // Verificar cache
    if (_siteId && _siteIdExpiry && Date.now() < _siteIdExpiry) {
        return _siteId;
    }

    try {
        const siteUrl = new URL(CONFIG.sharepoint.siteUrl);
        const hostname = siteUrl.hostname;
        const pathname = siteUrl.pathname;

        const data = await callGraph(`/sites/${hostname}:${pathname}`);
        
        _siteId = data.id;
        _siteIdExpiry = Date.now() + CACHE_DURATION;

        if (CONFIG.options.debug) {
            console.log('✅ Site ID obtenido:', _siteId);
        }

        return _siteId;

    } catch (error) {
        console.error('Error al obtener Site ID:', error);
        throw new Error(`No se pudo conectar con SharePoint: ${error.message}`);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * CREAR SOLICITUD EN SHAREPOINT
 * ═══════════════════════════════════════════════════════════════
 */
async function createSolicitudEnSharePoint(solicitudData) {
    try {
        const siteId = await getSiteId();
        const formData = solicitudData.formData;

        // Preparar campos para SharePoint
        const fields = {
            Title: solicitudData.folio,
            Folio: solicitudData.folio,
            TipoTramite: solicitudData.tipo,
            NombreSolicitante: formData.nombre_completo || 
                              formData.nombre_estudiante || 
                              formData.nombre_solicitante || 
                              userState.profile.nombre,
            EmailSolicitante: formData.correo || 
                             formData.correo_solicitante || 
                             userState.profile.email,
            EmailUsuarioM365: userState.profile.email,
            Matricula: formData.matricula || '',
            MontoSolicitado: parseFloat(formData.monto_total || 0),
            MontoAutorizado: 0,
            Estado: 'Pendiente',
            DatosCompletos: JSON.stringify(formData),
            FechaSolicitud: new Date().toISOString(),
            NotasCT: '',
            URLPdf: ''
        };

        const response = await callGraph(
            `/sites/${siteId}/lists/${CONFIG.sharepoint.listName}/items`,
            'POST',
            { fields }
        );

        if (CONFIG.options.debug) {
            console.log('✅ Solicitud creada en SharePoint:', response.id);
        }

        return response;

    } catch (error) {
        console.error('Error al crear solicitud en SharePoint:', error);
        throw new Error(`No se pudo guardar en SharePoint: ${error.message}`);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * SUBIR PDF A SHAREPOINT
 * ═══════════════════════════════════════════════════════════════
 */
async function uploadPDFToSharePoint(pdfBlob, folio, tipoTramite) {
    try {
        const siteId = await getSiteId();
        
        // Obtener carpeta según tipo de trámite
        const carpeta = CONFIG.sharepoint.folders[tipoTramite] || '06_Otros';
        const fileName = `${folio}.pdf`;

        // Convertir blob a array buffer
        const arrayBuffer = await pdfBlob.arrayBuffer();
        const token = await getAccessToken();

        // Subir usando Graph API
        const uploadUrl = `${GRAPH_BASE}/sites/${siteId}/drive/root:/${CONFIG.sharepoint.libraryName}/${carpeta}/${fileName}:/content`;

        const response = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/pdf'
            },
            body: arrayBuffer
        });

        if (!response.ok) {
            throw new Error(`Error al subir PDF: ${response.status}`);
        }

        const data = await response.json();
        const pdfUrl = data.webUrl;

        if (CONFIG.options.debug) {
            console.log('✅ PDF subido a SharePoint:', pdfUrl);
        }

        return pdfUrl;

    } catch (error) {
        console.error('Error al subir PDF:', error);
        throw new Error(`No se pudo subir el PDF: ${error.message}`);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * ACTUALIZAR SOLICITUD (Solo Admin)
 * ═══════════════════════════════════════════════════════════════
 */
async function updateSolicitudEnSharePoint(itemId, updates) {
    try {
        if (!userState.isAdmin) {
            throw new Error('No tienes permisos para editar solicitudes');
        }

        const siteId = await getSiteId();

        // Preparar campos a actualizar
        const fields = {};
        
        if (updates.estado !== undefined) {
            fields.Estado = updates.estado;
        }
        
        if (updates.montoAutorizado !== undefined) {
            fields.MontoAutorizado = parseFloat(updates.montoAutorizado || 0);
        }
        
        if (updates.notasCT !== undefined) {
            fields.NotasCT = updates.notasCT;
        }

        await callGraph(
            `/sites/${siteId}/lists/${CONFIG.sharepoint.listName}/items/${itemId}/fields`,
            'PATCH',
            fields
        );

        if (CONFIG.options.debug) {
            console.log('✅ Solicitud actualizada:', itemId);
        }

    } catch (error) {
        console.error('Error al actualizar solicitud:', error);
        throw new Error(`No se pudo actualizar: ${error.message}`);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * OBTENER SOLICITUDES DEL USUARIO ACTUAL
 * ═══════════════════════════════════════════════════════════════
 */
async function getSolicitudesUsuario() {
    try {
        const siteId = await getSiteId();
        const email = userState.profile.email;

        const query = `/sites/${siteId}/lists/${CONFIG.sharepoint.listName}/items?` +
            `$filter=fields/EmailUsuarioM365 eq '${email}'` +
            `&$select=id,fields` +
            `&$expand=fields` +
            `&$orderby=fields/FechaSolicitud desc` +
            `&$top=100`;

        const data = await callGraph(query);

        const solicitudes = (data.value || []).map(item => ({
            spId: item.id,
            ...item.fields
        }));

        if (CONFIG.options.debug) {
            console.log(`✅ Solicitudes obtenidas: ${solicitudes.length}`);
        }

        return solicitudes;

    } catch (error) {
        console.error('Error al obtener solicitudes del usuario:', error);
        throw new Error(`No se pudieron cargar tus solicitudes: ${error.message}`);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * OBTENER TODAS LAS SOLICITUDES (Solo Admin)
 * ═══════════════════════════════════════════════════════════════
 */
async function getAllSolicitudes() {
    try {
        if (!userState.isAdmin) {
            throw new Error('No tienes permisos para ver todas las solicitudes');
        }

        const siteId = await getSiteId();

        const query = `/sites/${siteId}/lists/${CONFIG.sharepoint.listName}/items?` +
            `$select=id,fields` +
            `&$expand=fields` +
            `&$orderby=fields/FechaSolicitud desc` +
            `&$top=500`;

        const data = await callGraph(query);

        const solicitudes = (data.value || []).map(item => ({
            spId: item.id,
            ...item.fields
        }));

        if (CONFIG.options.debug) {
            console.log(`✅ Todas las solicitudes obtenidas: ${solicitudes.length}`);
        }

        return solicitudes;

    } catch (error) {
        console.error('Error al obtener todas las solicitudes:', error);
        throw new Error(`No se pudieron cargar las solicitudes: ${error.message}`);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * ENVIAR EMAIL VÍA MICROSOFT GRAPH
 * ═══════════════════════════════════════════════════════════════
 */
async function sendEmailViaGraph(to, subject, htmlBody, ccEmails = []) {
    try {
        const emailData = {
            message: {
                subject: subject,
                body: {
                    contentType: 'HTML',
                    content: htmlBody
                },
                toRecipients: [
                    { emailAddress: { address: to } }
                ],
                importance: 'Normal'
            },
            saveToSentItems: true
        };

        // Agregar CC si hay
        if (ccEmails.length > 0) {
            emailData.message.ccRecipients = ccEmails.map(email => ({
                emailAddress: { address: email }
            }));
        }

        await callGraph('/me/sendMail', 'POST', emailData);

        if (CONFIG.options.debug) {
            console.log('✅ Email enviado a:', to);
        }

    } catch (error) {
        console.error('Error al enviar email:', error);
        // No lanzar error para no bloquear el flujo principal
        console.warn('⚠️ El email no se pudo enviar, pero la solicitud se guardó correctamente');
    }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * TEMPLATES DE EMAILS
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Email de confirmación al usuario
 */
function buildConfirmationEmailHTML(formData, folio, tipoTramite, pdfUrl) {
    const nombre = formData.nombre_completo || 
                  formData.nombre_estudiante || 
                  formData.nombre_solicitante || 
                  userState.profile.nombre;
                  
    const tipoNombre = FORMS_CONFIG[tipoTramite]?.title || tipoTramite;
    const fecha = new Date().toLocaleDateString('es-MX', { dateStyle: 'full' });
    const monto = formData.monto_total ? 
        `$${parseFloat(formData.monto_total).toLocaleString('es-MX')} MXN` : null;

    return `
    <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;border:1px solid #EDEBE9;border-radius:8px;overflow:hidden">
        <div style="background:#0078D4;padding:28px;text-align:center">
            <h1 style="color:#fff;margin:0;font-size:22px;font-weight:600">✅ Solicitud Recibida</h1>
            <p style="color:#c7e3ff;margin:8px 0 0;font-size:14px">${CONFIG.institucion.nombre}</p>
            <p style="color:#c7e3ff;margin:4px 0 0;font-size:13px">${CONFIG.institucion.universidad}</p>
        </div>
        
        <div style="padding:28px">
            <p style="font-size:16px;color:#333;margin:0 0 12px">Estimado/a <strong>${nombre}</strong>,</p>
            
            <p style="color:#605E5C;line-height:1.7;margin:12px 0">
                Su solicitud ha sido registrada exitosamente en el sistema del ${CONFIG.institucion.nombreCorto}. 
                El <strong>H. Consejo Técnico</strong> revisará su solicitud en las próximas sesiones.
            </p>
            
            <div style="background:#EBF3FB;border-left:4px solid #0078D4;padding:20px;border-radius:0 6px 6px 0;margin:24px 0">
                <h3 style="margin:0 0 14px;color:#0078D4;font-size:16px">📋 Detalles de la Solicitud</h3>
                <table style="width:100%;border-collapse:collapse;font-size:14px">
                    <tr>
                        <td style="padding:6px 0;font-weight:600;color:#605E5C;width:40%">Folio:</td>
                        <td style="padding:6px 0;color:#0078D4;font-weight:700;font-family:monospace">${folio}</td>
                    </tr>
                    <tr>
                        <td style="padding:6px 0;font-weight:600;color:#605E5C">Tipo de Trámite:</td>
                        <td style="padding:6px 0">${tipoNombre}</td>
                    </tr>
                    <tr>
                        <td style="padding:6px 0;font-weight:600;color:#605E5C">Fecha de Solicitud:</td>
                        <td style="padding:6px 0">${fecha}</td>
                    </tr>
                    ${monto ? `<tr>
                        <td style="padding:6px 0;font-weight:600;color:#605E5C">Monto Solicitado:</td>
                        <td style="padding:6px 0;font-weight:700;color:#107C10">${monto}</td>
                    </tr>` : ''}
                    <tr>
                        <td style="padding:6px 0;font-weight:600;color:#605E5C">Estado Actual:</td>
                        <td style="padding:6px 0">
                            <span style="background:#FFF4CE;color:#7A4F01;padding:3px 12px;border-radius:12px;font-size:12px;font-weight:600">
                                ⏳ Pendiente de Revisión
                            </span>
                        </td>
                    </tr>
                </table>
            </div>
            
            ${pdfUrl ? `
            <div style="text-align:center;margin:24px 0">
                <a href="${pdfUrl}" style="display:inline-block;background:#0078D4;color:#fff;padding:12px 28px;text-decoration:none;border-radius:6px;font-weight:600;font-size:14px">
                    📄 Ver Documento en SharePoint
                </a>
            </div>` : ''}
            
            <div style="background:#FAF9F8;border:1px solid #EDEBE9;border-radius:6px;padding:16px;margin:24px 0">
                <p style="margin:0;font-size:13px;color:#605E5C;line-height:1.6">
                    <strong>📌 Próximos pasos:</strong><br>
                    1. Recibirá una notificación cuando el Consejo Técnico revise su solicitud<br>
                    2. Puede consultar el estado en cualquier momento en "Mis Solicitudes"<br>
                    3. El CT puede solicitar información adicional si es necesario
                </p>
            </div>
            
            <p style="color:#605E5C;font-size:13px;line-height:1.7;margin:20px 0 0">
                Si tiene alguna duda, puede contactar al ${CONFIG.institucion.nombreCorto}:<br>
                📧 ${CONFIG.institucion.email}<br>
                ${CONFIG.institucion.telefono ? `📞 ${CONFIG.institucion.telefono}<br>` : ''}
                🌐 ${CONFIG.institucion.sitioWeb}
            </p>
        </div>
        
        <div style="background:#FAF9F8;padding:20px;text-align:center;border-top:1px solid #EDEBE9">
            <p style="margin:0;font-size:12px;color:#888;line-height:1.5">
                <strong>${CONFIG.institucion.nombre}</strong><br>
                ${CONFIG.institucion.universidad}<br>
                ${CONFIG.institucion.direccion}<br><br>
                <em>Este es un mensaje automático generado desde Microsoft SharePoint. Por favor, no responder directamente a este correo.</em>
            </p>
        </div>
    </div>`;
}

/**
 * Email de notificación al Consejo Técnico
 */
function buildAdminNotificationHTML(formData, folio, tipoTramite, pdfUrl) {
    const nombre = formData.nombre_completo || 
                  formData.nombre_estudiante || 
                  formData.nombre_solicitante || '—';
    const email = formData.correo || formData.correo_solicitante || userState.profile.email;
    const monto = formData.monto_total ? 
        `$${parseFloat(formData.monto_total).toLocaleString('es-MX')} MXN` : 'Sin monto especificado';
    const tipoNombre = FORMS_CONFIG[tipoTramite]?.title || tipoTramite;
    const tipoSolicitante = formData.tipo_solicitante || '—';

    // Construir resumen de datos clave
    let datosRelevantes = '';
    if (formData.titulo_actividad) {
        datosRelevantes += `<tr><td style="padding:9px 12px;font-weight:600;background:#F3F2F1">Actividad:</td><td style="padding:9px 12px;background:#F3F2F1">${formData.titulo_actividad}</td></tr>`;
    }
    if (formData.destino) {
        datosRelevantes += `<tr><td style="padding:9px 12px;font-weight:600">Destino:</td><td style="padding:9px 12px">${formData.destino}</td></tr>`;
    }
    if (formData.fecha_inicio) {
        datosRelevantes += `<tr><td style="padding:9px 12px;font-weight:600;background:#F3F2F1">Fecha Inicio:</td><td style="padding:9px 12px;background:#F3F2F1">${formData.fecha_inicio}</td></tr>`;
    }

    return `
    <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:620px;margin:0 auto;background:#fff;border:1px solid #EDEBE9;border-radius:8px;overflow:hidden">
        <div style="background:#107C10;padding:24px">
            <h1 style="color:#fff;margin:0;font-size:20px;font-weight:600">📬 Nueva Solicitud — ${CONFIG.institucion.nombreCorto}</h1>
            <p style="color:#DFF6DD;margin:6px 0 0;font-size:13px">Sistema de Gestión Académica Digital</p>
        </div>
        
        <div style="padding:24px">
            <div style="background:#EBF3FB;border:1px solid #B3D4F0;border-radius:6px;padding:16px;margin:0 0 20px">
                <p style="margin:0;font-size:14px;color:#106EBE">
                    <strong>⚡ Acción requerida:</strong> Nueva solicitud registrada en SharePoint y pendiente de revisión por el H. Consejo Técnico.
                </p>
            </div>
            
            <table style="width:100%;border-collapse:collapse;font-size:14px;border:1px solid #EDEBE9;border-radius:6px;overflow:hidden">
                <tr style="background:#0078D4">
                    <td colspan="2" style="padding:10px 12px;color:#fff;font-weight:600">Información General</td>
                </tr>
                <tr>
                    <td style="padding:9px 12px;font-weight:600;background:#F3F2F1;width:35%">Folio:</td>
                    <td style="padding:9px 12px;background:#F3F2F1;font-weight:700;color:#0078D4;font-family:monospace">${folio}</td>
                </tr>
                <tr>
                    <td style="padding:9px 12px;font-weight:600">Tipo de Trámite:</td>
                    <td style="padding:9px 12px;font-weight:600;color:#107C10">${tipoNombre}</td>
                </tr>
                <tr>
                    <td style="padding:9px 12px;font-weight:600;background:#F3F2F1">Solicitante:</td>
                    <td style="padding:9px 12px;background:#F3F2F1">${nombre}</td>
                </tr>
                <tr>
                    <td style="padding:9px 12px;font-weight:600">Tipo:</td>
                    <td style="padding:9px 12px">${tipoSolicitante}</td>
                </tr>
                <tr>
                    <td style="padding:9px 12px;font-weight:600;background:#F3F2F1">Email:</td>
                    <td style="padding:9px 12px;background:#F3F2F1"><a href="mailto:${email}" style="color:#0078D4;text-decoration:none">${email}</a></td>
                </tr>
                <tr>
                    <td style="padding:9px 12px;font-weight:600">Monto Solicitado:</td>
                    <td style="padding:9px 12px;font-size:15px;font-weight:700;color:#107C10">${monto}</td>
                </tr>
                <tr>
                    <td style="padding:9px 12px;font-weight:600;background:#F3F2F1">Fecha/Hora:</td>
                    <td style="padding:9px 12px;background:#F3F2F1">${new Date().toLocaleString('es-MX', {dateStyle:'full', timeStyle:'short'})}</td>
                </tr>
                ${datosRelevantes}
            </table>
            
            ${pdfUrl ? `
            <div style="margin:24px 0">
                <a href="${pdfUrl}" style="display:inline-block;background:#0078D4;color:#fff;padding:11px 24px;text-decoration:none;border-radius:4px;font-weight:600;font-size:14px;margin-right:10px">
                    📄 Ver PDF en SharePoint
                </a>
                <a href="${CONFIG.sharepoint.siteUrl}" style="display:inline-block;background:#fff;color:#0078D4;border:2px solid #0078D4;padding:9px 24px;text-decoration:none;border-radius:4px;font-weight:600;font-size:14px">
                    🗂️ Abrir Lista SharePoint
                </a>
            </div>` : ''}
            
            <details style="margin-top:20px;border:1px solid #EDEBE9;border-radius:6px;overflow:hidden">
                <summary style="cursor:pointer;padding:12px 16px;background:#FAF9F8;color:#0078D4;font-weight:600;font-size:13px;user-select:none">
                    📊 Ver datos completos del formulario
                </summary>
                <pre style="margin:0;padding:16px;background:#F3F2F1;font-size:11px;overflow:auto;max-height:300px;line-height:1.5">${JSON.stringify(formData, null, 2)}</pre>
            </details>
        </div>
        
        <div style="background:#FAF9F8;padding:16px;text-align:center;border-top:1px solid #EDEBE9">
            <p style="margin:0;font-size:11px;color:#888;line-height:1.5">
                ${CONFIG.institucion.nombre} · ${CONFIG.institucion.universidad}<br>
                Mensaje automático del Sistema de Gestión Académica Digital · Microsoft SharePoint
            </p>
        </div>
    </div>`;
}
