/**
 * ═══════════════════════════════════════════════════════════════
 * CITRO — Configuración para Universidad Veracruzana
 * Microsoft 365 SharePoint: CtTramites2026
 * ═══════════════════════════════════════════════════════════════
 */

const CONFIG = {
    
    // ━━━ AZURE ACTIVE DIRECTORY ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // Obtener en: https://portal.azure.com
    azure: {
        clientId: '0681dda0-70f4-4038-b901-d41ee738cc2e'  // ← App Registration → Overview
        tenantId: '3af6b5a8-d84b-4cfa-9942-c6af8638c3e8',  // ← App Registration → Overview
        
        // ⚠️ NO MODIFICAR estos permisos
        scopes: [
            'User.Read',              // Leer perfil del usuario
            'Sites.ReadWrite.All',    // Leer/escribir en SharePoint
            'Calendars.ReadWrite',    // Crear eventos en Outlook
            'Mail.Send'               // Enviar correos
        ]
    },

    // ━━━ SHAREPOINT UNIVERSIDAD VERACRUZANA ━━━━━━━━━━━━━━━━━━━
    sharepoint: {
        // URL del sitio (SIN slash al final)
        siteUrl: 'https://uvmx.sharepoint.com/sites/CtTramites2026',
        
        // Tenant de la UV
        tenant: 'uvmx',
        
        // Nombres de las listas/bibliotecas en SharePoint
        listName: 'SolicitudesCITRO',           // Lista principal
        libraryName: 'PDFs_Solicitudes',        // Biblioteca de PDFs
        
        // Carpetas organizadas por tipo de trámite
        folders: {
            apoyo_academico:     '01_Apoyo_Academico',
            aval_institucional:  '02_Aval_Institucional',
            apoyo_terceros:      '03_Apoyo_Terceros',
            comite_tutorial:     '04_Comite_Tutorial',
            solicitud_libre:     '05_Solicitud_Libre'
        }
    },

    // ━━━ POWER AUTOMATE (Opcional) ━━━━━━━━━━━━━━━━━━━━━━━━━━━
    powerAutomate: {
        // URL del flujo HTTP (dejar vacío si no se usa)
        flowUrl: '',
        
        // Habilitar notificaciones automáticas
        enabled: false
    },

    // ━━━ ADMINISTRADORES DEL SISTEMA ━━━━━━━━━━━━━━━━━━━━━━━━━
    // Emails con acceso al Panel de Administración
    admins: [
        'clopez@uv.mx',
        'rmenchaca@uv.mx',
        'carlolopezo@uv.mx'
        // Agregar más aquí...
    ],

    // ━━━ INFORMACIÓN INSTITUCIONAL ━━━━━━━━━━━━━━━━━━━━━━━━━━━
    institucion: {
        nombre:      'Centro de Investigaciones Tropicales',
        nombreCorto: 'CITRO',
        universidad: 'Universidad Veracruzana',
        email:       'ctecnicocitro@uv.mx,
        telefono:    '228-842-1800',
        direccion:   'Xalapa, Veracruz, México',
        sitioWeb:    'https://www.uv.mx/citro'
    },

    // ━━━ CONFIGURACIÓN DE CORREOS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    email: {
        // Email del Consejo Técnico (recibe todas las notificaciones)
        adminEmail: 'ctecnicocitro@uv.mx',
        
        // Enviar confirmación al usuario que envía la solicitud
        enviarConfirmacion: true,
        
        // CC al admin en el correo del usuario
        ccAdminEnConfirmacion: false,
        
        // Firma del email
        firmaEmail: 'H. Consejo Técnico del CITRO<br>Universidad Veracruzana'
    },

    // ━━━ OPCIONES DE SEGURIDAD Y VALIDACIÓN ━━━━━━━━━━━━━━━━━━
    options: {
        // Solo permitir cuentas institucionales @uv.mx
        soloEmailUV: true,
        dominioPermitido: 'uv.mx',
        
        // Plazo mínimo para solicitudes (días)
        plazoMinimoDias: 21,
        
        // Validación de montos
        montoMaximo: 100000,  // MXN
        requiereJustificacionSi: 50000,  // MXN - solicitar justificación si excede
        
        // Modo debug (mostrar logs en consola)
        debug: false,
        
        // Cache de datos (minutos)
        cacheDuration: 5
    },

    // ━━━ CONFIGURACIÓN DE FORMULARIOS ━━━━━━━━━━━━━━━━━━━━━━━━
    formularios: {
        // Campos obligatorios globales
        camposObligatorios: ['nombre_completo', 'correo', 'matricula'],
        
        // Formato de folio
        formatoFolio: {
            apoyo_academico:    'AAC',
            aval_institucional: 'AVI',
            apoyo_terceros:     'TER',
            comite_tutorial:    'CMT',
            solicitud_libre:    'LIB'
        },
        
        // Tipos de solicitante
        tiposSolicitante: [
            'Estudiante de Licenciatura',
            'Estudiante de Maestría',
            'Estudiante de Doctorado',
            'Académico',
            'Técnico Académico',
            'Personal Administrativo',
            'Externo'
        ]
    },

    // ━━━ VERSIÓN DEL SISTEMA ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    version: {
        numero: '1.0',
        fecha: 'Febrero 2026',
        nombre: 'Sistema CITRO M365 - Universidad Veracruzana'
    }
};

// ━━━ VALIDACIÓN DE CONFIGURACIÓN ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Verificar que la configuración esté completa
(function validarConfig() {
    const errores = [];
    
    if (CONFIG.azure.clientId.includes('PEGAR')) {
        errores.push('⚠️ Falta configurar Azure clientId');
    }
    
    if (CONFIG.azure.tenantId.includes('PEGAR')) {
        errores.push('⚠️ Falta configurar Azure tenantId');
    }
    
    if (!CONFIG.sharepoint.siteUrl.includes('uvmx.sharepoint.com')) {
        errores.push('⚠️ URL de SharePoint incorrecta');
    }
    
    if (CONFIG.admins.length === 0 || CONFIG.admins[0].includes('director.citro')) {
        console.warn('⚠️ Recuerda actualizar la lista de administradores en CONFIG.admins');
    }
    
    if (errores.length > 0) {
        console.error('❌ ERRORES DE CONFIGURACIÓN:');
        errores.forEach(e => console.error(e));
        console.error('→ Edita el archivo config-m365.js antes de usar el sistema');
    } else if (CONFIG.options.debug) {
        console.log('✅ Configuración validada correctamente');
    }
})();
