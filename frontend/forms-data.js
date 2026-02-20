/**
 * ═══════════════════════════════════════════════════════════════
 * CITRO - Configuración de Formularios V2.0
 * Incluye 5 tipos de formularios
 * ═══════════════════════════════════════════════════════════════
 */

const FORMS_CONFIG = {
    apoyo_academico: {
        title: 'Solicitud de Apoyo a Actividades Académicas',
        subtitle: 'Para viáticos, congresos, estancias y trabajo de campo',
        fields: [
            {
                name: 'nombre_completo',
                label: 'Nombre completo del solicitante',
                type: 'text',
                required: true,
                placeholder: 'Ej: Ana María García López'
            },
            {
                name: 'correo',
                label: 'Correo electrónico',
                type: 'email',
                required: true,
                placeholder: 'ejemplo@uv.mx'
            },
            {
                name: 'tipo_solicitante',
                label: 'Tipo de solicitante',
                type: 'select',
                required: true,
                options: [
                    'Estudiante de Maestría',
                    'Estudiante de Doctorado',
                    'Académico/Investigador',
                    'Técnico Académico',
                    'Otro'
                ]
            },
            {
                name: 'matricula',
                label: 'Matrícula o número de personal',
                type: 'text',
                required: true,
                placeholder: 'Ej: S20012345'
            },
            {
                name: 'tipo_actividad',
                label: 'Tipo de actividad',
                type: 'select',
                required: true,
                options: [
                    'Asistencia a evento académico (congreso, simposio, etc.)',
                    'Presentación de trabajo en evento',
                    'Curso o taller',
                    'Trabajo de campo',
                    'Estancia académica o de investigación',
                    'Otra'
                ]
            },
            {
                name: 'titulo_actividad',
                label: 'Título de la actividad',
                type: 'text',
                required: true,
                placeholder: 'Nombre del evento, curso, estancia, etc.'
            },
            {
                name: 'fecha_inicio',
                label: 'Fecha de inicio',
                type: 'date',
                required: true
            },
            {
                name: 'fecha_termino',
                label: 'Fecha de término',
                type: 'date',
                required: true
            },
            {
                name: 'destino',
                label: 'Lugar/destino',
                type: 'text',
                required: true,
                placeholder: 'Ciudad, Estado, País'
            },
            {
                name: 'justificacion',
                label: 'Justificación de la solicitud',
                type: 'textarea',
                required: true,
                rows: 5,
                placeholder: 'Explique por qué es importante esta actividad para su formación o investigación'
            },
            {
                name: 'proyecto_sirei',
                label: 'Proyecto SiREI asociado (si aplica)',
                type: 'text',
                required: false,
                placeholder: 'Número de proyecto'
            },
            {
                name: 'desglose_gastos',
                label: 'Desglose de gastos',
                type: 'textarea',
                required: true,
                rows: 4,
                placeholder: 'Ej: Transporte: $5,000 MXN\nHospedaje (3 noches): $3,000 MXN\nInscripción: $2,000 MXN',
                help: 'Incluya concepto y monto estimado de cada gasto'
            },
            {
                name: 'monto_total',
                label: 'Monto total solicitado (MXN)',
                type: 'number',
                required: true,
                placeholder: '10000'
            }
        ]
    },
    
    aval_institucional: {
        title: 'Solicitud de Aval Institucional',
        subtitle: 'Para representar al CITRO / UV en eventos académicos',
        fields: [
            {
                name: 'nombre_completo',
                label: 'Nombre completo del solicitante',
                type: 'text',
                required: true
            },
            {
                name: 'correo',
                label: 'Correo electrónico',
                type: 'email',
                required: true
            },
            {
                name: 'tipo_solicitante',
                label: 'Tipo de solicitante',
                type: 'select',
                required: true,
                options: [
                    'Estudiante de Maestría',
                    'Estudiante de Doctorado',
                    'Académico/Investigador',
                    'Egresado',
                    'Otro'
                ]
            },
            {
                name: 'matricula',
                label: 'Matrícula o número de personal',
                type: 'text',
                required: true
            },
            {
                name: 'tipo_participacion',
                label: 'Tipo de participación',
                type: 'select',
                required: true,
                options: [
                    'Ponencia oral',
                    'Presentación de póster',
                    'Organización de evento',
                    'Impartición de clase/taller',
                    'Moderador/Comentarista',
                    'Otra'
                ]
            },
            {
                name: 'nombre_evento',
                label: 'Nombre del evento',
                type: 'text',
                required: true
            },
            {
                name: 'fecha_inicio',
                label: 'Fecha de inicio',
                type: 'date',
                required: true
            },
            {
                name: 'fecha_termino',
                label: 'Fecha de término',
                type: 'date',
                required: true
            },
            {
                name: 'descripcion',
                label: 'Descripción de la participación',
                type: 'textarea',
                required: true,
                rows: 4,
                placeholder: 'Título de la ponencia, taller, o descripción de la actividad'
            },
            {
                name: 'relevancia',
                label: 'Relevancia para el CITRO',
                type: 'textarea',
                required: true,
                rows: 4,
                placeholder: 'Explique cómo esta participación beneficia al centro'
            },
            {
                name: 'lgac',
                label: 'LGAC relacionada',
                type: 'select',
                required: false,
                options: [
                    'Biodiversidad y Conservación',
                    'Ecología Funcional',
                    'Cambio Global',
                    'No aplica'
                ]
            }
        ]
    },
    
    apoyo_terceros: {
        title: 'Solicitud de Apoyo a Terceros — CITRO',
        subtitle: 'Para invitar ponentes o colaboradores externos',
        fields: [
            {
                name: 'nombre_solicitante',
                label: 'Nombre del solicitante (quien solicita el apoyo)',
                type: 'text',
                required: true
            },
            {
                name: 'correo_solicitante',
                label: 'Correo del solicitante',
                type: 'email',
                required: true
            },
            {
                name: 'matricula',
                label: 'Matrícula o número de personal',
                type: 'text',
                required: true
            },
            {
                name: 'nombre_tercero',
                label: 'Nombre completo del tercero (persona a apoyar)',
                type: 'text',
                required: true
            },
            {
                name: 'correo_tercero',
                label: 'Correo del tercero',
                type: 'email',
                required: true
            },
            {
                name: 'institucion_tercero',
                label: 'Institución del tercero',
                type: 'text',
                required: true
            },
            {
                name: 'relacion',
                label: 'Relación con el solicitante',
                type: 'select',
                required: true,
                options: [
                    'Director de tesis',
                    'Co-tutor',
                    'Asesor externo',
                    'Colaborador de investigación',
                    'Ponente invitado',
                    'Otra'
                ]
            },
            {
                name: 'tipo_actividad',
                label: 'Tipo de actividad',
                type: 'select',
                required: true,
                options: [
                    'Ponencia/Conferencia',
                    'Curso/Taller',
                    'Asesoría de investigación',
                    'Colaboración en proyecto',
                    'Visita académica',
                    'Otra'
                ]
            },
            {
                name: 'titulo_actividad',
                label: 'Título de la actividad',
                type: 'text',
                required: true
            },
            {
                name: 'fecha_inicio',
                label: 'Fecha de inicio',
                type: 'date',
                required: true
            },
            {
                name: 'fecha_termino',
                label: 'Fecha de término',
                type: 'date',
                required: true
            },
            {
                name: 'descripcion',
                label: 'Descripción de la actividad',
                type: 'textarea',
                required: true,
                rows: 4
            },
            {
                name: 'objetivo',
                label: 'Objetivo y pertinencia',
                type: 'textarea',
                required: true,
                rows: 4,
                placeholder: 'Explique por qué es importante traer a esta persona'
            },
            {
                name: 'tipos_apoyo',
                label: 'Tipos de apoyo solicitado',
                type: 'textarea',
                required: true,
                rows: 3,
                placeholder: 'Ej: Transporte, hospedaje, honorarios, alimentación'
            },
            {
                name: 'monto_total',
                label: 'Monto total estimado (MXN)',
                type: 'number',
                required: true
            },
            {
                name: 'desglose',
                label: 'Desglose del presupuesto',
                type: 'textarea',
                required: true,
                rows: 4,
                placeholder: 'Detalle de cada concepto y monto'
            }
        ]
    },
    
    comite_tutorial: {
        title: 'Solicitud de Comité Tutorial',
        subtitle: 'Maestría/Doctorado en Ecología Tropical',
        fields: [
            {
                name: 'nombre_estudiante',
                label: 'Nombre completo del estudiante',
                type: 'text',
                required: true
            },
            {
                name: 'correo',
                label: 'Correo electrónico',
                type: 'email',
                required: true
            },
            {
                name: 'programa',
                label: 'Programa de posgrado',
                type: 'select',
                required: true,
                options: [
                    'Maestría en Ecología Tropical',
                    'Doctorado en Ecología Tropical'
                ]
            },
            {
                name: 'matricula',
                label: 'Matrícula',
                type: 'text',
                required: true
            },
            {
                name: 'tema_tesis',
                label: 'Tema/título de tesis',
                type: 'textarea',
                required: true,
                rows: 3
            },
            {
                name: 'director_nombre',
                label: 'Nombre del Director de tesis',
                type: 'text',
                required: true
            },
            {
                name: 'director_institucion',
                label: 'Institución del Director',
                type: 'text',
                required: true
            },
            {
                name: 'cotutor_nombre',
                label: 'Nombre del Co-tutor (si aplica)',
                type: 'text',
                required: false
            },
            {
                name: 'cotutor_institucion',
                label: 'Institución del Co-tutor',
                type: 'text',
                required: false
            },
            {
                name: 'asesor1_nombre',
                label: 'Nombre del Asesor 1',
                type: 'text',
                required: true
            },
            {
                name: 'asesor1_institucion',
                label: 'Institución del Asesor 1',
                type: 'text',
                required: true
            },
            {
                name: 'asesor2_nombre',
                label: 'Nombre del Asesor 2 (si aplica)',
                type: 'text',
                required: false
            },
            {
                name: 'asesor2_institucion',
                label: 'Institución del Asesor 2',
                type: 'text',
                required: false
            },
            {
                name: 'asesor3_nombre',
                label: 'Nombre del Asesor 3 (si aplica)',
                type: 'text',
                required: false
            },
            {
                name: 'asesor3_institucion',
                label: 'Institución del Asesor 3',
                type: 'text',
                required: false
            },
            {
                name: 'cambio_miembro',
                label: '¿Se trata de cambio de algún miembro del comité?',
                type: 'select',
                required: true,
                options: ['No', 'Sí']
            },
            {
                name: 'motivo_cambio',
                label: 'Si es cambio, especifique el motivo',
                type: 'textarea',
                required: false,
                rows: 3
            }
        ]
    },
    
    // ═══════════════════════════════════════════════════════════════
    // NUEVO: SOLICITUD LIBRE
    // ═══════════════════════════════════════════════════════════════
    
    solicitud_libre: {
        title: 'Solicitud Libre al Consejo Técnico',
        subtitle: 'Para trámites no contemplados en los formatos anteriores',
        fields: [
            {
                name: 'nombre_completo',
                label: 'Nombre completo del solicitante',
                type: 'text',
                required: true,
                placeholder: 'Ej: Ana María García López'
            },
            {
                name: 'correo',
                label: 'Correo electrónico',
                type: 'email',
                required: true,
                placeholder: 'ejemplo@uv.mx'
            },
            {
                name: 'tipo_solicitante',
                label: 'Tipo de solicitante',
                type: 'select',
                required: true,
                options: [
                    'Estudiante de Maestría',
                    'Estudiante de Doctorado',
                    'Académico/Investigador',
                    'Técnico Académico',
                    'Personal Administrativo',
                    'Egresado',
                    'Externo',
                    'Otro'
                ]
            },
            {
                name: 'matricula',
                label: 'Matrícula o número de personal (si aplica)',
                type: 'text',
                required: false,
                placeholder: 'Ej: S20012345'
            },
            {
                name: 'asunto',
                label: 'Asunto de la solicitud',
                type: 'text',
                required: true,
                placeholder: 'Resuma en pocas palabras el motivo de su solicitud',
                help: 'Sea específico y conciso (máximo 100 caracteres)'
            },
            {
                name: 'descripcion',
                label: 'Descripción detallada',
                type: 'textarea',
                required: true,
                rows: 8,
                placeholder: 'Explique con detalle su solicitud:\n\n- ¿Qué está solicitando?\n- ¿Por qué lo necesita?\n- ¿Cuándo lo necesita?\n- ¿Cómo beneficia al CITRO/UV?\n- Cualquier información relevante adicional',
                help: 'Sea lo más específico posible para que el Consejo Técnico pueda evaluar adecuadamente su solicitud'
            },
            {
                name: 'fecha_relacionada',
                label: 'Fecha relacionada con la solicitud (si aplica)',
                type: 'date',
                required: false,
                help: 'Fecha del evento, actividad, plazo, etc.'
            },
            {
                name: 'requiere_presupuesto',
                label: '¿La solicitud requiere apoyo económico?',
                type: 'select',
                required: true,
                options: ['No', 'Sí']
            },
            {
                name: 'monto_total',
                label: 'Monto solicitado (MXN)',
                type: 'number',
                required: false,
                placeholder: '0',
                help: 'Solo si seleccionó "Sí" en la pregunta anterior'
            },
            {
                name: 'desglose_presupuesto',
                label: 'Desglose del presupuesto',
                type: 'textarea',
                required: false,
                rows: 4,
                placeholder: 'Si requiere apoyo económico, detalle los gastos:\n\nConcepto 1: $XXXX\nConcepto 2: $XXXX\n...',
                help: 'Solo si seleccionó "Sí" en apoyo económico'
            },
            {
                name: 'documentos_adjuntos',
                label: 'Documentos de respaldo',
                type: 'textarea',
                required: false,
                rows: 3,
                placeholder: 'Liste los documentos que adjuntará por separado:\n- Cotización\n- Carta de invitación\n- Programa del evento\netc.',
                help: 'Indique qué documentos proporcionará como evidencia o respaldo (adjuntarlos posteriormente por correo)'
            },
            {
                name: 'observaciones',
                label: 'Observaciones adicionales',
                type: 'textarea',
                required: false,
                rows: 3,
                placeholder: 'Cualquier información adicional que considere relevante'
            }
        ]
    }
};
