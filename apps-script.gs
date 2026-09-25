/**
 * Receptor de inscripciones — Programa de Transformación Digital, Camacol Caldas.
 *
 * Qué hace: recibe el envío del formulario, escribe una fila en la hoja "Inscripciones",
 * manda la confirmación a quien se inscribe y un aviso interno al equipo.
 *
 * Instalación: ver README.md
 */

// ───────────────────── Configuración ─────────────────────
const HOJA            = 'Inscripciones';
const AVISO_INTERNO_A = 'sergio.franco@camacolcaldas.org';   // separar varios con coma
const NOMBRE_REMITENTE= 'Camacol Capacita';
const CORREO_CONTACTO = 'sergio.franco@camacolcaldas.org';
const CELULAR_CONTACTO= '[CELULAR]';

// Orden de las columnas. Debe coincidir con los name= del formulario.
const COLUMNAS = [
  'fecha_envio','razon_social','nit','ciudad','direccion','afiliada',
  'responsable_nombre','responsable_cargo','responsable_correo','responsable_celular',
  'lider_nombre','lider_cargo','lider_correo','lider_celular','lider_rol',
  'proyecto_nombre','proyecto_tipo','proyecto_etapa','proyecto_area','proyecto_unidades','tiene_modelo',
  'madurez_bim','control_obra','software','expectativa',
  'factura_razon','factura_nit','factura_correo','factura_contacto',
  'acepta_dedicacion','acepta_valor','acepta_datos','origen'
];

const ENCABEZADOS = [
  'Fecha','Razón social','NIT','Ciudad','Dirección','Afiliada',
  'Responsable','Cargo responsable','Correo responsable','Celular responsable',
  'Líder','Cargo líder','Correo líder','Celular líder','Rol líder',
  'Proyecto','Tipo','Etapa','Área m2','Unidades','Tiene modelo BIM',
  'Madurez BIM','Control de obra','Software','Expectativa',
  'Razón social factura','NIT factura','Correo factura','Contacto factura',
  'Acepta dedicación','Acepta valor','Autoriza datos','Origen'
];

// ───────────────────── Recepción ─────────────────────
function doPost(e) {
  try {
    const p = e.parameter || {};

    // Antispam: si el campo oculto viene lleno, es un bot.
    if (p.website) return json({ ok: true });

    const hoja = obtenerHoja_();
    hoja.appendRow(COLUMNAS.map(c => p[c] || ''));

    enviarConfirmacion_(p);
    enviarAvisoInterno_(p);

    return json({ ok: true });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, mensaje: 'Receptor de inscripciones activo.' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function obtenerHoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA);
    hoja.appendRow(ENCABEZADOS);
    hoja.getRange(1, 1, 1, ENCABEZADOS.length)
        .setFontWeight('bold').setBackground('#164194').setFontColor('#FFFFFF');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// ───────────────────── Correos ─────────────────────
function enviarConfirmacion_(p) {
  const para = p.responsable_correo;
  if (!para) return;

  const asunto = 'Inscripción recibida · Programa de Transformación Digital';
  const cuerpo =
    'Buenos días, ' + (p.responsable_nombre || '') + '.\n\n' +
    'Recibimos la inscripción de ' + (p.razon_social || 'su empresa') + ' al programa ' +
    '"De la preconstrucción virtual a la gestión industrializada de proyectos".\n\n' +
    'Lo que sigue:\n\n' +
    '1. En los próximos días le enviaremos la factura a ' + (p.factura_correo || p.responsable_correo) + '.\n' +
    '2. Antes del inicio debe confirmarnos los otros dos participantes de su equipo. ' +
    'Recomendamos cubrir gerencia, programación y presupuestos, y dirección de obra.\n' +
    '3. El programa inicia el 13 de octubre. Las sesiones virtuales son martes y jueves ' +
    'de 6:30 a 8:30 p. m., y las dos jornadas presenciales en Manizales son el 10 y el 12 de noviembre.\n\n' +
    'Datos que registramos:\n' +
    '· Empresa: ' + (p.razon_social || '') + ' — NIT ' + (p.nit || '') + '\n' +
    '· Participante líder: ' + (p.lider_nombre || '') + ' (' + (p.lider_cargo || '') + ')\n' +
    '· Proyecto piloto: ' + (p.proyecto_nombre || '') + '\n\n' +
    'Si algo quedó mal registrado, respóndanos este correo y lo corregimos.\n\n' +
    'Cordialmente,\n\n' +
    'Camacol Capacita · Cámara Regional de la Construcción — Camacol Caldas\n' +
    CORREO_CONTACTO + ' · ' + CELULAR_CONTACTO;

  MailApp.sendEmail({ to: para, subject: asunto, body: cuerpo, name: NOMBRE_REMITENTE });
}

function enviarAvisoInterno_(p) {
  const asunto = 'Nueva inscripción: ' + (p.razon_social || 'sin razón social');
  const cuerpo =
    'Empresa: ' + (p.razon_social || '') + ' (NIT ' + (p.nit || '') + ') — afiliada: ' + (p.afiliada || '') + '\n' +
    'Inscribe: ' + (p.responsable_nombre || '') + ', ' + (p.responsable_cargo || '') + ' · ' +
    (p.responsable_correo || '') + ' · ' + (p.responsable_celular || '') + '\n' +
    'Líder: ' + (p.lider_nombre || '') + ', ' + (p.lider_cargo || '') + ' — ' + (p.lider_rol || '') + '\n' +
    'Proyecto: ' + (p.proyecto_nombre || '') + ' · ' + (p.proyecto_tipo || '') + ' · ' + (p.proyecto_etapa || '') + '\n' +
    'Modelo BIM: ' + (p.tiene_modelo || '') + ' · Madurez: ' + (p.madurez_bim || '') + '\n' +
    'Control de obra: ' + (p.control_obra || '') + '\n' +
    'Software: ' + (p.software || '') + '\n' +
    'Expectativa: ' + (p.expectativa || '') + '\n' +
    'Facturar a: ' + (p.factura_razon || p.razon_social || '') + ' — ' + (p.factura_correo || '') + '\n';

  MailApp.sendEmail({ to: AVISO_INTERNO_A, subject: asunto, body: cuerpo, name: NOMBRE_REMITENTE });
}
