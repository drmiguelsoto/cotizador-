/**
 * FIRMA DE PRESUPUESTOS · Dr. Miguel Soto Universe
 * Backend Apps Script (Web App: Ejecutar como YO · Acceso: Cualquiera)
 * Cuenta propietaria: miguelsotouniverse@gmail.com
 *
 * Sheet SELLADO: solo este script escribe en él.
 * Acciones:
 *   POST crearSolicitud  (Cotizador)  → token + liga para la paciente
 *   GET  obtener?t=      (Página)     → presupuesto + términos
 *   POST firmar          (Página)     → registra evidencia, sella PDF, hash
 *   GET  estado?ids=     (Cotizador)  → estado de firma por COT
 *   POST anular          (Cotizador)  → invalida una liga pendiente
 *
 * Primera vez: ejecutar setupInicial() desde el editor.
 */

var FIRMA_BASE_URL = 'https://firma.miguelsotomd.com/?t=';
var DIAS_VIGENCIA_LIGA = 30;
var TERMS_VERSION = 'T-2026-11-v1';
var CREADORES_PERMITIDOS = /@(gmail\.com|miguelsotomd\.com|renaste\.com|zorvix\.mx)$/i;

var COLS = ['Token','Folio','COT','PacienteId','Nombre','Telefono','Email','Total','Anticipo',
  'SnapshotJSON','SnapshotHash','TermsVersion','CreadoPor','FechaCreacion','Expira','Estado',
  'Aperturas','PrimeraApertura','UltimaApertura',
  'FechaFirma','NombreEscrito','Tel4','IP','UserAgent','ZonaHoraria','Pantalla',
  'FirmaPngId','PdfId','HashDocumento','EvidenciaJSON','CorreoEnviado','AnuladaPor'];

// ═════════════════════════════════════ TÉRMINOS (fuente única)
function terminos_(anticipoTxt) {
  return [
    { k:'apartado', t:'Apartado de fecha',
      p:'La fecha de su cirugía queda reservada únicamente con el pago de un anticipo de ' + anticipoTxt +
        ', que se aplica en su totalidad al importe de este presupuesto.' },
    { k:'destino', t:'Destino del anticipo',
      p:'Al reservar su fecha se bloquean en exclusiva el espacio quirúrgico, el equipo médico y la agenda del cirujano, ' +
        'y se programan los insumos de su procedimiento. Por ello el anticipo cubre costos que no se recuperan si la fecha se libera.' },
    { k:'cancelacion', t:'Si usted cancela',
      p:'Con más de 30 días naturales de anticipación a la fecha programada se le reembolsa el 50 % del anticipo; ' +
        'entre 30 y 15 días, el 25 %; con menos de 15 días, o si no se presenta, no hay reembolso.' },
    { k:'credito', t:'Crédito en lugar de reembolso',
      p:'En cualquier caso, en lugar del reembolso usted puede conservar el 100 % de lo pagado como crédito para una nueva fecha ' +
        'dentro de los 12 meses siguientes a su pago. Ese crédito puede transferirse, mediante aviso por escrito, a un familiar o amigo ' +
        'para cualquier procedimiento del Dr. Miguel Soto dentro del mismo plazo.' },
    { k:'parciales', t:'Pagos posteriores al anticipo',
      p:'Los pagos parciales que usted realice después del anticipo siguen la misma política de cancelación y de crédito descrita arriba.' },
    { k:'medico', t:'Si la cirugía no se realiza por indicación médica',
      p:'Si el Dr. Miguel Soto determina, con base en su valoración o estudios preoperatorios, que no es conveniente operarle, ' +
        'o si el consultorio necesita reprogramar su fecha, usted podrá elegir el reembolso íntegro de lo pagado o conservarlo como crédito, ' +
        'descontando únicamente los estudios y valoraciones ya realizados.' },
    { k:'presupuesto', t:'Alcance y vigencia del presupuesto',
      p:'Este presupuesto ampara solo los procedimientos descritos y tiene vigencia durante el año en que se entrega. ' +
        'Si el plan quirúrgico cambia, el presupuesto se modificará y se le presentará nuevamente.' },
    { k:'electronica', t:'Firma electrónica',
      p:'Usted acepta que su firma electrónica en este documento, junto con los datos de identificación y la constancia de fecha y hora ' +
        'que lo acompañan, tiene la misma validez que su firma autógrafa. Sus datos se tratan conforme al aviso de privacidad del consultorio.' }
  ];
}

// ═════════════════════════════════════ SETUP
function setupInicial() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SHEET_ID')) {
    var ss = SpreadsheetApp.create('Firmas de Presupuestos (SELLADO)');
    var sh = ss.getSheets()[0];
    sh.setName('Firmas');
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    var log = ss.insertSheet('Bitacora');
    log.getRange(1,1,1,5).setValues([['Fecha','Token','Evento','Detalle','Origen']]).setFontWeight('bold');
    log.setFrozenRows(1);
    props.setProperty('SHEET_ID', ss.getId());
  }
  if (!props.getProperty('FOLDER_ID')) {
    var root = DriveApp.createFolder('Presupuestos Firmados');
    props.setProperty('FOLDER_ID', root.getId());
  }
  Logger.log('SHEET_ID=' + props.getProperty('SHEET_ID'));
  Logger.log('FOLDER_ID=' + props.getProperty('FOLDER_ID'));
  // Fuerza los permisos de correo durante la autorización
  Logger.log('Cuota de correo: ' + MailApp.getRemainingDailyQuota());
}

// ═════════════════════════════════════ ROUTER
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'obtener') return json_(obtener_(p.t, 'web'));
    if (p.action === 'estado')  return json_(estado_(p.ids || ''));
    if (p.action === 'ping')    return json_({ ok:true, version:TERMS_VERSION });
    return json_({ ok:false, error:'Acción no reconocida' });
  } catch (err) {
    return json_({ ok:false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (x) { return json_({ ok:false, error:'JSON inválido' }); }
  try {
    if (body.action === 'crearSolicitud') return json_(crearSolicitud_(body));
    if (body.action === 'firmar')         return json_(firmar_(body));
    if (body.action === 'anular')         return json_(anular_(body));
    return json_({ ok:false, error:'Acción no reconocida' });
  } catch (err) {
    return json_({ ok:false, error: String(err && err.message || err) });
  }
}

// ═════════════════════════════════════ ACCIONES
function crearSolicitud_(b) {
  var q = b.cotizacion || {};
  var creador = String(b.createdBy || '');
  if (!CREADORES_PERMITIDOS.test(creador)) throw new Error('Usuario no autorizado para generar ligas de firma');
  if (!q.id || !q.nombre) throw new Error('Falta folio o nombre de la cotización');
  var tel = soloDigitos_(q.tel);
  if (tel.length < 10) throw new Error('La cotización no tiene un teléfono válido de la paciente');

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = hoja_();
    var data = sh.getDataRange().getValues();
    var c = idx_();
    var ahora = new Date();
    // Si ya hay una firmada para este COT, no se genera otra
    for (var i = 1; i < data.length; i++) {
      if (data[i][c.COT] === q.id && data[i][c.Estado] === 'firmada') {
        return { ok:true, yaFirmada:true, folio:data[i][c.Folio], fechaFirma: iso_(data[i][c.FechaFirma]) };
      }
    }
    // Una liga pendiente vigente se reutiliza (salvo que el presupuesto cambió)
    var snapshot = snapshot_(q);
    var snapJson = JSON.stringify(snapshot);
    var snapHash = sha256_(snapJson);
    for (var j = 1; j < data.length; j++) {
      var r = data[j];
      if (r[c.COT] === q.id && r[c.Estado] === 'pendiente') {
        if (r[c.SnapshotHash] === snapHash && new Date(r[c.Expira]) > ahora) {
          return { ok:true, token:r[c.Token], url:FIRMA_BASE_URL + r[c.Token], folio:r[c.Folio], reutilizada:true };
        }
        sh.getRange(j + 1, c.Estado + 1).setValue('reemplazada');
        log_(r[c.Token], 'reemplazada', 'Presupuesto modificado o liga vencida', creador);
      }
    }
    var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    var folio = 'FIR-' + String(data.length).padStart(4, '0');
    var expira = new Date(ahora.getTime() + DIAS_VIGENCIA_LIGA * 86400000);
    var row = new Array(COLS.length).fill('');
    row[c.Token] = token; row[c.Folio] = folio; row[c.COT] = q.id; row[c.PacienteId] = q.pacienteId || '';
    row[c.Nombre] = q.nombre; row[c.Telefono] = tel; row[c.Email] = q.email || '';
    row[c.Total] = snapshot.total; row[c.Anticipo] = snapshot.anticipo;
    row[c.SnapshotJSON] = snapJson; row[c.SnapshotHash] = snapHash; row[c.TermsVersion] = TERMS_VERSION;
    row[c.CreadoPor] = creador; row[c.FechaCreacion] = ahora; row[c.Expira] = expira;
    row[c.Estado] = 'pendiente'; row[c.Aperturas] = 0;
    sh.appendRow(row);
    log_(token, 'creada', folio + ' · ' + q.id, creador);
    return { ok:true, token:token, url:FIRMA_BASE_URL + token, folio:folio };
  } finally { lock.releaseLock(); }
}

function obtener_(token, origen) {
  var f = buscar_(token);
  if (!f) return { ok:false, error:'Esta liga no es válida. Por favor solicite una nueva a su asesora.' };
  var r = f.row, c = idx_();
  var est = r[c.Estado];
  if (est === 'firmada') {
    return { ok:true, estado:'firmada', folio:r[c.Folio], nombre:r[c.Nombre], cot:r[c.COT],
      fechaFirma: iso_(r[c.FechaFirma]), hash:r[c.HashDocumento] };
  }
  if (est !== 'pendiente') return { ok:false, error:'Esta liga ya no está vigente. Por favor solicite una nueva a su asesora.' };
  if (new Date(r[c.Expira]) < new Date()) return { ok:false, error:'Esta liga venció. Por favor solicite una nueva a su asesora.' };

  var sh = hoja_(), ahora = new Date();
  sh.getRange(f.rowNum, c.Aperturas + 1).setValue(Number(r[c.Aperturas] || 0) + 1);
  if (!r[c.PrimeraApertura]) sh.getRange(f.rowNum, c.PrimeraApertura + 1).setValue(ahora);
  sh.getRange(f.rowNum, c.UltimaApertura + 1).setValue(ahora);
  log_(token, 'apertura', '', origen);

  var snap = JSON.parse(r[c.SnapshotJSON]);
  return { ok:true, estado:'pendiente', folio:r[c.Folio], cot:r[c.COT], presupuesto:snap,
    terminos: terminos_(mxn_(snap.anticipo)), termsVersion:r[c.TermsVersion] };
}

function firmar_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var f = buscar_(b.t);
    if (!f) throw new Error('Liga no válida');
    var r = f.row, c = idx_(), sh = hoja_();
    if (r[c.Estado] === 'firmada') return { ok:true, yaFirmada:true, folio:r[c.Folio], hash:r[c.HashDocumento] };
    if (r[c.Estado] !== 'pendiente' || new Date(r[c.Expira]) < new Date()) throw new Error('Esta liga ya no está vigente');

    var snap = JSON.parse(r[c.SnapshotJSON]);
    var tel = String(r[c.Telefono]);
    var nombreEscrito = String(b.nombreEscrito || '').trim();
    var tel4 = soloDigitos_(b.tel4);
    var acepta = b.aceptaciones || [];
    var keys = terminos_('').map(function (x) { return x.k; });

    if (nombreEscrito.split(/\s+/).length < 2) throw new Error('Escriba su nombre completo');
    if (tel4 !== tel.slice(-4)) {
      log_(b.t, 'tel4_incorrecto', tel4, 'web');
      throw new Error('Los últimos 4 dígitos no coinciden con el teléfono registrado. Verifíquelos o comuníquese al 55 6223 4133.');
    }
    if (norm_(nombreEscrito) !== norm_(r[c.Nombre]) && !similar_(nombreEscrito, r[c.Nombre])) {
      throw new Error('El nombre no coincide con el registrado en su presupuesto (' + r[c.Nombre] + '). Escríbalo tal como aparece.');
    }
    for (var i = 0; i < keys.length; i++) if (acepta.indexOf(keys[i]) < 0) throw new Error('Debe aceptar cada uno de los términos');
    if (!b.firmaPng || String(b.firmaPng).indexOf('data:image/png;base64,') !== 0) throw new Error('Falta su firma');
    if (String(b.firmaPng).length > 400000) throw new Error('La imagen de la firma es demasiado grande');

    var ahora = new Date();
    var folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
    var base = r[c.Folio] + ' · ' + r[c.COT] + ' · ' + r[c.Nombre];
    var png = Utilities.newBlob(Utilities.base64Decode(String(b.firmaPng).split(',')[1]), 'image/png', base + ' · firma.png');
    var pngFile = folder.createFile(png);

    var evidencia = {
      folio:r[c.Folio], cot:r[c.COT], pacienteId:r[c.PacienteId], nombreRegistrado:r[c.Nombre],
      nombreEscrito:nombreEscrito, telefonoRegistrado:'••••••' + tel.slice(-4), tel4Verificado:true,
      ligaEnviadaA:'WhatsApp ' + tel, fechaHoraServidor:ahora.toISOString(),
      zonaHoraria:String(b.tz || ''), ip:String(b.ip || ''), userAgent:String(b.ua || '').slice(0, 300),
      pantalla:String(b.pantalla || ''), termsVersion:r[c.TermsVersion], aceptaciones:acepta,
      snapshotHash:r[c.SnapshotHash], firmaPngSha256:sha256Bytes_(png.getBytes()),
      aperturas:Number(r[c.Aperturas] || 0), primeraApertura:iso_(r[c.PrimeraApertura])
    };
    var canon = JSON.stringify({ presupuesto:snap, terminos:terminos_(mxn_(snap.anticipo)), evidencia:evidencia });
    var hashDoc = sha256_(canon);

    var html = pdfHtml_(snap, terminos_(mxn_(snap.anticipo)), evidencia, hashDoc, String(b.firmaPng), r[c.Folio]);
    var pdf = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(base + ' · ACEPTADA Y FIRMADA.pdf');
    var pdfFile = folder.createFile(pdf);
    // Copia del JSON canónico: permite recalcular el hash en cualquier momento
    folder.createFile(Utilities.newBlob(canon, 'application/json', base + ' · evidencia.json'));

    var correo = '';
    if (r[c.Email] && /@/.test(r[c.Email])) {
      try {
        MailApp.sendEmail({ to:r[c.Email], name:'Dr. Miguel Soto',
          subject:'Su presupuesto firmado · ' + r[c.COT],
          htmlBody:'<p>Estimada(o) ' + esc_(r[c.Nombre]) + ':</p><p>Le compartimos su presupuesto <b>' + r[c.COT] +
            '</b> con los términos de anticipo que usted aceptó y firmó el ' + fechaMx_(ahora) + '.</p>' +
            '<p>Para cualquier duda, comuníquese con nosotros al <b>55 6223 4133</b>.</p><p>Dr. Miguel Soto</p>',
          attachments:[pdf] });
        correo = r[c.Email];
      } catch (mx) { correo = 'error: ' + mx.message; }
    }

    var upd = {};
    upd[c.Estado] = 'firmada'; upd[c.FechaFirma] = ahora; upd[c.NombreEscrito] = nombreEscrito; upd[c.Tel4] = 'OK';
    upd[c.IP] = evidencia.ip; upd[c.UserAgent] = evidencia.userAgent; upd[c.ZonaHoraria] = evidencia.zonaHoraria;
    upd[c.Pantalla] = evidencia.pantalla; upd[c.FirmaPngId] = pngFile.getId(); upd[c.PdfId] = pdfFile.getId();
    upd[c.HashDocumento] = hashDoc; upd[c.EvidenciaJSON] = JSON.stringify(evidencia); upd[c.CorreoEnviado] = correo;
    Object.keys(upd).forEach(function (k) { sh.getRange(f.rowNum, Number(k) + 1).setValue(upd[k]); });
    log_(b.t, 'firmada', hashDoc, evidencia.ip);

    return { ok:true, folio:r[c.Folio], hash:hashDoc, fechaFirma:ahora.toISOString(),
      correo: correo && correo.indexOf('error') !== 0 ? correo : '',
      pdfBase64: Utilities.base64Encode(pdf.getBytes()), pdfNombre: pdf.getName() };
  } finally { lock.releaseLock(); }
}

function estado_(ids) {
  var lista = String(ids).split(',').map(function (s) { return s.trim(); }).filter(String);
  var out = {};
  var data = hoja_().getDataRange().getValues(), c = idx_();
  for (var i = 1; i < data.length; i++) {
    var r = data[i], id = r[c.COT];
    if (lista.length && lista.indexOf(id) < 0) continue;
    var est = r[c.Estado];
    if (est === 'reemplazada' || est === 'anulada') continue;
    if (out[id] && out[id].estado === 'firmada') continue;
    out[id] = { estado:est, folio:r[c.Folio], fechaFirma:iso_(r[c.FechaFirma]), aperturas:Number(r[c.Aperturas] || 0),
      url: est === 'pendiente' ? FIRMA_BASE_URL + r[c.Token] : '' };
  }
  return { ok:true, estados:out };
}

function anular_(b) {
  if (!CREADORES_PERMITIDOS.test(String(b.createdBy || ''))) throw new Error('No autorizado');
  var data = hoja_().getDataRange().getValues(), c = idx_(), n = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][c.COT] === b.cot && data[i][c.Estado] === 'pendiente') {
      hoja_().getRange(i + 1, c.Estado + 1).setValue('anulada');
      hoja_().getRange(i + 1, c.AnuladaPor + 1).setValue(b.createdBy);
      log_(data[i][c.Token], 'anulada', '', b.createdBy); n++;
    }
  }
  return { ok:true, anuladas:n };
}

// ═════════════════════════════════════ SNAPSHOT DEL PRESUPUESTO
// Congela lo que la paciente ve y firma; si el presupuesto cambia, cambia el hash.
function snapshot_(q) {
  var lineas = [];
  (q.selected || []).forEach(function (s, i) {
    var n = (s.proc && s.proc.n) || 'Procedimiento';
    lineas.push({ concepto:String(n), detalle: i === 0 ? 'Procedimiento principal' : 'Procedimiento adicional' });
  });
  (q.moduleItems || []).forEach(function (it) {
    lineas.push({ concepto:String(it.n || ''), detalle:(it.qty && it.qty > 1 ? it.qty + ' × ' : '') + (it.cat || '') });
  });
  (q.extras || []).forEach(function (x) { lineas.push({ concepto:String(x.n || 'Concepto adicional'), detalle:'' }); });
  return {
    cot:String(q.id), fechaPresupuesto:String(q.fecha || ''), nombre:String(q.nombre),
    lineas:lineas, total:Number(q.total) || 0, anticipo:Number(q.anticipo) || 10000,
    pdfPresupuesto:String(q.pdfLink || '')
  };
}

// ═════════════════════════════════════ PDF
function pdfHtml_(snap, terms, ev, hash, firmaPng, folio) {
  var L = snap.lineas.map(function (l) {
    return '<tr><td>' + esc_(l.concepto) + '</td><td class="m">' + esc_(l.detalle) + '</td></tr>'; }).join('');
  var T = terms.map(function (t, i) {
    return '<div class="t"><b>' + (i + 1) + '. ' + esc_(t.t) + '.</b> ' + esc_(t.p) + ' <span class="ok">✔ Aceptado</span></div>'; }).join('');
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Helvetica,Arial,sans-serif;color:#030707;font-size:9.5px;line-height:1.4;margin:18px 28px}' +
    '.hdr{color:#030707;padding:0 0 8px;border-bottom:2px solid #030707;letter-spacing:5px;font-size:13px}' +
    '.hdr small{display:block;color:#8A7B6C;font-size:7.5px;letter-spacing:3px;margin-top:3px}' +
    'h2{font-family:Georgia,serif;font-weight:normal;font-size:16px;margin:12px 0 2px}' +
    '.lbl{font-size:7px;letter-spacing:2.5px;text-transform:uppercase;color:#6B6059;margin:11px 0 5px;font-weight:bold}' +
    'table{width:100%;border-collapse:collapse}td{padding:5px 0;border-bottom:1px solid #E9E6E1;vertical-align:top}.m{color:#6B6059;text-align:right}' +
    '.tot{display:flex;justify-content:space-between;font-size:13px;margin-top:8px}.t{margin:0 0 7px}.ok{color:#3a7d44;font-size:9px;white-space:nowrap}' +
    '.firma{border:1px solid #D8D2CB;padding:10px;margin-top:6px}.firma img{height:70px}' +
    '.ev{font-family:Courier,monospace;font-size:8.5px;background:#F5F2EE;padding:8px;word-break:break-all;line-height:1.45}' +
    '</style></head><body>' +
    '<div class="hdr">DR. MIGUEL SOTO<small>ACEPTACIÓN DE PRESUPUESTO · ' + esc_(folio) + '</small></div>' +
    '<h2>' + esc_(snap.nombre) + '</h2><div style="color:#6B6059">Presupuesto ' + esc_(snap.cot) +
    (snap.fechaPresupuesto ? ' · emitido el ' + esc_(snap.fechaPresupuesto) : '') + '</div>' +
    '<div class="lbl">Procedimientos y conceptos</div><table>' + L + '</table>' +
    '<table style="margin-top:6px"><tr><td><b>Total del presupuesto</b></td><td class="m" style="color:#030707"><b>' + mxn_(snap.total) + '</b></td></tr>' +
    '<tr><td>Anticipo para apartar fecha</td><td class="m" style="color:#030707">' + mxn_(snap.anticipo) + '</td></tr></table>' +
    '<div class="lbl">Términos aceptados (' + esc_(ev.termsVersion) + ')</div>' + T +
    '<div class="lbl">Firma</div><div class="firma"><img src="' + firmaPng + '"><br>' + esc_(ev.nombreEscrito) +
    '<br><span style="color:#6B6059">Firmado electrónicamente el ' + fechaMx_(new Date(ev.fechaHoraServidor)) + ' (hora de la Ciudad de México)</span></div>' +
    '<div class="lbl">Constancia de evidencia</div><div class="ev">' +
    'Liga personal enviada a: ' + esc_(ev.ligaEnviadaA.replace(/\d(?=\d{4})/g, '•')) + '<br>' +
    'Verificación: nombre completo + últimos 4 dígitos del teléfono registrado ✔<br>' +
    'Fecha y hora (servidor, UTC): ' + esc_(ev.fechaHoraServidor) + '<br>' +
    'Zona horaria del dispositivo: ' + esc_(ev.zonaHoraria) + ' · IP: ' + esc_(ev.ip || 'no disponible') + '<br>' +
    'Dispositivo: ' + esc_(ev.userAgent) + '<br>' +
    'Huella del presupuesto (SHA-256): ' + esc_(ev.snapshotHash) + '<br>' +
    'Huella de la firma (SHA-256): ' + esc_(ev.firmaPngSha256) + '<br>' +
    '<b>Huella del documento (SHA-256): ' + esc_(hash) + '</b></div>' +
    '<div style="margin-top:6px;color:#6B6059;font-size:7.5px">Dr. Miguel Ángel Soto Miranda · Miramontes 2342, Coyoacán, C.P. 04460, CDMX · 55 6223 4133</div>' +
    '</body></html>';
}

// ═════════════════════════════════════ UTILIDADES
function hoja_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('Firmas'); }
function idx_() { var o = {}; COLS.forEach(function (k, i) { o[k] = i; }); return o; }
function buscar_(token) {
  if (!token || !/^[a-f0-9]{64}$/.test(String(token))) return null;
  var data = hoja_().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) if (data[i][0] === token) return { row:data[i], rowNum:i + 1 };
  return null;
}
function log_(token, evento, detalle, origen) {
  try { SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'))
    .getSheetByName('Bitacora').appendRow([new Date(), token, evento, detalle, origen]); } catch (e) {}
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sha256_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }
function sha256Bytes_(b) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, b)); }
function hex_(bytes) { return bytes.map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join(''); }
function soloDigitos_(s) { return String(s || '').replace(/\D/g, ''); }
function norm_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-zñ ]/g, '').replace(/\s+/g, ' ').trim(); }
// Tolera orden distinto o un segundo nombre/apellido faltante: exige que coincidan al menos 2 palabras
function similar_(a, b) {
  var A = norm_(a).split(' '), B = norm_(b).split(' ');
  var hits = A.filter(function (w) { return w.length > 1 && B.indexOf(w) >= 0; }).length;
  return hits >= Math.min(2, B.length);
}
function mxn_(n) { return '$' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits:0, maximumFractionDigits:2 }) + ' MXN'; }
function iso_(d) { return d ? new Date(d).toISOString() : ''; }
function fechaMx_(d) { return Utilities.formatDate(d, 'America/Mexico_City', "d 'de' MMMM 'de' yyyy, HH:mm 'h'").replace(/January|February|March|April|May|June|July|August|September|October|November|December/, function (m) {
  return { January:'enero', February:'febrero', March:'marzo', April:'abril', May:'mayo', June:'junio', July:'julio', August:'agosto', September:'septiembre', October:'octubre', November:'noviembre', December:'diciembre' }[m]; }); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (ch) { return { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[ch]; }); }
