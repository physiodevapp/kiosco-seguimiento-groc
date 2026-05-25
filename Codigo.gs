// ============================================================
//  KIOSKO SEGUIMIENTOS — Código.gs
//  Fisioterapia · Apps Script + Google Sheets
// ============================================================

// ── IDs y constantes ────────────────────────────────────────
const SHEET_ID        = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
const HOJA_FISIOS     = 'Fisios';
const HOJA_SEGUIM     = 'Seguimientos';
const DIAS_SEGUIM     = 10;

// Normaliza el número de colegiado eliminando ceros a la izquierda para comparación.
// Permite que el valor en Sheets sea número (11312) o texto ('011312') indistintamente.
function normColegiado(v) {
  return String(v).trim().replace(/^0+(\d)/, '$1');
}

// ── Menú personalizado en el Sheet ──────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠 Tests')
    .addItem('Enviar correo de prueba', 'testPlantillaRecordatorio')
    .addToUi();
}

// ── Punto de entrada web ────────────────────────────────────
function doGet(e) {
  const params = e.parameter;

  // Enlace de cierre desde email: ?accion=cerrar&id=ROW
  if (params.accion === 'cerrar' && params.id) {
    return cerrarSeguimientoDesdeEmail(params.id);
  }

  // Enlace de confirmación de reevaluación: ?accion=reevaluar&id=ROW
  if (params.accion === 'reevaluar' && params.id) {
    return reevaluarDesdeEmail(params.id);
  }

  // Kiosko normal
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Seguimientos · Fisioterapia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── API llamada desde el HTML (google.script.run) ───────────

/**
 * Valida el número de colegiado.
 * Devuelve { ok: true, nombre } o { ok: false }
 */
function validarColegiado(numColegiado) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const hoja  = ss.getSheetByName(HOJA_FISIOS);
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (normColegiado(datos[i][1]) === normColegiado(numColegiado)) {
      return { ok: true, nombre: datos[i][0], email: datos[i][2], fila: i + 1 };
    }
  }
  return { ok: false };
}

/**
 * Registra un nuevo seguimiento.
 * Devuelve { ok: true } o { ok: false, error }
 */
function registrarSeguimiento(numColegiado, fechaVisita, horaVisita) {
  try {
    const fisio = validarColegiado(numColegiado);
    if (!fisio.ok) return { ok: false, error: 'Colegiado no encontrado' };

    const ss   = SpreadsheetApp.openById(SHEET_ID);
    const hoja = ss.getSheetByName(HOJA_SEGUIM);

    // Comprobar duplicado: mismo fisio, misma fecha y hora
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (
        normColegiado(datos[i][2]) === normColegiado(numColegiado) &&
        datos[i][4] === fechaVisita &&
        datos[i][5] === horaVisita &&
        datos[i][6] === 'activo'
      ) {
        return { ok: false, error: 'Ya existe un seguimiento activo para esa franja horaria' };
      }
    }

    const ahora          = new Date();
    const proximaReev    = new Date(ahora);
    proximaReev.setDate(proximaReev.getDate() + DIAS_SEGUIM);
    const id             = 'SEG-' + ahora.getTime();

    const nextRow = hoja.getLastRow() + 1;
    hoja.getRange(nextRow, 3).setNumberFormat('@'); // preserve colegiado as text
    hoja.getRange(nextRow, 1, 1, 9).setValues([[
      id,
      fisio.nombre,
      numColegiado,
      fisio.email,
      fechaVisita,
      horaVisita,
      'activo',
      Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(proximaReev, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    ]]);

    return { ok: true, nombre: fisio.nombre };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Devuelve los seguimientos activos de un fisio.
 */
function obtenerSeguimientosFisio(numColegiado) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const hoja  = ss.getSheetByName(HOJA_SEGUIM);
  const datos = hoja.getDataRange().getValues();
  const lista = [];

  for (let i = 1; i < datos.length; i++) {
    if (
      normColegiado(datos[i][2]) === normColegiado(numColegiado) &&
      datos[i][6] === 'activo'
    ) {
      const fmt = (v, pattern) =>
        v instanceof Date
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), pattern)
          : String(v);
      lista.push({
        id:          String(datos[i][0]),
        fila:        i + 1,
        fecha:       fmt(datos[i][4], 'dd/MM/yyyy'),
        hora:        fmt(datos[i][5], 'HH:mm'),
        registro:    fmt(datos[i][7], 'dd/MM/yyyy'),
        proximaReev: fmt(datos[i][8], 'dd/MM/yyyy'),
      });
    }
  }
  return lista;
}

// ── Trigger diario de recordatorios ─────────────────────────

/**
 * Instala el trigger diario (ejecutar UNA sola vez desde el editor).
 */
function instalarTrigger() {
  // Elimina triggers previos del mismo nombre para no duplicar
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarRecordatoriosDiarios') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('enviarRecordatoriosDiarios')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

/**
 * Recorre los seguimientos activos y envía email cuando toca reevaluar.
 * También actualiza la próxima reevaluación sumando +10 días.
 */
function enviarRecordatoriosDiarios() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const hoja  = ss.getSheetByName(HOJA_SEGUIM);
  const datos = hoja.getDataRange().getValues();
  const hoy   = new Date();
  hoy.setHours(0, 0, 0, 0);

  const urlBase = ScriptApp.getService().getUrl();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][6] !== 'activo') continue;

    const proximaStr = datos[i][8]; // dd/MM/yyyy
    if (!proximaStr) continue;

    const partes  = String(proximaStr).split('/');
    const proxima = new Date(partes[2], partes[1] - 1, partes[0]);
    proxima.setHours(0, 0, 0, 0);

    if (proxima <= hoy) {
      const emailFisio  = datos[i][3];
      const nombreFisio = datos[i][1];
      const fechaVisita = datos[i][4];
      const horaVisita  = datos[i][5];
      const idSeg       = datos[i][0];
      const filaNum     = i + 1;

      const urlCerrar   = `${urlBase}?accion=cerrar&id=${filaNum}`;
      const urlReev     = `${urlBase}?accion=reevaluar&id=${filaNum}`;

      const asunto = `🔔 Reevaluación pendiente · Paciente ${fechaVisita} ${horaVisita}`;
      const cuerpo = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background:#f0f2f5;">
  <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#f0f2f5" style="background:#f0f2f5;">
    <tr>
      <td align="center" style="padding:32px 16px;">
        <table width="520" cellpadding="0" cellspacing="0" border="0" bgcolor="#f9f9f9" style="max-width:520px;background:#f9f9f9;border-radius:12px;">
          <tr>
            <td style="padding:32px 32px 0 32px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
              <h2 style="color:#1a1a2e;margin:0 0 16px 0;font-size:22px;font-weight:700;">Recordatorio de reevaluación</h2>
              <p style="color:#555555;margin:0 0 12px 0;font-size:15px;">Hola <strong>${nombreFisio}</strong>,</p>
              <p style="color:#555555;margin:0;font-size:15px;">Han pasado <strong>${DIAS_SEGUIM} días</strong> desde la primera visita de tu paciente:</p>
            </td>
          </tr>
          <tr>
            <td style="padding:20px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td bgcolor="#ffffff" style="background:#ffffff;border-left:4px solid #4a9eff;padding:16px 20px;border-radius:8px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;font-size:16px;font-weight:700;color:#1a1a2e;">
                    📅 ${fechaVisita} · ${horaVisita}
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:0 32px 16px 32px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
              <p style="color:#555555;margin:0;font-size:15px;">¿Qué quieres hacer?</p>
            </td>
          </tr>
          <tr>
            <td style="padding:0 32px 12px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td align="center" bgcolor="#4a9eff" style="border-radius:8px;background:#4a9eff;">
                    <a href="${urlReev}" target="_blank" style="display:block;padding:14px 24px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;font-size:15px;font-weight:600;color:#ffffff;text-decoration:none;text-align:center;">
                      ✅ Reevaluado — continuar seguimiento
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:0 32px 28px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td align="center" bgcolor="#ff6b6b" style="border-radius:8px;background:#ff6b6b;">
                    <a href="${urlCerrar}" target="_blank" style="display:block;padding:14px 24px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;font-size:15px;font-weight:600;color:#ffffff;text-decoration:none;text-align:center;">
                      🔴 Finalizar seguimiento
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:16px 32px 28px 32px;border-top:1px solid #e5e5e5;">
              <p style="font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;color:#aaaaaa;font-size:12px;margin:0;">Si no haces nada, recibirás otro recordatorio en ${DIAS_SEGUIM} días.</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

      MailApp.sendEmail({
        to:       emailFisio,
        subject:  asunto,
        htmlBody: cuerpo,
      });

      // Actualizar próxima reevaluación +10 días
      const nuevaProxima = new Date(proxima);
      nuevaProxima.setDate(nuevaProxima.getDate() + DIAS_SEGUIM);
      hoja.getRange(filaNum, 9).setValue(
        Utilities.formatDate(nuevaProxima, Session.getScriptTimeZone(), 'dd/MM/yyyy')
      );
    }
  }
}

// ── Test visual de plantilla de email ───────────────────────

/**
 * Envía un correo de prueba para revisar la plantilla visualmente.
 * Pide email y nº de fila mediante diálogos al ejecutar desde el editor.
 * Si se indica una fila real, carga sus datos y genera enlaces funcionales.
 * Si se deja vacío, usa datos ficticios y enlaces de ejemplo.
 */
function testPlantillaRecordatorio() {
  const ui = SpreadsheetApp.getUi();

  const resEmail = ui.prompt(
    'Test · Email destinatario',
    'Deja vacío para usar tu email (' + Session.getActiveUser().getEmail() + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resEmail.getSelectedButton() !== ui.Button.OK) return;
  const destinatario = resEmail.getResponseText().trim() || Session.getActiveUser().getEmail();

  const resFila = ui.prompt(
    'Test · Fila de Seguimientos',
    'Nº de fila en la hoja Seguimientos (deja vacío para datos ficticios):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resFila.getSelectedButton() !== ui.Button.OK) return;

  const urlDetectada = ScriptApp.getService().getUrl();
  const resUrl = ui.prompt(
    'Test · URL de la Web App',
    'URL base para los botones (deja vacío para usar la detectada: ' + urlDetectada + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resUrl.getSelectedButton() !== ui.Button.OK) return;
  const urlBase = resUrl.getResponseText().trim() || urlDetectada;

  let nombreFisio, fechaVisita, horaVisita, urlReev, urlCerrar;

  const filaNum = parseInt(resFila.getResponseText().trim(), 10);
  if (!isNaN(filaNum) && filaNum > 1) {
    const fila = SpreadsheetApp.openById(SHEET_ID)
      .getSheetByName(HOJA_SEGUIM)
      .getRange(filaNum, 1, 1, 9).getValues()[0];
    const fmt = (v, p) => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), p) : String(v);
    nombreFisio = fila[1];
    fechaVisita = fmt(fila[4], 'dd/MM/yyyy');
    horaVisita  = fmt(fila[5], 'HH:mm');
    urlReev     = `${urlBase}?accion=reevaluar&id=${filaNum}`;
    urlCerrar   = `${urlBase}?accion=cerrar&id=${filaNum}`;
  } else {
    nombreFisio = 'Ana García';
    fechaVisita = '15/05/2026';
    horaVisita  = '10:30';
    urlReev     = '#reev-test';
    urlCerrar   = '#cerrar-test';
  }

  const asunto = `🔔 [TEST] Reevaluación pendiente · Paciente ${fechaVisita} ${horaVisita}`;
  const cuerpo = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background:#f0f2f5;">
  <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#f0f2f5" style="background:#f0f2f5;">
    <tr>
      <td align="center" style="padding:32px 16px;">
        <table width="520" cellpadding="0" cellspacing="0" border="0" bgcolor="#f9f9f9" style="max-width:520px;background:#f9f9f9;border-radius:12px;">
          <tr>
            <td style="padding:32px 32px 0 32px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
              <h2 style="color:#1a1a2e;margin:0 0 16px 0;font-size:22px;font-weight:700;">Recordatorio de reevaluación</h2>
              <p style="color:#555555;margin:0 0 12px 0;font-size:15px;">Hola <strong>${nombreFisio}</strong>,</p>
              <p style="color:#555555;margin:0;font-size:15px;">Han pasado <strong>${DIAS_SEGUIM} días</strong> desde la primera visita de tu paciente:</p>
            </td>
          </tr>
          <tr>
            <td style="padding:20px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td bgcolor="#ffffff" style="background:#ffffff;border-left:4px solid #4a9eff;padding:16px 20px;border-radius:8px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;font-size:16px;font-weight:700;color:#1a1a2e;">
                    📅 ${fechaVisita} · ${horaVisita}
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:0 32px 16px 32px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
              <p style="color:#555555;margin:0;font-size:15px;">¿Qué quieres hacer?</p>
            </td>
          </tr>
          <tr>
            <td style="padding:0 32px 12px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td align="center" bgcolor="#4a9eff" style="border-radius:8px;background:#4a9eff;">
                    <a href="${urlReev}" target="_blank" style="display:block;padding:14px 24px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;font-size:15px;font-weight:600;color:#ffffff;text-decoration:none;text-align:center;">
                      ✅ Reevaluado — continuar seguimiento
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:0 32px 28px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td align="center" bgcolor="#ff6b6b" style="border-radius:8px;background:#ff6b6b;">
                    <a href="${urlCerrar}" target="_blank" style="display:block;padding:14px 24px;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;font-size:15px;font-weight:600;color:#ffffff;text-decoration:none;text-align:center;">
                      🔴 Finalizar seguimiento
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:16px 32px 28px 32px;border-top:1px solid #e5e5e5;">
              <p style="font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;color:#aaaaaa;font-size:12px;margin:0;">Si no haces nada, recibirás otro recordatorio en ${DIAS_SEGUIM} días.</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

  MailApp.sendEmail({ to: destinatario, subject: asunto, htmlBody: cuerpo });
  Logger.log(`Correo de prueba enviado a ${destinatario}`);
}

// ── Acciones desde enlaces del email ────────────────────────

function cerrarSeguimientoDesdeEmail(filaNum) {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const hoja = ss.getSheetByName(HOJA_SEGUIM);
  hoja.getRange(Number(filaNum), 7).setValue('finalizado');

  return HtmlService.createHtmlOutput(`
    <html><body style="font-family:sans-serif;text-align:center;padding:60px;background:#f9f9f9">
      <div style="max-width:400px;margin:auto;background:#fff;padding:40px;border-radius:16px;box-shadow:0 4px 20px rgba(0,0,0,0.08)">
        <div style="font-size:3em">✅</div>
        <h2 style="color:#1a1a2e">Seguimiento finalizado</h2>
        <p style="color:#555">El seguimiento ha sido cerrado correctamente.<br>Puedes cerrar esta ventana.</p>
      </div>
    </body></html>
  `);
}

function reevaluarDesdeEmail(filaNum) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const hoja  = ss.getSheetByName(HOJA_SEGUIM);
  const hoy   = new Date();
  const nueva = new Date(hoy);
  nueva.setDate(nueva.getDate() + DIAS_SEGUIM);

  hoja.getRange(Number(filaNum), 9).setValue(
    Utilities.formatDate(nueva, Session.getScriptTimeZone(), 'dd/MM/yyyy')
  );

  return HtmlService.createHtmlOutput(`
    <html><body style="font-family:sans-serif;text-align:center;padding:60px;background:#f9f9f9">
      <div style="max-width:400px;margin:auto;background:#fff;padding:40px;border-radius:16px;box-shadow:0 4px 20px rgba(0,0,0,0.08)">
        <div style="font-size:3em">🔄</div>
        <h2 style="color:#1a1a2e">Seguimiento renovado</h2>
        <p style="color:#555">El próximo recordatorio llegará en <strong>${DIAS_SEGUIM} días</strong>.<br>Puedes cerrar esta ventana.</p>
      </div>
    </body></html>
  `);
}
