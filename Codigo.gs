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
      const cuerpo = `
        <div style="font-family: 'Helvetica Neue', sans-serif; max-width: 520px; margin: 0 auto; padding: 32px; background: #f9f9f9; border-radius: 12px;">
          <h2 style="color: #1a1a2e; margin-bottom: 4px;">Recordatorio de reevaluación</h2>
          <p style="color: #555; margin-top: 0;">Hola <strong>${nombreFisio}</strong>,</p>
          <p style="color: #555;">Han pasado ${DIAS_SEGUIM} días desde la primera visita de tu paciente:</p>
          <div style="background: #fff; border-left: 4px solid #4a9eff; padding: 16px 20px; border-radius: 8px; margin: 20px 0;">
            <strong style="font-size: 1.1em;">📅 ${fechaVisita} · ${horaVisita}</strong>
          </div>
          <p style="color: #555;">¿Qué quieres hacer?</p>
          <div style="margin: 24px 0; display: flex; gap: 12px; flex-direction: column;">
            <a href="${urlReev}" style="display: block; text-align: center; background: #4a9eff; color: white; padding: 14px 24px; border-radius: 8px; text-decoration: none; font-weight: 600;">
              ✅ Reevaluado — continuar seguimiento
            </a>
            <a href="${urlCerrar}" style="display: block; text-align: center; background: #ff6b6b; color: white; padding: 14px 24px; border-radius: 8px; text-decoration: none; font-weight: 600;">
              🔴 Finalizar seguimiento
            </a>
          </div>
          <p style="color: #aaa; font-size: 0.8em;">Si no haces nada, recibirás otro recordatorio en ${DIAS_SEGUIM} días.</p>
        </div>
      `;

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
