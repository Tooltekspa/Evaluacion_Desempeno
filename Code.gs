/**
 * SISTEMA DE EVALUACIÓN DE DESEMPEÑO — TOOLTEK SpA
 * Backend Google Apps Script
 * ---------------------------------------------------------
 * Hojas requeridas en el Spreadsheet:
 *  1) "Nomina"      -> RUT_Evaluado | Nombre_Evaluado | Cargo | Email_Evaluado | RUT_Evaluador | Nombre_Evaluador | Email_Evaluador | Periodo
 *  2) "Procesos"     -> se crea y mantiene automáticamente (1 fila = 1 proceso evaluado+periodo)
 *  3) "Config"       -> Email_RRHH | Nombre_Empresa | Rut_Empresa
 *
 * Despliegue: Implementar como Web App (Aplicación Web)
 *  - Ejecutar como: Yo (el propietario)
 *  - Quién tiene acceso: Cualquier usuario con el enlace (o de la organización)
 */

const SHEET_NOMINA = 'Nomina';
const SHEET_PROCESOS = 'Procesos';
const SHEET_CONFIG = 'Config';


/**
 * SEGURIDAD ADMIN
 * La clave de administrador NUNCA se guarda en texto plano ni viaja en el HTML.
 * Se guarda como hash SHA-256 en la celda Config!Admin_Hash (columna "Admin_Hash").
 * Para cambiarla: calcula el SHA-256 de la nueva clave y pégalo en esa celda.
 * (Puedes usar la función utilitaria generarHashAdmin('tu_clave_nueva') desde
 * el editor de Apps Script -> Ejecutar, y leer el resultado en "Registro de ejecuciones".)
 */
function generarHashAdmin(claveTexto) {
  const hash = sha256(claveTexto);
  Logger.log(hash);
  return hash;
}

function sha256(texto) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto);
  return raw.map(function(b){ const v=(b<0?b+256:b).toString(16); return v.length===1?'0'+v:v; }).join('');
}

const COLS_PROCESO = [
  'ID', 'Periodo', 'RUT_Evaluado', 'Nombre_Evaluado', 'Cargo', 'Email_Evaluado',
  'RUT_Evaluador', 'Nombre_Evaluador', 'Email_Evaluador',
  // Autoevaluación (40%)
  'Auto_Comprender', 'Auto_Desear', 'Auto_Capacidad',
  'Auto_Trabajamos', 'Auto_Excelencia', 'Auto_Desafio', 'Auto_Confianza',
  'Auto_Comentario', 'Auto_Completado', 'Auto_FechaHora',
  // Jefatura (60%)
  'Jefe_Comprender', 'Jefe_Desear', 'Jefe_Capacidad',
  'Jefe_Trabajamos', 'Jefe_Excelencia', 'Jefe_Desafio', 'Jefe_Confianza',
  'Jefe_Comentario', 'Jefe_Completado', 'Jefe_FechaHora',
  // Resultado
  'CDC_Auto', 'Valores_Auto', 'Resultado_Auto',
  'CDC_Jefe', 'Valores_Jefe', 'Resultado_Jefe',
  'Resultado_Final', 'Clasificacion_Final',
  'Estado', 'InformeEnviado', 'FechaEnvio'
];

function idx(col) { return COLS_PROCESO.indexOf(col); }

function doGet(e) {
  return handle(e);
}
function doPost(e) {
  return handle(e);
}

function handle(e) {
  try {
    const params = e.parameter || {};
    let body = {};
    if (e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (err) {}
    }
    const action = params.action || body.action;
    let result;
    switch (action) {
      case 'login':
        result = login(body.rut || params.rut);
        break;
      case 'loginAdmin':
        result = loginAdmin(body.rut, body.clave, body.userAgent);
        break;
      case 'adminListarProcesos':
        result = adminListarProcesos(body.adminToken);
        break;
      case 'adminListarNomina':
        result = adminListarNomina(body.adminToken);
        break;
      case 'adminGuardarNomina':
        result = adminGuardarNomina(body.adminToken, body.fila, body.filaOriginalIndex);
        break;
      case 'adminEliminarNomina':
        result = adminEliminarNomina(body.adminToken, body.filaIndex);
        break;
      case 'adminReenviarInforme':
        result = adminReenviarInforme(body.adminToken, body.procesoId);
        break;
      case 'adminEliminarProceso':
        result = adminEliminarProceso(body.adminToken, body.procesoId);
        break;
      case 'adminListarPeriodosConInformes':
        result = adminListarPeriodosConInformes(body.adminToken);
        break;
      case 'adminDescargarTodosLosInformes':
        result = adminDescargarTodosLosInformes(body.adminToken, body.periodo);
        break;
      case 'getProceso':
        result = getProcesoParaUsuario(body.rut, body.procesoId);
        break;
      case 'guardarRespuestas':
        result = guardarRespuestas(body);
        break;
      case 'getPreview':
        result = getPreview(body.procesoId);
        break;
      case 'confirmarEnvio':
        result = confirmarEnvio(body.procesoId, body.comentarioFinalJefe);
        break;
      default:
        result = { ok: false, error: 'Acción no reconocida: ' + action };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message, stack: err.stack });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizaRut(rut) {
  if (!rut) return '';
  return rut.toString().toUpperCase().replace(/[^0-9K]/g, '');
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** ---------------- CACHÉ DE LECTURAS DEL SHEET ---------------- */
/**
 * Lee la nómina con caché (CacheService) de 5 minutos. La nómina cambia poco,
 * así que cachearla evita releer el Sheet completo en cada login, acelerando
 * notablemente la respuesta. El caché se invalida automáticamente al editar la
 * nómina desde el panel admin (ver invalidarCacheNomina).
 */
function getNominaData(ss) {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get('nomina_data');
  if (cacheado) {
    try { return JSON.parse(cacheado); } catch (e) {}
  }
  const nomina = ss.getSheetByName(SHEET_NOMINA);
  const data = nomina.getDataRange().getValues();
  try {
    cache.put('nomina_data', JSON.stringify(data), 300); // 5 minutos
  } catch (e) {
    // Si excede el límite de tamaño del caché (100KB), se omite silenciosamente
    Logger.log('Nómina muy grande para cachear: ' + e.message);
  }
  return data;
}

/** Invalida el caché de la nómina (llamar tras agregar/editar/eliminar registros). */
function invalidarCacheNomina() {
  try { CacheService.getScriptCache().remove('nomina_data'); } catch (e) {}
}

/** ---------------- LOGIN ---------------- */
function login(rutRaw) {
  const rut = normalizaRut(rutRaw);
  if (!rut) return { ok: false, error: 'RUT vacío' };

  const ss = getSS();
  const data = getNominaData(ss);
  const headers = data[0];
  const colRutEval = headers.indexOf('RUT_Evaluado');
  const colRutJefe = headers.indexOf('RUT_Evaluador');

  const procesosSheet = getOrCreateProcesos(ss);

  const rolesEncontrados = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rutEvaluado = normalizaRut(row[colRutEval]);
    const rutEvaluador = normalizaRut(row[colRutJefe]);

    if (rutEvaluado === rut || rutEvaluador === rut) {
      const procesoId = ensureProceso(ss, procesosSheet, row, headers);
      const rolUsuario = rutEvaluado === rut ? 'evaluado' : 'evaluador';
      rolesEncontrados.push({
        procesoId: procesoId,
        rol: rolUsuario,
        nombreEvaluado: row[headers.indexOf('Nombre_Evaluado')],
        nombreEvaluador: row[headers.indexOf('Nombre_Evaluador')],
        cargo: row[headers.indexOf('Cargo')],
        periodo: row[headers.indexOf('Periodo')]
      });
    }
  }

  if (rolesEncontrados.length === 0) {
    return { ok: false, error: 'RUT no encontrado en la nómina de evaluaciones vigentes.' };
  }

  return { ok: true, roles: rolesEncontrados };
}

/** Crea la fila de proceso si no existe (clave: RUT_Evaluado + Periodo) y retorna el ID */
function ensureProceso(ss, procesosSheet, nominaRow, nominaHeaders) {
  const rutEvaluado = normalizaRut(nominaRow[nominaHeaders.indexOf('RUT_Evaluado')]);
  const periodo = nominaRow[nominaHeaders.indexOf('Periodo')];
  const id = rutEvaluado + '_' + periodo;

  const data = procesosSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][idx('ID')] === id) return id;
  }

  const newRow = new Array(COLS_PROCESO.length).fill('');
  newRow[idx('ID')] = id;
  newRow[idx('Periodo')] = periodo;
  newRow[idx('RUT_Evaluado')] = nominaRow[nominaHeaders.indexOf('RUT_Evaluado')];
  newRow[idx('Nombre_Evaluado')] = nominaRow[nominaHeaders.indexOf('Nombre_Evaluado')];
  newRow[idx('Cargo')] = nominaRow[nominaHeaders.indexOf('Cargo')];
  newRow[idx('Email_Evaluado')] = nominaRow[nominaHeaders.indexOf('Email_Evaluado')];
  newRow[idx('RUT_Evaluador')] = nominaRow[nominaHeaders.indexOf('RUT_Evaluador')];
  newRow[idx('Nombre_Evaluador')] = nominaRow[nominaHeaders.indexOf('Nombre_Evaluador')];
  newRow[idx('Email_Evaluador')] = nominaRow[nominaHeaders.indexOf('Email_Evaluador')];
  newRow[idx('Auto_Completado')] = 'NO';
  newRow[idx('Jefe_Completado')] = 'NO';
  newRow[idx('Estado')] = 'PENDIENTE';
  newRow[idx('InformeEnviado')] = 'NO';

  procesosSheet.appendRow(newRow);
  return id;
}

function getOrCreateProcesos(ss) {
  let sheet = ss.getSheetByName(SHEET_PROCESOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROCESOS);
    sheet.appendRow(COLS_PROCESO);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findProcesoRow(procesosSheet, procesoId) {
  const data = procesosSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][idx('ID')] === procesoId) return { rowIndex: i + 1, row: data[i] };
  }
  return null;
}

/** ---------------- OBTENER ESTADO DE UN PROCESO PARA UN ROL ---------------- */
function getProcesoParaUsuario(rutRaw, procesoId) {
  const rut = normalizaRut(rutRaw);
  const ss = getSS();
  const procesosSheet = getOrCreateProcesos(ss);
  const found = findProcesoRow(procesosSheet, procesoId);
  if (!found) return { ok: false, error: 'Proceso no encontrado' };
  const row = found.row;

  const esEvaluado = normalizaRut(row[idx('RUT_Evaluado')]) === rut;
  const esEvaluador = normalizaRut(row[idx('RUT_Evaluador')]) === rut;
  if (!esEvaluado && !esEvaluador) return { ok: false, error: 'No autorizado para este proceso' };

  const rol = esEvaluado ? 'evaluado' : 'evaluador';
  const yaCompleto = rol === 'evaluado' ? row[idx('Auto_Completado')] === 'SI' : row[idx('Jefe_Completado')] === 'SI';
  const otroCompleto = rol === 'evaluado' ? row[idx('Jefe_Completado')] === 'SI' : row[idx('Auto_Completado')] === 'SI';

  return {
    ok: true,
    rol: rol,
    nombreEvaluado: row[idx('Nombre_Evaluado')],
    nombreEvaluador: row[idx('Nombre_Evaluador')],
    cargo: row[idx('Cargo')],
    periodo: row[idx('Periodo')],
    yaCompleto: yaCompleto,
    otroCompleto: otroCompleto,
    estado: row[idx('Estado')],
    respuestasPrevias: yaCompleto ? extraerRespuestas(row, rol) : null
  };
}

function extraerRespuestas(row, rol) {
  const p = rol === 'evaluado' ? 'Auto_' : 'Jefe_';
  return {
    comprender: row[idx(p + 'Comprender')],
    desear: row[idx(p + 'Desear')],
    capacidad: row[idx(p + 'Capacidad')],
    trabajamos: row[idx(p + 'Trabajamos')],
    excelencia: row[idx(p + 'Excelencia')],
    desafio: row[idx(p + 'Desafio')],
    confianza: row[idx(p + 'Confianza')],
    comentario: row[idx(p + 'Comentario')]
  };
}

/** ---------------- GUARDAR RESPUESTAS (evaluado o evaluador) ---------------- */
function guardarRespuestas(body) {
  const rut = normalizaRut(body.rut);
  const procesoId = body.procesoId;
  const r = body.respuestas; // {comprender, desear, capacidad, trabajamos, excelencia, desafio, confianza, comentario}

  const ss = getSS();
  const procesosSheet = getOrCreateProcesos(ss);
  const found = findProcesoRow(procesosSheet, procesoId);
  if (!found) return { ok: false, error: 'Proceso no encontrado' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const found2 = findProcesoRow(procesosSheet, procesoId);
    const row = found2.row;
    const rowIndex = found2.rowIndex;

    const esEvaluado = normalizaRut(row[idx('RUT_Evaluado')]) === rut;
    const esEvaluador = normalizaRut(row[idx('RUT_Evaluador')]) === rut;
    if (!esEvaluado && !esEvaluador) return { ok: false, error: 'No autorizado' };

    const prefix = esEvaluado ? 'Auto_' : 'Jefe_';
    const colsToWrite = {
      'Comprender': r.comprender, 'Desear': r.desear, 'Capacidad': r.capacidad,
      'Trabajamos': r.trabajamos, 'Excelencia': r.excelencia, 'Desafio': r.desafio,
      'Confianza': r.confianza, 'Comentario': r.comentario
    };

    Object.keys(colsToWrite).forEach(function (key) {
      const colIndex = idx(prefix + key) + 1; // 1-based
      procesosSheet.getRange(rowIndex, colIndex).setValue(colsToWrite[key]);
    });
    procesosSheet.getRange(rowIndex, idx(prefix + 'Completado') + 1).setValue('SI');
    procesosSheet.getRange(rowIndex, idx(prefix + 'FechaHora') + 1).setValue(new Date());

    // Releer fila actualizada
    const updatedRow = procesosSheet.getRange(rowIndex, 1, 1, COLS_PROCESO.length).getValues()[0];
    const autoListo = updatedRow[idx('Auto_Completado')] === 'SI';
    const jefeListo = updatedRow[idx('Jefe_Completado')] === 'SI';

    let resultado = null;
    if (autoListo && jefeListo) {
      resultado = calcularYGuardarResultado(procesosSheet, rowIndex, updatedRow);
      procesosSheet.getRange(rowIndex, idx('Estado') + 1).setValue('LISTO_PARA_PREVIEW');
      notificarAmbosListos(updatedRow);
    } else {
      procesosSheet.getRange(rowIndex, idx('Estado') + 1).setValue('EN_PROGRESO');
    }

    return {
      ok: true,
      autoCompleto: autoListo,
      jefeCompleto: jefeListo,
      ambosCompletos: autoListo && jefeListo,
      resultado: resultado
    };
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- CÁLCULO (réplica exacta de la planilla TOOLTEK) ---------------- */
function calcularBloque(comprender, desear, capacidad, trabajamos, excelencia, desafio, confianza) {
  const cdc = (Number(comprender) * 0.2) + (Number(desear) * 0.2) + (Number(capacidad) * 0.2);
  const valores = (Number(trabajamos) * 0.1) + (Number(excelencia) * 0.1) + (Number(desafio) * 0.1) + (Number(confianza) * 0.1);
  return { cdc: cdc, valores: valores };
}

function clasificar(nota) {
  if (nota <= 33.4) return 'MEJORA NECESARIA';
  if (nota <= 66.6) return 'EFICAZ';
  return 'CUMPLE Y/O SUPERA LAS EXPECTATIVAS';
}

function calcularYGuardarResultado(sheet, rowIndex, row) {
  const auto = calcularBloque(
    row[idx('Auto_Comprender')], row[idx('Auto_Desear')], row[idx('Auto_Capacidad')],
    row[idx('Auto_Trabajamos')], row[idx('Auto_Excelencia')], row[idx('Auto_Desafio')], row[idx('Auto_Confianza')]
  );
  const jefe = calcularBloque(
    row[idx('Jefe_Comprender')], row[idx('Jefe_Desear')], row[idx('Jefe_Capacidad')],
    row[idx('Jefe_Trabajamos')], row[idx('Jefe_Excelencia')], row[idx('Jefe_Desafio')], row[idx('Jefe_Confianza')]
  );

  const resultadoAuto = (auto.cdc + auto.valores) * 0.4;
  const resultadoJefe = (jefe.cdc + jefe.valores) * 0.6;
  const resultadoFinal = resultadoAuto + resultadoJefe;
  const clasificacion = clasificar(resultadoFinal);

  sheet.getRange(rowIndex, idx('CDC_Auto') + 1).setValue(round1(auto.cdc));
  sheet.getRange(rowIndex, idx('Valores_Auto') + 1).setValue(round1(auto.valores));
  sheet.getRange(rowIndex, idx('Resultado_Auto') + 1).setValue(round1(resultadoAuto));
  sheet.getRange(rowIndex, idx('CDC_Jefe') + 1).setValue(round1(jefe.cdc));
  sheet.getRange(rowIndex, idx('Valores_Jefe') + 1).setValue(round1(jefe.valores));
  sheet.getRange(rowIndex, idx('Resultado_Jefe') + 1).setValue(round1(resultadoJefe));
  sheet.getRange(rowIndex, idx('Resultado_Final') + 1).setValue(round1(resultadoFinal));
  sheet.getRange(rowIndex, idx('Clasificacion_Final') + 1).setValue(clasificacion);

  return {
    cdcAuto: round1(auto.cdc), valoresAuto: round1(auto.valores), resultadoAuto: round1(resultadoAuto),
    cdcJefe: round1(jefe.cdc), valoresJefe: round1(jefe.valores), resultadoJefe: round1(resultadoJefe),
    resultadoFinal: round1(resultadoFinal), clasificacion: clasificacion
  };
}

function round1(n) { return Math.round(Number(n) * 10) / 10; }

/** ---------------- NOTIFICAR A AMBOS QUE YA SE PUEDE VER LA PREVIEW ---------------- */
function notificarAmbosListos(row) {
  // No se envía informe aún; solo se podría notificar al evaluador que ya puede previsualizar.
  // (Opcional: enviar correo de aviso. Lo dejamos silencioso por defecto para no duplicar con el informe final.)
}

/** ---------------- PREVIEW DEL INFORME (antes de enviar) ---------------- */
function getPreview(procesoId) {
  const ss = getSS();
  const procesosSheet = getOrCreateProcesos(ss);
  const found = findProcesoRow(procesosSheet, procesoId);
  if (!found) return { ok: false, error: 'Proceso no encontrado' };
  const row = found.row;

  if (row[idx('Auto_Completado')] !== 'SI' || row[idx('Jefe_Completado')] !== 'SI') {
    return { ok: false, error: 'Aún falta que evaluado y/o evaluador completen su evaluación.' };
  }

  const config = getConfig(ss);

  return {
    ok: true,
    empresa: config.nombreEmpresa,
    rutEmpresa: config.rutEmpresa,
    periodo: row[idx('Periodo')],
    nombreEvaluado: row[idx('Nombre_Evaluado')],
    cargo: row[idx('Cargo')],
    nombreEvaluador: row[idx('Nombre_Evaluador')],
    auto: extraerRespuestas(row, 'evaluado'),
    jefe: extraerRespuestas(row, 'evaluador'),
    resultado: {
      cdcAuto: row[idx('CDC_Auto')], valoresAuto: row[idx('Valores_Auto')], resultadoAuto: row[idx('Resultado_Auto')],
      cdcJefe: row[idx('CDC_Jefe')], valoresJefe: row[idx('Valores_Jefe')], resultadoJefe: row[idx('Resultado_Jefe')],
      resultadoFinal: row[idx('Resultado_Final')], clasificacion: row[idx('Clasificacion_Final')]
    }
  };
}

function getConfig(ss) {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get('config_data');
  if (cacheado) {
    try { return JSON.parse(cacheado); } catch (e) {}
  }
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) return { emailRRHH: '', nombreEmpresa: 'TOOLTEK SpA', rutEmpresa: '76.435.761-2', adminHash: '' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data[1] || [];
  const config = {
    emailRRHH: row[headers.indexOf('Email_RRHH')] || '',
    nombreEmpresa: row[headers.indexOf('Nombre_Empresa')] || 'TOOLTEK SpA',
    rutEmpresa: row[headers.indexOf('Rut_Empresa')] || '76.435.761-2',
    adminHash: row[headers.indexOf('Admin_Hash')] || ''
  };
  try { cache.put('config_data', JSON.stringify(config), 600); } catch (e) {} // 10 minutos
  return config;
}

/** ---------------- CONFIRMAR ENVÍO (el evaluador aprueba la previsualización) ---------------- */
function confirmarEnvio(procesoId, comentarioFinalJefe) {
  const ss = getSS();
  const procesosSheet = getOrCreateProcesos(ss);
  const found = findProcesoRow(procesosSheet, procesoId);
  if (!found) return { ok: false, error: 'Proceso no encontrado' };
  const row = found.row;
  const rowIndex = found.rowIndex;

  if (comentarioFinalJefe !== undefined && comentarioFinalJefe !== null) {
    procesosSheet.getRange(rowIndex, idx('Jefe_Comentario') + 1).setValue(comentarioFinalJefe);
  }

  const updatedRow = procesosSheet.getRange(rowIndex, 1, 1, COLS_PROCESO.length).getValues()[0];
  const preview = getPreview(procesoId);
  if (!preview.ok) return preview;

  const html = construirInformeHTML(preview);
  const pdfBlob = construirInformePDF(preview);

  const config = getConfig(ss);
  const destinatarios = [
    updatedRow[idx('Email_Evaluado')],
    updatedRow[idx('Email_Evaluador')],
    config.emailRRHH
  ].filter(function (e) { return e && e.toString().trim() !== ''; }).join(',');

  if (destinatarios) {
    const nombreArchivo = 'Informe_Evaluacion_' + updatedRow[idx('Nombre_Evaluado')].toString().replace(/[^a-zA-Z0-9]+/g, '_') + '.pdf';
    MailApp.sendEmail({
      to: destinatarios,
      subject: 'Informe Evaluación de Desempeño — ' + updatedRow[idx('Nombre_Evaluado')] + ' (' + updatedRow[idx('Periodo')] + ')',
      htmlBody: html,
      attachments: [pdfBlob.setName(nombreArchivo)]
    });
  }

  procesosSheet.getRange(rowIndex, idx('Estado') + 1).setValue('ENVIADO');
  procesosSheet.getRange(rowIndex, idx('InformeEnviado') + 1).setValue('SI');
  procesosSheet.getRange(rowIndex, idx('FechaEnvio') + 1).setValue(new Date());

  return { ok: true, enviadoA: destinatarios };
}

/** Genera el mismo informe como PDF tamaño carta (Letter), usando HtmlService */
/**
 * Genera el mismo informe como PDF tamaño carta, a color, con logo.
 * NOTA TÉCNICA: HtmlService.getAs('application/pdf') no preserva bien los
 * colores de fondo (backgrounds). Por eso se sube el HTML y se fuerza su
 * conversión a Google Doc vía el servicio avanzado "Drive API" (Drive.Files.insert
 * con mimeType GOOGLE_DOCS) — esa conversión sí preserva CSS de colores —
 * y luego se exporta ese Doc como PDF. El archivo temporal se envía a la papelera.
 * Requiere: Servicios avanzados → Drive API activado en este proyecto.
 */
/**
 * Genera el informe como PDF tamaño carta, a color, con logo — usando Google Slides.
 * NOTA TÉCNICA: HtmlService.getAs('application/pdf') no preserva colores de fondo,
 * y Drive.Files.insert requiere servicios avanzados que pueden no estar disponibles
 * según el proyecto de Google Cloud. SlidesApp es 100% nativo (sin activaciones extra)
 * y exporta a PDF preservando colores y formas perfectamente.
 */
/** Función temporal SOLO para forzar el prompt de autorización de Slides. Ejecutar manualmente una vez, luego se puede borrar. */
function autorizarPermisoSlides() {
  const p = SlidesApp.create('TEST_AUTORIZACION_BORRAR');
  const id = p.getId();
  p.saveAndClose();
  DriveApp.getFileById(id).setTrashed(true);
  Logger.log('Permiso de Slides autorizado correctamente.');
}

/** Función temporal SOLO para forzar el prompt de autorización de Documents. Ejecutar manualmente una vez. */
function autorizarPermisoDocs() {
  const d = DocumentApp.create('TEST_AUTORIZACION_DOCS_BORRAR');
  const id = d.getId();
  d.saveAndClose();
  DriveApp.getFileById(id).setTrashed(true);
  Logger.log('Permiso de Documents autorizado correctamente.');
}

function construirInformePDF(p) {
  const r = p.resultado;
  const colorHex = colorClasificacionHex(r.clasificacion); // ej. '#FFB703'

  const doc = DocumentApp.create('TEMP_Informe_' + p.nombreEvaluado);
  const body = doc.getBody();
  body.setMarginTop(28).setMarginBottom(28).setMarginLeft(36).setMarginRight(36);
  // Tamaño carta (8.5in x 11in) en puntos — DocumentApp.PageSize no existe como setter directo,
  // pero el documento nuevo de Google Docs ya nace en tamaño carta (Letter) por defecto en cuentas US/LatAm.
  body.clear();

  /** Abreviación corta de la clasificación */
  function clasifAbrev(nota) {
    const c = clasificar(Number(nota));
    if (c === 'MEJORA NECESARIA') return 'M. Necesaria';
    if (c === 'EFICAZ') return 'Eficaz';
    return 'Supera';
  }
  function round1Local(n) { return Math.round(Number(n) * 10) / 10; }

  /** Tabla 1x1 usada como "tarjeta" de fondo gris, con el contenido pasado como párrafos previamente armados no aplica aquí;
   *  se usa simplemente para separar visualmente secciones mediante un borde superior de color. */
  function lineaSeparadora(colorLinea) {
    const tabla = body.appendTable([['']]);
    tabla.setBorderWidth(0);
    const celda = tabla.getCell(0, 0);
    celda.setBackgroundColor(colorLinea);
    celda.setPaddingTop(1).setPaddingBottom(1);
    tabla.setColumnWidth(0, 520);
  }

  /** Dibuja una barra de progreso como tabla de 2 columnas (relleno + resto) seguida de una nota a la derecha.
   *  Se simula con UNA tabla de 3 celdas: [relleno color][resto gris claro][nota gris oscuro con texto]. */
  function filaBarra(etiqueta, valorAuto, valorJefe) {
    const p1 = body.appendParagraph(etiqueta);
    p1.setFontSize(8.5).setForegroundColor('#222222').setSpacingBefore(4).setSpacingAfter(2);

    const tabla = body.appendTable([['', '', '', '', '', '']]);
    tabla.setBorderWidth(0);

    function pintarBarra(colAncho1, colAncho2, colAnchoNota, valor, colorBarra, notaTexto, offsetCol) {
      const anchoTotal = colAncho1 + colAncho2;
      const relleno = Math.max(8, anchoTotal * Math.min(Math.max(valor,0),100) / 100);
      const resto = Math.max(2, anchoTotal - relleno);

      const cRelleno = tabla.getCell(0, offsetCol);
      cRelleno.setBackgroundColor(colorBarra);
      cRelleno.setPaddingTop(2).setPaddingBottom(2);
      tabla.setColumnWidth(offsetCol, relleno);

      const cResto = tabla.getCell(0, offsetCol + 1);
      cResto.setBackgroundColor('#E2E8EE');
      cResto.setPaddingTop(2).setPaddingBottom(2);
      tabla.setColumnWidth(offsetCol + 1, resto);

      const cNota = tabla.getCell(0, offsetCol + 2);
      cNota.setBackgroundColor('#767676');
      cNota.setPaddingTop(2).setPaddingBottom(2);
      tabla.setColumnWidth(offsetCol + 2, colAnchoNota);
      const notaPar = cNota.getChild(0).asParagraph();
      notaPar.setText(notaTexto);
      notaPar.setForegroundColor('#FFFFFF').setBold(true).setFontSize(7.5);
      notaPar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }

    const colBarra = 145; // ancho aproximado de pista para cada lado
    const colNota = 70;
    pintarBarra(colBarra, 0, colNota, valorAuto, '#0081B0', valorAuto + ' · ' + clasifAbrev(valorAuto), 0);
    pintarBarra(colBarra, 0, colNota, valorJefe, '#25638F', valorJefe + ' · ' + clasifAbrev(valorJefe), 3);
  }

  /** Título de sección con color institucional */
  function tituloSeccion(texto, pesoTxt) {
    const par = body.appendParagraph(texto + '  ' + pesoTxt);
    par.editAsText().setBold(0, texto.length - 1, true).setFontSize(0, texto.length - 1, 12).setForegroundColor(0, texto.length - 1, '#25638F');
    const restoInicio = texto.length;
    if (pesoTxt) {
      par.editAsText().setFontSize(restoInicio, par.getText().length - 1, 8).setForegroundColor(restoInicio, par.getText().length - 1, '#6B7785');
    }
    par.setSpacingBefore(14).setSpacingAfter(6);
  }

  // ===================== HEADER (franja azul continua, logos superpuestos) =====================
  const headerTabla = body.appendTable([['', '', '']]);
  headerTabla.setBorderWidth(0);

  // Celda izquierda: logo TOOLTEK completo
  const celdaLogo = headerTabla.getCell(0, 0);
  celdaLogo.setBackgroundColor('#25638F');
  celdaLogo.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(10).setPaddingRight(4);
  headerTabla.setColumnWidth(0, 145);
  try {
    const logoBlob = Utilities.newBlob(Utilities.base64Decode(LOGO_TOOLTEK_COMPLETO_BASE64), 'image/png', 'logo.png');
    const img = celdaLogo.appendImage(logoBlob);
    img.setWidth(130); img.setHeight(130 * (97/271));
  } catch(e) {
    Logger.log('Error insertando logo TOOLTEK: ' + e.message);
  }

  // Celda central: título, misma franja azul
  const celdaTitulo = headerTabla.getCell(0, 1);
  celdaTitulo.setBackgroundColor('#25638F');
  celdaTitulo.setPaddingTop(14).setPaddingBottom(14).setPaddingLeft(6).setPaddingRight(6);
  headerTabla.setColumnWidth(1, 230);
  const parTitulo = celdaTitulo.getChild(0).asParagraph();
  parTitulo.setText('Informe de Evaluación de Desempeño');
  parTitulo.setForegroundColor('#FFFFFF').setBold(true).setFontSize(11);
  parTitulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const parSub = celdaTitulo.appendParagraph(p.empresa + ' · ' + p.periodo);
  parSub.setForegroundColor('#D6E8F2').setFontSize(8);
  parSub.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // Celda derecha: logo TOOLTEK Personas (costado contrario), misma franja azul
  const celdaIcono = headerTabla.getCell(0, 2);
  celdaIcono.setBackgroundColor('#25638F');
  celdaIcono.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(4).setPaddingRight(10);
  headerTabla.setColumnWidth(2, 145);
  try {
    const iconoBlob = Utilities.newBlob(Utilities.base64Decode(LOGO_PERSONAS_ICONO_BASE64), 'image/png', 'icono.png');
    const img2 = celdaIcono.appendImage(iconoBlob);
    img2.setWidth(130); img2.setHeight(130 * (111/268));
  } catch(e) {
    Logger.log('Error insertando icono personas: ' + e.message);
  }

  body.appendParagraph('').setSpacingAfter(10);

  // ===================== TARJETA RESUMEN: nombre/cargo + resultado final =====================
  const resumenTabla = body.appendTable([['', '']]);
  resumenTabla.setBorderWidth(0);
  const cInfo = resumenTabla.getCell(0, 0);
  cInfo.setBackgroundColor('#F4F6F8');
  resumenTabla.setColumnWidth(0, 360);
  const parNombre = cInfo.getChild(0).asParagraph();
  parNombre.setText(p.nombreEvaluado);
  parNombre.setBold(true).setFontSize(13).setForegroundColor('#1F2D3A');
  const parCargo = cInfo.appendParagraph(p.cargo + ' · Evaluador: ' + p.nombreEvaluador);
  parCargo.setFontSize(9).setForegroundColor('#6B7785');

  const cResultado = resumenTabla.getCell(0, 1);
  cResultado.setBackgroundColor(colorHex);
  resumenTabla.setColumnWidth(1, 130);
  const parResultado = cResultado.getChild(0).asParagraph();
  parResultado.setText(String(r.resultadoFinal) + ' · ' + r.clasificacion);
  parResultado.setBold(true).setFontSize(12).setForegroundColor('#FFFFFF');
  parResultado.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('').setSpacingAfter(8);

  // ===================== COMPARATIVO AUTOEVALUACIÓN / JEFATURA =====================
  const puntajeAutoReal = round1Local(Number(r.cdcAuto) + Number(r.valoresAuto));
  const puntajeJefeReal = round1Local(Number(r.cdcJefe) + Number(r.valoresJefe));

  const etiquetasComp = body.appendParagraph('AUTOEVALUACIÓN (40%)                                                              JEFATURA (60%)');
  etiquetasComp.setFontSize(7.5).setForegroundColor('#6B7785').setSpacingAfter(2);

  filaBarra('', puntajeAutoReal, puntajeJefeReal);

  body.appendParagraph('').setSpacingAfter(10);

  // ===================== SECCIÓN CDC =====================
  tituloSeccion('CDC', '(Ponderación 60% del bloque competencias)');
  filaBarra('Comprender (20%)', p.auto.comprender, p.jefe.comprender);
  filaBarra('Desear (20%)', p.auto.desear, p.jefe.desear);
  filaBarra('Capacidad de Hacerlo (20%)', p.auto.capacidad, p.jefe.capacidad);

  // ===================== SECCIÓN VALORES CORPORATIVOS =====================
  tituloSeccion('VALORES CORPORATIVOS', '(Ponderación 40% del bloque competencias)');
  filaBarra('Trabajamos Juntos Construyendo Valor (10%)', p.auto.trabajamos, p.jefe.trabajamos);
  filaBarra('Nos Comprometemos con un Servicio de Excelencia (10%)', p.auto.excelencia, p.jefe.excelencia);
  filaBarra('Nos Desafiamos a Mejorar Cada Día (10%)', p.auto.desafio, p.jefe.desafio);
  filaBarra('Cultivamos y Valoramos la Confianza (10%)', p.auto.confianza, p.jefe.confianza);

  body.appendParagraph('').setSpacingAfter(10);

  // ===================== COMENTARIO JEFATURA =====================
  const comentTabla = body.appendTable([['']]);
  comentTabla.setBorderWidth(0);
  const cComent = comentTabla.getCell(0, 0);
  cComent.setBackgroundColor('#F4F6F8');
  comentTabla.setColumnWidth(0, 490);
  const parLabelComent = cComent.getChild(0).asParagraph();
  parLabelComent.setText('COMENTARIO JEFATURA');
  parLabelComent.setBold(true).setFontSize(9).setForegroundColor('#25638F');
  const parTextoComent = cComent.appendParagraph(p.jefe.comentario || '—');
  parTextoComent.setFontSize(9).setForegroundColor('#222222');

  // ===================== FOOTER =====================
  const footerPar = body.appendParagraph(p.empresa + ' · Herramientas para construir el futuro');
  footerPar.setFontSize(7).setForegroundColor('#888888').setSpacingBefore(16);

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf');
  file.setTrashed(true);

  return pdfBlob;
}

function colorClasificacionHex(clasif) {
  if (clasif === 'MEJORA NECESARIA') return '#E63946';
  if (clasif === 'EFICAZ') return '#FFB703';
  return '#2A9D8F';
}

/** ---------------- DESCARGA MASIVA DE INFORMES (ZIP por periodo) ---------------- */

/** Devuelve la lista de periodos distintos que tienen al menos un informe enviado */
function adminListarPeriodosConInformes(token) {
  if (!validarAdminToken(token)) return { ok: false, error: 'Sesión de administrador inválida o expirada.' };
  const ss = getSS();
  const procesosSheet = getOrCreateProcesos(ss);
  const data = procesosSheet.getDataRange().getValues();
  const periodosSet = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][idx('InformeEnviado')] === 'SI') {
      periodosSet[data[i][idx('Periodo')]] = true;
    }
  }
  return { ok: true, periodos: Object.keys(periodosSet) };
}

/** Genera un ZIP con los PDFs de todos los procesos enviados de un periodo, lo sube a Drive y devuelve el link */
function adminDescargarTodosLosInformes(token, periodo) {
  if (!validarAdminToken(token)) return { ok: false, error: 'Sesión de administrador inválida o expirada.' };
  const ss = getSS();
  const procesosSheet = getOrCreateProcesos(ss);
  const data = procesosSheet.getDataRange().getValues();

  const blobs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idx('InformeEnviado')] !== 'SI') continue;
    if (periodo && row[idx('Periodo')] !== periodo) continue;

    const procesoId = row[idx('ID')];
    const preview = getPreview(procesoId);
    if (!preview.ok) continue;

    const pdfBlob = construirInformePDF(preview);
    const nombreArchivo = 'Informe_' + row[idx('Nombre_Evaluado')].toString().replace(/[^a-zA-Z0-9]+/g, '_') + '.pdf';
    blobs.push(pdfBlob.setName(nombreArchivo));
  }

  if (blobs.length === 0) {
    return { ok: false, error: 'No hay informes enviados para el periodo seleccionado.' };
  }

  const zipBlob = Utilities.zip(blobs, 'Informes_Evaluacion_' + (periodo || 'Todos') + '.zip');
  const archivoZip = DriveApp.createFile(zipBlob);
  archivoZip.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { ok: true, url: archivoZip.getDownloadUrl(), cantidad: blobs.length };
}

function colorClasificacion(clasif) {
  if (clasif === 'MEJORA NECESARIA') return '#e63946';
  if (clasif === 'EFICAZ') return '#ffb703';
  return '#2a9d8f';
}

function construirInformeHTML(p, paraPdf) {
  const r = p.resultado;
  const color = colorClasificacion(r.clasificacion);
  function fila(label, auto, jefe) {
    return '<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;">' + label + '</td>' +
      '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:center;">' + auto + '</td>' +
      '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:center;">' + jefe + '</td></tr>';
  }
  const cuerpo = '' +
    '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;color:#222;">' +
    '<div style="background:#25638f;padding:16px 24px;color:#fff;display:flex;align-items:center;gap:16px;">' +
    '<img src="data:image/png;base64,' + LOGO_TOOLTEK_BASE64 + '" alt="TOOLTEK" style="height:42px;width:auto;display:block;">' +
    '<div>' +
    '<h2 style="margin:0;">Informe de Evaluación de Desempeño</h2>' +
    '<p style="margin:4px 0 0;opacity:.85;">' + p.empresa + ' · ' + p.periodo + '</p>' +
    '</div>' +
    '</div>' +
    '<div style="padding:20px 24px;">' +
    '<p><strong>Evaluado:</strong> ' + p.nombreEvaluado + ' (' + p.cargo + ')<br>' +
    '<strong>Evaluador:</strong> ' + p.nombreEvaluador + '</p>' +
    '<div style="background:' + color + ';color:#fff;border-radius:8px;padding:14px 18px;margin:16px 0;text-align:center;">' +
    '<div style="font-size:28px;font-weight:bold;">' + r.resultadoFinal + '</div>' +
    '<div style="font-size:14px;letter-spacing:.5px;">' + r.clasificacion + '</div>' +
    '</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:16px;">' +
    '<tr style="background:#f4f6f8;"><th style="padding:6px 10px;text-align:left;">Atributo</th><th style="padding:6px 10px;">Autoevaluación (40%)</th><th style="padding:6px 10px;">Jefatura (60%)</th></tr>' +
    fila('Comprender', p.auto.comprender, p.jefe.comprender) +
    fila('Desear', p.auto.desear, p.jefe.desear) +
    fila('Capacidad de Hacerlo', p.auto.capacidad, p.jefe.capacidad) +
    fila('Trabajamos Juntos Construyendo Valor', p.auto.trabajamos, p.jefe.trabajamos) +
    fila('Nos Comprometemos con un Servicio de Excelencia', p.auto.excelencia, p.jefe.excelencia) +
    fila('Nos Desafiamos a Mejorar Cada Día', p.auto.desafio, p.jefe.desafio) +
    fila('Cultivamos y Valoramos la Confianza', p.auto.confianza, p.jefe.confianza) +
    '<tr style="background:#f4f6f8;font-weight:bold;">' +
    '<td style="padding:6px 10px;">CDC (60%) / Valores (40%)</td>' +
    '<td style="padding:6px 10px;text-align:center;">CDC ' + r.cdcAuto + ' · Val ' + r.valoresAuto + '</td>' +
    '<td style="padding:6px 10px;text-align:center;">CDC ' + r.cdcJefe + ' · Val ' + r.valoresJefe + '</td>' +
    '</tr>' +
    '</table>' +
    '<p><strong>Comentario Jefatura:</strong><br>' + (p.jefe.comentario || '—') + '</p>' +
    '<p style="font-size:12px;color:#888;margin-top:24px;">' + p.empresa + ' · Herramientas para construir el futuro</p>' +
    '</div></div>';

  if (!paraPdf) return cuerpo;

  // Versión imprimible: tamaño carta (8.5in x 11in) con márgenes estándar de 1.5cm
  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>@page{size:8.5in 11in;margin:1.5cm;} body{margin:0;}</style>' +
    '</head><body>' + cuerpo + '</body></html>';
}

/** =======================================================================
 *  MÓDULO ADMINISTRADOR
 *  - La clave nunca se compara ni se guarda en texto plano: se hashea (SHA-256)
 *    y se compara contra Config!Admin_Hash.
 *  - loginAdmin devuelve un token de sesión temporal (válido 2 horas, guardado
 *    en CacheService del servidor) en vez de re-enviar la clave en cada acción.
 * ======================================================================= */

function loginAdmin(rutRaw, clave, userAgent) {
  const ss = getSS();
  const config = getConfig(ss);
  if (!config.adminHash) {
    return { ok: false, error: 'No hay clave de administrador configurada. Define Config!Admin_Hash.' };
  }
  if (!clave || sha256(clave) !== config.adminHash) {
    registrarLoginAdmin(ss, rutRaw, 'FALLIDO', userAgent);
    return { ok: false, error: 'Clave de administrador incorrecta.' };
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('admin_' + token, normalizaRut(rutRaw) || 'admin', 7200); // 2 horas
  registrarLoginAdmin(ss, rutRaw, 'EXITOSO', userAgent);
  return { ok: true, adminToken: token };
}

/**
 * Registra un intento de login de administrador (exitoso o fallido) en la pestaña 'Log_Admin'.
 * La pestaña se crea automáticamente la primera vez con sus encabezados.
 * NOTA: Apps Script no expone la IP real del cliente en un Web App, por lo que se
 * registra el User-Agent del navegador (enviado desde el frontend) como identificador disponible.
 */
function registrarLoginAdmin(ss, rutRaw, resultado, userAgent) {
  try {
    let log = ss.getSheetByName('Log_Admin');
    if (!log) {
      log = ss.insertSheet('Log_Admin');
      log.appendRow(['Fecha/Hora', 'RUT ingresado', 'Resultado', 'Navegador (User-Agent)']);
      log.getRange(1, 1, 1, 4).setFontWeight('bold');
      log.setFrozenRows(1);
    }
    const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    log.appendRow([ahora, rutRaw || '(sin RUT)', resultado, userAgent || '(no disponible)']);
  } catch (e) {
    Logger.log('Error registrando login admin: ' + e.message);
  }
}

function validarAdminToken(token) {
  if (!token) return false;
  const cached = CacheService.getScriptCache().get('admin_' + token);
  return !!cached;
}

function adminListarProcesos(token) {
  if (!validarAdminToken(token)) return { ok: false, error: 'Sesión de administrador inválida o expirada.' };
  const ss = getSS();
  const procesosSheet = getOrCreateProcesos(ss);
  const data = procesosSheet.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    lista.push({
      id: row[idx('ID')],
      periodo: row[idx('Periodo')],
      nombreEvaluado: row[idx('Nombre_Evaluado')],
      cargo: row[idx('Cargo')],
      emailEvaluado: row[idx('Email_Evaluado')],
      nombreEvaluador: row[idx('Nombre_Evaluador')],
      emailEvaluador: row[idx('Email_Evaluador')],
      autoCompleto: row[idx('Auto_Completado')] === 'SI',
      jefeCompleto: row[idx('Jefe_Completado')] === 'SI',
      estado: row[idx('Estado')],
      resultadoFinal: row[idx('Resultado_Final')],
      clasificacion: row[idx('Clasificacion_Final')],
      informeEnviado: row[idx('InformeEnviado')] === 'SI'
    });
  }
  return { ok: true, procesos: lista };
}

function adminReenviarInforme(token, procesoId) {
  if (!validarAdminToken(token)) return { ok: false, error: 'Sesión de administrador inválida o expirada.' };
  return confirmarEnvio(procesoId, null);
}

function adminEliminarProceso(token, procesoId) {
  if (!validarAdminToken(token)) return { ok: false, error: 'Sesión de administrador inválida o expirada.' };
  const ss = getSS();
  const procesosSheet = getOrCreateProcesos(ss);
  const found = findProcesoRow(procesosSheet, procesoId);
  if (!found) return { ok: false, error: 'Proceso no encontrado.' };
  procesosSheet.deleteRow(found.rowIndex);
  return { ok: true };
}

function adminListarNomina(token) {
  if (!validarAdminToken(token)) return { ok: false, error: 'Sesión de administrador inválida o expirada.' };
  const ss = getSS();
  const sheet = ss.getSheetByName(SHEET_NOMINA);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const filas = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i].every(function (v) { return v === '' || v === null; })) continue;
    const obj = {};
    headers.forEach(function (h, idx2) { obj[h] = data[i][idx2]; });
    obj._rowIndex = i; // índice interno (0-based sobre data, sin contar encabezado)
    filas.push(obj);
  }
  return { ok: true, headers: headers, filas: filas };
}

/**
 * Agrega (si filaOriginalIndex es null/undefined) o edita (si viene definido)
 * una fila de la nómina. `fila` debe tener las mismas claves que los headers.
 */
function adminGuardarNomina(token, fila, filaOriginalIndex) {
  if (!validarAdminToken(token)) return { ok: false, error: 'Sesión de administrador inválida o expirada.' };
  const ss = getSS();
  const sheet = ss.getSheetByName(SHEET_NOMINA);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nuevaFila = headers.map(function (h) { return fila[h] !== undefined ? fila[h] : ''; });

  if (filaOriginalIndex === null || filaOriginalIndex === undefined) {
    sheet.appendRow(nuevaFila);
  } else {
    const rowNumber = Number(filaOriginalIndex) + 1; // +1 porque data incluye encabezado en índice 0
    if (rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      return { ok: false, error: 'Fila a editar no encontrada.' };
    }
    sheet.getRange(rowNumber + 1, 1, 1, headers.length).setValues([nuevaFila]);
  }
  invalidarCacheNomina();
  return { ok: true };
}

function adminEliminarNomina(token, filaIndex) {
  if (!validarAdminToken(token)) return { ok: false, error: 'Sesión de administrador inválida o expirada.' };
  const ss = getSS();
  const sheet = ss.getSheetByName(SHEET_NOMINA);
  const rowNumber = Number(filaIndex) + 1 + 1; // +1 header, +1 paso a 1-based
  if (rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return { ok: false, error: 'Fila a eliminar no encontrada.' };
  }
  sheet.deleteRow(rowNumber);
  invalidarCacheNomina();
  return { ok: true };
}
