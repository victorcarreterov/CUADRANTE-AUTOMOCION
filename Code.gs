/**
 * CUADRANTE AUTOMOCIÓN — Código servidor (Google Apps Script)
 * Aplicación de gestión de disponibilidad de personal por departamentos.
 *
 * INSTALACIÓN RÁPIDA:
 *  1. Crea una hoja de Google Sheets nueva.
 *  2. Extensiones > Apps Script. Pega este archivo como "Code.gs".
 *  3. Crea el archivo HTML "Index" y pega su contenido (las imágenes van incluidas).
 *  4. Ejecuta una vez la función configurarHoja() (menú desplegable de funciones > Ejecutar).
 *  5. Implementar > Nueva implementación > Aplicación web:
 *       - Ejecutar como: Tú
 *       - Acceso: Cualquier usuario con el enlace
 *  6. Abre la URL en el móvil y añádela a la pantalla de inicio.
 */

var TZ = Session.getScriptTimeZone();

// ---------------------------------------------------------------- SERVIR APP
function doGet() {
  try {
    var salida = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Cuadrante Automoción')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    // Icono del acceso directo: URL pública en Config > ICONO_URL (debe terminar en .png o .ico)
    try {
      var urlIcono = valorConfig_('ICONO_URL');
      if (urlIcono && /\.(png|ico)$/i.test(urlIcono)) salida.setFaviconUrl(urlIcono);
    } catch (e2) { /* sin icono configurado: continuar */ }
    return salida;
  } catch (e) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:20px">' +
      '<h2 style="color:#FF0032">No se pudo cargar la pantalla</h2>' +
      '<p><b>Error:</b> ' + e.message + '</p>' +
      '<p>Comprueba que existe un archivo llamado exactamente <b>Index.html</b>.</p></div>'
    );
  }
}

// ------------------------------------------------------- CONFIGURACIÓN INICIAL
/** Ejecutar UNA VEZ para crear la estructura de la hoja con datos de ejemplo. */
function configurarHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  crearPestana_(ss, 'Config', [['Clave', 'Valor']], [
    ['PASSWORD', '1234'],
    ['PASSWORD_INFORMES', '5678'],
    ['COLOR_FINDE', '#C9D2E0']
  ]);

  crearPestana_(ss, 'Departamentos', [['Departamento', 'Orden', 'Color', 'Icono']], [
    ['MECÁNICA', 1, '#002855', '🔧'],
    ['CHAPA Y PINTURA', 2, '#B03A2E', '🎨'],
    ['ELECTRICIDAD', 3, '#B7950B', '⚡'],
    ['ADMINISTRACIÓN', 4, '#1E8449', '📋']
  ]);

  crearPestana_(ss, 'Trabajadores', [['Nombre', 'Color', 'Departamento', 'Activo']], [
    ['Juan Pérez',    '#1E88E5', 'MECÁNICA', 'SI'],
    ['Ana García',    '#43A047', 'MECÁNICA', 'SI'],
    ['Luis Martín',   '#FB8C00', 'CHAPA Y PINTURA', 'SI'],
    ['Marta López',   '#8E24AA', 'ELECTRICIDAD', 'SI'],
    ['Pedro Sánchez', '#00ACC1', 'ADMINISTRACIÓN', 'SI']
  ]);

  crearPestana_(ss, 'Motivos', [['Motivo', 'Icono', 'Detalle']], [
    ['Almuerzo', '🍽️', 'NO'],
    ['Gestión personal', '👤', 'NO'],
    ['Gestión laboral', '💼', 'NO'],
    ['Taller', '🔧', 'SI'],
    ['Comisión de servicio', '🚔', 'NO'],
    ['Médico', '🏥', 'NO']
  ]);

  crearPestana_(ss, 'TiposVacaciones', [['Tipo', 'Icono', 'Color', 'Abreviatura']], [
    ['Vacaciones anuales', '🌴', '#43A047', 'VAC'],
    ['Viaje', '✈️', '#5B9BD5', 'VIAJE'],
    ['Asuntos propios', '🏠', '#F4B183', 'AP'],
    ['Libre compensado', '🔄', '#FF66B2', 'LC']
  ]);

  crearPestana_(ss, 'Vacaciones', [['Fecha', 'Trabajador', 'Tipo']], []);

  crearPestana_(ss, 'Festivos', [['Fecha', 'Descripción']], [
    ['2026-01-01', 'Año Nuevo'],
    ['2026-01-06', 'Reyes']
  ]);

  crearPestana_(ss, 'Registros', [['Marca temporal', 'Fecha', 'Hora', 'Trabajador', 'Departamento', 'Tipo', 'Motivo', 'Detalle']], []);

  SpreadsheetApp.getUi().alert('Hoja configurada correctamente.\n\nPersonaliza las pestañas Config (contraseña), Departamentos, Trabajadores y Motivos, y después implementa la aplicación web.');
}

function crearPestana_(ss, nombre, cabecera, filas) {
  var hoja = ss.getSheetByName(nombre);
  if (hoja) return; // no sobrescribir si ya existe
  hoja = ss.insertSheet(nombre);
  hoja.getRange(1, 1, 1, cabecera[0].length).setValues(cabecera)
      .setFontWeight('bold').setBackground('#002855').setFontColor('#FFFFFF');
  if (filas.length) hoja.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
  hoja.setFrozenRows(1);
}

// ----------------------------------------------------------------- UTILIDADES
function hoja_(nombre) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h) throw new Error('Falta la pestaña "' + nombre + '". Ejecuta configurarHoja().');
  return h;
}

function fechaTxt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').trim().substring(0, 10);
}

function horaTxt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  return String(v || '').trim();
}

function fechaCorta_(iso) { // 'aaaa-mm-dd' -> 'dd/mm/aaaa'
  var p = String(iso || '').split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : iso;
}

// -------------------------------------------------------------------- ACCESO
function valorConfig_(clave) {
  var datos = hoja_('Config').getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim().toUpperCase() === clave) {
      return String(datos[i][1]).trim();
    }
  }
  return null;
}

function passwordActual_() { return valorConfig_('PASSWORD'); }

/** Token derivado de la contraseña: si esta cambia, los dispositivos recordados caducan solos. */
function tokenDe_(pass) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pass + '|CUADRANTE-AUTOMOCION');
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

/** Devuelve {ok:true, token:'...'} si la contraseña es correcta. */
function comprobarPassword(pass) {
  var real = passwordActual_();
  if (real !== null && String(pass) === real) {
    return { ok: true, token: tokenDe_(real) };
  }
  return { ok: false };
}

/** Contraseña del apartado de informes (Config > PASSWORD_INFORMES). Si no está configurada, acceso libre. */
function comprobarPasswordInformes(pass) {
  var real = valorConfig_('PASSWORD_INFORMES');
  if (real === null || real === '') return true;
  return String(pass) === real;
}

/** Valida el token guardado en un dispositivo recordado. */
function comprobarToken(token) {
  var real = passwordActual_();
  return real !== null && String(token) === tokenDe_(real);
}

// -------------------------------------------------------- DATOS PRINCIPALES
/** Devuelve toda la configuración + estado del día para pintar la app. */
function getDatosApp() {
  var trabajadores = hoja_('Trabajadores').getDataRange().getValues().slice(1)
    .filter(function (f) { return f[0] && String(f[3] || 'SI').toUpperCase() !== 'NO'; })
    .map(function (f) {
      return { nombre: String(f[0]).trim(), color: String(f[1] || '#888888').trim(), departamento: String(f[2] || 'SIN DEPARTAMENTO').trim() };
    });

  var departamentos = hoja_('Departamentos').getDataRange().getValues().slice(1)
    .filter(function (f) { return f[0]; })
    .sort(function (a, b) { return (a[1] || 99) - (b[1] || 99); })
    .map(function (f) {
      return {
        nombre: String(f[0]).trim(),
        color: String(f[2] || '#002855').trim() || '#002855',
        icono: String(f[3] || '').trim()
      };
    });

  var motivos = hoja_('Motivos').getDataRange().getValues().slice(1)
    .filter(function (f) { return f[0]; })
    .map(function (f) {
      return {
        nombre: String(f[0]).trim(),
        icono: String(f[1] || '❓').trim(),
        pideDetalle: /^(SI|SÍ|S|X|YES|TRUE|1)$/i.test(String(f[2] || '').trim())
      };
    });

  var vacaciones = hoja_('Vacaciones').getDataRange().getValues().slice(1)
    .filter(function (f) { return f[0] && f[1]; })
    .map(function (f) {
      return { fecha: fechaTxt_(f[0]), trabajador: String(f[1]).trim(), tipo: String(f[2] || '').trim() };
    });

  var tiposVacaciones = [];
  var hTV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TiposVacaciones');
  if (hTV) {
    tiposVacaciones = hTV.getDataRange().getValues().slice(1)
      .filter(function (f) { return f[0]; })
      .map(function (f) {
        return {
          nombre: String(f[0]).trim(),
          icono: String(f[1] || '🌴').trim(),
          color: String(f[2] || '#8A94A6').trim() || '#8A94A6',
          abrev: String(f[3] || '').trim()
        };
      });
  }

  // Últimos registros de HOY por trabajador (para saber quién está fuera y por qué)
  var hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var regs = hoja_('Registros').getDataRange().getValues().slice(1);
  var estadoHoy = {};
  regs.forEach(function (f) {
    if (fechaTxt_(f[1]) !== hoy) return;
    estadoHoy[String(f[3]).trim()] = {
      tipo: String(f[5]).trim(),          // SALIDA | ENTRADA
      motivo: String(f[6] || '').trim(),
      detalle: String(f[7] || '').trim(),
      hora: horaTxt_(f[2])
    };
  });

  var festivos = [];
  var hFes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Festivos');
  if (hFes) {
    festivos = hFes.getDataRange().getValues().slice(1)
      .filter(function (f) { return f[0]; })
      .map(function (f) { return fechaTxt_(f[0]); });
  }

  return {
    hoy: hoy,
    colorFinde: valorConfig_('COLOR_FINDE') || '#C9D2E0',
    festivos: festivos,
    trabajadores: trabajadores,
    departamentos: departamentos,
    motivos: motivos,
    tiposVacaciones: tiposVacaciones,
    vacaciones: vacaciones,
    estadoHoy: estadoHoy
  };
}

// ------------------------------------------------------------------- FICHAJE
/**
 * Registra entradas o salidas de uno o varios trabajadores a la vez.
 * tipo: 'SALIDA' (deja de estar disponible) o 'ENTRADA' (vuelve a estar disponible).
 * Si motivo === 'Otro' y hay detalle, el detalle se añade a la lista de Motivos.
 */
function fichar(nombres, tipo, motivo, detalle) {
  if (!nombres || !nombres.length) throw new Error('No se ha seleccionado ningún trabajador.');
  tipo = tipo === 'ENTRADA' ? 'ENTRADA' : 'SALIDA';
  motivo = String(motivo || '').trim();
  detalle = String(detalle || '').trim();

  // Mapa trabajador -> departamento
  var deps = {};
  hoja_('Trabajadores').getDataRange().getValues().slice(1).forEach(function (f) {
    if (f[0]) deps[String(f[0]).trim()] = String(f[2] || '').trim();
  });

  var motivoFinal = motivo;
  var detalleFinal = detalle;      // texto adicional si el motivo lo pide (p. ej., nombre del taller)

  var ahora = new Date();
  var fecha = Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd');
  var hora  = Utilities.formatDate(ahora, TZ, 'HH:mm');
  var hReg = hoja_('Registros');
  var filas = nombres.map(function (n) {
    n = String(n).trim();
    return [ahora, fecha, hora, n, deps[n] || '', tipo,
            tipo === 'SALIDA' ? motivoFinal : '', tipo === 'SALIDA' ? detalleFinal : ''];
  });
  var filaInicio = hReg.getLastRow() + 1;
  hReg.getRange(filaInicio, 1, filas.length, filas[0].length).setValues(filas);
  // Fecha corta y hora, sin zona horaria ni segundos
  hReg.getRange(filaInicio, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  hReg.getRange(filaInicio, 3, filas.length, 1).setNumberFormat('hh:mm');

  return getDatosApp();
}

// ---------------------------------------------------------------- VACACIONES
/** Asigna una lista de fechas ('yyyy-MM-dd') a un trabajador. Evita duplicados. */
function asignarVacaciones(fechas, trabajador, tipo) {
  trabajador = String(trabajador).trim();
  tipo = String(tipo || '').trim();
  var h = hoja_('Vacaciones');
  var existentes = {};
  h.getDataRange().getValues().slice(1).forEach(function (f) {
    if (f[0] && f[1]) existentes[fechaTxt_(f[0]) + '|' + String(f[1]).trim()] = true;
  });
  var nuevas = [];
  fechas.forEach(function (f) {
    f = String(f).trim();
    if (!existentes[f + '|' + trabajador]) nuevas.push([f, trabajador, tipo]);
  });
  if (nuevas.length) {
    h.getRange(h.getLastRow() + 1, 1, nuevas.length, 3).setValues(nuevas);
  }
  return getDatosApp();
}

/** Asigna un tipo de vacaciones a un lote de celdas 'aaaa-mm-dd|Trabajador'. */
function asignarVacacionesLote(claves, tipo) {
  tipo = String(tipo || '').trim();
  var h = hoja_('Vacaciones');
  var existentes = {};
  h.getDataRange().getValues().slice(1).forEach(function (f) {
    if (f[0] && f[1]) existentes[fechaTxt_(f[0]) + '|' + String(f[1]).trim()] = true;
  });
  var nuevas = [];
  (claves || []).forEach(function (k) {
    var p = String(k).split('|');
    if (p.length < 2) return;
    var fecha = p[0].trim();
    var trab = p.slice(1).join('|').trim();
    if (!existentes[fecha + '|' + trab]) {
      nuevas.push([fecha, trab, tipo]);
      existentes[fecha + '|' + trab] = true;
    }
  });
  if (nuevas.length) {
    h.getRange(h.getLastRow() + 1, 1, nuevas.length, 3).setValues(nuevas);
  }
  return getDatosApp();
}

/** Elimina un día de vacaciones de un trabajador. */
function quitarVacacion(fecha, trabajador) {
  fecha = String(fecha).trim();
  trabajador = String(trabajador).trim();
  var h = hoja_('Vacaciones');
  var datos = h.getDataRange().getValues();
  for (var i = datos.length - 1; i >= 1; i--) {
    if (fechaTxt_(datos[i][0]) === fecha && String(datos[i][1]).trim() === trabajador) {
      h.deleteRow(i + 1);
    }
  }
  return getDatosApp();
}

// ------------------------------------------------------------------ INFORMES
/** Devuelve los registros entre dos fechas ('yyyy-MM-dd', ambas incluidas). */
function getRegistros(desde, hasta) {
  desde = String(desde || '0000-01-01').trim();
  hasta = String(hasta || '9999-12-31').trim();
  return hoja_('Registros').getDataRange().getValues().slice(1)
    .filter(function (f) {
      var d = fechaTxt_(f[1]);
      return f[3] && d >= desde && d <= hasta;
    })
    .map(function (f) {
      return {
        fecha: fechaTxt_(f[1]),
        hora: horaTxt_(f[2]),
        trabajador: String(f[3]).trim(),
        departamento: String(f[4] || '').trim(),
        tipo: String(f[5] || '').trim(),
        motivo: String(f[6] || '').trim(),
        detalle: String(f[7] || '').trim()
      };
    });
}


// ------------------------------------------------------------- INFORMES PDF
// Los PDF se generan construyendo un Google Doc temporal (única vía que respeta
// los fondos de color), exportándolo a PDF y enviándolo a la papelera.
var B64_LOGO_PN = ''
  + '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';

var AZUL_PDF = '#063268';
var AMARILLO_PDF = '#F5E200';
var GRIS_TEXTO_PDF = '#C9D2E0';
var AZUL_TEXTO = '#002855';

function docNuevo_() {
  var doc = DocumentApp.create('informe_tmp_' + Date.now());
  var body = doc.getBody();
  body.setMarginTop(24).setMarginBottom(24).setMarginLeft(34).setMarginRight(34);
  return doc;
}

/** Banda de color a todo el ancho con texto encima. */
function bandaDoc_(body, texto, colorFondo, colorTexto, tamano, alertaTexto) {
  var t = body.appendTable([['']]);
  t.setBorderWidth(0);
  var cell = t.getCell(0, 0);
  cell.setBackgroundColor(colorFondo);
  cell.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8);
  var p = cell.getChild(0).asParagraph();
  var tx = p.appendText(texto);
  tx.setForegroundColor(colorTexto).setBold(true).setFontSize(tamano);
  if (alertaTexto) {
    var al = p.appendText(alertaTexto);
    al.setForegroundColor(AMARILLO_PDF).setBold(true).setFontSize(tamano);
  }
  return t;
}

/** Cabecera común: banda azul con logo + título gris, franja amarilla y línea de día/hora. */
function cabeceraDoc_(body) {
  var t = body.appendTable([['']]);
  t.setBorderWidth(0);
  var cell = t.getCell(0, 0);
  cell.setBackgroundColor(AZUL_PDF);
  cell.setPaddingTop(9).setPaddingBottom(9).setPaddingLeft(10).setPaddingRight(10);
  var p = cell.getChild(0).asParagraph();
  var img = p.appendInlineImage(Utilities.newBlob(Utilities.base64Decode(B64_LOGO_PN), 'image/png'));
  var ratio = img.getWidth() / img.getHeight();
  img.setHeight(26);
  img.setWidth(Math.round(26 * ratio));
  var tx = p.appendText('    DISPONIBILIDAD AUTOMOCI\u00D3N');
  tx.setForegroundColor(GRIS_TEXTO_PDF).setBold(true).setFontSize(15);

  var t2 = body.appendTable([['']]);
  t2.setBorderWidth(0);
  var c2 = t2.getCell(0, 0);
  c2.setBackgroundColor(AMARILLO_PDF);
  c2.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  c2.getChild(0).asParagraph().editAsText().setFontSize(3);

  var ahora = new Date();
  var pf = body.appendParagraph('D\u00EDa: ' + Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy') +
                                ' \u00B7 Hora: ' + Utilities.formatDate(ahora, TZ, 'HH:mm'));
  pf.editAsText().setBold(true).setFontSize(10).setForegroundColor(AZUL_TEXTO);
}

/** Da el estilo base a una tabla de datos: bordes suaves, cabecera gris, letra pequeña. */
function estilizarTabla_(t, anchos) {
  t.setBorderColor('#B9C4D6');
  t.setBorderWidth(1);
  for (var ci = 0; ci < anchos.length; ci++) t.setColumnWidth(ci, anchos[ci]);
  for (var r = 0; r < t.getNumRows(); r++) {
    for (var col = 0; col < t.getRow(r).getNumCells(); col++) {
      var cell = t.getCell(r, col);
      cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(6).setPaddingRight(6);
      var txt = cell.editAsText();
      if (txt.getText() !== '') {
        txt.setFontSize(9).setForegroundColor(AZUL_TEXTO).setBold(false);
      }
      if (r === 0) {
        cell.setBackgroundColor('#EDF0F5');
        if (txt.getText() !== '') txt.setBold(true);
      }
    }
  }
}

function pdfDesdeDoc_(doc, nombre) {
  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  var blob = file.getAs(MimeType.PDF);
  file.setTrashed(true);
  return { nombre: nombre, base64: Utilities.base64Encode(blob.getBytes()) };
}

/** PDF estándar: estado de disponibilidad de hoy por departamentos. */
function generarInformePDF() {
  var d = getDatosApp();
  var pF = d.hoy.split('-');
  var fechaDM = pF[2] + '/' + pF[1];

  var iconos = {};
  d.motivos.forEach(function (m) { iconos[m.nombre.toLowerCase()] = m.icono; });
  var tvIconos = {};
  (d.tiposVacaciones || []).forEach(function (t) { tvIconos[t.nombre.toLowerCase()] = t.icono; });

  function estadoDe(nombre) {
    for (var i = 0; i < d.vacaciones.length; i++) {
      var v = d.vacaciones[i];
      if (v.trabajador === nombre && v.fecha === d.hoy) return { tipo: 'vac', tipoVac: v.tipo || 'Vacaciones' };
    }
    var r = d.estadoHoy[nombre];
    if (r && r.tipo === 'SALIDA') return { tipo: 'fuera', motivo: r.motivo, detalle: r.detalle, hora: r.hora };
    return { tipo: 'ok' };
  }

  var doc = docNuevo_();
  var body = doc.getBody();
  cabeceraDoc_(body);

  d.departamentos.forEach(function (dep) {
    var lista = d.trabajadores.filter(function (t) { return t.departamento === dep.nombre; });
    if (!lista.length) return;
    var nDisp = lista.filter(function (t) { return estadoDe(t.nombre).tipo === 'ok'; }).length;

    if (nDisp === 0) {
      bandaDoc_(body,
        (dep.icono ? dep.icono + ' ' : '') + dep.nombre + ' \u2014 ',
        dep.color, '#FFFFFF', 11,
        '\u26A0\uFE0F ' + nDisp + '/' + lista.length + ' Disponibles');
    } else {
      bandaDoc_(body,
        (dep.icono ? dep.icono + ' ' : '') + dep.nombre + ' \u2014 ' + nDisp + ' / ' + lista.length + ' disponibles',
        dep.color, '#FFFFFF', 11, null);
    }

    var filas = [['Nombre', 'Estado', 'Detalle']];
    var estados = [null];
    lista.forEach(function (t) {
      var e = estadoDe(t.nombre);
      if (e.tipo === 'ok') {
        filas.push([t.nombre, 'DISPONIBLE', '']);
        estados.push('ok');
      } else if (e.tipo === 'fuera') {
        filas.push([t.nombre, 'NO DISPONIBLE',
          (iconos[String(e.motivo).toLowerCase()] || '') + ' ' + e.motivo +
          (e.detalle ? ' (' + e.detalle + ')' : '') + ' \u00B7 desde ' + fechaDM + ' ' + e.hora]);
        estados.push('fuera');
      } else {
        filas.push([t.nombre,
          ((tvIconos[String(e.tipoVac).toLowerCase()] || '') + ' ' + String(e.tipoVac).toUpperCase()).trim(),
          'No disponible todo el d\u00EDa']);
        estados.push('vac');
      }
    });

    var tabla = body.appendTable(filas);
    estilizarTabla_(tabla, [135, 125, 270]);
    for (var r = 1; r < filas.length; r++) {
      var celdaEstado = tabla.getCell(r, 1).editAsText();
      celdaEstado.setBold(true);
      if (estados[r] === 'ok') celdaEstado.setForegroundColor('#1F9D55');
      else if (estados[r] === 'fuera') celdaEstado.setForegroundColor('#FF0032');
      else celdaEstado.setForegroundColor('#8A94A6');
    }
  });

  var fechaC = pF[2] + '-' + pF[1] + '-' + pF[0];
  return pdfDesdeDoc_(doc, 'Disponibilidad_' + fechaC + '.pdf');
}

/** PDF filtrado (apartado protegido): histórico de un periodo con filtros. */
function generarInformeFiltrado(desde, hasta, dep, trab) {
  var d = getDatosApp();
  dep = String(dep || '').trim();
  trab = String(trab || '').trim();

  var iconos = {};
  d.motivos.forEach(function (m) { iconos[m.nombre.toLowerCase()] = m.icono; });

  var regs = getRegistros(desde, hasta).filter(function (r) {
    return (!dep || r.departamento === dep) && (!trab || r.trabajador === trab);
  });

  function dm(f) { var p = String(f).split('-'); return p.length === 3 ? p[2] + '/' + p[1] : f; }

  var doc = docNuevo_();
  var body = doc.getBody();
  cabeceraDoc_(body);

  bandaDoc_(body,
    'HIST\u00D3RICO \u2014 ' + dm(desde) + ' a ' + dm(hasta) +
    (dep ? ' \u00B7 ' + dep : '') + (trab ? ' \u00B7 ' + trab : ''),
    AZUL_PDF, '#FFFFFF', 11, null);

  if (regs.length) {
    var filas = [['Fecha', 'Hora', 'Nombre', 'Departamento', 'Movimiento', 'Motivo']];
    var tipos = [null];
    regs.forEach(function (r) {
      filas.push([dm(r.fecha), r.hora, r.trabajador, r.departamento,
        r.tipo === 'SALIDA' ? 'Inicio de salida' : 'Fin de salida',
        r.motivo ? (iconos[String(r.motivo).toLowerCase()] || '') + ' ' + r.motivo + (r.detalle ? ' (' + r.detalle + ')' : '') : '']);
      tipos.push(r.tipo);
    });
    var tabla = body.appendTable(filas);
    estilizarTabla_(tabla, [50, 42, 95, 105, 88, 150]);
    for (var r2 = 1; r2 < filas.length; r2++) {
      var cm = tabla.getCell(r2, 4).editAsText();
      cm.setBold(true).setForegroundColor(tipos[r2] === 'SALIDA' ? '#FF0032' : '#1F9D55');
    }
    var pt = body.appendParagraph(regs.length + ' movimientos en el periodo.');
    pt.editAsText().setFontSize(9).setForegroundColor(AZUL_TEXTO);
  } else {
    var pv = body.appendParagraph('Sin movimientos con los filtros seleccionados.');
    pv.editAsText().setFontSize(10).setForegroundColor(AZUL_TEXTO);
  }

  return pdfDesdeDoc_(doc, 'Historico_' + String(desde) + '_a_' + String(hasta) + '.pdf');
}
