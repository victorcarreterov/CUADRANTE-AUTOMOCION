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
    ['PASSWORD', '1234']
  ]);

  crearPestana_(ss, 'Departamentos', [['Departamento', 'Orden']], [
    ['MECÁNICA', 1],
    ['CHAPA Y PINTURA', 2],
    ['ELECTRICIDAD', 3],
    ['ADMINISTRACIÓN', 4]
  ]);

  crearPestana_(ss, 'Trabajadores', [['Nombre', 'Color', 'Departamento', 'Activo']], [
    ['Juan Pérez',    '#1E88E5', 'MECÁNICA', 'SI'],
    ['Ana García',    '#43A047', 'MECÁNICA', 'SI'],
    ['Luis Martín',   '#FB8C00', 'CHAPA Y PINTURA', 'SI'],
    ['Marta López',   '#8E24AA', 'ELECTRICIDAD', 'SI'],
    ['Pedro Sánchez', '#00ACC1', 'ADMINISTRACIÓN', 'SI']
  ]);

  crearPestana_(ss, 'Motivos', [['Motivo', 'Icono']], [
    ['Almuerzo', '🍽️'],
    ['Gestión personal', '👤'],
    ['Gestión laboral', '💼'],
    ['Taller', '🔧'],
    ['Comisión de servicio', '🚔'],
    ['Médico', '🏥']
  ]);

  crearPestana_(ss, 'Vacaciones', [['Fecha', 'Trabajador']], []);

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
    .map(function (f) { return String(f[0]).trim(); });

  var motivos = hoja_('Motivos').getDataRange().getValues().slice(1)
    .filter(function (f) { return f[0]; })
    .map(function (f) { return { nombre: String(f[0]).trim(), icono: String(f[1] || '❓').trim() }; });

  var vacaciones = hoja_('Vacaciones').getDataRange().getValues().slice(1)
    .filter(function (f) { return f[0] && f[1]; })
    .map(function (f) { return { fecha: fechaTxt_(f[0]), trabajador: String(f[1]).trim() }; });

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
      hora: String(f[2] || '').trim()
    };
  });

  return {
    hoy: hoy,
    trabajadores: trabajadores,
    departamentos: departamentos,
    motivos: motivos,
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

  // "Otro": guardar el nuevo motivo en la lista configurable para próximas veces
  var motivoFinal = motivo;
  var detalleFinal = detalle;      // p. ej., el nombre del taller
  if (tipo === 'SALIDA' && motivo.toUpperCase() === 'OTRO' && detalle) {
    motivoFinal = detalle;
    detalleFinal = '';
    var hMot = hoja_('Motivos');
    var existentes = hMot.getDataRange().getValues().slice(1)
      .map(function (f) { return String(f[0]).trim().toLowerCase(); });
    if (existentes.indexOf(detalle.toLowerCase()) === -1) {
      hMot.appendRow([detalle, '📝']);
    }
  }

  var ahora = new Date();
  var fecha = Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd');
  var hora  = Utilities.formatDate(ahora, TZ, 'HH:mm');
  var hReg = hoja_('Registros');
  var filas = nombres.map(function (n) {
    n = String(n).trim();
    return [ahora, fecha, hora, n, deps[n] || '', tipo,
            tipo === 'SALIDA' ? motivoFinal : '', tipo === 'SALIDA' ? detalleFinal : ''];
  });
  hReg.getRange(hReg.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);

  return getDatosApp();
}

// ---------------------------------------------------------------- VACACIONES
/** Asigna una lista de fechas ('yyyy-MM-dd') a un trabajador. Evita duplicados. */
function asignarVacaciones(fechas, trabajador) {
  trabajador = String(trabajador).trim();
  var h = hoja_('Vacaciones');
  var existentes = {};
  h.getDataRange().getValues().slice(1).forEach(function (f) {
    if (f[0] && f[1]) existentes[fechaTxt_(f[0]) + '|' + String(f[1]).trim()] = true;
  });
  var nuevas = [];
  fechas.forEach(function (f) {
    f = String(f).trim();
    if (!existentes[f + '|' + trabajador]) nuevas.push([f, trabajador]);
  });
  if (nuevas.length) {
    h.getRange(h.getLastRow() + 1, 1, nuevas.length, 2).setValues(nuevas);
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
        hora: String(f[2] || '').trim(),
        trabajador: String(f[3]).trim(),
        departamento: String(f[4] || '').trim(),
        tipo: String(f[5] || '').trim(),
        motivo: String(f[6] || '').trim(),
        detalle: String(f[7] || '').trim()
      };
    });
}
