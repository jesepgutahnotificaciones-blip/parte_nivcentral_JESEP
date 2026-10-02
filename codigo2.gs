/* =====================================================
   CONFIGURACIÓN
   ===================================================== */

// ID del libro de Google Sheets (el que contiene USUARIOS, LISTADO_BASE, NOVEDADES...).
// Funciona tanto si el script está ligado al libro como si es un proyecto independiente.
var ID_LIBRO = '1kLYnWBgKgMfahllxWwe5ijls52YM3t_klg4XRnNaC0Q';

// While sea true, el mensaje de "Usuario o contraseña incorrectos" incluye datos de diagnóstico.
// Cuando el login funcione, cámbielo a false y vuelva a implementar.
var DEBUG_LOGIN = true;

// Marcador de versión.
// Al abrir  <URL del Web App>/exec?accion=version  debe responder con este texto.
// Si responde otra cosa (o "Acción no válida"), el despliegue está desactualizado:
// hay que pegar este código y hacer "Implementar > Nueva versión".
var VERSION_APP = 'JESEP-2026-10-02-r18';

// Separador entre el Tipo (columna E) y el nombre del funcionario (columna F).
// Cámbialo si prefieres otro formato, por ejemplo ' | ' o ' - '.
var SEPARADOR_NOVEDAD = ' - ';

function versionApp() {
  return { estado: true, version: VERSION_APP };
}

function abrirLibro_() {
  if (ID_LIBRO) return SpreadsheetApp.openById(ID_LIBRO);
  return SpreadsheetApp.getActiveSpreadsheet();
}


/* =====================================================
   ENRUTADOR PRINCIPAL (JSONP PARA GITHUB PAGES / WEB APP)
   ===================================================== */

/* =====================================================
   ENRUTADO DE MÓDULOS POR USUARIO
   - index_1: módulo de turnos / novedades generales (DISPONIBLE_*)
   - index_2: módulo por área (SGSST, VAC, PAS, CIT, HIS, PRO, UBL, GH)
   - ADMINISTRADOR: acceso a ambos
   ===================================================== */
function modulosDeUsuario_(usuarioId, rol) {
  var id = String(usuarioId || '').trim().toUpperCase();
  var esAdmin = (String(rol || '').trim().toUpperCase() === 'ADMINISTRADOR');

  // Administradores: ven los dos módulos y el selector de módulo se los muestra
  if (esAdmin) return ['index_1', 'index_2'];

  // Usuarios del módulo de áreas JESEP -> solo index_2
  var usuariosArea = {
    'SGSST_JESEP': 1, 'VAC_JESEP': 1, 'PAS_JESEP': 1, 'CIT_JESEP': 1,
    'HIS_JESEP': 1, 'PRO_JESEP': 1, 'UBL_JESEP': 1, 'GH_JESEP': 1
  };
  if (usuariosArea[id]) return ['index_2'];

  // Usuarios de turnos DISPONIBLE_* -> solo index_1
  if (id.indexOf('DISPONIBLE_') === 0) return ['index_1'];

  // Cualquier otro OPERADOR -> solo index_1 (comportamiento previo)
  return ['index_1'];
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var accion = p.accion;

  /* ---------------------------------------------------------------
     DESCARGA DIRECTA DE ARCHIVOS (PDF / XLSX)
     Devuelve el binario con Content-Disposition: attachment, de modo
     que el navegador lo descargue de inmediato. Nada queda en Drive.
     Se atiende antes del JSONP porque no usa callback.
     --------------------------------------------------------------- */
  if (accion === 'descargarReporte') {
    try {
      return servirArchivoReporte_(String(p.formato || 'pdf').toLowerCase(),
                                  String(p.filtro || 'A'));
    } catch (err) {
      // Incluye la versión en el texto para identificar siempre el código desplegado
      return ContentService
        .createTextOutput('[VERSION ' + VERSION_APP + ']\nNo se pudo generar el archivo.\n\n' +
          (err && err.message ? err.message : err))
        .setMimeType(ContentService.MimeType.TEXT);
    }
  }

  // Solo se permiten caracteres válidos en el nombre del callback (evita inyección de código)
  var callback = String(p.callback || '').replace(/[^a-zA-Z0-9_.$]/g, '');
  var argsRaw = p.args;

  function responder_(obj) {
    var json = JSON.stringify(obj);
    if (!callback) {
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  var args = [];
  try {
    if (argsRaw) args = JSON.parse(argsRaw);
  } catch (err) {
    return responder_({ __jsonp_error: true, mensaje: 'Error al procesar los argumentos.' });
  }

  // Si se abre el Web App sin parámetros, se devuelve el diagnóstico del esquema
  // de hojas. Así se puede verificar la instalación sin escribir query strings.
  if (!accion) {
    return responder_(diagnosticarEsquemaDataSafe());
  }

  var resultado;
  try {
    switch (accion) {
      case 'validarUsuario':
        resultado = validarUsuario(args[0], args[1]);
        break;
      case 'verificarSesion':
        resultado = verificarSesion(args[0]);
        break;
      case 'cerrarSesionCliente':
        resultado = cerrarSesionCliente(args[0]);
        break;
      case 'buscarFuncionario':
        resultado = buscarFuncionario(args[0], args[1]);
        break;
      case 'obtenerFichaFuncionario':
        resultado = obtenerFichaFuncionario(args[0], args[1]);
        break;
      case 'obtenerHistorialFuncionario':
        resultado = obtenerHistorialFuncionario(args[0], args[1]);
        break;
      case 'registrarNovedad':
        resultado = registrarNovedad(args[0], args[1]);
        break;
      case 'consultarPorTurno':
        resultado = consultarPorTurno(args[0], args[1]);
        break;
      case 'previsualizarReporteTurno':
        resultado = previsualizarReporteTurno(args[0], args[1]);
        break;
      case 'generarReporteTurno':
        resultado = generarReporteTurno(args[0], args[1]);
        break;
      case 'listarReportes':
        resultado = listarReportes(args[0], args[1]);
        break;
      case 'listarUsuarios':
        resultado = listarUsuarios(args[0]);
        break;
      case 'crearUsuario':
        resultado = crearUsuario(args[0], args[1], args[2], args[3], args[4]);
        break;
      case 'cambiarEstadoUsuario':
        resultado = cambiarEstadoUsuario(args[0], args[1], args[2]);
        break;
      case 'listarFuncionarios':
        resultado = listarFuncionarios(args[0]);
        break;
      case 'agregarFuncionario':
        resultado = agregarFuncionario(args[0], args[1]);
        break;
      case 'eliminarFuncionario':
        resultado = eliminarFuncionario(args[0], args[1]);
        break;
      case 'cambiarTurnoFuncionario':
        resultado = cambiarTurnoFuncionario(args[0], args[1], args[2]);
        break;
      case 'registrarHorarioFlexible':
        resultado = registrarHorarioFlexible(args[0], args[1]);
        break;
      case 'diagnosticarEsquema':
        resultado = diagnosticarEsquemaDataSafe();
        break;
      case 'version':
        resultado = versionApp();
        break;
      default:
        resultado = {
          estado: false,
          version: VERSION_APP,
          mensaje: (accion ? ('Acción no válida: ' + accion)
                           : 'Falta el parametro "accion". Abra el Web App con ?accion=<nombre>.'),
          accionesValidas: [
            'version', 'diagnosticarEsquema',
            'validarUsuario', 'verificarSesion', 'cerrarSesionCliente',
            'buscarFuncionario', 'obtenerFichaFuncionario', 'obtenerHistorialFuncionario',
            'registrarNovedad', 'consultarPorTurno', 'previsualizarReporteTurno',
            'generarReporteTurno', 'listarReportes', 'listarUsuarios',
            'crearUsuario', 'cambiarEstadoUsuario', 'listarFuncionarios',
            'agregarFuncionario', 'eliminarFuncionario',
            'cambiarTurnoFuncionario', 'registrarHorarioFlexible'
          ]
        };
    }
  } catch (error) {
    // El mensaje incluye la versión para saber siempre qué código falló
    resultado = {
      estado: false,
      version: VERSION_APP,
      mensaje: '[' + VERSION_APP + '] ' + error.toString()
    };
  }

  return responder_(resultado);
}


/* =====================================================
   1. GESTIÓN DE USUARIOS Y AUTENTICACIÓN
   ===================================================== */

function limpiarTexto_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, ' ')
    .trim();
}

function normalizarCab_(v) {
  return limpiarTexto_(v).toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// SHA-256 en hexadecimal (64 caracteres)
function sha256Hex_(texto) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(texto), Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// Compara la clave ingresada con la guardada en la hoja.
// Acepta clave en texto plano o guardada como hash SHA-256 (con o sin "salt").
function claveCoincide_(ingresada, almacenada, salt) {
  if (almacenada === ingresada) return true;
  var alm = almacenada.toLowerCase();
  if (/^[a-f0-9]{64}$/.test(alm)) {
    var candidatos = [sha256Hex_(ingresada)];
    if (salt) {
      candidatos.push(sha256Hex_(ingresada + salt));
      candidatos.push(sha256Hex_(salt + ingresada));
    }
    return candidatos.indexOf(alm) !== -1;
  }
  return false;
}

// Utilidad: ejecútela desde el editor para generar el hash de una clave nueva
// y pegarlo en la columna de clave de la hoja USUARIOS.
function generarHashClave() {
  Logger.log(sha256Hex_('ESCRIBA_AQUI_LA_CLAVE'));
}

function buscarCab_(cab, regex, excluir) {
  for (var i = 0; i < cab.length; i++) {
    if (regex.test(cab[i]) && !(excluir && excluir.test(cab[i]))) return i;
  }
  return -1;
}

/* =====================================================
   UTILIDADES PARA ENCONTRAR COLUMNAS (normaliza acentos,
   mayúsculas y espacios/guiones bajos en los encabezados)
   ===================================================== */
function normalizarCabColumna_(v) {
  return String(v === null || v === undefined ? '' : v)
    .toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[\s_\-]+/g, ' ')
    .trim();
}

// Devuelve el índice de la primera columna cuyo encabezado coincide con alguno de los patrones
function idxColumnaBase_(cabeceras, patrones, respaldo) {
  var cab = cabeceras.map(normalizarCabColumna_);
  var lista = Array.isArray(patrones) ? patrones : [patrones];
  for (var p = 0; p < lista.length; p++) {
    for (var i = 0; i < cab.length; i++) {
      if (lista[p].test(cab[i])) return i;
    }
  }
  return (respaldo === undefined) ? -1 : respaldo;
}

function validarUsuario(usuarioIngresado, claveIngresada) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) {
      var hojas = ss.getSheets();
      for (var k = 0; k < hojas.length; k++) {
        if (normalizarCab_(hojas[k].getName()) === 'USUARIOS') { hoja = hojas[k]; break; }
      }
    }
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja USUARIOS en el libro "' + ss.getName() + '".' };

    // getDisplayValues devuelve el texto tal como se ve en la hoja (conserva ceros a la izquierda)
    var datos = hoja.getDataRange().getDisplayValues();
    if (datos.length < 2) return { estado: false, mensaje: 'La hoja USUARIOS no contiene registros.' };

    // Buscar la fila de cabeceras en las primeras 10 filas
    var filaCab = -1, idxUser = -1, idxClave = -1, idxRol = -1, idxDep = -1, idxSalt = -1, idxTurno = -1, i;
    for (i = 0; i < Math.min(datos.length, 10); i++) {
      var cab = datos[i].map(normalizarCab_);
      var u = buscarCab_(cab, /USUARIO|USER|LOGIN/);
      var c = buscarCab_(cab, /CLAVE|CONTRASE|PASS|HASH/, /SALT/);
      if (u !== -1 && c !== -1) {
        filaCab = i; idxUser = u; idxClave = c;
        idxRol = buscarCab_(cab, /^ROL|PERFIL/);
        idxDep = buscarCab_(cab, /DEPENDENCIA/);
        idxSalt = buscarCab_(cab, /SALT/);
        idxTurno = buscarCab_(cab, /TURNO|GRUPO/);
        break;
      }
    }
    var cabDetectada = filaCab !== -1;
    if (!cabDetectada) { filaCab = 0; idxUser = 0; idxClave = 1; idxRol = 2; idxDep = 3; idxTurno = -1; }

    var userIn = limpiarTexto_(usuarioIngresado).toLowerCase();
    var passIn = limpiarTexto_(claveIngresada);
    var usuarioExiste = false, largoClaveHoja = 0;

    for (i = filaCab + 1; i < datos.length; i++) {
      var row = datos[i];
      if (limpiarTexto_(row[idxUser]).toLowerCase() !== userIn) continue;
      usuarioExiste = true;
      var rowClave = limpiarTexto_(row[idxClave]);
      largoClaveHoja = rowClave.length;
      var salt = idxSalt >= 0 ? limpiarTexto_(row[idxSalt]) : '';
      if (!claveCoincide_(passIn, rowClave, salt)) continue;

      var rolUsuario = String((idxRol >= 0 && row[idxRol]) ? row[idxRol] : 'OPERADOR').trim().toUpperCase();
      var turnoPermitido = idxTurno >= 0 && row[idxTurno] ? String(row[idxTurno]).trim().toUpperCase() : '';

      // Identificador normalizado del usuario (sin espacios, en mayúsculas)
      var usuarioId = String(row[idxUser] || '').trim().toUpperCase();

      // Determinar turno permitido basado en el nombre de usuario si no está en la hoja
      if (!turnoPermitido) {
        if (usuarioId.indexOf('DISPONIBLE_A') !== -1) turnoPermitido = 'A';
        else if (usuarioId.indexOf('DISPONIBLE_B') !== -1) turnoPermitido = 'B';
        else if (usuarioId.indexOf('DISPONIBLE_C') !== -1) turnoPermitido = 'C';
      }

      // Asignar módulos según el usuario y el rol
      var modulosAsignados = modulosDeUsuario_(usuarioId, rolUsuario);

      return {
        estado: true,
        token: 'SESION_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
        usuario: {
          usuario: row[idxUser],
          usuarioId: usuarioId,
          rol: rolUsuario,
          dependencia: (idxDep >= 0 && row[idxDep]) ? row[idxDep] : 'GENERAL',
          turnoPermitido: turnoPermitido
        },
        modulos: modulosAsignados
      };
    }

    var msg = 'Usuario o contraseña incorrectos.';
    if (DEBUG_LOGIN) {
      msg += ' [DEBUG libro="' + ss.getName() + '", filas=' + (datos.length - filaCab - 1) +
        ', cabecerasDetectadas=' + cabDetectada +
        ', colUsuario=' + (idxUser + 1) + ', colClave=' + (idxClave + 1) +
        ', usuarioExiste=' + usuarioExiste +
        (usuarioExiste ? ', largoClaveHoja=' + largoClaveHoja + ', largoClaveIngresada=' + passIn.length + ', columnaSalt=' + (idxSalt >= 0) : '') + ']';
    }
    return { estado: false, mensaje: msg };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function verificarSesion(token) {
  return { estado: true };
}

function cerrarSesionCliente(token) {
  return { estado: true };
}

// Prueba manual desde el editor de Apps Script: cambie usuario y clave y ejecute esta función.
function probarLogin() {
  Logger.log(JSON.stringify(validarUsuario('TU_USUARIO', 'TU_CLAVE')));
}


/* =====================================================
   2. FUNCIONARIOS Y CONSULTAS
   ===================================================== */

function buscarFuncionario(token, valor) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja LISTADO_BASE en el Google Sheet.' };

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: false, mensaje: 'La hoja LISTADO_BASE está vacía.' };

    var cabeceras = datos[0];
    var q = String(valor).toLowerCase();
    var resultados = [];

    for (var i = 1; i < datos.length; i++) {
      var filaTexto = datos[i].join(' ').toLowerCase();
      if (filaTexto.indexOf(q) !== -1) {
        var obj = {};
        for (var j = 0; j < cabeceras.length; j++) {
          obj[cabeceras[j]] = datos[i][j];
        }
        resultados.push(obj);
      }
    }

    if (resultados.length === 0) {
      return { estado: false, mensaje: 'No se encontraron funcionarios con ese criterio.' };
    }

    return { estado: true, datos: resultados };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function obtenerFichaFuncionario(token, cedula) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'Hoja LISTADO_BASE no encontrada.' };

    var datos = hoja.getDataRange().getValues();
    var cabeceras = datos[0];
    var idxCedula = -1;

    for (var j = 0; j < cabeceras.length; j++) {
      var c = String(cabeceras[j]).toUpperCase();
      if (c === 'CEDULA' || c === 'CC') { idxCedula = j; break; }
    }

    for (var i = 1; i < datos.length; i++) {
      if (idxCedula !== -1 && String(datos[i][idxCedula]).trim() === String(cedula).trim()) {
        var obj = {};
        for (var k = 0; k < cabeceras.length; k++) { obj[cabeceras[k]] = datos[i][k]; }
        return { estado: true, funcionario: obj };
      }
    }
    return { estado: false, mensaje: 'Funcionario no encontrado.' };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function obtenerHistorialFuncionario(token, cedula) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) return { estado: true, historial: [] };

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: true, historial: [] };

    var cabeceras = datos[0];
    var idxCedula = -1;
    for (var j = 0; j < cabeceras.length; j++) {
      var c = String(cabeceras[j]).toUpperCase();
      if (c === 'CEDULA' || c === 'CC') { idxCedula = j; break; }
    }

    var historial = [];
    for (var i = 1; i < datos.length; i++) {
      if (idxCedula !== -1 && String(datos[i][idxCedula]).trim() === String(cedula).trim()) {
        var obj = {};
        for (var k = 0; k < cabeceras.length; k++) { obj[cabeceras[k]] = datos[i][k]; }
        historial.push(obj);
      }
    }
    return { estado: true, historial: historial };
  } catch (err) {
    return { estado: false, historial: [] };
  }
}


/* =====================================================
   3. REGISTRO DE NOVEDADES
   ===================================================== */

// Novedades habilitadas por cada usuario del módulo index_2.
// Debe coincidir con NOVEDADES_POR_USUARIO del front-end.
function novedadesPermitidasDe_(usuarioId) {
  var mapa = {
    'SGSST_JESEP': ['EXCUSAS MEDICAS', 'RESTRICCIONES MEDICAS', 'LICENCIA DE MATERNIDAD'],
    'VAC_JESEP': ['PLAN VACACIONAL', 'VACACIONES EXTRAORDINARIAS', 'PLAN REDUCCION', 'VACACIONES DE RETIRO'],
    'PAS_JESEP': ['COMISIONES ESTUDIO', 'COMISION DE SERVICIO', 'LICENCIA DE PATERNIDAD', 'LICENCIA DE LUTO'],
    'CIT_JESEP': ['SUSPENCION', 'CITACION JUDICIAL (PERMISO)'],
    'HIS_JESEP': ['RETIROS 3 MESES DE ALTA', 'ELIMINAR USUARIO POR RETIRO'],
    'PRO_JESEP': ['CURSO DE ASCENSO'],
    'GH_JESEP': ['CAMBIO DE TURNO', 'HORARIO FLEXIBLE']
  };
  return mapa[String(usuarioId || '').trim().toUpperCase()] || null;
}

/* Normaliza el identificador de usuario (mayúsculas, sin espacios sobrantes) */
function normalizarUsuario_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, ' ')
    .trim()
    .toUpperCase();
}

// Consulta el ROL real del usuario directamente en la hoja USUARIOS,
// para no confiar en el rol que envía el cliente.
function esUsuarioAdmin_(usuario) {
  var id = normalizarUsuario_(usuario);
  if (!id) return false;
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) return false;

    var datos = hoja.getDataRange().getDisplayValues();
    if (datos.length < 2) return false;

    var filaCab = -1, idxUser = -1, idxRol = -1, i;
    for (i = 0; i < Math.min(datos.length, 10); i++) {
      var cab = datos[i].map(normalizarCab_);
      var u = buscarCab_(cab, /USUARIO|USER|LOGIN/);
      var c = buscarCab_(cab, /CLAVE|CONTRASE|PASS|HASH/, /SALT/);
      if (u !== -1 && c !== -1) {
        filaCab = i; idxUser = u;
        idxRol = buscarCab_(cab, /^ROL|PERFIL/);
        break;
      }
    }
    if (filaCab === -1) return false;

    for (i = filaCab + 1; i < datos.length; i++) {
      if (normalizarUsuario_(datos[i][idxUser]) === id) {
        var rol = idxRol >= 0 ? String(datos[i][idxRol] || '').trim().toUpperCase() : '';
        return (rol === 'ADMINISTRADOR' || rol === 'ADMIN');
      }
    }
  } catch (e) {
    return false;
  }
  return false;
}

/* =====================================================
   ESQUEMA DE LA HOJA NOVEDADES

   Columnas exigidas:
     B -> CEDULA   (datos del funcionario)
     E -> TIPO
     F -> NOVEDAD  (concatenada con el dato de la columna E)
     G -> DIAS
     H -> FECHA INICIAL
     I -> FECHA PRESENTACION
   Las columnas restantes se completan automaticamente con
   los datos del funcionario tomados de la hoja LISTADO_BASE.
   ===================================================== */

// Devuelve la letra de columna (A, B, C...) o '?' si el índice no es válido.
// Evita que getColumnLetter(0) lance una excepción.
function letraColumna_(indiceBase) {
  var n = Number(indiceBase);
  if (!n || n < 1 || n > 18278) return '?';
  try {
    return Utilities.getColumnLetter(n);
  } catch (e) {
    return '?';
  }
}

/* =====================================================
   DESCARGA DIRECTA DE REPORTES (PDF / XLSX)

   El archivo se genera en memoria y se devuelve con
   "Content-Disposition: attachment", por lo que el navegador lo
   guarda en la carpeta de descargas del usuario.
   No se crea ningún archivo en Drive.
   ===================================================== */
// createTextOutput() acepta un Blob y conserva su tipo de contenido.
// No se llama a setMimeType porque solo acepta el enumerado MimeType
// y el .xlsx no existe en ese enumerado. Tampoco se usa createOutput(),
// que no está disponible en todos los runtimes.
function responderBlob_(blob) {
  var salida;
  if (typeof ContentService.createTextOutput === 'function') {
    salida = ContentService.createTextOutput(blob);
  } else if (typeof ContentService.createOutput === 'function') {
    salida = ContentService.createOutput(blob);
  } else {
    throw new Error('ContentService no expone ningun metodo para enviar el archivo.');
  }

  // Algunos runtimes no ofrecen setHeaders ni setContentType.
  // Se aplican solo si existen; el blob ya lleva su tipo de contenido.
  try {
    if (typeof salida.setHeaders === 'function') {
      salida = salida.setHeaders({
        'Content-Disposition': 'attachment; filename="' + blob.getName() + '"',
        'Cache-Control': 'no-store, no-cache, must-revalidate'
      });
    }
  } catch (e) { /* se ignora */ }

  try {
    if (typeof salida.setContentType === 'function') {
      salida = salida.setContentType(blob.getContentType());
    }
  } catch (e2) { /* se ignora */ }

  return salida;
}

function servirArchivoReporte_(formato, filtro) {
  var resultado = consultarPorTurno('MODO_SIN_LOGIN', filtro);
  if (!resultado.estado) {
    throw new Error(resultado.mensaje || 'No se encontraron datos para el turno ' + filtro + '.');
  }

  var funcs = resultado.datos.funcionarios;
  if (!funcs.length) {
    throw new Error('No hay funcionarios registrados en el turno ' + filtro + '.');
  }

  var sello = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMMdd-HHmmss');
  var nombreBase = 'Reporte_Turno_' + filtro + '_' + sello;
  var esExcel = (formato === 'excel' || formato === 'xlsx');
  var blob;

  if (esExcel) {
    var excel = generarExcelArchivo_(funcs, filtro, sello);
    blob = excel.blob;
    blob.setName(excel.nombre);
    blob.setContentType(excel.tipo === 'csv'
      ? 'text/csv; charset=utf-8'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  } else {
    blob = generarPDFBlob_(funcs, filtro, sello);
    blob.setName(nombreBase + '.pdf');
    blob.setContentType('application/pdf');
  }

  // El blob ya lleva su propio tipo de contenido; se delega en responderBlob_.
  return responderBlob_(blob);
}

function filaCabeceraNovedades_(hoja) {
  var ultCol = hoja.getLastColumn();
  if (!ultCol || ultCol < 1) return -1;

  var maxFilas = Math.min(Math.max(hoja.getLastRow(), 1), 10);
  var datos;
  try {
    datos = hoja.getRange(1, 1, maxFilas, ultCol).getDisplayValues();
  } catch (e) {
    return -1;
  }

  for (var i = 0; i < datos.length; i++) {
    var cab = datos[i].map(normalizarCabColumna_);
    var tieneCedula = idxColumnaBase_(cab, [/^CEDULA$/, /^CC$/], -1) !== -1;
    var tieneTipo = idxColumnaBase_(cab, [/^TIPO$/], -1) !== -1;
    if (tieneCedula || tieneTipo) return i;
  }
  return -1;
}

// Traduce el nombre de una columna de NOVEDADES a su equivalente en LISTADO_BASE
var ALIAS_BASE_ = {
  'CEDULA': [/^CEDULA$/, /^CC$/, /^NUMERO DE CEDULA$/, /^N C$/],
  'GR': [/^GR$/, /^GRADO$/],
  'NIV': [/^NIV$/, /^NIVEL$/],
  'FUNCIONARIO': [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/, /^NOMBRES$/],
  'DEPENDENCIA': [/^DEPENDENCIA$/],
  'PERT': [/^PERT$/],
  'TURNO': [/^TURNO$/],
  'SEXO': [/^SEXO$/],
  'ESTADO CIVIL': [/^ESTADO CIVIL$/],
  'SITUACION LABORAL': [/^SITUACION LABORAL$/],
  'CORREO ELECTRONICO': [/^CORREO ELECTRONICO$/, /^CORREO$/, /^EMAIL$/],
  'FECHA NACIMIENTO': [/^FECHA NACIMIENTO$/, /^FECHA DE NACIMIENTO$/, /^FNAC$/],
  'COMUNICADO OFICIAL': [/^COMUNICADO OFICIAL$/, /^COMUNICADO$/],
  'MES': [/^MES$/, /^MES ANO$/],
  'DIA': [/^DIA$/]
};

function indiceBasePorAlias_(cabecerasBase, nombreColumnaNoveldad) {
  var clave = String(nombreColumnaNoveldad || '').trim().toUpperCase();
  var alias = ALIAS_BASE_[clave];
  if (!alias) {
    // Sin alias: se busca el mismo nombre exacto en LISTADO_BASE
    return idxColumnaBase_(cabecerasBase, [new RegExp('^' + clave.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$')], -1);
  }
  return idxColumnaBase_(cabecerasBase, alias, -1);
}

// Devuelve el renglón de LISTADO_BASE del funcionario (valores crudos)
function renglonBasePorCedula_(cedula) {
  var ss = abrirLibro_();
  var hoja = ss.getSheetByName('LISTADO_BASE');
  if (!hoja) return null;

  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return null;

  var cabeceras = datos[0];
  var idxCedula = idxColumnaBase_(cabeceras, [/^CEDULA$/, /^CC$/, /^NUMERO DE CEDULA$/, /^N C$/], 0);
  var cedulaIn = String(cedula || '').trim();

  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][idxCedula]).trim() === cedulaIn) {
      return { cabeceras: cabeceras, valores: datos[i], fila: i + 1, hoja: hoja };
    }
  }
  return null;
}

function mapaColumnasNovedades_(hoja) {
  var fCab = filaCabeceraNovedades_(hoja);
  if (fCab === -1) return null;

  var cabeceras = hoja.getRange(fCab + 1, 1, 1, hoja.getLastColumn()).getDisplayValues()[0];
  var cab = cabeceras.map(normalizarCabColumna_);

  // Las columnas B, E, F, G, H e I usan la POSICION fija que exige el usuario,
  // sin importar como esten rotuladas. Asi el dato siempre cae en la letra correcta.
  // Las columnas auxiliares si se detectan por nombre; si no existen, quedan en -1.
  return {
    filaCabecera: fCab,
    cabeceras: cabeceras,
    cabNorm: cab,
    numCols: hoja.getLastColumn(),
    cedula: 1,               // B
    tipo: 4,                 // E
    novedad: 5,              // F
    dias: 6,                 // G
    fechaInicial: 7,         // H
    fechaPresentacion: 8,    // I
    descripcion: idxColumnaBase_(cab, [/^DESCRIPCION$/, /^DESCRIPCIÓN$/], -1),
    observacion: idxColumnaBase_(cab, [/^OBSERVACION$/, /^OBSERVACIÓN$/], -1),
    rv: idxColumnaBase_(cab, [/^RV$/], -1),
    texto: idxColumnaBase_(cab, [/^TEXTO$/], -1),
    fechaRegistro: idxColumnaBase_(cab, [/^FECHA REGISTRO$/, /^FECHA DE REGISTRO$/], -1)
  };
}

function construirFilaNovedades_(m, datos) {
  var fila = [];
  for (var i = 0; i < m.numCols; i++) fila.push('');

  var tipo = String(datos.tipo || '').trim();
  // Nombre del funcionario elegido en la busqueda (es lo que va en la columna F)
  var funcionario = String(datos.nombreFuncionario || '').trim();
  // Texto de la novedad escrito por el usuario
  var detalle = String(datos.novedad || '').trim();

  // B = cédula
  fila[m.cedula] = datos.cedula || datos.cc || '';

  // E = Tipo
  fila[m.tipo] = tipo;

  // F = Tipo concatenado con el nombre del funcionario seleccionado
  if (tipo && funcionario) fila[m.novedad] = tipo + SEPARADOR_NOVEDAD + funcionario;
  else fila[m.novedad] = funcionario || tipo || detalle;

  // G / H / I
  if (m.dias >= 0) fila[m.dias] = datos.dias === undefined || datos.dias === null ? '' : datos.dias;
  if (m.fechaInicial >= 0) fila[m.fechaInicial] = datos.fechaInicial || '';
  if (m.fechaPresentacion >= 0) fila[m.fechaPresentacion] = datos.fechaPresentacion || '';

  // Columnas auxiliares (si existen en la hoja)
  if (m.descripcion >= 0) fila[m.descripcion] = detalle || datos.descripcion || '';
  if (m.observacion >= 0) fila[m.observacion] = datos.observacion || detalle || datos.descripcion || '';
  if (m.rv >= 0) fila[m.rv] = datos.rv || '';
  if (m.texto >= 0) fila[m.texto] = detalle || datos.descripcion || '';
  if (m.fechaRegistro >= 0) fila[m.fechaRegistro] = new Date();

  return fila;
}

// Rellena las columnas no usadas con los datos del funcionario en LISTADO_BASE
function completarDesdeBase_(fila, m, base) {
  if (!base) return;

  var fijas = {};
  [m.cedula, m.tipo, m.novedad, m.dias, m.fechaInicial, m.fechaPresentacion,
   m.descripcion, m.observacion, m.rv, m.texto, m.fechaRegistro].forEach(function(k) {
    if (k !== undefined && k >= 0) fijas[k] = true;
  });

  for (var c = 0; c < m.numCols; c++) {
    if (fijas[c]) continue;

    var nombreCol = m.cabNorm[c];
    if (!nombreCol) continue;

    var idxBase = indiceBasePorAlias_(base.cabeceras, nombreCol);
    if (idxBase === -1 || idxBase >= base.valores.length) continue;

    var valor = base.valores[idxBase];
    fila[c] = (valor === null || valor === undefined) ? '' : valor;
  }
}

function registrarNovedad(token, datosNovedad) {
  try {
    var tipo = String(datosNovedad.tipo || '').trim().toUpperCase();

    // ADMINISTRADOR puede registrar cualquier novedad
    var esAdmin = esUsuarioAdmin_(datosNovedad.usuario);

    if (!esAdmin) {
      var permitidas = novedadesPermitidasDe_(datosNovedad.usuario);
      if (!permitidas) {
        return { estado: false, mensaje: 'Su usuario no tiene novedades habilitadas.' };
      }
      var coincide = permitidas.some(function(n) { return n.toUpperCase() === tipo; });
      if (!coincide) {
        return {
          estado: false,
          mensaje: 'El tipo "' + datosNovedad.tipo + '" no está habilitado para su usuario.'
        };
      }
    }

    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) {
      hoja = ss.insertSheet('NOVEDADES');
      hoja.appendRow(['CEDULA', '', '', '', 'TIPO', 'NOVEDAD', 'DIAS', 'FECHA INICIAL', 'FECHA PRESENTACION', 'OBSERVACION', 'RV', 'FECHA REGISTRO']);
    }

    var mapa = mapaColumnasNovedades_(hoja);
    if (!mapa) {
      return { estado: false, mensaje: 'No se pudo leer el encabezado de la hoja NOVEDADES.' };
    }

    var base = renglonBasePorCedula_(datosNovedad.cedula || datosNovedad.cc);

    // Si el navegador no envió el nombre del funcionario, se toma de LISTADO_BASE
    if (!String(datosNovedad.nombreFuncionario || '').trim() && base) {
      var idxNombre = idxColumnaBase_(base.cabeceras,
        [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/, /^NOMBRES$/, /^APELLIDOS$/], -1);
      if (idxNombre !== -1 && idxNombre < base.valores.length) {
        var vNombre = base.valores[idxNombre];
        if (vNombre !== null && vNombre !== undefined && String(vNombre).trim() !== '') {
          datosNovedad.nombreFuncionario = String(vNombre).trim();
        }
      }
    }

    var fila = construirFilaNovedades_(mapa, datosNovedad);
    completarDesdeBase_(fila, mapa, base);

    hoja.appendRow(fila);

    return { estado: true, mensaje: 'Novedad registrada correctamente.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error al registrar novedad: ' + err.message };
  }
}


/* =====================================================
   4. TURNOS Y REPORTES (CON CRUCE DE COLUMNAS E Y F)
   ===================================================== */

function consultarPorTurno(token, filtro) {
  try {
    var ss = abrirLibro_();
    var hojaBase = ss.getSheetByName('LISTADO_BASE');
    if (!hojaBase) return { estado: false, mensaje: 'Hoja LISTADO_BASE no encontrada.' };

    var datosBase = hojaBase.getDataRange().getValues();
    if (datosBase.length < 2) return { estado: false, mensaje: 'La hoja LISTADO_BASE está vacía.' };

    var cabecerasBase = datosBase[0];

    // Leer hoja NOVEDADES para cruzar información
    var hojaNovedades = ss.getSheetByName('NOVEDADES');
    var datosNovedades = [];
    if (hojaNovedades) {
      datosNovedades = hojaNovedades.getDataRange().getValues();
    }

    var funcionarios = [];

    for (var i = 1; i < datosBase.length; i++) {
      var obj = {};
      for (var j = 0; j < cabecerasBase.length; j++) {
        obj[cabecerasBase[j]] = datosBase[i][j];
      }

      var turnoFila = String(obj.turno || obj.TURNO || '').trim().toUpperCase();
      if (filtro === 'SEPRI' || turnoFila === String(filtro).trim().toUpperCase()) {

        var cedulaFuncionario = String(obj.cedula || obj.CEDULA || '').trim();
        var novedadesCruzadas = [];

        // Cruce con la hoja NOVEDADES usando las mismas posiciones fijas
        // del registro: B = cedula, E = tipo, F = novedad, G = dias.
        if (datosNovedades.length > 1 && cedulaFuncionario !== '') {
          var idxCedulaNov = 1;
          var idxTipoNov = 4;
          var idxNovedadNov = 5;
          var idxDiasNov = 6;

          for (var n = 1; n < datosNovedades.length; n++) {
            var filaNov = datosNovedades[n];

            var vCed = filaNov[idxCedulaNov];
            var cedulaNov = (vCed === null || vCed === undefined) ? '' : String(vCed).trim();
            if (cedulaNov !== cedulaFuncionario) continue;

            var vTipo = filaNov[idxTipoNov];
            var vNovedad = filaNov[idxNovedadNov];
            var vDias = filaNov[idxDiasNov];

            var tipoNov = (vTipo === null || vTipo === undefined) ? '' : String(vTipo).trim();
            var novedadNov = (vNovedad === null || vNovedad === undefined) ? '' : String(vNovedad).trim();
            var diasNov = (vDias === null || vDias === undefined) ? '' : String(vDias).trim();

            // F ya viene concatenada ("TIPO - nombre"); solo se une si viene separada
            var concatenado = novedadNov || tipoNov;
            if (novedadNov && tipoNov && novedadNov.indexOf(tipoNov) === -1) {
              concatenado = tipoNov + ' - ' + novedadNov;
            }
            if (diasNov && diasNov !== '0') {
              concatenado += (concatenado ? ' ' : '') + '(' + diasNov + ' días)';
            }

            if (concatenado) {
              novedadesCruzadas.push({ NOVEDAD: concatenado });
            }
          }
        }

        obj.historialNovedades = novedadesCruzadas;
        funcionarios.push(obj);
      }
    }

    return {
      estado: true,
      datos: {
        filtro: filtro,
        funcionarios: funcionarios
      }
    };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function previsualizarReporteTurno(token, filtro) {
  try {
    var resultadoTurno = consultarPorTurno(token, filtro);
    if (!resultadoTurno.estado) return resultadoTurno;

    var funcs = resultadoTurno.datos.funcionarios;
    var consecutivo = 'PREVIEW';
    var html = construirHtmlReporte_(funcs, filtro, consecutivo);

    return { estado: true, html: html };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function generarReporteTurno(token, filtro) {
  var paso = 'consulta de datos';

  try {
    var resultadoTurno = consultarPorTurno(token, filtro);
    if (!resultadoTurno.estado) return resultadoTurno;

    var funcs = resultadoTurno.datos.funcionarios;
    var consecutivo = 'REP-' + Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMMdd-HHmmss');

    // --- PDF ---
    paso = 'construccion del HTML';
    var htmlReporte = construirHtmlReporte_(funcs, filtro, consecutivo);

    paso = 'conversion a PDF (HtmlService)';
    var pdfBlob = generarPDFBlobDesdeHtml_(htmlReporte);

    paso = 'codificacion del PDF en base64';
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    // --- Excel (xlsx real, o CSV de respaldo si Drive no convierte) ---
    paso = 'generacion del Excel';
    var excel = generarExcelArchivo_(funcs, filtro, consecutivo);

    paso = 'codificacion del Excel en base64';
    var excelBase64 = Utilities.base64Encode(excel.blob.getBytes());

    return {
      estado: true,
      version: VERSION_APP,
      datos: {
        filtro: filtro,
        consecutivo: consecutivo,
        tipoExcel: excel.tipo,
        nombrePDF: 'Reporte_Turno_' + filtro + '_' + consecutivo + '.pdf',
        nombreExcel: excel.nombre,
        pdfBase64: pdfBase64,
        excelBase64: excelBase64
      }
    };
  } catch (err) {
    return {
      estado: false,
      version: VERSION_APP,
      mensaje: '[' + VERSION_APP + '] FALLO EN: ' + paso + ' -> ' + err.message
    };
  }
}

// Construye el HTML del reporte (compartido por PDF y previsualización)
function construirHtmlReporte_(funcs, filtro, consecutivo) {
  var html = '<html><head><meta charset="utf-8"><style>';
  html += '@page{size:Letter landscape;margin:12mm;}';
  html += 'body{font-family:Arial,Helvetica,sans-serif;color:#1f2a24;}';
  html += 'h1{color:#01592F;font-size:16pt;margin:0 0 4px;}';
  html += '.meta{font-size:8pt;color:#51615a;margin-bottom:8px;}';
  html += 'table{width:100%;border-collapse:collapse;font-size:7.5pt;}';
  html += 'th{background:#01592F;color:#fff;padding:4px 5px;text-align:left;border:1px solid #01592F;}';
  html += 'td{padding:3px 5px;border:1px solid #cccccc;vertical-align:top;}';
  html += 'tr{page-break-inside:avoid;}';
  html += 'thead{display:table-header-group;}';
  html += '</style></head><body>';
  html += '<h1>REPORTE DE TURNO ' + escHtml_(filtro) + '</h1>';
  html += '<div class="meta"><b>Consecutivo:</b> ' + escHtml_(consecutivo) + ' &nbsp;|&nbsp; ';
  html += '<b>Fecha:</b> ' + escHtml_(Utilities.formatDate(new Date(), 'America/Bogota', "dd/MM/yyyy HH:mm")) + ' &nbsp;|&nbsp; ';
  html += '<b>Total funcionarios:</b> ' + funcs.length + '</div>';
  html += '<table><thead><tr>';
  html += '<th>#</th><th>C&eacute;dula</th><th>GR</th><th>Funcionario</th><th>Dependencia</th><th>Turno</th><th>Novedades</th>';
  html += '</tr></thead><tbody>';

  funcs.forEach(function(f, i) {
    var novedadesTexto = (f.historialNovedades || []).map(function(n) { return n.NOVEDAD; }).join(', ');
    html += '<tr>';
    html += '<td>' + (i + 1) + '</td>';
    html += '<td>' + escHtml_(f.cedula || f.CEDULA || '') + '</td>';
    html += '<td>' + escHtml_(f.grado || f.GR || '') + '</td>';
    html += '<td>' + escHtml_(f.funcionario || f.FUNCIONARIO || '') + '</td>';
    html += '<td>' + escHtml_(f.dependencia || f.DEPENDENCIA || '') + '</td>';
    html += '<td>' + escHtml_(f.turno || f.TURNO || '') + '</td>';
    html += '<td>' + escHtml_(novedadesTexto || 'S/N') + '</td>';
    html += '</tr>';
  });

  html += '</tbody></table></body></html>';
  return html;
}

// Escapa caracteres especiales para que no rompan el HTML del reporte
function escHtml_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Convierte el HTML del reporte en un Blob PDF.
// Se separa del resto para poder identificar en qué paso exacto se produce
// un error (HtmlService puede lanzar "Argumento no válido" con HTML muy largo).
function generarPDFBlobDesdeHtml_(html) {
  var htmlOutput = HtmlService.createHtmlOutput(html);
  var pdf = htmlOutput.getAs('application/pdf');
  if (!pdf || pdf.getBytes().length === 0) {
    throw new Error('El PDF generado está vacío.');
  }
  return pdf;
}

function generarPDFBlob_(funcs, filtro, consecutivo) {
  var html = construirHtmlReporte_(funcs, filtro, consecutivo);
  return generarPDFBlobDesdeHtml_(html)
    .setName('Reporte_Turno_' + filtro + '_' + consecutivo + '.pdf');
}

// Genera el Excel. Intenta un .xlsx real mediante Drive; si la conversión
// no está disponible, cae a un CSV construido en memoria (que Excel abre
// sin problemas y no requiere Drive).
// Devuelve { blob, nombre, tipo }
function generarExcelArchivo_(funcs, filtro, consecutivo) {
  var filas = funcs.map(function(f, i) {
    var novedadesTexto = (f.historialNovedades || []).map(function(n) { return n.NOVEDAD; }).join(', ');
    return [
      i + 1,
      f.cedula || f.CEDULA || '',
      f.grado || f.GR || '',
      f.funcionario || f.FUNCIONARIO || '',
      f.dependencia || f.DEPENDENCIA || '',
      f.turno || f.TURNO || '',
      novedadesTexto || 'S/N'
    ];
  });
  var encabezados = ['#', 'CÉDULA', 'GR', 'FUNCIONARIO', 'DEPENDENCIA', 'TURNO', 'NOVEDADES'];
  var nombreBase = 'Reporte_Turno_' + filtro + '_' + consecutivo;

  // --- Intento 1: .xlsx real ---
  try {
    var libro = SpreadsheetApp.create('tmp_' + consecutivo);
    try {
      var hoja = libro.getActiveSheet();
      try {
        var nombreHoja = ('TURNO ' + String(filtro))
          .replace(/[\/\*\?\[\]:\\]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
          .substring(0, 90) || 'TURNO';
        hoja.setName(nombreHoja);
      } catch (e) { /* conserva el nombre por defecto */ }

      hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
      if (filas.length) {
        hoja.getRange(2, 1, filas.length, encabezados.length).setValues(filas);
        var cab = hoja.getRange(1, 1, 1, encabezados.length);
        cab.setFontWeight('bold').setBackground('#01592F').setFontColor('#ffffff');
        hoja.setFrozenRows(1);
      }

      // Se usa la cadena MIME literal: el enumerado MimeType puede llegar nulo.
      var archivo = DriveApp.getFileById(libro.getId());
      var blob = archivo.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      if (blob && blob.getBytes().length > 0) {
        blob.setName(nombreBase + '.xlsx');
        return { blob: blob, nombre: nombreBase + '.xlsx', tipo: 'xlsx' };
      }
    } finally {
      try { DriveApp.getFileById(libro.getId()).setTrashed(true); } catch (e2) {}
    }
  } catch (e1) {
    // Se ignora y se intenta el CSV
  }

  // --- Intento 2: CSV en memoria (sin Drive) ---
  function celdaCsv_(v) {
    var s = (v === null || v === undefined) ? '' : String(v);
    if (s.indexOf('"') !== -1 || s.indexOf(',') !== -1 || s.indexOf('\n') !== -1) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  var lineas = [];
  lineas.push(encabezados.map(celdaCsv_).join(','));
  filas.forEach(function(fila) { lineas.push(fila.map(celdaCsv_).join(',')); });

  // BOM para que Excel reconozca UTF-8 (acentos y ñ)
  var csv = '\ufeff' + lineas.join('\r\n');
  var csvBlob = Utilities.newBlob(csv, 'text/csv; charset=utf-8', nombreBase + '.csv');
  return { blob: csvBlob, nombre: nombreBase + '.csv', tipo: 'csv' };
}

function listarReportes(token, limite) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('REPORTES');
    if (!hoja) return { estado: true, datos: [] };

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: true, datos: [] };

    var reportes = [];
    for (var i = 1; i < datos.length; i++) {
      reportes.push({
        CONSECUTIVO: datos[i][0],
        FECHA: datos[i][1],
        USUARIO: datos[i][2],
        TURNO: datos[i][3]
      });
    }

    // Ordenar por fecha descendente y limitar
    reportes.sort(function(a, b) {
      return new Date(b.FECHA) - new Date(a.FECHA);
    });

    if (limite && reportes.length > limite) {
      reportes = reportes.slice(0, limite);
    }

    return { estado: true, datos: reportes };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}


/* =====================================================
   5. GESTIÓN DE USUARIOS (SOLO ADMIN)
   ===================================================== */

function listarUsuarios(token) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja USUARIOS.' };

    var datos = hoja.getDataRange().getDisplayValues();
    if (datos.length < 2) return { estado: true, datos: [] };

    // Buscar cabeceras
    var filaCab = -1, idxUser = -1, idxRol = -1, idxDep = -1, idxEstado = -1, i;
    for (i = 0; i < Math.min(datos.length, 10); i++) {
      var cab = datos[i].map(normalizarCab_);
      var u = buscarCab_(cab, /USUARIO|USER|LOGIN/);
      var c = buscarCab_(cab, /CLAVE|CONTRASE|PASS|HASH/, /SALT/);
      if (u !== -1 && c !== -1) {
        filaCab = i; idxUser = u;
        idxRol = buscarCab_(cab, /^ROL|PERFIL/);
        idxDep = buscarCab_(cab, /DEPENDENCIA/);
        idxEstado = buscarCab_(cab, /ESTADO|ACTIVO|INACTIVO/);
        break;
      }
    }
    if (filaCab === -1) { filaCab = 0; idxUser = 0; idxRol = 2; idxDep = 3; idxEstado = -1; }

    var usuarios = [];
    for (i = filaCab + 1; i < datos.length; i++) {
      var row = datos[i];
      var usuario = limpiarTexto_(row[idxUser]);
      if (!usuario) continue;

      usuarios.push({
        USUARIO: usuario,
        ROL: idxRol >= 0 && row[idxRol] ? row[idxRol] : 'OPERADOR',
        DEPENDENCIA: idxDep >= 0 && row[idxDep] ? row[idxDep] : '',
        ESTADO: idxEstado >= 0 && row[idxEstado] ? row[idxEstado] : 'ACTIVO'
      });
    }

    return { estado: true, datos: usuarios };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function crearUsuario(token, usuario, clave, rol, dependencia) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja USUARIOS.' };

    var datos = hoja.getDataRange().getDisplayValues();

    // Buscar cabeceras
    var filaCab = -1, idxUser = -1, idxClave = -1, idxRol = -1, idxDep = -1, idxEstado = -1, i;
    for (i = 0; i < Math.min(datos.length, 10); i++) {
      var cab = datos[i].map(normalizarCab_);
      var u = buscarCab_(cab, /USUARIO|USER|LOGIN/);
      var c = buscarCab_(cab, /CLAVE|CONTRASE|PASS|HASH/, /SALT/);
      if (u !== -1 && c !== -1) {
        filaCab = i; idxUser = u; idxClave = c;
        idxRol = buscarCab_(cab, /^ROL|PERFIL/);
        idxDep = buscarCab_(cab, /DEPENDENCIA/);
        idxEstado = buscarCab_(cab, /ESTADO|ACTIVO|INACTIVO/);
        break;
      }
    }
    if (filaCab === -1) { filaCab = 0; idxUser = 0; idxClave = 1; idxRol = 2; idxDep = 3; idxEstado = -1; }

    // Verificar si el usuario ya existe
    var userIn = limpiarTexto_(usuario).toLowerCase();
    for (i = filaCab + 1; i < datos.length; i++) {
      if (limpiarTexto_(datos[i][idxUser]).toLowerCase() === userIn) {
        return { estado: false, mensaje: 'El usuario ya existe.' };
      }
    }

    // Crear nueva fila
    var nuevaFila = [];
    nuevaFila[idxUser] = usuario;
    nuevaFila[idxClave] = sha256Hex_(clave);
    if (idxRol >= 0) nuevaFila[idxRol] = rol;
    if (idxDep >= 0) nuevaFila[idxDep] = dependencia;
    if (idxEstado >= 0) nuevaFila[idxEstado] = 'ACTIVIVO';

    hoja.appendRow(nuevaFila);

    return { estado: true, mensaje: 'Usuario creado correctamente.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error creando usuario: ' + err.message };
  }
}

function cambiarEstadoUsuario(token, usuario, nuevoEstado) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja USUARIOS.' };

    var datos = hoja.getDataRange().getDisplayValues();

    // Buscar cabeceras
    var filaCab = -1, idxUser = -1, idxEstado = -1, i;
    for (i = 0; i < Math.min(datos.length, 10); i++) {
      var cab = datos[i].map(normalizarCab_);
      var u = buscarCab_(cab, /USUARIO|USER|LOGIN/);
      var c = buscarCab_(cab, /CLAVE|CONTRASE|PASS|HASH/, /SALT/);
      if (u !== -1 && c !== -1) {
        filaCab = i; idxUser = u;
        idxEstado = buscarCab_(cab, /ESTADO|ACTIVO|INACTIVO/);
        break;
      }
    }
    if (filaCab === -1) { filaCab = 0; idxUser = 0; idxEstado = -1; }

    // Buscar y actualizar
    var userIn = limpiarTexto_(usuario).toLowerCase();
    for (i = filaCab + 1; i < datos.length; i++) {
      if (limpiarTexto_(datos[i][idxUser]).toLowerCase() === userIn) {
        if (idxEstado >= 0) {
          hoja.getRange(i + 1, idxEstado + 1).setValue(nuevoEstado);
        }
        return { estado: true, mensaje: 'Estado actualizado correctamente.' };
      }
    }

    return { estado: false, mensaje: 'Usuario no encontrado.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error actualizando estado: ' + err.message };
  }
}


/* =====================================================
   6. GESTIÓN DE FUNCIONARIOS (SOLO UBL_JESEP)
   ===================================================== */

function listarFuncionarios(token) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja LISTADO_BASE.' };

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: true, datos: [] };

    var cabeceras = datos[0];
    var funcionarios = [];

    for (var i = 1; i < datos.length; i++) {
      var obj = {};
      for (var j = 0; j < cabeceras.length; j++) {
        obj[cabeceras[j]] = datos[i][j];
      }
      funcionarios.push(obj);
    }

    return { estado: true, datos: funcionarios };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

// Crea una fila con la MISMA longitud que los encabezados y ubica cada campo
// en la columna que realmente existe en LISTADO_BASE.
function agregarFuncionario(token, datos) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'No se encontro la hoja LISTADO_BASE.' };

    var existentes = hoja.getDataRange().getValues();
    var cabeceras = existentes[0];
    var numCols = cabeceras.length;

    var idxCedula = idxColumnaBase_(cabeceras, [/^CEDULA$/, /^CC$/, /^NUMERO DE CEDULA$/, /^N C$/], 0);
    var cedulaIn = String(datos.cedula || '').trim();
    if (!cedulaIn) return { estado: false, mensaje: 'La cedula es obligatoria.' };

    for (var i = 1; i < existentes.length; i++) {
      if (String(existentes[i][idxCedula]).trim() === cedulaIn) {
        return { estado: false, mensaje: 'Ya existe un funcionario registrado con esa cedula.' };
      }
    }

    // Fila vacia del mismo ancho que la hoja (evita desalinear columnas)
    var fila = [];
    for (var c = 0; c < numCols; c++) fila.push('');

    var campos = [
      { valor: datos.cedula,            patrones: [/^CEDULA$/, /^CC$/, /^NUMERO DE CEDULA$/] },
      { valor: datos.nivel,             patrones: [/^NIV$/, /^NIVEL$/] },
      { valor: datos.grado,             patrones: [/^GR$/, /^GRADO$/] },
      { valor: datos.funcionario,       patrones: [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/, /^NOMBRES$/] },
      { valor: datos.dependencia,       patrones: [/^DEPENDENCIA$/] },
      { valor: datos.pert,              patrones: [/^PERT$/] },
      { valor: datos.turno,             patrones: [/^TURNO$/] },
      { valor: datos.mes,               patrones: [/^MES$/, /^MES ANO$/, /^MESES ANO$/] },
      { valor: datos.dia,               patrones: [/^DIA$/] },
      { valor: datos.fechaNacimiento,   patrones: [/^FECHA NACIMIENTO$/, /^FECHA DE NACIMIENTO$/, /^FNAC$/] },
      { valor: datos.correo,            patrones: [/^CORREO ELECTRONICO$/, /^CORREO$/, /^EMAIL$/] },
      { valor: datos.sexo,              patrones: [/^SEXO$/] },
      { valor: datos.estadoCivil,       patrones: [/^ESTADO CIVIL$/] },
      { valor: datos.situacionLaboral,  patrones: [/^SITUACION LABORAL$/] },
      { valor: datos.comunicado,        patrones: [/^COMUNICADO OFICIAL$/, /^COMUNICADO$/] }
    ];

    var sinColumna = [];
    for (var k = 0; k < campos.length; k++) {
      var campo = campos[k];
      if (campo.valor === '' || campo.valor === null || campo.valor === undefined) continue;
      var destino = idxColumnaBase_(cabeceras, campo.patrones, -1);
      if (destino === -1) { sinColumna.push(campo.patrones[0].source); continue; }
      fila[destino] = campo.valor;
    }

    hoja.appendRow(fila);

    return {
      estado: true,
      mensaje: 'Funcionario agregado a LISTADO_BASE correctamente.' +
        (sinColumna.length ? ' Sin columna en la hoja para: ' + sinColumna.join(', ') : '')
    };
  } catch (err) {
    return { estado: false, mensaje: 'Error agregando funcionario: ' + err.message };
  }
}


function eliminarFuncionario(token, cedula) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja LISTADO_BASE.' };

    var datos = hoja.getDataRange().getValues();
    var idxCedula = -1;
    var cabeceras = datos[0];
    for (var j = 0; j < cabeceras.length; j++) {
      var c = String(cabeceras[j]).toUpperCase();
      if (c === 'CEDULA' || c === 'CC') { idxCedula = j; break; }
    }

    if (idxCedula === -1) return { estado: false, mensaje: 'No se encontró la columna de cédula.' };

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][idxCedula]).trim() === String(cedula).trim()) {
        hoja.deleteRow(i + 1);
        return { estado: true, mensaje: 'Funcionario eliminado correctamente.' };
      }
    }

    return { estado: false, mensaje: 'Funcionario no encontrado.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error eliminando funcionario: ' + err.message };
  }
}


/* =====================================================
   7. CAMBIO DE TURNO (SOLO GH_JESEP)
   ===================================================== */

function cambiarTurnoFuncionario(token, cedula, nuevoTurno) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja LISTADO_BASE.' };

    var turno = String(nuevoTurno || '').trim();
    if (!turno) return { estado: false, mensaje: 'Debe indicar el nuevo turno.' };

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: false, mensaje: 'La hoja LISTADO_BASE está vacía.' };

    var cabeceras = datos[0];
    var idxCedula = idxColumnaBase_(cabeceras, [/^CEDULA$/, /^CC$/, /^NUMERO DE CEDULA$/, /^N C$/], 0);
    // Columna H de LISTADO_BASE = índice 7
    var idxTurno = idxColumnaBase_(cabeceras, [/^TURNO$/], 7);

    if (idxTurno >= cabeceras.length) {
      return { estado: false, mensaje: 'La columna H (TURNO) no existe en LISTADO_BASE.' };
    }

    var cedulaIn = String(cedula || '').trim();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][idxCedula]).trim() === cedulaIn) {
        var turnoAnterior = String(datos[i][idxTurno] || '').trim();
        hoja.getRange(i + 1, idxTurno + 1).setValue(turno);
        SpreadsheetApp.flush();
        return {
          estado: true,
          mensaje: 'Turno actualizado de "' + turnoAnterior + '" a "' + turno + '" (columna ' + letraColumna_(idxTurno + 1) + ').',
          turnoAnterior: turnoAnterior,
          turnoNuevo: turno,
          columna: letraColumna_(idxTurno + 1)
        };
      }
    }

    return { estado: false, mensaje: 'Funcionario no encontrado en LISTADO_BASE.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error cambiando turno: ' + err.message };
  }
}

// Registra un horario flexible en la hoja NOVEDADES
function registrarHorarioFlexible(token, datosHorario) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) {
      hoja = ss.insertSheet('NOVEDADES');
      hoja.appendRow(['CEDULA', '', '', '', 'TIPO', 'NOVEDAD', 'DIAS', 'FECHA INICIAL', 'FECHA PRESENTACION', 'OBSERVACION', 'RV', 'FECHA REGISTRO']);
    }

    var mapa = mapaColumnasNovedades_(hoja);
    if (!mapa) {
      return { estado: false, mensaje: 'No se pudo leer el encabezado de la hoja NOVEDADES.' };
    }

    var listaDias = Array.isArray(datosHorario.dias) ? datosHorario.dias : [];
    var diasTexto = listaDias.join(', ');
    var numeroDias = listaDias.length;
    var descripcion = 'De ' + datosHorario.horaInicial + ' a ' + datosHorario.horaFinal + ' | Días: ' + diasTexto;

    var fila = construirFilaNovedades_(mapa, {
      cedula: datosHorario.cedula || datosHorario.cc,
      nombreFuncionario: datosHorario.nombreFuncionario || '',
      tipo: 'HORARIO FLEXIBLE',
      novedad: descripcion,
      dias: numeroDias,
      fechaInicial: datosHorario.fechaInicial || '',
      fechaPresentacion: datosHorario.fechaPresentacion || '',
      descripcion: descripcion,
      observacion: datosHorario.observacion || descripcion,
      texto: descripcion,
      rv: datosHorario.rv || ''
    });

    completarDesdeBase_(fila, mapa, renglonBasePorCedula_(datosHorario.cedula || datosHorario.cc));

    hoja.appendRow(fila);

    return { estado: true, mensaje: 'Horario flexible registrado correctamente.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error registrando horario: ' + err.message };
  }
}


/* =====================================================
   8. DIAGNÓSTICO
   Ejecute esta función desde el editor de Apps Script para
   revisar en qué columna cae cada dato en sus hojas reales.
   ===================================================== */

// Version segura: acumula el reporte y lo devuelve como dato.
// No usa SpreadsheetApp.getUi() (falla en scripts independientes).
function diagnosticarEsquemaData() {
  var L = [];
  function add(t) { L.push(String(t)); }

  var ss;
  try {
    ss = abrirLibro_();
    add('=== DIAGNOSTICO ===');
    add('Version en ejecucion: ' + VERSION_APP);
    add('Libro: ' + ss.getName());
  } catch (e) {
    add('ERROR al abrir el libro: ' + e);
    return { estado: false, version: VERSION_APP, reporte: L.join('\n') };
  }

  try {
    var nombres = [];
    var hojas = ss.getSheets();
    for (var i = 0; i < hojas.length; i++) nombres.push(hojas[i].getName());
    add('Hojas: ' + nombres.join(', '));
  } catch (e) {
    add('ERROR al listar hojas: ' + e);
  }
  add('');

  var m = null;

  // ---------- NOVEDADES ----------
  add('=== NOVEDADES ===');
  var nov = null;
  try { nov = ss.getSheetByName('NOVEDADES'); } catch (e) { add('ERROR getSheetByName: ' + e); }

  if (!nov) {
    add('La hoja NOVEDADES no existe todavia.');
  } else {
    try { add('Filas: ' + nov.getLastRow() + ' | Columnas: ' + nov.getLastColumn()); }
    catch (e) { add('ERROR getLastRow: ' + e); }

    try { m = mapaColumnasNovedades_(nov); }
    catch (e) { add('ERROR en mapaColumnasNovedades_: ' + e); m = null; }

    if (!m) {
      add('No se detecto fila de encabezados (se requiere una fila con CEDULA o TIPO).');
    } else {
      add('Fila de encabezados: ' + (m.filaCabecera + 1));
      add('');
      add('Campo         Detectada  Esperada  Estado');
      add('CEDULA        ' + letraColumna_(m.cedula + 1) + '          B          ' + (m.cedula === 1 ? 'OK' : 'REVISAR'));
      add('TIPO          ' + letraColumna_(m.tipo + 1) + '          E          ' + (m.tipo === 4 ? 'OK' : 'REVISAR'));
      add('NOVEDAD       ' + letraColumna_(m.novedad + 1) + '          F          ' + (m.novedad === 5 ? 'OK' : 'REVISAR'));
      add('DIAS          ' + letraColumna_(m.dias + 1) + '          G          ' + (m.dias === 6 ? 'OK' : 'REVISAR'));
      add('FECHA INICIAL ' + letraColumna_(m.fechaInicial + 1) + '          H          ' + (m.fechaInicial === 7 ? 'OK' : 'REVISAR'));
      add('FECHA PRESENT ' + letraColumna_(m.fechaPresentacion + 1) + '          I          ' + (m.fechaPresentacion === 8 ? 'OK' : 'REVISAR'));
      add('');
      add('Encabezados de NOVEDADES:');
      for (var c = 0; c < m.cabeceras.length; c++) {
        add('  ' + letraColumna_(c + 1) + ') ' + (m.cabeceras[c] || '(vacia)'));
      }
    }
  }

  add('');

  // ---------- LISTADO_BASE ----------
  add('=== LISTADO_BASE ===');
  var base = null;
  try { base = ss.getSheetByName('LISTADO_BASE'); } catch (e) { add('ERROR getSheetByName: ' + e); }

  var cabBase = [];
  var filasBase = 0;

  if (!base) {
    add('La hoja LISTADO_BASE no existe.');
  } else {
    try { add('Filas: ' + base.getLastRow() + ' | Columnas: ' + base.getLastColumn()); }
    catch (e) { add('ERROR getLastRow: ' + e); }

    try {
      var d = base.getDataRange().getValues();
      if (d && d.length > 0 && d[0]) { cabBase = d[0]; filasBase = d.length - 1; }
    } catch (e) { add('ERROR al leer LISTADO_BASE: ' + e); }

    if (!cabBase.length) {
      add('LISTADO_BASE esta vacia o sin encabezados legibles.');
    } else {
      add('Encabezados: ' + cabBase.join(' | '));
      add('Total funcionarios: ' + filasBase);
      add('Columna TURNO: ' + letraColumna_(idxColumnaBase_(cabBase, [/^TURNO$/], 7) + 1) + '   (H = 8)');
      add('Columna CEDULA: ' + letraColumna_(idxColumnaBase_(cabBase, [/^CEDULA$/, /^CC$/], 0) + 1) + '   (A = 1)');
    }
  }

  add('');
  add('=== Cruce NOVEDADES -> LISTADO_BASE ===');
  if (m && cabBase.length) {
    for (var c2 = 0; c2 < m.numCols; c2++) {
      var nombre = m.cabNorm[c2] || '';
      var destino = -1;
      try { destino = indiceBasePorAlias_(cabBase, nombre); } catch (e) { destino = -1; }
      add('  ' + letraColumna_(c2 + 1) + ') ' + (m.cabeceras[c2] || '(vacia)') +
        '  <-  ' + (destino === -1 ? 'SIN COINCIDENCIA' : ('se llena desde col ' + letraColumna_(destino + 1))));
    }
  } else if (!m) {
    add('  (no se pudo leer el esquema de NOVEDADES)');
  } else {
    add('  (no se pudo leer LISTADO_BASE)');
  }

  var reporte = L.join('\n');
  try { Logger.log(reporte); } catch (e) {}
  return { estado: true, version: VERSION_APP, reporte: reporte };
}

// Version para el editor: delega en la de datos.
function diagnosticarEsquema() {
  var res = diagnosticarEsquemaDataSafe();
  try { Logger.log(res.reporte); } catch (e) {}
  return res.reporte;
}

// Envoltura: nunca deja propagar una excepcion al editor.
function diagnosticarEsquemaDataSafe() {
  try {
    return diagnosticarEsquemaData();
  } catch (e) {
    return {
      estado: false,
      version: VERSION_APP,
      reporte: 'ERROR inesperado: ' + e + '\n' + (e && e.stack ? e.stack : '')
    };
  }
}
