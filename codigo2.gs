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
var VERSION_APP = 'JESEP-2026-10-07-r38';

// Separador entre el Tipo (columna E) y el nombre del funcionario (columna F).
// Cámbialo si prefieres otro formato, por ejemplo ' | ' o ' - '.
var SEPARADOR_NOVEDAD = ' - ';

function versionApp() {
  return { estado: true, version: VERSION_APP };
}

/* =====================================================
   DIAGNÓSTICO DE PERMISOS
   Ejecuta esta función desde el EDITOR de Apps Script
   (no desde la página web) y mira el resultado.
   Sirve para saber si el problema es el manifiesto,
   la autorización, o que se está editando otro proyecto.
   ===================================================== */
function revisarPermisos() {
  var L = [];
  function add(x) { L.push(String(x)); }

  add('VERSION  ' + VERSION_APP);
  add('ID_LIBRO ' + ID_LIBRO);
  add('');

  // 1) Contexto de ejecucion
  try {
    add('Usuario que ejecuta  : ' + Session.getEffectiveUser().getEmail());
  } catch (e0) {
    add('Usuario que ejecuta  : no disponible (' + e0.message + ')');
  }
  try {
    add('Usuario activo       : ' + Session.getActiveUser().getEmail());
  } catch (e1) {
    add('Usuario activo       : no disponible');
  }
  try {
    add('Script ID            : ' + ScriptApp.getService().getUrl());
  } catch (e2) {
    add('Script ID            : no disponible');
  }
  add('');

  // 2) Prueba directa de la hoja
  try {
    var ss = SpreadsheetApp.openById(ID_LIBRO);
    add('OK  SpreadsheetApp.openById -> ' + ss.getName());
    add('OK  Filas LISTADO_BASE  -> ' + ss.getSheetByName('LISTADO_BASE').getLastRow());
  } catch (err) {
    add('FALLO SpreadsheetApp.openById');
    add('     ' + err.message);
  }
  add('');

  // 3) Prueba de Drive (para el Excel)
  try {
    var tmp = SpreadsheetApp.create('tmp_permisos');
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
    add('OK  DriveApp crear/eliminar temporal');
  } catch (err2) {
    add('FALLO DriveApp -> ' + err2.message);
  }
  add('');
  add('Si las pruebas de arriba dicen OK pero la pagina web');
  add('falla, el problema es la AUTORIZACION o el DESPLIEGUE,');
  add('no el manifiesto.');
  add('Si fallan aqui, el manifiesto de ESTE proyecto sigue');
  add('sin el scope spreadsheets.');

  return L.join('\n');
}

// Abre el libro de trabajo.
// Si falta el permiso de hojas de calculo, reune datos de contexto
// (que proyecto responde y con que cuenta se ejecuta) para que el
// mensaje diga exactamente donde esta el problema.
function abrirLibro_() {
  try {
    if (ID_LIBRO) return SpreadsheetApp.openById(ID_LIBRO);
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    var msg = String((err && err.message) || err);

    if (msg.indexOf('spreadsheets') !== -1 || msg.indexOf('permisos') !== -1 ||
        msg.indexOf('Scopes') !== -1 || msg.indexOf('insufficient') !== -1) {

      var url = '(desconocido)';
      try { url = ScriptApp.getService().getUrl(); } catch (e1) {}

      var quien = '(desconocido)';
      try { quien = Session.getEffectiveUser().getEmail() || 'sin correo'; } catch (e2) {}

      var activa = '(desconocida)';
      try { var ua = Session.getActiveUser().getEmail(); activa = ua ? ua : 'ninguna (web app)'; } catch (e3) {}

      throw new Error(
        'SIN PERMISO DE HOJAS DE CALCULO. | ' +
        'Proyecto que responde: ' + url + ' | ' +
        'Ejecuta como: ' + quien + ' | ' +
        'Usuario activo: ' + activa + ' | ' +
        'ID_LIBRO en el codigo: ' + ID_LIBRO + ' | ' +
        'COMO ARREGLAR: en el editor de Apps Script abre Configuracion del proyecto, ' +
        'marca "Mostrar el archivo de manifiesto appsscript.json", BORRA todo el bloque ' +
        '"oauthScopes" (o pon solo https://www.googleapis.com/auth/spreadsheets y ' +
        'https://www.googleapis.com/auth/drive.file), pulsa Guardar, ' +
        'DESPUES ejecuta cualquier funcion desde el EDITOR y acepta la pantalla de ' +
        'autorizacion de Google (este paso es obligatorio), y solo entonces ve a ' +
        'Implementar > Nueva implementacion. Detalle original: ' + msg
      );
    }
    throw err;
  }
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

  // Usuarios de turnos DISPONIBLE_* -> solo index_1.
  // Se aceptan variantes: DISPONIBLE_A, "DISPONIBLE A", DISPONIBLE-A.
  if (/^DISPONIBLE[\s_\-]*[ABC]?$/.test(id)) return ['index_1'];

  // Usuarios del módulo de áreas JESEP -> solo index_2.
  // UBL_JESEP no va aquí: entra a index_1 (turnos), donde administra
  // el LISTADO_BASE desde la sección Funcionarios.
  // Se comprueba con hasOwnProperty para que un usuario llamado
  // "CONSTRUCTOR" o "TOSTRING" no herede una propiedad del objeto.
  var usuariosArea = {
    'SGSST_JESEP': 1, 'VAC_JESEP': 1, 'PAS_JESEP': 1, 'CIT_JESEP': 1,
    'HIS_JESEP': 1, 'PRO_JESEP': 1, 'GH_JESEP': 1
  };
  if (Object.prototype.hasOwnProperty.call(usuariosArea, id)) return ['index_2'];

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
      case 'iniciarSesionCompleta':
        resultado = iniciarSesionCompleta(args[0], args[1]);
        break;
      case 'diagnosticarCumpleanos':
        resultado = diagnosticarCumpleanos();
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
      case 'diagnosticarCruce':
        resultado = diagnosticarCruceDataSafe();
        break;
      case 'cumpleanerosHoy':
        resultado = cumpleanerosHoy(args[0]);
        break;
      case 'consultarPresentacionesHoy':
        resultado = consultarPresentacionesHoy(args[0], args[1]);
        break;
      case 'confirmarPresentacionesHoy':
        resultado = confirmarPresentacionesHoy(args[0], args[1]);
        break;
      case 'presentarVacacionesHoy':
        resultado = consultarPresentacionesHoy(args[0], '');
        break;
      case 'registrarNovedadRapida':
        resultado = registrarNovedadRapida(args[0], args[1], args[2]);
        break;
      case 'version':
        resultado = versionApp();
        break;
      case 'diagnosticarUsuarios':
        resultado = diagnosticarUsuarios();
        break;
      case 'listarNovedades':
        resultado = listarNovedades(args[0], args[1], args[2]);
        break;
      case 'diagnosticarCarga':
        resultado = diagnosticarCarga(args[0]);
        break;
      case 'eliminarNovedad':
        resultado = eliminarNovedad(args[0], args[1], args[2], args[3]);
        break;
      default:
        resultado = {
          estado: false,
          version: VERSION_APP,
          mensaje: (accion ? ('Acción no válida: ' + accion)
                           : 'Falta el parametro "accion". Abra el Web App con ?accion=<nombre>.'),
          accionesValidas: [
            'version', 'diagnosticarEsquema', 'diagnosticarCruce',
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

/* =====================================================
   FECHAS UTILES PARA EL INDICE DE NOVEDADES
   La fecha de presentacion YA esta calculada en la hoja NOVEDADES
   (columna K). Aqui solo se cuenta cuanto falta desde hoy.
   ===================================================== */

var DIAS_SEMANA_ = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];

/* Devuelve una fecha legible dd/mm/aaaa, o cadena vacia si no hay fecha. */
function fechaComoTexto_(v) {
  var p = partesFecha_(v, ZONA_);
  if (!p || !p.anio) return '';
  function dos(n) { return (n < 10 ? '0' : '') + n; }
  return dos(p.dia) + '/' + dos(p.mes) + '/' + p.anio;
}

/* Dias de calendario entre hoy y una fecha, y el nombre del dia.
   Devuelve estado VENCE cuando ya paso, HOY cuando es el mismo dia. */
function diasRestantesHasta_(v) {
  var p = partesFecha_(v, ZONA_);
  if (!p || !p.anio) {
    return { fecha: '', restantes: '', diaSemana: '', estado: 'SIN_FECHA' };
  }

  var hoy = partesFecha_(new Date(), ZONA_);

  // Medianoche de ambas fechas: evita que las horas del turno resten dias.
  var utcHoy = Date.UTC(hoy.anio, hoy.mes - 1, hoy.dia);
  var utcDestino = Date.UTC(p.anio, p.mes - 1, p.dia);
  var restantes = Math.round((utcDestino - utcHoy) / 86400000);

  var estado = restantes < 0 ? 'VENCIDO' : (restantes === 0 ? 'HOY' : 'PENDIENTE');

  // getUTCDay(): la fecha ya esta desplazada a UTC en la cuenta anterior.
  var diaSemana = DIAS_SEMANA_[(new Date(utcDestino).getUTCDay() + 7) % 7];

  return {
    fecha: fechaComoTexto_(v),
    restantes: restantes,
    diaSemana: diaSemana,
    estado: estado
  };
}

/* Usuarios que NO pueden administrar usuarios aunque su ROL en la hoja
   USUARIOS sea ADMINISTRADOR. Es una lista negra explicita para que el
   bloqueo no dependa de editar la hoja y no se pierda en una carga. */
var USUARIOS_SIN_ADMIN_USUARIOS_ = ['OFSEMANA'];

function esUsuarioBloqueadoAdmin_(usuario) {
  var id = normalizarUsuario_(usuario);
  // Sin identidad no se concede nada: es lo seguro ante un parametro vacio.
  if (!id) return true;
  return USUARIOS_SIN_ADMIN_USUARIOS_.some(function(u) {
    return normalizarUsuario_(u) === id;
  });
}

/* Unico punto de decision para administrar usuarios: rol real ADMINISTRADOR
   leido de la hoja USUARIOS y ausencia en la lista negra.
   Las funciones de la seccion Usuarios lo usan para validar en el servidor:
   ocultar el boton en el cliente no es proteccion, porque la URL es publica. */
function puedeAdministrarUsuarios_(usuario) {
  if (esUsuarioBloqueadoAdmin_(usuario)) return false;
  return esUsuarioAdmin_(usuario);
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
  var cedulaIn = normalizaCedula_(cedula);
  if (!cedulaIn) return null;

  for (var i = 1; i < datos.length; i++) {
    if (normalizaCedula_(datos[i][idxCedula]) === cedulaIn) {
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

  // Todas las columnas se localizan por su encabezado real en la hoja NOVEDADES:
//   A CEDULA | B GR | C NIV | D DEPENDENCIA | E TURNO | F FUNCIONARIO
//   G TIPO | H NOVEDAD | I Dias | J Fecha INICIAL | K Fecha PRESENTACION
//   L Observacion | M Placa_Chip | N TEXTO
// El numero entre parentesis es la posicion de respaldo por si el titulo falta.
return {
    filaCabecera: fCab,
    cabeceras: cabeceras,
    cabNorm: cab,
    numCols: hoja.getLastColumn(),
    cedula: idxColumnaBase_(cab, [/^CEDULA$/, /^CC$/], 0),              // A
    grado: idxColumnaBase_(cab, [/^GR$/, /^GRADO$/], 1),                // B
    nivel: idxColumnaBase_(cab, [/^NIV$/, /^NIVEL$/], 2),               // C
    dependencia: idxColumnaBase_(cab, [/^DEPENDENCIA$/], 3),            // D
    turno: idxColumnaBase_(cab, [/^TURNO$/], 4),                        // E
    funcionario: idxColumnaBase_(cab, [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/], 5), // F
    tipo: idxColumnaBase_(cab, [/^TIPO$/], 6),                          // G
    novedad: idxColumnaBase_(cab, [/^NOVEDAD$/], 7),                   // H
    dias: idxColumnaBase_(cab, [/^DIAS$/], 8),                         // I
    fechaInicial: idxColumnaBase_(cab, [/^FECHA INICIAL$/], 9),         // J
    fechaPresentacion: idxColumnaBase_(cab, [/^FECHA PRESENTACION$/], 10), // K
    observacion: idxColumnaBase_(cab, [/^OBSERVACION$/], 11),           // L
    placaChip: idxColumnaBase_(cab, [/^PLACA ?CHIP$/], 12),             // M
    texto: idxColumnaBase_(cab, [/^TEXTO$/], 13),                       // N
    fechaRegistro: idxColumnaBase_(cab, [/^FECHA REGISTRO$/, /^FECHA DE REGISTRO$/], -1)
  };
}

function construirFilaNovedades_(m, datos) {
  var fila = [];
  for (var i = 0; i < m.numCols; i++) fila.push('');

  var tipo = String(datos.tipo || '').trim();
  var detalle = String(datos.novedad || '').trim();
  var funcionario = String(datos.nombreFuncionario || '').trim();

  // A = cédula
  if (m.cedula >= 0) fila[m.cedula] = datos.cedula || datos.cc || '';

  // F = FUNCIONARIO (el seleccionado en la búsqueda)
  if (m.funcionario >= 0) fila[m.funcionario] = funcionario;

  // G = TIPO
  if (m.tipo >= 0) fila[m.tipo] = tipo;

  // H = NOVEDAD (texto escrito por el usuario)
  if (m.novedad >= 0) fila[m.novedad] = detalle;

  // I / J / K / L / N
  if (m.dias >= 0) fila[m.dias] = (datos.dias === undefined || datos.dias === null) ? '' : datos.dias;
  if (m.fechaInicial >= 0) fila[m.fechaInicial] = datos.fechaInicial || '';
  if (m.fechaPresentacion >= 0) fila[m.fechaPresentacion] = datos.fechaPresentacion || '';
  if (m.observacion >= 0) fila[m.observacion] = datos.observacion || datos.descripcion || '';
  if (m.texto >= 0) fila[m.texto] = datos.texto || detalle || datos.descripcion || '';
  if (m.fechaRegistro >= 0) fila[m.fechaRegistro] = new Date();

  return fila;
}

// Rellena las columnas que no se escribieron a mano con los datos del
// funcionario tomados de LISTADO_BASE. En la hoja NOVEDADES quedan
// cubiertos GR, NIV, DEPENDENCIA, TURNO y Placa_Chip.
function completarDesdeBase_(fila, m, base) {
  if (!base) return;

  var fijas = {};
  [m.cedula, m.funcionario, m.tipo, m.novedad, m.dias, m.fechaInicial,
   m.fechaPresentacion, m.observacion, m.texto, m.fechaRegistro].forEach(function(k) {
    if (k !== undefined && k >= 0) fijas[k] = true;
  });

  var cabNormBase = base.cabeceras.map(function(c) {
    return String(c).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s_\-]+/g, ' ').trim();
  });

  for (var c = 0; c < m.numCols; c++) {
    if (fijas[c]) continue;

    var nombreCol = m.cabNorm[c];
    if (!nombreCol) continue;

    var idxBase = indiceBasePorAlias_(base.cabeceras, nombreCol);
    if (idxBase === -1) idxBase = cabNormBase.indexOf(nombreCol);
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
      hoja.appendRow(['CEDULA', 'GR', 'NIV', 'DEPENDENCIA', 'TURNO', 'FUNCIONARIO', 'TIPO', 'NOVEDAD', 'Dias', 'Fecha INICIAL', 'Fecha PRESENTACION', 'Observacion', 'Placa_Chip', 'TEXTO']);
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
    var cabNormBase = cabecerasBase.map(normalizarCabColumna_);

    // Índices reales de cada campo. Se buscan por encabezado normalizado para
    // que funcionen tanto "CEDULA" como "Cédula", "GR"/"GRADO", etc.
    var iCedula  = idxColumnaBase_(cabNormBase, [/^CEDULA$/, /^CC$/, /^NUMERO DE CEDULA$/, /^N C$/], 0);
    var iTurno   = idxColumnaBase_(cabNormBase, [/^TURNO$/], 7);
    var iGrado   = idxColumnaBase_(cabNormBase, [/^GR$/, /^GRADO$/], -1);
    var iNombre  = idxColumnaBase_(cabNormBase, [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/, /^NOMBRES$/], -1);
    var iDepend  = idxColumnaBase_(cabNormBase, [/^DEPENDENCIA$/], -1);
    // Fecha del ultimo ascenso. El encabezado exacto no es uniforme entre
    // copias de la hoja, asi que se aceptan las variantes conocidas.
    var iAscenso = idxColumnaBase_(cabNormBase,
      [/^FECHA ULTIMO ASCENSO$/, /^FECHA DE ULTIMO ASCENSO$/, /^ULTIMO ASCENSO$/,
       /^FECHA ASCENSO$/, /^FECHA DE ASCENSO$/, /^FECHA_ASCENSO$/], -1);

    function valCampo_(fila, idx) {
      if (idx === -1 || idx === undefined || idx >= fila.length) return '';
      var v = fila[idx];
      return (v === null || v === undefined) ? '' : v;
    }

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
      // Objetos con claves normalizadas, para que el resto del código
      // no dependa de cómo estén escritos los encabezados.
      var filaBase = {
        cedula: valCampo_(datosBase[i], iCedula),
        turno: valCampo_(datosBase[i], iTurno),
        grado: valCampo_(datosBase[i], iGrado),
        funcionario: valCampo_(datosBase[i], iNombre),
        dependencia: valCampo_(datosBase[i], iDepend)
      };
      for (var j2 = 0; j2 < cabecerasBase.length; j2++) {
        var claveNorm = normalizarCabColumna_(cabecerasBase[j2]).replace(/[^A-Z0-9]/g, '');
        if (claveNorm && filaBase[claveNorm.toLowerCase()] === undefined) {
          filaBase[claveNorm.toLowerCase()] = datosBase[i][j2];
        }
      }

      var turnoFila = String(filaBase.turno || '').trim().toUpperCase();
      if (filtro === 'SEPRI' || turnoFila === String(filtro).trim().toUpperCase()) {

        var cedulaFuncionarioNorm = normalizaCedula_(filaBase.cedula);
        var novedadesCruzadas = [];

        // Cruce con la hoja NOVEDADES. Las posiciones corresponden a los
        // titulos reales de esa hoja:
        //   A CEDULA | G TIPO | H NOVEDAD | I Dias
        if (datosNovedades.length > 1 && cedulaFuncionarioNorm !== '') {
          var cabNormNov = (datosNovedades[0] || []).map(normalizarCabColumna_);
          var idxCedulaNov = idxColumnaBase_(cabNormNov, [/^CEDULA$/, /^CC$/], 0);
          var idxTipoNov = idxColumnaBase_(cabNormNov, [/^TIPO$/], 6);
          var idxNovedadNov = idxColumnaBase_(cabNormNov, [/^NOVEDAD$/], 7);
          var idxDiasNov = idxColumnaBase_(cabNormNov, [/^DIAS$/], 8);
          // J = Fecha INICIAL | K = Fecha PRESENTACION. La presentacion ya
          // esta calculada en la hoja: no se vuelve a derivar de los dias.
          var idxFInicNov = idxColumnaBase_(cabNormNov, [/^FECHA INICIAL$/], 9);
          var idxFPresNov = idxColumnaBase_(cabNormNov, [/^FECHA PRESENTACION$/], 10);

          for (var n = 1; n < datosNovedades.length; n++) {
            var filaNov = datosNovedades[n];

            // Se compara la cedula normalizada: tolera formato, separadores
            // y ceros a la izquierda (numero en una hoja, texto en la otra).
            var cedulaNov = normalizaCedula_(filaNov[idxCedulaNov]);
            if (cedulaNov === '' || cedulaNov !== cedulaFuncionarioNorm) continue;

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
              // NOVEDAD = texto concatenado (TIPO - detalle - dias)
              // TIPO = solo el tipo, para las vistas simplificadas
              // El resto lo consume el indice de novedades, que solo se
              // muestra a administradores.
              var presInfo = diasRestantesHasta_(filaNov[idxFPresNov]);
              var inicialInfo = filaNov[idxFInicNov];

              novedadesCruzadas.push({
                NOVEDAD: concatenado,
                TIPO: tipoNov,
                DIAS: diasNov,
                FECHA_INICIAL: fechaComoTexto_(inicialInfo),
                FECHA_PRESENTACION: presInfo.fecha,
                DIAS_RESTANTES: presInfo.restantes,
                DIA_SEMANA: presInfo.diaSemana,
                ESTADO_PRESENTACION: presInfo.estado
              });
            }
          }
        }

        obj.historialNovedades = novedadesCruzadas;
        obj.cedula = filaBase.cedula;
        obj.turno = filaBase.turno;
        obj.grado = filaBase.grado;
        obj.funcionario = filaBase.funcionario;
        obj.fechaAscenso = fechaComoTexto_(iAscenso === -1 ? '' : datosBase[i][iAscenso]);
        obj.dependencia = filaBase.dependencia;
        funcionarios.push(obj);
      }
    }

    return {
      estado: true,
      datos: {
        filtro: filtro,
        funcionarios: funcionarios,
        _diag: {
          iCedula: iCedula + 1,
          iTurno: iTurno + 1,
          iGrado: iGrado + 1,
          iNombre: iNombre + 1,
          iDepend: iDepend + 1,
          filasNovedades: Math.max(0, datosNovedades.length - 1),
          conNovedades: funcionarios.filter(function(f) { return tieneNovedades_(f); }).length,
          muestraCedulas: funcionarios.slice(0, 3).map(function(f) { return normalizaCedula_(f.cedula); })
        }
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
  html += '.grupo td{background:#e8f3ed;font-weight:bold;font-size:9pt;color:#01592F;padding:5px;border:1px solid #c8dfd2;}';
  html += '.grupo td.oscuro{background:#fdf3e3;color:#8a5a00;}';
  html += '.nov{background:#fff6f6;}';
  html += '</style></head><body>';

  var ordenados = ordenarParaReporte_(funcs);
  var sinNov = ordenados.filter(function(f) { return !tieneNovedades_(f); });
  var conNov = ordenados.filter(function(f) { return tieneNovedades_(f); });

  html += '<h1>REPORTE DE TURNO ' + escHtml_(filtro) + '</h1>';
  html += '<div class="meta"><b>Consecutivo:</b> ' + escHtml_(consecutivo) + ' &nbsp;|&nbsp; ';
  html += '<b>Fecha:</b> ' + escHtml_(Utilities.formatDate(new Date(), 'America/Bogota', "dd/MM/yyyy HH:mm")) + ' &nbsp;|&nbsp; ';
  html += '<b>Total:</b> ' + funcs.length +
    ' &nbsp;|&nbsp; <b>Sin novedades:</b> ' + sinNov.length +
    ' &nbsp;|&nbsp; <b>Con novedades:</b> ' + conNov.length + '</div>';
  html += '<table><thead><tr>';
  html += '<th>#</th><th>C&eacute;dula</th><th>GR</th><th>Funcionario</th><th>Dependencia</th><th>Turno</th><th>Novedades</th>';
  html += '</tr></thead><tbody>';

  var n = 0;

  // Sin filas separadoras: solo se resalta el color de los que tienen novedades.
  function filasDe_(lista, conNovedades) {
    for (var i = 0; i < lista.length; i++) {
      var f = lista[i];
      n++;
      var novedadesTexto = (f.historialNovedades || []).map(function(x) { return x.NOVEDAD; }).join(', ');
      html += '<tr' + (conNovedades ? ' class="nov"' : '') + '>';
      html += '<td>' + n + '</td>';
      html += '<td>' + escHtml_(f.cedula || f.CEDULA || '') + '</td>';
      html += '<td>' + escHtml_(f.grado || f.GR || '') + '</td>';
      html += '<td>' + escHtml_(f.funcionario || f.FUNCIONARIO || '') + '</td>';
      html += '<td>' + escHtml_(f.dependencia || f.DEPENDENCIA || '') + '</td>';
      html += '<td>' + escHtml_(f.turno || f.TURNO || '') + '</td>';
      html += '<td>' + escHtml_(novedadesTexto || 'S/N') + '</td>';
      html += '</tr>';
    }
  }

  filasDe_(sinNov, false);
  filasDe_(conNov, true);

  html += '</tbody></table></body></html>';
  return html;
}

// Escapa caracteres especiales para que no rompan el HTML del reporte
function escHtml_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* =====================================================
   ORDEN Y AGRUPACIÓN DEL REPORTE
   Primero los funcionarios SIN novedades, agrupados por GR.
   Al final los que SÍ tienen novedades, también agrupados por GR.
   ===================================================== */

// Normaliza la cédula para comparar sin importar formato,
// separadores ni ceros a la izquierda (numero vs texto).
function normalizaCedula_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v).replace(/[^0-9]/g, '');
  return s.replace(/^0+/, '');
}

function tieneNovedades_(f) {
  return !!(f && f.historialNovedades && f.historialNovedades.length);
}

function gradoDe_(f) {
  return String((f && (f.grado || f.GR)) || '').trim().toUpperCase();
}

// Orden jerárquico de grados exigido por la institución.
var ORDEN_GRADOS_ = [
  'MG', 'BG', 'CR', 'TC', 'MY', 'CT', 'TE', 'ST', 'CM', 'SC',
  'IJ', 'IT', 'SI', 'PT', 'PP', 'AXP'
];

function indiceGrado_(grado) {
  var g = String(grado || '').trim().toUpperCase();
  for (var i = 0; i < ORDEN_GRADOS_.length; i++) {
    if (ORDEN_GRADOS_[i] === g) return i;
  }
  return ORDEN_GRADOS_.length;   // los grados no listados van al final
}

function nombreDe_(f) {
  return String((f && (f.funcionario || f.FUNCIONARIO)) || '').trim().toUpperCase();
}

// Copia ordenada: sin novedades primero; dentro de cada grupo, por la
// jerarquia de grados exigida y luego por nombre.
function ordenarParaReporte_(funcs) {
  return (funcs || []).slice().sort(function(a, b) {
    var na = tieneNovedades_(a) ? 1 : 0;
    var nb = tieneNovedades_(b) ? 1 : 0;
    if (na !== nb) return na - nb;                    // sin novedades arriba
    var ia = indiceGrado_(gradoDe_(a));
    var ib = indiceGrado_(gradoDe_(b));
    if (ia !== ib) return ia - ib;                    // jerarquia de grados
    var xa = nombreDe_(a), xb = nombreDe_(b);
    if (xa !== xb) return xa < xb ? -1 : 1;           // por nombre
    return 0;
  });
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
  // Mismo orden que el PDF: sin novedades primero, agrupado por GR.
  var ordenados = ordenarParaReporte_(funcs);

  var filas = [];
  var n = 0;

  ordenados.forEach(function(f) {
    n++;
    var novedadesTexto = (f.historialNovedades || []).map(function(x) { return x.NOVEDAD; }).join(', ');
    filas.push([
      n,
      f.cedula || f.CEDULA || '',
      f.grado || f.GR || '',
      f.funcionario || f.FUNCIONARIO || '',
      f.dependencia || f.DEPENDENCIA || '',
      f.turno || f.TURNO || '',
      novedadesTexto || 'S/N'
    ]);
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

// Muestra la hoja USUARIOS con el módulo que le corresponde a cada cuenta.
// Sirve para verificar por que un usuario entra al módulo equivocado.
// No requiere contraseña: solo abre la URL con ?accion=diagnosticarUsuarios
function diagnosticarUsuarios() {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) {
      return { estado: false, version: VERSION_APP, mensaje: 'No existe la hoja USUARIOS.' };
    }

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) {
      return { estado: true, version: VERSION_APP, mensaje: 'La hoja USUARIOS no tiene filas.', datos: [] };
    }

    var cab = (datos[0] || []).map(normalizarCabColumna_);
    var iUser = idxColumnaBase_(cab, [/^USUARIO$/, /^USUARIOS?$/, /^LOGIN$/, /^NOMBRE$/], 0);
    var iRol = idxColumnaBase_(cab, [/^ROL$/, /^PERFIL$/, /^PERMISO$/], 1);
    var iTurno = idxColumnaBase_(cab, [/^TURNO$/, /^TURNO PERMITIDO$/], -1);

    var lista = [];
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      var nombre = iUser >= 0 ? String((fila[iUser] === null ? '' : fila[iUser])) : '';
      if (!nombre.trim()) continue;

      var rol = iRol >= 0 ? String((fila[iRol] === null ? '' : fila[iRol])) : '';
      var id = nombre.trim().toUpperCase();
      var turno = iTurno >= 0 ? String((fila[iTurno] === null ? '' : fila[iTurno])) : '';

      if (!turno.trim()) {
        if (id.indexOf('DISPONIBLE_A') !== -1) turno = 'A (inferido)';
        else if (id.indexOf('DISPONIBLE_B') !== -1) turno = 'B (inferido)';
        else if (id.indexOf('DISPONIBLE_C') !== -1) turno = 'C (inferido)';
      }

      var modulos = modulosDeUsuario_(id, rol);
      lista.push({
        USUARIO: nombre,
        ROL: rol,
        TURNO: turno,
        MODULOS: modulos.join(' + '),
        DESTINO: modulos.length === 1
          ? (modulos[0] === 'index_1' ? 'Turnos' : 'Áreas')
          : 'Selector de módulo'
      });
    }

    return {
      estado: true,
      version: VERSION_APP,
      total: lista.length,
      mensaje: 'Usuarios en la hoja USUARIOS. ' +
        'Si un usuario tipo DISPONIBLE aparece en "Áreas", su nombre no esta escrito como DISPONIBLE_A/B/C.',
      datos: lista
    };
  } catch (err) {
    return { estado: false, version: VERSION_APP, mensaje: 'Error en diagnostico de usuarios: ' + err.message };
  }
}

function listarUsuarios(token, quienLlama) {
  try {
    if (!puedeAdministrarUsuarios_(quienLlama)) {
      return { estado: false, mensaje: 'Su usuario no tiene permiso para administrar usuarios.' };
    }
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

function crearUsuario(token, quienLlama, usuario, clave, rol, dependencia) {
  try {
    if (!puedeAdministrarUsuarios_(quienLlama)) {
      return { estado: false, mensaje: 'Su usuario no tiene permiso para crear usuarios.' };
    }
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

function cambiarEstadoUsuario(token, quienLlama, usuario, nuevoEstado) {
  try {
    if (!puedeAdministrarUsuarios_(quienLlama)) {
      return { estado: false, mensaje: 'Su usuario no tiene permiso para cambiar el estado de usuarios.' };
    }
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
      // La columna MES pasa a guardar la fecha completa de presentacion.
      // El campo DIA se elimino del formulario y ya no se escribe.
      { valor: datos.fechaPresentacion, patrones: [/^MES$/, /^MES ANO$/, /^MESES ANO$/, /^FECHA PRESENTACION$/] },
      { valor: datos.fechaNacimiento,   patrones: [/^FECHA NACIMIENTO$/, /^FECHA DE NACIMIENTO$/, /^FNAC$/] },
      { valor: datos.correo,            patrones: [/^CORREO ELECTRONICO$/, /^CORREO$/, /^EMAIL$/] },
      { valor: datos.sexo,              patrones: [/^SEXO$/] },
      { valor: datos.estadoCivil,       patrones: [/^ESTADO CIVIL$/] },
      { valor: datos.situacionLaboral,  patrones: [/^SITUACION LABORAL$/] },
      { valor: datos.comunicado,        patrones: [/^COMUNICADO OFICIAL$/, /^COMUNICADO$/] },
      { valor: datos.fechaUltimoAscenso, patrones: [/^FECHA ULTIMO ASCENSO$/, /^FECHA DE ULTIMO ASCENSO$/, /^ULTIMO ASCENSO$/, /^FECHA ASCENSO$/, /^FECHA DE ASCENSO$/, /^FECHA_ASCENSO$/] },
      { valor: datos.numeroCurso,        patrones: [/^N CURSO$/, /^NUMERO DE CURSO$/, /^NUM CURSO$/, /^NCURSO$/, /^CURSO$/] }
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
   12. GESTION DE NOVEDADES (USUARIOS DE AREAS)
   SGSST_JESEP, VAC_JESEP, PAS_JESEP, CIT_JESEP, HIS_JESEP,
   PRO_JESEP y GH_JESEP pueden consultar y eliminar filas
   de la hoja NOVEDADES.
   ===================================================== */

// Verifica que el usuario exista en USUARIOS y sea de areas (o administrador).
// Acepta el libro ya abierto para no volver a abrirlo.
function esUsuarioArea_(usuario, libroPrevio) {
  var id = normalizarUsuario_(usuario);
  if (!id) return false;

  var areas = {
    'SGSST_JESEP': 1, 'VAC_JESEP': 1, 'PAS_JESEP': 1, 'CIT_JESEP': 1,
    'HIS_JESEP': 1, 'PRO_JESEP': 1, 'GH_JESEP': 1
  };

  try {
    var ss = libroPrevio || abrirLibro_();
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
        if (rol === 'ADMINISTRADOR' || rol === 'ADMIN') return true;
        return Object.prototype.hasOwnProperty.call(areas, id);
      }
    }
  } catch (e) {
    return false;
  }
  return false;
}

// Convierte el valor de una celda de fecha a texto dd/mm/aaaa.
function formatearFechaCelda_(v) {
  var p = partesFecha_(v, ZONA_);
  if (!p) return (v === null || v === undefined) ? '' : String(v);
  var dd = p.dia < 10 ? '0' + p.dia : String(p.dia);
  var mm = p.mes < 10 ? '0' + p.mes : String(p.mes);
  return dd + '/' + mm + '/' + p.anio;
}

// Lista las filas de NOVEDADES con su numero real dentro de la hoja,
// para poder borrarlas de forma precisa.
//
// Recibe un objeto de opciones para no enviar nunca toda la hoja:
//   { pagina, porPagina, cedula, tipo, turno }
//
// Devuelve solo la pagina solicitada, el total de coincidencias y el
// conteo por tipo (para pintar los filtros sin pedir otra consulta).
function listarNovedades(token, usuario, opciones) {
  try {
    var op = (opciones && typeof opciones === 'object') ? opciones : {};

    var pagina = Math.max(1, parseInt(op.pagina, 10) || 1);
    var porPagina = Math.min(200, Math.max(5, parseInt(op.porPagina, 10) || 25));
    var fCedula = String(op.cedula || '').trim();
    var fTipo = String(op.tipo || '').trim().toUpperCase();
    var fTurno = String(op.turno || '').trim().toUpperCase();

    // Se abre el libro una sola vez y se reutiliza para las dos hojas.
    var ss = abrirLibro_();

    if (!esUsuarioArea_(usuario, ss)) {
      return { estado: false, mensaje: 'Su usuario no puede gestionar las novedades.' };
    }

    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja NOVEDADES.' };

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) {
      return {
        estado: true, mensaje: 'La hoja NOVEDADES no tiene registros.',
        total: 0, pagina: 1, paginas: 0, porPagina: porPagina, datos: [], porTipo: {}
      };
    }

    var cab = (datos[0] || []).map(normalizarCabColumna_);
    var iCed = idxColumnaBase_(cab, [/^CEDULA$/, /^CC$/], 0);
    var iNom = idxColumnaBase_(cab, [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/], -1);
    var iGr = idxColumnaBase_(cab, [/^GR$/, /^GRADO$/], -1);
    var iTur = idxColumnaBase_(cab, [/^TURNO$/], -1);
    var iTipo = idxColumnaBase_(cab, [/^TIPO$/], -1);
    var iNov = idxColumnaBase_(cab, [/^NOVEDAD$/], -1);
    var iDias = idxColumnaBase_(cab, [/^DIAS$/], -1);
    var iFIni = idxColumnaBase_(cab, [/^FECHA INICIAL$/], -1);
    var iFPre = idxColumnaBase_(cab, [/^FECHA PRESENTACION$/], -1);

    var busqueda = fCedula.toUpperCase();
    var coincidencias = [];
    var porTipo = {};
    var turnosVistos = {};
    var totalHoja = 0;

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];

      // Ignora filas totalmente vacias
      var tieneAlgo = false;
      for (var c = 0; c < fila.length; c++) {
        if (fila[c] !== null && fila[c] !== undefined && String(fila[c]).trim() !== '') {
          tieneAlgo = true; break;
        }
      }
      if (!tieneAlgo) continue;
      totalHoja++;

      var sCed = iCed >= 0 ? String(fila[iCed] === null ? '' : fila[iCed]) : '';
      var sNom = iNom >= 0 ? String(fila[iNom] === null ? '' : fila[iNom]) : '';
      var sTur = iTur >= 0 ? String(fila[iTur] === null ? '' : fila[iTur]) : '';
      var sTipo = iTipo >= 0 ? String(fila[iTipo] === null ? '' : fila[iTipo]) : '';
      var sNov = iNov >= 0 ? String(fila[iNov] === null ? '' : fila[iNov]) : '';
      var sGr = iGr >= 0 ? String(fila[iGr] === null ? '' : fila[iGr]) : '';

      // Los conteos por tipo y turno se hacen sobre toda la hoja para
      // que los filtros reflejen el total real, no solo la pagina visible.
      if (sTipo) porTipo[sTipo] = (porTipo[sTipo] || 0) + 1;
      if (sTur) turnosVistos[sTur.trim().toUpperCase()] = 1;

      // Filtros
      if (fTipo && sTipo.trim().toUpperCase() !== fTipo) continue;
      if (fTurno && sTur.trim().toUpperCase() !== fTurno) continue;
      if (busqueda) {
        var heno = (sCed + ' ' + sNom + ' ' + sTipo + ' ' + sNov + ' ' + sGr).toUpperCase();
        if (heno.indexOf(busqueda) === -1) continue;
      }

      coincidencias.push({
        fila: i + 1,
        cedula: sCed,
        funcionario: sNom,
        grado: sGr,
        turno: sTur,
        tipo: sTipo,
        novedad: sNov,
        dias: iDias >= 0 ? String(fila[iDias] === null ? '' : fila[iDias]) : '',
        fechaInicial: iFIni >= 0 ? formatearFechaCelda_(fila[iFIni]) : '',
        fechaPresentacion: iFPre >= 0 ? formatearFechaCelda_(fila[iFPre]) : '',
        sello: sCed + '|' + (iTipo >= 0 ? fila[iTipo] : '') + '|' + (iNov >= 0 ? fila[iNov] : '')
      });
    }

    var total = coincidencias.length;
    var paginas = Math.ceil(total / porPagina);
    if (paginas === 0) paginas = 0;
    if (pagina > paginas && paginas > 0) pagina = paginas;

    var desde = (pagina - 1) * porPagina;
    var paginaDatos = coincidencias.slice(desde, desde + porPagina);

    return {
      estado: true,
      total: total,
      totalHoja: totalHoja,
      pagina: pagina,
      paginas: paginas,
      porPagina: porPagina,
      tipos: Object.keys(porTipo).sort(),
      porTipo: porTipo,
      turnos: Object.keys(turnosVistos).sort(),
      datos: paginaDatos
    };
  } catch (err) {
    return { estado: false, mensaje: 'Error listando novedades: ' + err.message };
  }
}

// Elimina una fila de NOVEDADES.
// Se exige el "sello" (cedula + tipo + novedad) para confirmar que la fila
// sigue siendo la misma que el usuario vio. Si la lista cambio entre la
// consulta y el borrado, se rechaza en vez de borrar el registro equivocado.
function eliminarNovedad(token, usuario, fila, sello) {
  try {
    var nFila = parseInt(fila, 10);
    if (isNaN(nFila) || nFila < 2) {
      return { estado: false, mensaje: 'Fila no válida.' };
    }

    var ss = abrirLibro_();

    if (!esUsuarioArea_(usuario, ss)) {
      return { estado: false, mensaje: 'Su usuario no puede eliminar novedades.' };
    }

    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja NOVEDADES.' };

    var datos = hoja.getDataRange().getValues();
    if (nFila > datos.length) {
      return { estado: false, mensaje: 'La fila ya no existe. Actualice la lista.' };
    }

    if (sello) {
      var cab = (datos[0] || []).map(normalizarCabColumna_);
      var iCed = idxColumnaBase_(cab, [/^CEDULA$/, /^CC$/], 0);
      var iTipo = idxColumnaBase_(cab, [/^TIPO$/], -1);
      var iNov = idxColumnaBase_(cab, [/^NOVEDAD$/], -1);
      var f = datos[nFila - 1];
      var actual = (iCed >= 0 ? String(f[iCed] === null ? '' : f[iCed]) : '') + '|' +
                   (iTipo >= 0 ? f[iTipo] : '') + '|' +
                   (iNov >= 0 ? f[iNov] : '');
      if (String(actual) !== String(sello)) {
        return {
          estado: false,
          mensaje: 'La lista cambió. Actualice y vuelva a intentarlo, para no borrar el registro equivocado.'
        };
      }
    }

    hoja.deleteRow(nFila);
    return { estado: true, mensaje: 'Novedad eliminada de la hoja NOVEDADES.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error eliminando novedad: ' + err.message };
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
      hoja.appendRow(['CEDULA', 'GR', 'NIV', 'DEPENDENCIA', 'TURNO', 'FUNCIONARIO', 'TIPO', 'NOVEDAD', 'Dias', 'Fecha INICIAL', 'Fecha PRESENTACION', 'Observacion', 'Placa_Chip', 'TEXTO']);
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

/* =====================================================
   9. AVISOS: CUMPLEAÑOS Y PRESENTACIÓN DE VACACIONES
   ===================================================== */

var ZONA_ = 'America/Bogota';

// Convierte un valor de celda (Date o texto) en año/mes/día.
function partesFecha_(v, zona) {
  if (v === null || v === undefined || v === '') return null;

  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    // Una sola llamada a formatDate en lugar de tres: al listar muchas
    // filas esta funcion se ejecuta cientos de veces y cada llamada
    // tiene un costo medible.
    var txt = Utilities.formatDate(v, zona || ZONA_, 'yyyy-MM-dd');
    var p = txt.split('-');
    return { anio: Number(p[0]), mes: Number(p[1]), dia: Number(p[2]) };
  }

  // Numero de serie de Sheets. getValues() devuelve Date en celdas con
  // formato de fecha, pero si la columna esta formateada como texto plano
  // puede llegar el numero crudo.
  if (typeof v === 'number' && isFinite(v)) {
    if (v < 1 || v > 80000) return null;
    var fs = Utilities.formatDate(new Date(v * 86400000), zona || ZONA_, 'yyyy-MM-dd');
    var q = fs.split('-');
    return { anio: Number(q[0]), mes: Number(q[1]), dia: Number(q[2]) };
  }

  var s = String(v).trim();
  var m = s.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
  if (!m) return null;

  var a = Number(m[1]);
  var b = Number(m[2]);
  var c = Number(m[3]);

  // yyyy-mm-dd: el anio va primero y es el unico grupo de 4 digitos.
  if (a > 31) return { anio: a, mes: b, dia: c };

  // dd-mm-aaaa: el anio va al final (formato que usa la institucion).
  if (c > 31) {
    if (c < 100) c += 2000;
    return { anio: c, mes: b, dia: a };
  }

  // Ambos grupos son de 1-2 digitos (dd-mm-yy): "15-08-90" es ambiguo,
  // puede ser 1990 o 2090. Se asume el siglo pasado, que es lo esperable
  // para una fecha de nacimiento.
  if (c < 100) c += 1900;
  return { anio: c, mes: b, dia: a };
}

function quitarAcentos_(s) {
  return String(s === null || s === undefined ? '' : s)
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

// Funcionarios que cumplen años hoy. Revisa FECHA_NACIMIENTO en LISTADO_BASE.
/* Calcula los cumpleaños de hoy sobre un libro YA abierto.
   Se separa de cumpleanerosHoy para que iniciarSesionCompleta pueda
   reutilizar la misma conexion en vez de abrirla dos veces. */
function cumpleanerosEnLibro_(ss) {
  var hoja = ss.getSheetByName('LISTADO_BASE');
  if (!hoja) return { fecha: '', lista: [] };

  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return { fecha: '', lista: [] };

  var cabeceras = datos[0];
  var cab = cabeceras.map(normalizarCabColumna_);
  var iCed = idxColumnaBase_(cab, [/^CEDULA$/, /^CC$/, /^NUMERO DE CEDULA$/], 0);
  var iNom = idxColumnaBase_(cab, [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/, /^NOMBRES$/], -1);
  var iGr = idxColumnaBase_(cab, [/^GR$/, /^GRADO$/], -1);
  var iDep = idxColumnaBase_(cab, [/^DEPENDENCIA$/], -1);
  // Columna M de LISTADO_BASE
  var iNac = idxColumnaBase_(cab, [/^FECHA NACIMIENTO$/, /^FECHA DE NACIMIENTO$/, /^FNAC$/], 12);

  var hoy = partesFecha_(new Date(), ZONA_);
  var lista = [];

  for (var i = 1; i < datos.length; i++) {
    var nac = partesFecha_(datos[i][iNac], ZONA_);
    if (!nac) continue;
    if (!nac.anio) continue;                     // sin año: no se puede calcular
    if (nac.mes !== hoy.mes || nac.dia !== hoy.dia) continue;

    var anios = hoy.anio - nac.anio;
    lista.push({
      cedula: String(datos[i][iCed] === null ? '' : datos[i][iCed]),
      funcionario: iNom >= 0 ? String(datos[i][iNom] || '') : '',
      grado: iGr >= 0 ? String(datos[i][iGr] || '') : '',
      dependencia: iDep >= 0 ? String(datos[i][iDep] || '') : '',
      anios: anios,
      cumpleanios: (anios % 10 === 0 && (anios % 100 !== 0)) ? '¡Feliz cumpleaños redondo!' : ''
    });
  }

  return { fecha: hoy.dia + '/' + hoy.mes + '/' + hoy.anio, lista: lista };
}

function cumpleanerosHoy(token) {
  try {
    var r = cumpleanerosEnLibro_(abrirLibro_());
    return { estado: true, fecha: r.fecha, datos: r.lista };
  } catch (err) {
    return { estado: false, mensaje: 'Error buscando cumpleaños: ' + err.message };
  }
}

/* Validacion + cumpleaños en un SOLO viaje de red.
   En redes lentas o moviles cada llamada al backend puede tardar 20-35 s y la
   conexion se cae antes de responder. Reducir el login de dos viajes a uno
   elimina de raiz esa exposicion. */
function iniciarSesionCompleta(usuarioIngresado, claveIngresada) {
  var r = validarUsuario(usuarioIngresado, claveIngresada);
  if (!r || r.estado !== true) return r;   // si la clave falla, no se calcula nada

  try {
    var c = cumpleanerosEnLibro_(abrirLibro_());
    r.cumpleanos = { fecha: c.fecha, datos: c.lista };
  } catch (err) {
    r.cumpleanos = { fecha: '', datos: [] };   // el login nunca debe fallar por esto
  }
  return r;
}

/* Diagnostico de la columna FECHA_NACIMIENTO (M de LISTADO_BASE).
   Reporta como viene cada celda realmente: fecha nativa, texto o numero.
   Sin esto no hay forma de saber por que los cumpleaños no aparecen. */
function diagnosticarCumpleanos() {
  var add = [];
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'No existe la hoja LISTADO_BASE.' };

    var datos = hoja.getDataRange().getValues();
    var cabeceras = datos[0] || [];

    add.push('Libro: ' + ss.getName());
    add.push('LISTADO_BASE: ' + datos.length + ' filas x ' + cabeceras.length + ' columnas');

    var cab = cabeceras.map(normalizarCabColumna_);
    var iNac = idxColumnaBase_(cab, [/^FECHA NACIMIENTO$/, /^FECHA DE NACIMIENTO$/, /^FNAC$/], -1);

    add.push('Columna FECHA_NACIMIENTO: indice ' + iNac +
      (iNac >= 0 ? ' -> ' + letraColumna_(iNac + 1) + '  titulo="' + cabeceras[iNac] + '"' : '  NO ENCONTRADA'));
    add.push('Indice fijo de la columna M: 12 (0-based)');

    var hoy = partesFecha_(new Date(), ZONA_);
    add.push('Fecha de hoy: ' + hoy.dia + '/' + hoy.mes + '/' + hoy.anio);

    if (iNac < 0) {
      add.push('');
      add.push('Encabezados reales de la hoja:');
      for (var c = 0; c < cabeceras.length; c++) {
        add.push('  ' + letraColumna_(c + 1) + ') "' + cabeceras[c] + '"');
      }
      return { estado: true, detalle: add.join('\n') };
    }

    var tipos = {}, ok = 0, vacias = 0, muestras = [];
    for (var i = 1; i < datos.length; i++) {
      var v = datos[i][iNac];
      if (v === null || v === undefined || String(v).trim() === '') { vacias++; continue; }
      var t = Object.prototype.toString.call(v);
      tipos[t] = (tipos[t] || 0) + 1;
      if (partesFecha_(v, ZONA_)) ok++;
      if (muestras.length < 10) {
        muestras.push('  fila ' + (i + 1) + ': ' + t.replace('[object ', '').replace(']', '') +
          '  crudo="' + String(v) + '"  ->  ' + JSON.stringify(partesFecha_(v, ZONA_)));
      }
    }

    add.push('');
    add.push('Tipos de valor encontrados:');
    for (var k in tipos) {
      if (Object.prototype.hasOwnProperty.call(tipos, k)) {
        add.push('  ' + k.replace('[object ', '').replace(']', '') + ': ' + tipos[k] + ' celdas');
      }
    }
    add.push('Celdas vacias: ' + vacias);
    add.push('Celdas que se pudieron interpretar como fecha: ' + ok);

    add.push('');
    add.push('Primeras 10 celdas con su interpretacion:');
    if (!muestras.length) add.push('  (no hay ninguna celda con dato)');
    muestras.forEach(function(m) { add.push(m); });

    var r = cumpleanerosEnLibro_(ss);
    add.push('');
    add.push('CUMPLEANEROS DETECTADOS HOY: ' + r.lista.length);
    r.lista.slice(0, 15).forEach(function(p) {
      add.push('  ' + p.funcionario + ' - ' + p.anios + ' anios');
    });

    return { estado: true, detalle: add.join('\n') };
  } catch (err) {
    return { estado: false, mensaje: err.message, detalle: add.join('\n') };
  }
}

// Funcionarios que deben presentarse HOY de vacaciones o permiso.
// Solo lectura: no escribe nada en ninguna hoja.
// Revisa NOVEDADES: K = Fecha PRESENTACION igual a hoy y
// G = TIPO con los conceptos VACACIONES o PERMISO.
function consultarPresentacionesHoy(token, filtro) {
  try {
    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) return { estado: true, datos: [] };

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: true, datos: [] };

    var cab = (datos[0] || []).map(normalizarCabColumna_);
    var iCed = idxColumnaBase_(cab, [/^CEDULA$/, /^CC$/], 0);
    var iNom = idxColumnaBase_(cab, [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/], -1);
    var iGr = idxColumnaBase_(cab, [/^GR$/, /^GRADO$/], -1);
    var iTur = idxColumnaBase_(cab, [/^TURNO$/], -1);
    var iTipo = idxColumnaBase_(cab, [/^TIPO$/], -1);
    var iNov = idxColumnaBase_(cab, [/^NOVEDAD$/], -1);
    var iDias = idxColumnaBase_(cab, [/^DIAS$/], -1);
    var iFPre = idxColumnaBase_(cab, [/^FECHA PRESENTACION$/], -1);

    var hoy = partesFecha_(new Date(), ZONA_);
    var filtroTurno = String(filtro || '').trim().toUpperCase();
    var lista = [];
    var vistos = {};

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];

      // Solo filas cuyo TIPO sea VACACIONES o PERMISO
      var tipo = iTipo >= 0 ? String(fila[iTipo] || '') : '';
      var tipoNorm = quitarAcentos_(tipo);
      if (tipoNorm.indexOf('VACACION') === -1 && tipoNorm.indexOf('PERMISO') === -1) continue;

      // Y cuya Fecha PRESENTACION (columna K) sea hoy
      var fPre = iFPre >= 0 ? partesFecha_(fila[iFPre], ZONA_) : null;
      if (!fPre) continue;
      if (fPre.anio !== hoy.anio || fPre.mes !== hoy.mes || fPre.dia !== hoy.dia) continue;

      var turno = iTur >= 0 ? String(fila[iTur] || '') : '';
      // Si se consulto un turno especifico, se filtra por el
      if (filtroTurno && filtroTurno !== 'TODOS' && String(turno).trim().toUpperCase() !== filtroTurno) continue;

      var cedula = iCed >= 0 ? String(fila[iCed] === null ? '' : fila[iCed]) : '';
      // Un funcionario no se repite aunque tenga varias novedades
      var clave = normalizaCedula_(cedula) || ('F' + i);
      if (vistos[clave]) continue;
      vistos[clave] = true;

      lista.push({
        cedula: cedula,
        funcionario: iNom >= 0 ? String(fila[iNom] || '') : '',
        grado: iGr >= 0 ? String(fila[iGr] || '') : '',
        turno: turno,
        tipo: tipo,
        novedad: iNov >= 0 ? String(fila[iNov] || '') : '',
        dias: iDias >= 0 ? String(fila[iDias] === null ? '' : fila[iDias]) : ''
      });
    }

    // Sin novedades de vacaciones ni permiso
    if (!lista.length) {
      return {
        estado: true,
        fecha: hoy.dia + '/' + hoy.mes + '/' + hoy.anio,
        vacio: true,
        mensaje: 'Hoy no se presenta de vacaciones o permiso. Ningún funcionario.',
        datos: []
      };
    }

    return {
      estado: true,
      fecha: hoy.dia + '/' + hoy.mes + '/' + hoy.anio,
      vacio: false,
      total: lista.length,
      datos: lista
    };
  } catch (err) {
    return { estado: false, mensaje: 'Error buscando presentaciones de hoy: ' + err.message };
  }
}

// Confirma quienes SI se presentaron hoy.
// A los marcados SI se les borra de NOVEDADES la fila de TIPO
// VACACIONES o PERMISO con Fecha PRESENTACION igual a hoy, con lo
// cual vuelven al listado S/N y se suman a la fuerza disponible.
// A los marcados NO no se les toca nada: la novedad sigue vigente.
function confirmarPresentacionesHoy(token, cedulas) {
  try {
    var entrada = Array.isArray(cedulas) ? cedulas : [];

    // Solo se aceptan los que trae "SI". El "NO" nunca se envia.
    var objetivo = {};
    for (var i = 0; i < entrada.length; i++) {
      var cc = normalizaCedula_(entrada[i]);
      if (cc) objetivo[cc] = true;
    }

    if (!Object.keys(objetivo).length) {
      return { estado: true, version: VERSION_APP, mensaje: 'No se marcó ningún funcionario como presentado.', filasEliminadas: 0, funcionarios: 0 };
    }

    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) {
      return { estado: false, version: VERSION_APP, mensaje: 'No existe la hoja NOVEDADES.' };
    }

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) {
      return { estado: true, version: VERSION_APP, mensaje: 'No hay novedades que actualizar.', filasEliminadas: 0, funcionarios: 0 };
    }

    var cab = (datos[0] || []).map(normalizarCabColumna_);
    var iCed = idxColumnaBase_(cab, [/^CEDULA$/, /^CC$/], 0);
    var iTipo = idxColumnaBase_(cab, [/^TIPO$/], -1);
    var iFPre = idxColumnaBase_(cab, [/^FECHA PRESENTACION$/], -1);

    var hoy = partesFecha_(new Date(), ZONA_);
    var borrar = [];
    var porCedula = {};

    for (var f = 1; f < datos.length; f++) {
      var fila = datos[f];

      // Solo TIPO VACACIONES o PERMISO
      var tipo = iTipo >= 0 ? String(fila[iTipo] || '') : '';
      var tipoNorm = quitarAcentos_(tipo);
      if (tipoNorm.indexOf('VACACION') === -1 && tipoNorm.indexOf('PERMISO') === -1) continue;

      // Y con Fecha PRESENTACION de hoy (columna K)
      var fPre = iFPre >= 0 ? partesFecha_(fila[iFPre], ZONA_) : null;
      if (!fPre) continue;
      if (fPre.anio !== hoy.anio || fPre.mes !== hoy.mes || fPre.dia !== hoy.dia) continue;

      var ccFila = iCed >= 0 ? normalizaCedula_(fila[iCed]) : '';
      if (!ccFila || !objetivo[ccFila]) continue;

      borrar.push(f + 1);
      porCedula[ccFila] = true;
    }

    // Se borra de abajo hacia arriba para no mover los indices
    for (var b = borrar.length - 1; b >= 0; b--) {
      hoja.deleteRow(borrar[b]);
    }

    var atendidos = Object.keys(porCedula).length;
    var sinCambio = Object.keys(objetivo).length - atendidos;

    var msg = 'Pasaron a S/N: ' + atendidos + ' funcionario(s). ' +
      'Se eliminaron ' + borrar.length + ' fila(s) de NOVEDADES.';
    if (sinCambio > 0) {
      msg += ' ' + sinCambio + ' no tenían novedad de vacaciones o permiso con fecha de hoy.';
    }

    return {
      estado: true,
      version: VERSION_APP,
      mensaje: msg,
      filasEliminadas: borrar.length,
      funcionarios: atendidos,
      sinCambio: sinCambio
    };
  } catch (err) {
    return { estado: false, version: VERSION_APP, mensaje: 'Error confirmando presentaciones: ' + err.message };
  }
}

/* =====================================================
   10. NOVEDADES RÁPIDAS DESDE LA TABLA DE TURNOS
   Permite asignar una novedad directamente sobre el listado
   del turno, sin pasar por el formulario "Registrar novedad".
   ===================================================== */

var NOVEDADES_RAPIDAS_ = [
  'OFICINA',
  'EN VACACIONES',
  'SERVICIO',
  'ESTADO DE GRAVIDEZ',
  'EXCUSA TOTAL',
  'PERMISO',
  'CITA MEDICA',
  'LICENCIA DE MATERNIDAD',
  'EXCUSADO DEL SERVICIO',
  'OTRA NOVEDAD',
  'RETARDADA',
  'EXCUSA PARCIAL',
  'CURSO DE ASCENSO'
];

// Valida el tipo recibido contra la lista blanca.
// Acepta "RETARDAD@" (como figura en la interfaz) y lo guarda como "RETARDADA".
function novedadRapidaValida_(tipo) {
  var t = quitarAcentos_(String(tipo || '').trim().toUpperCase());
  if (!t) return null;

  t = t.replace(/@/g, 'A').replace(/\s+/g, ' ').trim();

  for (var i = 0; i < NOVEDADES_RAPIDAS_.length; i++) {
    if (quitarAcentos_(NOVEDADES_RAPIDAS_[i]) === t) return NOVEDADES_RAPIDAS_[i];
  }
  return null;
}

// Registra una novedadrapida para un funcionario.
// No aplica la restriccion de novedades por usuario: es una operacion
// propia del modulo de turnos.
function registrarNovedadRapida(token, cedula, tipo) {
  try {
    var tipoOk = novedadRapidaValida_(tipo);
    if (!tipoOk) {
      return { estado: false, version: VERSION_APP, mensaje: 'Tipo de novedad no permitido.' };
    }

    var cedulaIn = normalizaCedula_(cedula);
    if (!cedulaIn) {
      return { estado: false, version: VERSION_APP, mensaje: 'Cédula inválida o vacía.' };
    }

    var ss = abrirLibro_();
    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) {
      return { estado: false, version: VERSION_APP, mensaje: 'No existe la hoja NOVEDADES.' };
    }

    var mapa = mapaColumnasNovedades_(hoja);
    if (!mapa) {
      return { estado: false, version: VERSION_APP, mensaje: 'No se pudo leer el encabezado de NOVEDADES.' };
    }

    var base = renglonBasePorCedula_(cedulaIn);
    if (!base) {
      return { estado: false, version: VERSION_APP, mensaje: 'Funcionario no encontrado en LISTADO_BASE.' };
    }

    // Nombre del funcionario para la columna F
    var nombre = '';
    var cabBase = base.cabeceras.map(normalizarCabColumna_);
    var idxNombre = idxColumnaBase_(cabBase,
      [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/, /^NOMBRES$/], -1);
    if (idxNombre !== -1 && idxNombre < base.valores.length) {
      nombre = String(base.valores[idxNombre] === null ? '' : base.valores[idxNombre]).trim();
    }

    var hoy = Utilities.formatDate(new Date(), ZONA_, 'yyyy-MM-dd');

    var fila = construirFilaNovedades_(mapa, {
      cedula: cedulaIn,
      nombreFuncionario: nombre,
      tipo: tipoOk,
      novedad: tipoOk,
      dias: '',
      fechaInicial: hoy,
      fechaPresentacion: hoy,
      observacion: '',
      texto: tipoOk
    });

    // GR, NIV, DEPENDENCIA, TURNO y Placa_Chip se copian de LISTADO_BASE
    completarDesdeBase_(fila, mapa, base);

    hoja.appendRow(fila);

    return {
      estado: true,
      version: VERSION_APP,
      mensaje: 'Novedad "' + tipoOk + '" registrada.',
      tipo: tipoOk,
      cedula: cedulaIn,
      funcionario: nombre
    };
  } catch (err) {
    return { estado: false, version: VERSION_APP, mensaje: 'Error registrando novedad: ' + err.message };
  }
}

// Mide cuanto tarda cada parte del backend para detectar cuellos de botella.
// No necesita token ni usuario: solo informa tiempos.
function diagnosticarCarga() {
  var t0 = Date.now();
  var L = [];
  L.push('diagnosticarCarga inicio');

  try {
    var tA = Date.now();
    var ss = abrirLibro_();
    L.push('abrirLibro_          : ' + (Date.now() - tA) + ' ms');

    var tB = Date.now();
    var nom = ss.getName();
    L.push('getName               : ' + (Date.now() - tB) + ' ms  -> ' + nom);

    var tC = Date.now();
    var hojas = ss.getSheets().map(function(h) { return h.getName(); });
    L.push('getSheets            : ' + (Date.now() - tC) + ' ms  -> ' + hojas.join(', '));

    var tD = Date.now();
    var us = ss.getSheetByName('USUARIOS');
    var filasUs = us ? us.getLastRow() : -1;
    L.push('getSheetByName USUAR  : ' + (Date.now() - tD) + ' ms  -> ' + filasUs + ' filas');

    var tE = Date.now();
    var nov = ss.getSheetByName('NOVEDADES');
    var filasNov = nov ? nov.getLastRow() : -1;
    L.push('getSheetByName NOVED  : ' + (Date.now() - tE) + ' ms  -> ' + filasNov + ' filas');

    var tF = Date.now();
    var datos = nov ? nov.getDataRange().getValues() : [];
    L.push('getValues NOVEDADES   : ' + (Date.now() - tF) + ' ms  -> ' + datos.length + ' filas');

    var tG = Date.now();
    var lista = listarNovedades('DIAG', 'VAC_JESEP', { pagina: 1, porPagina: 25 });
    L.push('listarNovedades(25)   : ' + (Date.now() - tG) + ' ms  estado=' + (lista && lista.estado) +
      ' total=' + (lista && lista.total !== undefined ? lista.total : '?') +
      ' paginas=' + (lista && lista.paginas !== undefined ? lista.paginas : '?') +
      (lista && lista.estado !== true ? '  msg=' + lista.mensaje : ''));

    var tH = Date.now();
    var filtrada = listarNovedades('DIAG', 'VAC_JESEP', { pagina: 1, porPagina: 25, cedula: '111' });
    L.push('listarNovedades(busc) : ' + (Date.now() - tH) + ' ms  total=' + (filtrada && filtrada.total));

    L.push('');
    L.push('TOTAL: ' + (Date.now() - t0) + ' ms');

    return { estado: true, version: VERSION_APP, reporte: L.join('\n') };
  } catch (err) {
    L.push('ERROR: ' + err.message);
    L.push('TOTAL hasta el fallo: ' + (Date.now() - t0) + ' ms');
    return { estado: false, version: VERSION_APP, reporte: L.join('\n') };
  }
}

/* =====================================================
   DIAGNÓSTICO DEL CRUCE LISTADO_BASE <-> NOVEDADES
   Muestra los datos reales para verificar por qué no se cruzan.
   ===================================================== */
function diagnosticarCruceDataSafe() {
  var L = [];
  function add(x) { L.push(String(x)); }

  try {
    var ss = abrirLibro_();
    add('VERSION ' + VERSION_APP);
    add('LIBRO ' + ss.getName());
    add('');

    // ---------- LISTADO_BASE ----------
    var base = ss.getSheetByName('LISTADO_BASE');
    if (!base) {
      add('ERROR: no existe LISTADO_BASE');
      return { estado: false, version: VERSION_APP, reporte: L.join('\n') };
    }
    var dB = base.getDataRange().getValues();
    var cabB = dB[0] || [];
    var cabNB = cabB.map(normalizarCabColumna_);

    add('=== LISTADO_BASE ===');
    add('Filas de datos: ' + Math.max(0, dB.length - 1));
    add('Columnas: ' + cabB.length);
    add('');
    add('Encabezados (letra = nombre):');
    for (var c = 0; c < cabB.length; c++) {
      add('  ' + letraColumna_(c + 1) + ' = ' + (cabB[c] === '' ? '(vacia)' : cabB[c]));
    }

    var iCed = idxColumnaBase_(cabNB, [/^CEDULA$/, /^CC$/, /^NUMERO DE CEDULA$/, /^N C$/], 0);
    var iTur = idxColumnaBase_(cabNB, [/^TURNO$/], 7);
    var iGrd = idxColumnaBase_(cabNB, [/^GR$/, /^GRADO$/], -1);
    var iNom = idxColumnaBase_(cabNB, [/^FUNCIONARIO$/, /^APELLIDOS Y NOMBRES$/, /^NOMBRE$/, /^NOMBRES$/], -1);

    add('');
    add('Indices calculados:');
    add('  CEDULA -> col ' + letraColumna_(iCed + 1));
    add('  TURNO  -> col ' + letraColumna_(iTur + 1));
    add('  GR     -> col ' + letraColumna_(iGrd + 1));
    add('  NOMBRE -> col ' + letraColumna_(iNom + 1));

    var cedulasBase = {};
    add('');
    add('Primeras 5 cedulas de LISTADO_BASE (como estan / normalizadas):');
    for (var i = 1; i <= 5 && i < dB.length; i++) {
      var crudo = dB[i][iCed];
      add('  fila ' + (i + 1) + ': ["' + crudo + '"]  ->  "' + normalizaCedula_(crudo) + '"');
      cedulasBase[normalizaCedula_(crudo)] = true;
    }

    // ---------- NOVEDADES ----------
    add('');
    add('=== NOVEDADES ===');
    var nov = ss.getSheetByName('NOVEDADES');
    if (!nov) {
      add('ERROR: no existe la hoja NOVEDADES');
      return { estado: false, version: VERSION_APP, reporte: L.join('\n') };
    }
    var dN = nov.getDataRange().getValues();
    add('Filas de datos: ' + Math.max(0, dN.length - 1));
    add('Columnas: ' + (dN[0] ? dN[0].length : 0));
    add('');
    add('Encabezados:');
    var cabN = dN[0] || [];
    for (var c2 = 0; c2 < cabN.length; c2++) {
      add('  ' + letraColumna_(c2 + 1) + ' = ' + (cabN[c2] === '' ? '(vacia)' : cabN[c2]));
    }

    add('');
    add('Primeras 5 filas de NOVEDADES (todas las columnas):');
    var coincidencias = 0;
    for (var n = 1; n <= 5 && n < dN.length; n++) {
      var fila = dN[n];
      var partes = [];
      for (var k = 0; k < fila.length; k++) {
        partes.push(letraColumna_(k + 1) + '=["' + fila[k] + '"]');
      }
      add('  fila ' + (n + 1) + ': ' + partes.join(' '));
      var cedN = normalizaCedula_(fila[1]);
      var hallada = cedulasBase[cedN] === true;
      if (hallada) coincidencias++;
      add('     -> cedula col B = "' + cedN + '"  |  ' + (hallada ? 'SI existe en LISTADO_BASE' : 'NO existe en LISTADO_BASE (muestra)'));
    }

    add('');
    add('RESULTADO: ' + coincidencias + ' de ' + Math.min(5, Math.max(0, dN.length - 1)) +
        ' filas de NOVEDADES coincidieron con la muestra de LISTADO_BASE.');
    add('Si es 0, el problema son las cedulas. Si coincide, el problema esta en las columnas E/F.');

  } catch (err) {
    add('');
    add('ERROR GENERAL: ' + err);
  }

  return { estado: true, version: VERSION_APP, reporte: L.join('\n') };
}

// Version segura: acumula el reporte y lo devuelve como dato.
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
