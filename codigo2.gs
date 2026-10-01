/* =====================================================
   ENRUTADOR PRINCIPAL (JSONP PARA GITHUB PAGES / WEB APP)
   ===================================================== */

function doGet(e) {
  var accion = e.parameter.accion;
  var callback = e.parameter.callback;
  var argsRaw = e.parameter.args;
  
  var args = [];
  try {
    if (argsRaw) {
      args = JSON.parse(argsRaw);
    }
  } catch (err) {
    var errorRespuesta = { __jsonp_error: true, mensaje: 'Error al procesar los argumentos.' };
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(errorRespuesta) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  var resultado;
  try {
    switch (accion) {
      case 'validarUsuario':
        resultado = validarUsuario(args[0], args[1]);
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
      default:
        resultado = { estado: false, mensaje: 'Acción no válida: ' + accion };
    }
  } catch (error) {
    resultado = { estado: false, mensaje: error.toString() };
  }

  var jsonString = JSON.stringify(resultado);
  var scriptOutput = callback + '(' + jsonString + ');';
  
  return ContentService.createTextOutput(scriptOutput)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}


/* =====================================================
   FUNCIÓN AUXILIAR DE VALIDACIÓN DE SESIÓN (BYPASS)
   ===================================================== */
function verificarToken_(token) {
  if (token === 'MODO_SIN_LOGIN') {
    return { estado: true, rol: 'ADMINISTRADOR' };
  }
  // Si en el futuro requieres tokens reales, puedes ponerlos aquí.
  // Por defecto, aceptamos cualquier token para evitar bloqueos.
  return { estado: true, rol: 'ADMINISTRADOR' };
}


/* =====================================================
   1. GESTIÓN DE USUARIOS Y AUTENTICACIÓN
   ===================================================== */

function validarUsuario(usuarioIngresado, claveIngresada) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaUsuarios = ss.getSheetByName('USUARIOS');
    
    if (!hojaUsuarios) {
      return { estado: false, mensaje: 'No se encontró la hoja "USUARIOS" en el libro de Excel/Drive.' };
    }

    var datos = hojaUsuarios.getDataRange().getValues();
    if (datos.length < 2) return { estado: false, mensaje: 'La hoja USUARIOS no contiene registros.' };
    
    var cabeceras = datos[0];
    var idxUser = cabeceras.indexOf('USUARIO');
    var idxClave = cabeceras.indexOf('CLAVE');
    var idxRol = cabeceras.indexOf('ROL');
    var idxDep = cabeceras.indexOf('DEPENDENCIA');
    var idxEstado = cabeceras.indexOf('ESTADO');

    for (var i = 1; i < datos.length; i++) {
      var row = datos[i];
      var user = String(row[idxUser]).trim();
      var clave = String(row[idxClave]).trim();
      var estado = idxEstado !== -1 ? String(row[idxEstado]).trim().toUpperCase() : 'ACTIVO';

      if (user.toLowerCase() === usuarioIngresado.trim().toLowerCase() && clave === claveIngresada) {
        if (estado === 'INACTIVO') {
          return { estado: false, mensaje: 'El usuario se encuentra inactivo.' };
        }

        var tokenUnico = 'token_' + Utilities.getUuid();
        
        var idxAcceso = cabeceras.indexOf('ULTIMO_ACCESO');
        if (idxAcceso !== -1) {
          hojaUsuarios.getRange(i + 1, idxAcceso + 1).setValue(new Date());
        }

        return {
          estado: true,
          token: tokenUnico,
          usuario: {
            usuario: user,
            rol: row[idxRol] || 'OPERADOR',
            dependencia: row[idxDep] || 'GENERAL'
          }
        };
      }
    }

    return { estado: false, mensaje: 'Usuario o contraseña incorrectos.' };
  } catch (error) {
    return { estado: false, mensaje: 'Error al validar usuario: ' + error.message };
  }
}

function cerrarSesionCliente(token) {
  return { estado: true };
}

function listarUsuarios(token) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) return { estado: false, mensaje: 'Hoja USUARIOS no encontrada.' };

    var datos = hoja.getDataRange().getValues();
    var cabeceras = datos[0];
    var usuarios = [];

    for (var i = 1; i < datos.length; i++) {
      var obj = {};
      for (var j = 0; j < cabeceras.length; j++) {
        obj[cabeceras[j]] = datos[i][j];
      }
      usuarios.push(obj);
    }

    return { estado: true, datos: usuarios };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function crearUsuario(token, usuario, clave, rol, dependencia) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) return { estado: false, mensaje: 'Hoja USUARIOS no encontrada.' };

    hoja.appendRow([usuario, clave, rol, dependencia, 'ACTIVO', new Date()]);
    return { estado: true, mensaje: 'Usuario creado exitosamente.' };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function cambiarEstadoUsuario(token, usuario, nuevoEstado) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('USUARIOS');
    var datos = hoja.getDataRange().getValues();
    var idxUser = datos[0].indexOf('USUARIO');
    var idxEstado = datos[0].indexOf('ESTADO');

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][idxUser]).trim() === String(usuario).trim()) {
        hoja.getRange(i + 1, idxEstado + 1).setValue(nuevoEstado);
        return { estado: true, mensaje: 'Estado actualizado correctamente.' };
      }
    }
    return { estado: false, mensaje: 'Usuario no encontrado.' };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}


/* =====================================================
   2. FUNCIONARIOS Y CONSULTAS
   ===================================================== */

function buscarFuncionario(token, valor) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja LISTADO_BASE.' };

    var datos = hoja.getDataRange().getValues();
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
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    var datos = hoja.getDataRange().getValues();
    var cabeceras = datos[0];
    var idxCedula = -1;

    for (var j = 0; j < cabeceras.length; j++) {
      var c = String(cabeceras[j]).toUpperCase();
      if (c === 'CEDULA' || c === 'CC') {
        idxCedula = j;
        break;
      }
    }

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][idxCedula]).trim() === String(cedula).trim()) {
        var obj = {};
        for (var j = 0; j < cabeceras.length; j++) {
          obj[cabeceras[j]] = datos[i][j];
        }
        return { estado: true, funcionario: obj };
      }
    }

    return { estado: false, mensaje: 'Funcionario no encontrado.' };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function obtenerHistorialFuncionario(token, cedula) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) return { estado: true, historial: [] }; 

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: true, historial: [] };

    var cabeceras = datos[0];
    var idxCedula = -1;
    for (var j = 0; j < cabeceras.length; j++) {
      var c = String(cabeceras[j]).toUpperCase();
      if (c === 'CEDULA' || c === 'CC') {
        idxCedula = j;
        break;
      }
    }

    var historial = [];
    for (var i = 1; i < datos.length; i++) {
      if (idxCedula !== -1 && String(datos[i][idxCedula]).trim() === String(cedula).trim()) {
        var obj = {};
        for (var j = 0; j < cabeceras.length; j++) {
          obj[cabeceras[j]] = datos[i][j];
        }
        historial.push(obj);
      }
    }

    return { estado: true, historial: historial };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}


/* =====================================================
   3. REGISTRO DE NOVEDADES
   ===================================================== */

function registrarNovedad(token, datosNovedad) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('NOVEDADES');
    
    if (!hoja) {
      hoja = ss.insertSheet('NOVEDADES');
      hoja.appendRow(['CEDULA', 'NOVEDAD', 'TIPO', 'DESCRIPCION', 'Dias', 'Fecha INICIAL', 'Fecha PRESENTACION', 'Observacion', 'RV', 'TEXTO', 'FECHA_REGISTRO']);
    }

    hoja.appendRow([
      datosNovedad.cedula || datosNovedad.cc,
      datosNovedad.novedad,
      datosNovedad.tipo,
      datosNovedad.descripcion,
      datosNovedad.dias,
      datosNovedad.fechaInicial,
      datosNovedad.fechaPresentacion,
      datosNovedad.observacion,
      datosNovedad.rv,
      datosNovedad.descripcion || '',
      new Date()
    ]);

    return { estado: true, mensaje: 'Novedad registrada correctamente.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error al registrar novedad: ' + err.message };
  }
}


/* =====================================================
   4. TURNOS Y REPORTES
   ===================================================== */

function consultarPorTurno(token, filtro) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'Hoja LISTADO_BASE no encontrada.' };

    var datos = hoja.getDataRange().getValues();
    var cabeceras = datos[0];
    
    var idxTurno = -1;
    for (var j = 0; j < cabeceras.length; j++) {
      if (String(cabeceras[j]).toUpperCase() === 'TURNO') {
        idxTurno = j;
        break;
      }
    }

    var funcionarios = [];
    for (var i = 1; i < datos.length; i++) {
      var obj = {};
      for (var j = 0; j < cabeceras.length; j++) {
        obj[cabeceras[j]] = datos[i][j];
      }
      
      var turnoFila = String(obj.turno || obj.TURNO || '').trim().toUpperCase();
      if (filtro === 'SEPRI' || turnoFila === String(filtro).trim().toUpperCase()) {
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
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var resultadoTurno = consultarPorTurno(token, filtro);
    if (!resultadoTurno.estado) return resultadoTurno;

    var funcs = resultadoTurno.datos.funcionarios;
    var html = '<h3 style="font-family:Arial;">Reporte de Turno: ' + filtro + '</h3>';
    html += '<table border="1" cellpadding="5" style="border-collapse:collapse;width:100%;font-family:Arial;font-size:12px;">';
    html += '<tr style="background:#01592F;color:white;"><th>Cédula</th><th>Grado</th><th>Funcionario</th><th>Dependencia</th><th>Turno</th></tr>';
    
    funcs.forEach(function(f) {
      html += '<tr>';
      html += '<td>' + (f.cedula || f.CEDULA || '') + '</td>';
      html += '<td>' + (f.grado || f.GR || '') + '</td>';
      html += '<td>' + (f.funcionario || f.FUNCIONARIO || '') + '</td>';
      html += '<td>' + (f.dependencia || f.DEPENDENCIA || '') + '</td>';
      html += '<td>' + (f.turno || f.TURNO || '') + '</td>';
      html += '</tr>';
    });
    html += '</table>';

    return { estado: true, html: html };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function generarReporteTurno(token, filtro) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    var consecutivo = 'REP-' + Date.now();
    return {
      estado: true,
      datos: {
        filtro: filtro,
        consecutivo: consecutivo,
        urlPDF: '#',
        urlExcel: '#'
      }
    };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function listarReportes(token, limite) {
  var verificacion = verificarToken_(token);
  if (!verificacion.estado) return { estado: false, mensaje: 'Sesión vencida o no válida.' };

  try {
    return { estado: true, datos: [] };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function doGet(e) {
  var accion = e.parameter.accion;
  var callback = e.parameter.callback;
  var argsRaw = e.parameter.args;
  
  var args = [];
  try {
    if (argsRaw) {
      args = JSON.parse(argsRaw);
    }
  } catch (err) {
    var errorRespuesta = { __jsonp_error: true, mensaje: 'Error al procesar los argumentos.' };
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(errorRespuesta) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  var resultado;
  try {
    switch (accion) {
      case 'validarUsuario':
        resultado = validarUsuario(args[0], args[1]);
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
      default:
        resultado = { estado: false, mensaje: 'Acción no válida: ' + accion };
    }
  } catch (error) {
    resultado = { estado: false, mensaje: error.toString() };
  }

  var jsonString = JSON.stringify(resultado);
  var scriptOutput = callback + '(' + jsonString + ');';
  
  return ContentService.createTextOutput(scriptOutput)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}


/* =====================================================
   1. GESTIÓN DE USUARIOS Y AUTENTICACIÓN
   ===================================================== */

function validarUsuario(usuarioIngresado, claveIngresada) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaUsuarios = ss.getSheetByName('USUARIOS');
    
    if (!hojaUsuarios) {
      return { estado: false, mensaje: 'No se encontró la hoja "USUARIOS" en el libro de Excel/Drive.' };
    }

    var datos = hojaUsuarios.getDataRange().getValues();
    if (datos.length < 2) return { estado: false, mensaje: 'La hoja USUARIOS no contiene registros.' };
    
    var cabeceras = datos[0];
    var idxUser = cabeceras.indexOf('USUARIO');
    var idxClave = cabeceras.indexOf('CLAVE');
    var idxRol = cabeceras.indexOf('ROL');
    var idxDep = cabeceras.indexOf('DEPENDENCIA');
    var idxEstado = cabeceras.indexOf('ESTADO');

    for (var i = 1; i < datos.length; i++) {
      var row = datos[i];
      var user = String(row[idxUser]).trim();
      var clave = String(row[idxClave]).trim();
      var estado = idxEstado !== -1 ? String(row[idxEstado]).trim().toUpperCase() : 'ACTIVO';

      if (user.toLowerCase() === usuarioIngresado.trim().toLowerCase() && clave === claveIngresada) {
        if (estado === 'INACTIVO') {
          return { estado: false, mensaje: 'El usuario se encuentra inactivo.' };
        }

        var tokenUnico = 'token_' + Utilities.getUuid();
        
        var idxAcceso = cabeceras.indexOf('ULTIMO_ACCESO');
        if (idxAcceso !== -1) {
          hojaUsuarios.getRange(i + 1, idxAcceso + 1).setValue(new Date());
        }

        return {
          estado: true,
          token: tokenUnico,
          usuario: {
            usuario: user,
            rol: row[idxRol] || 'OPERADOR',
            dependencia: row[idxDep] || 'GENERAL'
          }
        };
      }
    }

    return { estado: false, mensaje: 'Usuario o contraseña incorrectos.' };
  } catch (error) {
    return { estado: false, mensaje: 'Error al validar usuario: ' + error.message };
  }
}

function cerrarSesionCliente(token) {
  return { estado: true };
}

function listarUsuarios(token) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) return { estado: false, mensaje: 'Hoja USUARIOS no encontrada.' };

    var datos = hoja.getDataRange().getValues();
    var cabeceras = datos[0];
    var usuarios = [];

    for (var i = 1; i < datos.length; i++) {
      var obj = {};
      for (var j = 0; j < cabeceras.length; j++) {
        obj[cabeceras[j]] = datos[i][j];
      }
      usuarios.push(obj);
    }

    return { estado: true, datos: usuarios };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function crearUsuario(token, usuario, clave, rol, dependencia) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('USUARIOS');
    if (!hoja) return { estado: false, mensaje: 'Hoja USUARIOS no encontrada.' };

    hoja.appendRow([usuario, clave, rol, dependencia, 'ACTIVO', new Date()]);
    return { estado: true, mensaje: 'Usuario creado exitosamente.' };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function cambiarEstadoUsuario(token, usuario, nuevoEstado) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('USUARIOS');
    var datos = hoja.getDataRange().getValues();
    var idxUser = datos[0].indexOf('USUARIO');
    var idxEstado = datos[0].indexOf('ESTADO');

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][idxUser]).trim() === String(usuario).trim()) {
        hoja.getRange(i + 1, idxEstado + 1).setValue(nuevoEstado);
        return { estado: true, mensaje: 'Estado actualizado correctamente.' };
      }
    }
    return { estado: false, mensaje: 'Usuario no encontrado.' };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}


/* =====================================================
   2. FUNCIONARIOS Y CONSULTAS
   ===================================================== */

function buscarFuncionario(token, valor) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'No se encontró la hoja LISTADO_BASE.' };

    var datos = hoja.getDataRange().getValues();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    var datos = hoja.getDataRange().getValues();
    var cabeceras = datos[0];
    var idxCedula = -1;

    for (var j = 0; j < cabeceras.length; j++) {
      var c = String(cabeceras[j]).toUpperCase();
      if (c === 'CEDULA' || c === 'CC') {
        idxCedula = j;
        break;
      }
    }

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][idxCedula]).trim() === String(cedula).trim()) {
        var obj = {};
        for (var j = 0; j < cabeceras.length; j++) {
          obj[cabeceras[j]] = datos[i][j];
        }
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('NOVEDADES');
    if (!hoja) return { estado: true, historial: [] }; // Si no existe la hoja, retorna vacío limpiamente

    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: true, historial: [] };

    var cabeceras = datos[0];
    var idxCedula = -1;
    for (var j = 0; j < cabeceras.length; j++) {
      var c = String(cabeceras[j]).toUpperCase();
      if (c === 'CEDULA' || c === 'CC') {
        idxCedula = j;
        break;
      }
    }

    var historial = [];
    for (var i = 1; i < datos.length; i++) {
      if (idxCedula !== -1 && String(datos[i][idxCedula]).trim() === String(cedula).trim()) {
        var obj = {};
        for (var j = 0; j < cabeceras.length; j++) {
          obj[cabeceras[j]] = datos[i][j];
        }
        historial.push(obj);
      }
    }

    return { estado: true, historial: historial };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}


/* =====================================================
   3. REGISTRO DE NOVEDADES
   ===================================================== */

function registrarNovedad(token, datosNovedad) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('NOVEDADES');
    
    if (!hoja) {
      hoja = ss.insertSheet('NOVEDADES');
      hoja.appendRow(['CEDULA', 'NOVEDAD', 'TIPO', 'DESCRIPCION', 'Dias', 'Fecha INICIAL', 'Fecha PRESENTACION', 'Observacion', 'RV', 'TEXTO', 'FECHA_REGISTRO']);
    }

    hoja.appendRow([
      datosNovedad.cedula || datosNovedad.cc,
      datosNovedad.novedad,
      datosNovedad.tipo,
      datosNovedad.descripcion,
      datosNovedad.dias,
      datosNovedad.fechaInicial,
      datosNovedad.fechaPresentacion,
      datosNovedad.observacion,
      datosNovedad.rv,
      datosNovedad.descripcion || '',
      new Date()
    ]);

    return { estado: true, mensaje: 'Novedad registrada correctamente.' };
  } catch (err) {
    return { estado: false, mensaje: 'Error al registrar novedad: ' + err.message };
  }
}


/* =====================================================
   4. TURNOS Y REPORTES
   ===================================================== */

function consultarPorTurno(token, filtro) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('LISTADO_BASE');
    if (!hoja) return { estado: false, mensaje: 'Hoja LISTADO_BASE no encontrada.' };

    var datos = hoja.getDataRange().getValues();
    var cabeceras = datos[0];
    
    var idxTurno = -1;
    for (var j = 0; j < cabeceras.length; j++) {
      if (String(cabeceras[j]).toUpperCase() === 'TURNO') {
        idxTurno = j;
        break;
      }
    }

    var funcionarios = [];
    for (var i = 1; i < datos.length; i++) {
      var obj = {};
      for (var j = 0; j < cabeceras.length; j++) {
        obj[cabeceras[j]] = datos[i][j];
      }
      
      var turnoFila = String(obj.turno || obj.TURNO || '').trim().toUpperCase();
      if (filtro === 'SEPRI' || turnoFila === String(filtro).trim().toUpperCase()) {
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
    var html = '<h3 style="font-family:Arial;">Reporte de Turno: ' + filtro + '</h3>';
    html += '<table border="1" cellpadding="5" style="border-collapse:collapse;width:100%;font-family:Arial;font-size:12px;">';
    html += '<tr style="background:#01592F;color:white;"><th>Cédula</th><th>Grado</th><th>Funcionario</th><th>Dependencia</th><th>Turno</th></tr>';
    
    funcs.forEach(function(f) {
      html += '<tr>';
      html += '<td>' + (f.cedula || f.CEDULA || '') + '</td>';
      html += '<td>' + (f.grado || f.GR || '') + '</td>';
      html += '<td>' + (f.funcionario || f.FUNCIONARIO || '') + '</td>';
      html += '<td>' + (f.dependencia || f.DEPENDENCIA || '') + '</td>';
      html += '<td>' + (f.turno || f.TURNO || '') + '</td>';
      html += '</tr>';
    });
    html += '</table>';

    return { estado: true, html: html };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function generarReporteTurno(token, filtro) {
  try {
    var consecutivo = 'REP-' + Date.now();
    return {
      estado: true,
      datos: {
        filtro: filtro,
        consecutivo: consecutivo,
        urlPDF: '#',
        urlExcel: '#'
      }
    };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function listarReportes(token, limite) {
  try {
    return { estado: true, datos: [] };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}