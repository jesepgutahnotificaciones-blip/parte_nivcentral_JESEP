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
   1. GESTIÓN DE USUARIOS Y AUTENTICACIÓN
   ===================================================== */

function validarUsuario(usuarioIngresado, claveIngresada) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaUsuarios = ss.getSheetByName('USUARIOS');
    if (!hojaUsuarios) return { estado: false, mensaje: 'No se encontró la hoja "USUARIOS".' };

    var datos = hojaUsuarios.getDataRange().getValues();
    if (datos.length < 2) return { estado: false, mensaje: 'La hoja USUARIOS no contiene registros.' };
    
    var cabeceras = datos[0].map(function(c) { return String(c).trim().toUpperCase(); });
    
    // Buscar los índices de manera flexible
    var idxUser = cabeceras.indexOf('USUARIO');
    var idxClave = cabeceras.indexOf('CLAVE');
    var idxRol = cabeceras.indexOf('ROL');
    var idxDep = cabeceras.indexOf('DEPENDENCIA');

    // Si no encuentra por nombres de cabecera predeterminados, asume las columnas estándar (A=0, B=1, C=2, D=3)
    if (idxUser === -1) idxUser = 0;
    if (idxClave === -1) idxClave = 1;
    if (idxRol === -1) idxRol = 2;
    if (idxDep === -1) idxDep = 3;

    var userClean = String(usuarioIngresado).trim().toLowerCase();
    var passClean = String(claveIngresada).trim();

    for (var i = 1; i < datos.length; i++) {
      var row = datos[i];
      var rowUser = String(row[idxUser] || '').trim().toLowerCase();
      var rowClave = String(row[idxClave] || '').trim();

      if (rowUser === userClean && rowClave === passClean) {
        return {
          estado: true,
          token: 'MODO_SIN_LOGIN',
          usuario: { 
            usuario: row[idxUser], 
            rol: row[idxRol] || 'ADMINISTRADOR', 
            dependencia: row[idxDep] || 'GENERAL' 
          }
        };
      }
    }
    return { estado: false, mensaje: 'Usuario o contraseña incorrectos.' };
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
        for (var j = 0; j < cabeceras.length; j++) { obj[cabeceras[j]] = datos[i][j]; }
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
        for (var j = 0; j < cabeceras.length; j++) { obj[cabeceras[j]] = datos[i][j]; }
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
   4. TURNOS Y REPORTES (CON CRUCE DE COLUMNAS E Y F)
   ===================================================== */

function consultarPorTurno(token, filtro) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaBase = ss.getSheetByName('LISTADO_BASE');
    if (!hojaBase) return { estado: false, mensaje: 'Hoja LISTADO_BASE não encontrada.' };

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

        // Cruce con la hoja NOVEDADES (Concatenando columna E [índice 4] y F [índice 5])
        if (datosNovedades.length > 1 && cedulaFuncionario !== '') {
          var cabecerasNov = datosNovedades[0];
          var idxCedulaNov = -1;
          for (var c = 0; c < cabecerasNov.length; c++) {
            var nombreCab = String(cabecerasNov[c]).toUpperCase();
            if (nombreCab === 'CEDULA' || nombreCab === 'CC') {
              idxCedulaNov = c;
              break;
            }
          }

          if (idxCedulaNov !== -1) {
            for (var n = 1; n < datosNovedades.length; n++) {
              var filaNov = datosNovedades[n];
              var cedulaNov = String(filaNov[idxCedulaNov]).trim();
              if (cedulaNov === cedulaFuncionario) {
                var valorE = filaNov[4] !== undefined && filaNov[4] !== null ? String(filaNov[4]).trim() : '';
                var valorF = filaNov[5] !== undefined && filaNov[5] !== null ? String(filaNov[5]).trim() : '';
                
                var concatenado = '';
                if (valorE && valorF) {
                  concatenado = valorE + ' - ' + valorF;
                } else {
                  concatenado = valorE + valorF;
                }

                if (concatenado) {
                  novedadesCruzadas.push({ NOVEDAD: concatenado });
                }
              }
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
    var html = '<h3 style="font-family:Arial;">Reporte de Turno: ' + filtro + '</h3>';
    html += '<table border="1" cellpadding="5" style="border-collapse:collapse;width:100%;font-family:Arial;font-size:12px;">';
    html += '<tr style="background:#01592F;color:white;"><th>Cédula</th><th>Grado</th><th>Funcionario</th><th>Dependencia</th><th>Turno</th><th>Novedades</th></tr>';
    
    funcs.forEach(function(f) {
      var novedadesTexto = (f.historialNovedades || []).map(function(n) { return n.NOVEDAD; }).join(', ');
      html += '<tr>';
      html += '<td>' + (f.cedula || f.CEDULA || '') + '</td>';
      html += '<td>' + (f.grado || f.GR || '') + '</td>';
      html += '<td>' + (f.funcionario || f.FUNCIONARIO || '') + '</td>';
      html += '<td>' + (f.dependencia || f.DEPENDENCIA || '') + '</td>';
      html += '<td>' + (f.turno || f.TURNO || '') + '</td>';
      html += '<td>' + novedadesTexto + '</td>';
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