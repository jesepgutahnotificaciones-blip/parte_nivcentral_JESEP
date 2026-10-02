/* =====================================================
   CONFIGURACIÓN
   ===================================================== */

// ID del libro de Google Sheets (el que contiene USUARIOS, LISTADO_BASE, NOVEDADES...).
// Funciona tanto si el script está ligado al libro como si es un proyecto independiente.
var ID_LIBRO = '1kLYnWBgKgMfahllxWwe5ijls52YM3t_klg4XRnNaC0Q';

// Mientras sea true, el mensaje de "Usuario o contraseña incorrectos" incluye datos de diagnóstico.
// Cuando el login funcione, cámbielo a false y vuelva a implementar.
var DEBUG_LOGIN = true;

function abrirLibro_() {
  if (ID_LIBRO) return SpreadsheetApp.openById(ID_LIBRO);
  return SpreadsheetApp.getActiveSpreadsheet();
}


/* =====================================================
   ENRUTADOR PRINCIPAL (JSONP PARA GITHUB PAGES / WEB APP)
   ===================================================== */

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var accion = p.accion;
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
      default:
        resultado = { estado: false, mensaje: 'Acción no válida: ' + accion };
    }
  } catch (error) {
    resultado = { estado: false, mensaje: error.toString() };
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

      // Determinar turno permitido basado en el nombre de usuario si no está en la hoja
      if (!turnoPermitido) {
        var userUpper = userIn.toUpperCase();
        if (userUpper.indexOf('DISPONIBLE_A') !== -1) turnoPermitido = 'A';
        else if (userUpper.indexOf('DISPONIBLE_B') !== -1) turnoPermitido = 'B';
        else if (userUpper.indexOf('DISPONIBLE_C') !== -1) turnoPermitido = 'C';
      }

      // Asignar módulos según el rol o permisos institucionales
      var modulosAsignados = ['index_1', 'index_2'];
      if (rolUsuario === 'OPERADOR') {
        modulosAsignados = ['index_1'];
      }

      return {
        estado: true,
        token: 'SESION_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
        usuario: {
          usuario: row[idxUser],
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

function registrarNovedad(token, datosNovedad) {
  try {
    var ss = abrirLibro_();
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
    var resultadoTurno = consultarPorTurno(token, filtro);
    if (!resultadoTurno.estado) return resultadoTurno;

    var funcs = resultadoTurno.datos.funcionarios;
    var consecutivo = 'REP-' + Date.now();

    // Generar PDF usando DocumentApp
    var pdfBlob = generarPDFBlob_(funcs, filtro, consecutivo);
    var pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var urlPDF = pdfFile.getUrl();

    // Generar Excel (CSV) usando SpreadsheetApp
    var excelBlob = generarExcelBlob_(funcs, filtro, consecutivo);
    var excelFile = DriveApp.createFile(excelBlob);
    excelFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var urlExcel = excelFile.getUrl();

    // Guardar registro en hoja REPORTES
    try {
      var ss = abrirLibro_();
      var hojaReportes = ss.getSheetByName('REPORTES');
      if (!hojaReportes) {
        hojaReportes = ss.insertSheet('REPORTES');
        hojaReportes.appendRow(['CONSECUTIVO', 'FECHA', 'USUARIO', 'TURNO', 'URL_PDF', 'URL_EXCEL']);
      }
      hojaReportes.appendRow([consecutivo, new Date(), 'USUARIO', filtro, urlPDF, urlExcel]);
    } catch (e) {
      // No crítico si no se puede guardar el registro
    }

    return {
      estado: true,
      datos: {
        filtro: filtro,
        consecutivo: consecutivo,
        urlPDF: urlPDF,
        urlExcel: urlExcel
      }
    };
  } catch (err) {
    return { estado: false, mensaje: err.message };
  }
}

function generarPDFBlob_(funcs, filtro, consecutivo) {
  var html = '<html><head><style>';
  html += 'body{font-family:Arial,sans-serif;margin:20px;}';
  html += 'h2{color:#01592F;}';
  html += 'table{width:100%;border-collapse:collapse;font-size:11px;}';
  html += 'th{background:#01592F;color:white;padding:8px;text-align:left;}';
  html += 'td{padding:6px;border:1px solid #ddd;}';
  html += 'tr:nth-child(even){background:#f9f9f9;}';
  html += '</style></head><body>';
  html += '<h2>REPORTE DE TURNO ' + filtro + '</h2>';
  html += '<p><strong>Consecutivo:</strong> ' + consecutivo + '</p>';
  html += '<p><strong>Fecha:</strong> ' + new Date().toLocaleString() + '</p>';
  html += '<table><thead><tr>';
  html += '<th>Cédula</th><th>Grado</th><th>Funcionario</th><th>Dependencia</th><th>Turno</th><th>Novedades</th>';
  html += '</tr></thead><tbody>';

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

  html += '</tbody></table></body></html>';

  var blob = Utilities.newBlob(html, 'text/html', 'reporte.html');
  return blob.getAs('application/pdf');
}

function generarExcelBlob_(funcs, filtro, consecutivo) {
  var ss = SpreadsheetApp.create('Reporte_Turno_' + filtro + '_' + consecutivo);
  var hoja = ss.getActiveSheet();

  // Encabezados
  hoja.appendRow(['CÉDULA', 'GRADO', 'FUNCIONARIO', 'DEPENDENCIA', 'TURNO', 'NOVEDADES']);

  // Datos
  funcs.forEach(function(f) {
    var novedadesTexto = (f.historialNovedades || []).map(function(n) { return n.NOVEDAD; }).join(', ');
    hoja.appendRow([
      f.cedula || f.CEDULA || '',
      f.grado || f.GR || '',
      f.funcionario || f.FUNCIONARIO || '',
      f.dependencia || f.DEPENDENCIA || '',
      f.turno || f.TURNO || '',
      novedadesTexto
    ]);
  });

  // Formato
  var rango = hoja.getRange(1, 1, 1, 6);
  rango.setFontWeight('bold');
  rango.setBackground('#01592F');
  rango.setFontColor('white');

  var csv = '';
  var datos = hoja.getDataRange().getValues();
  for (var i = 0; i < datos.length; i++) {
    csv += datos[i].join(',') + '\n';
  }

  var blob = Utilities.newBlob(csv, 'text/csv', 'Reporte_Turno_' + filtro + '_' + consecutivo + '.csv');
  DriveApp.getFileById(ss.getId()).setTrashed(true);

  return blob;
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
        TURNO: datos[i][3],
        URL: datos[i][4] || ''
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
