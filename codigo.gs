/*******************************************************
 * SISTEMA DE GESTIÓN DE NOVEDADES
 * POLICÍA NACIONAL — JESEP / GUTAH
 * CÓDIGO.GS — VERSIÓN COMPLETA (AJUSTADA)
 *
 * Cambios respecto a la versión anterior:
 *  1. El token fijo del módulo Fuerza Disponible solo puede usar
 *     una lista corta de acciones (ver ACCIONES_TOKEN_DISPONIBLE).
 *  2. previsualizarSoloNovedades ahora filtra por turno y escapa HTML.
 *  3. consultarPorTurno acepta "A" o "TURNO A" en la hoja.
 *  4. Se limita el largo de los textos recibidos al registrar novedades.
 *******************************************************/

const HOJA_BASE       = 'LISTADO_BASE';
const HOJA_NOVEDADES  = 'NOVEDADES';
const HOJA_USUARIOS   = 'USUARIOS';
const HOJA_AUDITORIA  = 'AUDITORIA';
const HOJA_REPORTES   = 'REPORTES';

/* =====================================================
   TOKEN FIJO DEL MÓDULO FUERZA DISPONIBLE (index_1)
   ===================================================== */
const TOKEN_DISPONIBLE = 'MODULO_DISPONIBLE_AUTOMATICO';

// Acciones de solo lectura que puede usar el token fijo
const ACCIONES_TOKEN_DISPONIBLE = [
  'consultarPorTurno',
  'buscarFuncionario',
  'obtenerFichaFuncionario',
  'obtenerHistorialFuncionario',
  'previsualizarSoloNovedades',
  'obtenerTiposNovedadPermitidos',
  'verificarSesion'
];

// true  = el index_1 puede guardar novedades (registrarNovedad) con el token fijo
// false = el token fijo es 100% solo lectura (el botón "Guardar novedad" dejará de funcionar)
const PERMITIR_REGISTRO_CON_TOKEN_FIJO = true;

// Tipos de novedad que el token fijo puede registrar (los del select del HTML)
const TIPOS_TOKEN_FIJO = ['SERVICIO', 'PERMISO', 'OTRA NOVEDAD'];

// Largo máximo de los textos libres que llegan desde el cliente
const MAX_TEXTO = 300;

function accionPermitidaConTokenFijo_(accion) {
  if (ACCIONES_TOKEN_DISPONIBLE.indexOf(accion) !== -1) return true;
  if (PERMITIR_REGISTRO_CON_TOKEN_FIJO && accion === 'registrarNovedad') return true;
  return false;
}

/* =====================================================
   CONTROL DE ACCESO POR MÓDULO
   ===================================================== */
const USUARIOS_INDEX_1 = ['DISPONIBLE_A', 'DISPONIBLE_B', 'DISPONIBLE_C'];

const USUARIOS_INDEX_2 = [
  'SGSST_JESEP', 'VAC_JESEP', 'PAS_JESEP', 'CIT_JESEP',
  'PRO_JESEP', 'UBL_JESEP', 'GH_JESEP', 'CAP_JESEP'
];

const DURACION_SESION_MS = 8 * 60 * 60 * 1000; // 8 horas

function claveUsuario_(valor) {
  return normalizarTexto_(valor).replace(/_+/g, '_');
}

function esAdministrador_(rol) {
  return normalizarTexto_(rol) === 'ADMINISTRADOR';
}

function modulosPermitidos_(usuario, rol) {
  if (esAdministrador_(rol)) return ['index_1', 'index_2'];
  const clave = claveUsuario_(usuario);
  const modulos = [];
  if (USUARIOS_INDEX_1.map(claveUsuario_).indexOf(clave) !== -1) modulos.push('index_1');
  if (USUARIOS_INDEX_2.map(claveUsuario_).indexOf(clave) !== -1) modulos.push('index_2');
  return modulos;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.accion) {
    return ContentService.createTextOutput('API JESEP activa').setMimeType(ContentService.MimeType.TEXT);
  }
  return manejarLlamadaAPI_(e);
}

function obtenerAccionesPermitidas_() {
  return {
    validarUsuario: validarUsuario,
    cerrarSesionCliente: cerrarSesionCliente,
    verificarSesion: verificarSesion,
    buscarFuncionario: buscarFuncionario,
    consultarPorTurno: consultarPorTurno,
    obtenerFichaFuncionario: obtenerFichaFuncionario,
    obtenerHistorialFuncionario: obtenerHistorialFuncionario,
    registrarNovedad: registrarNovedad,
    obtenerTiposNovedadPermitidos: obtenerTiposNovedadPermitidos,
    registrarNovedadPorUsuario: registrarNovedadPorUsuario,
    previsualizarSoloNovedades: previsualizarSoloNovedades,
    agregarFuncionarioUBL: agregarFuncionarioUBL,
    eliminarFuncionarioUBL: eliminarFuncionarioUBL,
    eliminarNovedadUBL: eliminarNovedadUBL
  };
}

function manejarLlamadaAPI_(e) {
  const callbackCrudo = texto_(e.parameter.callback);
  const callback = /^[a-zA-Z0-9_]+$/.test(callbackCrudo) ? callbackCrudo : '';

  if (!callback) {
    return ContentService.createTextOutput('console.error("Callback inválido.");').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  let resultado;
  try {
    const accion = texto_(e.parameter.accion);
    const funcion = obtenerAccionesPermitidas_()[accion];
    if (!funcion) throw new Error('Acción no permitida: ' + accion);

    let argumentos = [];
    if (e.parameter.args) argumentos = JSON.parse(e.parameter.args);
    if (!Array.isArray(argumentos)) argumentos = [];

    // El token fijo solo puede ejecutar las acciones autorizadas
    if (argumentos[0] === TOKEN_DISPONIBLE && !accionPermitidaConTokenFijo_(accion)) {
      throw new Error('Esta acción no está permitida para el módulo Fuerza Disponible.');
    }

    resultado = funcion.apply(null, argumentos);
    if (resultado === undefined) resultado = null;
  } catch (error) {
    resultado = { __jsonp_error: true, estado: false, mensaje: error.message };
  }

  return ContentService.createTextOutput(callback + '(' + JSON.stringify(resultado) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSS_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function obtenerHoja_(nombre) { return getSS_().getSheetByName(nombre); }

function texto_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function limitar_(valor) {
  return texto_(valor).substring(0, MAX_TEXTO);
}

function normalizarTexto_(valor) {
  return String(valor || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
}

function normalizarCedula_(valor) {
  return String(valor || '').replace(/[.\-\s]/g, '').trim();
}

// "TURNO A", "turno a", " A " -> "A"
function normalizarTurno_(valor) {
  return normalizarTexto_(valor).replace(/^TURNO\s*/, '').trim();
}

function escaparHtml_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function formatearCelda_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return valor;
}

function sha256_(texto) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(texto || ''), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}

function respuestaOK_(mensaje, datos) {
  return { estado: true, mensaje: mensaje || '', datos: datos !== undefined ? datos : null };
}

function respuestaError_(mensaje) {
  return { estado: false, mensaje: mensaje || 'Error en la operación.' };
}

function crearSesion_(usuario) {
  const token = sha256_(Utilities.getUuid() + '|' + new Date().getTime());
  const json = JSON.stringify({ usuario: usuario.usuario, rol: usuario.rol, dependencia: usuario.dependencia, modulos: usuario.modulos || [], creado: new Date().getTime() });
  PropertiesService.getScriptProperties().setProperty('SESION_' + token, json);
  return token;
}

function obtenerSesion_(token) {
  if (!token) return null;
  const valor = PropertiesService.getScriptProperties().getProperty('SESION_' + token);
  if (!valor) return null;
  const sesion = JSON.parse(valor);
  if (new Date().getTime() - sesion.creado > DURACION_SESION_MS) {
    PropertiesService.getScriptProperties().deleteProperty('SESION_' + token);
    return null;
  }
  return sesion;
}

/* =====================================================
   VALIDACIÓN DE SESIÓN
   (el token fijo queda restringido en manejarLlamadaAPI_)
   ===================================================== */
function validarSesion_(token) {
  if (token === TOKEN_DISPONIBLE) {
    return { usuario: 'DISPONIBLE_A', rol: 'OPERADOR', dependencia: 'JESEP', modulos: ['index_1'] };
  }
  const sesion = obtenerSesion_(token);
  if (!sesion) throw new Error('Sesión vencida o no válida.');
  return sesion;
}

function cerrarSesionCliente(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('SESION_' + token);
  return { estado: true };
}

function validarUsuario(usuario, clave) {
  const hoja = obtenerHoja_(HOJA_USUARIOS);
  const datos = hoja.getDataRange().getValues();
  const bus = claveUsuario_(usuario);
  const pass = sha256_(clave);

  for (let i = 1; i < datos.length; i++) {
    if (claveUsuario_(datos[i][0]) === bus && datos[i][1] === pass) {
      const u = { usuario: datos[i][0], rol: datos[i][2], dependencia: datos[i][3] };
      u.modulos = modulosPermitidos_(u.usuario, u.rol);

      if (!u.modulos.length) {
        throw new Error('Su usuario no tiene un módulo asignado. Comuníquese con el administrador.');
      }

      const token = crearSesion_(u);
      return { estado: true, token: token, usuario: u, modulos: u.modulos };
    }
  }
  throw new Error('Usuario o contraseña incorrectos.');
}

function verificarSesion(token, modulo) {
  const sesion = validarSesion_(token);
  if (modulo && sesion.modulos.indexOf(modulo) === -1) {
    throw new Error('No tiene autorización para este módulo.');
  }
  return { estado: true, usuario: { usuario: sesion.usuario, rol: sesion.rol, dependencia: sesion.dependencia }, modulos: sesion.modulos };
}

function buscarFuncionario(token, valor) {
  validarSesion_(token);
  const hoja = obtenerHoja_(HOJA_BASE);
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return respuestaOK_('Resultados de búsqueda', []);

  const enc = datos[0].map(v => normalizarTexto_(v));
  const iCc = enc.indexOf('CEDULA') >= 0 ? enc.indexOf('CEDULA') : 0;
  const iGr = enc.indexOf('GR') >= 0 ? enc.indexOf('GR') : 2;
  const iNom = enc.indexOf('FUNCIONARIO') >= 0 ? enc.indexOf('FUNCIONARIO') : 3;
  const iDep = enc.indexOf('DEPENDENCIA') >= 0 ? enc.indexOf('DEPENDENCIA') : (enc.indexOf('DEP') >= 0 ? enc.indexOf('DEP') : 4);
  const iPert = enc.indexOf('PERT') >= 0 ? enc.indexOf('PERT') : 6;
  const iTur = enc.indexOf('TURNO') >= 0 ? enc.indexOf('TURNO') : 7;

  const crit = normalizarTexto_(valor);
  if (!crit) return respuestaOK_('Resultados de búsqueda', []);
  const res = [];

  for (let i = 1; i < datos.length; i++) {
    const cc = normalizarCedula_(datos[i][iCc]);
    const nom = normalizarTexto_(datos[i][iNom]);
    const dep = normalizarTexto_(datos[i][iDep]);
    const pert = normalizarTexto_(datos[i][iPert]);
    const gr = normalizarTexto_(datos[i][iGr]);

    if (cc.indexOf(crit) !== -1 || nom.indexOf(crit) !== -1 || dep.indexOf(crit) !== -1 || pert === crit || pert.indexOf(crit) !== -1 || gr.indexOf(crit) !== -1) {
      res.push({
        cedula: cc,
        grado: datos[i][iGr],
        funcionario: datos[i][iNom],
        dependencia: datos[i][iDep],
        pert: datos[i][iPert],
        turno: datos[i][iTur]
      });
    }
  }
  return respuestaOK_('Resultados de búsqueda', res);
}

/* =====================================================
   CONSULTAR POR TURNO (Fuerza Disponible)
   ===================================================== */
function consultarPorTurno(token, turno) {
  try {
    validarSesion_(token);
    const hoja = obtenerHoja_(HOJA_BASE);
    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return respuestaOK_('Turno vacío', []);

    const enc = datos[0].map(v => normalizarTexto_(v));
    const iCc = enc.indexOf('CEDULA') >= 0 ? enc.indexOf('CEDULA') : 0;
    const iGr = enc.indexOf('GR') >= 0 ? enc.indexOf('GR') : 2;
    const iNom = enc.indexOf('FUNCIONARIO') >= 0 ? enc.indexOf('FUNCIONARIO') : 3;
    const iDep = enc.indexOf('DEPENDENCIA') >= 0 ? enc.indexOf('DEPENDENCIA') : (enc.indexOf('DEP') >= 0 ? enc.indexOf('DEP') : 4);
    const iTur = enc.indexOf('TURNO') >= 0 ? enc.indexOf('TURNO') : 7;

    const turnoBusqueda = normalizarTurno_(turno);
    const res = [];

    for (let i = 1; i < datos.length; i++) {
      const tur = normalizarTurno_(datos[i][iTur]);
      if (tur === turnoBusqueda) {
        res.push({
          cedula: normalizarCedula_(datos[i][iCc]),
          grado: datos[i][iGr],
          funcionario: datos[i][iNom],
          dependencia: datos[i][iDep],
          turno: datos[i][iTur]
        });
      }
    }
    return respuestaOK_('Fuerza del turno ' + turno, res);
  } catch (error) {
    return respuestaError_('Error en servidor: ' + error.message);
  }
}

function obtenerFichaFuncionario(token, cedula) {
  validarSesion_(token);
  const hoja = obtenerHoja_(HOJA_BASE);
  const datos = hoja.getDataRange().getValues();
  const cc = normalizarCedula_(cedula);

  const enc = datos[0].map(v => normalizarTexto_(v));
  const iCc = enc.indexOf('CEDULA') >= 0 ? enc.indexOf('CEDULA') : 0;
  const iNiv = enc.indexOf('NIV') >= 0 ? enc.indexOf('NIV') : 1;
  const iGr = enc.indexOf('GR') >= 0 ? enc.indexOf('GR') : 2;
  const iNom = enc.indexOf('FUNCIONARIO') >= 0 ? enc.indexOf('FUNCIONARIO') : 3;
  const iDep = enc.indexOf('DEPENDENCIA') >= 0 ? enc.indexOf('DEPENDENCIA') : 4;
  const iPert = enc.indexOf('PERT') >= 0 ? enc.indexOf('PERT') : 6;
  const iTur = enc.indexOf('TURNO') >= 0 ? enc.indexOf('TURNO') : 7;

  for (let i = 1; i < datos.length; i++) {
    if (normalizarCedula_(datos[i][iCc]) === cc) {
      return {
        estado: true,
        funcionario: {
          cedula: cc,
          nivel: datos[i][iNiv],
          grado: datos[i][iGr],
          funcionario: datos[i][iNom],
          dependencia: datos[i][iDep],
          pert: datos[i][iPert],
          turno: datos[i][iTur]
        }
      };
    }
  }
  return respuestaError_('Funcionario no encontrado.');
}

function obtenerHistorialFuncionario(token, cedula) {
  validarSesion_(token);
  const hoja = obtenerHoja_(HOJA_NOVEDADES);
  const datos = hoja.getDataRange().getValues();
  const cc = normalizarCedula_(cedula);
  const hist = [];

  for (let i = 1; i < datos.length; i++) {
    if (normalizarCedula_(datos[i][2]) === cc) {
      hist.push({
        fila: i + 1,
        GR: datos[i][0],
        NOVEDAD: datos[i][4],
        Dias: datos[i][6],
        'Fecha INICIAL': formatearCelda_(datos[i][7]),
        'Fecha PRESENTACION': formatearCelda_(datos[i][8])
      });
    }
  }
  return { estado: true, historial: hist };
}

const TIPOS_NOVEDAD_POR_USUARIO = {
  'SGSST_JESEP': ['EXCUSA MEDICA', 'RESTRICCIONES MEDICA', 'LICIENCIA DE MATERNIDAD'],
  'VAC_JESEP': ['PLAN VACACIONAL', 'VACACIONES EXTRAORDINARIAS', 'PLAN REDUCCION', 'VACACIONES DE RETIRO'],
  'PAS_JESEP': ['COMISION DE ESTUDIO', 'COMISION DE SERVICIO', 'LICENCIA DE PATERNIDAD', 'LICENCIA DE LUTO'],
  'CIT_JESEP': ['SUSPENCION', 'CITACION JUDICIAL (PERMISO)'],
  'HIS_JESEP': ['RETIROS 3 MESES DE ALTA', 'ELIMINAR USURIO POR RETIRO'],
  'PRO_JESEP': ['CURSO DE ASCENSO ESPOL','CURSO DE ASCENSO ESJIM'],
  'GH_JESEP': ['CAMBIO DE TURNO', 'HORARIO FLEXIBLE','PERMISO'],
  'CAP_JESEP': ['CURSO MANDATORIO', 'CAPACITACION']
};

function obtenerTiposNovedadPermitidos(token) {
  const sesion = validarSesion_(token);
  const usr = claveUsuario_(sesion.usuario);
  const tipos = TIPOS_NOVEDAD_POR_USUARIO[usr] || ['SERVICIO', 'PERMISO', 'OTRA NOVEDAD'];
  return respuestaOK_('Tipos asignados', { tipos: tipos });
}

function registrarNovedadPorUsuario(token, datos) {
  const sesion = validarSesion_(token);
  const tipo = datos.novedad;

  if (tipo === 'CAMBIO DE TURNO') {
    const hojaBase = obtenerHoja_(HOJA_BASE);
    const datosBase = hojaBase.getDataRange().getValues();
    const cc = normalizarCedula_(datos.cc || datos.cedula);
    const nuevoTurno = datos.nuevoTurno;

    const enc = datosBase[0].map(v => normalizarTexto_(v));
    const iCc = enc.indexOf('CEDULA') >= 0 ? enc.indexOf('CEDULA') : 0;
    const iTur = enc.indexOf('TURNO') >= 0 ? enc.indexOf('TURNO') : 7;

    for (let i = 1; i < datosBase.length; i++) {
      if (normalizarCedula_(datosBase[i][iCc]) === cc) {
        hojaBase.getRange(i + 1, iTur + 1).setValue(nuevoTurno);
        return respuestaOK_('Cambio de turno registrado en LISTADO_BASE correctamente.');
      }
    }
    return respuestaError_('Funcionario no encontrado.');
  }

  return registrarNovedad(token, datos);
}

function registrarNovedad(token, datos) {
  validarSesion_(token);

  // Con el token fijo solo se permiten los tipos del módulo Fuerza Disponible
  if (token === TOKEN_DISPONIBLE &&
      TIPOS_TOKEN_FIJO.indexOf(normalizarTexto_(datos.novedad)) === -1) {
    return respuestaError_('Tipo de novedad no permitido para este módulo.');
  }

  const cc = normalizarCedula_(datos.cc || datos.cedula || datos.CEDULA);
  if (!cc) return respuestaError_('No se recibió número de cédula válido.');

  const hojaBase = obtenerHoja_(HOJA_BASE);
  const datosBase = hojaBase.getDataRange().getValues();
  if (datosBase.length < 2) return respuestaError_('HOJA LISTADO_BASE vacía.');

  const encBase = datosBase[0].map(v => normalizarTexto_(v));
  const iCc = encBase.indexOf('CEDULA') >= 0 ? encBase.indexOf('CEDULA') : 0;
  const iGr = encBase.indexOf('GR') >= 0 ? encBase.indexOf('GR') : 2;
  const iNom = encBase.indexOf('FUNCIONARIO') >= 0 ? encBase.indexOf('FUNCIONARIO') : 3;
  const iDep = encBase.indexOf('DEPENDENCIA') >= 0 ? encBase.indexOf('DEPENDENCIA') : 4;
  const iNiv = encBase.indexOf('NIV') >= 0 ? encBase.indexOf('NIV') : 1;
  const iTur = encBase.indexOf('TURNO') >= 0 ? encBase.indexOf('TURNO') : 7;

  let gr = datos.grado || datos.GR || '';
  let funcionario = datos.funcionario || datos.FUNCIONARIO || '';
  let dependencia = datos.dependencia || '';
  let nivel = datos.nivel || '';
  let turno = datos.turno || '';

  for (let i = 1; i < datosBase.length; i++) {
    if (normalizarCedula_(datosBase[i][iCc]) === cc) {
      gr = datosBase[i][iGr];
      funcionario = datosBase[i][iNom];
      dependencia = datosBase[i][iDep];
      nivel = datosBase[i][iNiv];
      turno = datosBase[i][iTur];
      break;
    }
  }

  if (!funcionario) {
    return respuestaError_('El funcionario con CC ' + cc + ' no existe en LISTADO_BASE.');
  }

  const tipoNovedadColE = limitar_(datos.novedad);

  let descripcionColF = '';
  const reub = limitar_(datos.descripcion);
  const detalleExtra = limitar_(datos.nombreNovedad);

  if (reub && detalleExtra) {
    descripcionColF = reub + ' "' + detalleExtra + '"';
  } else if (reub) {
    descripcionColF = reub;
  } else {
    descripcionColF = detalleExtra;
  }

  const hojaNov = obtenerHoja_(HOJA_NOVEDADES);

  const nuevaFila = [
    gr, funcionario, cc, dependencia, tipoNovedadColE, descripcionColF,
    limitar_(datos.dias), limitar_(datos.fechaInicial), limitar_(datos.fechaPresentacion),
    '', '', nivel, turno, '', ''
  ];

  hojaNov.appendRow(nuevaFila);
  return respuestaOK_('Novedad guardada exitosamente.');
}

/* =====================================================
   PREVISUALIZAR NOVEDADES (filtradas por turno)
   ===================================================== */
function previsualizarSoloNovedades(token, filtro) {
  validarSesion_(token);
  const turnoFiltro = normalizarTurno_(filtro);

  // Mapa cédula -> turno actual desde LISTADO_BASE (por si la fila de NOVEDADES no tiene turno)
  const mapaTurno = {};
  const datosBase = obtenerHoja_(HOJA_BASE).getDataRange().getValues();
  if (datosBase.length > 1) {
    const enc = datosBase[0].map(v => normalizarTexto_(v));
    const iCc = enc.indexOf('CEDULA') >= 0 ? enc.indexOf('CEDULA') : 0;
    const iTur = enc.indexOf('TURNO') >= 0 ? enc.indexOf('TURNO') : 7;
    for (let i = 1; i < datosBase.length; i++) {
      mapaTurno[normalizarCedula_(datosBase[i][iCc])] = normalizarTurno_(datosBase[i][iTur]);
    }
  }

  const datosNov = obtenerHoja_(HOJA_NOVEDADES).getDataRange().getValues();
  let filasHTML = '';

  for (let i = 1; i < datosNov.length; i++) {
    const cc = normalizarCedula_(datosNov[i][2]);
    const turnoFila = normalizarTurno_(datosNov[i][12]) || mapaTurno[cc] || '';
    if (turnoFiltro && turnoFila !== turnoFiltro) continue;

    filasHTML += '<tr>' +
      '<td>' + escaparHtml_(datosNov[i][0]) + '</td>' +
      '<td>' + escaparHtml_(datosNov[i][1]) + '</td>' +
      '<td>' + escaparHtml_(datosNov[i][2]) + '</td>' +
      '<td>' + escaparHtml_(datosNov[i][4]) + '</td>' +
      '<td>' + escaparHtml_(datosNov[i][5]) + '</td>' +
      '<td>' + escaparHtml_(datosNov[i][6]) + '</td>' +
      '<td>' + escaparHtml_(formatearCelda_(datosNov[i][7])) + '</td>' +
      '</tr>';
  }

  const html = '<style>.tbl-rep{width:100%;border-collapse:collapse;}.tbl-rep th,.tbl-rep td{border:1px solid #ddd;padding:8px;font-size:12px;}.tbl-rep th{background:#01592F;color:white;}</style>' +
    '<h5>Reporte de Novedades Registradas - Turno ' + escaparHtml_(turnoFiltro) + '</h5>' +
    '<div class="table-wrapper"><table class="tbl-rep"><thead><tr><th>GR</th><th>FUNCIONARIO</th><th>CÉDULA</th><th>NOVEDAD</th><th>DETALLE</th><th>DÍAS</th><th>INICIO</th></tr></thead><tbody>' +
    (filasHTML || '<tr><td colspan="7">No hay novedades registradas.</td></tr>') +
    '</tbody></table></div>';

  return { estado: true, html: html };
}

function agregarFuncionarioUBL(token, datos) {
  const sesion = validarSesion_(token);
  if (claveUsuario_(sesion.usuario) !== 'UBL_JESEP' && !esAdministrador_(sesion.rol)) {
    throw new Error('No autorizado.');
  }

  const hoja = obtenerHoja_(HOJA_BASE);
  const fila = [
    datos.cedula, datos.niv, datos.gr, datos.funcionario, datos.dependencia,
    datos.dependencia, datos.pert, datos.turno, '', '', datos.mes, datos.dia,
    datos.fechaNacimiento, datos.correo, datos.sexo, datos.estadoCivil,
    datos.situacionLaboral, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
    datos.comunicadoOficial || ''
  ];

  hoja.appendRow(fila);
  return respuestaOK_('Funcionario agregado exitosamente a LISTADO_BASE.');
}

function eliminarFuncionarioUBL(token, cedula) {
  const sesion = validarSesion_(token);
  if (claveUsuario_(sesion.usuario) !== 'UBL_JESEP' && !esAdministrador_(sesion.rol)) {
    throw new Error('No autorizado.');
  }

  const hoja = obtenerHoja_(HOJA_BASE);
  const datos = hoja.getDataRange().getValues();
  const cc = normalizarCedula_(cedula);

  for (let i = 1; i < datos.length; i++) {
    if (normalizarCedula_(datos[i][0]) === cc) {
      hoja.deleteRow(i + 1);
      return respuestaOK_('Funcionario eliminado correctamente de LISTADO_BASE.');
    }
  }
  return respuestaError_('Funcionario no encontrado.');
}

function eliminarNovedadUBL(token, numFila) {
  const sesion = validarSesion_(token);
  if (claveUsuario_(sesion.usuario) !== 'UBL_JESEP' && !esAdministrador_(sesion.rol)) {
    throw new Error('No autorizado.');
  }

  const fila = parseInt(numFila, 10);
  const hoja = obtenerHoja_(HOJA_NOVEDADES);
  if (!fila || fila < 2 || fila > hoja.getLastRow()) {
    return respuestaError_('Número de fila no válido.');
  }

  hoja.deleteRow(fila);
  return respuestaOK_('Novedad de la fila ' + fila + ' eliminada correctamente.');
}
