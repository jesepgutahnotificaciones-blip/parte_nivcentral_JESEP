/*******************************************************
 * SISTEMA DE GESTIÓN DE NOVEDADES
 * POLICÍA NACIONAL — JESEP / GUTAH
 *
 * CÓDIGO.GS
 * PARTE 1 — CONFIGURACIÓN, UTILIDADES Y SEGURIDAD
 *
 * HOJAS:
 *   LISTADO_BASE
 *   NOVEDADES
 *   USUARIOS
 *   AUDITORIA
 *   REPORTES
 *   CONFIG
 *
 * IMPORTANTE:
 * Este archivo debe construirse por partes.
 * NO mezclar con versiones anteriores.
 *******************************************************/


/* =====================================================
   1. CONFIGURACIÓN GENERAL
   ===================================================== */

const HOJA_BASE       = 'LISTADO_BASE';
const HOJA_NOVEDADES  = 'NOVEDADES';
const HOJA_USUARIOS   = 'USUARIOS';
const HOJA_AUDITORIA  = 'AUDITORIA';
const HOJA_REPORTES   = 'REPORTES';
const HOJA_CONFIG     = 'CONFIG';


/*
 * Tiempo de duración de la sesión.
 * 6 horas.
 */
const SESION_HORAS = 6;

const SESION_SEGUNDOS =
  SESION_HORAS * 60 * 60;


/*
 * Los tres turnos básicos.
 *
 * NO se utilizarán para limitar
 * la consulta por los nuevos filtros.
 *
 * Los filtros A, B, C, SEPRI,
 * GURIN y NO APLICA serán tratados
 * específicamente en consultarPorTurno().
 */
const TURNOS_VALIDOS = [
  'A',
  'B',
  'C'
];


/* =====================================================
   2. ESTRUCTURA DE LA HOJA NOVEDADES
   ===================================================== */

const ENCABEZADOS_NOVEDADES = [

  'GR',
  'APELLIDOS Y NOMBRES',
  'CC',
  'DEPENDENCIA',
  'NOVEDAD',
  'DESCRIPCION',
  'Dias',
  'Fecha INICIAL',
  'Fecha PRESENTACION',
  'Observacion',
  'RV',
  'NIV',
  'Turno',
  'Placa_Chip',
  'TEXTO'

];


/* =====================================================
   3. DOGET
   ===================================================== */

function doGet(e) {

  /*
   * Si viene "accion" y "callback" en la URL, es una
   * llamada de API/JSONP desde el frontend externo
   * (GitHub Pages), no una carga de la página HTML.
   */
  if (
    e &&
    e.parameter &&
    e.parameter.accion &&
    e.parameter.callback
  ) {

    return manejarLlamadaAPI_(e);

  }


  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle(
      'Sistema de Gestión de Novedades'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* =====================================================
   3B. API EXTERNA (JSONP) PARA GITHUB PAGES
   ===================================================== */

/*
 * Lista blanca de funciones que pueden invocarse desde
 * el frontend externo. Cualquier nombre que no esté
 * aquí se rechaza, aunque exista como función en el
 * proyecto — esto evita exponer funciones internas
 * (las que terminan en "_") o funciones de
 * administración no pensadas para llamarse así.
 */
function obtenerAccionesPermitidas_() {

  return {

    validarUsuario: validarUsuario,
    cerrarSesionCliente: cerrarSesionCliente,
    obtenerDatosSesion: obtenerDatosSesion,

    buscarFuncionario: buscarFuncionario,
    obtenerFichaFuncionario: obtenerFichaFuncionario,
    obtenerHistorialFuncionario: obtenerHistorialFuncionario,

    registrarNovedad: registrarNovedad,

    consultarPorTurno: consultarPorTurno,
    consultarTurno: consultarTurno,

    previsualizarReporteTurno: previsualizarReporteTurno,
    listarReportes: listarReportes,

    listarUsuarios: listarUsuarios,
    crearUsuario: crearUsuario,
    cambiarEstadoUsuario: cambiarEstadoUsuario

  };

}


function manejarLlamadaAPI_(e) {

  const callbackCrudo =
    texto_(e.parameter.callback);


  /*
   * El nombre de callback se inserta directamente en
   * la respuesta JavaScript, así que se valida de forma
   * estricta para evitar inyección de código.
   */
  const callback =
    /^[a-zA-Z0-9_]+$/.test(callbackCrudo)
      ? callbackCrudo
      : '';

  if (!callback) {

    return ContentService
      .createTextOutput(
        'console.error("Callback inválido.");'
      )
      .setMimeType(
        ContentService.MimeType.JAVASCRIPT
      );

  }


  let resultado;

  try {

    const accion =
      texto_(e.parameter.accion);

    const funciones =
      obtenerAccionesPermitidas_();

    const funcion =
      funciones[accion];

    if (!funcion) {

      throw new Error(
        'Acción no permitida: ' + accion
      );

    }


    let argumentos = [];

    if (e.parameter.args) {

      try {

        argumentos =
          JSON.parse(e.parameter.args);

      } catch (errorParseo) {

        throw new Error(
          'Argumentos inválidos.'
        );

      }

    }

    if (!Array.isArray(argumentos)) {
      argumentos = [];
    }


    resultado =
      funcion.apply(null, argumentos);

    if (resultado === undefined) {
      resultado = null;
    }

  } catch (error) {

    resultado = {

      __jsonp_error: true,

      estado: false,

      mensaje:
        error && error.message
          ? error.message
          : 'Se presentó un error.'

    };

  }


  const cuerpo =
    callback +
    '(' +
    JSON.stringify(resultado) +
    ');';


  return ContentService
    .createTextOutput(cuerpo)
    .setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );

}


/* =====================================================
   4. INCLUIR ARCHIVOS HTML
   ===================================================== */

function include(nombre) {

  return HtmlService
    .createHtmlOutputFromFile(nombre)
    .getContent();

}


/* =====================================================
   5. CONEXIÓN CON EL SPREADSHEET
   ===================================================== */

function getSS_() {

  return SpreadsheetApp
    .getActiveSpreadsheet();

}


/* =====================================================
   6. OBTENER HOJA
   ===================================================== */

function obtenerHoja_(nombre) {

  const ss = getSS_();

  const hoja =
    ss.getSheetByName(nombre);

  if (!hoja) {

    throw new Error(
      'No existe la hoja: ' + nombre
    );

  }

  return hoja;

}


/* =====================================================
   7. OBTENER DATOS DE UNA HOJA
   ===================================================== */

function obtenerDatosHoja_(nombre) {

  const hoja =
    obtenerHoja_(nombre);

  const rango =
    hoja.getDataRange();

  return rango.getValues();

}


/* =====================================================
   8. CONVERSIÓN SEGURA A TEXTO
   ===================================================== */

function texto_(valor) {

  if (
    valor === null ||
    valor === undefined
  ) {

    return '';

  }

  return String(valor).trim();

}


/* =====================================================
   9. NORMALIZACIÓN DE TEXTO
   ===================================================== */

function normalizarTexto_(valor) {

  return String(valor || '')
    .normalize('NFD')
    .replace(
      /[\u0300-\u036f]/g,
      ''
    )
    .trim()
    .toUpperCase();

}


/* =====================================================
   10. NORMALIZACIÓN DE CÉDULA
   ===================================================== */

function normalizarCedula_(valor) {

  return String(valor || '')
    .replace(/[.\-\s]/g, '')
    .trim();

}


/* =====================================================
   11. NORMALIZACIÓN DE NOMBRE
   ===================================================== */

function normalizarNombre_(valor) {

  return normalizarTexto_(valor)
    .replace(/\s+/g, ' ')
    .trim();

}


/* =====================================================
   12. HASH SHA-256
   ===================================================== */

function sha256_(texto) {

  texto =
    String(texto || '');

  const bytes =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      texto,
      Utilities.Charset.UTF_8
    );

  return bytes
    .map(function(byte) {

      const valor =
        byte < 0
          ? byte + 256
          : byte;

      return (
        '0' +
        valor.toString(16)
      ).slice(-2);

    })
    .join('');

}


/* =====================================================
   13. GENERAR TOKEN DE SESIÓN
   ===================================================== */

function generarTokenSesion_() {

  const datos =
    Utilities.getUuid() +
    '|' +
    new Date().getTime() +
    '|' +
    Math.random();

  return sha256_(datos);

}


/* =====================================================
   14. CREAR SESIÓN
   ===================================================== */

function crearSesion_(usuario) {

  const token =
    generarTokenSesion_();

  const datosSesion = {

    usuario:
      usuario.usuario,

    rol:
      usuario.rol,

    dependencia:
      usuario.dependencia,

    creado:
      new Date().getTime()

  };


  /*
   * Se usa PropertiesService en lugar de CacheService:
   * CacheService no garantiza lectura inmediata
   * después de escribir (comportamiento "best-effort"),
   * lo que provocaba sesiones inválidas justo después
   * de iniciar sesión. PropertiesService es persistente
   * y consistente. La expiración se controla a mano
   * con el campo "creado" en obtenerSesion_().
   */
  PropertiesService
    .getScriptProperties()
    .setProperty(

      'SESION_' + token,

      JSON.stringify(
        datosSesion
      )

    );


  return token;

}


/* =====================================================
   15. OBTENER SESIÓN
   ===================================================== */

function obtenerSesion_(token) {

  if (!token) {

    return null;

  }


  const propiedades =
    PropertiesService
      .getScriptProperties();

  const clave =
    'SESION_' + token;

  const valor =
    propiedades.getProperty(
      clave
    );


  if (!valor) {

    return null;

  }


  let sesion;

  try {

    sesion =
      JSON.parse(
        valor
      );

  } catch (e) {

    return null;

  }


  /*
   * Expiración manual: PropertiesService
   * no expira solo, así que se controla
   * con el campo "creado".
   */
  const ahora =
    new Date().getTime();

  if (
    !sesion.creado ||
    (ahora - sesion.creado) > (SESION_SEGUNDOS * 1000)
  ) {

    propiedades.deleteProperty(
      clave
    );

    return null;

  }


  return sesion;

}


/* =====================================================
   16. VALIDAR SESIÓN
   ===================================================== */

function validarSesion_(token) {

  const sesion =
    obtenerSesion_(token);

  if (!sesion) {

    throw new Error(
      'Sesión no válida o expirada.'
    );

  }

  return sesion;

}

/* =====================================================
   17. CERRAR SESIÓN
   ===================================================== */

function cerrarSesion(token) {

  if (!token) {

    return {

      estado: true,

      mensaje:
        'Sesión cerrada.'

    };

  }


  PropertiesService
    .getScriptProperties()
    .deleteProperty(
      'SESION_' + token
    );


  return {

    estado: true,

    mensaje:
      'Sesión cerrada correctamente.'

  };

}


/* =====================================================
   18. VALIDAR ROL
   ===================================================== */

function validarRol_(sesion, rolesPermitidos) {

  if (!sesion) {

    throw new Error(
      'Sesión no válida.'
    );

  }


  const rol =
    normalizarTexto_(
      sesion.rol
    );


  const permitidos =
    rolesPermitidos.map(
      function(rolPermitido) {

        return normalizarTexto_(
          rolPermitido
        );

      }
    );


  if (
    permitidos.indexOf(rol) === -1
  ) {

    throw new Error(
      'No tiene permisos para realizar esta operación.'
    );

  }


  return true;

}


/* =====================================================
   19. VALIDAR TURNO BÁSICO
   ===================================================== */

function validarTurno_(turno) {

  const valor =
    normalizarTexto_(
      turno
    );


  if (
    TURNOS_VALIDOS.indexOf(valor) === -1
  ) {

    throw new Error(
      'El turno indicado no es válido.'
    );

  }


  return valor;

}


/* =====================================================
   20. FORMATEAR FECHA
   ===================================================== */

function formatearFecha_(fecha) {

  if (!fecha) {

    return '';

  }


  try {

    return Utilities.formatDate(

      new Date(fecha),

      Session.getScriptTimeZone(),

      'dd/MM/yyyy'

    );

  } catch (e) {

    return fecha;

  }

}


/* =====================================================
   21. FORMATEAR FECHA Y HORA
   ===================================================== */

function formatearFechaHora_(fecha) {

  if (!fecha) {

    return '';

  }


  try {

    return Utilities.formatDate(

      new Date(fecha),

      Session.getScriptTimeZone(),

      'dd/MM/yyyy HH:mm:ss'

    );

  } catch (e) {

    return fecha;

  }

}


/* =====================================================
   22. RESPUESTA ESTÁNDAR EXITOSA
   ===================================================== */

/*
 * Esta función acepta DOS formas:
 *
 * A)
 * respuestaOK_(
 *   'Mensaje',
 *   datos
 * )
 *
 * B)
 * respuestaOK_({
 *   estado: 'OK',
 *   datos: ...
 * })
 *
 * Se deja así para evitar incompatibilidades
 * entre las diferentes funciones del sistema.
 */

function respuestaOK_(mensaje, datos) {


  /*
   * FORMA B:
   *
   * respuestaOK_({
   *   estado: 'OK',
   *   ...
   * })
   */

  if (
    mensaje &&
    typeof mensaje === 'object' &&
    datos === undefined
  ) {

    return {

      estado: true,

      mensaje:
        mensaje.mensaje || '',

      datos:
        mensaje

    };

  }


  /*
   * FORMA A:
   *
   * respuestaOK_(
   *   'Mensaje',
   *   datos
   * )
   */

  return {

    estado: true,

    mensaje:
      mensaje || '',

    datos:
      datos !== undefined
        ? datos
        : null

  };

}


/* =====================================================
   23. RESPUESTA ESTÁNDAR DE ERROR
   ===================================================== */

function respuestaError_(mensaje, error) {

  return {

    estado: false,

    mensaje:
      mensaje ||
      (
        error &&
        error.message
          ? error.message
          : 'Se presentó un error.'
      ),

    datos: null

  };

}


/* =====================================================
   FIN PARTE 1
   ===================================================== */
   /* =====================================================
   24. VALIDACIÓN DE USUARIO
   ===================================================== */

function validarUsuario(usuario, clave) {

  usuario =
    String(usuario || '')
      .trim();

  clave =
    String(clave || '');


  if (!usuario) {

    throw new Error(
      'Debe ingresar el usuario.'
    );

  }


  if (!clave) {

    throw new Error(
      'Debe ingresar la contraseña.'
    );

  }


  const hoja =
    obtenerHoja_(
      HOJA_USUARIOS
    );


  const datos =
    hoja
      .getDataRange()
      .getValues();


  if (
    datos.length < 2
  ) {

    throw new Error(
      'No existen usuarios registrados en el sistema.'
    );

  }


  const encabezados =
    datos[0].map(
      function(valor) {

        return normalizarTexto_(
          valor
        );

      }
    );


  const indiceUsuario =
    encabezados.indexOf(
      'USUARIO'
    );

  const indicePassword =
    encabezados.indexOf(
      'PASSWORD_HASH'
    );

  const indiceRol =
    encabezados.indexOf(
      'ROL'
    );

  const indiceDependencia =
    encabezados.indexOf(
      'DEPENDENCIA'
    );

  const indiceEstado =
    encabezados.indexOf(
      'ESTADO'
    );

  const indiceUltimoAcceso =
    encabezados.indexOf(
      'ULTIMO_ACCESO'
    );


  if (
    indiceUsuario === -1 ||
    indicePassword === -1 ||
    indiceRol === -1 ||
    indiceDependencia === -1 ||
    indiceEstado === -1
  ) {

    throw new Error(
      'La hoja USUARIOS no tiene la estructura requerida.'
    );

  }


  const usuarioBuscado =
    normalizarTexto_(
      usuario
    );


  const hashIngresado =
    sha256_(
      clave
    );


  let usuarioEncontrado =
    null;


  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    const fila =
      datos[i];


    const usuarioHoja =
      normalizarTexto_(
        fila[indiceUsuario]
      );


    if (
      usuarioHoja ===
      usuarioBuscado
    ) {

      usuarioEncontrado = {

        fila: i + 1,

        usuario:
          String(
            fila[indiceUsuario] || ''
          ).trim(),

        passwordHash:
          String(
            fila[indicePassword] || ''
          ).trim(),

        rol:
          String(
            fila[indiceRol] || ''
          ).trim(),

        dependencia:
          String(
            fila[indiceDependencia] || ''
          ).trim(),

        estado:
          String(
            fila[indiceEstado] || ''
          ).trim(),

        indiceUltimoAcceso:
          indiceUltimoAcceso

      };

      break;

    }

  }


  if (!usuarioEncontrado) {

    throw new Error(
      'Usuario o contraseña incorrectos.'
    );

  }


  const estado =
    normalizarTexto_(
      usuarioEncontrado.estado
    );


  if (
    estado !== 'ACTIVO'
  ) {

    throw new Error(
      'El usuario se encuentra inactivo.'
    );

  }


  if (
    usuarioEncontrado.passwordHash !==
    hashIngresado
  ) {

    throw new Error(
      'Usuario o contraseña incorrectos.'
    );

  }


  const usuarioSesion = {

    usuario:
      usuarioEncontrado.usuario,

    rol:
      usuarioEncontrado.rol,

    dependencia:
      usuarioEncontrado.dependencia

  };


  const token =
    crearSesion_(
      usuarioSesion
    );


  /*
   * Actualizar último acceso.
   */

  if (
    usuarioEncontrado.indiceUltimoAcceso !== -1
  ) {

    hoja
      .getRange(

        usuarioEncontrado.fila,

        usuarioEncontrado.indiceUltimoAcceso + 1

      )
      .setValue(
        new Date()
      );

  }


  /*
   * Registrar ingreso en auditoría.
   */

  registrarAuditoria_(

    usuarioEncontrado.usuario,

    'INICIO DE SESIÓN',

    '',

    'Ingreso exitoso al sistema.'

  );


  return {

    estado: true,

    mensaje:
      'Inicio de sesión exitoso.',

    token:
      token,

    usuario: {

      usuario:
        usuarioEncontrado.usuario,

      rol:
        usuarioEncontrado.rol,

      dependencia:
        usuarioEncontrado.dependencia

    }

  };

}


/* =====================================================
   25. CERRAR SESIÓN DESDE EL CLIENTE
   ===================================================== */

function cerrarSesionCliente(token) {

  return cerrarSesion(
    token
  );

}


/* =====================================================
   26. OBTENER DATOS DE LA SESIÓN
   ===================================================== */

function obtenerDatosSesion(token) {

  const sesion =
    validarSesion_(
      token
    );


  return {

    estado: true,

    usuario:
      sesion.usuario,

    rol:
      sesion.rol,

    dependencia:
      sesion.dependencia,

    creado:
      sesion.creado

  };

}


/* =====================================================
   27. AUDITORÍA
   ===================================================== */

function registrarAuditoria_(
  usuario,
  accion,
  cedula,
  detalle
) {

  try {

    const hoja =
      obtenerHoja_(
        HOJA_AUDITORIA
      );


    /*
     * Estructura:
     *
     * A FECHA
     * B USUARIO
     * C ACCION
     * D CC
     * E DETALLE
     * F ORIGEN
     */

    hoja.appendRow([

      new Date(),

      usuario || '',

      accion || '',

      cedula || '',

      detalle || '',

      'WEB APP'

    ]);


  } catch (e) {

    /*
     * La auditoría no debe impedir
     * una operación válida del sistema.
     */

    console.error(
      'Error registrando auditoría: ' +
      e.message
    );

  }

}


/* =====================================================
   28. CONSULTAR USUARIO ACTUAL
   ===================================================== */

function consultarUsuario(token) {

  const sesion =
    validarSesion_(
      token
    );


  return respuestaOK_(

    'Usuario consultado correctamente.',

    {

      usuario:
        sesion.usuario,

      rol:
        sesion.rol,

      dependencia:
        sesion.dependencia

    }

  );

}


/* =====================================================
   29. CREAR USUARIO
   ===================================================== */

function crearUsuario(
  token,
  usuario,
  clave,
  rol,
  dependencia
) {

  const sesion =
    validarSesion_(
      token
    );


  validarRol_(
    sesion,
    [
      'ADMINISTRADOR'
    ]
  );


  usuario =
    String(
      usuario || ''
    ).trim();


  clave =
    String(
      clave || ''
    );


  rol =
    String(
      rol || ''
    ).trim();


  dependencia =
    String(
      dependencia || ''
    ).trim();


  if (!usuario) {

    throw new Error(
      'Debe indicar el usuario.'
    );

  }


  if (!clave) {

    throw new Error(
      'Debe indicar la contraseña.'
    );

  }


  if (!rol) {

    throw new Error(
      'Debe indicar el rol.'
    );

  }


  if (!dependencia) {

    throw new Error(
      'Debe indicar la dependencia.'
    );

  }


  const rolNormalizado =
    normalizarTexto_(
      rol
    );


  if (
    [
      'ADMINISTRADOR',
      'OPERADOR'
    ].indexOf(
      rolNormalizado
    ) === -1
  ) {

    throw new Error(
      'El rol debe ser ADMINISTRADOR u OPERADOR.'
    );

  }


  const hoja =
    obtenerHoja_(
      HOJA_USUARIOS
    );


  const datos =
    hoja
      .getDataRange()
      .getValues();


  if (
    datos.length === 0
  ) {

    throw new Error(
      'La hoja USUARIOS no tiene encabezados.'
    );

  }


  const encabezados =
    datos[0].map(
      function(valor) {

        return normalizarTexto_(
          valor
        );

      }
    );


  const indiceUsuario =
    encabezados.indexOf(
      'USUARIO'
    );

  const indicePassword =
    encabezados.indexOf(
      'PASSWORD_HASH'
    );

  const indiceRol =
    encabezados.indexOf(
      'ROL'
    );

  const indiceDependencia =
    encabezados.indexOf(
      'DEPENDENCIA'
    );

  const indiceEstado =
    encabezados.indexOf(
      'ESTADO'
    );

  const indiceFechaCreacion =
    encabezados.indexOf(
      'FECHA_CREACION'
    );

  const indiceUltimoAcceso =
    encabezados.indexOf(
      'ULTIMO_ACCESO'
    );


  if (
    indiceUsuario === -1 ||
    indicePassword === -1 ||
    indiceRol === -1 ||
    indiceDependencia === -1 ||
    indiceEstado === -1
  ) {

    throw new Error(
      'La hoja USUARIOS no tiene la estructura requerida.'
    );

  }


  /*
   * Verificar que el usuario
   * no exista previamente.
   */

  const usuarioNormalizado =
    normalizarTexto_(
      usuario
    );


  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    const existente =
      normalizarTexto_(
        datos[i][indiceUsuario]
      );


    if (
      existente ===
      usuarioNormalizado
    ) {

      throw new Error(
        'El usuario ya existe.'
      );

    }

  }


  /*
   * Crear nueva fila respetando
   * exactamente las columnas
   * existentes en USUARIOS.
   */

  const nuevaFila =
    new Array(
      encabezados.length
    ).fill('');


  nuevaFila[
    indiceUsuario
  ] =
    usuario;


  nuevaFila[
    indicePassword
  ] =
    sha256_(
      clave
    );


  nuevaFila[
    indiceRol
  ] =
    rolNormalizado;


  nuevaFila[
    indiceDependencia
  ] =
    dependencia;


  nuevaFila[
    indiceEstado
  ] =
    'ACTIVO';


  if (
    indiceFechaCreacion !== -1
  ) {

    nuevaFila[
      indiceFechaCreacion
    ] =
      new Date();

  }


  if (
    indiceUltimoAcceso !== -1
  ) {

    nuevaFila[
      indiceUltimoAcceso
    ] =
      '';

  }


  hoja.appendRow(
    nuevaFila
  );


  registrarAuditoria_(

    sesion.usuario,

    'CREACIÓN DE USUARIO',

    '',

    'Usuario creado: ' +
    usuario +
    ' | Rol: ' +
    rolNormalizado +
    ' | Dependencia: ' +
    dependencia

  );


  return respuestaOK_(

    'Usuario creado correctamente.',

    {

      usuario:
        usuario,

      rol:
        rolNormalizado,

      dependencia:
        dependencia,

      estado:
        'ACTIVO'

    }

  );

}


/* =====================================================
   30. LISTAR USUARIOS
   ===================================================== */

function listarUsuarios(token) {

  const sesion =
    validarSesion_(
      token
    );


  validarRol_(
    sesion,
    [
      'ADMINISTRADOR'
    ]
  );


  const hoja =
    obtenerHoja_(
      HOJA_USUARIOS
    );


  const datos =
    hoja
      .getDataRange()
      .getValues();


  if (
    datos.length < 2
  ) {

    return respuestaOK_(

      'No existen usuarios registrados.',

      []

    );

  }


  const encabezados =
    datos[0].map(
      function(valor) {

        return normalizarTexto_(
          valor
        );

      }
    );


  const indiceUsuario =
    encabezados.indexOf(
      'USUARIO'
    );

  const indiceRol =
    encabezados.indexOf(
      'ROL'
    );

  const indiceDependencia =
    encabezados.indexOf(
      'DEPENDENCIA'
    );

  const indiceEstado =
    encabezados.indexOf(
      'ESTADO'
    );

  const indiceFechaCreacion =
    encabezados.indexOf(
      'FECHA_CREACION'
    );

  const indiceUltimoAcceso =
    encabezados.indexOf(
      'ULTIMO_ACCESO'
    );


  const resultado = [];


  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    const fila =
      datos[i];


    const usuario =
      {

        fila:
          i + 1,

        usuario:
          indiceUsuario !== -1
            ? fila[indiceUsuario]
            : '',

        rol:
          indiceRol !== -1
            ? fila[indiceRol]
            : '',

        dependencia:
          indiceDependencia !== -1
            ? fila[indiceDependencia]
            : '',

        estado:
          indiceEstado !== -1
            ? fila[indiceEstado]
            : '',

        fechaCreacion:
          indiceFechaCreacion !== -1
            ? formatearFechaHora_(
                fila[indiceFechaCreacion]
              )
            : '',

        ultimoAcceso:
          indiceUltimoAcceso !== -1
            ? formatearFechaHora_(
                fila[indiceUltimoAcceso]
              )
            : ''

      };


    resultado.push(
      usuario
    );

  }


  return respuestaOK_(

    'Usuarios consultados correctamente.',

    resultado

  );

}


/* =====================================================
   31. CAMBIAR ESTADO DE USUARIO
   ===================================================== */

function cambiarEstadoUsuario(
  token,
  usuario,
  nuevoEstado
) {

  const sesion =
    validarSesion_(
      token
    );


  validarRol_(
    sesion,
    [
      'ADMINISTRADOR'
    ]
  );


  usuario =
    String(
      usuario || ''
    ).trim();


  nuevoEstado =
    normalizarTexto_(
      nuevoEstado
    );


  if (!usuario) {

    throw new Error(
      'Debe indicar el usuario.'
    );

  }


  if (
    [
      'ACTIVO',
      'INACTIVO'
    ].indexOf(
      nuevoEstado
    ) === -1
  ) {

    throw new Error(
      'El estado debe ser ACTIVO o INACTIVO.'
    );

  }


  const hoja =
    obtenerHoja_(
      HOJA_USUARIOS
    );


  const datos =
    hoja
      .getDataRange()
      .getValues();


  const encabezados =
    datos[0].map(
      function(valor) {

        return normalizarTexto_(
          valor
        );

      }
    );


  const indiceUsuario =
    encabezados.indexOf(
      'USUARIO'
    );

  const indiceEstado =
    encabezados.indexOf(
      'ESTADO'
    );


  if (
    indiceUsuario === -1 ||
    indiceEstado === -1
  ) {

    throw new Error(
      'La hoja USUARIOS no tiene la estructura requerida.'
    );

  }


  const usuarioBuscado =
    normalizarTexto_(
      usuario
    );


  let filaEncontrada =
    -1;


  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    if (
      normalizarTexto_(
        datos[i][indiceUsuario]
      ) ===
      usuarioBuscado
    ) {

      filaEncontrada =
        i + 1;

      break;

    }

  }


  if (
    filaEncontrada === -1
  ) {

    throw new Error(
      'No se encontró el usuario indicado.'
    );

  }


  /*
   * Evitar que el administrador
   * se desactive accidentalmente
   * a sí mismo.
   */

  if (
    normalizarTexto_(
      sesion.usuario
    ) ===
    usuarioBuscado &&
    nuevoEstado ===
    'INACTIVO'
  ) {

    throw new Error(
      'No puede desactivar su propio usuario mientras tiene la sesión activa.'
    );

  }


  hoja
    .getRange(

      filaEncontrada,

      indiceEstado + 1

    )
    .setValue(
      nuevoEstado
    );


  registrarAuditoria_(

    sesion.usuario,

    'CAMBIO DE ESTADO DE USUARIO',

    '',

    'Usuario: ' +
    usuario +
    ' | Nuevo estado: ' +
    nuevoEstado

  );


  return respuestaOK_(

    'Estado del usuario actualizado correctamente.',

    {

      usuario:
        usuario,

      estado:
        nuevoEstado

    }

  );

}


/* =====================================================
   FIN PARTE 2
   ===================================================== */
   /*******************************************************
 * PARTE 3
 * LISTADO_BASE Y CONSULTA DE FUNCIONARIOS
 *******************************************************/


/**
 * Obtiene los encabezados de la hoja LISTADO_BASE.
 * Se leen dinámicamente para evitar depender de posiciones
 * fijas de columnas.
 */
function obtenerEncabezadosBase_() {

  const hoja = obtenerHoja_(HOJA_BASE);

  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_BASE + '".');
  }

  const ultimaColumna = hoja.getLastColumn();

  if (ultimaColumna < 1) {
    throw new Error(
      'La hoja "' + HOJA_BASE + '" no contiene columnas.'
    );
  }

  return hoja
    .getRange(1, 1, 1, ultimaColumna)
    .getValues()[0]
    .map(function(encabezado) {
      return texto_(encabezado);
    });
}


/**
 * Obtiene el índice de una columna por nombre.
 *
 * Retorna:
 *   0 = primera columna
 *   1 = segunda columna
 *   etc.
 *
 * Retorna -1 si no encuentra la columna.
 */
function obtenerIndiceColumna_(encabezados, nombre) {

  const buscado = normalizarTexto_(nombre);

  for (let i = 0; i < encabezados.length; i++) {

    if (normalizarTexto_(encabezados[i]) === buscado) {
      return i;
    }
  }

  return -1;
}


/**
 * Construye el objeto completo de un funcionario
 * a partir de una fila de LISTADO_BASE.
 *
 * Se conservan TODOS los campos existentes en la hoja.
 */
function construirFuncionario_(encabezados, fila, numeroFila) {

  const funcionario = {};

  // Información completa según encabezados
  encabezados.forEach(function(encabezado, indice) {

    if (!encabezado) {
      return;
    }

    funcionario[encabezado] = fila[indice];
  });


  // Índices de campos principales
  const iCedula = obtenerIndiceColumna_(
    encabezados,
    'CEDULA'
  );

  const iNivel = obtenerIndiceColumna_(
    encabezados,
    'NIV'
  );

  const iGrado = obtenerIndiceColumna_(
    encabezados,
    'GR'
  );

  const iNombre = obtenerIndiceColumna_(
    encabezados,
    'FUNCIONARIO'
  );

  const iDependencia = obtenerIndiceColumna_(
    encabezados,
    'DEPENDENCIA'
  );

  const iTurno = obtenerIndiceColumna_(
    encabezados,
    'TURNO'
  );

  const iPlaca = obtenerIndiceColumna_(
    encabezados,
    'Placa_Chip'
  );


  // Propiedades normalizadas que utilizará el sistema
  funcionario.fila = numeroFila;

  funcionario.cedula =
    iCedula >= 0
      ? normalizarCedula_(fila[iCedula])
      : '';

  funcionario.nivel =
    iNivel >= 0
      ? texto_(fila[iNivel])
      : '';

  funcionario.grado =
    iGrado >= 0
      ? texto_(fila[iGrado])
      : '';

  funcionario.funcionario =
    iNombre >= 0
      ? texto_(fila[iNombre])
      : '';

  funcionario.dependencia =
    iDependencia >= 0
      ? texto_(fila[iDependencia])
      : '';

  funcionario.turno =
    iTurno >= 0
      ? texto_(fila[iTurno])
      : '';

  funcionario.placaChip =
    iPlaca >= 0
      ? texto_(fila[iPlaca])
      : '';


  return funcionario;
}


/**
 * Determina si el usuario puede consultar
 * la información de un funcionario.
 *
 * ADMINISTRADOR:
 *   Puede consultar cualquier dependencia.
 *
 * OPERADOR:
 *   Solo puede consultar funcionarios de su
 *   misma dependencia.
 */
function usuarioPuedeConsultarFuncionario_(
  sesion,
  funcionario
) {

  if (!sesion || !funcionario) {
    return false;
  }

  const rol = normalizarTexto_(sesion.rol);

  // Administrador tiene acceso general
  if (rol === 'ADMINISTRADOR') {
    return true;
  }

  // Operador: validar dependencia
  if (rol === 'OPERADOR') {

    const dependenciaUsuario =
      normalizarTexto_(sesion.dependencia);

    const dependenciaFuncionario =
      normalizarTexto_(funcionario.dependencia);

    return (
      dependenciaUsuario !== '' &&
      dependenciaFuncionario !== '' &&
      dependenciaUsuario === dependenciaFuncionario
    );
  }

  return false;
}


/**
 * Busca funcionarios por:
 * - Cédula
 * - Nombre
 *
 * Permite coincidencias parciales por nombre.
 *
 * Máximo de resultados: 50.
 */
function buscarFuncionario(tokenSesion, valor) {

  const sesion = validarSesion_(tokenSesion);

  const criterio = texto_(valor).trim();

  if (!criterio) {

    return respuestaError_(
      'Debe ingresar una cédula o un nombre para realizar la búsqueda.'
    );
  }


  const hoja = obtenerHoja_(HOJA_BASE);

  if (!hoja) {

    return respuestaError_(
      'No existe la hoja "' + HOJA_BASE + '".'
    );
  }


  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaFila < 2) {

    return respuestaOK_(
      'No existen funcionarios registrados.',
      []
    );
  }


  const encabezados = obtenerEncabezadosBase_();

  const datos = hoja
    .getRange(
      2,
      1,
      ultimaFila - 1,
      ultimaColumna
    )
    .getValues();


  const criterioCedula =
    normalizarCedula_(criterio);

  const criterioNombre =
    normalizarNombre_(criterio);


  const resultados = [];


  datos.forEach(function(fila, indice) {

    if (resultados.length >= 50) {
      return;
    }


    const funcionario =
      construirFuncionario_(
        encabezados,
        fila,
        indice + 2
      );


    // Validación de permisos
    if (
      !usuarioPuedeConsultarFuncionario_(
        sesion,
        funcionario
      )
    ) {
      return;
    }


    const cedulaFuncionario =
      normalizarCedula_(
        funcionario.cedula
      );


    const nombreFuncionario =
      normalizarNombre_(
        funcionario.funcionario
      );


    let coincide = false;


    // Búsqueda por cédula
    if (
      criterioCedula !== '' &&
      cedulaFuncionario.indexOf(
        criterioCedula
      ) !== -1
    ) {

      coincide = true;
    }


    // Búsqueda por nombre
    if (
      criterioNombre !== '' &&
      nombreFuncionario.indexOf(
        criterioNombre
      ) !== -1
    ) {

      coincide = true;
    }


    if (coincide) {
      resultados.push(funcionario);
    }
  });


  // Registrar consulta
  registrarAuditoria_(
    sesion.usuario,
    'BUSQUEDA_FUNCIONARIO',
    criterioCedula,
    'Criterio: ' + criterio +
    ' | Resultados: ' + resultados.length
  );


  return respuestaOK_(
    resultados.length > 0
      ? 'Consulta realizada correctamente.'
      : 'No se encontraron funcionarios con el criterio indicado.',
    resultados
  );
}


/**
 * Consulta exacta de un funcionario por cédula.
 */
function consultarFuncionario(token, cedula) {

  const sesion = validarSesion_(token);

  const cedulaBuscada =
    normalizarCedula_(cedula);


  if (!cedulaBuscada) {

    return respuestaError_(
      'Debe indicar una cédula válida.'
    );
  }


  const hoja = obtenerHoja_(HOJA_BASE);

  if (!hoja) {

    return respuestaError_(
      'No existe la hoja "' + HOJA_BASE + '".'
    );
  }


  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaFila < 2) {

    return respuestaOK_(
      'No existen funcionarios registrados.',
      null
    );
  }


  const encabezados =
    obtenerEncabezadosBase_();


  const indiceCedula =
    obtenerIndiceColumna_(
      encabezados,
      'CEDULA'
    );


  if (indiceCedula < 0) {

    return respuestaError_(
      'No se encontró la columna CEDULA en LISTADO_BASE.'
    );
  }


  const datos = hoja
    .getRange(
      2,
      1,
      ultimaFila - 1,
      ultimaColumna
    )
    .getValues();


  let encontrado = null;


  for (
    let i = 0;
    i < datos.length;
    i++
  ) {

    const cedulaFila =
      normalizarCedula_(
        datos[i][indiceCedula]
      );


    if (
      cedulaFila === cedulaBuscada
    ) {

      const funcionario =
        construirFuncionario_(
          encabezados,
          datos[i],
          i + 2
        );


      if (
        !usuarioPuedeConsultarFuncionario_(
          sesion,
          funcionario
        )
      ) {

        return respuestaError_(
          'El usuario no tiene permisos para consultar este funcionario.'
        );
      }


      encontrado = funcionario;
      break;
    }
  }


  // Registrar consulta
  registrarAuditoria_(
    sesion.usuario,
    'CONSULTA_FUNCIONARIO',
    cedulaBuscada,
    encontrado
      ? 'Funcionario encontrado.'
      : 'Funcionario no encontrado.'
  );


  if (!encontrado) {

    return respuestaOK_(
      'No se encontró un funcionario con la cédula indicada.',
      null
    );
  }


  return respuestaOK_(
    'Funcionario encontrado correctamente.',
    encontrado
  );
}


/**
 * Obtiene la ficha completa de un funcionario.
 *
 * Esta función queda como punto de acceso para el frontend
 * cuando se requiera cargar toda la información del funcionario.
 */
function obtenerFichaFuncionario(token, cedula) {

  const resultado =
    consultarFuncionario(
      token,
      cedula
    );


  if (
    !resultado ||
    resultado.estado !== true
  ) {

    return resultado;
  }


  const funcionario =
    resultado.datos;


  return respuestaOK_(
    'Ficha del funcionario cargada correctamente.',
    {
      funcionario: funcionario,
      fila: funcionario
        ? funcionario.fila
        : null
    }
  );
}


/**
 * Diagnóstico de LISTADO_BASE.
 *
 * Solo disponible para ADMINISTRADOR.
 *
 * Permite comprobar:
 * - cantidad de filas
 * - cantidad de columnas
 * - encabezados
 * - registros con cédula
 * - registros sin cédula
 * - distribución de dependencias
 * - distribución de turnos
 */
function diagnosticarListadoBase(token) {

  const sesion = validarSesion_(token);

  validarRol_(
    sesion,
    ['ADMINISTRADOR']
  );


  const hoja =
    obtenerHoja_(HOJA_BASE);


  if (!hoja) {

    return respuestaError_(
      'No existe la hoja "' + HOJA_BASE + '".'
    );
  }


  const ultimaFila =
    hoja.getLastRow();

  const ultimaColumna =
    hoja.getLastColumn();


  if (ultimaFila < 1) {

    return respuestaOK_(
      'La hoja LISTADO_BASE está vacía.',
      {
        filas: 0,
        columnas: ultimaColumna,
        encabezados: [],
        registrosConCedula: 0,
        registrosSinCedula: 0,
        dependencias: {},
        turnos: {}
      }
    );
  }


  const encabezados =
    obtenerEncabezadosBase_();


  if (ultimaFila < 2) {

    return respuestaOK_(
      'LISTADO_BASE contiene únicamente encabezados.',
      {
        filas: 0,
        columnas: ultimaColumna,
        encabezados: encabezados,
        registrosConCedula: 0,
        registrosSinCedula: 0,
        dependencias: {},
        turnos: {}
      }
    );
  }


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  const indiceCedula =
    obtenerIndiceColumna_(
      encabezados,
      'CEDULA'
    );


  const indiceDependencia =
    obtenerIndiceColumna_(
      encabezados,
      'DEPENDENCIA'
    );


  const indiceTurno =
    obtenerIndiceColumna_(
      encabezados,
      'TURNO'
    );


  let registrosConCedula = 0;
  let registrosSinCedula = 0;


  const dependencias = {};
  const turnos = {};


  datos.forEach(function(fila) {

    // Cédula
    if (
      indiceCedula >= 0 &&
      normalizarCedula_(
        fila[indiceCedula]
      ) !== ''
    ) {

      registrosConCedula++;

    } else {

      registrosSinCedula++;
    }


    // Dependencia
    if (indiceDependencia >= 0) {

      const dependencia =
        texto_(
          fila[indiceDependencia]
        ) || '(SIN DEPENDENCIA)';


      dependencias[dependencia] =
        (dependencias[dependencia] || 0) + 1;
    }


    // Turno
    if (indiceTurno >= 0) {

      const turno =
        texto_(
          fila[indiceTurno]
        ) || '(SIN TURNO)';


      turnos[turno] =
        (turnos[turno] || 0) + 1;
    }
  });


  registrarAuditoria_(
    sesion.usuario,
    'DIAGNOSTICO_LISTADO_BASE',
    '',
    'Filas: ' + (ultimaFila - 1) +
    ' | Columnas: ' + ultimaColumna
  );


  return respuestaOK_(
    'Diagnóstico de LISTADO_BASE realizado correctamente.',
    {
      filas: ultimaFila - 1,
      columnas: ultimaColumna,
      encabezados: encabezados,
      registrosConCedula: registrosConCedula,
      registrosSinCedula: registrosSinCedula,
      dependencias: dependencias,
      turnos: turnos
    }
  );
}


/*******************************************************
 * FIN PARTE 3
 *******************************************************/

 /*******************************************************
 * PARTE 4
 * GESTIÓN DE NOVEDADES
 *
 * HOJA: NOVEDADES
 *
 * COLUMNAS OBLIGATORIAS:
 * A  = GR
 * B  = APELLIDOS Y NOMBRES
 * C  = CC
 * D  = DEPENDENCIA
 * E  = NOVEDAD
 * F  = DESCRIPCION
 * G  = Dias
 * H  = Fecha INICIAL
 * I  = Fecha PRESENTACION
 * J  = Observacion
 * K  = RV
 * L  = NIV
 * M  = Turno
 * N  = Placa_Chip
 * O  = TEXTO
 *******************************************************/


/**
 * Obtiene los encabezados de la hoja NOVEDADES.
 */
function obtenerEncabezadosNovedades_() {

  const hoja = obtenerHoja_(HOJA_NOVEDADES);

  if (!hoja) {
    throw new Error(
      'No existe la hoja "' + HOJA_NOVEDADES + '".'
    );
  }

  const ultimaColumna = hoja.getLastColumn();

  if (ultimaColumna < 1) {
    throw new Error(
      'La hoja "' + HOJA_NOVEDADES + '" no contiene encabezados.'
    );
  }

  return hoja
    .getRange(1, 1, 1, ultimaColumna)
    .getValues()[0]
    .map(function(encabezado) {
      return texto_(encabezado);
    });
}


/**
 * Verifica que NOVEDADES conserve exactamente
 * los 15 encabezados establecidos por el sistema.
 */
function validarEstructuraNovedades_() {

  const hoja = obtenerHoja_(HOJA_NOVEDADES);

  if (!hoja) {
    throw new Error(
      'No existe la hoja "' + HOJA_NOVEDADES + '".'
    );
  }

  const encabezados =
    obtenerEncabezadosNovedades_();


  if (
    encabezados.length !==
    ENCABEZADOS_NOVEDADES.length
  ) {

    throw new Error(
      'La hoja NOVEDADES debe contener exactamente ' +
      ENCABEZADOS_NOVEDADES.length +
      ' columnas. Actualmente contiene ' +
      encabezados.length +
      '.'
    );
  }


  for (
    let i = 0;
    i < ENCABEZADOS_NOVEDADES.length;
    i++
  ) {

    if (
      normalizarTexto_(encabezados[i]) !==
      normalizarTexto_(
        ENCABEZADOS_NOVEDADES[i]
      )
    ) {

      throw new Error(
        'Estructura incorrecta en NOVEDADES. ' +
        'La columna ' +
        String.fromCharCode(65 + i) +
        ' debe ser "' +
        ENCABEZADOS_NOVEDADES[i] +
        '" y actualmente es "' +
        encabezados[i] +
        '".'
      );
    }
  }


  return true;
}


/**
 * Obtiene el índice de una columna de NOVEDADES.
 */
function obtenerIndiceColumnaNovedades_(
  encabezados,
  nombre
) {

  const buscado =
    normalizarTexto_(nombre);


  for (
    let i = 0;
    i < encabezados.length;
    i++
  ) {

    if (
      normalizarTexto_(encabezados[i]) ===
      buscado
    ) {

      return i;
    }
  }


  return -1;
}


/**
 * Convierte una fecha recibida desde el frontend
 * a un objeto Date válido.
 *
 * Acepta:
 * - Date
 * - yyyy-MM-dd
 * - dd/MM/yyyy
 * - dd-MM-yyyy
 */
function convertirFechaNovedad_(valor) {

  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {
    return '';
  }


  if (
    Object.prototype.toString.call(valor) ===
    '[object Date]'
  ) {

    if (!isNaN(valor.getTime())) {
      return valor;
    }

    return '';
  }


  const textoFecha =
    String(valor).trim();


  if (!textoFecha) {
    return '';
  }


  // Formato yyyy-MM-dd
  let coincidencia =
    textoFecha.match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );


  if (coincidencia) {

    const fecha =
      new Date(
        Number(coincidencia[1]),
        Number(coincidencia[2]) - 1,
        Number(coincidencia[3])
      );


    if (!isNaN(fecha.getTime())) {
      return fecha;
    }
  }


  // Formato dd/MM/yyyy
  coincidencia =
    textoFecha.match(
      /^(\d{2})\/(\d{2})\/(\d{4})$/
    );


  if (coincidencia) {

    const fecha =
      new Date(
        Number(coincidencia[3]),
        Number(coincidencia[2]) - 1,
        Number(coincidencia[1])
      );


    if (!isNaN(fecha.getTime())) {
      return fecha;
    }
  }


  // Formato dd-MM-yyyy
  coincidencia =
    textoFecha.match(
      /^(\d{2})-(\d{2})-(\d{4})$/
    );


  if (coincidencia) {

    const fecha =
      new Date(
        Number(coincidencia[3]),
        Number(coincidencia[2]) - 1,
        Number(coincidencia[1])
      );


    if (!isNaN(fecha.getTime())) {
      return fecha;
    }
  }


  // Último intento
  const fechaGenerica =
    new Date(textoFecha);


  if (
    !isNaN(
      fechaGenerica.getTime()
    )
  ) {

    return fechaGenerica;
  }


  return '';
}


/**
 * Normaliza un valor para guardar en NOVEDADES.
 */
function normalizarValorNovedad_(valor) {

  if (
    valor === null ||
    valor === undefined
  ) {
    return '';
  }

  if (
    Object.prototype.toString.call(valor) ===
    '[object Date]'
  ) {
    return valor;
  }

  return texto_(valor);
}


/**
 * Construye el texto consolidado de una novedad.
 *
 * Este campo se utiliza para facilitar búsquedas,
 * filtros, impresión y trazabilidad.
 */
function construirTextoNovedad_(datos) {

  const partes = [];


  if (datos.gr) {
    partes.push(
      texto_(datos.gr)
    );
  }


  if (datos.funcionario) {
    partes.push(
      texto_(datos.funcionario)
    );
  }


  if (datos.cc) {
    partes.push(
      texto_(datos.cc)
    );
  }


  if (datos.dependencia) {
    partes.push(
      texto_(datos.dependencia)
    );
  }


  if (datos.novedad) {
    partes.push(
      texto_(datos.novedad)
    );
  }


  if (datos.descripcion) {
    partes.push(
      texto_(datos.descripcion)
    );
  }


  if (datos.observacion) {
    partes.push(
      texto_(datos.observacion)
    );
  }


  if (datos.turno) {
    partes.push(
      texto_(datos.turno)
    );
  }


  return partes.join(' | ');
}


/**
 * Busca un funcionario en LISTADO_BASE y devuelve
 * su ficha completa.
 */
function obtenerFuncionarioParaNovedad_(
  cedula
) {

  const hoja =
    obtenerHoja_(HOJA_BASE);


  if (!hoja) {
    throw new Error(
      'No existe la hoja "' + HOJA_BASE + '".'
    );
  }


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2 ||
    ultimaColumna < 1
  ) {

    return null;
  }


  const encabezados =
    obtenerEncabezadosBase_();


  const indiceCedula =
    obtenerIndiceColumna_(
      encabezados,
      'CEDULA'
    );


  if (indiceCedula < 0) {

    throw new Error(
      'No se encontró la columna CEDULA en LISTADO_BASE.'
    );
  }


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  const cedulaBuscada =
    normalizarCedula_(cedula);


  for (
    let i = 0;
    i < datos.length;
    i++
  ) {

    const cedulaFila =
      normalizarCedula_(
        datos[i][indiceCedula]
      );


    if (
      cedulaFila ===
      cedulaBuscada
    ) {

      return construirFuncionario_(
        encabezados,
        datos[i],
        i + 2
      );
    }
  }


  return null;
}


/**
 * Obtiene las novedades existentes de un funcionario.
 *
 * Se utiliza para determinar su historial completo.
 */
function obtenerHistorialNovedadesInterno_(
  cedula
) {

  validarEstructuraNovedades_();


  const hoja =
    obtenerHoja_(HOJA_NOVEDADES);


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2 ||
    ultimaColumna < 1
  ) {

    return [];
  }


  const encabezados =
    obtenerEncabezadosNovedades_();


  const indiceCC =
    obtenerIndiceColumnaNovedades_(
      encabezados,
      'CC'
    );


  if (indiceCC < 0) {

    throw new Error(
      'No se encontró la columna CC en NOVEDADES.'
    );
  }


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  const cedulaBuscada =
    normalizarCedula_(cedula);


  const resultados = [];


  datos.forEach(function(fila, indice) {

    const cedulaFila =
      normalizarCedula_(
        fila[indiceCC]
      );


    if (
      cedulaFila !==
      cedulaBuscada
    ) {

      return;
    }


    const registro = {};


    encabezados.forEach(
      function(encabezado, columna) {

        registro[encabezado] =
          fila[columna];
      }
    );


    registro.fila =
      indice + 2;


    resultados.push(
      registro
    );
  });


  // Ordenar del registro más reciente
  // al más antiguo.
  resultados.sort(
    function(a, b) {

      const fechaA =
        a['Fecha INICIAL'] instanceof Date
          ? a['Fecha INICIAL'].getTime()
          : 0;

      const fechaB =
        b['Fecha INICIAL'] instanceof Date
          ? b['Fecha INICIAL'].getTime()
          : 0;


      return fechaB - fechaA;
    }
  );


  return resultados;
}


/**
 * Determina si una novedad continúa activa.
 *
 * Criterio:
 *
 * 1. Si no tiene Fecha PRESENTACION,
 *    se considera activa.
 *
 * 2. Si tiene Fecha PRESENTACION igual o
 *    posterior al día actual, continúa activa.
 *
 * 3. Si la Fecha PRESENTACION es anterior
 *    al día actual, se considera finalizada.
 */
function novedadEstaActiva_(fechaPresentacion) {

  if (
    fechaPresentacion === null ||
    fechaPresentacion === undefined ||
    fechaPresentacion === ''
  ) {

    return true;
  }


  let fecha;


  if (
    Object.prototype.toString.call(
      fechaPresentacion
    ) === '[object Date]'
  ) {

    fecha =
      new Date(
        fechaPresentacion.getTime()
      );

  } else {

    fecha =
      convertirFechaNovedad_(
        fechaPresentacion
      );
  }


  if (!fecha) {
    return true;
  }


  const hoy =
    new Date();


  hoy.setHours(
    0,
    0,
    0,
    0
  );


  fecha.setHours(
    0,
    0,
    0,
    0
  );


  return fecha.getTime() >= hoy.getTime();
}


/**
 * Obtiene únicamente las novedades activas
 * de un funcionario.
 */
function obtenerNovedadesActivasCC_(
  cedula
) {

  const historial =
    obtenerHistorialNovedadesInterno_(
      cedula
    );


  return historial.filter(
    function(registro) {

      return novedadEstaActiva_(
        registro['Fecha PRESENTACION']
      );
    }
  );
}


/**
 * Registra una nueva novedad.
 *
 * Parámetros esperados desde el frontend:
 *
 * {
 *   cc: '',
 *   novedad: '',
 *   descripcion: '',
 *   dias: '',
 *   fechaInicial: '',
 *   fechaPresentacion: '',
 *   observacion: '',
 *   rv: ''
 * }
 */
function registrarNovedad(
  token,
  datos
) {

  const sesion =
    validarSesion_(token);


  validarEstructuraNovedades_();


  if (
    !datos ||
    typeof datos !== 'object'
  ) {

    return respuestaError_(
      'No se recibieron los datos de la novedad.'
    );
  }


  const cc =
    normalizarCedula_(
      datos.cc ||
      datos.CC ||
      datos.cedula
    );


  if (!cc) {

    return respuestaError_(
      'Debe indicar la cédula del funcionario.'
    );
  }


  const novedad =
    texto_(
      datos.novedad ||
      datos.NOVEDAD
    ).trim();


  if (!novedad) {

    return respuestaError_(
      'Debe seleccionar o registrar el tipo de novedad.'
    );
  }


  /*
   * Primero se verifica que el funcionario
   * exista realmente en LISTADO_BASE.
   */
  const funcionario =
    obtenerFuncionarioParaNovedad_(
      cc
    );


  if (!funcionario) {

    return respuestaError_(
      'La cédula ' +
      cc +
      ' no se encuentra registrada en LISTADO_BASE.'
    );
  }


  /*
   * Verificación de permisos.
   */
  if (
    !usuarioPuedeConsultarFuncionario_(
      sesion,
      funcionario
    )
  ) {

    return respuestaError_(
      'El usuario no tiene permisos para registrar novedades de este funcionario.'
    );
  }


  /*
   * Datos enviados por el frontend.
   */
  const descripcion =
    texto_(
      datos.descripcion ||
      datos.DESCRIPCION
    ).trim();


  const observacion =
    texto_(
      datos.observacion ||
      datos.Observacion ||
      datos.OBSERVACION
    ).trim();


  const rv =
    texto_(
      datos.rv ||
      datos.RV
    ).trim();


  const dias =
    texto_(
      datos.dias ||
      datos.Dias
    ).trim();


  /*
   * Fechas.
   */
  const fechaInicial =
    convertirFechaNovedad_(
      datos.fechaInicial ||
      datos['Fecha INICIAL'] ||
      datos.fecha_inicial
    );


  const fechaPresentacion =
    convertirFechaNovedad_(
      datos.fechaPresentacion ||
      datos['Fecha PRESENTACION'] ||
      datos.fecha_presentacion
    );


  /*
   * Validación básica de fechas.
   */
  if (
    (
      datos.fechaInicial ||
      datos['Fecha INICIAL'] ||
      datos.fecha_inicial
    ) &&
    !fechaInicial
  ) {

    return respuestaError_(
      'La Fecha INICIAL no tiene un formato válido.'
    );
  }


  if (
    (
      datos.fechaPresentacion ||
      datos['Fecha PRESENTACION'] ||
      datos.fecha_presentacion
    ) &&
    !fechaPresentacion
  ) {

    return respuestaError_(
      'La Fecha PRESENTACION no tiene un formato válido.'
    );
  }


  /*
   * Construcción del registro.
   *
   * IMPORTANTE:
   * Los datos de identidad del funcionario NO se toman
   * directamente del navegador.
   *
   * Se obtienen nuevamente de LISTADO_BASE.
   */
  const registro = {

    gr:
      funcionario.grado,

    funcionario:
      funcionario.funcionario,

    cc:
      funcionario.cedula,

    dependencia:
      funcionario.dependencia,

    novedad:
      novedad,

    descripcion:
      descripcion,

    dias:
      dias,

    fechaInicial:
      fechaInicial,

    fechaPresentacion:
      fechaPresentacion,

    observacion:
      observacion,

    rv:
      rv,

    nivel:
      funcionario.nivel,

    turno:
      funcionario.turno,

    placaChip:
      funcionario.placaChip
  };


  registro.texto =
    construirTextoNovedad_(
      registro
    );


  /*
   * Obtener hoja.
   */
  const hoja =
    obtenerHoja_(HOJA_NOVEDADES);


  if (!hoja) {

    return respuestaError_(
      'No existe la hoja "' +
      HOJA_NOVEDADES +
      '".'
    );
  }


  /*
   * Se agrega exactamente una fila
   * con las 15 columnas establecidas.
   */
  const nuevaFila = [

    registro.gr,                    // A GR
    registro.funcionario,           // B APELLIDOS Y NOMBRES
    registro.cc,                    // C CC
    registro.dependencia,           // D DEPENDENCIA
    registro.novedad,               // E NOVEDAD
    registro.descripcion,           // F DESCRIPCION
    registro.dias,                  // G Dias
    registro.fechaInicial,          // H Fecha INICIAL
    registro.fechaPresentacion,     // I Fecha PRESENTACION
    registro.observacion,           // J Observacion
    registro.rv,                    // K RV
    registro.nivel,                 // L NIV
    registro.turno,                 // M Turno
    registro.placaChip,             // N Placa_Chip
    registro.texto                  // O TEXTO
  ];


  hoja.appendRow(
    nuevaFila
  );


  const filaRegistrada =
    hoja.getLastRow();


  /*
   * Formato de fechas.
   */
  if (fechaInicial) {

    hoja
      .getRange(
        filaRegistrada,
        8
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );
  }


  if (fechaPresentacion) {

    hoja
      .getRange(
        filaRegistrada,
        9
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );
  }


  /*
   * Auditoría.
   */
  registrarAuditoria_(
    sesion.usuario,
    'REGISTRAR_NOVEDAD',
    cc,
    'Novedad: ' +
    novedad +
    ' | Funcionario: ' +
    funcionario.funcionario +
    ' | Fila: ' +
    filaRegistrada
  );


  return respuestaOK_(
    'Novedad registrada correctamente.',
    {
      fila: filaRegistrada,
      registro: nuevaFila,
      funcionario: funcionario
    }
  );
}


/**
 * Obtiene el historial completo de un funcionario.
 */
function obtenerHistorialFuncionario(
  token,
  cedula
) {

  const sesion =
    validarSesion_(token);


  const cc =
    normalizarCedula_(cedula);


  if (!cc) {

    return respuestaError_(
      'Debe indicar una cédula válida.'
    );
  }


  const funcionario =
    obtenerFuncionarioParaNovedad_(
      cc
    );


  if (!funcionario) {

    return respuestaError_(
      'El funcionario no se encuentra registrado en LISTADO_BASE.'
    );
  }


  if (
    !usuarioPuedeConsultarFuncionario_(
      sesion,
      funcionario
    )
  ) {

    return respuestaError_(
      'El usuario no tiene permisos para consultar este funcionario.'
    );
  }


  const historial =
    obtenerHistorialNovedadesInterno_(
      cc
    );


  registrarAuditoria_(
    sesion.usuario,
    'CONSULTAR_HISTORIAL_NOVEDADES',
    cc,
    'Registros encontrados: ' +
    historial.length
  );


  return respuestaOK_(
    'Historial consultado correctamente.',
    {
      funcionario: funcionario,
      historial: historial,
      total: historial.length
    }
  );
}


/**
 * Obtiene únicamente las novedades activas
 * de un funcionario.
 */
function obtenerNovedadesActivas(
  token,
  cedula
) {

  const sesion =
    validarSesion_(token);


  const cc =
    normalizarCedula_(cedula);


  if (!cc) {

    return respuestaError_(
      'Debe indicar una cédula válida.'
    );
  }


  const funcionario =
    obtenerFuncionarioParaNovedad_(
      cc
    );


  if (!funcionario) {

    return respuestaError_(
      'El funcionario no se encuentra registrado en LISTADO_BASE.'
    );
  }


  if (
    !usuarioPuedeConsultarFuncionario_(
      sesion,
      funcionario
    )
  ) {

    return respuestaError_(
      'El usuario no tiene permisos para consultar este funcionario.'
    );
  }


  const activas =
    obtenerNovedadesActivasCC_(
      cc
    );


  return respuestaOK_(
    'Novedades activas consultadas correctamente.',
    {
      funcionario: funcionario,
      novedades: activas,
      total: activas.length
    }
  );
}


/**
 * Consulta el historial completo utilizando
 * el objeto funcionario cuando ya se encuentra
 * cargado en pantalla.
 */
function consultarHistorialFuncionario(
  token,
  cedula
) {

  return obtenerHistorialFuncionario(
    token,
    cedula
  );
}


/**
 * Consulta todas las novedades registradas.
 *
 * ADMINISTRADOR:
 *   Puede consultar toda la información.
 *
 * OPERADOR:
 *   Solo obtiene novedades de su dependencia.
 */
function listarNovedades(
  token,
  limite
) {

  const sesion =
    validarSesion_(token);


  validarEstructuraNovedades_();


  const hoja =
    obtenerHoja_(HOJA_NOVEDADES);


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2
  ) {

    return respuestaOK_(
      'No existen novedades registradas.',
      {
        novedades: [],
        total: 0
      }
    );
  }


  const encabezados =
    obtenerEncabezadosNovedades_();


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  let maximo =
    Number(limite);


  if (
    !isFinite(maximo) ||
    maximo <= 0
  ) {

    maximo = 500;
  }


  maximo =
    Math.min(
      maximo,
      2000
    );


  const resultados = [];


  datos.forEach(
    function(fila, indice) {

      if (
        resultados.length >= maximo
      ) {
        return;
      }


      const registro = {};


      encabezados.forEach(
        function(encabezado, columna) {

          registro[encabezado] =
            fila[columna];
        }
      );


      registro.fila =
        indice + 2;


      /*
       * Los operadores solo pueden ver
       * novedades de su dependencia.
       */
      if (
        normalizarTexto_(sesion.rol) ===
        'OPERADOR'
      ) {

        const dependencia =
          normalizarTexto_(
            registro['DEPENDENCIA']
          );


        const dependenciaUsuario =
          normalizarTexto_(
            sesion.dependencia
          );


        if (
          dependencia !==
          dependenciaUsuario
        ) {

          return;
        }
      }


      resultados.push(
        registro
      );
    }
  );


  /*
   * Más recientes primero.
   */
  resultados.sort(
    function(a, b) {

      const fechaA =
        a['Fecha INICIAL'] instanceof Date
          ? a['Fecha INICIAL'].getTime()
          : 0;


      const fechaB =
        b['Fecha INICIAL'] instanceof Date
          ? b['Fecha INICIAL'].getTime()
          : 0;


      return fechaB - fechaA;
    }
  );


  registrarAuditoria_(
    sesion.usuario,
    'LISTAR_NOVEDADES',
    '',
    'Registros retornados: ' +
    resultados.length
  );


  return respuestaOK_(
    'Novedades consultadas correctamente.',
    {
      novedades: resultados,
      total: resultados.length
    }
  );
}


/**
 * Busca novedades por cédula, nombre,
 * dependencia o texto consolidado.
 */
function buscarNovedades(
  token,
  criterio
) {

  const sesion =
    validarSesion_(token);


  validarEstructuraNovedades_();


  const textoBuscado =
    normalizarTexto_(
      criterio
    );


  if (!textoBuscado) {

    return respuestaError_(
      'Debe ingresar un criterio de búsqueda.'
    );
  }


  const hoja =
    obtenerHoja_(HOJA_NOVEDADES);


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2
  ) {

    return respuestaOK_(
      'No existen novedades registradas.',
      []
    );
  }


  const encabezados =
    obtenerEncabezadosNovedades_();


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  const resultados = [];


  datos.forEach(
    function(fila, indice) {

      const registro = {};


      encabezados.forEach(
        function(encabezado, columna) {

          registro[encabezado] =
            fila[columna];
        }
      );


      registro.fila =
        indice + 2;


      /*
       * Restricción por dependencia
       * para operadores.
       */
      if (
        normalizarTexto_(sesion.rol) ===
        'OPERADOR'
      ) {

        const dependencia =
          normalizarTexto_(
            registro['DEPENDENCIA']
          );


        const dependenciaUsuario =
          normalizarTexto_(
            sesion.dependencia
          );


        if (
          dependencia !==
          dependenciaUsuario
        ) {

          return;
        }
      }


      /*
       * Se busca en todos los campos
       * de la fila.
       */
      const textoFila =
        encabezados
          .map(
            function(encabezado) {

              return texto_(
                registro[encabezado]
              );
            }
          )
          .join(' ');


      if (
        normalizarTexto_(
          textoFila
        ).indexOf(
          textoBuscado
        ) !== -1
      ) {

        resultados.push(
          registro
        );
      }
    }
  );


  /*
   * Limitar resultados para evitar
   * respuestas excesivamente grandes.
   */
  const resultadosFinales =
    resultados.slice(
      0,
      500
    );


  registrarAuditoria_(
    sesion.usuario,
    'BUSCAR_NOVEDADES',
    '',
    'Criterio: ' +
    texto_(criterio) +
    ' | Resultados: ' +
    resultadosFinales.length
  );


  return respuestaOK_(
    resultadosFinales.length > 0
      ? 'Consulta realizada correctamente.'
      : 'No se encontraron novedades con el criterio indicado.',
    resultadosFinales
  );
}


/**
 * Consulta resumida de novedades activas
 * de todos los funcionarios que pertenecen
 * a la dependencia del usuario.
 *
 * Es útil para el tablero.
 */
function listarNovedadesActivas(
  token
) {

  const sesion =
    validarSesion_(token);


  validarEstructuraNovedades_();


  const hoja =
    obtenerHoja_(HOJA_NOVEDADES);


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2
  ) {

    return respuestaOK_(
      'No existen novedades registradas.',
      {
        novedades: [],
        total: 0
      }
    );
  }


  const encabezados =
    obtenerEncabezadosNovedades_();


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  const resultados = [];


  datos.forEach(
    function(fila, indice) {

      const registro = {};


      encabezados.forEach(
        function(encabezado, columna) {

          registro[encabezado] =
            fila[columna];
        }
      );


      registro.fila =
        indice + 2;


      if (
        !novedadEstaActiva_(
          registro['Fecha PRESENTACION']
        )
      ) {

        return;
      }


      /*
       * Operador: únicamente su dependencia.
       */
      if (
        normalizarTexto_(sesion.rol) ===
        'OPERADOR'
      ) {

        if (
          normalizarTexto_(
            registro['DEPENDENCIA']
          ) !==
          normalizarTexto_(
            sesion.dependencia
          )
        ) {

          return;
        }
      }


      resultados.push(
        registro
      );
    }
  );


  /*
   * Más recientes primero.
   */
  resultados.sort(
    function(a, b) {

      const fechaA =
        a['Fecha INICIAL'] instanceof Date
          ? a['Fecha INICIAL'].getTime()
          : 0;


      const fechaB =
        b['Fecha INICIAL'] instanceof Date
          ? b['Fecha INICIAL'].getTime()
          : 0;


      return fechaB - fechaA;
    }
  );


  registrarAuditoria_(
    sesion.usuario,
    'LISTAR_NOVEDADES_ACTIVAS',
    '',
    'Total activas: ' +
    resultados.length
  );


  return respuestaOK_(
    'Novedades activas consultadas correctamente.',
    {
      novedades: resultados,
      total: resultados.length
    }
  );
}


/*******************************************************
 * FIN PARTE 4
 *******************************************************/

 /*******************************************************
 * PARTE 5
 * CONSULTA COMPLETA POR TURNO
 *
 * FILTROS DISPONIBLES:
 *
 * A
 * B
 * C
 * SEPRI
 * GURIN
 * NO APLICA
 *
 * REGLAS:
 *
 * A         -> TURNO exactamente A
 * B         -> TURNO exactamente B
 * C         -> TURNO exactamente C
 * SEPRI     -> TURNO contiene SEPRI
 * GURIN     -> TURNO contiene GURIN
 * NO APLICA -> TURNO contiene NO APLICA
 *              O TURNO vacío
 *******************************************************/


/**
 * Normaliza el filtro recibido desde el frontend.
 */
function normalizarFiltroTurno_(filtro) {

  let valor =
    texto_(filtro).trim();


  valor =
    normalizarTexto_(valor);


  /*
   * Convierte variantes habituales.
   */
  if (
    valor === 'NO APLICA' ||
    valor === 'NO_APLICA' ||
    valor === 'NO-APLICA'
  ) {

    return 'NO APLICA';
  }


  if (valor === 'SEPRI') {
    return 'SEPRI';
  }


  if (valor === 'GURIN') {
    return 'GURIN';
  }


  if (
    valor === 'A' ||
    valor === 'B' ||
    valor === 'C'
  ) {

    return valor;
  }


  return '';
}


/**
 * Determina si un valor de TURNO pertenece
 * al filtro solicitado.
 */
function turnoCoincideFiltro_(
  turno,
  filtro
) {

  const valorTurno =
    normalizarTexto_(
      turno
    ).trim();


  const filtroNormalizado =
    normalizarFiltroTurno_(
      filtro
    );


  if (!filtroNormalizado) {
    return false;
  }


  /*
   * A, B y C:
   * coincidencia EXACTA.
   */
  if (
    filtroNormalizado === 'A' ||
    filtroNormalizado === 'B' ||
    filtroNormalizado === 'C'
  ) {

    return (
      valorTurno ===
      filtroNormalizado
    );
  }


  /*
   * SEPRI:
   * acepta SEPRI A, SEPRI B, SEPRI C,
   * SEPRI y cualquier otra variante
   * que contenga la palabra SEPRI.
   */
  if (
    filtroNormalizado === 'SEPRI'
  ) {

    return (
      valorTurno.indexOf('SEPRI') !== -1
    );
  }


  /*
   * GURIN:
   * acepta cualquier valor que contenga GURIN.
   *
   * Ejemplo:
   * ESC – GURIN
   */
  if (
    filtroNormalizado === 'GURIN'
  ) {

    return (
      valorTurno.indexOf('GURIN') !== -1
    );
  }


  /*
   * NO APLICA:
   *
   * Incluye:
   * NO APLICA - A
   * NO APLICA - B
   * NO APLICA - C
   * NO APLICA
   *
   * Y también los registros sin turno.
   */
  if (
    filtroNormalizado === 'NO APLICA'
  ) {

    return (
      valorTurno === '' ||
      valorTurno.indexOf('NO APLICA') !== -1
    );
  }


  return false;
}


/**
 * Obtiene una fila completa de LISTADO_BASE
 * y le agrega información relacionada con
 * novedades.
 */
function construirRegistroTurno_(
  encabezados,
  fila,
  numeroFila
) {

  const funcionario =
    construirFuncionario_(
      encabezados,
      fila,
      numeroFila
    );


  /*
   * Historial de novedades.
   */
  let historial = [];


  try {

    if (
      funcionario.cedula
    ) {

      historial =
        obtenerHistorialNovedadesInterno_(
          funcionario.cedula
        );
    }

  } catch (error) {

    /*
     * La consulta principal de LISTADO_BASE
     * no se detiene si NOVEDADES presenta
     * una inconsistencia.
     */
    historial = [];
  }


  /*
   * Novedades actualmente activas.
   */
  const activas =
    historial.filter(
      function(registro) {

        return novedadEstaActiva_(
          registro['Fecha PRESENTACION']
        );
      }
    );


  /*
   * Se agregan propiedades auxiliares.
   */
  funcionario.historialNovedades =
    historial;


  funcionario.novedadesActivas =
    activas;


  funcionario.totalNovedades =
    historial.length;


  funcionario.totalNovedadesActivas =
    activas.length;


  funcionario.tieneNovedadActiva =
    activas.length > 0;


  return funcionario;
}


/**
 * Consulta completa por turno.
 *
 * ESTA ES LA FUNCIÓN PRINCIPAL QUE UTILIZARÁ
 * EL FRONTEND PARA LOS BOTONES DE TURNOS.
 *
 * Ejemplos:
 *
 * consultarPorTurno(token, 'A')
 * consultarPorTurno(token, 'B')
 * consultarPorTurno(token, 'C')
 * consultarPorTurno(token, 'SEPRI')
 * consultarPorTurno(token, 'GURIN')
 * consultarPorTurno(token, 'NO APLICA')
 */
function consultarPorTurno(
  token,
  filtro
) {

  // ============================================================
  // 1. VALIDAR SESIÓN Y FILTRO
  // ============================================================

  const sesion = validarSesion_(token);

  const filtroNormalizado =
    normalizarFiltroTurno_(filtro);

  if (!filtroNormalizado) {
    return respuestaError_(
      'Filtro de turno no válido. ' +
      'Los filtros permitidos son: A, B, C, SEPRI, GURIN y NO APLICA.'
    );
  }

  // ============================================================
  // 2. LEER LISTADO_BASE UNA SOLA VEZ
  // ============================================================

  const hoja = obtenerHoja_(HOJA_BASE);

  if (!hoja) {
    return respuestaError_(
      'No existe la hoja "' + HOJA_BASE + '".'
    );
  }

  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaFila < 2 || ultimaColumna < 1) {
    return respuestaOK_(
      'No existen funcionarios registrados en LISTADO_BASE.',
      {
        filtro: filtroNormalizado,
        funcionarios: [],
        total: 0,
        estadisticas: {
          total: 0,
          conNovedadActiva: 0,
          sinNovedadActiva: 0
        },
        fechaConsulta: formatearFechaHora_(new Date())
      }
    );
  }

  const encabezados = obtenerEncabezadosBase_();

  const indiceTurno = obtenerIndiceColumna_(
    encabezados,
    'TURNO'
  );

  if (indiceTurno < 0) {
    return respuestaError_(
      'No se encontró la columna TURNO en LISTADO_BASE.'
    );
  }

  const datos = hoja.getRange(
    2,
    1,
    ultimaFila - 1,
    ultimaColumna
  ).getValues();

  // ============================================================
  // 3. FILTRAR FUNCIONARIOS DEL TURNO
  // ============================================================

  const resultados = [];
  const cedulasTurno = {};

  datos.forEach(function(fila, indice) {

    const turno = fila[indiceTurno];

    if (!turnoCoincideFiltro_(
      turno,
      filtroNormalizado
    )) {
      return;
    }

    const funcionario = construirFuncionario_(
      encabezados,
      fila,
      indice + 2
    );

    if (!usuarioPuedeConsultarFuncionario_(
      sesion,
      funcionario
    )) {
      return;
    }

    const cedula = normalizarCedula_(
      funcionario.cedula
    );

    if (cedula) {
      cedulasTurno[cedula] = true;
    }

    // Inicializar información de novedades.
    funcionario.historialNovedades = [];
    funcionario.novedadesActivas = [];
    funcionario.totalNovedades = 0;
    funcionario.totalNovedadesActivas = 0;
    funcionario.tieneNovedadActiva = false;

    resultados.push(funcionario);
  });

  // ============================================================
  // 4. SI NO HAY RESULTADOS, TERMINAR AQUÍ
  // ============================================================

  if (resultados.length === 0) {

    registrarAuditoria_(
      sesion.usuario,
      'CONSULTA_POR_TURNO',
      '',
      'Filtro: ' + filtroNormalizado +
      ' | Registros: 0'
    );

    return respuestaOK_(
      'No se encontraron funcionarios para el filtro "' +
      filtroNormalizado + '".',
      {
        filtro: filtroNormalizado,
        funcionarios: [],
        total: 0,
        estadisticas: {
          total: 0,
          conNovedadActiva: 0,
          sinNovedadActiva: 0
        },
        fechaConsulta: formatearFechaHora_(new Date())
      }
    );
  }

  // ============================================================
  // 5. LEER NOVEDADES UNA SOLA VEZ
  // ============================================================

  const novedadesPorCedula = {};
  const hojaNovedades = obtenerHoja_(HOJA_NOVEDADES);

  if (hojaNovedades) {

    const ultimaFilaNovedades =
      hojaNovedades.getLastRow();

    const ultimaColumnaNovedades =
      hojaNovedades.getLastColumn();

    if (
      ultimaFilaNovedades >= 2 &&
      ultimaColumnaNovedades >= 1
    ) {

      const encabezadosNovedades =
        obtenerEncabezadosNovedades_();

      const indiceCC =
        obtenerIndiceColumnaNovedades_(
          encabezadosNovedades,
          'CC'
        );

      if (indiceCC >= 0) {

        const datosNovedades =
          hojaNovedades.getRange(
            2,
            1,
            ultimaFilaNovedades - 1,
            ultimaColumnaNovedades
          ).getValues();

        datosNovedades.forEach(function(
          fila,
          indice
        ) {

          const cedula = normalizarCedula_(
            fila[indiceCC]
          );

          // Solo indexar novedades de funcionarios
          // que pertenecen al turno consultado.
          if (!cedula || !cedulasTurno[cedula]) {
            return;
          }

          const registro = {};

          encabezadosNovedades.forEach(
            function(encabezado, columna) {
              registro[encabezado] = fila[columna];
            }
          );

          registro.fila = indice + 2;

          if (!novedadesPorCedula[cedula]) {
            novedadesPorCedula[cedula] = [];
          }

          novedadesPorCedula[cedula].push(registro);
        });
      }
    }
  }

  // ============================================================
  // 6. ASOCIAR NOVEDADES A LOS FUNCIONARIOS
  // ============================================================

  resultados.forEach(function(funcionario) {

    const cedula = normalizarCedula_(
      funcionario.cedula
    );

    const historial = cedula &&
      novedadesPorCedula[cedula]
        ? novedadesPorCedula[cedula]
        : [];

    historial.sort(function(a, b) {

      const fechaA =
        a['Fecha INICIAL'] instanceof Date
          ? a['Fecha INICIAL'].getTime()
          : 0;

      const fechaB =
        b['Fecha INICIAL'] instanceof Date
          ? b['Fecha INICIAL'].getTime()
          : 0;

      return fechaB - fechaA;
    });

    const activas = historial.filter(
      function(registro) {
        return novedadEstaActiva_(
          registro['Fecha PRESENTACION']
        );
      }
    );

    funcionario.historialNovedades = historial;
    funcionario.novedadesActivas = activas;
    funcionario.totalNovedades = historial.length;
    funcionario.totalNovedadesActivas = activas.length;
    funcionario.tieneNovedadActiva = activas.length > 0;
  });

  // ============================================================
  // 7. ORDENAR RESULTADOS
  // ============================================================

  resultados.sort(function(a, b) {

    const nombreA = normalizarNombre_(
      a.funcionario
    );

    const nombreB = normalizarNombre_(
      b.funcionario
    );

    return nombreA.localeCompare(
      nombreB,
      'es',
      {
        sensitivity: 'base'
      }
    );
  });

  // ============================================================
  // 8. ESTADÍSTICAS
  // ============================================================

  const estadisticas = {

    total: resultados.length,

    conNovedadActiva:
      resultados.filter(function(item) {
        return item.tieneNovedadActiva === true;
      }).length,

    sinNovedadActiva:
      resultados.filter(function(item) {
        return item.tieneNovedadActiva !== true;
      }).length
  };

  // ============================================================
  // 9. AUDITORÍA
  // ============================================================

  registrarAuditoria_(
    sesion.usuario,
    'CONSULTA_POR_TURNO',
    '',
    'Filtro: ' + filtroNormalizado +
    ' | Registros: ' + resultados.length
  );

  // ============================================================
  // 10. RESPUESTA FINAL
  // ============================================================

  return respuestaOK_(
    'Consulta por turno realizada correctamente.',
    {
      filtro: filtroNormalizado,
      funcionarios: resultados,
      total: resultados.length,
      estadisticas: estadisticas,
      fechaConsulta: formatearFechaHora_(new Date())
    }
  );
}


/**
 * Alias para el frontend.
 *
 * Permite utilizar:
 * consultarTurno(...)
 *
 * o:
 * consultarPorTurno(...)
 */
function consultarTurno(
  token,
  filtro
) {

  return consultarPorTurno(
    token,
    filtro
  );
}


/**
 * Devuelve únicamente el listado resumido
 * de funcionarios de un turno.
 *
 * Es útil para tablas donde no se necesita
 * enviar todo el historial de novedades.
 */
function listarFuncionariosPorTurno(
  token,
  filtro
) {

  const resultado =
    consultarPorTurno(
      token,
      filtro
    );


  if (
    !resultado ||
    resultado.estado !== true
  ) {

    return resultado;
  }


  const datos =
    resultado.datos || {};


  const funcionarios =
    datos.funcionarios || [];


  const listado =
    funcionarios.map(
      function(funcionario) {

        return {

          fila:
            funcionario.fila,

          cedula:
            funcionario.cedula,

          nivel:
            funcionario.nivel,

          grado:
            funcionario.grado,

          funcionario:
            funcionario.funcionario,

          dependencia:
            funcionario.dependencia,

          turno:
            funcionario.turno,

          placaChip:
            funcionario.placaChip,

          totalNovedades:
            funcionario.totalNovedades,

          totalNovedadesActivas:
            funcionario.totalNovedadesActivas,

          tieneNovedadActiva:
            funcionario.tieneNovedadActiva
        };
      }
    );


  return respuestaOK_(
    'Listado resumido generado correctamente.',
    {
      filtro:
        datos.filtro,

      funcionarios:
        listado,

      total:
        listado.length
    }
  );
}


/**
 * Devuelve los conteos de los filtros
 * disponibles para el usuario.
 *
 * No utiliza TURNOS_VALIDOS porque SEPRI,
 * GURIN y NO APLICA tienen reglas especiales.
 */
function obtenerConteosTurnos(
  token
) {

  const sesion =
    validarSesion_(token);


  const hoja =
    obtenerHoja_(HOJA_BASE);


  if (!hoja) {

    return respuestaError_(
      'No existe la hoja "' +
      HOJA_BASE +
      '".'
    );
  }


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2
  ) {

    return respuestaOK_(
      'No existen funcionarios registrados.',
      {
        A: 0,
        B: 0,
        C: 0,
        SEPRI: 0,
        GURIN: 0,
        'NO APLICA': 0
      }
    );
  }


  const encabezados =
    obtenerEncabezadosBase_();


  const indiceTurno =
    obtenerIndiceColumna_(
      encabezados,
      'TURNO'
    );


  if (
    indiceTurno < 0
  ) {

    return respuestaError_(
      'No se encontró la columna TURNO en LISTADO_BASE.'
    );
  }


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  const conteos = {

    A: 0,

    B: 0,

    C: 0,

    SEPRI: 0,

    GURIN: 0,

    'NO APLICA': 0
  };


  datos.forEach(
    function(fila, indice) {

      const turno =
        fila[indiceTurno];


      /*
       * Construir funcionario únicamente
       * para validar permisos.
       */
      const funcionario =
        construirFuncionario_(
          encabezados,
          fila,
          indice + 2
        );


      if (
        !usuarioPuedeConsultarFuncionario_(
          sesion,
          funcionario
        )
      ) {

        return;
      }


      /*
       * Contabilizar cada filtro.
       */
      if (
        turnoCoincideFiltro_(
          turno,
          'A'
        )
      ) {

        conteos.A++;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'B'
        )
      ) {

        conteos.B++;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'C'
        )
      ) {

        conteos.C++;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'SEPRI'
        )
      ) {

        conteos.SEPRI++;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'GURIN'
        )
      ) {

        conteos.GURIN++;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'NO APLICA'
        )
      ) {

        conteos['NO APLICA']++;
      }
    }
  );


  registrarAuditoria_(
    sesion.usuario,
    'CONSULTA_CONTEOS_TURNOS',
    '',
    JSON.stringify(conteos)
  );


  return respuestaOK_(
    'Conteos de turnos obtenidos correctamente.',
    conteos
  );
}


/**
 * Diagnóstico específico de turnos.
 *
 * Solo ADMINISTRADOR.
 *
 * Permite comprobar qué valores reales existen
 * en la columna TURNO y cómo son clasificados.
 */
function diagnosticarTurnos(
  token
) {

  const sesion =
    validarSesion_(token);


  validarRol_(
    sesion,
    ['ADMINISTRADOR']
  );


  const hoja =
    obtenerHoja_(HOJA_BASE);


  if (!hoja) {

    return respuestaError_(
      'No existe la hoja "' +
      HOJA_BASE +
      '".'
    );
  }


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2
  ) {

    return respuestaOK_(
      'No existen registros para diagnosticar.',
      {}
    );
  }


  const encabezados =
    obtenerEncabezadosBase_();


  const indiceTurno =
    obtenerIndiceColumna_(
      encabezados,
      'TURNO'
    );


  if (
    indiceTurno < 0
  ) {

    return respuestaError_(
      'No se encontró la columna TURNO.'
    );
  }


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  const valores = {};


  datos.forEach(
    function(fila) {

      const turnoOriginal =
        texto_(
          fila[indiceTurno]
        );


      const clave =
        turnoOriginal ||
        '(VACÍO)';


      valores[clave] =
        (valores[clave] || 0) + 1;
    }
  );


  /*
   * Clasificación según los botones
   * actualmente definidos.
   */
  const clasificacion = {

    A: 0,

    B: 0,

    C: 0,

    SEPRI: 0,

    GURIN: 0,

    'NO APLICA': 0,

    'SIN CLASIFICAR': 0
  };


  datos.forEach(
    function(fila) {

      const turno =
        fila[indiceTurno];


      let clasificado =
        false;


      if (
        turnoCoincideFiltro_(
          turno,
          'A'
        )
      ) {

        clasificacion.A++;
        clasificado = true;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'B'
        )
      ) {

        clasificacion.B++;
        clasificado = true;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'C'
        )
      ) {

        clasificacion.C++;
        clasificado = true;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'SEPRI'
        )
      ) {

        clasificacion.SEPRI++;
        clasificado = true;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'GURIN'
        )
      ) {

        clasificacion.GURIN++;
        clasificado = true;
      }


      if (
        turnoCoincideFiltro_(
          turno,
          'NO APLICA'
        )
      ) {

        clasificacion['NO APLICA']++;
        clasificado = true;
      }


      if (!clasificado) {

        clasificacion['SIN CLASIFICAR']++;
      }
    }
  );


  registrarAuditoria_(
    sesion.usuario,
    'DIAGNOSTICO_TURNOS',
    '',
    'Diagnóstico de valores de TURNO.'
  );


  return respuestaOK_(
    'Diagnóstico de turnos realizado correctamente.',
    {
      valoresOriginales: valores,
      clasificacion: clasificacion
    }
  );
}


/*******************************************************
 * FIN PARTE 5
 *******************************************************/

 /*******************************************************
 * PARTE 6
 * REPORTES Y PREPARACIÓN PARA PDF
 *
 * HOJA: REPORTES
 *
 * COLUMNAS:
 * A = CONSECUTIVO
 * B = FECHA
 * C = USUARIO
 * D = TURNO
 * E = ARCHIVO
 * F = URL
 *******************************************************/


/**
 * Obtiene los encabezados de la hoja REPORTES.
 */
function obtenerEncabezadosReportes_() {

  const hoja =
    obtenerHoja_(HOJA_REPORTES);


  if (!hoja) {

    throw new Error(
      'No existe la hoja "' +
      HOJA_REPORTES +
      '".'
    );
  }


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaColumna < 1
  ) {

    return [];
  }


  return hoja
    .getRange(
      1,
      1,
      1,
      ultimaColumna
    )
    .getValues()[0]
    .map(
      function(valor) {
        return texto_(valor);
      }
    );
}


/**
 * Verifica que la hoja REPORTES tenga
 * la estructura esperada.
 *
 * Si la hoja está vacía, crea los encabezados.
 */
function validarEstructuraReportes_() {

  const hoja =
    obtenerHoja_(HOJA_REPORTES);


  if (!hoja) {

    throw new Error(
      'No existe la hoja "' +
      HOJA_REPORTES +
      '".'
    );
  }


  const encabezadosEsperados = [
    'CONSECUTIVO',
    'FECHA',
    'USUARIO',
    'TURNO',
    'ARCHIVO',
    'URL'
  ];


  const ultimaColumna =
    hoja.getLastColumn();


  /*
   * Si la hoja está completamente vacía,
   * crear estructura.
   */
  if (
    ultimaColumna === 0
  ) {

    hoja
      .getRange(
        1,
        1,
        1,
        encabezadosEsperados.length
      )
      .setValues([
        encabezadosEsperados
      ]);

    return true;
  }


  const encabezados =
    obtenerEncabezadosReportes_();


  /*
   * Si no tiene encabezados,
   * escribirlos.
   */
  if (
    encabezados.length === 0
  ) {

    hoja
      .getRange(
        1,
        1,
        1,
        encabezadosEsperados.length
      )
      .setValues([
        encabezadosEsperados
      ]);

    return true;
  }


  /*
   * Validar las primeras seis columnas.
   */
  for (
    let i = 0;
    i < encabezadosEsperados.length;
    i++
  ) {

    const actual =
      normalizarTexto_(
        encabezados[i] || ''
      );


    const esperado =
      normalizarTexto_(
        encabezadosEsperados[i]
      );


    if (
      actual !== esperado
    ) {

      throw new Error(
        'Estructura incorrecta en REPORTES. ' +
        'La columna ' +
        String.fromCharCode(65 + i) +
        ' debe ser "' +
        encabezadosEsperados[i] +
        '".'
      );
    }
  }


  return true;
}


/**
 * Obtiene el siguiente consecutivo.
 *
 * Formato:
 *
 * REP-000001
 * REP-000002
 * REP-000003
 */
function obtenerSiguienteConsecutivoReporte_() {

  const hoja =
    obtenerHoja_(HOJA_REPORTES);


  const ultimaFila =
    hoja.getLastRow();


  if (
    ultimaFila < 2
  ) {

    return 'REP-000001';
  }


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        1
      )
      .getValues();


  let mayor =
    0;


  datos.forEach(
    function(fila) {

      const valor =
        texto_(
          fila[0]
        );


      if (!valor) {
        return;
      }


      /*
       * Extraer número.
       *
       * Ejemplo:
       * REP-000025
       */
      const coincidencia =
        valor.match(
          /(\d+)$/
        );


      if (
        coincidencia
      ) {

        const numero =
          Number(
            coincidencia[1]
          );


        if (
          numero > mayor
        ) {

          mayor = numero;
        }
      }
    }
  );


  const siguiente =
    mayor + 1;


  return (
    'REP-' +
    String(siguiente)
      .padStart(
        6,
        '0'
      )
  );
}


/**
 * Escapa caracteres HTML.
 *
 * Evita que datos provenientes de las hojas
 * sean interpretados como código HTML.
 */
function escaparHTML_(valor) {

  if (
    valor === null ||
    valor === undefined
  ) {

    return '';
  }


  return String(valor)
    .replace(
      /&/g,
      '&amp;'
    )
    .replace(
      /</g,
      '&lt;'
    )
    .replace(
      />/g,
      '&gt;'
    )
    .replace(
      /"/g,
      '&quot;'
    )
    .replace(
      /'/g,
      '&#039;'
    );
}


/**
 * Convierte cualquier valor recibido
 * en texto para impresión.
 */
function valorParaPDF_(valor) {

  if (
    valor === null ||
    valor === undefined
  ) {

    return '';
  }


  if (
    Object.prototype.toString.call(
      valor
    ) === '[object Date]'
  ) {

    return formatearFecha_(
      valor
    );
  }


  return texto_(valor);
}


/**
 * Construye el HTML completo del listado
 * filtrado por turno.
 *
 * Este HTML será enviado al navegador
 * para impresión / PDF.
 */
function construirHTMLReporteTurno_(
  filtro,
  funcionarios,
  consecutivo,
  usuario
) {

  const fecha =
    new Date();


  const fechaTexto =
    formatearFechaHora_(
      fecha
    );


  /*
   * Si no existen funcionarios,
   * igualmente se genera un reporte
   * indicando ausencia de resultados.
   */
  const total =
    funcionarios.length;


  /*
   * Construcción de filas.
   */
  let filasHTML = '';


  funcionarios.forEach(
    function(funcionario, indice) {

      const numero =
        indice + 1;


      const tieneNovedad =
        funcionario.tieneNovedadActiva
          ? 'SÍ'
          : 'NO';


      filasHTML +=
        '<tr>' +

          '<td class="centro">' +
            numero +
          '</td>' +

          '<td>' +
            escaparHTML_(
              funcionario.cedula
            ) +
          '</td>' +

          '<td>' +
            escaparHTML_(
              funcionario.nivel
            ) +
          '</td>' +

          '<td>' +
            escaparHTML_(
              funcionario.grado
            ) +
          '</td>' +

          '<td>' +
            escaparHTML_(
              funcionario.funcionario
            ) +
          '</td>' +

          '<td>' +
            escaparHTML_(
              funcionario.dependencia
            ) +
          '</td>' +

          '<td class="centro">' +
            escaparHTML_(
              funcionario.turno
            ) +
          '</td>' +

          '<td>' +
            escaparHTML_(
              funcionario.placaChip
            ) +
          '</td>' +

          '<td class="centro">' +
            funcionario.totalNovedades +
          '</td>' +

          '<td class="centro">' +
            tieneNovedad +
          '</td>' +

        '</tr>';
    }
  );


  /*
   * HTML completo.
   */
  const html =
`<!DOCTYPE html>
<html lang="es">
<head>

<meta charset="UTF-8">

<title>
Listado por turno - ${escaparHTML_(filtro)}
</title>

<style>

@page {
  size: landscape;
  margin: 12mm;
}

* {
  box-sizing: border-box;
}

body {
  font-family: Arial, Helvetica, sans-serif;
  margin: 0;
  padding: 0;
  color: #111;
  font-size: 10px;
}

.encabezado {
  width: 100%;
  border-bottom: 2px solid #222;
  padding-bottom: 8px;
  margin-bottom: 10px;
}

.titulo {
  text-align: center;
  font-size: 18px;
  font-weight: bold;
  margin-bottom: 4px;
}

.subtitulo {
  text-align: center;
  font-size: 12px;
  margin-bottom: 8px;
}

.informacion {
  width: 100%;
  display: table;
  margin-bottom: 10px;
}

.informacion-fila {
  display: table-row;
}

.informacion-celda {
  display: table-cell;
  padding: 3px 8px 3px 0;
}

.etiqueta {
  font-weight: bold;
}

table {
  width: 100%;
  border-collapse: collapse;
  table-layout: auto;
}

thead {
  display: table-header-group;
}

tfoot {
  display: table-footer-group;
}

tr {
  page-break-inside: avoid;
}

th {
  background: #e8e8e8;
  border: 1px solid #555;
  padding: 5px;
  text-align: center;
  font-weight: bold;
  font-size: 9px;
}

td {
  border: 1px solid #777;
  padding: 4px;
  vertical-align: middle;
}

.centro {
  text-align: center;
}

.pie {
  margin-top: 10px;
  border-top: 1px solid #777;
  padding-top: 5px;
  font-size: 8px;
}

.sin-resultados {
  text-align: center;
  padding: 20px;
  font-weight: bold;
}

@media print {

  .no-imprimir {
    display: none !important;
  }

}

</style>

</head>

<body>

<div class="encabezado">

  <div class="titulo">
    POLICÍA NACIONAL DE COLOMBIA
  </div>

  <div class="subtitulo">
    SISTEMA DE GESTIÓN DE NOVEDADES
  </div>

  <div class="subtitulo">
    LISTADO DE FUNCIONARIOS POR TURNO
  </div>

</div>


<div class="informacion">

  <div class="informacion-fila">

    <div class="informacion-celda">
      <span class="etiqueta">
        Filtro:
      </span>
      ${escaparHTML_(filtro)}
    </div>

    <div class="informacion-celda">
      <span class="etiqueta">
        Total funcionarios:
      </span>
      ${total}
    </div>

  </div>


  <div class="informacion-fila">

    <div class="informacion-celda">
      <span class="etiqueta">
        Usuario:
      </span>
      ${escaparHTML_(usuario)}
    </div>

    <div class="informacion-celda">
      <span class="etiqueta">
        Fecha:
      </span>
      ${escaparHTML_(fechaTexto)}
    </div>

  </div>


  <div class="informacion-fila">

    <div class="informacion-celda">
      <span class="etiqueta">
        Reporte:
      </span>
      ${escaparHTML_(consecutivo)}
    </div>

  </div>

</div>


<table>

<thead>

<tr>

<th>
#
</th>

<th>
CÉDULA
</th>

<th>
NIV
</th>

<th>
GRADO
</th>

<th>
FUNCIONARIO
</th>

<th>
DEPENDENCIA
</th>

<th>
TURNO
</th>

<th>
PLACA / CHIP
</th>

<th>
NOVEDADES
</th>

<th>
NOVEDAD ACTIVA
</th>

</tr>

</thead>


<tbody>

${
  filasHTML ||
  '<tr>' +
  '<td colspan="10" class="sin-resultados">' +
  'No se encontraron funcionarios para el filtro seleccionado.' +
  '</td>' +
  '</tr>'
}

</tbody>

</table>


<div class="pie">

  Reporte generado mediante el
  Sistema de Gestión de Novedades.

  <br>

  Consecutivo:
  ${escaparHTML_(consecutivo)}

</div>


</body>
</html>`;


  return html;
}


/**
 * Genera el reporte de funcionarios por turno.
 *
 * El resultado contiene:
 *
 * - consecutivo
 * - filtro
 * - total
 * - funcionarios
 * - html
 *
 * El HTML se puede abrir en el navegador
 * y utilizar "Imprimir → Guardar como PDF".
 */
function generarReporteTurno(
  token,
  filtro
) {

  const sesion =
    validarSesion_(token);


  /*
   * Validar estructura.
   */
  validarEstructuraReportes_();


  /*
   * Normalizar filtro.
   */
  const filtroNormalizado =
    normalizarFiltroTurno_(
      filtro
    );


  if (
    !filtroNormalizado
  ) {

    return respuestaError_(
      'Filtro de turno no válido.'
    );
  }


  /*
   * Consultar funcionarios.
   */
  const consulta =
    consultarPorTurno(
      token,
      filtroNormalizado
    );


  if (
    !consulta ||
    consulta.estado !== true
  ) {

    return consulta;
  }


  const datos =
    consulta.datos || {};


  const funcionarios =
    datos.funcionarios || [];


  /*
   * Consecutivo.
   */
  const consecutivo =
    obtenerSiguienteConsecutivoReporte_();


  /*
   * Construir HTML.
   */
  const html =
    construirHTMLReporteTurno_(
      filtroNormalizado,
      funcionarios,
      consecutivo,
      sesion.usuario
    );


  /*
   * Registrar reporte.
   */
  const hoja =
    obtenerHoja_(HOJA_REPORTES);


  const filaReporte = [

    consecutivo,

    new Date(),

    sesion.usuario,

    filtroNormalizado,

    'Listado_' +
    filtroNormalizado.replace(
      /\s+/g,
      '_'
    ) +
    '_' +
    consecutivo,

    ''
  ];


  hoja.appendRow(
    filaReporte
  );


  const filaNueva =
    hoja.getLastRow();


  /*
   * Formato de fecha.
   */
  hoja
    .getRange(
      filaNueva,
      2
    )
    .setNumberFormat(
      'dd/MM/yyyy HH:mm:ss'
    );


  /*
   * Auditoría.
   */
  registrarAuditoria_(
    sesion.usuario,
    'GENERAR_REPORTE_TURNO',
    '',
    'Reporte: ' +
    consecutivo +
    ' | Filtro: ' +
    filtroNormalizado +
    ' | Total: ' +
    funcionarios.length
  );


  return respuestaOK_(
    'Reporte generado correctamente.',
    {
      consecutivo:
        consecutivo,

      filtro:
        filtroNormalizado,

      total:
        funcionarios.length,

      funcionarios:
        funcionarios,

      html:
        html,

      fecha:
        formatearFechaHora_(
          new Date()
        )
    }
  );
}


/**
 * Genera únicamente la información necesaria
 * para una vista previa del PDF.
 *
 * No registra un reporte nuevo.
 */
function previsualizarReporteTurno(
  token,
  filtro
) {

  const sesion =
    validarSesion_(token);


  const filtroNormalizado =
    normalizarFiltroTurno_(
      filtro
    );


  if (
    !filtroNormalizado
  ) {

    return respuestaError_(
      'Filtro de turno no válido.'
    );
  }


  const consulta =
    consultarPorTurno(
      token,
      filtroNormalizado
    );


  if (
    !consulta ||
    consulta.estado !== true
  ) {

    return consulta;
  }


  const datos =
    consulta.datos || {};


  const funcionarios =
    datos.funcionarios || [];


  /*
   * Para la vista previa se utiliza un
   * consecutivo temporal.
   */
  const consecutivo =
    'PREVISUALIZACION';


  const html =
    construirHTMLReporteTurno_(
      filtroNormalizado,
      funcionarios,
      consecutivo,
      sesion.usuario
    );


  return respuestaOK_(
    'Vista previa generada correctamente.',
    {
      filtro:
        filtroNormalizado,

      total:
        funcionarios.length,

      funcionarios:
        funcionarios,

      html:
        html
    }
  );
}


/**
 * Consulta el historial de reportes generados.
 *
 * ADMINISTRADOR:
 *   Todos los reportes.
 *
 * OPERADOR:
 *   Únicamente sus propios reportes.
 */
function listarReportes(
  token,
  limite
) {

  const sesion =
    validarSesion_(token);


  validarEstructuraReportes_();


  const hoja =
    obtenerHoja_(HOJA_REPORTES);


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  if (
    ultimaFila < 2
  ) {

    return respuestaOK_(
      'No existen reportes registrados.',
      {
        reportes: [],
        total: 0
      }
    );
  }


  const encabezados =
    obtenerEncabezadosReportes_();


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getValues();


  let maximo =
    Number(limite);


  if (
    !isFinite(maximo) ||
    maximo <= 0
  ) {

    maximo = 100;
  }


  maximo =
    Math.min(
      maximo,
      500
    );


  const reportes = [];


  datos.forEach(
    function(fila, indice) {

      if (
        reportes.length >= maximo
      ) {

        return;
      }


      const reporte = {};


      encabezados.forEach(
        function(encabezado, columna) {

          reporte[encabezado] =
            fila[columna];
        }
      );


      reporte.fila =
        indice + 2;


      /*
       * Operador:
       * solo puede consultar sus reportes.
       */
      if (
        normalizarTexto_(sesion.rol) ===
        'OPERADOR'
      ) {

        if (
          normalizarTexto_(
            reporte['USUARIO']
          ) !==
          normalizarTexto_(
            sesion.usuario
          )
        ) {

          return;
        }
      }


      reportes.push(
        reporte
      );
    }
  );


  /*
   * Más recientes primero.
   */
  reportes.sort(
    function(a, b) {

      const fechaA =
        a['FECHA'] instanceof Date
          ? a['FECHA'].getTime()
          : 0;


      const fechaB =
        b['FECHA'] instanceof Date
          ? b['FECHA'].getTime()
          : 0;


      return fechaB - fechaA;
    }
  );


  registrarAuditoria_(
    sesion.usuario,
    'CONSULTAR_REPORTES',
    '',
    'Reportes retornados: ' +
    reportes.length
  );


  return respuestaOK_(
    'Reportes consultados correctamente.',
    {
      reportes:
        reportes,

      total:
        reportes.length
    }
  );
}


/**
 * Consulta un reporte específico por consecutivo.
 */
function consultarReporte(
  token,
  consecutivo
) {

  const sesion =
    validarSesion_(token);


  const consecutivoBuscado =
    texto_(
      consecutivo
    ).trim();


  if (
    !consecutivoBuscado
  ) {

    return respuestaError_(
      'Debe indicar el consecutivo del reporte.'
    );
  }


  validarEstructuraReportes_();


  const hoja =
    obtenerHoja_(HOJA_REPORTES);


  const ultimaFila =
    hoja.getLastRow();


  if (
    ultimaFila < 2
  ) {

    return respuestaOK_(
      'No existen reportes registrados.',
      null
    );
  }


  const encabezados =
    obtenerEncabezadosReportes_();


  const indiceConsecutivo =
    obtenerIndiceColumna_(
      encabezados,
      'CONSECUTIVO'
    );


  const indiceUsuario =
    obtenerIndiceColumna_(
      encabezados,
      'USUARIO'
    );


  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        hoja.getLastColumn()
      )
      .getValues();


  let encontrado =
    null;


  for (
    let i = 0;
    i < datos.length;
    i++
  ) {

    const valor =
      texto_(
        datos[i][indiceConsecutivo]
      ).trim();


    if (
      valor !==
      consecutivoBuscado
    ) {

      continue;
    }


    /*
     * Restricción del operador.
     */
    if (
      normalizarTexto_(sesion.rol) ===
      'OPERADOR'
    ) {

      const usuarioReporte =
        normalizarTexto_(
          datos[i][indiceUsuario]
        );


      if (
        usuarioReporte !==
        normalizarTexto_(
          sesion.usuario
        )
      ) {

        return respuestaError_(
          'El usuario no tiene permisos para consultar este reporte.'
        );
      }
    }


    encontrado = {};


    encabezados.forEach(
      function(encabezado, columna) {

        encontrado[encabezado] =
          datos[i][columna];
      }
    );


    encontrado.fila =
      i + 2;


    break;
  }


  if (!encontrado) {

    return respuestaOK_(
      'No se encontró el reporte solicitado.',
      null
    );
  }


  registrarAuditoria_(
    sesion.usuario,
    'CONSULTAR_REPORTE',
    '',
    'Consecutivo: ' +
    consecutivoBuscado
  );


  return respuestaOK_(
    'Reporte consultado correctamente.',
    encontrado
  );
}


/**
 * Diagnóstico de la hoja REPORTES.
 *
 * Solo ADMINISTRADOR.
 */
function diagnosticarReportes(
  token
) {

  const sesion =
    validarSesion_(token);


  validarRol_(
    sesion,
    ['ADMINISTRADOR']
  );


  validarEstructuraReportes_();


  const hoja =
    obtenerHoja_(HOJA_REPORTES);


  const ultimaFila =
    hoja.getLastRow();


  const ultimaColumna =
    hoja.getLastColumn();


  const encabezados =
    obtenerEncabezadosReportes_();


  const registros =
    Math.max(
      0,
      ultimaFila - 1
    );


  registrarAuditoria_(
    sesion.usuario,
    'DIAGNOSTICO_REPORTES',
    '',
    'Registros: ' +
    registros
  );


  return respuestaOK_(
    'Diagnóstico de REPORTES realizado correctamente.',
    {
      filas:
        registros,

      columnas:
        ultimaColumna,

      encabezados:
        encabezados
    }
  );
}


/*******************************************************
 * FIN PARTE 6
 *******************************************************/