/**
 * Seguimiento de pedidos – Mis Angelitos Baby
 *
 * Una sola web (Aplicación web de Apps Script):
 *   - Sin parámetros        → página pública para que el cliente consulte su pedido.
 *   - ...?p=admin           → panel privado (con clave) para crear y actualizar pedidos.
 * La hoja de Google Sheets solo hace de base de datos.
 */

const HOJA_PEDIDOS = 'Pedidos';
const HOJA_HISTORIAL = 'Historial';
const HOJA_CONFIG = 'Config';

const ESTADOS = [
  'Recibido',
  'Pago confirmado',
  'En preparación',
  'Enviado',
  'En tránsito',
  'Entregado',
  'Cancelado'
];

const MENSAJES_ESTADO = {
  'Recibido': 'Recibimos tu pedido. En cuanto confirmemos el pago te avisaremos.',
  'Pago confirmado': 'Confirmamos tu pago. ¡Gracias por confiar en nosotros!',
  'En preparación': 'Estamos preparando tu pedido con mucho cariño.',
  'Enviado': 'Tu pedido ya salió de nuestras manos y va en camino.',
  'En tránsito': 'Tu pedido está en tránsito y llegará pronto.',
  'Entregado': 'Tu pedido fue entregado. ¡Esperamos que lo disfruten mucho!',
  'Cancelado': 'Tu pedido fue cancelado. Si tienes dudas, escríbenos y te ayudamos.'
};

// Posición (1-based) de cada columna en la hoja Pedidos.
const COL = {
  ID: 1, FECHA: 2, CLIENTE: 3, EMAIL: 4, TELEFONO: 5, PRODUCTOS: 6, TOTAL: 7,
  DIRECCION: 8, ESTADO: 9, TRANSPORTADORA: 10, GUIA: 11, NOTA: 12, ACTUALIZADO: 13
};

const ENCABEZADOS_PEDIDOS = [
  'ID Pedido', 'Fecha', 'Cliente', 'Email', 'Teléfono', 'Productos', 'Total',
  'Dirección / Ciudad', 'Estado', 'Transportadora', 'Guía', 'Nota para el cliente',
  'Última actualización'
];

const ENCABEZADOS_HISTORIAL = ['Fecha', 'Pedido', 'Tipo', 'Estado anterior', 'Estado nuevo', 'Nota', 'Correo'];

const CONFIG_POR_DEFECTO = [
  ['NOMBRE_TIENDA', 'Mis Angelitos Baby'],
  ['NOTIFICAR_CLIENTE', 'SI'],
  ['EMAIL_RESPUESTA', ''],
  ['URL_RASTREO_TRANSPORTADORA', '']
];

const MAX_INTENTOS_CONSULTA = 8;   // consultas del cliente, por pedido y por hora
const MAX_INTENTOS_LOGIN = 5;      // intentos de clave del panel, por 15 minutos
const DURACION_SESION_S = 6 * 3600;

/* ============================== Menú ============================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Pedidos')
    .addItem('🔗 Ver enlaces de la web', 'mostrarEnlaces')
    .addToUi();
}

function mostrarEnlaces() {
  const url = obtenerUrlWeb_();
  const ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Aún no has publicado la web.\n\nVe a Implementar > Nueva implementación > Aplicación web ' +
      '(Ejecutar como: yo, Acceso: cualquier persona) y vuelve a abrir esta opción.');
    return;
  }
  ui.alert('Página para tus clientes:\n' + url + '\n\nPanel de administración (solo tú):\n' + url + '?p=admin');
}

function obtenerUrlWeb_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (err) {
    return '';
  }
}

/* ===================== Hojas y configuración ===================== */

/** Crea las hojas la primera vez que se usa; no hace falta ejecutar nada a mano. */
function asegurarHojas_() {
  const ss = SpreadsheetApp.getActive();

  let pedidos = ss.getSheetByName(HOJA_PEDIDOS);
  if (!pedidos) {
    pedidos = ss.insertSheet(HOJA_PEDIDOS);
    pedidos.getRange(1, 1, 1, ENCABEZADOS_PEDIDOS.length).setValues([ENCABEZADOS_PEDIDOS])
      .setFontWeight('bold').setBackground('#f8bbd0').setWrap(true);
    pedidos.setFrozenRows(1);
    pedidos.getRange(2, COL.FECHA, 2000, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    pedidos.getRange(2, COL.ACTUALIZADO, 2000, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    pedidos.getRange(2, COL.TOTAL, 2000, 1).setNumberFormat('$#,##0');
    pedidos.getRange(2, COL.TELEFONO, 2000, 1).setNumberFormat('@');
    pedidos.setColumnWidth(COL.PRODUCTOS, 260);
    pedidos.setColumnWidth(COL.DIRECCION, 220);
    pedidos.setColumnWidth(COL.NOTA, 240);
  }

  let historial = ss.getSheetByName(HOJA_HISTORIAL);
  if (!historial) {
    historial = ss.insertSheet(HOJA_HISTORIAL);
    historial.getRange(1, 1, 1, ENCABEZADOS_HISTORIAL.length).setValues([ENCABEZADOS_HISTORIAL])
      .setFontWeight('bold').setBackground('#f8bbd0');
    historial.setFrozenRows(1);
    historial.getRange(2, 1, 5000, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }

  if (!ss.getSheetByName(HOJA_CONFIG)) {
    const config = ss.insertSheet(HOJA_CONFIG);
    config.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]).setFontWeight('bold').setBackground('#f8bbd0');
    config.getRange(2, 1, CONFIG_POR_DEFECTO.length, 2).setValues(CONFIG_POR_DEFECTO);
    config.setColumnWidth(1, 240);
    config.setColumnWidth(2, 360);
  }
  return ss;
}

function leerConfig_() {
  const config = {};
  CONFIG_POR_DEFECTO.forEach(function (par) { config[par[0]] = par[1]; });
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONFIG);
  if (hoja && hoja.getLastRow() > 1) {
    hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(function (fila) {
      if (fila[0]) config[String(fila[0]).trim()] = String(fila[1]).trim();
    });
  }
  return config;
}

/* =========================== Páginas web =========================== */

function doGet(e) {
  asegurarHojas_();
  const tienda = leerConfig_().NOMBRE_TIENDA;
  const admin = e && e.parameter && e.parameter.p === 'admin';

  const plantilla = HtmlService.createTemplate(admin ? HTML_ADMIN : HTML_SEGUIMIENTO);
  plantilla.tienda = tienda;
  plantilla.estados = admin ? ESTADOS : ESTADOS.filter(function (x) { return x !== 'Cancelado'; });
  return plantilla.evaluate()
    .setTitle((admin ? 'Panel de pedidos · ' : 'Seguimiento de pedido · ') + tienda)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ===================== Acceso al panel (clave) ===================== */

function hashClave_(clave, sal) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + clave, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function nuevaSesion_() {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('adm_' + token, '1', DURACION_SESION_S);
  return token;
}

function exigirAdmin_(token) {
  if (!token || !CacheService.getScriptCache().get('adm_' + token)) throw new Error('SESION');
}

/** ¿Ya existe una clave? Si no, la página pide crearla. */
function estadoPanel() {
  return { claveCreada: !!PropertiesService.getScriptProperties().getProperty('ADMIN_HASH') };
}

function crearClave(clave) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ADMIN_HASH')) throw new Error('La clave ya fue creada.');
  clave = String(clave || '');
  if (clave.length < 8) throw new Error('La clave debe tener al menos 8 caracteres.');
  const sal = Utilities.getUuid();
  props.setProperties({ ADMIN_SAL: sal, ADMIN_HASH: hashClave_(clave, sal) });
  return nuevaSesion_();
}

function iniciarSesion(clave) {
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('ADMIN_HASH');
  if (!hash) throw new Error('Primero crea tu clave.');

  const cache = CacheService.getScriptCache();
  const intentos = Number(cache.get('login_intentos') || 0);
  if (intentos >= MAX_INTENTOS_LOGIN) throw new Error('Demasiados intentos. Espera 15 minutos.');

  if (hashClave_(String(clave || ''), props.getProperty('ADMIN_SAL')) !== hash) {
    cache.put('login_intentos', String(intentos + 1), 900);
    throw new Error('Clave incorrecta.');
  }
  cache.remove('login_intentos');
  return nuevaSesion_();
}

/* ======================= Funciones del panel ======================= */

function listarPedidos(token) {
  exigirAdmin_(token);
  const ss = asegurarHojas_();
  const hoja = ss.getSheetByName(HOJA_PEDIDOS);
  if (hoja.getLastRow() < 2) return [];
  const tz = ss.getSpreadsheetTimeZone();
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, ENCABEZADOS_PEDIDOS.length).getValues()
    .filter(function (f) { return String(f[COL.ID - 1]).trim() !== ''; })
    .map(function (f) { return pedidoDesdeFila_(f, tz); })
    .reverse();
}

function pedidoDesdeFila_(f, tz) {
  const fecha = function (d) {
    return d instanceof Date ? Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm') : '';
  };
  return {
    id: String(f[COL.ID - 1]),
    fecha: fecha(f[COL.FECHA - 1]),
    cliente: String(f[COL.CLIENTE - 1]),
    email: String(f[COL.EMAIL - 1]),
    telefono: String(f[COL.TELEFONO - 1]),
    productos: String(f[COL.PRODUCTOS - 1]),
    total: Number(f[COL.TOTAL - 1]) || 0,
    direccion: String(f[COL.DIRECCION - 1]),
    estado: String(f[COL.ESTADO - 1]),
    transportadora: String(f[COL.TRANSPORTADORA - 1]),
    guia: String(f[COL.GUIA - 1]),
    nota: String(f[COL.NOTA - 1]),
    actualizado: fecha(f[COL.ACTUALIZADO - 1])
  };
}

function crearPedido(token, datos) {
  exigirAdmin_(token);
  const cliente = String(datos.cliente || '').trim();
  const email = String(datos.email || '').trim();
  const telefono = String(datos.telefono || '').trim();
  const productos = String(datos.productos || '').trim();
  const total = Number(datos.total) || 0;
  const direccion = String(datos.direccion || '').trim();

  if (!cliente || !productos) throw new Error('El cliente y los productos son obligatorios.');
  if (!email && !telefono) throw new Error('Escribe el email o el teléfono (el cliente lo usa para consultar).');
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('El email no es válido.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = asegurarHojas_();
    const hoja = ss.getSheetByName(HOJA_PEDIDOS);
    const id = siguienteId_(hoja);
    const ahora = new Date();
    const estado = ESTADOS[0];

    const fila = [];
    fila[COL.ID - 1] = id;
    fila[COL.FECHA - 1] = ahora;
    fila[COL.CLIENTE - 1] = cliente;
    fila[COL.EMAIL - 1] = email;
    fila[COL.TELEFONO - 1] = telefono;
    fila[COL.PRODUCTOS - 1] = productos;
    fila[COL.TOTAL - 1] = total;
    fila[COL.DIRECCION - 1] = direccion;
    fila[COL.ESTADO - 1] = estado;
    fila[COL.TRANSPORTADORA - 1] = '';
    fila[COL.GUIA - 1] = '';
    fila[COL.NOTA - 1] = '';
    fila[COL.ACTUALIZADO - 1] = ahora;

    hoja.appendRow(fila);
    const correo = notificarCliente_(fila, estado);
    registrarHistorial_(ahora, id, 'Estado', '', estado, 'Pedido creado', correo);
    return { pedido: pedidoDesdeFila_(fila, ss.getSpreadsheetTimeZone()), correo: correo };
  } finally {
    lock.releaseLock();
  }
}

function actualizarPedido(token, id, cambios) {
  exigirAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = asegurarHojas_();
    const hoja = ss.getSheetByName(HOJA_PEDIDOS);
    const fila = buscarFila_(hoja, id);
    if (!fila) throw new Error('No se encontró el pedido ' + id);

    const v = hoja.getRange(fila, 1, 1, ENCABEZADOS_PEDIDOS.length).getValues()[0];
    const estadoAnterior = String(v[COL.ESTADO - 1]);
    const estado = String(cambios.estado || estadoAnterior);
    if (ESTADOS.indexOf(estado) === -1) throw new Error('Estado no válido.');

    const transportadora = String(cambios.transportadora || '').trim();
    const guia = String(cambios.guia || '').trim();
    const nota = String(cambios.nota || '').trim();
    const huboCambioDetalle = transportadora !== String(v[COL.TRANSPORTADORA - 1]) ||
      guia !== String(v[COL.GUIA - 1]) || nota !== String(v[COL.NOTA - 1]);

    v[COL.ESTADO - 1] = estado;
    v[COL.TRANSPORTADORA - 1] = transportadora;
    v[COL.GUIA - 1] = guia;
    v[COL.NOTA - 1] = nota;
    const ahora = new Date();
    v[COL.ACTUALIZADO - 1] = ahora;

    let correo = 'No aplica';
    if (estado !== estadoAnterior) {
      correo = notificarCliente_(v, estado);
      registrarHistorial_(ahora, id, 'Estado', estadoAnterior, estado, nota, correo);
    } else if (huboCambioDetalle) {
      registrarHistorial_(ahora, id, 'Actualización', estado, estado,
        [transportadora && 'Transportadora: ' + transportadora, guia && 'Guía: ' + guia, nota]
          .filter(Boolean).join(' · '), 'No aplica');
    } else {
      return { pedido: pedidoDesdeFila_(v, ss.getSpreadsheetTimeZone()), correo: 'Sin cambios' };
    }

    hoja.getRange(fila, 1, 1, ENCABEZADOS_PEDIDOS.length).setValues([v]);
    return { pedido: pedidoDesdeFila_(v, ss.getSpreadsheetTimeZone()), correo: correo };
  } finally {
    lock.releaseLock();
  }
}

function buscarFila_(hoja, id) {
  const ultima = hoja.getLastRow();
  if (ultima < 2) return 0;
  const ids = hoja.getRange(2, COL.ID, ultima - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim().toUpperCase() === String(id).trim().toUpperCase()) return i + 2;
  }
  return 0;
}

function siguienteId_(hoja) {
  let max = 0;
  const ultima = hoja.getLastRow();
  if (ultima > 1) {
    hoja.getRange(2, COL.ID, ultima - 1, 1).getValues().forEach(function (f) {
      const m = /^MA-(\d+)$/i.exec(String(f[0]).trim());
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return 'MA-' + ('0000' + (max + 1)).slice(-4);
}

function registrarHistorial_(fecha, pedido, tipo, anterior, nuevo, nota, correo) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_HISTORIAL);
  if (hoja) hoja.appendRow([fecha, pedido, tipo, anterior, nuevo, nota, correo]);
}

/* ============================== Correos ============================== */

function notificarCliente_(fila, estado) {
  const config = leerConfig_();
  if (config.NOTIFICAR_CLIENTE.toUpperCase() !== 'SI') return 'Desactivado';

  const email = String(fila[COL.EMAIL - 1] || '').trim();
  if (!email) return 'Sin email';

  const tienda = config.NOMBRE_TIENDA;
  const pedido = {
    id: String(fila[COL.ID - 1]),
    cliente: String(fila[COL.CLIENTE - 1]),
    productos: String(fila[COL.PRODUCTOS - 1]),
    transportadora: String(fila[COL.TRANSPORTADORA - 1] || ''),
    guia: String(fila[COL.GUIA - 1] || ''),
    nota: String(fila[COL.NOTA - 1] || '')
  };

  try {
    const opciones = {
      name: tienda,
      htmlBody: plantillaCorreo_(pedido, estado, tienda, obtenerUrlWeb_(), config.URL_RASTREO_TRANSPORTADORA)
    };
    if (config.EMAIL_RESPUESTA) opciones.replyTo = config.EMAIL_RESPUESTA;
    MailApp.sendEmail(email, tienda + ' · Pedido ' + pedido.id + ': ' + estado,
      MENSAJES_ESTADO[estado] || estado, opciones);
    return 'Enviado';
  } catch (err) {
    console.error('Error enviando correo del pedido ' + pedido.id + ': ' + err);
    return 'Error: ' + err.message;
  }
}

function esc_(texto) {
  return String(texto).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function plantillaCorreo_(p, estado, tienda, urlSeguimiento, urlRastreo) {
  const filas = [];
  if (p.transportadora) filas.push('<b>Transportadora:</b> ' + esc_(p.transportadora));
  if (p.guia) filas.push('<b>Guía:</b> ' + esc_(p.guia));
  if (p.nota) filas.push('<b>Nota:</b> ' + esc_(p.nota));

  let botones = '';
  if (urlSeguimiento) {
    botones += '<p><a href="' + esc_(urlSeguimiento) + '" style="background:#ec7fa9;color:#fff;' +
      'padding:10px 18px;border-radius:20px;text-decoration:none;display:inline-block">Ver seguimiento</a></p>';
  }
  if (/^https?:\/\//i.test(urlRastreo || '') && p.guia) {
    botones += '<p><a href="' + esc_(urlRastreo) + '">Rastrear con la transportadora</a></p>';
  }

  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;color:#444">' +
    '<h2 style="color:#ec7fa9;margin-bottom:4px">' + esc_(tienda) + '</h2>' +
    '<p>Hola ' + esc_(p.cliente) + ',</p>' +
    '<p>' + esc_(MENSAJES_ESTADO[estado] || estado) + '</p>' +
    '<div style="background:#fdf1f5;border-radius:12px;padding:14px 16px">' +
    '<p style="margin:0 0 6px"><b>Pedido:</b> ' + esc_(p.id) + '</p>' +
    '<p style="margin:0 0 6px"><b>Estado:</b> ' + esc_(estado) + '</p>' +
    '<p style="margin:0 0 6px"><b>Productos:</b> ' + esc_(p.productos) + '</p>' +
    filas.map(function (f) { return '<p style="margin:0 0 6px">' + f + '</p>'; }).join('') +
    '</div>' + botones +
    '<p style="font-size:12px;color:#999">Para consultar tu pedido necesitarás el número ' +
    esc_(p.id) + ' y tu email o teléfono.</p></div>';
}

/* ====================== Consulta pública (cliente) ====================== */

/**
 * Exige número de pedido + email o teléfono para que no se puedan
 * consultar pedidos ajenos.
 */
function buscarPedido(idIngresado, contacto) {
  const id = String(idIngresado || '').trim().toUpperCase();
  const dato = String(contacto || '').trim();
  const generico = { ok: false, mensaje: 'No encontramos un pedido con esos datos. Revisa el número y tu email o teléfono.' };
  if (!id || !dato) return { ok: false, mensaje: 'Escribe el número de pedido y tu email o teléfono.' };

  const cache = CacheService.getScriptCache();
  const claveIntentos = 'intentos_' + id;
  const intentos = Number(cache.get(claveIntentos) || 0);
  if (intentos >= MAX_INTENTOS_CONSULTA) {
    return { ok: false, mensaje: 'Demasiados intentos. Vuelve a probar en una hora o escríbenos.' };
  }
  cache.put(claveIntentos, String(intentos + 1), 3600);

  const ss = asegurarHojas_();
  const hoja = ss.getSheetByName(HOJA_PEDIDOS);
  if (hoja.getLastRow() < 2) return generico;

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, ENCABEZADOS_PEDIDOS.length).getValues();
  const fila = datos.find(function (f) { return String(f[COL.ID - 1]).trim().toUpperCase() === id; });
  if (!fila || !contactoCoincide_(fila, dato)) return generico;

  const tz = ss.getSpreadsheetTimeZone();
  const p = pedidoDesdeFila_(fila, tz);
  return {
    ok: true,
    pedido: {
      id: p.id,
      fecha: p.fecha,
      cliente: p.cliente.split(' ')[0],
      productos: p.productos,
      estado: p.estado,
      mensaje: MENSAJES_ESTADO[p.estado] || '',
      transportadora: p.transportadora,
      guia: p.guia,
      nota: p.nota,
      actualizado: p.actualizado,
      urlRastreo: leerConfig_().URL_RASTREO_TRANSPORTADORA
    }
  };
}

function contactoCoincide_(fila, contacto) {
  const c = contacto.toLowerCase();
  if (c.indexOf('@') !== -1) {
    return c === String(fila[COL.EMAIL - 1]).trim().toLowerCase();
  }
  const digitos = c.replace(/\D/g, '');
  const telefono = String(fila[COL.TELEFONO - 1]).replace(/\D/g, '');
  return digitos.length >= 6 && telefono.length >= 6 && telefono.slice(-digitos.length) === digitos;
}
