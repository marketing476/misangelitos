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

/* ===================== Páginas HTML (incluidas en este mismo archivo) ===================== */

const HTML_SEGUIMIENTO = String.raw`<!DOCTYPE html>
<html lang="es">
<head>
  <base target="_top">
  <meta charset="utf-8">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    :root { --rosa: #ec7fa9; --rosa-claro: #fdf1f5; }
    body { background: var(--rosa-claro); font-family: system-ui, sans-serif; }
    .tarjeta { max-width: 560px; margin: 24px auto; border: 0; border-radius: 20px; box-shadow: 0 6px 24px rgba(236,127,169,.18); }
    h1 { color: var(--rosa); font-size: 1.5rem; }
    .btn-rosa { background: var(--rosa); color: #fff; border-radius: 20px; }
    .btn-rosa:hover { background: #d96a96; color: #fff; }
    .paso { display: flex; gap: 12px; padding-bottom: 18px; position: relative; }
    .paso:not(:last-child)::before { content: ''; position: absolute; left: 11px; top: 24px; bottom: 0; width: 2px; background: #e0e0e0; }
    .punto { width: 24px; height: 24px; border-radius: 50%; background: #e0e0e0; flex: none; }
    .paso.hecho .punto { background: var(--rosa); }
    .paso.hecho:not(:last-child)::before { background: var(--rosa); }
    .paso.actual .punto { box-shadow: 0 0 0 5px rgba(236,127,169,.3); }
    .paso:not(.hecho) { color: #aaa; }
  </style>
</head>
<body>
<div class="container px-3">
  <div class="card tarjeta"><div class="card-body p-4">
    <h1 class="text-center mb-1"><?= tienda ?></h1>
    <p class="text-center text-muted">Sigue tu pedido</p>

    <form id="formulario" class="mb-3">
      <div class="mb-3">
        <label class="form-label" for="pedido">Número de pedido</label>
        <input id="pedido" class="form-control" placeholder="MA-0001" required autocomplete="off">
      </div>
      <div class="mb-3">
        <label class="form-label" for="contacto">Tu email o teléfono</label>
        <input id="contacto" class="form-control" required autocomplete="off">
      </div>
      <button id="boton" class="btn btn-rosa w-100" type="submit">Consultar</button>
    </form>

    <div id="error" class="alert alert-warning d-none"></div>

    <div id="resultado" class="d-none">
      <hr>
      <p class="mb-1">Hola <b id="r-cliente"></b> 👋</p>
      <p id="r-mensaje" class="text-muted"></p>
      <div id="r-linea" class="my-3"></div>
      <div class="small">
        <div><b>Pedido:</b> <span id="r-id"></span></div>
        <div><b>Productos:</b> <span id="r-productos"></span></div>
        <div id="r-envio" class="d-none"><b>Envío:</b> <span id="r-transportadora"></span> <span id="r-guia"></span>
          <a id="r-rastreo" class="d-none" target="_blank" rel="noopener">Rastrear</a></div>
        <div id="r-nota-fila" class="d-none"><b>Nota:</b> <span id="r-nota"></span></div>
        <div class="text-muted mt-2">Última actualización: <span id="r-actualizado"></span></div>
      </div>
    </div>
  </div></div>
</div>

<script>
  const ESTADOS = <?!= JSON.stringify(estados) ?>;
  const $ = function (id) { return document.getElementById(id); };

  $('formulario').addEventListener('submit', function (ev) {
    ev.preventDefault();
    $('boton').disabled = true;
    $('boton').textContent = 'Buscando…';
    $('error').classList.add('d-none');
    $('resultado').classList.add('d-none');
    google.script.run
      .withSuccessHandler(mostrar)
      .withFailureHandler(function () { mostrar({ ok: false, mensaje: 'Ocurrió un error. Inténtalo de nuevo.' }); })
      .buscarPedido($('pedido').value, $('contacto').value);
  });

  function mostrar(r) {
    $('boton').disabled = false;
    $('boton').textContent = 'Consultar';
    if (!r.ok) {
      $('error').textContent = r.mensaje;
      $('error').classList.remove('d-none');
      return;
    }
    const p = r.pedido;
    $('r-cliente').textContent = p.cliente;
    $('r-mensaje').textContent = p.mensaje;
    $('r-id').textContent = p.id;
    $('r-productos').textContent = p.productos;
    $('r-actualizado').textContent = p.actualizado;

    const hayEnvio = p.transportadora || p.guia;
    $('r-envio').classList.toggle('d-none', !hayEnvio);
    $('r-transportadora').textContent = p.transportadora;
    $('r-guia').textContent = p.guia ? '· Guía ' + p.guia : '';
    const enlace = $('r-rastreo');
    if (p.urlRastreo && /^https?:\/\//i.test(p.urlRastreo) && p.guia) {
      enlace.href = p.urlRastreo;
      enlace.classList.remove('d-none');
    } else {
      enlace.classList.add('d-none');
    }
    $('r-nota').textContent = p.nota;
    $('r-nota-fila').classList.toggle('d-none', !p.nota);

    linea(p.estado);
    $('resultado').classList.remove('d-none');
  }

  function linea(estadoActual) {
    const cont = $('r-linea');
    cont.textContent = '';
    if (estadoActual === 'Cancelado') {
      const aviso = document.createElement('div');
      aviso.className = 'alert alert-danger mb-0';
      aviso.textContent = 'Pedido cancelado';
      cont.appendChild(aviso);
      return;
    }
    const idx = ESTADOS.indexOf(estadoActual);
    ESTADOS.forEach(function (estado, i) {
      const paso = document.createElement('div');
      paso.className = 'paso' + (i <= idx ? ' hecho' : '') + (i === idx ? ' actual' : '');
      const punto = document.createElement('div');
      punto.className = 'punto';
      const texto = document.createElement('div');
      texto.textContent = estado;
      if (i === idx) texto.style.fontWeight = '600';
      paso.appendChild(punto);
      paso.appendChild(texto);
      cont.appendChild(paso);
    });
  }
</script>
</body>
</html>
`;

const HTML_ADMIN = String.raw`<!DOCTYPE html>
<html lang="es">
<head>
  <base target="_top">
  <meta charset="utf-8">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    :root { --rosa: #ec7fa9; --rosa-claro: #fdf1f5; }
    body { background: var(--rosa-claro); font-family: system-ui, sans-serif; }
    h1, h2 { color: var(--rosa); }
    .btn-rosa { background: var(--rosa); color: #fff; border-radius: 20px; }
    .btn-rosa:hover { background: #d96a96; color: #fff; }
    .caja { border: 0; border-radius: 16px; box-shadow: 0 4px 16px rgba(236,127,169,.15); }
    .badge-estado { background: var(--rosa); }
    .badge-Entregado { background: #43a047; } .badge-Cancelado { background: #e53935; }
  </style>
</head>
<body>
<div class="container py-4" style="max-width: 820px">
  <h1 class="h3 mb-3">Panel de pedidos · <?= tienda ?></h1>

  <!-- Acceso -->
  <div id="vista-login" class="card caja d-none" style="max-width:380px"><div class="card-body">
    <h2 id="login-titulo" class="h5"></h2>
    <p id="login-ayuda" class="small text-muted"></p>
    <form id="form-login">
      <input id="clave" type="password" class="form-control mb-2" placeholder="Clave" required autocomplete="current-password">
      <input id="clave2" type="password" class="form-control mb-2 d-none" placeholder="Repite la clave" autocomplete="new-password">
      <div id="login-error" class="alert alert-danger py-1 d-none"></div>
      <button class="btn btn-rosa w-100" type="submit">Entrar</button>
    </form>
  </div></div>

  <!-- Panel -->
  <div id="vista-panel" class="d-none">
    <div class="text-end mb-2"><button id="salir" class="btn btn-sm btn-outline-secondary">Salir</button></div>

    <div class="card caja mb-4"><div class="card-body">
      <h2 class="h5">Nuevo pedido</h2>
      <form id="form-nuevo">
        <div class="mb-2"><label class="form-label mb-0 small">Cliente *</label><input name="cliente" class="form-control" required></div>
        <div class="row g-2 mb-2">
          <div class="col-sm"><label class="form-label mb-0 small">Email</label><input name="email" type="email" class="form-control"></div>
          <div class="col-sm"><label class="form-label mb-0 small">Teléfono</label><input name="telefono" class="form-control"></div>
        </div>
        <div class="mb-2"><label class="form-label mb-0 small">Productos *</label><textarea name="productos" rows="3" class="form-control" required></textarea></div>
        <div class="row g-2 mb-3">
          <div class="col-sm-4"><label class="form-label mb-0 small">Total</label><input name="total" type="number" min="0" step="any" class="form-control"></div>
          <div class="col-sm"><label class="form-label mb-0 small">Dirección / ciudad</label><input name="direccion" class="form-control"></div>
        </div>
        <div id="nuevo-msg" class="alert py-1 d-none"></div>
        <button id="btn-nuevo" class="btn btn-rosa" type="submit">Crear pedido</button>
      </form>
    </div></div>

    <div class="d-flex gap-2 mb-3">
      <input id="buscar" class="form-control" placeholder="Buscar por número, cliente o producto">
      <select id="filtro" class="form-select" style="max-width:190px"></select>
    </div>
    <div id="lista"></div>
  </div>
</div>

<script>
  const ESTADOS = <?!= JSON.stringify(estados) ?>;
  const $ = function (id) { return document.getElementById(id); };
  let token = null;
  let pedidos = [];
  let creandoClave = false;

  try { token = localStorage.getItem('adm_token'); } catch (e) { token = null; }

  function guardarToken(t) {
    token = t;
    try { if (t) localStorage.setItem('adm_token', t); else localStorage.removeItem('adm_token'); } catch (e) {}
  }

  function llamar(nombre, args, ok, fallo) {
    const r = google.script.run.withSuccessHandler(ok).withFailureHandler(function (e) {
      if (/SESION/.test(e.message)) { salir(); return; }
      if (fallo) fallo(e);
    });
    r[nombre].apply(r, args);
  }

  function mostrarLogin() {
    $('vista-panel').classList.add('d-none');
    $('vista-login').classList.remove('d-none');
    llamar('estadoPanel', [], function (s) {
      creandoClave = !s.claveCreada;
      $('login-titulo').textContent = creandoClave ? 'Crea tu clave' : 'Entra a tu panel';
      $('login-ayuda').textContent = creandoClave
        ? 'Es la primera vez. Elige una clave de al menos 8 caracteres; la usarás siempre para entrar.' : '';
      $('clave2').classList.toggle('d-none', !creandoClave);
      $('clave2').required = creandoClave;
    });
  }

  function salir() { guardarToken(null); mostrarLogin(); }
  $('salir').addEventListener('click', salir);

  $('form-login').addEventListener('submit', function (ev) {
    ev.preventDefault();
    const err = $('login-error');
    err.classList.add('d-none');
    if (creandoClave && $('clave').value !== $('clave2').value) {
      err.textContent = 'Las dos claves no coinciden.';
      err.classList.remove('d-none');
      return;
    }
    const alEntrar = function (t) { guardarToken(t); $('clave').value = ''; $('clave2').value = ''; abrirPanel(); };
    const alFallar = function (e) { err.textContent = e.message.replace(/^Error: /, ''); err.classList.remove('d-none'); };
    llamar(creandoClave ? 'crearClave' : 'iniciarSesion', [$('clave').value], alEntrar, alFallar);
  });

  function abrirPanel() {
    $('vista-login').classList.add('d-none');
    $('vista-panel').classList.remove('d-none');
    cargar();
  }

  function cargar() {
    llamar('listarPedidos', [token], function (lista) { pedidos = lista; dibujar(); });
  }

  /* ---------- Nuevo pedido ---------- */
  $('form-nuevo').addEventListener('submit', function (ev) {
    ev.preventDefault();
    const datos = {};
    new FormData(ev.target).forEach(function (v, k) { datos[k] = v; });
    const boton = $('btn-nuevo'), msg = $('nuevo-msg');
    boton.disabled = true; boton.textContent = 'Guardando…'; msg.classList.add('d-none');
    llamar('crearPedido', [token, datos], function (r) {
      boton.disabled = false; boton.textContent = 'Crear pedido';
      ev.target.reset();
      msg.className = 'alert alert-success py-1';
      msg.textContent = 'Pedido ' + r.pedido.id + ' creado. Correo: ' + r.correo;
      pedidos.unshift(r.pedido);
      dibujar();
    }, function (e) {
      boton.disabled = false; boton.textContent = 'Crear pedido';
      msg.className = 'alert alert-danger py-1';
      msg.textContent = e.message.replace(/^Error: /, '');
    });
  });

  /* ---------- Lista ---------- */
  const filtro = $('filtro');
  [['', 'Todos los estados']].concat(ESTADOS.map(function (e) { return [e, e]; })).forEach(function (par) {
    const o = document.createElement('option');
    o.value = par[0]; o.textContent = par[1];
    filtro.appendChild(o);
  });
  filtro.addEventListener('change', dibujar);
  $('buscar').addEventListener('input', dibujar);

  function el(tag, clase, texto) {
    const e = document.createElement(tag);
    if (clase) e.className = clase;
    if (texto !== undefined) e.textContent = texto;
    return e;
  }

  function dibujar() {
    const q = $('buscar').value.trim().toLowerCase();
    const est = filtro.value;
    const lista = $('lista');
    lista.textContent = '';
    const visibles = pedidos.filter(function (p) {
      if (est && p.estado !== est) return false;
      return !q || (p.id + ' ' + p.cliente + ' ' + p.productos).toLowerCase().indexOf(q) !== -1;
    });
    if (!visibles.length) {
      lista.appendChild(el('p', 'text-muted text-center', 'No hay pedidos para mostrar.'));
      return;
    }
    visibles.forEach(function (p) { lista.appendChild(tarjeta(p)); });
  }

  function campo(etiqueta, valor, clase) {
    const col = el('div', clase);
    col.appendChild(el('label', 'form-label mb-0 small', etiqueta));
    const input = el('input', 'form-control form-control-sm');
    input.value = valor;
    col.appendChild(input);
    return { col: col, input: input };
  }

  function tarjeta(p) {
    const card = el('div', 'card caja mb-3');
    const body = el('div', 'card-body');
    card.appendChild(body);

    const cab = el('div', 'd-flex justify-content-between align-items-start mb-1');
    const izq = el('div');
    izq.appendChild(el('b', '', p.id + ' · ' + p.cliente));
    izq.appendChild(el('div', 'small text-muted', p.fecha + (p.total ? ' · $' + p.total.toLocaleString('es-CO') : '')));
    cab.appendChild(izq);
    cab.appendChild(el('span', 'badge badge-estado badge-' + p.estado.replace(/\s/g, ''), p.estado));
    body.appendChild(cab);

    body.appendChild(el('div', 'small', p.productos));
    const contacto = [p.email, p.telefono, p.direccion].filter(Boolean).join(' · ');
    if (contacto) body.appendChild(el('div', 'small text-muted mb-2', contacto));

    const fila = el('div', 'row g-2 mt-1');
    const colEstado = el('div', 'col-sm-4');
    colEstado.appendChild(el('label', 'form-label mb-0 small', 'Estado'));
    const select = el('select', 'form-select form-select-sm');
    ESTADOS.forEach(function (e) {
      const o = el('option', '', e); o.value = e; if (e === p.estado) o.selected = true; select.appendChild(o);
    });
    colEstado.appendChild(select);
    const transp = campo('Transportadora', p.transportadora, 'col-sm-4');
    const guia = campo('Guía', p.guia, 'col-sm-4');
    fila.appendChild(colEstado); fila.appendChild(transp.col); fila.appendChild(guia.col);
    body.appendChild(fila);

    const nota = campo('Nota para el cliente (la verá en su seguimiento)', p.nota, 'mt-2');
    body.appendChild(nota.col);

    const pie = el('div', 'd-flex align-items-center gap-2 mt-2');
    const boton = el('button', 'btn btn-rosa btn-sm', 'Guardar cambios');
    const msg = el('span', 'small text-muted');
    pie.appendChild(boton); pie.appendChild(msg);
    body.appendChild(pie);

    boton.addEventListener('click', function () {
      boton.disabled = true; msg.textContent = 'Guardando…';
      llamar('actualizarPedido', [token, p.id, {
        estado: select.value, transportadora: transp.input.value, guia: guia.input.value, nota: nota.input.value
      }], function (r) {
        const i = pedidos.findIndex(function (x) { return x.id === p.id; });
        if (i !== -1) pedidos[i] = r.pedido;
        dibujar();
      }, function (e) {
        boton.disabled = false; msg.textContent = e.message.replace(/^Error: /, '');
      });
    });
    return card;
  }

  if (token) abrirPanel(); else mostrarLogin();
</script>
</body>
</html>
`;
