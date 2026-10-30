/**
 * Seguimiento de pedidos – Mis Angelitos Baby
 *
 * Hoja de cálculo = base de datos. Cada vez que cambias el "Estado" de un pedido:
 *   1. se registra en la hoja Historial,
 *   2. se avisa por correo al cliente,
 *   3. el cliente ve el avance en la página pública de seguimiento (Web App).
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

const COLORES_ESTADO = {
  'Recibido': '#e8eaf6',
  'Pago confirmado': '#e0f2f1',
  'En preparación': '#fff8e1',
  'Enviado': '#e3f2fd',
  'En tránsito': '#e1f5fe',
  'Entregado': '#e8f5e9',
  'Cancelado': '#ffebee'
};

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
  DIRECCION: 8, ESTADO: 9, TRANSPORTADORA: 10, GUIA: 11, NOTA: 12,
  ACTUALIZADO: 13, NOTIFICADO: 14
};

const ENCABEZADOS_PEDIDOS = [
  'ID Pedido', 'Fecha', 'Cliente', 'Email', 'Teléfono', 'Productos', 'Total',
  'Dirección / Ciudad', 'Estado', 'Transportadora', 'Guía', 'Nota para el cliente',
  'Última actualización', 'Estado notificado (automático)'
];

const ENCABEZADOS_HISTORIAL = ['Fecha', 'Pedido', 'Tipo', 'Estado anterior', 'Estado nuevo', 'Nota', 'Correo'];

const CONFIG_POR_DEFECTO = [
  ['NOMBRE_TIENDA', 'Mis Angelitos Baby'],
  ['NOTIFICAR_CLIENTE', 'SI'],
  ['EMAIL_RESPUESTA', ''],
  ['URL_RASTREO_TRANSPORTADORA', '']
];

const MAX_INTENTOS_CONSULTA = 8; // por pedido y por hora

/* ============================== Menú ============================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Pedidos')
    .addItem('➕ Nuevo pedido', 'mostrarNuevoPedido')
    .addItem('🔗 Ver enlace de seguimiento', 'mostrarEnlaceSeguimiento')
    .addSeparator()
    .addItem('⚙️ Configuración inicial', 'configuracionInicial')
    .addToUi();
}

function mostrarNuevoPedido() {
  const html = HtmlService.createHtmlOutput(HTML_NUEVO_PEDIDO).setWidth(420).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nuevo pedido');
}

function mostrarEnlaceSeguimiento() {
  const url = obtenerUrlSeguimiento_();
  const ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Aún no has publicado la Web App.\n\nVe a Implementar > Nueva implementación > Aplicación web ' +
      '(Ejecutar como: yo, Acceso: cualquier persona) y vuelve a abrir esta opción.');
    return;
  }
  ui.alert('Enlace de seguimiento para tus clientes:\n\n' + url);
}

/* ======================== Configuración inicial ======================== */

function configuracionInicial() {
  const ss = SpreadsheetApp.getActive();

  // Pedidos
  const pedidos = ss.getSheetByName(HOJA_PEDIDOS) || ss.insertSheet(HOJA_PEDIDOS);
  pedidos.getRange(1, 1, 1, ENCABEZADOS_PEDIDOS.length).setValues([ENCABEZADOS_PEDIDOS])
    .setFontWeight('bold').setBackground('#f8bbd0').setWrap(true);
  pedidos.setFrozenRows(1);
  pedidos.getRange(2, COL.FECHA, 1000, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  pedidos.getRange(2, COL.ACTUALIZADO, 1000, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  pedidos.getRange(2, COL.TOTAL, 1000, 1).setNumberFormat('$#,##0');
  pedidos.getRange(2, COL.NOTIFICADO, 1000, 1).setFontColor('#9e9e9e');
  pedidos.setColumnWidth(COL.PRODUCTOS, 260);
  pedidos.setColumnWidth(COL.DIRECCION, 220);
  pedidos.setColumnWidth(COL.NOTA, 240);

  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(ESTADOS, true).setAllowInvalid(false).build();
  pedidos.getRange(2, COL.ESTADO, 1000, 1).setDataValidation(regla);

  const rangoEstado = pedidos.getRange(2, COL.ESTADO, 1000, 1);
  pedidos.setConditionalFormatRules(ESTADOS.map(function (estado) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(estado).setBackground(COLORES_ESTADO[estado])
      .setRanges([rangoEstado]).build();
  }));

  // Historial
  const historial = ss.getSheetByName(HOJA_HISTORIAL) || ss.insertSheet(HOJA_HISTORIAL);
  historial.getRange(1, 1, 1, ENCABEZADOS_HISTORIAL.length).setValues([ENCABEZADOS_HISTORIAL])
    .setFontWeight('bold').setBackground('#f8bbd0');
  historial.setFrozenRows(1);
  historial.getRange(2, 1, 5000, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // Config
  let config = ss.getSheetByName(HOJA_CONFIG);
  if (!config) {
    config = ss.insertSheet(HOJA_CONFIG);
    config.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]).setFontWeight('bold').setBackground('#f8bbd0');
    config.getRange(2, 1, CONFIG_POR_DEFECTO.length, 2).setValues(CONFIG_POR_DEFECTO);
    config.setColumnWidth(1, 240);
    config.setColumnWidth(2, 360);
  }

  // Disparador instalable (el onEdit simple no puede enviar correos).
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'alEditar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('alEditar').forSpreadsheet(ss).onEdit().create();

  SpreadsheetApp.getUi().alert('Listo ✅\n\nSe crearon las hojas Pedidos, Historial y Config, ' +
    'y se activó el aviso automático al cambiar el Estado de un pedido.');
}

/* =========================== Configuración =========================== */

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

function obtenerUrlSeguimiento_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (err) {
    return '';
  }
}

/* ======================== Edición de la hoja ======================== */

/** Disparador instalable onEdit. */
function alEditar(e) {
  if (!e || !e.range) return;
  const hoja = e.range.getSheet();
  if (hoja.getName() !== HOJA_PEDIDOS) return;

  const primeraFila = Math.max(2, e.range.getRow());
  const ultimaFila = e.range.getRow() + e.range.getNumRows() - 1;
  if (ultimaFila < 2) return;

  const c0 = e.range.getColumn();
  const c1 = c0 + e.range.getNumColumns() - 1;
  const toca = function (col) { return c0 <= col && c1 >= col; };
  if (!toca(COL.ESTADO) && !toca(COL.GUIA) && !toca(COL.NOTA) && !toca(COL.TRANSPORTADORA)) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    for (let fila = primeraFila; fila <= ultimaFila; fila++) {
      procesarFila_(hoja, fila, toca(COL.ESTADO));
    }
  } finally {
    lock.releaseLock();
  }
}

function procesarFila_(hoja, fila, cambioEstado) {
  const v = hoja.getRange(fila, 1, 1, ENCABEZADOS_PEDIDOS.length).getValues()[0];
  const id = String(v[COL.ID - 1]).trim();
  if (!id) return;

  const estado = String(v[COL.ESTADO - 1]).trim();
  const notificado = String(v[COL.NOTIFICADO - 1]).trim();
  const ahora = new Date();

  if (cambioEstado && estado && estado !== notificado) {
    const resultadoCorreo = notificarCliente_(v, estado);
    registrarHistorial_(ahora, id, 'Estado', notificado, estado, String(v[COL.NOTA - 1]), resultadoCorreo);
    hoja.getRange(fila, COL.NOTIFICADO).setValue(estado);
  } else {
    registrarHistorial_(ahora, id, 'Actualización', estado, estado,
      armarDetalleActualizacion_(v), 'No aplica');
  }
  hoja.getRange(fila, COL.ACTUALIZADO).setValue(ahora);
}

function armarDetalleActualizacion_(v) {
  const partes = [];
  const transportadora = String(v[COL.TRANSPORTADORA - 1]).trim();
  const guia = String(v[COL.GUIA - 1]).trim();
  const nota = String(v[COL.NOTA - 1]).trim();
  if (transportadora) partes.push('Transportadora: ' + transportadora);
  if (guia) partes.push('Guía: ' + guia);
  if (nota) partes.push(nota);
  return partes.join(' · ');
}

function registrarHistorial_(fecha, pedido, tipo, anterior, nuevo, nota, correo) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_HISTORIAL);
  if (!hoja) return;
  hoja.appendRow([fecha, pedido, tipo, anterior, nuevo, nota, correo]);
}

/* ============================ Nuevo pedido ============================ */

/** Llamado desde el diálogo NuevoPedido.html. */
function crearPedido(datos) {
  const cliente = String(datos.cliente || '').trim();
  const email = String(datos.email || '').trim();
  const telefono = String(datos.telefono || '').trim();
  const productos = String(datos.productos || '').trim();
  const total = Number(datos.total) || 0;
  const direccion = String(datos.direccion || '').trim();

  if (!cliente || !productos) throw new Error('El cliente y los productos son obligatorios.');
  if (!email && !telefono) throw new Error('Indica el email o el teléfono del cliente (se usa para el seguimiento).');
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('El email no es válido.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_PEDIDOS);
    if (!hoja) throw new Error('Primero ejecuta ⚙️ Configuración inicial.');

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
    fila[COL.NOTIFICADO - 1] = estado;

    hoja.appendRow(fila);
    const correo = notificarCliente_(fila, estado);
    registrarHistorial_(ahora, id, 'Estado', '', estado, 'Pedido creado', correo);
    return id;
  } finally {
    lock.releaseLock();
  }
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

/* ============================== Correos ============================== */

function notificarCliente_(fila, estado) {
  const config = leerConfig_();
  if (config.NOTIFICAR_CLIENTE.toUpperCase() !== 'SI') return 'Desactivado';

  const email = String(fila[COL.EMAIL - 1]).trim();
  if (!email) return 'Sin email';

  const tienda = config.NOMBRE_TIENDA;
  const pedido = {
    id: String(fila[COL.ID - 1]),
    cliente: String(fila[COL.CLIENTE - 1]),
    productos: String(fila[COL.PRODUCTOS - 1]),
    transportadora: String(fila[COL.TRANSPORTADORA - 1]),
    guia: String(fila[COL.GUIA - 1]),
    nota: String(fila[COL.NOTA - 1])
  };

  try {
    const opciones = {
      name: tienda,
      htmlBody: plantillaCorreo_(pedido, estado, tienda, obtenerUrlSeguimiento_(), config.URL_RASTREO_TRANSPORTADORA)
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
  if (urlRastreo && p.guia) {
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

/* ========================= Web App (cliente) ========================= */

function doGet() {
  const tienda = leerConfig_().NOMBRE_TIENDA;
  const plantilla = HtmlService.createTemplate(HTML_SEGUIMIENTO);
  plantilla.tienda = tienda;
  plantilla.estados = ESTADOS.filter(function (e) { return e !== 'Cancelado'; });
  return plantilla.evaluate()
    .setTitle('Seguimiento de pedido · ' + tienda)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Llamado desde la página pública. Exige número de pedido + email o teléfono
 * para que no se puedan consultar pedidos ajenos.
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

  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(HOJA_PEDIDOS);
  if (!hoja || hoja.getLastRow() < 2) return generico;

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, ENCABEZADOS_PEDIDOS.length).getValues();
  const fila = datos.find(function (f) { return String(f[COL.ID - 1]).trim().toUpperCase() === id; });
  if (!fila || !contactoCoincide_(fila, dato)) return generico;

  const tz = ss.getSpreadsheetTimeZone();
  const formato = function (d) {
    return d instanceof Date ? Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm') : '';
  };

  return {
    ok: true,
    pedido: {
      id: String(fila[COL.ID - 1]),
      fecha: formato(fila[COL.FECHA - 1]),
      cliente: String(fila[COL.CLIENTE - 1]).split(' ')[0],
      productos: String(fila[COL.PRODUCTOS - 1]),
      estado: String(fila[COL.ESTADO - 1]),
      mensaje: MENSAJES_ESTADO[String(fila[COL.ESTADO - 1])] || '',
      transportadora: String(fila[COL.TRANSPORTADORA - 1]),
      guia: String(fila[COL.GUIA - 1]),
      nota: String(fila[COL.NOTA - 1]),
      actualizado: formato(fila[COL.ACTUALIZADO - 1]),
      urlRastreo: leerConfig_().URL_RASTREO_TRANSPORTADORA
    },
    historial: historialPublico_(ss, String(fila[COL.ID - 1]), formato)
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

function historialPublico_(ss, id, formato) {
  const hoja = ss.getSheetByName(HOJA_HISTORIAL);
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, ENCABEZADOS_HISTORIAL.length).getValues()
    .filter(function (f) { return String(f[1]).trim().toUpperCase() === id.toUpperCase() && f[2] === 'Estado'; })
    .map(function (f) { return { fecha: formato(f[0]), estado: String(f[4]), nota: String(f[5]) }; })
    .reverse();
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

const HTML_NUEVO_PEDIDO = String.raw`<!DOCTYPE html>
<html lang="es">
<head>
  <base target="_top">
  <meta charset="utf-8">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <style> .btn-rosa { background:#ec7fa9; color:#fff; } .btn-rosa:hover { background:#d96a96; color:#fff; } </style>
</head>
<body class="p-3">
<form id="f">
  <div class="mb-2"><label class="form-label mb-0">Cliente *</label><input name="cliente" class="form-control form-control-sm" required></div>
  <div class="row g-2 mb-2">
    <div class="col"><label class="form-label mb-0">Email</label><input name="email" type="email" class="form-control form-control-sm"></div>
    <div class="col"><label class="form-label mb-0">Teléfono</label><input name="telefono" class="form-control form-control-sm"></div>
  </div>
  <div class="mb-2"><label class="form-label mb-0">Productos *</label><textarea name="productos" rows="3" class="form-control form-control-sm" required></textarea></div>
  <div class="mb-2"><label class="form-label mb-0">Total</label><input name="total" type="number" min="0" step="any" class="form-control form-control-sm"></div>
  <div class="mb-3"><label class="form-label mb-0">Dirección / ciudad</label><input name="direccion" class="form-control form-control-sm"></div>
  <div id="msg" class="alert alert-danger py-1 d-none"></div>
  <button id="b" class="btn btn-rosa w-100" type="submit">Crear pedido</button>
</form>
<script>
  document.getElementById('f').addEventListener('submit', function (ev) {
    ev.preventDefault();
    const b = document.getElementById('b'), msg = document.getElementById('msg');
    const datos = {};
    new FormData(ev.target).forEach(function (v, k) { datos[k] = v; });
    b.disabled = true; b.textContent = 'Guardando…'; msg.classList.add('d-none');
    google.script.run
      .withSuccessHandler(function (id) { alert('Pedido ' + id + ' creado ✅'); google.script.host.close(); })
      .withFailureHandler(function (e) { b.disabled = false; b.textContent = 'Crear pedido'; msg.textContent = e.message; msg.classList.remove('d-none'); })
      .crearPedido(datos);
  });
</script>
</body>
</html>
`;
