/**
 * BACKEND para "Mis Gastos - Personal" + sync de Ventas / Stocks / Campañas
 * ---------------------------------------------------------
 * ⚠️ ESPEJO, NO LA FUENTE DE VERDAD: este archivo es una copia de referencia
 * del código real, que vive en el editor de Apps Script del Sheet "Timeless -
 * Ventas e Inventario" (Extensiones → Apps Script) y es compartido con el
 * proyecto del dashboard de Timeless (otro chat/carpeta). Si algo no cuadra
 * con lo que ves acá, CONFÍA en lo que está deployado ahí, no en este archivo
 * — cópialo de vuelta a mano si hace falta. (Pasó una vez: este archivo tenía
 * solo la parte de gastos, sin Ventas/Stocks/Campañas/Cashback/Compras/
 * Seguimiento/Instagram que ya existían en el deploy real.)
 *
 * Recibe cada gasto que envía la app y lo agrega como una fila nueva (o
 * ACTUALIZA la fila si la ID ya existe, para que las ediciones hechas en la
 * app —fecha, monto, categoría, nota— se reflejen en Sheets) en la pestaña
 * "Gastos".
 *
 * También recibe varios tipos de sync (full-replace de una pestaña) y CRUD
 * puntual para otras secciones del dashboard:
 *   - ventasSync    → pestaña "VentasDetalle"  (sync-ventas.ps1 leyendo "Resultados ventas 2026.xlsx")
 *   - stocksSync    → pestaña "Stocks"         (sync-ventas.ps1 leyendo "Venta accs.xlsx")
 *   - campanasSync  → pestaña "Campanas"       (gasto real de Meta Ads por día y campaña)
 *   - cashbackSync  → pestaña "Cashback"       (retiros de cashback, desde esta misma app)
 *   - compraGuardar / compraEliminar / compraFoto → pestaña "Compras" (planificación de accesorios)
 *   - seguimientoGuardar / seguimientoEliminar     → pestaña "Seguimiento" (17TRACK)
 *
 * Los sync masivos (Ventas/Stocks/Campañas/Cashback) son full-replace porque
 * el Excel / Meta / la app son la fuente de verdad: así nunca hay filas
 * duplicadas ni desfasadas entre un sync y otro.
 */

var SHEET_NAME = 'Gastos';

var VENTAS_SHEET_NAME = 'VentasDetalle';
var VENTAS_HEADERS = ['Fecha', 'Producto', 'Venta', 'Utilidad'];

var STOCKS_SHEET_NAME = 'Stocks';
// "StockInvertido" al FINAL a propósito (no correr las columnas de arriba):
// costo unitario EXACTO por bloque × stock que le queda a cada bloque,
// sumado bloque por bloque (no un promedio) — igual al criterio de la
// columna "Stock invertido" de tu pestaña resumen Stocks del Excel. Antes el
// dashboard promediaba (Invertido total / Cantidad pedido total) × Stock, que
// da un número ligeramente distinto si un bloque viejo costaba menos/más que
// el nuevo.
var STOCKS_HEADERS = ['Producto', 'Precio', 'Vendidos', 'Stock', 'Ganancia bruta posible', 'Ganancia neta posible', 'Invertido', 'Cantidad pedido', 'Fecha pedido', 'Plataforma', 'StockInvertido'];

// OJO: la pestaña real se llama "Campanas" SIN ñ (Google la nombró así al
// importar el CSV "Campanas.csv"). Debe coincidir para no crear una pestaña
// nueva y perder el gid publicado que lee el dashboard.
var CAMPANAS_SHEET_NAME = 'Campanas';
var CAMPANAS_HEADERS = ['Fecha', 'Campaña', 'Gasto'];

// Gasto y conversaciones POR ANUNCIO (no por campaña) — para el semáforo de
// costo por conversación del dashboard. Se llena sola desde syncMetaAdsAutomatico()
// junto con Campanas; se lee en vivo por ?action=anunciosMeta (no necesita
// "Publicar en la web" como las demás pestañas de sync masivo).
var ANUNCIOS_META_SHEET_NAME = 'AnunciosMeta';
var ANUNCIOS_META_HEADERS = ['Fecha', 'Anuncio', 'Campaña', 'Gasto', 'Conversaciones'];

// Pestaña "Compras" — planificación de accesorios para traer (bloques de
// compra). A diferencia de Ventas/Stocks/Campañas (que son sync masivo desde
// Excel/Meta), esta se edita directo desde el dashboard: Alberto crea, edita
// y borra items uno por uno desde su celular o PC.
var COMPRAS_SHEET_NAME = 'Compras';
var COMPRAS_HEADERS = ['ID', 'Nombre', 'Estado', 'FechaInicio', 'FechaFin', 'PrecioTotal', 'ProductosJSON', 'FotoURL', 'Notas', 'CreadoEn'];

// Pestaña "Seguimiento" — un paquete por fila (número de tracking + los
// productos que vienen dentro + la tienda). Alberto la llena desde el
// dashboard; el estado lo rellena solo sync17Track() con la API de 17TRACK.
var SEGUIMIENTO_SHEET_NAME = 'Seguimiento';
// "Link" y "Proveedor" van al FINAL a propósito: sync17Track escribe en
// columnas fijas por número (Estado=6 … Registrado=11), así que agregar
// columnas al final no corre esos índices. "Proveedor" es qué página arma el
// botón ↗ (parcelsapp / mailamericas / 17track / personalizado=usa "Link").
var SEGUIMIENTO_HEADERS = ['ID', 'Tracking', 'Productos', 'Plataforma', 'FechaPedido', 'Estado', 'Descripcion', 'Ubicacion', 'ActualizadoEn', 'Archivado', 'Registrado', 'CreadoEn', 'Link', 'Proveedor'];

// Pestaña "Cashback" — retiros de cashback de la tarjeta de Alberto,
// registrados desde la app "Mis Gastos - Personal" (chat aparte). Full-
// replace como Ventas/Stocks/Campañas porque los retiros se pueden editar o
// borrar ahí. El dashboard los usa para calcular cuánto cashback ya
// "recuperó" cada mes contra los gastos PERSONALES (FIFO cronológico hacia
// adelante — ver getCashbackUsadoPorMes en app.js). Se lee en vivo por
// ?action=cashback (como Compras/Seguimiento/AnunciosMeta) — no necesita
// "Publicar en la web".
var CASHBACK_SHEET_NAME = 'Cashback';
var CASHBACK_HEADERS = ['Fecha', 'Monto', 'Nota'];

// Carpeta de Drive donde se guardan las fotos de los bloques de compra.
var FOTOS_FOLDER_NAME = 'Timeless - Fotos accesorios';

// Recibe cada gasto (POST) y lo agrega como fila nueva, o un sync de datos.
function doPost(e) {
  var data = JSON.parse(e.postData.contents);

  if (data.type === 'ventasSync')   return handleVentasSync_(data);
  if (data.type === 'stocksSync')   return handleStocksSync_(data);
  if (data.type === 'campanasSync') return handleCampanasSync_(data);
  if (data.type === 'cashbackSync') return handleCashbackSync_(data);
  if (data.type === 'compraGuardar')   return handleCompraGuardar_(data);
  if (data.type === 'compraEliminar')  return handleCompraEliminar_(data);
  if (data.type === 'compraFoto')      return handleCompraFoto_(data);
  if (data.type === 'seguimientoGuardar')  return handleSeguimientoGuardar_(data);
  if (data.type === 'seguimientoEliminar') return handleSeguimientoEliminar_(data);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // evita filas duplicadas si llegan dos a la vez

    var sheet = getOrCreateSheet_();

    // Encabezados la primera vez.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID', 'Fecha', 'Categoría', 'Monto', 'Nota', 'Registrado en']);
    }

    // Si la ID ya existe es una EDICIÓN (o un reintento): actualiza esa fila
    // en vez de descartarla o duplicarla. (Fix 18-ago-2026: antes cualquier ID
    // repetida se trataba como duplicado y se ignoraba — así que editar un
    // gasto en la app nunca se reflejaba en Sheets.)
    var fecha = data.date ? new Date(data.date) : new Date();
    var filaGasto = data.id ? findRowById_(sheet, data.id) : -1;

    if (filaGasto > 0) {
      sheet.getRange(filaGasto, 2, 1, 4).setValues([[
        fecha,
        data.category || '',
        Number(data.amount) || 0,
        data.note || ''
      ]]);
      return json_({ ok: true, updated: true });
    }

    sheet.appendRow([
      data.id || '',
      fecha,
      data.category || '',
      Number(data.amount) || 0,
      data.note || '',
      new Date()
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// Formatos de columna (evitan que un formato viejo, ej. la columna Gasto que
// estaba como FECHA, altere cómo se muestran/exportan los datos):
// '@' = texto, '0.00' = número 2 decimales, '0' = entero, 'dd/mm/yyyy' = fecha.
var FMT_VENTAS   = ['dd/mm/yyyy', '@', '0.00', '0.00'];
var FMT_STOCKS   = ['@', '0.00', '0', '0', '0.00', '0.00', '0.00', '0', 'dd/mm/yyyy', '@', '0.00'];
var FMT_CAMPANAS = ['dd/mm/yyyy', '@', '0.00'];
var FMT_CASHBACK = ['dd/mm/yyyy', '0.00', '@'];
var FMT_ANUNCIOS_META = ['dd/mm/yyyy', '@', '@', '0.00', '0'];

// Reemplaza TODA la pestaña VentasDetalle con las filas que manda el Excel.
function handleVentasSync_(data) {
  return fullReplace_(VENTAS_SHEET_NAME, VENTAS_HEADERS, data.rows, function (r) {
    return [new Date(r.fecha), String(r.producto), Number(r.venta) || 0, Number(r.utilidad) || 0];
  }, FMT_VENTAS);
}

// Reemplaza TODA la pestaña Stocks con las filas que manda el Excel "Venta accs.xlsx".
// "Cantidad pedido" es la cantidad ORIGINAL del pedido (columna "Cantidad" del
// Excel) — se usa para proyectar ingresos/ganancia de pedidos aún sin llegar
// (con stock=0, "Ganancia bruta/neta posible" da 0 porque esas son sobre el
// stock actual, no sobre el pedido completo).
// "Fecha pedido" / "Plataforma" salen del bloque de color de la columna A del
// Excel (la tienda y la "(dd/mm)" que Alberto escribe al costado de cada
// pedido): el dashboard las usa para mostrar "pedido el 3 ago · hace 5 días"
// en Pedidos por llegar.
function handleStocksSync_(data) {
  return fullReplace_(STOCKS_SHEET_NAME, STOCKS_HEADERS, data.rows, function (r) {
    return [
      String(r.producto),
      Number(r.precio) || 0,
      Number(r.vendidos) || 0,
      Number(r.stock) || 0,
      Number(r.ganBruta) || 0,
      Number(r.ganNeta) || 0,
      Number(r.invertido) || 0,
      Number(r.cantidadPedido) || 0,
      fechaLocal_(r.fechaPedido),
      String(r.plataforma || ''),
      Number(r.stockInvertido) || 0
    ];
  }, FMT_STOCKS);
}

// "yyyy-MM-dd" -> Date en hora LOCAL (mediodía). Ojo: new Date('2026-08-03')
// lo interpreta como UTC medianoche, que en Lima (-5) cae el 2/08 a las 19:00 y
// se guardaría con un día menos. Por eso se arma con año/mes/día sueltos.
function fechaLocal_(iso) {
  if (!iso) return '';
  var p = String(iso).split('-');
  if (p.length !== 3) return '';
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
}

// Reemplaza TODA la pestaña Campañas con el gasto real por día/campaña de Meta Ads.
function handleCampanasSync_(data) {
  return fullReplace_(CAMPANAS_SHEET_NAME, CAMPANAS_HEADERS, data.rows, function (r) {
    return [new Date(r.fecha), String(r.campana), Number(r.gasto) || 0];
  }, FMT_CAMPANAS);
}

// Reemplaza TODA la pestaña Cashback con los retiros que manda la app de
// gastos. r.date llega como 'YYYY-MM-DDTHH:mm:ss' SIN "Z" a propósito (así
// new Date() lo lee como hora local, no UTC — mismo cuidado que fechaLocal_
// arriba, pero acá ya viene con hora incluida así que no hace falta armarla
// a mano).
function handleCashbackSync_(data) {
  return fullReplace_(CASHBACK_SHEET_NAME, CASHBACK_HEADERS, data.rows, function (r) {
    return [new Date(r.date), Number(r.amount) || 0, String(r.note || '')];
  }, FMT_CASHBACK);
}

// Lee Cashback EN VIVO (sin publicar a la web) para ?action=cashback.
function getCashbackVivo_() {
  var sheet = getOrCreateNamedSheet_(CASHBACK_SHEET_NAME);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, CASHBACK_HEADERS.length).getValues();
  return rows.filter(function (r) { return r[0] && Number(r[1]) > 0; }).map(function (r) {
    return {
      date: r[0] ? new Date(r[0]).toISOString() : '',
      amount: Number(r[1]) || 0,
      note: String(r[2] || '')
    };
  });
}

// Reemplaza TODA la pestaña AnunciosMeta con el gasto + conversaciones reales
// por día/anuncio de Meta Ads.
function handleAnunciosMetaSync_(data) {
  return fullReplace_(ANUNCIOS_META_SHEET_NAME, ANUNCIOS_META_HEADERS, data.rows, function (r) {
    return [new Date(r.fecha), String(r.anuncio), String(r.campana), Number(r.gasto) || 0, Number(r.conversaciones) || 0];
  }, FMT_ANUNCIOS_META);
}

// Lee AnunciosMeta EN VIVO (sin publicar a la web) para ?action=anunciosMeta.
function getAnunciosMetaVivo_() {
  var sheet = getOrCreateNamedSheet_(ANUNCIOS_META_SHEET_NAME);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, ANUNCIOS_META_HEADERS.length).getValues();
  return rows.filter(function (r) { return r[1]; }).map(function (r) {
    return {
      fecha: r[0] ? new Date(r[0]).toISOString() : '',
      anuncio: String(r[1] || ''),
      campana: String(r[2] || ''),
      gasto: Number(r[3]) || 0,
      conversaciones: Number(r[4]) || 0
    };
  });
}

// ==================== SYNC AUTOMÁTICO DE META ADS ====================
// Jala el gasto real por día y campaña DIRECTO de Meta (sin pasar por Claude)
// y actualiza la pestaña Campañas solo. Pensado para correr con un trigger de
// tiempo (Activadores → Añadir activador → syncMetaAdsAutomatico → basado en
// tiempo → cada 8 horas = 3 veces al día).
//
// REQUIERE: guardar el token en Configuración del proyecto → Propiedades del
// script → clave "META_ACCESS_TOKEN" (token de Usuario del sistema, permiso
// ads_read, expiración "Nunca" — generado en Meta Business Manager).
var META_AD_ACCOUNT_ID = '1273938200074568';
var META_API_VERSION = 'v20.0';

function syncMetaAdsAutomatico() {
  var token = PropertiesService.getScriptProperties().getProperty('META_ACCESS_TOKEN');
  if (!token) {
    Logger.log('Falta guardar META_ACCESS_TOKEN en Configuración del proyecto → Propiedades del script.');
    return;
  }

  var hoy = new Date();
  var hace30 = new Date();
  hace30.setDate(hace30.getDate() - 30);
  var since = Utilities.formatDate(hace30, 'America/Lima', 'yyyy-MM-dd');
  var until = Utilities.formatDate(hoy, 'America/Lima', 'yyyy-MM-dd');

  var url = 'https://graph.facebook.com/' + META_API_VERSION + '/act_' + META_AD_ACCOUNT_ID + '/insights'
    + '?level=campaign'
    + '&fields=campaign_name,spend'
    + '&time_range=' + encodeURIComponent(JSON.stringify({ since: since, until: until }))
    + '&time_increment=1'
    + '&limit=500'
    + '&access_token=' + token;

  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(resp.getContentText());

  if (json.error) {
    Logger.log('Error de Meta: ' + JSON.stringify(json.error));
    return;
  }

  var rows = (json.data || [])
    .filter(function (d) { return Number(d.spend) > 0; })
    .map(function (d) {
      return { fecha: d.date_start + 'T12:00:00', campana: d.campaign_name, gasto: Number(d.spend) };
    });

  if (!rows.length) {
    Logger.log('Meta no devolvió gasto en el rango de 30 días.');
    return;
  }

  handleCampanasSync_({ rows: rows });
  Logger.log('Sync automático de Meta Ads OK: ' + rows.length + ' filas.');

  syncMetaAdsPorAnuncio_(token, since, until);
}

// Segunda pasada: gasto + conversaciones POR ANUNCIO (nivel "ad", no "campaign"),
// para el semáforo de costo por conversación del dashboard. Aparte de la de
// arriba porque Meta no deja pedir level=campaign y level=ad en la misma
// llamada. "actions" trae varios tipos de evento por fila; solo se suma el
// que empieza con "messaging_conversation_started" (conversaciones iniciadas
// por WhatsApp/Messenger, que es la métrica que usan los anuncios de Alberto).
function syncMetaAdsPorAnuncio_(token, since, until) {
  var url = 'https://graph.facebook.com/' + META_API_VERSION + '/act_' + META_AD_ACCOUNT_ID + '/insights'
    + '?level=ad'
    + '&fields=ad_name,campaign_name,spend,actions'
    + '&time_range=' + encodeURIComponent(JSON.stringify({ since: since, until: until }))
    + '&time_increment=1'
    + '&limit=500'
    + '&access_token=' + token;

  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(resp.getContentText());

  if (json.error) {
    Logger.log('Error de Meta (por anuncio): ' + JSON.stringify(json.error));
    return;
  }

  var rows = (json.data || [])
    .filter(function (d) { return Number(d.spend) > 0; })
    .map(function (d) {
      return {
        fecha: d.date_start + 'T12:00:00',
        anuncio: d.ad_name,
        campana: d.campaign_name,
        gasto: Number(d.spend),
        conversaciones: extraerConversaciones_(d.actions)
      };
    });

  if (!rows.length) {
    Logger.log('Meta no devolvió gasto por anuncio en el rango de 30 días.');
    return;
  }

  handleAnunciosMetaSync_({ rows: rows });
  Logger.log('Sync automático de Meta Ads por anuncio OK: ' + rows.length + ' filas.');
}

function extraerConversaciones_(actions) {
  if (!actions) return 0;
  var total = 0;
  actions.forEach(function (a) {
    if (String(a.action_type || '').indexOf('messaging_conversation_started') !== -1) {
      total += Number(a.value) || 0;
    }
  });
  return total;
}

// Crea o actualiza UN bloque de compra (a diferencia de los sync de arriba,
// esto NO es full-replace: solo toca la fila de ese ID). Si data.compra.id ya
// existe en la pestaña, actualiza esa fila; si no, la crea con un ID nuevo.
function handleCompraGuardar_(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var c = data.compra || {};
    var sheet = getOrCreateNamedSheet_(COMPRAS_SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COMPRAS_HEADERS);
      sheet.getRange(1, 1, 1, COMPRAS_HEADERS.length).setFontWeight('bold');
    }

    var id = c.id;
    var rowIdx = id ? findRowById_(sheet, id) : -1;
    var esNuevo = rowIdx === -1;
    if (esNuevo) id = Utilities.getUuid();

    var fila = [
      id,
      String(c.nombre || ''),
      String(c.estado || 'Nuevo'),
      c.fechaInicio ? new Date(c.fechaInicio) : '',
      c.fechaFin ? new Date(c.fechaFin) : '',
      Number(c.precioTotal) || 0,
      JSON.stringify(c.productos || []),
      JSON.stringify(c.fotos || []),
      String(c.notas || ''),
      esNuevo ? new Date() : (c.creadoEn ? new Date(c.creadoEn) : new Date())
    ];

    if (esNuevo) {
      sheet.appendRow(fila);
    } else {
      sheet.getRange(rowIdx, 1, 1, COMPRAS_HEADERS.length).setValues([fila]);
    }
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// Borra UN bloque de compra por ID.
function handleCompraEliminar_(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sheet = getOrCreateNamedSheet_(COMPRAS_SHEET_NAME);
    var rowIdx = findRowById_(sheet, data.id);
    if (rowIdx === -1) return json_({ ok: false, error: 'No existe ese ID' });
    sheet.deleteRow(rowIdx);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// Busca la fila (1-indexed) cuya columna A coincide con el ID. -1 si no existe.
function findRowById_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

// Sube una foto (base64, mandada por el dashboard ya redimensionada en el
// navegador) a una carpeta de Drive y la deja visible por link, para poder
// mostrarla como <img> en el dashboard público.
function handleCompraFoto_(data) {
  try {
    var folder = getOrCreateFolder_(FOTOS_FOLDER_NAME);
    var bytes = Utilities.base64Decode(data.base64);
    var blob = Utilities.newBlob(bytes, data.mimeType || 'image/jpeg', data.filename || (Utilities.getUuid() + '.jpg'));
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
    return json_({ ok: true, url: url, fileId: file.getId() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function getOrCreateFolder_(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// Función de UN SOLO USO para autorizar el acceso a Drive: en el editor,
// elige "autorizarDrive" en el desplegable de arriba (NO "doPost" — esa
// espera datos de una petición web real y truena si se ejecuta a mano) y
// dale a ▶ Ejecutar. Va a pedir permiso de Drive; acéptalo. No hace falta
// volver a Implementar: el permiso queda a nivel de cuenta/proyecto y el
// sitio ya desplegado lo usa de inmediato.
function autorizarDrive() {
  var folder = getOrCreateFolder_(FOTOS_FOLDER_NAME);
  Logger.log('Listo. Carpeta de fotos: ' + folder.getUrl());
}

// Devuelve todos los bloques de compra en vivo (sin esperar el retraso de
// "Publicar en la web"), para que el dashboard vea sus propios cambios al
// instante después de guardar.
function getComprasVivo_() {
  var sheet = getOrCreateNamedSheet_(COMPRAS_SHEET_NAME);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, COMPRAS_HEADERS.length).getValues();
  return rows.filter(function (r) { return r[0]; }).map(function (r) {
    var productos = [];
    try { productos = JSON.parse(r[6] || '[]'); } catch (ignore) {}
    // Columna de fotos: antes guardaba UNA url de texto plano; ahora guarda un
    // JSON con VARIAS. Si es una fila vieja (no es JSON válido), la migramos
    // sola a un array de 1 elemento para que el dashboard no note la diferencia.
    var fotos = [];
    var rawFotos = r[7];
    if (rawFotos) {
      try { fotos = JSON.parse(rawFotos); if (!Array.isArray(fotos)) fotos = [String(rawFotos)]; }
      catch (ignore) { fotos = [String(rawFotos)]; }
    }
    return {
      id: String(r[0]),
      nombre: String(r[1]),
      estado: String(r[2]),
      fechaInicio: r[3] ? new Date(r[3]).toISOString() : '',
      fechaFin: r[4] ? new Date(r[4]).toISOString() : '',
      precioTotal: Number(r[5]) || 0,
      productos: productos,
      fotos: fotos,
      notas: String(r[8] || ''),
      creadoEn: r[9] ? new Date(r[9]).toISOString() : ''
    };
  });
}

// Helper común: limpia una pestaña existente (conserva su gid publicado) y
// reescribe encabezados + filas. Busca la hoja por NOMBRE, nunca crea una nueva
// si ya existe (así no se rompe la URL CSV publicada).
function fullReplace_(sheetName, headers, rows, mapFn, colFormats) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!rows || !rows.length) {
      return json_({ ok: false, error: 'Sin filas para sincronizar (' + sheetName + ')' });
    }

    var sheet = getOrCreateNamedSheet_(sheetName);

    // Reset TOTAL de la pestaña antes de reescribir. clearContents() NO basta:
    // la pestaña "Campanas" original tenía la columna Gasto formateada como
    // FECHA y celdas COMBINADAS, así que los números se mostraban como fechas
    // (ej. 18.88 -> "17.01" = 17/ene) y media data salía vacía. breakApart()
    // deshace las combinaciones, clear() borra contenido + formato, y abajo se
    // fuerza el formato correcto por columna. A prueba de balas.
    var maxR = sheet.getMaxRows(), maxC = sheet.getMaxColumns();
    sheet.getRange(1, 1, maxR, maxC).breakApart();
    sheet.clear();

    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    var values = rows.map(mapFn);
    var dataRange = sheet.getRange(2, 1, values.length, headers.length);
    dataRange.setValues(values);

    // Forzar el formato de cada columna (sobrescribe cualquier formato viejo).
    if (colFormats) {
      for (var c = 0; c < colFormats.length; c++) {
        sheet.getRange(2, c + 1, values.length, 1).setNumberFormat(colFormats[c]);
      }
    }

    return json_({ ok: true, count: values.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// Permite abrir la URL en el navegador para comprobar que está viva, y sirve
// ?action=compras / ?action=instagram para leer datos EN VIVO (sin el retraso
// del CSV publicado, para que el dashboard vea todo al instante).
// ==================== SEGUIMIENTO DE PEDIDOS (17TRACK) ====================
// Alberto registra en el dashboard cada paquete (número de tracking + qué
// productos vienen dentro + la tienda). sync17Track() corre con un activador de
// tiempo, le pregunta a la API de 17TRACK por el último estado de cada uno y lo
// escribe en la pestaña "Seguimiento", que el dashboard lee por ?action=seguimiento.
//
// REQUIERE: Configuración del proyecto → Propiedades del script → clave
// "T17_TOKEN" con la API key de https://api.17track.net (cuenta gratis).
// La cuota gratis se consume al REGISTRAR cada número (una sola vez); las
// consultas de estado posteriores no gastan cuota. Por eso el registro se
// marca en la columna "Registrado" y nunca se repite.
var T17_BASE = 'https://api.17track.net/track/v2.2/';

// Estados en los que ya vale la pena estar atento (dispara correo de aviso).
var T17_ESTADOS_AVISO = ['OutForDelivery', 'AvailableForPickup', 'Delivered', 'DeliveryFailure', 'Exception'];

function t17Estado_(status) {
  var map = {
    NotFound: 'Sin información',
    InfoReceived: 'Info recibida',
    InTransit: 'En camino',
    Expired: 'Demorado',
    AvailableForPickup: 'Listo para recoger',
    OutForDelivery: 'En reparto',
    DeliveryFailure: 'Entrega fallida',
    Delivered: 'Entregado',
    Exception: 'Problema'
  };
  return map[status] || status || '';
}

function t17Post_(endpoint, payload) {
  var token = PropertiesService.getScriptProperties().getProperty('T17_TOKEN');
  if (!token) throw new Error('Falta T17_TOKEN en Propiedades del script.');
  var res = UrlFetchApp.fetch(T17_BASE + endpoint, {
    method: 'post',
    contentType: 'application/json',
    headers: { '17token': token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var txt = res.getContentText();
  var body;
  try { body = JSON.parse(txt); } catch (err) { throw new Error(endpoint + ': respuesta no-JSON: ' + txt.slice(0, 300)); }
  return body;
}

function getSeguimientoVivo_() {
  var sheet = getOrCreateNamedSheet_(SEGUIMIENTO_SHEET_NAME);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, SEGUIMIENTO_HEADERS.length).getValues();
  return rows.filter(function (r) { return r[0]; }).map(function (r) {
    return {
      id: String(r[0]),
      tracking: String(r[1] || ''),
      productos: String(r[2] || ''),
      plataforma: String(r[3] || ''),
      fechaPedido: r[4] ? new Date(r[4]).toISOString() : '',
      estado: String(r[5] || ''),
      descripcion: String(r[6] || ''),
      ubicacion: String(r[7] || ''),
      actualizadoEn: r[8] ? new Date(r[8]).toISOString() : '',
      archivado: String(r[9] || '') === 'si',
      creadoEn: r[11] ? new Date(r[11]).toISOString() : '',
      link: String(r[12] || ''),
      proveedor: String(r[13] || '')
    };
  });
}

function handleSeguimientoGuardar_(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = data.paquete || {};
    var sheet = getOrCreateNamedSheet_(SEGUIMIENTO_SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(SEGUIMIENTO_HEADERS);
      sheet.getRange(1, 1, 1, SEGUIMIENTO_HEADERS.length).setFontWeight('bold');
    }

    var id = p.id;
    var rowIdx = id ? findRowById_(sheet, id) : -1;
    var esNuevo = rowIdx === -1;
    if (esNuevo) id = Utilities.getUuid();

    // Al editar se conservan las columnas que llena el sync (estado, etc.);
    // solo se pisan si cambió el número de tracking (ahí ya no aplican).
    var previo = esNuevo ? [] : sheet.getRange(rowIdx, 1, 1, SEGUIMIENTO_HEADERS.length).getValues()[0];
    var trackingNuevo = String(p.tracking || '').trim().toUpperCase();
    var cambioTracking = !esNuevo && String(previo[1] || '').toUpperCase() !== trackingNuevo;

    var fila = [
      id,
      trackingNuevo,
      String(p.productos || ''),
      String(p.plataforma || ''),
      p.fechaPedido ? fechaLocal_(String(p.fechaPedido).slice(0, 10)) : '',
      cambioTracking || esNuevo ? '' : previo[5],
      cambioTracking || esNuevo ? '' : previo[6],
      cambioTracking || esNuevo ? '' : previo[7],
      cambioTracking || esNuevo ? '' : previo[8],
      p.archivado ? 'si' : '',
      cambioTracking || esNuevo ? '' : previo[10],
      esNuevo ? new Date() : (previo[11] || new Date()),
      String(p.link || ''),
      String(p.proveedor || '')
    ];

    if (esNuevo) sheet.appendRow(fila);
    else sheet.getRange(rowIdx, 1, 1, SEGUIMIENTO_HEADERS.length).setValues([fila]);

    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function handleSeguimientoEliminar_(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sheet = getOrCreateNamedSheet_(SEGUIMIENTO_SHEET_NAME);
    var rowIdx = findRowById_(sheet, data.id);
    if (rowIdx === -1) return json_({ ok: false, error: 'No existe ese ID' });
    sheet.deleteRow(rowIdx);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// Baja el último estado de cada paquete NO archivado. Pensado para un activador
// de tiempo cada 4 horas (Activadores → Añadir → sync17Track → cada 4 horas).
function sync17Track() {
  var sheet = getOrCreateNamedSheet_(SEGUIMIENTO_SHEET_NAME);
  var last = sheet.getLastRow();
  if (last < 2) { Logger.log('Seguimiento: sin paquetes.'); return; }

  var rows = sheet.getRange(2, 1, last - 1, SEGUIMIENTO_HEADERS.length).getValues();
  var activos = [];
  for (var i = 0; i < rows.length; i++) {
    var tracking = String(rows[i][1] || '').trim().toUpperCase();
    if (!tracking) continue;
    if (String(rows[i][9] || '') === 'si') continue; // archivado a mano
    activos.push({ fila: i + 2, idx: i, tracking: tracking });
  }
  if (!activos.length) { Logger.log('Seguimiento: nada activo que consultar.'); return; }

  // 1) Registrar los que nunca se registraron (esto es lo único que gasta cuota).
  var porRegistrar = activos.filter(function (a) { return String(rows[a.idx][10] || '') !== 'si'; });
  if (porRegistrar.length) {
    var reg = t17Post_('register', porRegistrar.map(function (a) { return { number: a.tracking, auto_detection: true }; }));
    Logger.log('register: ' + JSON.stringify(reg).slice(0, 500));
    // Se marca como registrado tanto si lo acepta como si lo rechaza por "ya
    // existe": en ambos casos no hay que volver a gastar cuota en ese número.
    porRegistrar.forEach(function (a) { sheet.getRange(a.fila, 11).setValue('si'); });
  }

  // 2) Pedir el estado de todos (esto NO gasta cuota). Máx. 40 por llamada.
  var avisos = [];
  for (var b = 0; b < activos.length; b += 40) {
    var lote = activos.slice(b, b + 40);
    var resp = t17Post_('gettrackinfo', lote.map(function (a) { return { number: a.tracking }; }));
    var aceptados = (resp && resp.data && resp.data.accepted) || [];
    for (var j = 0; j < aceptados.length; j++) {
      var it = aceptados[j];
      var info = it.track_info || {};
      var estado = (info.latest_status && info.latest_status.status) || '';
      var ev = info.latest_event || {};
      var destino = null;
      for (var k = 0; k < lote.length; k++) {
        if (lote[k].tracking === String(it.number || '').toUpperCase()) { destino = lote[k]; break; }
      }
      if (!destino) continue;

      var estadoAntes = String(rows[destino.idx][5] || '');
      sheet.getRange(destino.fila, 6, 1, 4).setValues([[
        estado,
        String(ev.description || ''),
        String(ev.location || ''),
        ev.time_iso ? new Date(ev.time_iso) : ''
      ]]);

      if (estado && estado !== estadoAntes && T17_ESTADOS_AVISO.indexOf(estado) !== -1) {
        avisos.push({
          tracking: destino.tracking,
          productos: String(rows[destino.idx][2] || ''),
          estado: estado,
          descripcion: String(ev.description || ''),
          ubicacion: String(ev.location || '')
        });
      }
    }
    var rechazados = (resp && resp.data && resp.data.rejected) || [];
    if (rechazados.length) Logger.log('gettrackinfo rechazados: ' + JSON.stringify(rechazados).slice(0, 500));
  }

  if (avisos.length) enviarAviso17Track_(avisos);
  Logger.log('Seguimiento: ' + activos.length + ' paquete(s) consultados, ' + avisos.length + ' aviso(s).');
}

// Correo de aviso cuando un paquete cambia a un estado que importa. Es la forma
// más simple de que suene el celular: no hace falta servidor de push, llega
// como cualquier correo de Gmail.
function enviarAviso17Track_(avisos) {
  var lineas = avisos.map(function (a) {
    return '📦 ' + (a.productos || a.tracking) + '\n' +
      '   ' + t17Estado_(a.estado) + (a.ubicacion ? ' · ' + a.ubicacion : '') + '\n' +
      '   ' + a.descripcion + '\n' +
      '   ' + a.tracking;
  }).join('\n\n');
  var titulo = avisos.length === 1
    ? 'Timeless · ' + (avisos[0].productos || avisos[0].tracking) + ' — ' + t17Estado_(avisos[0].estado)
    : 'Timeless · ' + avisos.length + ' pedidos cambiaron de estado';
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), titulo, lineas);
}

// Para probar a mano desde el editor: muestra en el registro lo que responde
// 17TRACK, sin escribir nada en la hoja.
function probar17Track() {
  var sheet = getOrCreateNamedSheet_(SEGUIMIENTO_SHEET_NAME);
  var last = sheet.getLastRow();
  if (last < 2) { Logger.log('No hay paquetes en la pestaña Seguimiento.'); return; }
  var nums = sheet.getRange(2, 2, last - 1, 1).getValues()
    .map(function (r) { return String(r[0] || '').trim().toUpperCase(); })
    .filter(function (n) { return n; });
  Logger.log('Trackings en la hoja: ' + nums.join(', '));
  Logger.log(JSON.stringify(t17Post_('gettrackinfo', nums.map(function (n) { return { number: n }; })), null, 2));
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;
  if (action === 'compras') {
    return json_({ ok: true, compras: getComprasVivo_() });
  }
  if (action === 'instagram') {
    return json_({ ok: true, instagram: getInstagramVivo_() });
  }
  if (action === 'seguimiento') {
    return json_({ ok: true, seguimiento: getSeguimientoVivo_() });
  }
  if (action === 'anunciosMeta') {
    return json_({ ok: true, anunciosMeta: getAnunciosMetaVivo_() });
  }
  if (action === 'cashback') {
    return json_({ ok: true, cashback: getCashbackVivo_() });
  }
  return json_({ ok: true, service: 'Mis Gastos - Personal + sync Ventas/Stocks/Campañas/Instagram', version: 'v6-cashback', sheet: SHEET_NAME });
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

// Busca una pestaña por nombre. 1) match exacto; 2) match tolerante a acentos,
// mayúsculas y espacios (así "Campañas" encuentra "Campanas", etc.) para nunca
// crear una pestaña nueva por accidente y romper la URL CSV publicada;
// 3) recién si no existe ninguna, la crea.
function getOrCreateNamedSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  var norm = normalizarNombre_(name);
  var todas = ss.getSheets();
  for (var i = 0; i < todas.length; i++) {
    if (normalizarNombre_(todas[i].getName()) === norm) return todas[i];
  }
  return ss.insertSheet(name);
}

function normalizarNombre_(s) {
  return String(s).toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // quita acentos/ñ→n
    .replace(/\s+/g, ' ').trim();
}

// NOTA: ya no se usa (reemplazada por findRowById_ + update en doPost, fix
// 18-ago-2026). Se deja por si algo más la referencia; se puede borrar.
function idYaExiste_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return true;
  }
  return false;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== INSTAGRAM ====================
// Métricas orgánicas de Instagram (seguidores + crecimiento, alcance, visitas
// al perfil y mejores publicaciones), vía Instagram Graph API con el MISMO
// token del sync de ads (ahora con permisos instagram_basic +
// instagram_manage_insights). Escribe 3 pestañas que el dashboard lee en vivo
// por ?action=instagram. Corre en su propio trigger cada 8h (syncInstagram).
var IG_USER_ID = '17841456244505708';

function syncInstagram() {
  var token = PropertiesService.getScriptProperties().getProperty('META_ACCESS_TOKEN');
  if (!token) { Logger.log('Falta META_ACCESS_TOKEN.'); return; }
  var V = META_API_VERSION;
  var TZ = 'America/Lima';

  // 1) Datos básicos de la cuenta (permiso instagram_basic).
  var basico = igGet_(V, IG_USER_ID, { fields: 'username,followers_count,media_count' }, token);
  if (basico.error) { Logger.log('IG básico error: ' + JSON.stringify(basico.error)); return; }
  var seguidores = Number(basico.followers_count) || 0;
  var publicaciones = Number(basico.media_count) || 0;

  // 2) Snapshot diario de seguidores (upsert por fecha) para medir crecimiento.
  var hoyStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  upsertSeguidores_(hoyStr, seguidores, publicaciones);

  // 3) Series diarias de alcance / visitas / nuevos seguidores (30 días).
  var hace30 = new Date(); hace30.setDate(hace30.getDate() - 30);
  var since = Math.floor(hace30.getTime() / 1000);
  var until = Math.floor(Date.now() / 1000);
  var alcance = igInsightSerie_(V, IG_USER_ID, 'reach', since, until, token, TZ);
  var visitas = igInsightSerie_(V, IG_USER_ID, 'profile_views', since, until, token, TZ);
  var nuevos  = igInsightSerie_(V, IG_USER_ID, 'follower_count', since, until, token, TZ);

  var fechas = {};
  [alcance, visitas, nuevos].forEach(function (m) { Object.keys(m).forEach(function (k) { fechas[k] = 1; }); });
  var diaRows = Object.keys(fechas).sort().map(function (f) {
    return { fecha: f, alcance: alcance[f] || 0, visitas: visitas[f] || 0, nuevos: nuevos[f] || 0 };
  });
  if (diaRows.length) {
    fullReplace_('InstagramDia', ['Fecha', 'Alcance', 'VisitasPerfil', 'NuevosSeguidores'], diaRows,
      function (r) { return [r.fecha + 'T12:00:00', r.alcance, r.visitas, r.nuevos]; },
      ['yyyy-mm-dd', '0', '0', '0']);
  }

  // 4) Mejores publicaciones por interacción (like+coment), con alcance/guardados.
  var media = igGet_(V, IG_USER_ID + '/media', {
    fields: 'caption,media_type,media_product_type,timestamp,permalink,thumbnail_url,media_url,like_count,comments_count',
    limit: '25'
  }, token);
  var items = (media && media.data) ? media.data : [];
  items.forEach(function (m) { m._eng = (Number(m.like_count) || 0) + (Number(m.comments_count) || 0); });
  items.sort(function (a, b) { return b._eng - a._eng; });
  var top = items.slice(0, 15);
  top.forEach(function (m) {
    m._reach = 0; m._saved = 0;
    var ins = igGet_(V, m.id + '/insights', { metric: 'reach,saved' }, token);
    if (ins && ins.error) ins = igGet_(V, m.id + '/insights', { metric: 'reach' }, token); // reels/tipos raros
    if (ins && ins.data) {
      ins.data.forEach(function (d) {
        var val = (d.values && d.values[0]) ? Number(d.values[0].value) : 0;
        if (d.name === 'reach') m._reach = val;
        if (d.name === 'saved') m._saved = val;
      });
    }
  });
  if (top.length) {
    fullReplace_('InstagramMedia',
      ['Tipo', 'Fecha', 'Caption', 'Likes', 'Comentarios', 'Alcance', 'Guardados', 'Permalink', 'Thumbnail'],
      top,
      function (m) {
        var tipo = m.media_product_type === 'REELS' ? 'Reel'
          : (m.media_type === 'VIDEO' ? 'Video'
          : (m.media_type === 'CAROUSEL_ALBUM' ? 'Carrusel' : 'Foto'));
        var cap = String(m.caption || '').replace(/\s+/g, ' ').slice(0, 140);
        return [tipo, m.timestamp || '', cap, Number(m.like_count) || 0, Number(m.comments_count) || 0,
                m._reach || 0, m._saved || 0, String(m.permalink || ''), String(m.thumbnail_url || m.media_url || '')];
      },
      ['@', 'yyyy-mm-dd', '@', '0', '0', '0', '0', '@', '@']);
  }

  Logger.log('Sync Instagram OK: ' + seguidores + ' seguidores, ' + diaRows.length + ' días, ' + top.length + ' posts.');
}

// GET a la Graph API devolviendo objeto JSON parseado (o {error:...}).
function igGet_(version, node, params, token) {
  var qs = Object.keys(params || {}).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  var url = 'https://graph.facebook.com/' + version + '/' + node + '?' + (qs ? qs + '&' : '') + 'access_token=' + token;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  try { return JSON.parse(resp.getContentText()); } catch (e) { return { error: String(e) }; }
}

// Serie diaria {yyyy-MM-dd: valor} de una métrica de cuenta. {} si la métrica
// no está disponible (así el sync no se cae si Meta deprecia alguna).
function igInsightSerie_(version, node, metric, since, until, token, tz) {
  var out = {};
  var j = igGet_(version, node + '/insights', { metric: metric, period: 'day', since: String(since), until: String(until) }, token);
  if (!j || j.error || !j.data || !j.data.length) return out;
  (j.data[0].values || []).forEach(function (v) {
    out[Utilities.formatDate(new Date(v.end_time), tz, 'yyyy-MM-dd')] = Number(v.value) || 0;
  });
  return out;
}

// Guarda/actualiza el total de seguidores + nº de publicaciones de HOY (una fila
// por día) para poder graficar el crecimiento a lo largo del tiempo.
function upsertSeguidores_(fechaStr, total, publicaciones) {
  var sheet = getOrCreateNamedSheet_('InstagramSeguidores');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Fecha', 'Seguidores', 'Publicaciones']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  var last = sheet.getLastRow();
  var fechas = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getValues() : [];
  for (var i = 0; i < fechas.length; i++) {
    var f = fechas[i][0];
    var fs = (f instanceof Date) ? Utilities.formatDate(f, 'America/Lima', 'yyyy-MM-dd') : String(f).slice(0, 10);
    if (fs === fechaStr) {
      sheet.getRange(i + 2, 2, 1, 2).setValues([[total, publicaciones]]);
      return;
    }
  }
  sheet.appendRow([fechaStr + 'T12:00:00', total, publicaciones]);
}

// Lee las 3 pestañas de Instagram y arma el JSON que consume el dashboard.
function getInstagramVivo_() {
  var out = { username: '', seguidores: 0, publicaciones: 0, dia: [], seguidoresHist: [], media: [] };

  var segSheet = getOrCreateNamedSheet_('InstagramSeguidores');
  var sl = segSheet.getLastRow();
  if (sl >= 2) {
    var segRows = segSheet.getRange(2, 1, sl - 1, 3).getValues();
    out.seguidoresHist = segRows.filter(function (r) { return r[0]; }).map(function (r) {
      return { fecha: toIso_(r[0]), total: Number(r[1]) || 0 };
    });
    if (out.seguidoresHist.length) {
      out.seguidores = out.seguidoresHist[out.seguidoresHist.length - 1].total;
      out.publicaciones = Number(segRows[segRows.length - 1][2]) || 0;
    }
  }

  var diaSheet = getOrCreateNamedSheet_('InstagramDia');
  var dl = diaSheet.getLastRow();
  if (dl >= 2) {
    var diaRows = diaSheet.getRange(2, 1, dl - 1, 4).getValues();
    out.dia = diaRows.filter(function (r) { return r[0]; }).map(function (r) {
      return { fecha: toIso_(r[0]), alcance: Number(r[1]) || 0, visitas: Number(r[2]) || 0, nuevos: Number(r[3]) || 0 };
    });
  }

  var medSheet = getOrCreateNamedSheet_('InstagramMedia');
  var ml = medSheet.getLastRow();
  if (ml >= 2) {
    var medRows = medSheet.getRange(2, 1, ml - 1, 9).getValues();
    out.media = medRows.filter(function (r) { return r[7]; }).map(function (r) {
      return { tipo: String(r[0]), fecha: toIso_(r[1]), caption: String(r[2]),
               likes: Number(r[3]) || 0, comentarios: Number(r[4]) || 0, alcance: Number(r[5]) || 0,
               guardados: Number(r[6]) || 0, permalink: String(r[7]), thumbnail: String(r[8]) };
    });
  }
  return out;
}

function toIso_(v) {
  if (v instanceof Date) return v.toISOString();
  return String(v);
}
