// ============================================================
// TIENDA VIRTUAL — GOOGLE APPS SCRIPT BACKEND
// Fases 1 y 2: Inicialización + API Web App
// Copia este archivo completo en el editor de Apps Script
// vinculado a tu Google Sheets.
// ============================================================

// ============================================================
// SECCIÓN 1 — INICIALIZACIÓN DE LA BASE DE DATOS (FASE 1)
// ============================================================

/**
 * Ejecuta esta función UNA VEZ para crear todas las hojas
 * con sus encabezados y datos de ejemplo.
 * Si las hojas ya existen, NO las sobreescribe.
 */
function inicializarTienda() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Hoja: productos ---
    var colsProductos = [
      'id','nombre','descripcion','precio','precio_oferta',
      'categoria','subcategoria','imagen_url','stock','activo',
      'es_gancho','productos_gancho_ids','destacado','orden',
      'fecha_creacion','fecha_modificacion'
    ];
    var hProductos = crearHojaConEncabezados_(ss, 'productos', colsProductos);

    // --- Hoja: categorias ---
    var colsCategorias = [
      'id','nombre','descripcion','imagen_url','activo','orden','fecha_creacion'
    ];
    var hCategorias = crearHojaConEncabezados_(ss, 'categorias', colsCategorias);

    // --- Hoja: pedidos ---
    var colsPedidos = [
      'id','fecha','nombre_cliente','telefono','email','direccion',
      'ciudad','barrio','notas','productos_json','subtotal',
      'domicilio','total','estado','numero_whatsapp_enviado'
    ];
    crearHojaConEncabezados_(ss, 'pedidos', colsPedidos);

    // --- Hoja: configuracion (clave/valor) ---
    var hConfig = ss.getSheetByName('configuracion');
    if (!hConfig) {
      hConfig = ss.insertSheet('configuracion');
      var configHeaders = [['clave', 'valor']];
      hConfig.getRange(1, 1, 1, 2).setValues(configHeaders);
      formatearEncabezados_(hConfig, 2);
      hConfig.setFrozenRows(1);

      // Valores por defecto
      var configDefaults = [
        ['nombre_tienda', 'Mi Tienda de Lentes'],
        ['telefono_whatsapp', '573000000000'],
        ['moneda', 'COP'],
        ['mensaje_whatsapp_intro', 'Hola! Quiero realizar el siguiente pedido:'],
        ['mensaje_whatsapp_cierre', 'Por favor confirmar disponibilidad y tiempo de entrega. Gracias!'],
        ['formulario_campos', 'nombre,telefono,email,direccion,ciudad,barrio,notas'],
        ['formulario_campos_requeridos', 'nombre,telefono,direccion,ciudad'],
        ['domicilio_medellin', '5000'],
        ['domicilio_itagui', '6000'],
        ['domicilio_envigado', '6000'],
        ['domicilio_bello', '7000'],
        ['domicilio_sabaneta', '6000'],
        ['domicilio_otro', '8000'],
        ['color_primario', '#0f172a'],
        ['color_secundario', '#6366f1'],
        ['logo_url', ''],
        ['banner_url', ''],
        ['whatsapp_mensaje_producto', 'Producto: {nombre} x{cantidad} - ${precio}'],
        ['admin_token', 'cambia_este_token_ahora'],
        ['admin_email', ''],
        ['tagline_tienda', 'Los mejores lentes de contacto al mejor precio']
      ];
      hConfig.getRange(2, 1, configDefaults.length, 2).setValues(configDefaults);
      hConfig.autoResizeColumns(1, 2);
    }

    // --- Hoja: zonas_domicilio ---
    var colsZonas = ['id','nombre_zona','municipios','valor_domicilio','activo'];
    var hZonas = crearHojaConEncabezados_(ss, 'zonas_domicilio', colsZonas);
    if (hZonas.getLastRow() < 2) {
      var zonasData = [
        [generarId(), 'Medellín', 'Medellín,El Poblado,Laureles,Envigado', 5000, true],
        [generarId(), 'Área Metropolitana Norte', 'Bello,Copacabana,Girardota', 7000, true],
        [generarId(), 'Área Metropolitana Sur', 'Itagüí,Sabaneta,La Estrella,Caldas', 6000, true],
        [generarId(), 'Otro', 'Otro', 10000, true]
      ];
      hZonas.getRange(2, 1, zonasData.length, colsZonas.length).setValues(zonasData);
    }

    // --- Datos de ejemplo: categorias ---
    if (hCategorias.getLastRow() < 2) {
      var ahora = new Date().toISOString();
      var catData = [
        [generarId(), 'Lentes de Color', 'Lentes de contacto de colores para cambiar tu look', '', true, 1, ahora],
        [generarId(), 'Lentes Correctivos', 'Lentes de contacto correctivos con y sin color', '', true, 2, ahora],
        [generarId(), 'Accesorios', 'Líquidos limpiadores, estuches y accesorios para tus lentes', '', true, 3, ahora]
      ];
      hCategorias.getRange(2, 1, catData.length, colsCategorias.length).setValues(catData);
    }

    // --- Datos de ejemplo: productos ---
    if (hProductos.getLastRow() < 2) {
      var ahora = new Date().toISOString();
      var idAcc1 = generarId();
      var idAcc2 = generarId();
      var productoData = [
        // Lente de color — producto principal (no gancho)
        [
          generarId(),
          'Lentes Air Optix Colors — Azul Zafiro',
          'Lentes de contacto de color mensual con tecnología SmartShield. Cómodos todo el día. Presentación x2 unidades.',
          85000, 72000, 'Lentes de Color', 'Mensual',
          'https://images.unsplash.com/photo-1574258495973-f010dfbb5371?w=400',
          50, true, false, idAcc1 + ',' + idAcc2,
          true, 1, ahora, ahora
        ],
        // Accesorio 1 — gancho
        [
          idAcc1,
          'Líquido Limpiador Renu MultiPlus 360ml',
          'Solución multipropósito para lentes de contacto blandos. Limpia, desinfecta y lubrica. Hasta 3 meses de uso.',
          35000, '', 'Accesorios', 'Líquidos',
          'https://images.unsplash.com/photo-1559305616-3f99cd43e353?w=400',
          100, true, true, '', false, 2, ahora, ahora
        ],
        // Accesorio 2 — gancho
        [
          idAcc2,
          'Estuche Premium para Lentes de Contacto',
          'Estuche doble con espejo y pinzas incluidas. Colores surtidos. Ideal para llevar a cualquier parte.',
          12000, 9000, 'Accesorios', 'Estuches',
          'https://images.unsplash.com/photo-1586495777744-4e6232bf4400?w=400',
          200, true, true, '', false, 3, ahora, ahora
        ]
      ];
      hProductos.getRange(2, 1, productoData.length, colsProductos.length).setValues(productoData);
    }

    SpreadsheetApp.getUi().alert('✅ Tienda inicializada correctamente\n\nSe crearon las 5 hojas con datos de ejemplo.\n\n⚠️ IMPORTANTE: Cambia el admin_token en la hoja "configuracion" antes de publicar.');

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error al inicializar:\n' + e.message);
  }
}

/** Crea una hoja si no existe y le pone encabezados formateados */
function crearHojaConEncabezados_(ss, nombre, columnas) {
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.getRange(1, 1, 1, columnas.length).setValues([columnas]);
    formatearEncabezados_(hoja, columnas.length);
    hoja.setFrozenRows(1);
    hoja.autoResizeColumns(1, columnas.length);
  }
  return hoja;
}

/** Aplica formato visual a la fila de encabezados */
function formatearEncabezados_(hoja, numColumnas) {
  var rango = hoja.getRange(1, 1, 1, numColumnas);
  rango.setBackground('#1e293b');
  rango.setFontColor('#ffffff');
  rango.setFontWeight('bold');
  rango.setFontSize(11);
}

// ============================================================
// SECCIÓN 2 — FUNCIONES HELPER DE CONFIGURACIÓN
// ============================================================

/** Obtiene el valor de una clave en la hoja configuracion */
function obtenerConfiguracion(clave) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('configuracion');
    if (!hoja) return null;
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === String(clave).trim()) {
        return datos[i][1];
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

/** Guarda o actualiza una clave en la hoja configuracion */
function guardarConfiguracion(clave, valor) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('configuracion');
    if (!hoja) return false;
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === String(clave).trim()) {
        hoja.getRange(i + 1, 2).setValue(valor);
        return true;
      }
    }
    // No existe la clave, insertar nueva fila
    hoja.appendRow([clave, valor]);
    return true;
  } catch (e) {
    return false;
  }
}

/** Genera un ID único basado en timestamp + random */
function generarId() {
  return 'ID' + Date.now() + Math.random().toString(36).substr(2, 5).toUpperCase();
}

// ============================================================
// SECCIÓN 3 — PUNTOS DE ENTRADA WEB APP (FASE 2)
// ============================================================

/** Maneja las peticiones GET — sirve HTML o JSON API */
function doGet(e) {
  try {
    var params = e ? (e.parameter || {}) : {};
    var action = params.action || '';

    // Si hay action, devolver JSON API (compatibilidad con GitHub Pages)
    if (action) {
      var resultado;
      switch (action) {
        case 'getProductos':
          resultado = accion_getProductos(params);
          break;
        case 'getProducto':
          resultado = accion_getProducto(params);
          break;
        case 'getCategorias':
          resultado = accion_getCategorias(params);
          break;
        case 'getConfiguracion':
          resultado = accion_getConfiguracion(params);
          break;
        case 'getZonasDomicilio':
          resultado = accion_getZonasDomicilio(params);
          break;
        case 'calcularDomicilio':
          resultado = accion_calcularDomicilio(params);
          break;
        case 'getPedidos':
          resultado = accion_getPedidos(params);
          break;
        default:
          resultado = { success: false, error: 'Acción no reconocida: ' + action };
      }
      return crearRespuesta(resultado);
    }

    // Sin action: servir página HTML
    var page = params.page || 'tienda';
    var fileName = (page === 'admin') ? 'admin' : 'tienda';
    var title = (page === 'admin') ? 'Panel Administrador' : 'Mi Tienda';

    var html = HtmlService.createHtmlOutputFromFile(fileName)
      .setTitle(title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return html;
  } catch (err) {
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + err.message + '</p>');
  }
}

/** Maneja las peticiones POST (admin + guardar pedido público) */
function doPost(e) {
  try {
    var body = {};
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      body = e.parameter || {};
    }

    var action = body.action || (e.parameter ? e.parameter.action : '') || '';

    // Acciones que NO requieren token (pedido público)
    var accionesPublicas = ['guardarPedido'];
    var requiereToken = accionesPublicas.indexOf(action) === -1;

    if (requiereToken) {
      var tokenConfig = obtenerConfiguracion('admin_token');
      if (!tokenConfig || body.adminToken !== String(tokenConfig).trim()) {
        return crearRespuesta({ success: false, error: 'No autorizado' });
      }
    }

    var resultado;
    switch (action) {
      case 'guardarPedido':
        resultado = accion_guardarPedido(body);
        break;
      case 'guardarProducto':
        resultado = accion_guardarProducto(body);
        break;
      case 'eliminarProducto':
        resultado = accion_eliminarProducto(body);
        break;
      case 'guardarCategoria':
        resultado = accion_guardarCategoria(body);
        break;
      case 'eliminarCategoria':
        resultado = accion_eliminarCategoria(body);
        break;
      case 'guardarConfiguracion':
        resultado = accion_guardarConfiguracionAdmin(body);
        break;
      case 'getPedidos':
        resultado = accion_getPedidos(body);
        break;
      case 'actualizarEstadoPedido':
        resultado = accion_actualizarEstadoPedido(body);
        break;
      case 'guardarZona':
        resultado = accion_guardarZona(body);
        break;
      case 'eliminarZona':
        resultado = accion_eliminarZona(body);
        break;
      default:
        resultado = { success: false, error: 'Acción no reconocida: ' + action };
    }

    return crearRespuesta(resultado);
  } catch (e) {
    return crearRespuesta({ success: false, error: 'Error interno: ' + e.message });
  }
}

/** Construye la respuesta HTTP con headers CORS */
function crearRespuesta(datos) {
  var json = JSON.stringify(datos);
  var output = ContentService.createTextOutput(json);
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
// SECCIÓN 4 — ACCIONES PÚBLICAS (doGet)
// ============================================================

function accion_getProductos(params) {
  try {
    var productos = hojaToObjetos('productos');
    var soloActivos = params.soloActivos !== 'false';

    if (soloActivos) {
      productos = productos.filter(function(p) { return p.activo === true || p.activo === 'true'; });
    }

    if (params.categoria && params.categoria !== 'todos') {
      var cat = params.categoria.toLowerCase().trim();
      productos = productos.filter(function(p) {
        return String(p.categoria || '').toLowerCase().trim() === cat;
      });
    }

    if (params.soloDestacados === 'true') {
      productos = productos.filter(function(p) { return p.destacado === true; });
    }

    // Ordenar por campo 'orden'
    productos.sort(function(a, b) {
      return (Number(a.orden) || 999) - (Number(b.orden) || 999);
    });

    return { success: true, data: productos };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_getProducto(params) {
  try {
    var id = params.id || '';
    var productos = hojaToObjetos('productos');
    var producto = null;

    for (var i = 0; i < productos.length; i++) {
      if (String(productos[i].id) === String(id)) {
        producto = productos[i];
        break;
      }
    }

    if (!producto) return { success: false, error: 'Producto no encontrado' };

    // Expandir productos_gancho_ids
    var gancho_ids = [];
    if (producto.productos_gancho_ids && String(producto.productos_gancho_ids).trim() !== '') {
      gancho_ids = String(producto.productos_gancho_ids).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    }

    var productos_gancho = [];
    if (gancho_ids.length > 0) {
      for (var j = 0; j < productos.length; j++) {
        if (gancho_ids.indexOf(String(productos[j].id)) !== -1) {
          productos_gancho.push(productos[j]);
        }
      }
    }

    producto.productos_gancho = productos_gancho;
    return { success: true, data: producto };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_getCategorias(params) {
  try {
    var categorias = hojaToObjetos('categorias');
    categorias = categorias.filter(function(c) { return c.activo === true; });
    categorias.sort(function(a, b) {
      return (Number(a.orden) || 999) - (Number(b.orden) || 999);
    });
    return { success: true, data: categorias };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_getConfiguracion(params) {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('configuracion');
    if (!hoja) return { success: false, error: 'Hoja configuracion no encontrada' };

    var datos = hoja.getDataRange().getValues();
    // Claves que NO se exponen públicamente
    var clavesSensibles = ['admin_token', 'admin_email'];
    var config = {};

    for (var i = 1; i < datos.length; i++) {
      var clave = String(datos[i][0]).trim();
      var valor = datos[i][1];
      if (clave && clavesSensibles.indexOf(clave) === -1) {
        config[clave] = valor;
      }
    }

    return { success: true, data: config };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_getZonasDomicilio(params) {
  try {
    var zonas = hojaToObjetos('zonas_domicilio');
    zonas = zonas.filter(function(z) { return z.activo === true; });
    return { success: true, data: zonas };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_calcularDomicilio(params) {
  try {
    var ciudad = normalizarTexto_(params.ciudad || '');
    if (!ciudad) return { success: false, error: 'Ciudad requerida' };

    var zonas = hojaToObjetos('zonas_domicilio');
    var zonaOtro = null;
    var zonaEncontrada = null;

    for (var i = 0; i < zonas.length; i++) {
      var z = zonas[i];
      if (!z.activo) continue;
      var municipios = String(z.municipios || '').split(',').map(function(m) {
        return normalizarTexto_(m);
      });

      if (normalizarTexto_(z.nombre_zona) === 'otro' || municipios.indexOf('otro') !== -1) {
        zonaOtro = z;
      }

      if (municipios.indexOf(ciudad) !== -1) {
        zonaEncontrada = z;
        break;
      }
    }

    var zonaResultado = zonaEncontrada || zonaOtro;
    if (!zonaResultado) return { success: false, error: 'No se encontró zona de domicilio' };

    return {
      success: true,
      valor: Number(zonaResultado.valor_domicilio) || 0,
      zona: zonaResultado.nombre_zona
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_getPedidos(params) {
  try {
    var pedidos = hojaToObjetos('pedidos');

    if (params.estado && params.estado !== 'todos') {
      pedidos = pedidos.filter(function(p) {
        return String(p.estado).toLowerCase() === String(params.estado).toLowerCase();
      });
    }

    if (params.fechaDesde) {
      var desde = new Date(params.fechaDesde);
      pedidos = pedidos.filter(function(p) {
        return new Date(p.fecha) >= desde;
      });
    }

    if (params.fechaHasta) {
      var hasta = new Date(params.fechaHasta);
      hasta.setHours(23, 59, 59);
      pedidos = pedidos.filter(function(p) {
        return new Date(p.fecha) <= hasta;
      });
    }

    // Ordenar del más reciente al más antiguo
    pedidos.sort(function(a, b) {
      return new Date(b.fecha) - new Date(a.fecha);
    });

    return { success: true, data: pedidos };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// SECCIÓN 5 — ACCIONES ADMIN (doPost)
// ============================================================

function accion_guardarPedido(body) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('pedidos');
    if (!hoja) return { success: false, error: 'Hoja pedidos no encontrada' };

    var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var ahora = new Date().toISOString();
    var id = generarId();

    var pedido = {
      id: id,
      fecha: ahora,
      nombre_cliente: body.nombre_cliente || '',
      telefono: body.telefono || '',
      email: body.email || '',
      direccion: body.direccion || '',
      ciudad: body.ciudad || '',
      barrio: body.barrio || '',
      notas: body.notas || '',
      productos_json: typeof body.productos_json === 'string'
        ? body.productos_json
        : JSON.stringify(body.productos_json || []),
      subtotal: Number(body.subtotal) || 0,
      domicilio: Number(body.domicilio) || 0,
      total: Number(body.total) || 0,
      estado: 'pendiente',
      numero_whatsapp_enviado: obtenerConfiguracion('telefono_whatsapp') || ''
    };

    var fila = objetoToFila(pedido, headers);
    hoja.appendRow(fila);

    // Generar mensaje de WhatsApp
    var config = {};
    var hConfig = ss.getSheetByName('configuracion');
    if (hConfig) {
      var configData = hConfig.getDataRange().getValues();
      for (var i = 1; i < configData.length; i++) {
        config[String(configData[i][0]).trim()] = configData[i][1];
      }
    }

    var mensajeWA = generarMensajeWhatsapp(pedido, config);

    // Notificación por email si hay admin_email configurado
    try {
      var adminEmail = config['admin_email'];
      if (adminEmail && String(adminEmail).trim() !== '') {
        var productos = [];
        try { productos = JSON.parse(pedido.productos_json); } catch(e2) {}
        var listaProductosHtml = productos.map(function(p) {
          return '<li>' + p.nombre + ' x' + p.cantidad + ' — $' + formatPrecio_(p.precio) + '</li>';
        }).join('');

        var asunto = 'Nuevo pedido #' + id + ' — ' + pedido.nombre_cliente + ' — $' + formatPrecio_(pedido.total);
        var cuerpoHtml = '<h2>Nuevo pedido recibido</h2>' +
          '<p><strong>ID:</strong> ' + id + '</p>' +
          '<p><strong>Cliente:</strong> ' + pedido.nombre_cliente + '</p>' +
          '<p><strong>Teléfono:</strong> ' + pedido.telefono + '</p>' +
          '<p><strong>Email:</strong> ' + pedido.email + '</p>' +
          '<p><strong>Dirección:</strong> ' + pedido.direccion + ', ' + pedido.barrio + ', ' + pedido.ciudad + '</p>' +
          '<h3>Productos:</h3><ul>' + listaProductosHtml + '</ul>' +
          '<p><strong>Subtotal:</strong> $' + formatPrecio_(pedido.subtotal) + '</p>' +
          '<p><strong>Domicilio:</strong> $' + formatPrecio_(pedido.domicilio) + '</p>' +
          '<p><strong>TOTAL:</strong> $' + formatPrecio_(pedido.total) + '</p>' +
          (pedido.notas ? '<p><strong>Notas:</strong> ' + pedido.notas + '</p>' : '');

        MailApp.sendEmail({
          to: String(adminEmail).trim(),
          subject: asunto,
          htmlBody: cuerpoHtml
        });
      }
    } catch (mailErr) {
      // No lanzar error si el email falla
      Logger.log('Error enviando email: ' + mailErr.message);
    }

    return { success: true, id: id, mensaje_whatsapp: mensajeWA };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_guardarProducto(body) {
  try {
    return guardarEnHoja_('productos', body);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_eliminarProducto(body) {
  try {
    return eliminarDeHoja_('productos', body.id);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_guardarCategoria(body) {
  try {
    return guardarEnHoja_('categorias', body);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_eliminarCategoria(body) {
  try {
    return eliminarDeHoja_('categorias', body.id);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_guardarConfiguracionAdmin(body) {
  try {
    var pares = body.config || body;
    var omitir = ['action', 'adminToken'];
    for (var clave in pares) {
      if (omitir.indexOf(clave) === -1) {
        guardarConfiguracion(clave, pares[clave]);
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_actualizarEstadoPedido(body) {
  try {
    var estadosValidos = ['pendiente', 'confirmado', 'enviado', 'entregado', 'cancelado'];
    if (estadosValidos.indexOf(body.estado) === -1) {
      return { success: false, error: 'Estado inválido' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('pedidos');
    if (!hoja) return { success: false, error: 'Hoja pedidos no encontrada' };

    var datos = hoja.getDataRange().getValues();
    var headers = datos[0];
    var idxId = headers.indexOf('id');
    var idxEstado = headers.indexOf('estado');

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][idxId]) === String(body.id)) {
        hoja.getRange(i + 1, idxEstado + 1).setValue(body.estado);
        return { success: true };
      }
    }

    return { success: false, error: 'Pedido no encontrado' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_guardarZona(body) {
  try {
    return guardarEnHoja_('zonas_domicilio', body);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function accion_eliminarZona(body) {
  try {
    return eliminarDeHoja_('zonas_domicilio', body.id);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// SECCIÓN 6 — HELPERS GENÉRICOS
// ============================================================

/**
 * Convierte una hoja completa en array de objetos JS.
 * La fila 1 son los headers, cada fila siguiente es un objeto.
 */
function hojaToObjetos(nombreHoja) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return [];

  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return [];

  var headers = datos[0].map(function(h) { return String(h).trim(); });
  var resultado = [];

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    // Ignorar filas donde la primera columna esté vacía
    if (fila[0] === '' || fila[0] === null || fila[0] === undefined) continue;

    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = fila[j];
      // Convertir strings "TRUE"/"FALSE" a booleanos
      if (val === 'TRUE' || val === true) { val = true; }
      else if (val === 'FALSE' || val === false) { val = false; }
      // Convertir fechas a ISO string
      else if (val instanceof Date) { val = val.toISOString(); }
      // Dejar números como números
      else if (typeof val === 'number') { val = val; }
      // Todo lo demás como string (trim)
      else { val = String(val).trim(); }

      obj[headers[j]] = val;
    }
    resultado.push(obj);
  }

  return resultado;
}

/**
 * Convierte un objeto JS a array de valores en el orden de los headers.
 */
function objetoToFila(objeto, headers) {
  return headers.map(function(h) {
    var val = objeto[h];
    if (val === undefined || val === null) return '';
    if (typeof val === 'boolean') return val;
    if (typeof val === 'object') return JSON.stringify(val);
    return val;
  });
}

/**
 * Guarda o actualiza un registro en una hoja genérica.
 * Si el objeto tiene 'id' y existe en la hoja, actualiza.
 * Si no, inserta fila nueva.
 */
function guardarEnHoja_(nombreHoja, datos) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return { success: false, error: 'Hoja ' + nombreHoja + ' no encontrada' };

  var todosLosDatos = hoja.getDataRange().getValues();
  var headers = todosLosDatos[0];
  var ahora = new Date().toISOString();

  // Omitir campos de control del body
  var omitir = ['action', 'adminToken'];
  var objeto = {};
  for (var k in datos) {
    if (omitir.indexOf(k) === -1) {
      objeto[k] = datos[k];
    }
  }

  var id = objeto.id || '';
  var filaExistente = -1;

  if (id) {
    var idxId = headers.indexOf('id');
    if (idxId !== -1) {
      for (var i = 1; i < todosLosDatos.length; i++) {
        if (String(todosLosDatos[i][idxId]) === String(id)) {
          filaExistente = i + 1; // fila real en la hoja (1-indexed)
          break;
        }
      }
    }
  }

  if (filaExistente > 0) {
    // Actualizar fila existente
    objeto.fecha_modificacion = ahora;
    var fila = objetoToFila(objeto, headers);
    hoja.getRange(filaExistente, 1, 1, headers.length).setValues([fila]);
  } else {
    // Insertar nueva fila
    objeto.id = generarId();
    objeto.fecha_creacion = ahora;
    objeto.fecha_modificacion = ahora;
    var filaNew = objetoToFila(objeto, headers);
    hoja.appendRow(filaNew);
    id = objeto.id;
  }

  return { success: true, id: id };
}

/**
 * Elimina una fila de una hoja por id.
 */
function eliminarDeHoja_(nombreHoja, id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return { success: false, error: 'Hoja no encontrada' };

  var datos = hoja.getDataRange().getValues();
  var headers = datos[0];
  var idxId = headers.indexOf('id');
  if (idxId === -1) return { success: false, error: 'Columna id no encontrada' };

  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][idxId]) === String(id)) {
      hoja.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, error: 'Registro no encontrado' };
}

/**
 * Normaliza texto: minúsculas, sin tildes, sin espacios extremos.
 */
function normalizarTexto_(texto) {
  if (!texto) return '';
  return String(texto)
    .toLowerCase()
    .trim()
    .replace(/[áàäâ]/g, 'a')
    .replace(/[éèëê]/g, 'e')
    .replace(/[íìïî]/g, 'i')
    .replace(/[óòöô]/g, 'o')
    .replace(/[úùüû]/g, 'u')
    .replace(/ñ/g, 'n');
}

/**
 * Formatea un número como precio con separadores de miles.
 */
function formatPrecio_(num) {
  return Number(num || 0).toLocaleString('es-CO');
}

// ============================================================
// SECCIÓN 7 — GENERADOR DE MENSAJE WHATSAPP
// ============================================================

/**
 * Construye el mensaje de WhatsApp formateado con los datos del pedido.
 * @param {Object} pedido - Objeto pedido
 * @param {Object} config - Objeto con toda la configuración
 * @returns {string} Mensaje listo para WhatsApp
 */
function generarMensajeWhatsapp(pedido, config) {
  try {
    var intro = config['mensaje_whatsapp_intro'] || 'Hola! Quiero realizar el siguiente pedido:';
    var cierre = config['mensaje_whatsapp_cierre'] || 'Por favor confirmar disponibilidad y tiempo de entrega. Gracias!';

    // Parsear productos
    var productos = [];
    try {
      var pJson = typeof pedido.productos_json === 'string'
        ? JSON.parse(pedido.productos_json)
        : (pedido.productos_json || []);
      productos = pJson;
    } catch(e) {}

    var listaProductos = productos.map(function(p) {
      var precioUsar = p.precio_oferta && Number(p.precio_oferta) > 0 && Number(p.precio_oferta) < Number(p.precio)
        ? Number(p.precio_oferta)
        : Number(p.precio);
      return '  • ' + p.nombre + ' x' + p.cantidad + ' — $' + formatPrecio_(precioUsar);
    }).join('\n');

    var fecha = new Date(pedido.fecha);
    var fechaStr = fecha.toLocaleDateString('es-CO', {
      year: 'numeric', month: 'long', day: 'numeric',
      hour: '2-digit', minute: '2-digit'
    });

    var notas = pedido.notas && String(pedido.notas).trim() !== ''
      ? '\n📝 Notas: ' + pedido.notas
      : '';

    var email = pedido.email && String(pedido.email).trim() !== ''
      ? '\n  • Email: ' + pedido.email
      : '';

    var mensaje = intro + '\n\n' +
      '📋 *PEDIDO #' + pedido.id + '*\n' +
      '📅 ' + fechaStr + '\n\n' +
      '👤 *Datos del cliente:*\n' +
      '  • Nombre: ' + pedido.nombre_cliente + '\n' +
      '  • Teléfono: ' + pedido.telefono + email + '\n\n' +
      '📦 *Productos:*\n' + listaProductos + '\n\n' +
      '📍 *Dirección de entrega:*\n' +
      '  ' + pedido.direccion +
      (pedido.barrio ? ', ' + pedido.barrio : '') +
      ', ' + pedido.ciudad +
      notas + '\n\n' +
      '💰 *Resumen:*\n' +
      '  • Subtotal: $' + formatPrecio_(pedido.subtotal) + '\n' +
      '  • Domicilio: $' + formatPrecio_(pedido.domicilio) + '\n' +
      '  • *TOTAL: $' + formatPrecio_(pedido.total) + '*\n\n' +
      cierre;

    return mensaje;
  } catch (e) {
    return 'Error generando mensaje: ' + e.message;
  }
}

// ============================================================
// SECCIÓN 8 — FUNCIONES PUENTE PARA google.script.run
// Estas funciones son llamadas desde tienda.html y admin.html
// cuando se sirven directamente desde Apps Script.
// ============================================================

/**
 * API pública — llamada desde tienda.html vía google.script.run
 * @param {string} action - Nombre de la acción (getProductos, getCategorias, etc.)
 * @param {Object} params - Parámetros opcionales
 * @returns {Object} Resultado de la acción
 */
function gsApi(action, params) {
  params = params || {};
  switch (action) {
    case 'getProductos':      return accion_getProductos(params);
    case 'getProducto':       return accion_getProducto(params);
    case 'getCategorias':     return accion_getCategorias(params);
    case 'getConfiguracion':  return accion_getConfiguracion(params);
    case 'getZonasDomicilio': return accion_getZonasDomicilio(params);
    case 'calcularDomicilio': return accion_calcularDomicilio(params);
    case 'guardarPedido':     return accion_guardarPedido(params);
    default: return { success: false, error: 'Acción no reconocida: ' + action };
  }
}

/**
 * Validar token de admin — llamada desde admin.html vía google.script.run
 * @param {string} token - Token ingresado por el usuario
 * @returns {Object} {success: true/false}
 */
function gsAdminLogin(token) {
  try {
    var tokenConfig = obtenerConfiguracion('admin_token');
    if (!tokenConfig || String(token).trim() !== String(tokenConfig).trim()) {
      return { success: false, error: 'Token incorrecto' };
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * API admin — llamada desde admin.html vía google.script.run
 * Valida el token antes de ejecutar la acción.
 * @param {string} token - Token de admin
 * @param {string} action - Nombre de la acción
 * @param {Object} params - Parámetros de la acción
 * @returns {Object} Resultado de la acción
 */
function gsAdminApi(token, action, params) {
  try {
    var tokenConfig = obtenerConfiguracion('admin_token');
    if (!tokenConfig || String(token).trim() !== String(tokenConfig).trim()) {
      return { success: false, error: 'No autorizado' };
    }
    params = params || {};
    switch (action) {
      case 'getProductos':          return accion_getProductos(params);
      case 'getCategorias':         return accion_getCategorias(params);
      case 'getConfiguracion':      return accion_getConfiguracionCompleta_();
      case 'getZonasDomicilio':     return accion_getZonasDomicilio(params);
      case 'getPedidos':            return accion_getPedidos(params);
      case 'guardarProducto':       return accion_guardarProducto(params);
      case 'eliminarProducto':      return accion_eliminarProducto(params);
      case 'guardarCategoria':      return accion_guardarCategoria(params);
      case 'eliminarCategoria':     return accion_eliminarCategoria(params);
      case 'guardarConfiguracion':  return accion_guardarConfiguracionAdmin(params);
      case 'actualizarEstadoPedido':return accion_actualizarEstadoPedido(params);
      case 'guardarZona':           return accion_guardarZona(params);
      case 'eliminarZona':          return accion_eliminarZona(params);
      default: return { success: false, error: 'Acción no reconocida: ' + action };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Versión completa de getConfiguracion que INCLUYE campos sensibles.
 * Solo se usa desde gsAdminApi (ya validó token).
 */
function accion_getConfiguracionCompleta_() {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('configuracion');
    if (!hoja) return { success: false, error: 'Hoja configuracion no encontrada' };
    var datos = hoja.getDataRange().getValues();
    var config = {};
    for (var i = 1; i < datos.length; i++) {
      var clave = String(datos[i][0]).trim();
      if (clave) config[clave] = datos[i][1];
    }
    return { success: true, data: config };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
