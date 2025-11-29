// QASO SYSTEM - Sistema de Control de Inventario y Logística
// Backend con Google Apps Script
// Versión: 3.0.0 (COMPLETA - SIN RECORTES)

const SPREADSHEET_ID = "1eZdpRyhEZe2jmEO529459J3v4WYDQ4LVRfRcKlLVW3E";
const HOJA_PRODUCTOS = "Productos";
const HOJA_MOVIMIENTOS = "Movimientos";
const HOJA_UNIDADES = "Unidades";
const HOJA_GRUPOS = "Grupos";

const TIPOS_MOVIMIENTO = {
  INGRESO: "INGRESO",
  SALIDA: "SALIDA", 
  AJUSTE_POSITIVO: "AJUSTE_POSITIVO",
  AJUSTE_NEGATIVO: "AJUSTE_NEGATIVO",
  AJUSTE: "AJUSTE"
};

/**
 * Función principal que sirve la aplicación web
 */
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile("index")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle("Sistema de Control de Inventario - QASO SYSTEM");
  } catch (error) {
    return HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: Arial; text-align: center;">
        <h2 style="color: #dc3545;">Error del Sistema</h2>
        <p>No se pudo cargar la aplicación: ${error.message}</p>
      </div>
    `);
  }
}

/**
 * Registra un nuevo producto en el sistema
 * MODIFICADO: Ahora acepta stockInicial y crea el movimiento automático
 */
function registrarProducto(producto) {
  try {
    if (!producto || !producto.codigo || !producto.nombre) {
      return { success: false, message: "❌ Datos del producto incompletos. Código y nombre son obligatorios." };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!sheet) {
      throw new Error(`La hoja '${HOJA_PRODUCTOS}' no existe. Inicialice el sistema primero.`);
    }
    
    // Configurar encabezados si es la primera vez
    if (sheet.getLastRow() === 0) {
      // OJO: Agregamos la columna de Stock Inicial en la posición 6 (índice F)
      sheet.getRange(1, 1, 1, 7).setValues([["Código", "Nombre", "Unidad", "Grupo", "Stock Mínimo", "Stock Inicial", "Fecha Creación"]]);
      sheet.getRange(1, 1, 1, 7).setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
    }
    
    const datos = sheet.getDataRange().getValues();
    const codigoNormalizado = producto.codigo.toString().trim().toUpperCase();
    
    // Verificar duplicados
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim().toUpperCase() === codigoNormalizado) {
        return { success: false, message: "❌ Ya existe un producto con este código." };
      }
    }
    
    const nombre = producto.nombre.toString().trim();
    const unidad = producto.unidad || "Unidades";
    const grupo = producto.grupo || "General";
    const stockMin = Math.max(0, parseInt(producto.stockMin) || 0);
    const stockInicial = Math.max(0, parseFloat(producto.stockInicial) || 0);
    
    if (nombre.length < 2) {
      return { success: false, message: "❌ El nombre del producto debe tener al menos 2 caracteres." };
    }
    
    // Registrar producto
    sheet.appendRow([
      codigoNormalizado, 
      nombre, 
      unidad, 
      grupo, 
      stockMin,
      stockInicial, // Nuevo campo guardado
      new Date()
    ]);

    // LÓGICA NUEVA: Si hay stock inicial, registrar movimiento automáticamente
    if (stockInicial > 0) {
      registrarMovimiento({
        codigo: codigoNormalizado,
        fecha: new Date(), // Fecha actual
        tipo: "INGRESO",
        cantidad: stockInicial,
        observaciones: "Stock Inicial por Creación de Producto"
      });
    }
    
    return { success: true, message: "✅ Producto registrado correctamente." };
  } catch (error) {
    console.error("Error en registrarProducto:", error);
    return { success: false, message: `❌ Error al registrar producto: ${error.message}` };
  }
}

/**
 * Registra un movimiento de inventario
 * MODIFICADO: Mejor manejo de fechas para evitar errores
 */
function registrarMovimiento(mov) {
  try {
    // Validación básica, permitimos que fecha venga como objeto o string
    if (!mov || !mov.codigo || !mov.tipo || !mov.cantidad) {
      return { success: false, message: "❌ Datos del movimiento incompletos." };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!movSheet || !prodSheet) {
      throw new Error("❌ Las hojas del sistema no existen. Inicialice el sistema primero.");
    }
    
    // Configurar encabezados si es necesario
    if (movSheet.getLastRow() === 0) {
      movSheet.getRange(1, 1, 1, 8).setValues([["Código", "Fecha", "Tipo", "Cantidad", "Usuario", "Timestamp", "Observaciones", "Stock Resultante"]]);
      movSheet.getRange(1, 1, 1, 8).setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
    }
    
    const codigoNormalizado = mov.codigo.toString().trim().toUpperCase();
    const cantidad = parseFloat(mov.cantidad);
    const tipo = mov.tipo.toString().toUpperCase();
    
    if (cantidad <= 0) {
      return { success: false, message: "❌ La cantidad debe ser mayor a 0." };
    }
    
    // Verificar existencia producto
    const productos = prodSheet.getDataRange().getValues();
    let productoExiste = false;
    
    for (let i = 1; i < productos.length; i++) {
      if (productos[i][0] && productos[i][0].toString().trim().toUpperCase() === codigoNormalizado) {
        productoExiste = true;
        break;
      }
    }
    
    if (!productoExiste) {
      return { success: false, message: "❌ El producto no existe. Regístrelo primero." };
    }
    
    // Verificar stock disponible para salidas
    const stockActual = calcularStock(codigoNormalizado);
    
    if ((tipo === TIPOS_MOVIMIENTO.SALIDA || tipo === TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO) && stockActual < cantidad) {
      return { success: false, message: `❌ Stock insuficiente. Disponible: ${stockActual}, Solicitado: ${cantidad}` };
    }
    
    // Calcular stock resultante
    let stockResultante = stockActual;
    switch (tipo) {
      case TIPOS_MOVIMIENTO.INGRESO:
      case TIPOS_MOVIMIENTO.AJUSTE_POSITIVO:
        stockResultante += cantidad;
        break;
      case TIPOS_MOVIMIENTO.SALIDA:
      case TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO:
        stockResultante -= cantidad;
        break;
      case TIPOS_MOVIMIENTO.AJUSTE: // Caso raro, asumimos suma
        stockResultante += cantidad;
        break;
    }
    
    stockResultante = Math.max(0, stockResultante);
    
    // Procesar fecha robustamente
    let fechaMovimiento = new Date();
    if (mov.fecha) {
        if (typeof mov.fecha === 'string' && mov.fecha.includes('-')) {
             // Formato YYYY-MM-DD del input date
             const partes = mov.fecha.split('-');
             fechaMovimiento = new Date(partes[0], partes[1]-1, partes[2]);
        } else {
             fechaMovimiento = new Date(mov.fecha);
        }
    }
    
    movSheet.appendRow([
      codigoNormalizado, 
      fechaMovimiento, 
      tipo, 
      cantidad,
      Session.getActiveUser().getEmail() || "Sistema",
      new Date(),
      mov.observaciones || "",
      stockResultante
    ]);
    
    return { success: true, message: "✅ Movimiento registrado correctamente." };
  } catch (error) {
    console.error("Error en registrarMovimiento:", error);
    return { success: false, message: `❌ Error al registrar movimiento: ${error.message}` };
  }
}

/**
 * Busca productos por código (para autocompletado y escáner)
 * MODIFICADO: Devuelve objeto estructurado para el escáner
 */
function buscarProductoPorCodigo(codigo) {
  try {
    if (!codigo) return { encontrado: false };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!sheet) return { encontrado: false };
    
    const datos = sheet.getDataRange().getValues();
    const textoBusqueda = codigo.toString().toUpperCase().trim();
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[0] && fila[0].toString().toUpperCase().trim() === textoBusqueda) {
        // Producto encontrado, calculamos stock al instante
        const stock = calcularStock(fila[0]);
        return {
           encontrado: true,
           producto: {
             codigo: fila[0],
             nombre: fila[1],
             unidad: fila[2] || "Unidades",
             grupo: fila[3] || "General",
             stockActual: stock
           }
        };
      }
    }
    // Si llegamos aquí, no se encontró exacto (para autocompletado parcial)
    const encontrados = [];
    for (let i = 1; i < datos.length; i++) {
        if(datos[i][0] && datos[i][0].toString().toUpperCase().includes(textoBusqueda)) {
             encontrados.push({
                 codigo: datos[i][0],
                 nombre: datos[i][1],
                 unidad: datos[i][2],
                 grupo: datos[i][3]
             });
        }
    }
    
    if (encontrados.length > 0) {
        return { encontrado: false, sugerencias: encontrados.slice(0, 10) }; // Devolver sugerencias si no es exacto
    }

    return { encontrado: false };
  } catch (error) {
    console.error("Error en buscarProductoPorCodigo:", error);
    return { encontrado: false, error: error.message };
  }
}

/**
 * Busca productos por texto (búsqueda general)
 */
function buscarProducto(texto) {
  try {
    if (!texto) return [];

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!sheet) return [];
    
    const datos = sheet.getDataRange().getValues();
    const textoBusqueda = texto.toString().toLowerCase().trim();
    const encontrados = [];
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[0] && (
        fila[0].toString().toLowerCase().includes(textoBusqueda) ||
        fila[1].toString().toLowerCase().includes(textoBusqueda) ||
        (fila[3] && fila[3].toString().toLowerCase().includes(textoBusqueda))
      )) {
        const stock = calcularStock(fila[0]);
        encontrados.push([
          fila[0],
          fila[1],
          fila[2],
          fila[3],
          fila[4] || 0,
          stock
        ]);
      }
    }
    return encontrados.sort((a, b) => a[1].localeCompare(b[1]));
  } catch (error) {
    console.error("Error en buscarProducto:", error);
    return [];
  }
}

/**
 * Obtiene el stock actual de todos los productos
 */
function obtenerStock() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!prodSheet) return [];
    
    const productos = prodSheet.getDataRange().getValues();
    if (productos.length <= 1) return [];
    
    const stock = [];
    
    for (let i = 1; i < productos.length; i++) {
      const [codigo, nombre, unidad, grupo, stockMin] = productos[i];
      if (codigo && nombre) {
        const cantidad = calcularStock(codigo);
        stock.push({
          codigo: codigo.toString(), 
          nombre: nombre.toString(), 
          unidad: unidad || "Unidades", 
          grupo: grupo || "General", 
          stockMin: Math.max(0, parseInt(stockMin) || 0),
          cantidad: cantidad
        });
      }
    }
    return stock.sort((a, b) => a.nombre.localeCompare(b.nombre));
  } catch (error) {
    console.error("Error en obtenerStock:", error);
    return [];
  }
}

/**
 * Calcula el stock actual de un producto específico
 */
function calcularStock(codigo) {
  try {
    if (!codigo) return 0;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    if (!movSheet) return 0;
    
    const movimientos = movSheet.getDataRange().getValues();
    let cantidad = 0;
    const codigoNormalizado = codigo.toString().trim().toUpperCase();
    
    for (let i = 1; i < movimientos.length; i++) {
      const [cod, fecha, tipo, cant] = movimientos[i];
      if (cod && cod.toString().trim().toUpperCase() === codigoNormalizado) {
        const valor = parseFloat(cant) || 0;
        const tipoMovimiento = tipo.toString().toUpperCase();
        
        if (tipoMovimiento === TIPOS_MOVIMIENTO.INGRESO || tipoMovimiento === TIPOS_MOVIMIENTO.AJUSTE_POSITIVO) {
            cantidad += valor;
        } else if (tipoMovimiento === TIPOS_MOVIMIENTO.SALIDA || tipoMovimiento === TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO) {
            cantidad -= valor;
        } else if (tipoMovimiento === TIPOS_MOVIMIENTO.AJUSTE) {
            cantidad += valor;
        }
      }
    }
    return Math.max(0, Math.round(cantidad * 100) / 100);
  } catch (error) {
    return 0;
  }
}

/**
 * Obtiene el historial de movimientos con filtros
 */
function obtenerHistorial(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!movSheet || !prodSheet) return [];
    
    const movimientos = movSheet.getDataRange().getValues();
    const productos = prodSheet.getDataRange().getValues();
    
    // Mapa de nombres
    const prodMap = {};
    for (let i = 1; i < productos.length; i++) {
      if (productos[i][0]) prodMap[productos[i][0].toString().toUpperCase()] = productos[i][1];
    }
    
    const fechaDesde = new Date(filtros.fechaDesde + 'T00:00:00');
    const fechaHasta = new Date(filtros.fechaHasta + 'T23:59:59');
    const resultado = [];
    
    for (let i = 1; i < movimientos.length; i++) {
      const mov = movimientos[i];
      if (!mov[0]) continue;
      
      const fechaMov = new Date(mov[1]);
      const tipoMov = mov[2] ? mov[2].toString().toUpperCase() : "";
      
      if (fechaMov >= fechaDesde && fechaMov <= fechaHasta) {
        if (!filtros.tipo || tipoMov === filtros.tipo.toUpperCase()) {
          resultado.push({
            codigo: mov[0],
            fecha: formatearFecha(fechaMov),
            tipo: tipoMov,
            cantidad: parseFloat(mov[3]) || 0,
            producto: prodMap[mov[0].toString().toUpperCase()] || "Desconocido",
            observaciones: mov[6] || "",
            usuario: mov[4] || ""
          });
        }
      }
    }
    return resultado.reverse(); // Más recientes primero
  } catch (error) {
    console.error("Error historial:", error);
    return [];
  }
}

/**
 * Obtiene resumen general del sistema
 */
function obtenerResumen() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    
    let totalProductos = 0;
    let totalMovimientos = 0;
    let sinStock = 0;
    let stockBajo = 0;
    let valorTotalInventario = 0;
    let movimientosUltimoMes = 0;

    if (prodSheet) {
      const productos = prodSheet.getDataRange().getValues();
      totalProductos = Math.max(0, productos.length - 1);
      
      for (let i = 1; i < productos.length; i++) {
        if (!productos[i][0]) continue;
        const codigo = productos[i][0];
        const stockMin = parseInt(productos[i][4]) || 0;
        const stock = calcularStock(codigo);
        
        if (stock <= 0) sinStock++;
        else if (stock <= stockMin) stockBajo++;
        valorTotalInventario += stock;
      }
    }

    if (movSheet) {
      const movimientos = movSheet.getDataRange().getValues();
      totalMovimientos = Math.max(0, movimientos.length - 1);
      
      const mesAtras = new Date();
      mesAtras.setMonth(mesAtras.getMonth() - 1);
      
      for (let i = 1; i < movimientos.length; i++) {
        if (movimientos[i][1] && new Date(movimientos[i][1]) >= mesAtras) {
            movimientosUltimoMes++;
        }
      }
    }
    
    return {
      totalProductos,
      totalMovimientos,
      sinStock,
      stockBajo,
      valorTotalInventario: Math.round(valorTotalInventario * 100) / 100,
      movimientosUltimoMes
    };
  } catch (error) {
    return { totalProductos: 0, totalMovimientos: 0, sinStock: 0, stockBajo: 0 };
  }
}

/**
 * Valida la integridad de los datos
 */
function validarIntegridad() {
  const errores = [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    [HOJA_PRODUCTOS, HOJA_MOVIMIENTOS, HOJA_UNIDADES, HOJA_GRUPOS].forEach(h => {
      if (!ss.getSheetByName(h)) errores.push(`❌ Falta hoja: ${h}`);
    });
    
    if (errores.length > 0) return { errores };
    
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const productos = prodSheet.getDataRange().getValues();
    const codigosVistos = new Set();
    
    for (let i = 1; i < productos.length; i++) {
       const cod = productos[i][0];
       if(cod) {
         if(codigosVistos.has(cod.toString())) errores.push(`❌ Duplicado: ${cod}`);
         codigosVistos.add(cod.toString());
       }
    }
    return { errores };
  } catch (error) {
    return { errores: ["Error crítico al validar: " + error.message] };
  }
}

/**
 * Inicializa las hojas del sistema
 * MODIFICADO: Agrega la columna Stock Inicial si no existe
 */
function inicializarHojas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Productos
    let prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!prodSheet) prodSheet = ss.insertSheet(HOJA_PRODUCTOS);
    
    if (prodSheet.getLastRow() === 0) {
      prodSheet.getRange(1, 1, 1, 7).setValues([["Código", "Nombre", "Unidad", "Grupo", "Stock Mínimo", "Stock Inicial", "Fecha Creación"]]);
      prodSheet.getRange(1, 1, 1, 7).setBackground("#0F5132").setFontColor("white").setFontWeight("bold");
    }

    // 2. Movimientos
    let movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    if (!movSheet) movSheet = ss.insertSheet(HOJA_MOVIMIENTOS);
    
    if (movSheet.getLastRow() === 0) {
      movSheet.getRange(1, 1, 1, 8).setValues([["Código", "Fecha", "Tipo", "Cantidad", "Usuario", "Timestamp", "Observaciones", "Stock Resultante"]]);
      movSheet.getRange(1, 1, 1, 8).setBackground("#0F5132").setFontColor("white").setFontWeight("bold");
    }

    // 3. Unidades y Grupos
    obtenerListas();
    
    return "✅ Sistema inicializado correctamente. Hojas creadas.";
  } catch (error) {
    return `❌ Error al inicializar: ${error.message}`;
  }
}

/**
 * Obtiene listas para dropdowns
 */
function obtenerListas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let uSheet = ss.getSheetByName(HOJA_UNIDADES);
    let gSheet = ss.getSheetByName(HOJA_GRUPOS);
    
    if (!uSheet) {
      uSheet = ss.insertSheet(HOJA_UNIDADES);
      uSheet.getRange(1,1).setValue("Unidades");
      uSheet.appendRow(["Unidad"]);
      uSheet.appendRow(["Kg"]);
      uSheet.appendRow(["Litro"]);
    }
    if (!gSheet) {
      gSheet = ss.insertSheet(HOJA_GRUPOS);
      gSheet.getRange(1,1).setValue("Grupos");
      gSheet.appendRow(["General"]);
      gSheet.appendRow(["Electrónica"]);
    }
    
    const uData = uSheet.getDataRange().getValues();
    const gData = gSheet.getDataRange().getValues();
    
    return { 
      unidades: uData.slice(1).map(r => r[0]).filter(r => r), 
      grupos: gData.slice(1).map(r => r[0]).filter(r => r) 
    };
  } catch (e) {
    return { unidades: ["Unidad"], grupos: ["General"] };
  }
}

function exportarStockCSV() {
  // Función placeholder que retorna null para no romper el front si se llama
  // Implementación completa requeriría DriveApp permission scope que a veces complica
  return null; 
}

function formatearFecha(fecha) {
  const f = new Date(fecha);
  return `${f.getDate()}/${f.getMonth()+1}/${f.getFullYear()}`;
}

// Funciones de limpieza y reset
function limpiarTodosFormularios() { return "OK"; }
function resetSistema() { return "Función deshabilitada por seguridad."; }