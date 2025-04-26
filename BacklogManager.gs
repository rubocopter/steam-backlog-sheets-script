/**
 * ===================================================================================
 *                                BACKLOG MANAGER SCRIPT
 * ===================================================================================
 *
 * Autor: [rubocopter]
 * Versión: 3.2 - Añadida Actualización Horas Steam
 * Fecha: 2025-04-26 (Modificado)
 *
 * DESCRIPCIÓN:
 * Este script ayuda a gestionar un backlog de videojuegos en Google Sheets.
 * Permite actualizar la biblioteca desde Steam, buscar puntuaciones de Metacritic
 * (con caché), actualizar horas jugadas de juegos Steam, y fusionar duplicados.
 * Utiliza nombres de encabezado para encontrar columnas, haciéndolo más robusto.
 * Muestra el estado actual en una hoja '_ScriptStatus'.
 *
 * CONFIGURACIÓN INICIAL REQUERIDA:
 * ---------------------------------
 * 1.  **API Keys y Steam ID:** Ve a "Archivo" -> "Propiedades del proyecto" -> "Propiedades del script".
 * 2.  Añade/Verifica las siguientes propiedades:
 *     -   `STEAM_USER_ID`: Tu ID numérico de Steam de 64 bits.
 *     -   `STEAM_API_KEY`: Tu clave de API de Steam.
 *     -   `RAWG_API_KEY`: Tu clave de API de RAWG.io.
 * 3.  Guarda las propiedades.
 *
 * 4.  **Nombre Hoja Principal:** Modifica la constante `MAIN_SHEET_NAME` más abajo
 *     en este código para que coincida EXACTAMENTE con el nombre de la pestaña
 *     de tu hoja de cálculo principal (donde están los datos del backlog).
 *
 * 5.  **Encabezados de Hoja:** Asegúrate de que tu hoja principal tenga una fila
 *     de encabezado (Fila 1) con al menos los siguientes nombres de columna EXACTOS
 *     (mayúsculas/minúsculas no importan al buscar, pero el texto sí):
 *     -   Nombre
 *     -   Plataforma(s)
 *     -   Estado
 *     -   Horas jugadas
 *     -   Metacritic
 *     -   Steam App ID
 *     El script encontrará automáticamente en qué columna está cada uno.
 *     (Opcionales usados: Notas, Extra, Fecha de Finalización)
 *
 * 6.  **Autorización:** La primera vez que ejecutes una función, Google pedirá
 *     permiso. Revisa y permite el acceso.
 *
 * 7.  **Hoja de Estado:** El script creará automáticamente una hoja llamada
 *     `_ScriptStatus` para mostrar el progreso. Puedes ocultarla si lo deseas.
 *
 * FLUJO DE TRABAJO RECOMENDADO:
 * -----------------------------
 * A tener en cuenta el límite de ejecución de Google Sheets que son 6 minutos, 
 * si acaba el tiempo volver a ejecutar el paso
 
 * 1.  **`1️⃣ Actualizar Biblioteca Steam`**: Añade juegos nuevos. Corrige
 *     manualmente cualquier "ValveTestApp..." que aparezca usando el AppID.
 * 2.  **`2️⃣ Actualizar Metacritic Faltante`**: Busca scores. Puede tardar y
 *     requerir varias ejecuciones (usa caché para acelerar reanudaciones).
 * 3.  **`3️⃣ Actualizar Horas Jugadas (Steam)`**: Sincroniza el tiempo de juego con Steam.
 * 4.  **`4️⃣🧹 Fusionar Duplicados (por Nombre)`**: Combina duplicados. Puede requerir varias
 *     ejecuciones (dependiendo del tamaño de la lista).
 *
 * ===================================================================================
 */

// ==========================================================================
// ===                      CONFIGURACIÓN GLOBAL                          ===
// ==========================================================================

// --- Claves para guardar el progreso en ScriptProperties ---
const PROGRESS_KEY_METACRITIC = 'lastRow_Metacritic';
const PROGRESS_KEY_FUSION = 'lastGroup_Fusion';

// --- Nombres de Encabezado Esperados (sensible a espacios, no a mayúsculas/minúsculas al buscar) ---
const HEADER_NOMBRE = "Nombre";
const HEADER_PLATAFORMA = "Plataforma(s)";
const HEADER_ESTADO = "Estado";
const HEADER_NOTAS = "Notas"; // Opcional, usado como fallback si existe
const HEADER_EXTRA = "Extra"; // Opcional, usado como fallback si existe
const HEADER_HORAS = "Horas jugadas";
const HEADER_METACRITIC = "Metacritic";
const HEADER_FECHA_FIN = "Fecha de Finalización"; // Opcional, usado como fallback si existe
const HEADER_STEAM_ID = "Steam App ID";

// --- Nombre de la Hoja Principal (IMPORTANTE: Cambia esto al nombre real de tu hoja) ---
const MAIN_SHEET_NAME = "Backlog"; // <---- ¡¡¡ MODIFICA ESTO !!!

// --- Hoja de Estado ---
const STATUS_SHEET_NAME = "_ScriptStatus";

// ==========================================================================
// ===                      FUNCIONES DEL MENÚ                            ===
// ==========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🎮 Backlog Manager');
  menu.addItem('1️⃣ Actualizar Biblioteca Steam', 'iniciarActualizacionSteam');
  menu.addItem('2️⃣ Actualizar Metacritic Faltante', 'iniciarActualizacionDatosFaltantes');
  menu.addItem('3️⃣ Actualizar Horas Jugadas (Steam)', 'iniciarActualizarHorasSteam'); // <-- NUEVO
  menu.addItem('4️⃣🧹 Fusionar Duplicados (por Nombre)', 'iniciarFusionDuplicados'); // <-- Renumerado
  menu.addSeparator();
  const resetMenu = ui.createMenu('🔄 Reiniciar Progreso');
  resetMenu.addItem('Reiniciar Progreso de Metacritic', 'reiniciarProgresoMetacritic');
  resetMenu.addItem('Reiniciar Progreso de Fusión de Duplicados', 'reiniciarProgresoFusion');
  menu.addSubMenu(resetMenu);
  menu.addToUi();
}

// ==========================================================================
// ===                   FUNCIONES DE INICIO DE TAREAS                    ===
// ==========================================================================
function iniciarActualizacionSteam() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('🚀 Iniciando actualización de Steam...\n\n⚠️ Puede añadir juegos "ValveTestApp..." que deberás corregir manualmente.\n\nProceso rápido, no se reanuda.');
  _updateStatus("Actualizando Steam", "Iniciando...");
  try {
    const resultado = _ejecutarActualizacionSteam();
    onProcessComplete(resultado);
  } catch (error) { Logger.log("Error capturado en iniciarActualizacionSteam: " + error + "\nStack: " + error.stack); onProcessFailure(error); }
}

function iniciarActualizacionDatosFaltantes() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastRow = scriptProperties.getProperty(PROGRESS_KEY_METACRITIC);
  let message = '🔍 Iniciando búsqueda de Metacritic...\n(Usa caché para acelerar reanudaciones)\n';
  if (lastRow) { message += `🔄 Reanudando desde aprox. fila ${parseInt(lastRow) + 2}.\n`; }
  message += 'Puede tardar y ejecutarse en partes. Ver progreso en hoja `_ScriptStatus`.';
  ui.alert(message);
  _updateStatus("Actualizando Metacritic", "Iniciando...");
  try {
    const resultado = _ejecutarActualizacionDatosFaltantes();
    onProcessComplete(resultado);
  } catch (error) { Logger.log("Error capturado en iniciarActualizacionDatosFaltantes: " + error + "\nStack: " + error.stack); onProcessFailure(error, PROGRESS_KEY_METACRITIC); }
}

// --- NUEVA FUNCIÓN DE INICIO ---
function iniciarActualizarHorasSteam() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('⏱️ Iniciando actualización de Horas Jugadas de Steam...\n\nSe conectará a la API de Steam para obtener los tiempos más recientes.\n\nProceso relativamente rápido, no se reanuda.');
  _updateStatus("Actualizando Horas Steam", "Iniciando...");
  try {
    const resultado = _ejecutarActualizacionHorasSteam(); // Llamaremos a la función principal
    onProcessComplete(resultado);
  } catch (error) {
    Logger.log("Error capturado en iniciarActualizarHorasSteam: " + error + "\nStack: " + error.stack);
    onProcessFailure(error); // No necesita clave de progreso
  }
}
// --- FIN NUEVA FUNCIÓN DE INICIO ---

function iniciarFusionDuplicados() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastGroup = scriptProperties.getProperty(PROGRESS_KEY_FUSION);
  let message = '🧹 Iniciando fusión de duplicados por nombre (Optimizado)...\n(Corrige "TestApps" manualmente antes).\n';
  if (lastGroup) { message += `🔄 Reanudando post-grupo "${lastGroup}".\n`; }
  message += 'Puede requerir varias ejecuciones. Ver progreso en hoja `_ScriptStatus`.';
  ui.alert(message);
  _updateStatus("Fusionando Duplicados", "Iniciando...");
  try {
    const resultado = _ejecutarFusionDuplicadosPorNombre();
    onProcessComplete(resultado);
  } catch (error) { Logger.log("Error capturado en iniciarFusionDuplicados: " + error + "\nStack: " + error.stack); onProcessFailure(error, PROGRESS_KEY_FUSION); }
}

// ==========================================================================
// ===                 FUNCIONES DE REINICIO DE PROGRESO                  ===
// ==========================================================================
function reiniciarProgresoMetacritic() {
  try {
    PropertiesService.getScriptProperties().deleteProperty(PROGRESS_KEY_METACRITIC);
    SpreadsheetApp.getUi().alert('✅ Progreso de Metacritic reiniciado.\n(La caché existente expirará automáticamente).');
    Logger.log(`Progreso reiniciado para: ${PROGRESS_KEY_METACRITIC}. Caché no eliminada explícitamente.`);
  } catch (e) { Logger.log(`Error al reiniciar progreso Metacritic: ${e}`); SpreadsheetApp.getUi().alert(`❌ Error al reiniciar: ${e.message}`); }
}

function reiniciarProgresoFusion() {
  try { PropertiesService.getScriptProperties().deleteProperty(PROGRESS_KEY_FUSION); SpreadsheetApp.getUi().alert('✅ Progreso de fusión reiniciado.'); Logger.log(`Progreso reiniciado para: ${PROGRESS_KEY_FUSION}`); }
  catch (e) { Logger.log(`Error al reiniciar progreso Fusión: ${e}`); SpreadsheetApp.getUi().alert(`❌ Error al reiniciar: ${e.message}`); }
}
// ==========================================================================
// ===                  CALLBACKS DE ÉXITO Y FALLO                        ===
// ==========================================================================
function onProcessComplete(mensaje) {
  const mensajeFinal = mensaje || "Proceso completado."; Logger.log("Proceso completado: " + mensajeFinal);
  _updateStatus("Idle", ""); // Volver a estado inactivo
  if (mensajeFinal.toLowerCase().includes("parcialmente completada") || mensajeFinal.toLowerCase().includes("parcial")) { SpreadsheetApp.getUi().alert('⏳ Ejecución Parcial\n\n' + mensajeFinal); }
  else { SpreadsheetApp.getUi().alert('✅ Éxito\n\n' + mensajeFinal); }
}

function onProcessFailure(error, progressKey) {
  const errorObj = (error instanceof Error) ? error : new Error(String(error)); let mensajeError = "❌ Error en el script:\n\n" + errorObj.message;
  let logMessage = `Error en proceso [${progressKey || 'N/A'}]: ${errorObj.message}\nStack: ${errorObj.stack}`; Logger.log(logMessage);
  _updateStatus("Error", errorObj.message.substring(0, 150)); // Actualizar estado a Error con más detalle

  // Añadir detalles al mensaje de error según el tipo
  if (errorObj.message.includes("maximum execution time")) { mensajeError = `⏳ Tiempo excedido.\n\nProceso detenido.`; const props = PropertiesService.getScriptProperties(); if (progressKey && props.getProperty(progressKey)) { mensajeError += `\n\n✅ Progreso guardado. Re-ejecuta para continuar.`; } else if (progressKey) { mensajeError += `\n\n⚠️ No se guardó progreso. Re-ejecuta para reintentar.`; } else { mensajeError += `\n\n(Sin reanudación).`; } }
  else if (errorObj.message.includes("Identifier") && errorObj.message.includes("has already been declared")) { mensajeError += `\n\n⚠️ Error Sintaxis: Variable duplicada. Revisa el editor de scripts.`; }
  else if (errorObj.message.includes("Propiedad no configurada") || errorObj.message.includes("propiedad") && errorObj.message.includes("no configurada")) { mensajeError += `\n\n🔑 Falta una API Key o SteamID. Revisa Propiedades del Script.`; }
  else if (errorObj.message.includes("encabezado requerido no encontrado") || errorObj.message.includes("Hoja principal") || errorObj.message.includes("nombre de la hoja principal")) { mensajeError += `\n\n📄 Error de Hoja/Encabezado:\n1. Asegúrate que el nombre en la constante MAIN_SHEET_NAME ("${MAIN_SHEET_NAME}") sea EXACTO al de tu pestaña.\n2. Verifica que los encabezados requeridos (ej. "${errorObj.message.split('"')[1] || '???'}") existan en la Fila 1 de esa hoja.`; }
  else if (errorObj.message.includes("parámetros") && errorObj.message.includes("getRange")) { mensajeError += `\n\n📊 Error interno al obtener rango. Verifica que todos los encabezados requeridos (${HEADER_NOMBRE}, ${HEADER_STEAM_ID}, etc.) existan en la Fila 1 de "${MAIN_SHEET_NAME}".`; }
  else if (errorObj.message.includes("no devolvió resultado válido") || errorObj.message.includes("no devolvió un array")) { mensajeError += `\n\n📡 Error al obtener datos de la API de Steam (GetOwnedGames). Revisa tu conexión o el estado de la API de Steam. Consulta los logs para detalles.`; }
  else if (errorObj.message.includes("JSON inválido") || errorObj.message.includes("estructura JSON inesperada")) { mensajeError += `\n\n📡 Error en la respuesta de la API de Steam (GetOwnedGames). Respuesta recibida pero no se pudo procesar. Consulta los logs.`; }
  else if (errorObj.message.includes("Error Autorización") && errorObj.message.includes("API Steam")) { mensajeError += `\n\n🔑 Error de autenticación con Steam. Verifica tu API Key, Steam ID y la privacidad de tu perfil (debe ser público o 'solo amigos' con detalles de juegos públicos).`; }
  else if (errorObj.message.includes("contactando API Steam")) { mensajeError += `\n\n🌐 Error de conexión con la API de Steam (${errorObj.message.match(/\d+/)?.[0] || '?'}). Revisa tu conexión o el estado de Steam.`; }


  SpreadsheetApp.getUi().alert(mensajeError + "\n\nConsulta Logs ('Extensiones' -> 'Apps Script' -> 'Ejecuciones') y hoja '_ScriptStatus'.");
}
// ==========================================================================
// ===                      FUNCIONES AUXILIARES                          ===
// ==========================================================================
function _getProperty(key) { const v = PropertiesService.getScriptProperties().getProperty(key); if (!v) { const m = `Propiedad ${key} no configurada. Revisa Propiedades del Script.`; Logger.log(`ERROR: ${m}`); throw new Error(m); } return v; }
function _normalizeName(name) { if (!name || typeof name !== 'string') return ''; return name.toLowerCase().replace(/™|®|©|:|'|"|,|\.|\?|!|&/g, '').replace(/-/g, ' ').replace(/\s+/g, ' ').trim(); }
function _convertColumnIndexToLetter(column) { let t, l = ''; column = Math.max(1, Math.floor(column)); while (column > 0) { t = (column - 1) % 26; l = String.fromCharCode(t + 65) + l; column = (column - t - 1) / 26; } return l; }

function _getColumnIndexes(sheet) {
  if (!sheet || sheet.getLastRow() < 1) { throw new Error("Hoja inválida o vacía para buscar encabezados."); }
  const headersRaw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indexes = {};
  let foundHeaders = 0;
  headersRaw.forEach((header, index) => {
    if (header && typeof header === 'string' && header.trim() !== '') {
      indexes[header.trim().toLowerCase()] = index + 1;
      foundHeaders++;
    }
  });
  if (foundHeaders === 0) { throw new Error("No se encontraron encabezados válidos en la Fila 1."); }
  return indexes;
}

function _getRequiredColIndex(colIndexes, headerName, isRequired = true) {
  if (!colIndexes) { throw new Error("Objeto de índices de columna no válido."); }
  const index = colIndexes[headerName.toLowerCase()];
  if (!index && isRequired) { throw new Error(`Error Crítico: El encabezado requerido "${headerName}" no encontrado en la Fila 1.`); }
  return index || null;
}

function _updateStatus(task, progress) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); let statusSheet = ss.getSheetByName(STATUS_SHEET_NAME);
    if (!statusSheet) { statusSheet = ss.insertSheet(STATUS_SHEET_NAME); statusSheet.getRange("A1:B1").setValues([["Tarea Actual:", "Progreso:"]]).setFontWeight("bold"); statusSheet.setColumnWidth(1, 150); statusSheet.setColumnWidth(2, 350); SpreadsheetApp.flush(); } // Flush solo al crear
    statusSheet.getRange("A2:B2").clearContent().setValues([[task, progress]]);
    // SpreadsheetApp.flush(); // Eliminado para mejorar rendimiento UI en bucles largos
  } catch (e) { Logger.log(`Error al actualizar la hoja de estado: ${e}`); } // No relanzar
}
// ==========================================================================
// ===            FUNCIONES PRINCIPALES DE PROCESAMIENTO (_ejecutar...)   ===
// ==========================================================================

function _ejecutarActualizacionSteam() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!hoja) throw new Error(`Hoja principal "${MAIN_SHEET_NAME}" no encontrada.`);

  const colIndexes = _getColumnIndexes(hoja);
  const idxNombre = _getRequiredColIndex(colIndexes, HEADER_NOMBRE);
  const idxPlat = _getRequiredColIndex(colIndexes, HEADER_PLATAFORMA);
  const idxEstado = _getRequiredColIndex(colIndexes, HEADER_ESTADO);
  const idxHoras = _getRequiredColIndex(colIndexes, HEADER_HORAS);
  const idxMetacritic = _getRequiredColIndex(colIndexes, HEADER_METACRITIC);
  const idxSteamId = _getRequiredColIndex(colIndexes, HEADER_STEAM_ID); // Lanza error si falta

  const STEAM_ID = _getProperty('STEAM_USER_ID');
  const STEAM_API_KEY = _getProperty('STEAM_API_KEY');

  // Asegurar fila de encabezado congelada
  if (hoja.getLastRow() === 0) {
    const expectedHeaders = [HEADER_NOMBRE, HEADER_PLATAFORMA, HEADER_ESTADO, HEADER_NOTAS, HEADER_EXTRA, HEADER_HORAS, HEADER_METACRITIC, HEADER_FECHA_FIN, HEADER_STEAM_ID];
    const headerRange = hoja.getRange(1, 1, 1, Math.max(expectedHeaders.length, hoja.getMaxColumns() || expectedHeaders.length));
    headerRange.setValues([expectedHeaders.slice(0, headerRange.getNumColumns())]);
    hoja.setFrozenRows(1); SpreadsheetApp.flush(); Logger.log("Encabezados creados/congelados.");
  } else { if (hoja.getFrozenRows() < 1) { hoja.setFrozenRows(1); } }

  // Obtener IDs existentes
  const ultimaFila = hoja.getLastRow(); const appIdsExistentes = new Set();
  if (ultimaFila > 1) {
    try {
      hoja.getRange(2, idxSteamId, ultimaFila - 1, 1).getValues().forEach(r => {
        const id = String(r[0]).trim(); if (id && !isNaN(parseInt(id))) { appIdsExistentes.add(id); }
      });
    } catch (e) { Logger.log(`Error leyendo IDs existentes (Col ${idxSteamId}): ${e}`); throw new Error(`Error leyendo columna "${HEADER_STEAM_ID}". ¿Existe?`); }
  }
  Logger.log(`Encontrados ${appIdsExistentes.size} App IDs existentes.`);

  _updateStatus("Actualizando Steam", "Obteniendo biblioteca API..."); Logger.log('Obteniendo juegos Steam...');
  let juegos;
  try { juegos = _obtenerTodosLosJuegosDeSteam(STEAM_ID, STEAM_API_KEY); }
  catch (e) { Logger.log(`Error capturado directamente desde _obtenerTodos: ${e.message}`); throw e; }

  if (!Array.isArray(juegos)) { Logger.log(`ERR INESPERADO: _obtenerTodos no devolvió array. Recibido: ${juegos}`); throw new Error("_obtenerTodosLosJuegosDeSteam no devolvió resultado válido."); }

  let news = 0; const total = juegos.length; Logger.log(`Total juegos API Steam: ${total}`);
  if (total === 0) { _updateStatus("Actualizando Steam", "Completo (0 juegos)"); return "ℹ️ No se encontraron juegos Steam en la API (puede ser perfil privado)."; }

  const filasNuevas = []; const lastColNum = hoja.getLastColumn();
  juegos.forEach((juego, index) => {
    const appId = String(juego.appid); const nombre = (juego.name && String(juego.name).trim()) ? String(juego.name).trim() : `Juego Desconocido (${appId})`;
    const progressMsg = `Procesando ${index + 1} / ${total}`;
    // Actualizar estado con menos frecuencia
    if ((index + 1) % 100 === 0 || index + 1 === total) { Logger.log(`Proc. Steam: (${index + 1}/${total}) ${nombre}`); _updateStatus("Actualizando Steam", progressMsg); }
    if (!appIdsExistentes.has(appId)) {
      const horas = juego.playtime_forever ? Math.round((juego.playtime_forever / 60) * 10) / 10 : 0;
      const fila = Array(lastColNum).fill('');
      fila[idxNombre - 1] = nombre; fila[idxPlat - 1] = 'STEAM'; fila[idxEstado - 1] = 'Pendiente'; fila[idxHoras - 1] = horas; fila[idxSteamId - 1] = appId; fila[idxMetacritic - 1] = ''; // Asume que Metacritic está presente
      filasNuevas.push(fila); appIdsExistentes.add(appId); news++;
    }
  });

  if (filasNuevas.length > 0) { Logger.log(`Añadiendo ${filasNuevas.length} juegos...`); hoja.getRange(hoja.getLastRow() + 1, 1, filasNuevas.length, lastColNum).setValues(filasNuevas); Logger.log("Filas añadidas."); SpreadsheetApp.flush(); } // Flush después de añadir
  else { Logger.log("No se encontraron juegos nuevos."); }

  const numRows = hoja.getLastRow() - 1; if (numRows > 1) { _updateStatus("Actualizando Steam", "Ordenando..."); Logger.log("Ordenando..."); hoja.getRange(2, 1, numRows, lastColNum).sort({ column: idxNombre, ascending: true }); Logger.log("Hoja ordenada."); SpreadsheetApp.flush(); } // Flush después de ordenar

  const totalFin = hoja.getLastRow() - 1; const msg = `🎉 Actualización Steam OK. Añadidos: ${news} (TestApps a corregir manualmente). Total: ${totalFin}.`;
  _updateStatus("Actualizando Steam", "Completo"); Logger.log(msg); return msg;
}


function _ejecutarActualizacionDatosFaltantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const hoja = ss.getSheetByName(MAIN_SHEET_NAME); if (!hoja) throw new Error(`Hoja "${MAIN_SHEET_NAME}" no encontrada.`);
  const colIndexes = _getColumnIndexes(hoja); const idxNombre = _getRequiredColIndex(colIndexes, HEADER_NOMBRE); const idxMetacritic = _getRequiredColIndex(colIndexes, HEADER_METACRITIC);
  const RAWG_API_KEY = _getProperty('RAWG_API_KEY'); const rango = hoja.getDataRange(); const vals = rango.getValues(); const props = PropertiesService.getScriptProperties(); const cache = CacheService.getScriptCache();
  let updated = 0; const totalRows = vals.length - 1; let processed = 0; if (totalRows <= 0) { _updateStatus("Metacritic", "Completo (0)"); return "ℹ️ No hay datos."; }
  let startIdx = 1; const saved = props.getProperty(PROGRESS_KEY_METACRITIC); if (saved) { startIdx = parseInt(saved) + 1; startIdx = Math.min(startIdx, vals.length); Logger.log(`Reanudando Metacritic idx: ${startIdx} (Fila ${startIdx + 1})`); } else { Logger.log(`Iniciando Metacritic Fila 2.`); }
  Logger.log(`Procesando Metacritic filas ${startIdx + 1} a ${totalRows + 1}.`);

  for (let i = startIdx; i < vals.length; i++) {
    const row = vals[i]; const name = String(row[idxNombre - 1] || '').trim(); const meta = row[idxMetacritic - 1]; const rIdxSheet = i + 1; processed++;
    const progressMsg = `Fila ${rIdxSheet}/${vals.length}`;
    // Actualizar estado con menos frecuencia
    if (processed % 50 === 0 || i === vals.length - 1 || processed === 1) { Logger.log(`Metacritic: ${progressMsg} ('${name}')`); _updateStatus("Metacritic", progressMsg); }
    const needsMeta = (v) => v === 'N/A' || v === null || v === undefined || String(v).trim() === '';
    if (!name || name.startsWith('ValveTestApp') || !needsMeta(meta)) { if (name.startsWith('ValveTestApp') && needsMeta(meta)) { try { hoja.getRange(rIdxSheet, idxMetacritic).setValue('N/A'); } catch (e) { Logger.log(`Error setting N/A for ValveTestApp row ${rIdxSheet}: ${e}`); } } props.setProperty(PROGRESS_KEY_METACRITIC, i.toString()); continue; }
    try { const newMeta = _buscarMetacritic(name, RAWG_API_KEY, cache); if (newMeta !== 'N/A' && !isNaN(parseInt(newMeta))) { if (String(meta) !== String(newMeta)) { hoja.getRange(rIdxSheet, idxMetacritic).setValue(newMeta); updated++; } } else if (String(meta).trim() === '') { hoja.getRange(rIdxSheet, idxMetacritic).setValue('N/A'); } props.setProperty(PROGRESS_KEY_METACRITIC, i.toString()); }
    catch (e) { Logger.log(`Error grave Metacritic Fila ${rIdxSheet} ('${name}'): ${e}`); throw e; }
  }
  Logger.log(`Bucle Metacritic OK. Último índice: ${vals.length - 1}`); props.deleteProperty(PROGRESS_KEY_METACRITIC); Logger.log(`Progreso ${PROGRESS_KEY_METACRITIC} eliminado.`);
  SpreadsheetApp.flush(); const msg = `✅ Metacritic COMPLETO. Scores act/añadidos: ${updated}. Filas proc: ${processed}.`;
  _updateStatus("Metacritic", "Completo"); Logger.log(msg); return msg;
}

// --- NUEVA FUNCIÓN PRINCIPAL ---
function _ejecutarActualizacionHorasSteam() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!hoja) throw new Error(`Hoja principal "${MAIN_SHEET_NAME}" no encontrada.`);

  const colIndexes = _getColumnIndexes(hoja);
  const idxSteamId = _getRequiredColIndex(colIndexes, HEADER_STEAM_ID);
  const idxHoras = _getRequiredColIndex(colIndexes, HEADER_HORAS);
  const idxNombre = _getRequiredColIndex(colIndexes, HEADER_NOMBRE); // Para logs

  const STEAM_ID = _getProperty('STEAM_USER_ID');
  const STEAM_API_KEY = _getProperty('STEAM_API_KEY');

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) {
    _updateStatus("Actualizando Horas Steam", "Completo (0 filas)");
    return "ℹ️ No hay datos en la hoja para actualizar.";
  }

  // --- Paso 1: Obtener los datos actuales de la hoja ---
  _updateStatus("Actualizando Horas Steam", "Leyendo datos de la hoja...");
  const rangoDatos = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn());
  const valoresActuales = rangoDatos.getValues();
  Logger.log(`Leídas ${valoresActuales.length} filas de datos.`);

  // Crear un mapa para acceso rápido: steamAppId -> { rowIndex (0-based), currentHours, nombre }
  const mapaJuegosHoja = new Map();
  valoresActuales.forEach((fila, index) => {
    const appId = String(fila[idxSteamId - 1] || '').trim();
    if (appId && !isNaN(parseInt(appId))) {
      mapaJuegosHoja.set(appId, {
        rowIndex: index, // Índice basado en 0 relativo al rango leído (valoresActuales)
        currentHours: fila[idxHoras - 1],
        nombre: fila[idxNombre - 1] || `Fila ${index + 2}` // Para logs
      });
    }
  });
  Logger.log(`Encontrados ${mapaJuegosHoja.size} juegos con Steam App ID en la hoja.`);
  if (mapaJuegosHoja.size === 0) {
     _updateStatus("Actualizando Horas Steam", "Completo (0 Steam IDs)");
    return "ℹ️ No se encontraron juegos con Steam App ID en la columna correspondiente.";
  }

  // --- Paso 2: Obtener los datos actualizados de Steam ---
  _updateStatus("Actualizando Horas Steam", "Obteniendo biblioteca de Steam API...");
  Logger.log('Obteniendo datos actualizados de juegos Steam...');
  let juegosSteamApi;
  try {
    juegosSteamApi = _obtenerTodosLosJuegosDeSteam(STEAM_ID, STEAM_API_KEY);
  } catch (e) {
    Logger.log(`Error al obtener datos de Steam API: ${e.message}`);
    throw e; // Relanzar para que onProcessFailure lo capture
  }

  if (!Array.isArray(juegosSteamApi)) {
    Logger.log(`Error inesperado: _obtenerTodosLosJuegosDeSteam no devolvió un array. Recibido: ${juegosSteamApi}`);
    throw new Error("_obtenerTodosLosJuegosDeSteam no devolvió resultado válido.");
  }
  Logger.log(`API de Steam devolvió información para ${juegosSteamApi.length} juegos.`);

  // Crear un mapa para acceso rápido: appId -> playtime_forever (minutos)
  const mapaHorasSteam = new Map();
  juegosSteamApi.forEach(juego => {
    mapaHorasSteam.set(String(juego.appid), juego.playtime_forever || 0);
  });

  // --- Paso 3: Comparar y preparar actualizaciones ---
  _updateStatus("Actualizando Horas Steam", "Comparando y preparando actualizaciones...");
  let updatesNeeded = 0;
  const updatesBatch = []; // Almacenará { rangeA1: string, value: number }

  mapaJuegosHoja.forEach((infoJuegoHoja, appId) => {
    if (mapaHorasSteam.has(appId)) {
      const minutosSteam = mapaHorasSteam.get(appId);
      const horasSteamCalculadas = Math.round((minutosSteam / 60) * 10) / 10; // Horas con 1 decimal

      // Comprobar si el valor necesita actualización
      // Usar parseFloat para manejar números almacenados como texto y comparar numéricamente
      const valorActualHoras = infoJuegoHoja.currentHours;
      let necesitaUpdate = false;

      if (valorActualHoras === null || valorActualHoras === undefined || valorActualHoras === '' || typeof valorActualHoras === 'string' && valorActualHoras.trim() === '' || isNaN(parseFloat(valorActualHoras))) {
         // Si el valor actual está vacío o no es un número válido, actualizamos
         necesitaUpdate = true;
      } else {
         // Si hay un valor numérico, comparamos (con una pequeña tolerancia por si acaso)
         if (Math.abs(parseFloat(valorActualHoras) - horasSteamCalculadas) > 0.01) {
            necesitaUpdate = true;
         }
      }

      if (necesitaUpdate) {
        updatesNeeded++;
        const filaSheet = infoJuegoHoja.rowIndex + 2; // +1 por índice 0-based, +1 por cabecera
        const celdaHorasA1 = `${_convertColumnIndexToLetter(idxHoras)}${filaSheet}`;
        updatesBatch.push({ rangeA1: celdaHorasA1, value: horasSteamCalculadas });
        // Logger.log(`Actualización necesaria para ${infoJuegoHoja.nombre} (AppID: ${appId}): '${valorActualHoras}' -> ${horasSteamCalculadas}`);
      }
    } else {
       // Logger.log(`AppID ${appId} (${infoJuegoHoja.nombre}) encontrado en hoja pero no en la API de Steam actual. Se ignora.`);
       // Podría ser un juego que ya no está en la biblioteca o un ID incorrecto.
    }
  });

  Logger.log(`Se necesitan ${updatesNeeded} actualizaciones de horas.`);

  // --- Paso 4: Aplicar actualizaciones (si hay) ---
  if (updatesBatch.length > 0) {
    _updateStatus("Actualizando Horas Steam", `Aplicando ${updatesBatch.length} actualizaciones...`);
    Logger.log(`Aplicando ${updatesBatch.length} actualizaciones...`);

    // Optimizacion: Usar setValues en rangos contiguos si es posible,
    // pero dado que las actualizaciones pueden ser dispersas, la escritura celda a celda es más simple.
    // Para muchas actualizaciones (>100-200), considerar agrupar por rangos si el rendimiento se ve afectado.
    updatesBatch.forEach(update => {
      try {
        hoja.getRange(update.rangeA1).setValue(update.value);
      } catch(e) {
         Logger.log(`Error al actualizar celda ${update.rangeA1} con valor ${update.value}: ${e}`);
         // Continuar con las demás actualizaciones
      }
    });
    SpreadsheetApp.flush(); // Asegurar que los cambios se escriban
    Logger.log("Actualizaciones aplicadas.");
  } else {
    Logger.log("No se requirieron actualizaciones de horas.");
  }

  const msg = `✅ Actualización de Horas Steam OK. ${updatesNeeded} juegos actualizados. Total juegos Steam en hoja: ${mapaJuegosHoja.size}.`;
  _updateStatus("Actualizando Horas Steam", "Completo");
  Logger.log(msg);
  return msg;
}
// --- FIN NUEVA FUNCIÓN PRINCIPAL ---

function _ejecutarFusionDuplicadosPorNombre() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const hoja = ss.getSheetByName(MAIN_SHEET_NAME); if (!hoja) throw new Error(`Hoja "${MAIN_SHEET_NAME}" no encontrada.`);
  const colIndexes = _getColumnIndexes(hoja); const idxNombre = _getRequiredColIndex(colIndexes, HEADER_NOMBRE); const idxPlat = _getRequiredColIndex(colIndexes, HEADER_PLATAFORMA); const idxHoras = _getRequiredColIndex(colIndexes, HEADER_HORAS); const idxSteamId = _getRequiredColIndex(colIndexes, HEADER_STEAM_ID);
  const lastR = hoja.getLastRow(); const lastC = hoja.getLastColumn(); if (lastR <= 1) { _updateStatus("Fusión", "Completo (0)"); return "ℹ️ No hay datos."; }
  const props = PropertiesService.getScriptProperties(); const range = hoja.getRange(1, 1, lastR, lastC); const vals = range.getValues(); const lastCLetter = _convertColumnIndexToLetter(lastC);

  _updateStatus("Fusión", "Fase 1: Agrupando..."); Logger.log("F1: Agrupando..."); const groups = {};
  for (let i = 1; i < vals.length; i++) { const row = vals[i]; const name = row[idxNombre - 1]; if (!name || String(name).trim() === '' || String(name).startsWith('ValveTestApp')) continue; const normN = _normalizeName(name); const rIdx = i + 1; if (!normN) continue; if (!groups[normN]) groups[normN] = []; groups[normN].push({ rIdx: rIdx, data: { [idxPlat]: row[idxPlat - 1], [idxHoras]: row[idxHoras - 1], [idxSteamId]: row[idxSteamId - 1] }, allRowData: row }); }
  const numG = Object.keys(groups).length; Logger.log(`Agrupación OK. ${numG} nombres únicos (no TestApp).`);

  let resume = false; let lastGroup = props.getProperty(PROGRESS_KEY_FUSION); if (lastGroup) { resume = true; Logger.log(`Reanudando Fusión post: "${lastGroup}"`); } else { Logger.log("Iniciando Fusión."); }
  const toClear = []; const toUpdate = []; let proc = 0; let skipped = 0; const groupNames = Object.keys(groups).sort(); Logger.log("F2: Procesando..."); let ok = true;
  _updateStatus("Fusión", `F2: Procesando ${numG} grupos...`);

  try {
    for (const normName of groupNames) {
      const currentIdx = proc + skipped + 1; const progMsg = `Grupo ${currentIdx}/${groupNames.length} ('${normName.substring(0,30)}')`;
      // Actualizar estado con menos frecuencia
      if (proc % 50 === 0 || proc === 1 || currentIdx === groupNames.length) { _updateStatus("Fusión", progMsg); Logger.log(`Fusión: ${progMsg}`); }
      if (resume) { if (normName === lastGroup) { resume = false; Logger.log(`Reanudación encontrada en "${normName}". Continuando...`); continue; } else { skipped++; continue; } }
      const group = groups[normName]; if (group.length <= 1) { props.setProperty(PROGRESS_KEY_FUSION, normName); continue; } proc++;
      const steamR = [], nonS = []; group.forEach(it => { const sId = String(it.data[idxSteamId] || '').trim(); if (sId && !isNaN(parseInt(sId))) { steamR.push(it); } else { nonS.push(it); } }); let keep = null; let mark = []; if (steamR.length > 0) { // Priorizar Steam si existe
        keep = steamR.sort((a, b) => (parseFloat(b.data[idxHoras]) || 0) - (parseFloat(a.data[idxHoras]) || 0))[0]; // Mantener el de Steam con más horas
        group.forEach(it => { if (it.rIdx !== keep.rIdx) mark.push(it); });
      } else if (nonS.length >= 2) { // Si no hay Steam, buscar el "más completo"
        const isE = (c) => c === '' || c == 'N/A' || c == null || c == undefined || String(c).trim() === '';
        keep = nonS.sort((a, b) => a.allRowData.filter(isE).length - b.allRowData.filter(isE).length)[0]; // Mantener el con menos celdas vacías/N/A
        nonS.forEach(it => { if (it.rIdx !== keep.rIdx) mark.push(it); });
      } else { props.setProperty(PROGRESS_KEY_FUSION, normName); continue; } // Solo 1 entrada (1 steam ó 1 non-steam) o (1 steam y 1+ non-steam que ya se marcaron)
      if (keep) { // Combinar Plataformas
        const plats = new Set(); group.forEach(it => { const pS = String(it.data[idxPlat] || '').trim().toUpperCase(); if (pS) pS.split(',').map(p => p.trim()).filter(p => p).forEach(p => plats.add(p)); }); const newP = Array.from(plats).sort().join(', ');
        const keepRowData = hoja.getRange(keep.rIdx, 1, 1, lastC).getValues()[0]; // Leer datos actuales de la fila a mantener
        const curP = String(keepRowData[idxPlat - 1] || '').trim().toUpperCase().split(',').map(p => p.trim()).filter(p => p).sort().join(', ');
        if (curP !== newP) { toUpdate.push({ rIdx: keep.rIdx, cIdx: idxPlat, newV: newP }); Logger.log(` Fila ${keep.rIdx} ('${normName}'): Plataforma actualizada a "${newP}"`);}
        mark.forEach(it => toClear.push(it.rIdx)); } props.setProperty(PROGRESS_KEY_FUSION, normName);
    }
  } catch (e) { ok = false; Logger.log(`Err F2 Fusión: ${e}\nStack: ${e.stack}`); throw e; }

  let changed = false; let cleared = 0;
  if (toUpdate.length > 0 || toClear.length > 0) {
    _updateStatus("Fusión", "F3: Aplicando cambios..."); Logger.log(`F3: Aplicando ${toUpdate.length} act. y ${toClear.length} limpiezas...`); try { if (toUpdate.length > 0) { Logger.log("Actualizando plats..."); toUpdate.forEach(u => { try { hoja.getRange(u.rIdx, u.cIdx).setValue(u.newV); } catch (e) { Logger.log(`Error al actualizar Plat en fila ${u.rIdx}: ${e}`); } }); Logger.log(`${toUpdate.length} celdas act.`); SpreadsheetApp.flush(); Utilities.sleep(200); } if (toClear.length > 0) { const uRows = [...new Set(toClear)].sort((a, b) => b - a); // Ordenar descendente para evitar problemas al borrar
      cleared = uRows.length; Logger.log(`Limpiando ${cleared} filas...`);
      // Borrar filas en lugar de limpiar contenido para evitar filas vacías
      for (const rowIndex of uRows) {
        try {
            hoja.deleteRow(rowIndex);
            // Logger.log(` Fila ${rowIndex} eliminada.`);
            // Ajustar índices de 'toUpdate' si es necesario (aunque ya se aplicaron)
            // En este flujo, 'toUpdate' se aplica antes de borrar, así que no hay problema.
        } catch(e) {
            Logger.log(`Error al BORRAR fila ${rowIndex}: ${e}. Intentando limpiar contenido...`);
            try { hoja.getRange(`A${rowIndex}:${lastCLetter}${rowIndex}`).clearContent(); } catch (e2) { Logger.log(` Fallo al limpiar fila ${rowIndex}: ${e2}`); }
        }
      }
      Logger.log(`${cleared} filas eliminadas/limpiadas.`); SpreadsheetApp.flush(); Utilities.sleep(200);
      // No es necesario ordenar aquí si se borran filas, la hoja se reajusta.
      // Si solo se limpiara contenido, sí sería necesario ordenar después.
       } changed = true; } catch (e) { ok = false; Logger.log(`Err F3 Fusión: ${e}\nStack: ${e.stack}`); throw new Error(`Error al aplicar cambios: ${e.message}.`); }
  } else { Logger.log("No hubo cambios que aplicar en esta ejecución."); if (proc > 0) changed = true; /* Se procesaron grupos aunque no resultaran en cambios */ }

  const finished = ok && !resume; // Si ok es true y no estábamos en modo resumen (o ya salimos de él)
  if (finished) { props.deleteProperty(PROGRESS_KEY_FUSION); Logger.log(`Fusión completa. Progreso ${PROGRESS_KEY_FUSION} eliminado.`); const msg = `✅ Fusión COMPLETA. Filas eliminadas/limpiadas: ${cleared}. Plats act: ${toUpdate.length}. Grupos proc: ${proc + skipped}.`; _updateStatus("Fusión", "Completo"); Logger.log(msg); return msg; }
  else { const lastG = props.getProperty(PROGRESS_KEY_FUSION) || "N/A"; const msg = `⏳ Fusión PARCIAL (procesado hasta ~"${lastG.substring(0,30)}"). Cambios aplicados: ${changed ? 'Sí' : 'No'} (Limpiados: ${cleared}, Plats act: ${toUpdate.length}).\nVuelve a ejecutar para continuar.`; _updateStatus("Fusión", "Parcial (Reanudar)"); Logger.log(msg + ` (OK: ${ok}, Resume Flag: ${resume}, Último Grupo Procesado: ${lastG})`); return msg; }
}


// ==========================================================================
// ===                      LLAMADAS A APIs EXTERNAS                      ===
// ==========================================================================
function _obtenerTodosLosJuegosDeSteam(steamId, apiKey) {
  const url = `https://api.steampowered.com/IPlayerService/GetOwnedGames/v1/?key=${apiKey}&steamid=${steamId}&include_appinfo=true&include_played_free_games=true&format=json`;
  const options = { method: 'get', muteHttpExceptions: true, deadline: 45 }; // Deadline ligeramente aumentado
  let response;
  try {
    Logger.log("API Steam (GetOwnedGames)... Contactando URL: " + url.replace(apiKey, "REDACTED")); // Log URL sin API Key
    response = UrlFetchApp.fetch(url, options);
    const sc = response.getResponseCode(); const ct = response.getContentText() || "";
    Logger.log(`API Steam (GetOwnedGames)... Respuesta Recibida - Status: ${sc}, Longitud Contenido: ${ct.length}`);
    if (sc === 200) {
      let data; try { data = JSON.parse(ct); } catch (e) { Logger.log(`API Steam OK (200) pero falló JSON.parse: ${e}. Cont: ${ct.substring(0, 1000)}`); throw new Error("Respuesta API Steam OK pero JSON inválido."); }
      if (data && data.response && data.response.games !== undefined) {
        const gc = data.response.game_count || 0; if (gc === 0 && Array.isArray(data.response.games) && data.response.games.length === 0) { Logger.log("API Steam OK: 0 juegos encontrados."); return []; }
        const jv = data.response.games.filter(j => j && j.appid); Logger.log(`API Steam OK: ${gc} juegos reportados, ${jv.length} con AppID devueltos.`); return jv;
      } else { Logger.log(`API Steam OK (200) pero sin 'response' o 'response.games'. Cont: ${ct.substring(0, 500)}`); throw new Error("Respuesta API Steam OK pero estructura JSON inesperada."); }
    } else if (sc === 401 || sc === 403) { Logger.log(`Error Auth (${sc}) API Steam (GetOwnedGames). Key/ID/Perfil privado? Cont: ${ct.substring(0, 500)}`); throw new Error(`Error Autorización (${sc}) API Steam.`); }
    else { Logger.log(`Error ${sc} API Steam (GetOwnedGames). Respuesta: ${ct.substring(0, 500)}`); throw new Error(`Error ${sc} contactando API Steam.`); }
  } catch (e) {
    // Capturar errores de UrlFetchApp (timeouts, DNS, etc.)
    if (e.message.includes("timed out") || e.message.includes("Timeout")) {
        Logger.log(`Error de Timeout contactando API Steam: ${e.message}`);
        throw new Error("Timeout contactando API Steam.");
    } else if (e instanceof Error && e.message.startsWith("Error Autorización") || e.message.startsWith("Error ") && e.message.includes("contactando API Steam") || e.message.includes("JSON inválido") || e.message.includes("estructura JSON inesperada")) {
        // Si ya es un error que hemos lanzado nosotros, relanzarlo tal cual
        Logger.log(`Error gestionado dentro de _obtenerTodosLosJuegosDeSteam: ${e.message}`);
        throw e;
    } else {
        // Error genérico de conexión o de UrlFetchApp
        Logger.log(`Error NO controlado en _obtenerTodosLosJuegosDeSteam: ${e.message}\nStack: ${e.stack}`);
        throw new Error(`Fallo Crítico al obtener biblioteca de Steam: ${e.message}`);
    }
  }
}

function _buscarMetacritic(nombre, apiKey, cache) {
  const nL = nombre.replace(/™|®|©/g, '').replace(/\s+/g, ' ').trim(); if (!nL || nL.startsWith('ValveTestApp')) { return 'N/A'; }
  const cacheKey = `metacritic_${nL.toLowerCase().replace(/[^a-z0-9]/g, '')}`; const cachedValue = cache.get(cacheKey);
  if (cachedValue !== null) {
      // Logger.log(`Cache HIT para Metacritic: "${nL}" -> ${cachedValue}`);
      return cachedValue;
  }
  // Logger.log(`Cache MISS para Metacritic: "${nL}". Consultando API RAWG...`);

  const q = encodeURIComponent(nL); const url = `https://api.rawg.io/api/games?key=${apiKey}&search=${q}&page_size=1`;
  const opt = { method: 'get', muteHttpExceptions: true, deadline: 20, headers: { 'User-Agent': 'GAS-BacklogMgr/1.0' } }; let score = 'N/A'; let apiCalled = false;
  try {
    const r = UrlFetchApp.fetch(url, opt); apiCalled = true; const sc = r.getResponseCode(); const ct = r.getContentText();
    if (sc === 200) {
        let d;
        try { d = JSON.parse(ct); } catch (e) { Logger.log(`RAWG OK(200) pero JSON inválido para "${nL}": ${e}. Cont: ${ct.substring(0,500)}`); d = null; }
        if (d && d.results && d.results.length > 0) {
            const firstResult = d.results[0];
            // Comprobar coincidencia de nombre (simple, para evitar resultados muy dispares)
            const apiNameNorm = _normalizeName(firstResult.name || "");
            const searchNameNorm = _normalizeName(nL);
            if (apiNameNorm.includes(searchNameNorm) || searchNameNorm.includes(apiNameNorm)) { // Coincidencia razonable
                 if (firstResult.metacritic !== null && firstResult.metacritic !== undefined && !isNaN(parseInt(firstResult.metacritic))) {
                    score = parseInt(firstResult.metacritic); // Asegurar que es número
                 } else {
                    // Logger.log(`RAWG OK para "${nL}", pero sin Metacritic score (o no numérico): ${firstResult.metacritic}`);
                 }
            } else {
                // Logger.log(`RAWG encontró "${firstResult.name}" para "${nL}", pero no coincide lo suficiente. Ignorando.`);
            }
        } else {
             // Logger.log(`RAWG OK(200) pero sin resultados para "${nL}".`);
        }
    }
    else if (sc === 404) {/*Logger.log(`RAWG 404 Not Found para: "${nL}".`);*/ }
    else if (sc === 429) { Logger.log(`RAWG 429 Too Many Requests para "${nL}". Esperando 10s...`); Utilities.sleep(10000); apiCalled = false; /* No poner en caché ni pausar extra */ }
    else { Logger.log(`Error ${sc} RAWG para "${nL}". Respuesta:${ct.substring(0, 300)}`); }
  } catch (e) { Logger.log(`Error conexión/fetch RAWG para "${nL}": ${e}`); apiCalled = true; /* Asumir que se intentó llamar */ }

  // Guardar en caché el resultado ('N/A' o el número como string)
  cache.put(cacheKey, String(score), 21600); // Guardar en caché por 6h
  // Logger.log(`Cache SET para Metacritic: "${nL}" -> ${String(score)}`);

  if (apiCalled) { Utilities.sleep(1100); } // Pausa estándar SOLO si se llamó a la API RAWG (evitar rate limiting)
  return score;
}
