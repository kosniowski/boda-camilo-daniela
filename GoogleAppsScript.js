// ═══════════════════════════════════════════════════════════════
//  🌿 GOOGLE APPS SCRIPT — Boda Camilo & Daniela
//  Pega este código en script.google.com (nuevo proyecto)
//  y despliégalo como "Web App" (acceso: Cualquiera)
// ═══════════════════════════════════════════════════════════════

// ► Reemplaza con el ID de tu Google Sheets
//   (está en la URL: docs.google.com/spreadsheets/d/ESTE_ID/edit)
const SPREADSHEET_ID = "TU_SPREADSHEET_ID_AQUI";

const SS  = SpreadsheetApp.openById(SPREADSHEET_ID);

// Nombres de las hojas (tabs) en tu Sheets
const SHEET_GROUPS        = "Grupos";
const SHEET_CONFIRMATIONS = "Confirmaciones";

// ───────────────────────────────────────────
//  GET  — leer datos
// ───────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === "getGroups") {
      result = { groups: getGroups() };

    } else if (action === "checkConfirmed") {
      const familyId = e.parameter.familyId;
      result = { confirmed: isConfirmed(familyId) };

    } else if (action === "getConfirmations") {
      result = { confirmations: getAllConfirmations() };

    } else {
      result = { error: "Acción no válida" };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────────────────────
//  POST  — guardar datos
// ───────────────────────────────────────────
function doPost(e) {
  let body, result;
  try {
    body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "confirm") {
      if (isConfirmed(body.familyId)) {
        result = { success: false, message: "Ya confirmado" };
      } else {
        saveConfirmation(body);
        result = { success: true };
      }

    } else if (action === "uploadGroups") {
      uploadGroups(body.groups);
      result = { success: true, count: body.groups.length };

    } else {
      result = { error: "Acción no válida" };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────────────────────
//  FUNCIONES INTERNAS
// ───────────────────────────────────────────

function getGroups() {
  ensureSheet(SHEET_GROUPS, ["id", "name", "maxGuests"]);
  const sheet = SS.getSheetByName(SHEET_GROUPS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    id:        String(row[0]),
    name:      String(row[1]),
    maxGuests: parseInt(row[2]) || 2,
  }));
}

function isConfirmed(familyId) {
  ensureSheet(SHEET_CONFIRMATIONS, ["familyId","familyName","guestCount","phone","email","timestamp"]);
  const sheet = SS.getSheetByName(SHEET_CONFIRMATIONS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  return data.slice(1).some(row => String(row[0]) === String(familyId));
}

function getAllConfirmations() {
  ensureSheet(SHEET_CONFIRMATIONS, ["familyId","familyName","guestCount","phone","email","timestamp"]);
  const sheet = SS.getSheetByName(SHEET_CONFIRMATIONS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    familyId:   String(row[0]),
    familyName: String(row[1]),
    guestCount: parseInt(row[2]) || 0,
    phone:      String(row[3]),
    email:      String(row[4]),
    timestamp:  String(row[5]),
  }));
}

function saveConfirmation(data) {
  ensureSheet(SHEET_CONFIRMATIONS, ["familyId","familyName","guestCount","phone","email","timestamp"]);
  const sheet = SS.getSheetByName(SHEET_CONFIRMATIONS);
  sheet.appendRow([
    data.familyId   || "",
    data.familyName || "",
    data.guestCount || 0,
    data.phone      || "",
    data.email      || "",
    new Date().toISOString(),
  ]);
}

function uploadGroups(groups) {
  ensureSheet(SHEET_GROUPS, ["id", "name", "maxGuests"]);
  const sheet = SS.getSheetByName(SHEET_GROUPS);

  // Limpiar datos anteriores (menos encabezado)
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  // Insertar nuevos grupos
  groups.forEach(g => {
    sheet.appendRow([g.id, g.name, g.maxGuests]);
  });
}

// Crea la hoja si no existe y pone el encabezado
function ensureSheet(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#5C6B2E")
      .setFontColor("white")
      .setFontWeight("bold");
  }
}

// ═══════════════════════════════════════════════════════════════
//  INSTRUCCIONES DE CONFIGURACIÓN:
//
//  1. Ve a  https://script.google.com/  → Nuevo proyecto
//  2. Pega este código completo
//  3. Reemplaza SPREADSHEET_ID con el ID de tu Google Sheets
//  4. Haz clic en "Implementar" → "Nueva implementación"
//  5. Tipo: "Aplicación web"
//     - Ejecutar como: Yo (tu cuenta)
//     - Quién tiene acceso: Cualquier usuario
//  6. Autoriza los permisos cuando te lo pida
//  7. Copia la URL que te da y pégala en WeddingRSVP.jsx
//     como valor de SCRIPT_URL
//
//  ¡Listo! El sistema guardará todo en Google Sheets automáticamente.
// ═══════════════════════════════════════════════════════════════
