// ============================================================
// CATÁLOGO DE ROPA — Google Apps Script
// Pegá este código en Extensiones → Apps Script de tu Sheet
// ============================================================

const SHEET_NAME = "Hoja 1"; // Cambiá si tu hoja tiene otro nombre

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = JSON.parse(e.postData.contents);

    // Si es una actualización de producto existente
    if (data.accion === "actualizar") {
      const fila = data.fila;
      sheet.getRange(fila, 1, 1, 6).setValues([[
        data.nombre,
        data.precio,
        data.categoria,
        data.stock,
        data.descripcion || "",
        data.estado || "activo"
      ]]);
      return response({ ok: true, mensaje: "Producto actualizado" });
    }

    // Si es eliminar
    if (data.accion === "eliminar") {
      sheet.deleteRow(data.fila);
      return response({ ok: true, mensaje: "Producto eliminado" });
    }

    // Si es agregar (por defecto)
    sheet.appendRow([
      data.nombre,
      data.precio,
      data.categoria,
      data.stock,
      data.descripcion || "",
      data.estado || "activo"
    ]);

    return response({ ok: true, mensaje: "Producto agregado" });

  } catch (err) {
    return response({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const rows = sheet.getDataRange().getValues();

    // La primera fila son los encabezados, saltearla
    const productos = rows.slice(1).map((row, index) => ({
      fila: index + 2, // número de fila real en Sheets (empieza en 2)
      nombre:      row[0] || "",
      precio:      row[1] || 0,
      categoria:   row[2] || "",
      stock:       row[3] || 0,
      descripcion: row[4] || "",
      estado:      row[5] || "activo"
    })).filter(p => p.nombre !== ""); // ignorar filas vacías

    return response({ ok: true, productos });

  } catch (err) {
    return response({ ok: false, error: err.toString() });
  }
}

function response(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
