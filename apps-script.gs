// ============================================================
//  EPUE CONNECT — Google Apps Script
//  Pega este código en: Google Sheet > Extensiones > Apps Script
//  Luego: Implementar > Nueva implementación > Aplicación web
//  Ejecutar como: Tu cuenta | Acceso: Cualquier usuario
// ============================================================

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "addAttendance") {
      // Añade una sola fila al sheet del mes correspondiente
      addAttendanceRow(body.rec);
      return ok();
    }
    if (action === "syncAll") {
      // Sincronización completa de todos los datos
      syncAttendance(body.att || []);
      syncMembers(body.mem || [], body.mem || []);
      syncVisitors(body.vis || []);
      return ok();
    }
    return error("Unknown action");
  } catch (err) {
    return error(err.message);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: "EPUE Connect API activa ✅" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
function error(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ASISTENCIA: una fila por servicio, organizado por mes ──────────────────
// Estructura en Google Sheets:
//   Hoja "Enero 2025"  → todos los servicios de enero 2025
//   Hoja "Febrero 2025" → todos los servicios de febrero 2025
//   ...
//   Hoja "Asistencia_Completa" → todos los registros acumulados

function getMonthSheetName(dateStr) {
  // dateStr = "2025-01-12" → "Enero 2025"
  const months = {
    "01":"Enero","02":"Febrero","03":"Marzo","04":"Abril",
    "05":"Mayo","06":"Junio","07":"Julio","08":"Agosto",
    "09":"Septiembre","10":"Octubre","11":"Noviembre","12":"Diciembre"
  };
  const parts = dateStr.split("-");
  return `${months[parts[1]] || parts[1]} ${parts[0]}`;
}

function getOrCreateSheet(name, headerRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headerRow) {
      sheet.appendRow(headerRow);
      styleHeader(sheet, headerRow.length, "#00338D");
    }
  }
  return sheet;
}

function styleHeader(sheet, numCols, color) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground(color || "#00338D")
       .setFontColor("#ffffff")
       .setFontWeight("bold")
       .setFontSize(11);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, numCols);
}

const ATT_HEADERS = ["Fecha", "Niños", "Jóvenes", "Adultos", "Ancianos", "Visitas", "TOTAL"];

function addAttendanceRow(rec) {
  if (!rec || !rec.date) return;
  const monthName = getMonthSheetName(rec.date);
  const row = [rec.date, rec.children, rec.youth, rec.adults, rec.elderly, rec.visits, rec.total];

  // 1. Add to monthly sheet
  const monthSheet = getOrCreateSheet(monthName, ATT_HEADERS);
  // Check if date already exists → update instead of duplicate
  const data = monthSheet.getDataRange().getValues();
  let updated = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === rec.date) {
      monthSheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      updated = true;
      break;
    }
  }
  if (!updated) {
    monthSheet.appendRow(row);
    // Sort by date after appending
    if (monthSheet.getLastRow() > 2) {
      const dataRange = monthSheet.getRange(2, 1, monthSheet.getLastRow() - 1, row.length);
      dataRange.sort(1);
    }
  }
  monthSheet.autoResizeColumns(1, ATT_HEADERS.length);

  // 2. Also update/add to "Asistencia_Completa" sheet
  const fullSheet = getOrCreateSheet("Asistencia_Completa", ATT_HEADERS);
  const fullData = fullSheet.getDataRange().getValues();
  let fullUpdated = false;
  for (let i = 1; i < fullData.length; i++) {
    if (fullData[i][0] === rec.date) {
      fullSheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      fullUpdated = true;
      break;
    }
  }
  if (!fullUpdated) {
    fullSheet.appendRow(row);
    if (fullSheet.getLastRow() > 2) {
      fullSheet.getRange(2, 1, fullSheet.getLastRow() - 1, row.length).sort(1);
    }
  }
  fullSheet.autoResizeColumns(1, ATT_HEADERS.length);
}

// ── SYNC COMPLETO: ASISTENCIA ───────────────────────────────────────────────
function syncAttendance(data) {
  // Rebuilds all monthly sheets + Asistencia_Completa
  const byMonth = {};
  data.forEach(rec => {
    const m = getMonthSheetName(rec.date);
    if (!byMonth[m]) byMonth[m] = [];
    byMonth[m].push(rec);
  });

  Object.entries(byMonth).forEach(([month, recs]) => {
    const sheet = getOrCreateSheet(month, ATT_HEADERS);
    sheet.clearContents();
    sheet.appendRow(ATT_HEADERS);
    styleHeader(sheet, ATT_HEADERS.length, "#00338D");
    recs.sort((a,b) => a.date.localeCompare(b.date))
        .forEach(r => sheet.appendRow([r.date, r.children, r.youth, r.adults, r.elderly, r.visits, r.total]));
    sheet.autoResizeColumns(1, ATT_HEADERS.length);
  });

  // Full sheet
  const full = getOrCreateSheet("Asistencia_Completa", ATT_HEADERS);
  full.clearContents();
  full.appendRow(ATT_HEADERS);
  styleHeader(full, ATT_HEADERS.length, "#00338D");
  data.sort((a,b) => a.date.localeCompare(b.date))
      .forEach(r => full.appendRow([r.date, r.children, r.youth, r.adults, r.elderly, r.visits, r.total]));
  // Totals row
  if (data.length > 0) {
    const tot = ["TOTAL",
      data.reduce((s,r)=>s+r.children,0), data.reduce((s,r)=>s+r.youth,0),
      data.reduce((s,r)=>s+r.adults,0),  data.reduce((s,r)=>s+r.elderly,0),
      data.reduce((s,r)=>s+r.visits,0),  data.reduce((s,r)=>s+r.total,0)
    ];
    full.appendRow(tot);
    const totRange = full.getRange(full.getLastRow(), 1, 1, tot.length);
    totRange.setBackground("#F0AB00").setFontWeight("bold").setFontColor("#00338D");
  }
  full.autoResizeColumns(1, ATT_HEADERS.length);
}

// ── SYNC COMPLETO: MIEMBROS ──────────────────────────────────────────────────
// parents ahora es un arreglo de objetos: [{id, role}], role: "father"|"mother"|"guardian"
// allMembers se usa para resolver el ID del padre/madre a su nombre completo
function formatParents(parents, allMembers) {
  if (!parents || !parents.length) return "";
  const roleLabels = { father: "Padre", mother: "Madre", guardian: "Tutor/a" };
  return parents.map(function(p) {
    const info = allMembers.filter(function(m) { return m.id === p.id; })[0];
    const roleLabel = roleLabels[p.role] || p.role || "Padre";
    if (!info) return p.id + "(" + roleLabel + ")";
    return info.names + " " + info.lastNames + "(" + roleLabel + ")";
  }).join(" | ");
}

function syncMembers(data, allMembers) {
  const sheet = getOrCreateSheet("Miembros");
  sheet.clearContents();
  const headers = ["ID","Congregación","Estado","Nombres","Apellidos","F.Nacimiento","Lugar Nac.",
    "Dirección","GSM","Correo","Estado Civil","Cónyuge","F.Matrimonio",
    "F.Bautizo","Lugar Bautizo","Pastor Bautizo","Esp.Santo","F.Esp.Santo",
    "Ocupación","Escolaridad","Menor","Padres / Tutores"];
  sheet.appendRow(headers);
  styleHeader(sheet, headers.length, "#00338D");
  data.forEach(function(m) {
    sheet.appendRow([
      m.id, m.congregation, m.status, m.names, m.lastNames, m.birthDate, m.birthPlace,
      m.address, m.phone, m.email, m.marital, m.spouse, m.weddingDate,
      m.baptismDate, m.baptismPlace, m.pastor, m.spirit ? "Sí" : "No", m.spiritDate,
      m.occupation, m.education, m.isMinor ? "Sí" : "No",
      formatParents(m.parents, allMembers || data)
    ]);
  });
  sheet.autoResizeColumns(1, headers.length);
}

// ── SYNC COMPLETO: VISITAS / SIMPATIZANTES ──────────────────────────────────
function statusLabel(status) {
  const labels = {
    new: "Nueva visita",
    regular: "Visita regular",
    sympathizer: "Simpatizante",
    prebaptism: "Pre-bautismo"
  };
  return labels[status] || "Nueva visita";
}

function syncVisitors(data) {
  const sheet = getOrCreateSheet("Visitas");
  sheet.clearContents();
  const headers = ["ID","Nombres","Apellidos","F.Nacimiento","Dirección","GSM","Correo",
    "F.Registro","Estado","REFAM","N.º Visitas","Notas"];
  sheet.appendRow(headers);
  styleHeader(sheet, headers.length, "#009FDA");
  data.forEach(function(v) {
    sheet.appendRow([
      v.id, v.names, v.lastNames, v.birthDate, v.address, v.phone, v.email,
      v.regDate, statusLabel(v.status), v.refam ? "Sí" : "No", v.visitCount || 1, v.notes
    ]);
  });
  sheet.autoResizeColumns(1, headers.length);
}
