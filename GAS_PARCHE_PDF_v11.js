// ═══════════════════════════════════════════════════════════════
// PARCHE GAS v11.1 — PDFs de OM a Drive  (APLICADO y publicado el 15/11/2026,
// versión 31 de la implementación de POSTECSA-Materiales)
// Carpeta destino: "INFORMES PDF OM" (dentro de OM POSTECSA)
// Cambios en el script publicado (Código.gs):
//   · doPost: if (body.accion === 'subir_pdf') { return subirPdfOM_(body); }
//   · doGet : if (accion === 'get_pdf_link')   { return getPdfLink_(p.otNum); }
//   · funciones de abajo al final del archivo.
// Los links quedan en la hoja nueva "PDFs-OM" (N_OM, Fecha, Archivo, Link_PDF);
// no se toca OMs-Pendientes (la col. 13 es Ts_Cerrada).
// ═══════════════════════════════════════════════════════════════

var PDF_FOLDER_ID = '1ytXqUSe6ZOXktOfRDE61MRRGxKtUkf5-'; // INFORMES PDF OM
var HOJA_PDFS = 'PDFs-OM';

function _pdfsH_() {
  var ss = _ss(), h = ss.getSheetByName(HOJA_PDFS);
  if (!h) {
    h = ss.insertSheet(HOJA_PDFS);
    h.appendRow(['N_OM','Fecha','Archivo','Link_PDF']);
    h.getRange(1,1,1,4).setBackground('#E65100').setFontColor('#fff').setFontWeight('bold');
  }
  return h;
}

function _pdfFolderMes_(mes) {
  var root;
  try { root = DriveApp.getFolderById(PDF_FOLDER_ID); }
  catch (ex) {
    var it = DriveApp.getFoldersByName('INFORMES PDF OM');
    root = it.hasNext() ? it.next() : DriveApp.createFolder('INFORMES PDF OM');
  }
  var sub = root.getFoldersByName(mes);
  return sub.hasNext() ? sub.next() : root.createFolder(mes);
}

function subirPdfOM_(body) {
  var otNum  = String(body.otNum || '');
  var nombre = body.nombre || ('OM_' + otNum + '_POSTECSA.pdf');
  if (!body.pdf) return _err('PDF vacio');
  var blob = Utilities.newBlob(Utilities.base64Decode(body.pdf), 'application/pdf', nombre);
  var mes  = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM');
  var folder = _pdfFolderMes_(mes);
  var previos = folder.getFilesByName(nombre);
  while (previos.hasNext()) previos.next().setTrashed(true);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var link = file.getUrl();
  var h = _pdfsH_(), rows = h.getDataRange().getValues(), hecho = false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === otNum) {
      h.getRange(i+1,2,1,3).setValues([[new Date(), nombre, link]]);
      hecho = true; break;
    }
  }
  if (!hecho) h.appendRow([otNum, new Date(), nombre, link]);
  return _ok({guardado: true, link: link});
}

function getPdfLink_(otNum) {
  var rows = _pdfsH_().getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(otNum)) return _ok({link: String(rows[i][3] || '')});
  }
  return _ok({link: ''});
}
