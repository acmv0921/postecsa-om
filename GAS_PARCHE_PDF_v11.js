// ═══════════════════════════════════════════════════════════════
// PARCHE GAS — PDFs de OM a Drive (v11)  ·  POSTECSA OM v8.0
// Carpeta destino: "INFORMES PDF OM" (dentro de OM POSTECSA)
// ═══════════════════════════════════════════════════════════════
// CÓMO APLICAR (3 pasos, en el proyecto "POSTECSA-Materiales"):
//  1) Pegue este bloque completo al final del script.
//  2) En doPost, dentro de   if(body.accion==='subir_pdf'){ ... }
//     reemplace TODO el contenido por:   return subirPdfOM_(body);
//  3) En doGet agregue (junto a los demás if de accion):
//       if(e.parameter.accion==='get_pdf_link') return getPdfLink_(e.parameter.otNum);
//     Luego: Implementar > Administrar implementaciones > Editar > Nueva versión.
//
// CORRIGE: la versión anterior escribía el link en la columna 13 de
// OMs-Pendientes, que es Ts_Cerrada (la sobrescribía). Ahora el link
// solo se guarda en Resumen-OMs, col 12 (Link_PDF).
// ═══════════════════════════════════════════════════════════════

var PDF_FOLDER_ID = '1ytXqUSe6ZOXktOfRDE61MRRGxKtUkf5-'; // INFORMES PDF OM

function _pdfFolderMes_(mes) {
  var root;
  try { root = DriveApp.getFolderById(PDF_FOLDER_ID); }
  catch (ex) { // respaldo: carpeta por nombre en el Drive del script
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

  // Si ya existe un PDF de esa OM en la carpeta del mes, se reemplaza (evita duplicados)
  var previos = folder.getFilesByName(nombre);
  while (previos.hasNext()) previos.next().setTrashed(true);

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var link = file.getUrl();

  var hr = _resH(), rr = hr.getDataRange().getValues();
  for (var j = 1; j < rr.length; j++) {
    if (String(rr[j][0]) === otNum) { hr.getRange(j + 1, 12).setValue(link); break; }
  }
  return _ok({guardado: true, link: link});
}

function getPdfLink_(otNum) {
  var hr = _resH(), rr = hr.getDataRange().getValues();
  for (var j = 1; j < rr.length; j++) {
    if (String(rr[j][0]) === String(otNum)) return _ok({link: String(rr[j][11] || '')});
  }
  return _ok({link: ''});
}
