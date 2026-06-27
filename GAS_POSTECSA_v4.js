// ═══════════════════════════════════════════════════
// GAS POSTECSA v5 — Con CORS via HtmlService
// ═══════════════════════════════════════════════════

var SHEET_ID = '1IlxWD6QiCsntzyjEnJfE5KlB36Q39wJsQLnR0HE9gM4';
var HOJA_MAT = 'Hoja 1';
var HOJA_OMS = 'OMs-Pendientes';

function _ss()   { return SpreadsheetApp.openById(SHEET_ID); }
function _matH() { return _ss().getSheetByName(HOJA_MAT); }
function _omsH() {
  var ss=_ss(), h=ss.getSheetByName(HOJA_OMS);
  if(!h){
    h=ss.insertSheet(HOJA_OMS);
    h.appendRow(['num','fecha','area','equipo','falla','tipo_mant','prioridad','mec_cc','mec_nom','obs_i','estado','pdf_link']);
  }
  return h;
}

// Respuesta JSON con headers CORS explícitos via HtmlService
function _resp(data) {
  var json = JSON.stringify(data);
  var output = ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function _ok(data)  { return _resp(Object.assign({ok:true},  data||{})); }
function _err(msg)  { return _resp({ok:false, error:msg}); }

// ══════════════════════════════════════════════════
// doGet — lectura (CORS automático en GAS publicado)
// ══════════════════════════════════════════════════
function doGet(e) {
  var p      = (e && e.parameter) ? e.parameter : {};
  var accion = p.accion || '';

  try {
    if (!accion || accion === 'materiales') {
      var h    = _matH();
      var rows = h ? h.getDataRange().getValues() : [];
      var mats = [];
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[0] && !r[1]) continue;
        mats.push({codigo:String(r[0]||''), nombre:String(r[1]||''),
                   unidad:String(r[2]||'Und'), precio:Number(r[3])||0,
                   stock:Number(r[4])||0,
                   otNum:String(r[5]||''), equipo:String(r[6]||''),
                   mecanico:String(r[7]||''), fecha:String(r[8]||''),
                   cant:Number(r[9])||0,
                   desc:String(r[1]||''), um:String(r[2]||'Und'),
                   n:String(r[1]||''), u:String(r[2]||'Und'), p:Number(r[3])||0});
      }
      return _ok({materiales: mats});
    }

    if (accion === 'mis_oms') {
      var cc   = String(p.cc || '');
      var rows = _omsH().getDataRange().getValues();
      var oms  = [];
      for (var j = 1; j < rows.length; j++) {
        var row    = rows[j];
        var estado = String(row[10] || 'PENDIENTE');
        if (String(row[7]) === cc && estado !== 'BORRADA') {
          oms.push({num:String(row[0]),fecha:String(row[1]),area:String(row[2]),
                    equipo:String(row[3]),falla:String(row[4]),tipo_mant:String(row[5]),
                    prioridad:String(row[6]),mec_cc:String(row[7]),mec_nom:String(row[8]),
                    obs_i:String(row[9]),estado:estado});
        }
      }
      return _ok({oms: oms});
    }

    if (accion === 'todas_oms') {
      var rows = _omsH().getDataRange().getValues();
      var oms  = [];
      for (var k = 1; k < rows.length; k++) {
        var row3 = rows[k];
        if (!row3[0]) continue;
        oms.push({num:String(row3[0]),fecha:String(row3[1]),area:String(row3[2]),
                  equipo:String(row3[3]),falla:String(row3[4]),tipo_mant:String(row3[5]),
                  prioridad:String(row3[6]),mec_cc:String(row3[7]),mec_nom:String(row3[8]),
                  obs_i:String(row3[9]),estado:String(row3[10]||'PENDIENTE')});
      }
      return _ok({oms: oms});
    }

    if (accion === 'borrar_om') {
      var num  = String(p.num || '');
      var h4   = _omsH();
      var rows = h4.getDataRange().getValues();
      for (var m = 1; m < rows.length; m++) {
        if (String(rows[m][0]) === num) {
          h4.getRange(m+1, 11).setValue('BORRADA');
          return _ok({borrada: num});
        }
      }
      return _err('OM no encontrada: ' + num);
    }

    if (accion === 'get_pdf_link') {
      var otNum = String(p.otNum || '');
      var rows  = _omsH().getDataRange().getValues();
      for (var n = 1; n < rows.length; n++) {
        if (String(rows[n][0]) === otNum && rows[n][11]) {
          return _ok({link: String(rows[n][11])});
        }
      }
      return _err('Link no encontrado');
    }

    if (accion === 'guardar_material_get') {
      var hm = _matH();
      hm.appendRow([
        p.cod||'', p.desc||'', p.um||'Und',
        Number(p.precio)||0, 0,
        p.otNum||'', p.equipo||'', p.mecanico||'',
        p.fecha||'', Number(p.cant)||1
      ]);
      return _ok({guardado: true});
    }

    return _err('Accion no reconocida: ' + accion);
  } catch(ex) {
    return _err('Error interno: ' + ex.toString());
  }
}

// ══════════════════════════════════════════════════
// doPost — escritura
// ══════════════════════════════════════════════════
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch(ex) {}
  var accion = body.accion || '';

  try {
    if (accion === 'crear_om') {
      var h    = _omsH();
      var num  = String(body.num || '');
      var rows = h.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === num) return _ok({ya_existe: true});
      }
      h.appendRow([num, body.fecha||'', body.area||'', body.equipo||'',
                   body.falla||'', body.tipo_mant||'', body.prioridad||'MEDIA',
                   String(body.mec_cc||''), body.mec_nom||'', body.obs_i||'',
                   'PENDIENTE', '']);
      return _ok({creada: num});
    }

    if (accion === 'tomar_om') {
      var h2   = _omsH();
      var rows = h2.getDataRange().getValues();
      for (var j = 1; j < rows.length; j++) {
        if (String(rows[j][0]) === String(body.num||'')) {
          h2.getRange(j+1, 11).setValue('TOMADA');
          return _ok({tomada: body.num});
        }
      }
      return _err('OM no encontrada');
    }

    if (accion === 'actualizar_estado_om') {
      var h3   = _omsH();
      var rows = h3.getDataRange().getValues();
      for (var k = 1; k < rows.length; k++) {
        if (String(rows[k][0]) === String(body.num||'')) {
          h3.getRange(k+1, 11).setValue(body.estado || 'PENDIENTE');
          return _ok({actualizado: body.num});
        }
      }
      return _err('OM no encontrada');
    }

    if (accion === 'borrar_om') {
      var h4   = _omsH();
      var rows = h4.getDataRange().getValues();
      for (var m = 1; m < rows.length; m++) {
        if (String(rows[m][0]) === String(body.num||'')) {
          h4.getRange(m+1, 11).setValue('BORRADA');
          return _ok({borrada: body.num});
        }
      }
      return _err('OM no encontrada');
    }

    if (accion === 'guardar_material') {
      var hm = _matH();
      hm.appendRow([body.cod||'', body.desc||'', body.um||'Und',
                    Number(body.precio)||0, 0, body.otNum||'',
                    body.equipo||'', body.mecanico||'', body.fecha||'',
                    Number(body.cant)||1]);
      return _ok({guardado: true});
    }

    if (accion === 'subir_pdf') {
      var h6   = _omsH();
      var rows = h6.getDataRange().getValues();
      for (var p2 = 1; p2 < rows.length; p2++) {
        if (String(rows[p2][0]) === String(body.otNum||'')) {
          h6.getRange(p2+1, 12).setValue(body.link || '');
          return _ok({guardado: true});
        }
      }
      return _err('OM no encontrada para PDF');
    }

    if (accion === 'limpiar_todo') {
      var hc  = _matH();
      var last = hc.getLastRow();
      if (last > 1) hc.deleteRows(2, last-1);
      return _ok({limpiado: true});
    }

    if (accion === 'borrar_material') {
      var hb   = _matH();
      var rows = hb.getDataRange().getValues();
      for (var bi = 1; bi < rows.length; bi++) {
        if (String(rows[bi][1]).toLowerCase() === String(body.desc||'').toLowerCase()) {
          hb.deleteRow(bi+1);
          return _ok({borrado: body.desc});
        }
      }
      return _err('Material no encontrado');
    }

    return _err('Accion POST no reconocida: ' + accion);
  } catch(ex) {
    return _err('Error interno: ' + ex.toString());
  }
}
