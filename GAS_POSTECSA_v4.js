// ═══════════════════════════════════════════════════
// GAS POSTECSA v4 — Google Apps Script
// Sheet ID: 1IlxWD6QiCsntzyjEnJfE5KlB36Q39wJsQLnR0HE9gM4
// Hojas: "Hoja 1" (materiales) | "OMs-Pendientes" (supervisor)
// ═══════════════════════════════════════════════════

var SHEET_ID = '1IlxWD6QiCsntzyjEnJfE5KlB36Q39wJsQLnR0HE9gM4';
var HOJA_MAT = 'Hoja 1';
var HOJA_OMS = 'OMs-Pendientes';

function _ss()  { return SpreadsheetApp.openById(SHEET_ID); }
function _matH(){ return _ss().getSheetByName(HOJA_MAT); }
function _omsH(){
  var ss=_ss(), h=ss.getSheetByName(HOJA_OMS);
  if(!h){ h=ss.insertSheet(HOJA_OMS); h.appendRow(['num','fecha','area','equipo','falla','tipo_mant','prioridad','mec_cc','mec_nom','obs_i','estado']); }
  return h;
}

function _cors(output){
  return output
    .setMimeType(ContentService.MimeType.JSON);
}

function _ok(data){
  return _cors(ContentService.createTextOutput(JSON.stringify(Object.assign({ok:true},data||{}))));
}

function _err(msg){
  return _cors(ContentService.createTextOutput(JSON.stringify({ok:false,error:msg})));
}

// ══════════════════════════════════════════════════
// doGet
// ══════════════════════════════════════════════════
function doGet(e){
  var p=e.parameter||{};
  var accion=p.accion||'';
  var cb=p.callback||'';

  function _wrap(output){
    if(cb){
      return ContentService.createTextOutput(cb+'('+output.getContent()+');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return output;
  }

  // Materiales (default)
  if(!accion || accion==='materiales'){
    var h=_matH();
    if(!h) return _wrap(_err('Hoja materiales no encontrada'));
    var data=h.getDataRange().getValues();
    var mats=[];
    for(var i=1;i<data.length;i++){
      var r=data[i];
      if(!r[0]&&!r[1]) continue;
      mats.push({codigo:r[0]||'',nombre:r[1]||'',unidad:r[2]||'Und',precio:Number(r[3])||0,stock:Number(r[4])||0});
    }
    return _wrap(_ok({materiales:mats}));
  }

  // OMs asignadas a un mecánico
  if(accion==='mis_oms'){
    var cc=p.cc||'';
    var h2=_omsH();
    var data2=h2.getDataRange().getValues();
    var oms=[];
    for(var j=1;j<data2.length;j++){
      var row=data2[j];
      var estado=String(row[10]||'PENDIENTE');
      if(String(row[7])===String(cc) && estado!=='BORRADA'){
        oms.push({num:row[0],fecha:row[1],area:row[2],equipo:row[3],falla:row[4],tipo_mant:row[5],prioridad:row[6],mec_cc:row[7],mec_nom:row[8],obs_i:row[9],estado:estado});
      }
    }
    return _wrap(_ok({oms:oms}));
  }

  // Todas las OMs (supervisor)
  if(accion==='todas_oms'){
    var h3=_omsH();
    var data3=h3.getDataRange().getValues();
    var oms2=[];
    for(var k=1;k<data3.length;k++){
      var row3=data3[k];
      if(!row3[0]) continue;
      oms2.push({num:row3[0],fecha:row3[1],area:row3[2],equipo:row3[3],falla:row3[4],tipo_mant:row3[5],prioridad:row3[6],mec_cc:row3[7],mec_nom:row3[8],obs_i:row3[9],estado:String(row3[10]||'PENDIENTE')});
    }
    return _wrap(_ok({oms:oms2}));
  }

  // Borrar OM por GET (para evitar CORS en POST)
  if(accion==='borrar_om'){
    var num=p.num||'';
    var h4=_omsH();
    var data4=h4.getDataRange().getValues();
    for(var m=1;m<data4.length;m++){
      if(String(data4[m][0])===String(num)){
        h4.getRange(m+1,11).setValue('BORRADA');
        return _wrap(_ok({borrada:num}));
      }
    }
    return _wrap(_err('OM no encontrada: '+num));
  }

  // Link PDF en Drive
  if(accion==='get_pdf_link'){
    var otNum=p.otNum||'';
    var h5=_omsH();
    var data5=h5.getDataRange().getValues();
    for(var n=1;n<data5.length;n++){
      if(String(data5[n][0])===String(otNum) && data5[n][11]){
        return _wrap(_ok({link:data5[n][11]}));
      }
    }
    return _wrap(_err('Link no encontrado'));
  }

  return _wrap(_err('Accion GET no reconocida: '+accion));
}

// ══════════════════════════════════════════════════
// doPost
// ══════════════════════════════════════════════════
function doPost(e){
  var body={};
  try{ body=JSON.parse(e.postData.contents||'{}'); }catch(ex){}
  var accion=body.accion||'';

  if(accion==='crear_om'){
    var h=_omsH();
    var num=String(body.num||'');
    // Verificar que no existe ya
    var data=h.getDataRange().getValues();
    for(var i=1;i<data.length;i++){
      if(String(data[i][0])===num) return _ok({ya_existe:true});
    }
    h.appendRow([num,body.fecha||'',body.area||'',body.equipo||'',body.falla||'',
                 body.tipo_mant||'',body.prioridad||'MEDIA',
                 String(body.mec_cc||''),body.mec_nom||'',body.obs_i||'','PENDIENTE','']);
    return _ok({creada:num});
  }

  if(accion==='tomar_om'){
    var h2=_omsH();
    var data2=h2.getDataRange().getValues();
    for(var j=1;j<data2.length;j++){
      if(String(data2[j][0])===String(body.num||'')){
        h2.getRange(j+1,11).setValue('TOMADA');
        return _ok({tomada:body.num});
      }
    }
    return _err('OM no encontrada');
  }

  if(accion==='actualizar_estado_om'){
    var h3=_omsH();
    var data3=h3.getDataRange().getValues();
    for(var k=1;k<data3.length;k++){
      if(String(data3[k][0])===String(body.num||'')){
        h3.getRange(k+1,11).setValue(body.estado||'PENDIENTE');
        return _ok({actualizado:body.num});
      }
    }
    return _err('OM no encontrada');
  }

  if(accion==='borrar_om'){
    var h4=_omsH();
    var data4=h4.getDataRange().getValues();
    for(var m=1;m<data4.length;m++){
      if(String(data4[m][0])===String(body.num||'')){
        h4.getRange(m+1,11).setValue('BORRADA');
        return _ok({borrada:body.num});
      }
    }
    return _err('OM no encontrada');
  }

  if(accion==='guardar_material'){
    var hm=_matH();
    hm.appendRow([body.cod||'',body.desc||'',body.um||'Und',Number(body.precio)||0,0,
                  body.otNum||'',body.equipo||'',body.mecanico||'',body.fecha||'',Number(body.cant)||1]);
    return _ok({guardado:true});
  }

  if(accion==='subir_pdf'){
    var h6=_omsH();
    var data6=h6.getDataRange().getValues();
    for(var p2=1;p2<data6.length;p2++){
      if(String(data6[p2][0])===String(body.otNum||'')){
        h6.getRange(p2+1,12).setValue(body.link||'');
        return _ok({guardado:true});
      }
    }
    return _err('OM no encontrada para PDF');
  }

  if(accion==='limpiar_todo'){
    var hc=_matH();
    var lastRow=hc.getLastRow();
    if(lastRow>1) hc.deleteRows(2,lastRow-1);
    return _ok({limpiado:true});
  }

  if(accion==='borrar_material'){
    var hb=_matH();
    var db=hb.getDataRange().getValues();
    for(var bi=1;bi<db.length;bi++){
      if(String(db[bi][1]).toLowerCase()===String(body.desc||'').toLowerCase()){
        hb.deleteRow(bi+1);
        return _ok({borrado:body.desc});
      }
    }
    return _err('Material no encontrado');
  }

  return _err('Accion POST no reconocida: '+accion);
}
