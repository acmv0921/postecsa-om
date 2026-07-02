// ═══════════════════════════════════════════════════
// GAS POSTECSA v10 — Resumen + KPIs + Drive PDF
// Sheet: 1IlxWD6QiCsntzyjEnJfE5KlB36Q39wJsQLnR0HE9gM4
// ═══════════════════════════════════════════════════

var SHEET_ID    = '1IlxWD6QiCsntzyjEnJfE5KlB36Q39wJsQLnR0HE9gM4';
var HOJA_MAT    = 'Hoja 1';
var HOJA_OMS    = 'OMs-Pendientes';
var HOJA_RES    = 'Resumen-OMs';
var HOJA_KPI    = 'KPIs';
var FOLDER_NAME = 'POSTECSA-OM-PDFs';

var ESTADO_PRIO = {CERRADA:4, TOMADA:3, PENDIENTE:2, BORRADA:1};
var MIN_RATE    = 132.63; // $/min mano de obra

function _ss()    { return SpreadsheetApp.openById(SHEET_ID); }
function _matH()  { return _ss().getSheetByName(HOJA_MAT); }
function _omsH()  {
  var ss=_ss(), h=ss.getSheetByName(HOJA_OMS);
  if(!h){
    h=ss.insertSheet(HOJA_OMS);
    h.appendRow(['N_OM','Fecha','Area','Equipo','Falla','Prioridad',
                 'Mecanico_CC','Mecanico_Nom','Observacion','Estado',
                 'Ts_Creacion','Ts_Tomada','Ts_Cerrada']);
    h.getRange(1,1,1,13).setBackground('#E65100').setFontColor('#fff').setFontWeight('bold');
  }
  return h;
}
function _resH()  {
  var ss=_ss(), h=ss.getSheetByName(HOJA_RES);
  if(!h){
    h=ss.insertSheet(HOJA_RES);
    h.appendRow(['N_OM','Fecha_Cierre','Area','Equipo','Falla','Mecanico',
                 'Minutos_Labor','Costo_MO','Costo_Repuestos','Costo_Total',
                 'Num_Repuestos','Link_PDF','Observacion']);
    h.getRange(1,1,1,13).setBackground('#1565C0').setFontColor('#fff').setFontWeight('bold');
  }
  return h;
}
function _kpiH()  {
  var ss=_ss(), h=ss.getSheetByName(HOJA_KPI);
  if(!h){
    h=ss.insertSheet(HOJA_KPI);
    h.appendRow(['Tipo','Nombre','Total_OMs','Total_Minutos','Costo_Total',
                 'Costo_Promedio','Tiempo_Promedio_Min','Ultima_Actualizacion']);
    h.getRange(1,1,1,8).setBackground('#263238').setFontColor('#fff').setFontWeight('bold');
  }
  return h;
}

function _resp(d) {
  return ContentService.createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);
}
function _ok(d)  { return _resp(Object.assign({ok:true}, d||{})); }
function _err(m) { return _resp({ok:false, error:m}); }

// ── Obtener/crear carpeta Drive del mes ────────────
function _getFolder(mes) {
  var root = null;
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  root = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  var subName = mes || Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM');
  var subs = root.getFoldersByName(subName);
  return subs.hasNext() ? subs.next() : root.createFolder(subName);
}

// ── Deduplicar OMs ────────────────────────────────
function _deduplicar() {
  var h=_omsH(), rows=h.getDataRange().getValues();
  if(rows.length<=1) return 0;
  var mapa={};
  for(var i=1;i<rows.length;i++){
    var num=String(rows[i][0]||'').trim();
    var estado=String(rows[i][9]||'PENDIENTE').trim();
    if(!num) continue;
    var prio=ESTADO_PRIO[estado]||0;
    if(!mapa[num]||prio>mapa[num].prio) mapa[num]={idx:i,prio:prio};
  }
  var ganadoras={};
  for(var n in mapa) ganadoras[mapa[n].idx]=true;
  var borradas=0;
  for(var j=rows.length-1;j>=1;j--){
    var nm=String(rows[j][0]||'').trim();
    if(nm&&!ganadoras[j]){h.deleteRow(j+1);borradas++;}
  }
  return borradas;
}

// ── Registrar OM cerrada en Resumen-OMs ───────────
function _registrarResumen(p) {
  try {
    var hr = _resH();
    // Verificar si ya existe
    var rows = hr.getDataRange().getValues();
    for(var i=1;i<rows.length;i++){
      if(String(rows[i][0])===String(p.num||'')) return; // Ya registrada
    }
    var costoMO   = Number(p.minutos||0) * MIN_RATE;
    var costoRep  = Number(p.costo_repuestos||0);
    var costoTot  = costoMO + costoRep;
    hr.appendRow([
      p.num||'',
      p.fecha_cierre||new Date().toLocaleDateString('es-CO'),
      p.area||'',
      p.equipo||'',
      p.falla||'',
      p.mecanico||'',
      Number(p.minutos)||0,
      Math.round(costoMO),
      Math.round(costoRep),
      Math.round(costoTot),
      Number(p.num_repuestos)||0,
      p.link_pdf||'',
      p.obs_f||''
    ]);
    // Actualizar KPIs
    _actualizarKPIs();
  } catch(e) {
    Logger.log('Error en _registrarResumen: '+e);
  }
}

// ── Actualizar KPIs desde Resumen-OMs ─────────────
function _actualizarKPIs() {
  try {
    var hr   = _resH();
    var rows = hr.getDataRange().getValues();
    if(rows.length<=1) return;
    var datos=rows.slice(1).filter(function(r){return r[0];});
    
    // Agrupar por equipo y por mecánico
    var porEquipo={}, porMecanico={};
    datos.forEach(function(r){
      var eq=String(r[3]||'').trim();
      var mc=String(r[5]||'').trim();
      var min=Number(r[6])||0;
      var cost=Number(r[9])||0;
      if(eq){
        if(!porEquipo[eq]) porEquipo[eq]={oms:0,min:0,cost:0};
        porEquipo[eq].oms++; porEquipo[eq].min+=min; porEquipo[eq].cost+=cost;
      }
      if(mc){
        if(!porMecanico[mc]) porMecanico[mc]={oms:0,min:0,cost:0};
        porMecanico[mc].oms++; porMecanico[mc].min+=min; porMecanico[mc].cost+=cost;
      }
    });
    
    var hk=_kpiH();
    // Limpiar datos anteriores (mantener header)
    var lastRow=hk.getLastRow();
    if(lastRow>1) hk.deleteRows(2,lastRow-1);
    
    var ahora=new Date().toLocaleString('es-CO');
    // Escribir por equipo
    for(var eq in porEquipo){
      var d=porEquipo[eq];
      hk.appendRow(['EQUIPO',eq,d.oms,d.min,Math.round(d.cost),
        Math.round(d.cost/d.oms),Math.round(d.min/d.oms),ahora]);
    }
    // Escribir por mecánico
    for(var mc in porMecanico){
      var dm=porMecanico[mc];
      hk.appendRow(['MECANICO',mc,dm.oms,dm.min,Math.round(dm.cost),
        Math.round(dm.cost/dm.oms),Math.round(dm.min/dm.oms),ahora]);
    }
  } catch(e) {
    Logger.log('Error en _actualizarKPIs: '+e);
  }
}

// ══════════════════════════════════════════════════
// doGet
// ══════════════════════════════════════════════════
function doGet(e) {
  var p=e&&e.parameter?e.parameter:{};
  var accion=p.accion||'';
  try {

    if(accion==='crear_om_get'){
      var h=_omsH(), num=String(p.num||'');
      if(!num) return _err('Numero requerido');
      var rows=h.getDataRange().getValues();
      for(var i=1;i<rows.length;i++){
        if(String(rows[i][0])===num&&String(rows[i][9])!=='BORRADA')
          return _ok({ya_existe:true});
      }
      h.appendRow([num,p.fecha||'',p.area||'',p.equipo||'',p.falla||'',
                   p.prioridad||'MEDIA',String(p.mec_cc||''),p.mec_nom||'',
                   p.obs_i||'','PENDIENTE',new Date().toISOString(),'','']);
      _deduplicar();
      return _ok({creada:num});
    }

    if(accion==='tomar_om_get'){
      var h2=_omsH(), rows=h2.getDataRange().getValues();
      for(var j=1;j<rows.length;j++){
        if(String(rows[j][0])===String(p.num||'')&&
           String(rows[j][9])!=='BORRADA'&&String(rows[j][9])!=='CERRADA'){
          h2.getRange(j+1,10).setValue('TOMADA');
          h2.getRange(j+1,12).setValue(new Date().toISOString());
          return _ok({tomada:p.num});
        }
      }
      return _err('OM no encontrada');
    }

    if(accion==='actualizar_estado_get'){
      var h3=_omsH(), rows=h3.getDataRange().getValues();
      for(var k=1;k<rows.length;k++){
        if(String(rows[k][0])===String(p.num||'')&&String(rows[k][9])!=='BORRADA'){
          h3.getRange(k+1,10).setValue(p.estado||'PENDIENTE');
          if(p.estado==='CERRADA'){
            h3.getRange(k+1,13).setValue(new Date().toISOString());
            // Registrar en Resumen-OMs
            _registrarResumen({
              num:p.num, fecha_cierre:new Date().toLocaleDateString('es-CO'),
              area:String(rows[k][2]||''), equipo:String(rows[k][3]||''),
              falla:String(rows[k][4]||''), mecanico:String(rows[k][7]||''),
              minutos:Number(p.minutos)||0,
              costo_repuestos:Number(p.costo_repuestos)||0,
              num_repuestos:Number(p.num_repuestos)||0,
              obs_f:p.obs_f||''
            });
          }
          return _ok({actualizado:p.num,estado:p.estado});
        }
      }
      return _err('OM no encontrada: '+p.num);
    }

    if(accion==='borrar_om'){
      var h4=_omsH(), rows=h4.getDataRange().getValues();
      for(var m=1;m<rows.length;m++){
        if(String(rows[m][0])===String(p.num||'')){
          h4.getRange(m+1,10).setValue('BORRADA');
          return _ok({borrada:p.num});
        }
      }
      return _err('OM no encontrada: '+p.num);
    }

    if(accion==='guardar_material_get'){
      var hm=_matH();
      if(!hm){
        hm=_ss().insertSheet(HOJA_MAT);
        hm.appendRow(['Fecha','OT','Equipo','Mecanico','Material','Cantidad','UM','Precio']);
      }
      hm.appendRow([
        p.fecha||new Date().toLocaleDateString('es-CO'),
        p.otNum||'', p.equipo||'', p.mecanico||'', p.desc||'',
        Number(p.cant)||1, p.um||'Und', Number(p.precio)||0
      ]);
      return _ok({guardado:true});
    }

    if(accion==='deduplicar'){
      return _ok({borradas:_deduplicar()});
    }

    if(accion==='todas_oms'){
      var rows=_omsH().getDataRange().getValues(), oms=[];
      for(var a=1;a<rows.length;a++){
        var r=rows[a];
        if(!r[0]) continue;
        oms.push({num:r[0],fecha:r[1],area:r[2],equipo:r[3],
                  falla:r[4],prioridad:r[5],mec_cc:String(r[6]),
                  mec_nom:r[7],obs_i:r[8],estado:String(r[9]||'PENDIENTE'),
                  tipo_mant:'CORRECTIVO'});
      }
      return _ok({oms:oms});
    }

    if(accion==='mis_oms'){
      var cc=String(p.cc||'');
      var rows=_omsH().getDataRange().getValues(), oms=[];
      for(var b=1;b<rows.length;b++){
        var r=rows[b], estado=String(r[9]||'PENDIENTE');
        if(String(r[6])===cc&&estado!=='BORRADA')
          oms.push({num:r[0],fecha:r[1],area:r[2],equipo:r[3],
                    falla:r[4],prioridad:r[5],mec_cc:String(r[6]),
                    mec_nom:r[7],obs_i:r[8],estado:estado,tipo_mant:'CORRECTIVO'});
      }
      return _ok({oms:oms});
    }

    if(accion==='get_kpis'){
      var rows=_kpiH().getDataRange().getValues(), kpis=[];
      for(var c=1;c<rows.length;c++){
        var r=rows[c];
        if(!r[0]) continue;
        kpis.push({tipo:r[0],nombre:r[1],total_oms:r[2],total_min:r[3],
                   costo_total:r[4],costo_prom:r[5],tiempo_prom:r[6]});
      }
      return _ok({kpis:kpis});
    }

    if(accion==='get_resumen'){
      var rows=_resH().getDataRange().getValues(), res=[];
      for(var d=1;d<rows.length;d++){
        var r=rows[d];
        if(!r[0]) continue;
        res.push({num:r[0],fecha:r[1],area:r[2],equipo:r[3],falla:r[4],
                  mecanico:r[5],minutos:r[6],costo_mo:r[7],costo_rep:r[8],
                  costo_total:r[9],num_rep:r[10],link_pdf:r[11],obs:r[12]});
      }
      return _ok({resumen:res});
    }

    return _err('Accion no reconocida: '+accion);

  } catch(ex) {
    return _err('Error: '+ex.toString());
  }
}

// ══════════════════════════════════════════════════
// doPost — subir PDF a Drive + guardar link en Sheet
// ══════════════════════════════════════════════════
function doPost(e) {
  var body={};
  try{body=JSON.parse(e.postData.contents||'{}');}catch(ex){}
  try {
    if(body.accion==='subir_pdf'){
      var otNum  = String(body.otNum||'');
      var nombre = body.nombre||('OM-'+otNum+'.pdf');
      var pdfB64 = body.pdf||'';
      if(!pdfB64) return _err('PDF vacio');

      // Decodificar y subir a Drive
      var bytes  = Utilities.base64Decode(pdfB64);
      var blob   = Utilities.newBlob(bytes,'application/pdf',nombre);
      var mes    = Utilities.formatDate(new Date(),'America/Bogota','yyyy-MM');
      var folder = _getFolder(mes);
      var file   = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var link   = file.getUrl();

      // Guardar link en OMs-Pendientes
      var h=_omsH(), rows=h.getDataRange().getValues();
      for(var i=1;i<rows.length;i++){
        if(String(rows[i][0])===otNum){
          h.getRange(i+1,13).setValue(link); // Col M = Ts_Cerrada → usar col N para link
          break;
        }
      }

      // Actualizar link en Resumen-OMs
      var hr=_resH(), rrows=hr.getDataRange().getValues();
      for(var j=1;j<rrows.length;j++){
        if(String(rrows[j][0])===otNum){
          hr.getRange(j+1,12).setValue(link);
          break;
        }
      }
      return _ok({guardado:true, link:link});
    }

    if(body.accion==='registrar_cierre'){
      _registrarResumen(body);
      return _ok({registrado:true});
    }

    if(body.accion==='limpiar_todo'){
      var h=_matH();
      if(h){var l=h.getLastRow();if(l>1)h.deleteRows(2,l-1);}
      return _ok({limpiado:true});
    }

    if(body.accion==='recalcular_kpis'){
      _actualizarKPIs();
      return _ok({actualizado:true});
    }

    return _err('Accion POST no reconocida: '+body.accion);
  } catch(ex) {
    return _err('Error: '+ex.toString());
  }
}
