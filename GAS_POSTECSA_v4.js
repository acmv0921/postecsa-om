// ═══════════════════════════════════════════════════
// GAS POSTECSA v9 — Con deduplicación automática
// Hoja 1: Fecha|OT|Equipo|Mecanico|Material|Cantidad|UM|Precio
// OMs-Pendientes: N_OM|Fecha|Area|Equipo|Falla|Prioridad|
//                 Mecanico_CC|Mecanico_Nom|Observacion|Estado|
//                 Ts_Creacion|Ts_Tomada|Ts_Cerrada
// ═══════════════════════════════════════════════════

var SHEET_ID = '1IlxWD6QiCsntzyjEnJfE5KlB36Q39wJsQLnR0HE9gM4';
var HOJA_MAT = 'Hoja 1';
var HOJA_OMS = 'OMs-Pendientes';

// Prioridad de estados (mayor = más importante)
var ESTADO_PRIO = {CERRADA:4, TOMADA:3, PENDIENTE:2, BORRADA:1};

function _ss()   { return SpreadsheetApp.openById(SHEET_ID); }
function _matH() { return _ss().getSheetByName(HOJA_MAT); }
function _omsH() {
  var ss = _ss(), h = ss.getSheetByName(HOJA_OMS);
  if (!h) {
    h = ss.insertSheet(HOJA_OMS);
    h.appendRow(['N_OM','Fecha','Area','Equipo','Falla','Prioridad',
                 'Mecanico_CC','Mecanico_Nom','Observacion','Estado',
                 'Ts_Creacion','Ts_Tomada','Ts_Cerrada']);
  }
  return h;
}


// ── EXTERNOS ──────────────────────────────────────────────────
var HOJA_EXT = 'Externos';

function _extH(){
  var ss = _ss(), h = ss.getSheetByName(HOJA_EXT);
  if (!h) {
    h = ss.insertSheet(HOJA_EXT);
    h.appendRow(['ID','Nombre','Especialidad','Activo','Fecha_Registro']);
    // Pre-cargar los 6 base
    var base = [
      ['EXT001','MONTACARGAS VARELA','VEHÍCULOS (CATERPILLAR, YALE)',true],
      ['EXT002','REFRIGERACIÓN Y ELÉCTRICOS SNB SAS','AIRES ACONDICIONADOS Y CHILLER',true],
      ['EXT003','TECHOS RENTABLES','PANELES SOLARES',true],
      ['EXT004','HIDROSERVIS Y EQUIPOS SAS','BOMBAS DE RIEGO Y SUMINISTRO DE AGUA',true],
      ['EXT005','DERCO','VEHÍCULO KOMATSU',true],
      ['EXT006','NAVITRANS','VEHÍCULO MINICARGADOR (CASE)',true]
    ];
    var ts = new Date().toISOString();
    base.forEach(function(b){ h.appendRow([b[0],b[1],b[2],b[3],ts]); });
  }
  return h;
}

function _resp(d) {
  return ContentService.createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);
}
function _ok(d)  { return _resp(Object.assign({ok:true}, d||{})); }
function _err(m) { return _resp({ok:false, error:m}); }

// ── DEDUPLICAR OMs ─────────────────────────────────
function _deduplicar() {
  var h = _omsH();
  var rows = h.getDataRange().getValues();
  if (rows.length <= 1) return 0;

  // Agrupar por N_OM — guardar el índice de la fila ganadora
  var mapa = {}; // num → {fila_idx, prio, estado}
  for (var i = 1; i < rows.length; i++) {
    var num    = String(rows[i][0] || '').trim();
    var estado = String(rows[i][9] || 'PENDIENTE').trim();
    if (!num) continue;
    var prio = ESTADO_PRIO[estado] || 0;
    if (!mapa[num] || prio > mapa[num].prio) {
      mapa[num] = {idx: i, prio: prio, estado: estado};
    }
  }

  // Marcar filas a eliminar (las que NO son ganadoras)
  var ganadoras = {};
  for (var num in mapa) ganadoras[mapa[num].idx] = true;

  // Eliminar de abajo hacia arriba para no desplazar índices
  var borradas = 0;
  for (var j = rows.length - 1; j >= 1; j--) {
    var n = String(rows[j][0] || '').trim();
    if (n && !ganadoras[j]) {
      h.deleteRow(j + 1);
      borradas++;
    }
  }
  return borradas;
}


// ── INICIALIZACIÓN MANUAL ──────────────────────────────────────
// Ejecutar UNA VEZ desde el editor para crear todas las hojas
function inicializar() {
  _omsH();   // crea OMs-Pendientes si no existe
  _extH();   // crea Externos con los 6 proveedores base
  var hm = _ss().getSheetByName(HOJA_MAT);
  if (!hm) {
    hm = _ss().insertSheet(HOJA_MAT);
    hm.appendRow(['Fecha','OT','Equipo','Mecanico','Material','Cantidad','UM','Precio']);
  }
  Logger.log('✅ Hojas creadas: Hoja 1, OMs-Pendientes, Externos');
  return 'OK — revisa el Sheet, ya deben aparecer las 3 pestañas';
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var accion = p.accion || '';
  try {

    // ── CREAR OM ──────────────────────────────────
    if (accion === 'crear_om_get') {
      var h   = _omsH();
      var num = String(p.num || '');
      if (!num) return _err('Numero requerido');
      var rows = h.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === num && String(rows[i][9]) !== 'BORRADA')
          return _ok({ya_existe: true});
      }
      h.appendRow([
        num,
        p.fecha     || '',
        p.area      || '',
        p.equipo    || '',
        p.falla     || '',
        p.prioridad || 'MEDIA',
        String(p.mec_cc || ''),
        p.mec_nom   || '',
        p.obs_i     || '',
        'PENDIENTE',
        new Date().toISOString(),
        '', ''
      ]);
      // Deduplicar al crear
      _deduplicar();
      return _ok({creada: num});
    }

    // ── TOMAR OM ──────────────────────────────────
    if (accion === 'tomar_om_get') {
      var h2 = _omsH();
      var rows = h2.getDataRange().getValues();
      for (var j = 1; j < rows.length; j++) {
        if (String(rows[j][0]) === String(p.num||'') &&
            String(rows[j][9]) !== 'BORRADA' &&
            String(rows[j][9]) !== 'CERRADA') {
          h2.getRange(j+1, 10).setValue('TOMADA');
          h2.getRange(j+1, 12).setValue(new Date().toISOString());
          return _ok({tomada: p.num});
        }
      }
      return _err('OM no encontrada');
    }

    // ── ACTUALIZAR ESTADO ─────────────────────────
    if (accion === 'actualizar_estado_get') {
      var h3 = _omsH();
      var rows = h3.getDataRange().getValues();
      for (var k = 1; k < rows.length; k++) {
        if (String(rows[k][0]) === String(p.num||'') &&
            String(rows[k][9]) !== 'BORRADA') {
          h3.getRange(k+1, 10).setValue(p.estado || 'PENDIENTE');
          if (p.estado === 'CERRADA')
            h3.getRange(k+1, 13).setValue(new Date().toISOString());
          return _ok({actualizado: p.num, estado: p.estado});
        }
      }
      return _err('OM no encontrada: ' + p.num);
    }

    // ── BORRAR OM ─────────────────────────────────
    if (accion === 'borrar_om') {
      var h4 = _omsH();
      var rows = h4.getDataRange().getValues();
      for (var m = 1; m < rows.length; m++) {
        if (String(rows[m][0]) === String(p.num||'')) {
          h4.getRange(m+1, 10).setValue('BORRADA');
          return _ok({borrada: p.num});
        }
      }
      return _err('OM no encontrada: ' + p.num);
    }

    // ── DEDUPLICAR MANUAL ─────────────────────────
    if (accion === 'deduplicar') {
      var n = _deduplicar();
      return _ok({borradas: n});
    }

    // ── BORRAR MATERIALES DE UN OT (limpieza antes de re-enviar) ───────
    if (accion === 'borrar_ot_materiales') {
      var otDel = String(p.otNum || '');
      if (!otDel) return _err('otNum requerido');
      var hm = _ss().getSheetByName(HOJA_MAT);
      if (!hm) return _ok({ borradas: 0 });
      var rows = hm.getDataRange().getValues();
      var borradas = 0;
      for (var i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][1]) === otDel) {
          hm.deleteRow(i + 1);
          borradas++;
        }
      }
      return _ok({ borradas: borradas, otNum: otDel });
    }

    // ── GUARDAR MATERIAL → Hoja 1 ─────────────────
    if (accion === 'guardar_material_get') {
      var hm = _matH();
      if (!hm) {
        hm = _ss().insertSheet(HOJA_MAT);
        hm.appendRow(['Fecha','OT','Equipo','Mecanico','Material','Cantidad','UM','Precio']);
      }
      hm.appendRow([
        p.fecha    || new Date().toLocaleDateString('es-CO'),
        p.otNum    || '',
        p.equipo   || '',
        p.mecanico || '',
        p.desc     || '',
        Number(p.cant)   || 1,
        p.um       || 'Und',
        Number(p.precio) || 0
      ]);
      return _ok({guardado: true});
    }

    // ── LEER TODAS LAS OMs ────────────────────────
    if (accion === 'todas_oms') {
      var rows = _omsH().getDataRange().getValues();
      var oms = [];
      for (var a = 1; a < rows.length; a++) {
        var r = rows[a];
        if (!r[0]) continue;
        oms.push({
          num:r[0], fecha:r[1], area:r[2], equipo:r[3],
          falla:r[4], prioridad:r[5], mec_cc:String(r[6]),
          mec_nom:r[7], obs_i:r[8],
          estado:String(r[9]||'PENDIENTE'),
          tipo_mant:'CORRECTIVO'
        });
      }
      return _ok({oms: oms});
    }

    // ── LEER OMs DE UN MECÁNICO ───────────────────
    if (accion === 'mis_oms') {
      var cc = String(p.cc||'');
      var rows = _omsH().getDataRange().getValues();
      var oms = [];
      for (var b = 1; b < rows.length; b++) {
        var r = rows[b];
        var estado = String(r[9]||'PENDIENTE');
        if (String(r[6]) === cc && estado !== 'BORRADA') {
          oms.push({
            num:r[0], fecha:r[1], area:r[2], equipo:r[3],
            falla:r[4], prioridad:r[5], mec_cc:String(r[6]),
            mec_nom:r[7], obs_i:r[8], estado:estado,
            tipo_mant:'CORRECTIVO'
          });
        }
      }
      return _ok({oms: oms});
    }


    // ── LISTAR EXTERNOS ───────────────────────────────────────
    if (accion === 'listar_externos') {
      var rows = _extH().getDataRange().getValues();
      var externos = [];
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[0]) continue;
        externos.push({id:String(r[0]),nombre:String(r[1]),especialidad:String(r[2]),activo:r[3]!==false&&r[3]!=='FALSE'&&r[3]!==0});
      }
      return _ok({externos: externos});
    }

    // ── AGREGAR EXTERNO ───────────────────────────────────────
    if (accion === 'agregar_externo') {
      var id = String(p.id || '');
      var nombre = String(p.nombre || '');
      if (!id || !nombre) return _err('ID y nombre requeridos');
      var h = _extH();
      var rows = h.getDataRange().getValues();
      // Verificar duplicado
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === id) return _ok({ya_existe: true});
      }
      h.appendRow([id, nombre, p.especialidad || '', true, new Date().toISOString()]);
      return _ok({agregado: id});
    }


    // ── GET_DATOS — endpoint para Gerencia (sin parámetros) ───────
    // Gerencia llama fetch(GU) sin acción → retorna todos los registros de Hoja 1
    if (accion === '' || accion === 'get_datos') {
      var hm = _ss().getSheetByName(HOJA_MAT);
      if (!hm) return _ok({ datos: [] });
      var rows = hm.getDataRange().getValues();
      var datos = [];
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[0] && !r[1] && !r[4]) continue; // fila vacía
        if (String(r[4]) === 'Sin materiales' || String(r[4]) === '') continue;
        datos.push({
          fecha:    String(r[0] || ''),
          otNum:    String(r[1] || ''),
          equipo:   String(r[2] || ''),
          mecanico: String(r[3] || ''),
          desc:     String(r[4] || ''),
          cant:     Number(r[5] || 0),
          um:       String(r[6] || 'Und'),
          precio:   Number(r[7] || 0)
        });
      }
      return _ok({ datos: datos });
    }

    return _err('Accion no reconocida: ' + accion);

  } catch(ex) {
    return _err('Error: ' + ex.toString());
  }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents||'{}'); } catch(ex){}
  try {
    if (body.accion === 'limpiar_todo') {
      var h = _matH();
      if (h) { var l=h.getLastRow(); if(l>1) h.deleteRows(2,l-1); }
      return _ok({limpiado:true});
    }
    return _err('Accion POST no reconocida: '+body.accion);
  } catch(ex) {
    return _err('Error: '+ex.toString());
  }
}
