// ═══════════════════════════════════════════════════════════════════════════
//  AURA DE ORIENTE — Apps Script v4.4
//
//  INSTALACIÓN EN GOOGLE SHEETS NUEVO:
//  1. Crea un Google Sheets VACÍO nuevo
//  2. Menú → Extensiones → Apps Script
//  3. Pega este código completo, guarda (Ctrl+S)
//  4. Implementar → Nueva implementación → Aplicación web
//     · Ejecutar como: Yo  · Acceso: Cualquier persona
//  5. Autoriza permisos → Copia la URL /exec
//  6. En la app → ⚙️ → Pega URL → Guardar URL
//  7. En la app → ⚙️ → "Inicializar hojas"
//  8. En la app → ⚙️ → "Guardar en Google"
// ═══════════════════════════════════════════════════════════════════════════

const FOLDER_ID = ''; // ID carpeta Drive para imágenes ('' = auto)

const SHEETS = {
  productos:              ['id','clave','nombre','marca','fuente','precio','concentracion','genero','familia','uso','temporada','top','cor','fondo','desc','dur','proy','intensidad','extra','perfume_principal','imarca','inombre','mio','fav','pros','verificado','no_vendible','oculto','img_id','img_url','imgs','img_main_id','origen_id'],
  inventario:             ['id','pid','nombre','marca','tipo','pres','stock','costo','precio','nota','no_venta','consig','cuds','cprecio','cprov','ccom','ccond','cvenc','comprometidas'],
  ventas:                 ['id','fecha','cliente','clienteId','producto','tipoProd','presProd','marcaProd','invId','pkgId','uds','utype','precioU','descuento','total','pagado','quinc','prox','entregadas','esPaquete','observaciones','pkgProds','historial_entregas'],
  clientes:               ['id','nombre','tel','ciudad','notas','compras','saldo'],
  proveedores:            ['id','nombre','tel','concepto','total','pagado','tipo'],
  compras:                ['id','fecha','proveedor','provId','producto','prodId','tipo','pres','uds','costo','total','notas','estado','fechaRec','udsRecibidas','motivoDiff','anticipo','anticipo_forma'],
  paquetes:               ['id','nombre','desc','precio','normal','prods','status','fechaIni','fechaFin'],
  pedidos:                ['id','fecha','fechaEst','cliente','clienteId','articulo','productoId','proveedor','proveedorId','uds','precio','observaciones','cumplido','eliminado'],
  // pagos_proveedor removed — abonos are tracked in historial_prov (tipo:'abono')
  listas:                  null,  // Columnar format: headers=list names, values in rows
  historial_inv:          ['invId','producto','marca','tipo','pres','fecha','stockAntes','stockDespues','diferencia','motivo'],
  historial_prov:         ['provId','proveedor','fecha','tipo','monto','forma','motivo'],
  pagos_ventas:           ['ventaId','cliente','producto','tipoProd','presProd','fecha','monto','tipo','forma'],
  historial_ventas:       ['ventaId','cliente','producto','fecha','totalAntes','totalDespues','precioUAntes','precioUDespues','udsAntes','udsDespues','descAntes','descDespues','motivo'],
  confirmaciones_compras: ['compraId','proveedor','producto','fecha','udsPedidas','udsRecibidas','diferencia','motivo'],
  historial_compras:      ['compraId','proveedor','producto','fecha','tipo','cambios','motivo']
};

const MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
function normalizeDate(v) {
  if (!v && v !== 0) return ''; if (v === '-') return '-';
  const s = String(v).trim(); if (!s) return '';
  if (/^\d{1,2}\/[A-Za-z]{3}\/\d{4}$/.test(s)) return s;
  if (v instanceof Date && !isNaN(v)) return pad(v.getDate())+'/'+MESES[v.getMonth()]+'/'+v.getFullYear();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[3]+'/'+MESES[parseInt(iso[2])-1]+'/'+iso[1];
  const n = parseFloat(s);
  if (!isNaN(n) && n > 25000 && /^\d+(\.\d+)?$/.test(s)) {
    const d = new Date(Date.UTC(1899,11,30)+n*86400000);
    return pad(d.getUTCDate())+'/'+MESES[d.getUTCMonth()]+'/'+d.getUTCFullYear();
  }
  return s;
}
function pad(n) { return String(n).padStart(2,'0'); }
const DATE_COLS = new Set(['fecha','prox','cvenc','fechaRec','fechaIni','fechaFin']);
function serialize(col, val) {
  if (val===null||val===undefined) return '';
  if (val instanceof Date) return normalizeDate(val);
  if (typeof val==='object') return JSON.stringify(val);
  if (DATE_COLS.has(col)&&val!==''&&val!=='-') return normalizeDate(val);
  return val;
}
function rowToObj(cols, row) {
  const obj = {};
  cols.forEach((col,i) => {
    let v = row[i];
    if (v instanceof Date) v = normalizeDate(v);
    else if (typeof v==='string'&&/^\d{4}-\d{2}-\d{2}/.test(v)) v = normalizeDate(v);
    obj[col] = (v===null||v===undefined)?'':v;
  });
  return obj;
}
function getSpreadsheet() { return SpreadsheetApp.getActiveSpreadsheet(); }
function getImagesFolder() {
  if (FOLDER_ID) { try { return DriveApp.getFolderById(FOLDER_ID); } catch(e) {} }
  const name='Aura de Oriente - Imagenes';
  const f=DriveApp.getFoldersByName(name);
  return f.hasNext()?f.next():DriveApp.createFolder(name);
}
function getOrCreateSheet(ss,name,columns) {
  let s=ss.getSheetByName(name);
  if(!s){s=ss.insertSheet(name);writeHeader(s,columns);}
  return s;
}
function writeHeader(sheet,cols) {
  const r=sheet.getRange(1,1,1,cols.length);
  r.setValues([cols]);r.setBackground('#0e0b07').setFontColor('#c9922a').setFontWeight('bold');
  sheet.setFrozenRows(1);
}
function ensureColumns(sheet,required) {
  const lc=sheet.getLastColumn();
  const current=lc>0?sheet.getRange(1,1,1,lc).getValues()[0]:[];
  required.filter(c=>!current.includes(c)).forEach(col=>{
    const next=sheet.getLastColumn()+1;const cell=sheet.getRange(1,next);
    cell.setValue(col);cell.setBackground('#0e0b07').setFontColor('#c9922a').setFontWeight('bold');
  });
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
}
function upsertRow(sheet,cols,obj) {
  const arr=cols.map(col=>serialize(col,obj[col]));
  const idx=cols.indexOf('id');
  if(idx<0){sheet.appendRow(arr);return;}
  const newId=String(obj.id||'');const last=sheet.getLastRow();
  if(last>=2){
    const ids=sheet.getRange(2,idx+1,last-1,1).getValues();
    for(let i=0;i<ids.length;i++){if(String(ids[i][0])===newId){sheet.getRange(i+2,1,1,cols.length).setValues([arr]);return;}}
  }
  sheet.appendRow(arr);
}

function doGet(e){return route((e.parameter||{}).action||'',e.parameter||{},null);}
function doPost(e){let b={};try{b=JSON.parse(e.postData.contents);}catch(err){}return route(b.action||'',{},b);}
function ok(data){return ContentService.createTextOutput(JSON.stringify({ok:true,data})).setMimeType(ContentService.MimeType.JSON);}
function fail(msg){Logger.log('ERROR: '+msg);return ContentService.createTextOutput(JSON.stringify({ok:false,error:msg})).setMimeType(ContentService.MimeType.JSON);}
function route(action,params,body){
  try{
    switch(action){
      case 'ping':           return ok(doPing());
      case 'init_sheets':    return ok(doInitSheets());
      case 'save_all':       return ok(doSaveAll(body.data));
      case 'load':           return ok(doLoad());
      case 'load_listas':    var _ll=loadListas();return ok({listas_columnar:_ll,listas:Object.keys(_ll).map(function(k){return {key:k,valores:_ll[k].join('|')};})});
      case 'save_listas':    saveListas(body.data||{});return ok({saved:true});
      case 'generar_plantilla': return ok(doGenerarPlantilla());
      case 'upsert':         return ok(doUpsert(body.sheet,body.row));
      case 'upload_image':   return ok(doUploadImage(body.filename,body.base64,body.mime));
      case 'delete_image':   return ok(doDeleteImage(body.fileId));
      case 'scrape_perfume': return ok(doScrape(body.url));
      case 'debug_headers':  return ok(doDebugHeaders());
      default:               return fail('Accion desconocida: '+action);
    }
  }catch(e){return fail(e.message);}
}

function doPing(){
  const ss=getSpreadsheet();const vs=ss.getSheetByName('ventas');
  const lm=vs?vs.getLastRow()+'_'+(vs.getLastRow()>1?vs.getRange(vs.getLastRow(),1).getValue():'0'):'0_0';
  return{status:'ok',ts:new Date().toISOString(),last_modified:String(lm)};
}

function doInitSheets(){
  const ss=getSpreadsheet();const created=[],updated=[];
  const SAMPLE={
    productos:[
      {id:1,nombre:'Sheikh Al Layl',marca:'Miraj',fuente:'Oud Miraj',precio:950,concentracion:'EDP',intensidad:'Alta',genero:'Mujer',familia:'Oriental',uso:'Noche',temporada:'Invierno',top:'Cafe|Bergamota',cor:'Jazmin|Rosa',fondo:'Vainilla|Ambar',desc:'Fragancia oriental intensa con cafe negro y vainilla.',dur:'+12 horas',proy:'Excelente (deja estela)',extra:'',perfume_principal:'',imarca:'YSL',inombre:'Black Opium',mio:'TRUE',fav:'TRUE',pros:'FALSE',verificado:'TRUE',no_vendible:'FALSE',oculto:'FALSE',img_id:'',img_url:'',imgs:'[]',img_main_id:'',origen_id:''},
      {id:2,nombre:'Oud Al Malik',marca:'Lattafa',fuente:'Oud Miraj',precio:1100,concentracion:'Parfum',genero:'Hombre',familia:'Oud|Amaderado',uso:'Noche|Formal',temporada:'Invierno',top:'Pimienta|Cardamomo',cor:'Oud|Sandalo',fondo:'Vetiver|Ambar',desc:'Oud ahumado con sandalo. Distincion y poder.',dur:'+12 horas',proy:'Excelente (deja estela)',extra:'',perfume_principal:'',imarca:'Tom Ford',inombre:'Oud Wood',mio:'TRUE',fav:'FALSE',pros:'FALSE',verificado:'TRUE',no_vendible:'FALSE',oculto:'FALSE',img_id:'',img_url:'',imgs:'[]',img_main_id:'',origen_id:''},
      {id:3,nombre:'Noor Al Sabah',marca:'Al Haramain',fuente:'Fragrantica',precio:780,concentracion:'EDP',genero:'Unisex',familia:'Floral|Fresco',uso:'Dia|Casual',temporada:'Todo el anio',top:'Bergamota|Limon',cor:'Rosa|Peonia',fondo:'Almizcle|Sandalo',desc:'Floral fresco ligero. Uso diario.',dur:'6-8 horas',proy:'Buena',extra:'',perfume_principal:'',imarca:'Chanel',inombre:'Chance Eau Tendre',mio:'TRUE',fav:'FALSE',pros:'FALSE',verificado:'FALSE',no_vendible:'FALSE',oculto:'FALSE',img_id:'',img_url:'',imgs:'[]',img_main_id:'',origen_id:''}
    ],
    inventario:[
      {id:101,pid:1,nombre:'Sheikh Al Layl',marca:'Miraj',tipo:'Perfume',pres:'30ml',stock:8,costo:320,precio:950,nota:'Pedido inicial',consig:'FALSE',cuds:0,cprecio:0,cprov:'',ccom:0,ccond:'',cvenc:'',comprometidas:0},
      {id:102,pid:2,nombre:'Oud Al Malik',marca:'Lattafa',tipo:'Perfume',pres:'50ml',stock:5,costo:480,precio:1100,nota:'',consig:'TRUE',cuds:3,cprecio:1100,cprov:'Oud Miraj MX',ccom:20,ccond:'30 dias',cvenc:'30/Jun/2026',comprometidas:0},
      {id:103,pid:3,nombre:'Noor Al Sabah',marca:'Al Haramain',tipo:'Decant',pres:'10ml',stock:12,costo:95,precio:280,nota:'Decants propios',consig:'FALSE',cuds:0,cprecio:0,cprov:'',ccom:0,ccond:'',cvenc:'',comprometidas:0}
    ],
    ventas:[
      {id:301,fecha:'15/Abr/2026',cliente:'Ana Garcia',clienteId:201,producto:'Sheikh Al Layl Perfume 30ml',invId:101,pkgId:'',uds:2,utype:'propio',precioU:950,descuento:100,total:1700,pagado:900,quinc:200,prox:'15/May/2026',entregadas:0,esPaquete:'FALSE',observaciones:'Pago en abonos quincenales',pkgProds:'[]'},
      {id:302,fecha:'02/May/2026',cliente:'Laura Martinez',clienteId:202,producto:'Oud Al Malik Perfume 50ml',invId:102,pkgId:'',uds:1,utype:'propio',precioU:1100,descuento:0,total:1100,pagado:1100,quinc:200,prox:'-',entregadas:1,esPaquete:'FALSE',observaciones:'',pkgProds:'[]'}
    ],
    clientes:[
      {id:201,nombre:'Ana Garcia',tel:'8112345678',ciudad:'Monterrey',notas:'Prefiere florales',compras:1,saldo:800},
      {id:202,nombre:'Laura Martinez',tel:'8198765432',ciudad:'San Pedro',notas:'Cliente frecuente',compras:2,saldo:0},
      {id:203,nombre:'Sofia Reyes',tel:'8187654321',ciudad:'Guadalupe',notas:'Interesada en ouds',compras:0,saldo:0}
    ],
    proveedores:[
      {id:501,nombre:'Oud Miraj MX',tel:'',concepto:'Proveedor principal fragancias arabes',total:3200,pagado:1500,tipo:'compra'},
      {id:502,nombre:'Fragrantica Imports',tel:'',concepto:'Decants y muestras importadas',total:1275,pagado:0,tipo:'compra'}
    ],
    compras:[
      {id:401,fecha:'15/Abr/2026',proveedor:'Oud Miraj MX',provId:501,producto:'Sheikh Al Layl',prodId:1,tipo:'Perfume',pres:'30ml',uds:10,costo:320,total:3200,notas:'Pedido inicial',estado:'recibido',fechaRec:'17/Abr/2026',udsRecibidas:10,motivoDiff:''},
      {id:402,fecha:'04/May/2026',proveedor:'Fragrantica Imports',provId:502,producto:'Noor Al Sabah',prodId:3,tipo:'Decant',pres:'10ml',uds:15,costo:85,total:1275,notas:'Segundo pedido',estado:'pendiente',fechaRec:'',udsRecibidas:'',motivoDiff:''}
    ],
    pedidos:[],
    paquetes:[{id:601,nombre:'Kit Duo Oriental',desc:'Dos iconos orientales para noche',precio:1800,normal:2050,prods:'[{"invId":101,"nombre":"Sheikh Al Layl","tipo":"Perfume","pres":"30ml","uds":1},{"invId":102,"nombre":"Oud Al Malik","tipo":"Perfume","pres":"50ml","uds":1}]',status:'activo',fechaIni:'',fechaFin:''}],
    listas:[
      {key:'marcas_arabes',valores:'Miraj|Lattafa|Al Haramain|Rasasi|Ajmal|Armaf|Swiss Arabian|Nabeel|Ard Al Zaafaran|Zimaya'},
      {key:'marcas_inspiracion',valores:'YSL|Chanel|Dior|Tom Ford|Creed|Lancome|Guerlain|Armani|Burberry|Paco Rabanne|Versace|Givenchy|Carolina Herrera|Hugo Boss|Davidoff'},
      {key:'familias',valores:'Oriental|Floral|Amaderado|Fresco|Oud|Gourmand|Acuatico|Citrico|Chipre|Fougere|Aromatico'},
      {key:'concentraciones',valores:'EDP|EDT|EDC|Parfum|Attar|Oil|Extrait'},
      {key:'presentaciones',valores:'5ml|10ml|15ml|20ml|30ml|50ml|60ml|75ml|100ml|125ml|150ml|200ml'},
      {key:'notas_top',valores:'Bergamota|Limon|Mandarina|Naranja|Pomelo|Pimienta|Cardamomo|Azafran|Lavanda|Menta|Jengibre|Nuez moscada|Canela'},
      {key:'notas_corazon',valores:'Rosa|Jazmin|Iris|Ylang Ylang|Geranio|Violeta|Tuberosa|Lirio|Peonia|Orquidea|Oud|Neroli|Azahar'},
      {key:'notas_fondo',valores:'Vainilla|Sandalo|Almizcle|Ambar|Vetiver|Cedro|Pachuli|Benjui|Incienso|Mirra|Cuero|Balsamo|Tonka'},
      {key:'usos',valores:'Dia|Noche|Casual|Formal|Especial|Oficina|Deportivo|Romantico'}
    ],
    historial_inv:[
      {invId:101,producto:'Sheikh Al Layl',marca:'Miraj',tipo:'Perfume',pres:'30ml',fecha:'17/Abr/2026',stockAntes:0,stockDespues:10,diferencia:10,motivo:'Recepcion compra #401'},
      {invId:102,producto:'Oud Al Malik',marca:'Lattafa',tipo:'Perfume',pres:'50ml',fecha:'17/Abr/2026',stockAntes:0,stockDespues:5,diferencia:5,motivo:'Carga inicial'}
    ],
    historial_prov:[
      {provId:501,proveedor:'Oud Miraj MX',fecha:'17/Abr/2026',tipo:'compra_recibida',monto:3200,motivo:'Sheikh Al Layl x10'},
      {provId:501,proveedor:'Oud Miraj MX',fecha:'25/Abr/2026',tipo:'abono',monto:1500,motivo:'Pago parcial SPEI'}
    ],
    pagos_ventas:[
      {ventaId:301,cliente:'Ana Garcia',producto:'Sheikh Al Layl Perfume 30ml',fecha:'15/Abr/2026',monto:700,tipo:'real'},
      {ventaId:301,cliente:'Ana Garcia',producto:'Sheikh Al Layl Perfume 30ml',fecha:'30/Abr/2026',monto:200,tipo:'real'},
      {ventaId:302,cliente:'Laura Martinez',producto:'Oud Al Malik Perfume 50ml',fecha:'02/May/2026',monto:1100,tipo:'real'}
    ],
    historial_ventas:[],
    confirmaciones_compras:[{compraId:401,proveedor:'Oud Miraj MX',producto:'Sheikh Al Layl',fecha:'17/Abr/2026',udsPedidas:10,udsRecibidas:10,diferencia:0,motivo:''}],
    historial_compras:[{compraId:401,proveedor:'Oud Miraj MX',producto:'Sheikh Al Layl',fecha:'15/Abr/2026',tipo:'creacion',cambios:'Pedido creado',motivo:'Registro inicial'}]
  };

  Object.entries(SHEETS).forEach(([name,cols])=>{
    if(!cols){
      // Special-format sheet (e.g. listas) — just ensure it exists; saveListas handles content
      if(!ss.getSheetByName(name)){ss.insertSheet(name);created.push(name);}
      else{updated.push(name);}
      return;
    }
    const ex=ss.getSheetByName(name);
    if(!ex){
      const s=ss.insertSheet(name);writeHeader(s,cols);
      const sample=SAMPLE[name];
      if(sample&&sample.length>0){
        const rows=sample.map(rec=>cols.map(col=>serialize(col,rec[col])));
        s.getRange(2,1,rows.length,cols.length).setValues(rows);
      }
      created.push(name);
    }else{ensureColumns(ex,cols);updated.push(name);}
  });

  ['Sheet1','Hoja1','Hoja 1'].forEach(n=>{
    const s=ss.getSheetByName(n);
    if(s&&s.getLastRow()<=1&&ss.getSheets().length>1){try{ss.deleteSheet(s);}catch(e){}}
  });

  // Sheet reorder removed — caused crashes with user's extra sheets.
  // App sheets are created/updated above; order in Sheets UI is cosmetic only.
  Logger.log('doInitSheets complete. Created: '+created.join(',')+' Updated: '+updated.join(','));

  // Initialize listas — merge defaults with existing, never overwrite
  const _defaultListas={marcas_arabes:['Miraj','Lattafa','Al Haramain','Rasasi','Armaf'],
    marcas_inspiracion:['Chanel','YSL','Dior','Tom Ford','Creed'],
    familias:['Oriental','Floral','Amaderado','Oud','Gourmand','Fresco','Citrico'],
    concentraciones:['EDP','EDT','Parfum','EDP Intense','Attar'],
    presentaciones:['30ml','50ml','60ml','100ml','125ml','3ml decant','10ml decant'],
    notas_top:['Bergamota','Limon','Naranja','Menta','Cardamomo','Azafran'],
    notas_corazon:['Rosa','Jazmin','Iris','Lavanda','Geranio','Ylang Ylang'],
    notas_fondo:['Oud','Ambar','Vainilla','Almizcle','Patchouli','Cedro','Santal'],
    usos:['Dia','Noche','Casual','Formal','Especial','Oficina','Deportivo'],
    duraciones:['Menos de 2h','2-4h','4-6h','6-8h','8-12h','Mas de 12h'],
    proyecciones:['Muy cercana','Cercana','Moderada','Notable','Intensa','Bestial'],
    temporadas:['Primavera','Verano','Otono','Invierno','Todo el anio'],
    fuentes:['Oud Miraj','Fragrantica','Proveedor propio','Otro'],
    intensidades:['Ligera','Media','Alta','Fuerte'],
    formas_pago:['Efectivo','Transferencia','Tarjeta'],
    tipos_producto:['Perfume','Decant','Tester']};
  const _existing=loadListas();
  // Only add default keys that don't exist yet — never overwrite user data
  Object.keys(_defaultListas).forEach(k=>{
    if(!_existing[k]||_existing[k].length===0)_existing[k]=_defaultListas[k];
  });
  saveListas(_existing);
  return{created,updated,url:ss.getUrl(),total:Object.keys(SHEETS).length};
}

function loadListas(){
  const ss=getSpreadsheet();
  const sheet=ss.getSheetByName('listas');
  if(!sheet||sheet.getLastRow()<1)return {};
  const lc=sheet.getLastColumn(),lr=sheet.getLastRow();
  if(lc<1)return {};
  const all=sheet.getRange(1,1,lr,lc).getValues();
  const keys=all[0];
  const result={};
  keys.forEach(function(k,ci){
    if(!k)return;
    result[String(k)]=[];
    for(var ri=1;ri<all.length;ri++){
      var val=all[ri][ci];
      if(val!==''&&val!==null&&val!==undefined)result[String(k)].push(String(val));
    }
  });
  return result;
}

function saveListas(listsObj){
  const ss=getSpreadsheet();
  let sheet=ss.getSheetByName('listas');
  if(!sheet)sheet=ss.insertSheet('listas');
  sheet.clearContents();
  const keys=Object.keys(listsObj).filter(k=>Array.isArray(listsObj[k])&&listsObj[k].length>0);
  if(!keys.length)return;
  const maxLen=Math.max(...keys.map(k=>listsObj[k].length));
  const data=[keys];
  for(let i=0;i<maxLen;i++)
    data.push(keys.map(k=>listsObj[k][i]!==undefined?listsObj[k][i]:''));
  sheet.getRange(1,1,data.length,keys.length).setValues(data);
  const hdr=sheet.getRange(1,1,1,keys.length);
  hdr.setBackground('#0e0b07');hdr.setFontColor('#c9922a');
  hdr.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function doSaveAll(data){
  if(!data)throw new Error('Sin datos');
  const ss=getSpreadsheet();const saved={};
function save(name,records){
  if(records===null||records===undefined)return;  // skip null/undefined
    if(!SHEETS[name])return;  // skip special-format sheets (e.g. listas)
    const sheet=getOrCreateSheet(ss,name,SHEETS[name]);
    const cols=ensureColumns(sheet,SHEETS[name]);
    if(sheet.getLastRow()>1)sheet.getRange(2,1,sheet.getLastRow()-1,cols.length).clearContent();
    if(records.length>0){const rows=records.map(rec=>cols.map(col=>serialize(col,rec[col])));sheet.getRange(2,1,rows.length,cols.length).setValues(rows);}
    saved[name]=records.length;
  }
  save('productos',              data.productos              ||[]);
  save('inventario',             data.inventario             ||[]);
  save('ventas',                 data.ventas                 ||[]);
  save('clientes',               data.clientes               ||[]);
  save('proveedores',            data.proveedores            ||[]);
  save('compras',                data.compras                ||[]);
  save('paquetes',               data.paquetes               ||[]);
  save('pedidos',                data.pedidos                ||[]);
  // pagos_proveedor removed — no longer a separate sheet
  saveListas(data.listas_columnar||data.listas||{});  // columnar format
  save('historial_inv',          data.historial_inv          ||[]);
  save('historial_prov',         data.historial_prov         ||[]);
  save('pagos_ventas',           data.pagos_ventas           ||[]);
  save('historial_ventas',       data.historial_ventas       ||[]);
  save('confirmaciones_compras', data.confirmaciones_compras ||[]);
  save('historial_compras',      data.historial_compras      ||[]);
  return{saved,ts:normalizeDate(new Date())};
}

function doLoad(){
  const ss=getSpreadsheet();const result={};
  Object.keys(SHEETS).forEach(name=>{
    if(!SHEETS[name]){result[name]=[];return;}  // skip special-format sheets (listas)
    const sheet=ss.getSheetByName(name);
    if(!sheet){result[name]=[];return;}
    const lr=sheet.getLastRow();const lc=sheet.getLastColumn();
    if(lr<2||lc<1){result[name]=[];return;}  // empty sheet guard
    const cols=sheet.getRange(1,1,1,lc).getValues()[0].filter(h=>h!=='');
    if(!cols.length){result[name]=[];return;}
    const rows=sheet.getRange(2,1,lr-1,cols.length).getValues();
    result[name]=rows.map(r=>rowToObj(cols,r)).filter(o=>String(o[cols[0]]||'')!=='');
  });
  // Reconstituir arrays en los objetos padre
  result.inventario=(result.inventario||[]).map(i=>({...i,ajustes:(result.historial_inv||[]).filter(h=>String(h.invId)===String(i.id))}));
  result.proveedores=(result.proveedores||[]).map(p=>({...p,historial:(result.historial_prov||[]).filter(h=>String(h.provId)===String(p.id))}));
  result.ventas=(result.ventas||[]).map(v=>({...v,pagos:(result.pagos_ventas||[]).filter(p=>String(p.ventaId)===String(v.id)),ajustes:(result.historial_ventas||[]).filter(a=>String(a.ventaId)===String(v.id))}));
  result.compras=(result.compras||[]).map(c=>({...c,confirmaciones:(result.confirmaciones_compras||[]).filter(cf=>String(cf.compraId)===String(c.id)),historial:(result.historial_compras||[]).filter(h=>String(h.compraId)===String(c.id))}));
  // pagos_prov derived from historial_prov filtered by tipo='abono'
  // Load listas in columnar format
  result.listas_columnar=loadListas();
  result.listas=Object.keys(result.listas_columnar).map(function(k){return {key:k,valores:(result.listas_columnar[k]||[]).join('|')};});
  result.pagos_prov=(result.historial_prov||[]).filter(h=>h.tipo==='abono');
  return result;
}

function doUpsert(sheetName,row){
  if(!sheetName||!row)throw new Error('sheet y row requeridos');
  const ss=getSpreadsheet();
  const cols=SHEETS[sheetName]||Object.keys(row);
  const sheet=getOrCreateSheet(ss,sheetName,cols);
  upsertRow(ensureColumns(sheet,cols)===cols?sheet:sheet,ensureColumns(sheet,cols),row);
  return{sheet:sheetName,id:row.id||'N/A'};
}

function doUploadImage(filename,base64Data,mimeType){
  if(!base64Data)throw new Error('base64 requerido');
  if(!filename)filename='aura_'+Date.now()+'.jpg';
  if(!mimeType)mimeType='image/jpeg';
  const folder=getImagesFolder();
  const blob=Utilities.newBlob(Utilities.base64Decode(base64Data),mimeType,filename);
  const file=folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  const fileId=file.getId();
  return{fileId,url:'https://lh3.googleusercontent.com/d/'+fileId,filename,folderId:folder.getId()};
}

function doDeleteImage(fileId){
  if(!fileId)return{deleted:false};
  try{DriveApp.getFileById(fileId).setTrashed(true);return{deleted:true,fileId};}
  catch(e){return{deleted:false,reason:e.message};}
}

function doDebugHeaders(){
  const ss=getSpreadsheet();const out={};
  ss.getSheets().forEach(s=>{
    const lc=s.getLastColumn();
    if(lc>0)try{out[s.getName()]=s.getRange(1,1,1,lc).getValues()[0];}catch(e){out[s.getName()]=['ERROR: '+e.message];}
  });
  return out;
}

// ── SCRAPE ─────────────────────────────────────────────────────────────────
function doScrape(url){
  if(!url)throw new Error('URL requerida');
  const cleanUrl=url.split('?')[0].replace(/\/$/,'');
  let productData=null;

  // Intentar Shopify JSON endpoint primero
  try{
    const jr=UrlFetchApp.fetch(cleanUrl+'.json',{muteHttpExceptions:true,headers:{'User-Agent':'Mozilla/5.0 (compatible)'}});
    if(jr.getResponseCode()===200){const j=JSON.parse(jr.getContentText());productData=j.product||null;}
  }catch(e){Logger.log('JSON endpoint: '+e.message);}

  const resp=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:true,headers:{
    'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Accept':'text/html,application/xhtml+xml;q=0.9,*/*;q=0.8','Accept-Language':'es-MX,es;q=0.9,en-US;q=0.8','Cache-Control':'no-cache'
  }});
  if(resp.getResponseCode()!==200)throw new Error('HTTP '+resp.getResponseCode());
  const html=resp.getContentText();

  function clean(s){if(!s)return '';return s.replace(/<[^>]+>/g,' ').replace(/&amp;/g,'&').replace(/&nbsp;/g,' ').replace(/&#\d+;/g,c=>String.fromCharCode(parseInt(c.slice(2,-1)))).replace(/&[a-z]+;/g,' ').replace(/\s+/g,' ').trim();}

  const d={fuente:'',nombre:'',marca:'',imarca:'',inombre:'',desc:'',familia:[],top:[],cor:[],fondo:[],genero:'Unisex',concentracion:'',dur:'',proy:''};

  if(url.includes('oudmiraj')){
    d.fuente='Oud Miraj';
    if(productData){
      d.nombre=clean(productData.title||'');
      d.marca=clean(productData.vendor||'Oud Miraj');
      const rawBody=productData.body_html||'';
      d.desc=clean(rawBody).substring(0,600);
      const tags=Array.isArray(productData.tags)?productData.tags:String(productData.tags||'').split(',').map(t=>t.trim()).filter(Boolean);
      tags.forEach(tag=>{const t=tag.toLowerCase().trim();
        if(/mujer|woman|women|femenin/.test(t))d.genero='Mujer';
        else if(/hombre|^man$|^men$|masculin/.test(t))d.genero='Hombre';
        if(/^edp$|eau\s*de\s*parfum/.test(t))d.concentracion='EDP';
        else if(/^edt$/.test(t))d.concentracion='EDT';
        else if(/^parfum$|extrait/.test(t))d.concentracion='Parfum';
        else if(/attar|oil\b/.test(t))d.concentracion='Attar';
        if(/oriental/.test(t)&&!d.familia.includes('Oriental'))d.familia.push('Oriental');
        if(/\bfloral\b/.test(t)&&!d.familia.includes('Floral'))d.familia.push('Floral');
        if(/\boud\b/.test(t)&&!d.familia.includes('Oud'))d.familia.push('Oud');
        if(/wood|amader/.test(t)&&!d.familia.includes('Amaderado'))d.familia.push('Amaderado');
      });
      parseNotes_(rawBody||d.desc,d);
    }else{
      const h1=html.match(/<h1[^>]*class="[^"]*(?:product[_-]?title|title)[^"]*"[^>]*>([\s\S]{2,120}?)<\/h1>/i)||html.match(/<h1[^>]*>([\s\S]{2,120}?)<\/h1>/i);
      if(h1)d.nombre=clean(h1[1]);
      const vend=html.match(/<(?:span|a|div)[^>]*class="[^"]*vendor[^"]*"[^>]*>([\s\S]{2,50}?)<\/(?:span|a|div)>/i);
      d.marca=vend?clean(vend[1]):'Oud Miraj';
      const dp=[
        /<div[^>]+class="[^"]*(?:product-description|product__description|rte)[^"]*"[^>]*>([\s\S]{20,3000}?)<\/div>/i,
        /<meta[^>]+name="description"[^>]+content="([^"]{20,500})"/i,
        /<meta[^>]+property="og:description"[^>]+content="([^"]{20,500})"/i
      ];
      for(const p of dp){const m=html.match(p);if(m){d.desc=clean(m[1]).substring(0,600);break;}}
      if(/para\s+mujer|for\s+women/i.test(html))d.genero='Mujer';
      else if(/para\s+hombre|for\s+men/i.test(html))d.genero='Hombre';
      if(/\bEDP\b|Eau\s+de\s+Parfum/i.test(html))d.concentracion='EDP';
      else if(/\bEDT\b/i.test(html))d.concentracion='EDT';
      else if(/\bParfum\b|Extrait/i.test(html))d.concentracion='Parfum';
      parseNotes_(html,d);
    }
    // Inspiración
    const BRANDS='(?:YSL|Yves Saint Laurent|Chanel|Dior|Tom Ford|Creed|Guerlain|Armani|Lancome|Hermes|Versace|Givenchy|Paco Rabanne|Hugo Boss|Burberry|Carolina Herrera|Davidoff|Bvlgari|Prada)';
    const inspo=[
      /inspired?\s+by[\s:]+([^.<"\n]{3,80})/i,
      /inspirad[ao]\s+en[\s:]+([^.<"\n]{3,80})/i,
      /similar\s+(?:to|a)[\s:]+([^.<"\n]{3,80})/i,
      new RegExp('('+BRANDS+')\\s*[-–]\\s*([^.<"\\n]{3,60})','i')
    ];
    for(const p of inspo){
      const m=html.match(p)||(d.desc&&d.desc.match(p));
      if(m){
        if(m[1]&&m[2]){d.imarca=clean(m[1]);d.inombre=clean(m[2]);}
        else{d.inombre=clean(m[1]||'').substring(0,60);}
        if(d.inombre)break;
      }
    }
  }else if(url.includes('fragrantica')){
    d.fuente='Fragrantica';
    const h1=html.match(/<h1[^>]*itemprop="name"[^>]*>([\s\S]{2,80}?)<\/h1>/i)||html.match(/<h1[^>]*>([\s\S]{2,80}?)<\/h1>/i);
    if(h1)d.nombre=clean(h1[1]);
    const brand=html.match(/itemprop="brand"[\s\S]{0,400}?<span[^>]*>([\s\S]{2,60}?)<\/span>/i);
    if(brand)d.imarca=clean(brand[1]);
    if(/para\s+mujer|for\s+women/i.test(html))d.genero='Mujer';
    else if(/para\s+hombre|for\s+men/i.test(html))d.genero='Hombre';
    parseNotes_(html,d);
  }else{
    throw new Error('Solo Oud Miraj y Fragrantica son soportados. URL: '+url);
  }

  d.familia=[...new Set(d.familia)].slice(0,4);
  d.top=[...new Set(d.top)].slice(0,8);
  d.cor=[...new Set(d.cor)].slice(0,8);
  d.fondo=[...new Set(d.fondo)].slice(0,8);
  d.nombre=d.nombre.replace(/\s+/g,' ').trim();
  if(!d.nombre)throw new Error('No se pudo extraer el nombre. Agrega manualmente.');
  Logger.log('Scrape OK: '+d.nombre+' | top:'+d.top.join(','));
  return d;
}

function parseNotes_(text,d){
  function get(pat){
    const m=text.match(new RegExp(pat,'i'));if(!m)return[];
    // Strip HTML, split on ALL common separators to get EVERY note
    const raw=(m[1]||'').replace(/<[^>]+>/g,' ').replace(/&[a-z]+;/g,' ');
    return raw.split(/[,;|\/\u00b7\u2022\n]/).map(n=>n.replace(/\s+/g,' ').trim()).filter(n=>n.length>2&&n.length<45&&!/^\d+$/.test(n)&&n!=='and'&&n!=='y').slice(0,12);
  }
  if(!d.top.length)d.top=get('(?:top|opening|salida|apertura)\\s*notes?\\s*[:\\-\u2013]\\s*([^<\\n\\.]{5,400})');
  if(!d.cor.length)d.cor=get('(?:heart|middle|coraz[o\u00f3]n|medio)\\s*notes?\\s*[:\\-\u2013]\\s*([^<\\n\\.]{5,400})');
  if(!d.fondo.length)d.fondo=get('(?:base|fond[eo])\\s*notes?\\s*[:\\-\u2013]\\s*([^<\\n\\.]{5,400})');
  if(!d.top.length&&!d.cor.length&&!d.fondo.length){
    const gen=get('(?:notas?|notes?)\\s*[:\\-\u2013]\\s*([^<\\n\\.]{5,500})');
    if(gen.length>0){const t=Math.ceil(gen.length/3);d.top=gen.slice(0,t);d.cor=gen.slice(t,t*2);d.fondo=gen.slice(t*2);}
  }
  const plain=text.replace(/<[^>]+>/g,' ').toLowerCase();
  if(/\boriental\b/.test(plain)&&!d.familia.includes('Oriental'))d.familia.push('Oriental');
  if(/\bfloral\b/.test(plain)&&!d.familia.includes('Floral'))d.familia.push('Floral');
  if(/\boud\b|\baoud\b/.test(plain)&&!d.familia.includes('Oud'))d.familia.push('Oud');
  if(/\bwood(y)?\b|\bamader/.test(plain)&&!d.familia.includes('Amaderado'))d.familia.push('Amaderado');
  if(/\bfresh\b|\bfresc\b|\bcitrus\b/.test(plain)&&!d.familia.includes('Fresco'))d.familia.push('Fresco');
  if(/\bgourm|\bvaini|\bchocol/.test(plain)&&!d.familia.includes('Gourmand'))d.familia.push('Gourmand');
}

function testPing(){Logger.log(JSON.stringify(doPing()));}
function testInit(){Logger.log(JSON.stringify(doInitSheets()));}
function testLoad(){const d=doLoad();Logger.log('Productos:'+d.productos.length+' Ventas:'+d.ventas.length+' Inv:'+d.inventario.length);}
function testScrape(){
  try{const r=doScrape('https://oudmiraj.com/products/bharara-pharaoh-ramesses-i-edp');Logger.log(JSON.stringify({nombre:r.nombre,marca:r.marca,top:r.top,inspo:r.imarca+' '+r.inombre}));}
  catch(e){Logger.log('ERROR: '+e.message);}
}

// ═══════════════════════════════════════════════════════════════════════════
//  GENERAR PLANTILLA DE IMPORTACION (Google Sheets con validaciones)
// ═══════════════════════════════════════════════════════════════════════════
function doGenerarPlantilla() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Delete existing plantilla spreadsheet if any (by name) ───────────────
  var PLANTILLA_NAME = 'Plantilla Aura de Oriente';
  var existing = DriveApp.getFilesByName(PLANTILLA_NAME);
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  // ── Create new spreadsheet ────────────────────────────────────────────────
  var pl = SpreadsheetApp.create(PLANTILLA_NAME);
  var plId = pl.getId();

  // ── Read LISTS from the main sheet ───────────────────────────────────────
  var listas = loadListas();  // returns {marcas_arabes:[...], concentraciones:[...], ...}
  var intensidades = listas.intensidades && listas.intensidades.length
    ? listas.intensidades : ['Ligera','Media','Alta','Fuerte'];
  var fuentes = ['Oud Miraj','Fragrantica','Proveedor propio','Otro'];
  var generos = ['Mujer','Hombre','Unisex'];
  var sinono  = ['SI','NO'];

  // ── Column definitions (order matters — drives validation ranges) ─────────
  var listsDef = [
    {key:'marcas_arabes',      vals: listas.marcas_arabes      || []},
    {key:'concentraciones',    vals: listas.concentraciones    || []},
    {key:'marcas_inspiracion', vals: listas.marcas_inspiracion || []},
    {key:'familias',           vals: listas.familias           || []},
    {key:'usos',               vals: listas.usos               || []},
    {key:'temporadas',         vals: listas.temporadas         || []},
    {key:'duraciones',         vals: listas.duraciones         || []},
    {key:'proyecciones',       vals: listas.proyecciones       || []},
    {key:'intensidades',       vals: intensidades},
    {key:'notas_top',          vals: listas.notas_top          || []},
    {key:'notas_corazon',      vals: listas.notas_corazon      || []},
    {key:'notas_fondo',        vals: listas.notas_fondo        || []},
    {key:'si_no',              vals: sinono},
    {key:'genero',             vals: generos},
    {key:'fuente',             vals: fuentes},
  ];

  // ── BUILD LISTAS SHEET ───────────────────────────────────────────────────
  var shListas = pl.getSheets()[0];
  shListas.setName('Listas');
  shListas.setTabColor('#4a90d9');

  var maxRows = 0;
  listsDef.forEach(function(d){ if(d.vals.length > maxRows) maxRows = d.vals.length; });

  // Write headers (row 1)
  var listsHeaders = listsDef.map(function(d){ return d.key; });
  shListas.getRange(1, 1, 1, listsHeaders.length).setValues([listsHeaders])
    .setFontWeight('bold').setBackground('#2a5f5f').setFontColor('#ffffff');

  // Write values
  for (var r = 0; r < maxRows; r++) {
    var row = listsDef.map(function(d){ return d.vals[r] || ''; });
    shListas.getRange(r + 2, 1, 1, row.length).setValues([row]);
  }
  shListas.setFrozenRows(1);
  // Auto-resize columns
  for (var c = 1; c <= listsDef.length; c++) shListas.autoResizeColumn(c);
  // Protect Listas sheet (warning only — can't fully lock without owner)
  var prot = shListas.protect().setDescription('Lista de valores de referencia');
  prot.setWarningOnly(true);

  // Helper: get column letter for a listsDef key (1-based col index → A, B, ...)
  function listsCol(key) {
    for (var i = 0; i < listsDef.length; i++) {
      if (listsDef[i].key === key) {
        return String.fromCharCode(65 + i); // A=0, B=1, ...
      }
    }
    return 'A';
  }
  function listsRange(key) {
    var c = listsCol(key);
    var n = 0;
    listsDef.forEach(function(d){ if(d.key===key) n = d.vals.length; });
    return 'Listas!$' + c + '$2:$' + c + '$' + (n + 1);
  }

  // ── BUILD PERFUMES SHEET ─────────────────────────────────────────────────
  var shPerf = pl.insertSheet('Perfumes', 0); // insert at position 0 (first)
  shPerf.setTabColor('#c9922a');

  var headers = [
    'nombre_arabe','marca_arabe','concentracion','fuente','precio',
    'inspo_marca','inspo_nombre','genero','familia_olfativa','uso_sugerido',
    'temporada','notas_salida','notas_corazon','notas_fondo',
    'descripcion','duracion','proyeccion','intensidad','notas_extra',
    'mi_negocio','favorito','prospecto','verificado','no_vendible','oculto'
  ];
  // Map header name → 1-based column index
  var hIdx = {};
  headers.forEach(function(h,i){ hIdx[h] = i+1; });

  var DATA_START = 6;  // first data row (after 4 info rows + 1 header row)
  var DATA_END   = 105; // last data row (100 data rows)
  var NUM_COLS   = headers.length;

  // Info rows
  shPerf.getRange(1,1).setValue('AURA DE ORIENTE - Plantilla de importacion incremental de Perfumes')
    .setFontWeight('bold').setFontSize(12);
  shPerf.getRange(2,1).setValue('Para multiples valores en un campo usa | como separador. Ej: Oriental|Floral')
    .setFontColor('#7a6a55').setFontStyle('italic');
  shPerf.getRange(3,1).setValue('Tipo, Presentacion e Inventario se configuran por separado en la seccion Inventario')
    .setFontColor('#7a6a55').setFontStyle('italic');
  shPerf.getRange(1,1,1,NUM_COLS).merge();
  shPerf.getRange(2,1,1,NUM_COLS).merge();
  shPerf.getRange(3,1,1,NUM_COLS).merge();

  // Header row (row 5)
  var hdrRange = shPerf.getRange(5, 1, 1, NUM_COLS);
  hdrRange.setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1e1810')
    .setFontColor('#f2c46a')
    .setHorizontalAlignment('center');

  // Sample row (row 6)
  var sample = [
    'Sheikh Al Layl','Miraj','EDP','Oud Miraj','950','YSL','Black Opium',
    'Mujer','Oriental','Noche','Invierno',
    'Cafe|Bergamota','Jazmin|Rosa','Vainilla|Ambar',
    'Fragancia oriental intensa.','+12 horas','Excelente (deja estela)','Alta','Top ventas',
    'SI','SI','NO','SI','NO','NO'
  ];
  shPerf.getRange(6, 1, 1, NUM_COLS).setValues([sample])
    .setBackground('#fef3c7').setFontStyle('italic');

  // Freeze header row and first column
  shPerf.setFrozenRows(5);
  shPerf.setFrozenColumns(1);

  // Column widths
  shPerf.setColumnWidth(hIdx['nombre_arabe'],   180);
  shPerf.setColumnWidth(hIdx['descripcion'],    300);
  shPerf.setColumnWidth(hIdx['notas_extra'],    200);
  shPerf.setColumnWidth(hIdx['inspo_nombre'],   160);
  for (var ci = 1; ci <= NUM_COLS; ci++) {
    if (shPerf.getColumnWidth(ci) < 120) shPerf.setColumnWidth(ci, 120);
  }

  // Alternate row banding on data area
  shPerf.getRange(DATA_START, 1, DATA_END - DATA_START + 1, NUM_COLS)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  // ── DATA VALIDATIONS with dropdown ───────────────────────────────────────
  // Columns that get a dropdown (single-value fields)
  var dvCols = [
    {h:'marca_arabe',   listKey:'marcas_arabes'},
    {h:'concentracion', listKey:'concentraciones'},
    {h:'fuente',        listKey:'fuente'},
    {h:'inspo_marca',   listKey:'marcas_inspiracion'},
    {h:'genero',        listKey:'genero'},
    {h:'duracion',      listKey:'duraciones'},
    {h:'proyeccion',    listKey:'proyecciones'},
    {h:'intensidad',    listKey:'intensidades'},
    {h:'mi_negocio',    listKey:'si_no'},
    {h:'favorito',      listKey:'si_no'},
    {h:'prospecto',     listKey:'si_no'},
    {h:'verificado',    listKey:'si_no'},
    {h:'no_vendible',   listKey:'si_no'},
    {h:'oculto',        listKey:'si_no'},
  ];
  dvCols.forEach(function(dv) {
    var col = hIdx[dv.h];
    var range = shPerf.getRange(DATA_START, col, DATA_END - DATA_START + 1, 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(
        shListas.getRange(listsRange(dv.listKey).replace('Listas!','')),
        true  // showDropdown = true
      )
      .setAllowInvalid(true)   // allow values not in list (shows warning icon only)
      .setHelpText('Selecciona un valor de la lista o escribe uno personalizado')
      .build();
    range.setDataValidation(rule);
  });

  // ── CONDITIONAL FORMATTING: red background if value not in list ───────────
  // For ALL validated columns including multi-value ones
  var cfCols = [
    {h:'marca_arabe',     listKey:'marcas_arabes'},
    {h:'concentracion',   listKey:'concentraciones'},
    {h:'fuente',          listKey:'fuente'},
    {h:'inspo_marca',     listKey:'marcas_inspiracion'},
    {h:'genero',          listKey:'genero'},
    {h:'familia_olfativa',listKey:'familias'},
    {h:'uso_sugerido',    listKey:'usos'},
    {h:'temporada',       listKey:'temporadas'},
    {h:'notas_salida',    listKey:'notas_top'},
    {h:'notas_corazon',   listKey:'notas_corazon'},
    {h:'notas_fondo',     listKey:'notas_fondo'},
    {h:'duracion',        listKey:'duraciones'},
    {h:'proyeccion',      listKey:'proyecciones'},
    {h:'intensidad',      listKey:'intensidades'},
    {h:'mi_negocio',      listKey:'si_no'},
    {h:'favorito',        listKey:'si_no'},
    {h:'prospecto',       listKey:'si_no'},
    {h:'verificado',      listKey:'si_no'},
    {h:'no_vendible',     listKey:'si_no'},
    {h:'oculto',          listKey:'si_no'},
  ];
  cfCols.forEach(function(cf) {
    var col = hIdx[cf.h];
    var colLetter = String.fromCharCode(64 + col); // 1-based → A=1
    var range = shPerf.getRange(DATA_START, col, DATA_END - DATA_START + 1, 1);
    // Formula: cell not empty AND not found in list
    var listR = listsRange(cf.listKey);
    var formula = '=AND(' + colLetter + DATA_START + '<>"",COUNTIF(' + listR + ',' + colLetter + DATA_START + ')=0)';
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground('#ffcccc')
      .setRanges([range])
      .build();
    var rules = shPerf.getConditionalFormatRules();
    rules.push(rule);
    shPerf.setConditionalFormatRules(rules);
  });

  // ── Share file (anyone with link can view/edit) ───────────────────────────
  var file = DriveApp.getFileById(plId);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  var url = 'https://docs.google.com/spreadsheets/d/' + plId + '/edit';

  return { plantilla_url: url, nombre: PLANTILLA_NAME };
}
