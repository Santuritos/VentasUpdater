/**
 * =============================================================
 * Actualizador de Ventas – Versión 1.6
 * (Reconstrucción total – Precio desde Datos_ISBN + Woocommerce)
 * -------------------------------------------------------------
 * Fecha:             16 abr 2025
 * Autor:             TuNombre o ChatGPT
 * Descripción:       Sincroniza datos entre Ventas, Datos_ISBN,
 *                    Hamelyn, Momox, Ammareal, Ebay, Todocoleccion
 *                    y Woocommerce. Actualiza Posición en todas las hojas
 *                    y Precio en proveedores + Woocommerce desde Datos_ISBN.
 *                    Además crea enlaces internos en columna ID.
 * =============================================================
 */

function normalizeKey(key) {
  return String(key).trim().toUpperCase();
}

function getHeaderIndex(headers, colName) {
  var idx = headers.indexOf(colName);
  if (idx < 0 && colName.toLowerCase()==="posicion")
    idx = headers.indexOf("Posición");
  return idx;
}

function getPriceFromDatos(normKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datos = ss.getSheetByName("Datos_ISBN");
  if (!datos) return;
  var data = datos.getDataRange().getValues();
  if (data.length<2) return;
  var h = data[0], iISBN = h.indexOf("ISBN"), iPrecio = h.indexOf("Precio");
  if (iISBN<0||iPrecio<0) return;
  for(var i=1;i<data.length;i++){
    if (normalizeKey(data[i][iISBN])===normKey) return data[i][iPrecio];
  }
}

function onEdit(e){
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("IS_SYNCING")==="true") return;
  props.setProperty("IS_SYNCING","true");
  try{
    var s=e.range.getSheet(), n=s.getName();
    var allowed=["Ventas","Datos_ISBN","Hamelyn","Momox","Ammareal","Ebay","Todocoleccion","Woocommerce"];
    if (allowed.indexOf(n)<0) return;
    var headers=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
    var keyField=(n.match(/Hamelyn|Momox|Ammareal/)?"ISBN13":"ISBN");
    var iKey=headers.indexOf(keyField), iPos=getHeaderIndex(headers,'Posicion');
    if(iKey<0||iPos<0)return;
    var row=s.getRange(e.range.getRow(),1,1,headers.length).getValues()[0];
    var key=row[iKey]; if(!key) return;
    var normKey=normalizeKey(key), pos=row[iPos];
    var price = getPriceFromDatos(normKey);
    if(n==="Ventas"){
      updateCell("Datos_ISBN","Posicion",normKey,pos);
      ["Hamelyn","Momox","Ammareal"].forEach(function(p){
        updateCell(p,"Posición",normKey,pos);
        if(price!==undefined) updateCell(p,"Precio",normKey,price);
      });
    } else if(n==="Datos_ISBN"){
      updateCell("Ventas","Posicion",normKey,pos);
      ["Hamelyn","Momox","Ammareal","Ebay","Todocoleccion"].forEach(function(p){
        updateCell(p,p.match(/Hamelyn|Momox|Ammareal/)?"Posición":"Posicion",normKey,pos);
      });
      // Woocommerce: actualizar precio desde Datos_ISBN
      if(price!==undefined) updateCell("Woocommerce","Precio",normKey,price);
    } else if(n.match(/Hamelyn|Momox|Ammareal/)){
      updateCell("Ventas","Posicion",normKey,pos);
      if(price!==undefined){
        var col="Precio "+n;
        updateCell("Ventas",col,normKey,price);
      }
      updateCell("Datos_ISBN","Posicion",normKey,pos);
      ["Hamelyn","Momox","Ammareal"].filter(function(p){return p!==n;}).forEach(function(p){
        updateCell(p,"Posición",normKey,pos);
      });
    } else if(n==="Ebay"||n==="Todocoleccion"){
      updateCell("Datos_ISBN","Posicion",normKey,pos);
      updateCell("Ventas","Posicion",normKey,pos);
    } else if(n==="Woocommerce"){
      // Si editas Woocommerce, tomar precio y propagar a Datos_ISBN y Ventas?
      // Aquí solo actualizamos precio desde Datos_ISBN
      if(price!==undefined) updateCell("Woocommerce","Precio",normKey,price);
    }
  }catch(err){Logger.log(err)}
  finally{props.deleteProperty("IS_SYNCING");}
}

function updateCell(sheetName,colName,normKey,newVal){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(sheetName);
  if(!sh)return; var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var keyField=(sheetName.match(/Hamelyn|Momox|Ammareal/)?"ISBN13":"ISBN");
  if(sheetName==='Woocommerce') keyField='Título';
  var iKey=h.indexOf(keyField), iCol=h.indexOf(colName);
  if(iCol<0&&colName.toLowerCase()==='posicion')iCol=h.indexOf('Posición');
  if(iKey<0||iCol<0)return;
  var data=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  for(var i=0;i<data.length;i++){
    var ck=data[i][iKey]; if(!ck) continue;
    if(normalizeKey(ck)===normKey&&newVal!==undefined&&newVal!==''){
      sh.getRange(i+2,iCol+1).setValue(newVal);
    }
  }
}

function syncAllVentasOptimized(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), vs=ss.getSheetByName('Ventas');
  var vd=vs.getDataRange().getValues(); if(vd.length<2)return;
  var h=vd[0], iKey=h.indexOf('ISBN'), iPos=getHeaderIndex(h,'Posicion');
  if(iKey<0||iPos<0)return;
  var dict={};
  for(var r=1;r<vd.length;r++){var key=vd[r][iKey]; if(key) dict[normalizeKey(key)]={pos:vd[r][iPos],price:getPriceFromDatos(normalizeKey(key))};}
  var targets=[
    {nm:'Datos_ISBN',key:'ISBN',upd:{pos:'Posicion'}},
    {nm:'Hamelyn',key:'ISBN13',upd:{pos:'Posición',price:'Precio'}},
    {nm:'Momox',key:'ISBN13',upd:{pos:'Posición',price:'Precio'}},
    {nm:'Ammareal',key:'ISBN13',upd:{pos:'Posición',price:'Precio'}},
    {nm:'Ebay',key:'ISBN',upd:{pos:'Posicion'}},
    {nm:'Todocoleccion',key:'ISBN',upd:{pos:'Posicion'}},
    {nm:'Woocommerce',key:'Título',upd:{price:'Precio'}}
  ];
  targets.forEach(function(t){
    var sh=ss.getSheetByName(t.nm); if(!sh)return;
    var data=sh.getDataRange().getValues(); if(data.length<2)return;
    var hh=data[0], iK=hh.indexOf(t.key), iP=getHeaderIndex(hh,t.upd.pos), iPr=t.upd.price?hh.indexOf(t.upd.price):-1;
    if(iK<0||iP<0)return;
    for(var i=1;i<data.length;i++){var row=data[i], k=row[iK]; if(!k)continue;var d= dict[normalizeKey(k)]; if(d){if(d.pos!==undefined) row[iP]=d.pos; if(iPr>=0&&d.price!==undefined) row[iPr]=d.price;}};
    sh.getRange(1,1,data.length,data[0].length).setValues(data);
  });
}

function getDatosISBNRow(normKey){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName('Datos_ISBN');
  if(!sh)return; var d=sh.getDataRange().getValues(); if(d.length<2)return;
  var h=d[0], i=h.indexOf('ISBN'); if(i<0)return;
  for(var r=1;r<d.length;r++){ if(normalizeKey(d[r][i])===normKey) return r+1;} }

function updateIDLinksForSheet(name,ssId,gid){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(name);
  if(!sh)return; var d=sh.getDataRange().getValues(); if(d.length<2)return;
  var keyField=(name.match(/Hamelyn|Momox|Ammareal/)?'ISBN13':(name==='Woocommerce'?'Título':'ISBN'));
  var h=d[0], iK=h.indexOf(keyField), iID=h.indexOf('ID'); if(iID<0)return;
  for(var r=1;r<d.length;r++){var k=d[r][iK]; if(!k)continue;var f=getDatosISBNRow(normalizeKey(k)); if(!f)continue;
    var url='https://docs.google.com/spreadsheets/d/'+ssId+'/edit#gid='+gid+'&range=A'+f;
    var cell=sh.getRange(r+1,iID+1),txt=cell.getValue()||k;
    cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(txt).setLinkUrl(url).build());
  }
}

function updateIDLinksAcrossSheets(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),id=ss.getId(),gid=ss.getSheetByName('Datos_ISBN').getSheetId();
  ['Ventas','Datos_ISBN','Hamelyn','Momox','Ammareal','Ebay','Todocoleccion','Woocommerce']
    .forEach(function(n){updateIDLinksForSheet(n,id,gid);});
  Logger.log('Links actualizados');
}
