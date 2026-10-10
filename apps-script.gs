// ============================================================
//  SMART FIT CUSCO · Registro en Google Sheets (v5)
//  Pega TODO esto en: tu Google Sheet -> Extensiones -> Apps Script
//  1) Ejecuta una vez la función  configurar()  (pide permisos)
//  2) Implementar -> Nueva implementación -> Aplicación web
//     (Ejecutar como: Yo · Acceso: Cualquier persona) -> copia la URL /exec
// ============================================================
const PESTANAS = {
  Pedidos:     ['fecha','pedido_id','estado','canal','cliente','telefono','items','cantidad','total','sesion','notas'],
  Tallaje:     ['fecha','sesion','paso','detalle','prenda','talla','color','silueta','altura','pecho','cintura','cadera','segundos'],
  Opiniones:   ['fecha','sesion','momento','prenda','talla','estrellas','gusto','comentario'],
  Encuesta:    ['fecha','sesion','estrellas','gusto','uso','util','conf','usar','pago','nps','comentario','silueta','altura','pecho','cintura','cadera'],
  Contactos:   ['fecha','cliente','correo','mensaje','estado'],
  Suscriptores:['fecha','correo'],
  Cambios:     ['fecha','sesion','prenda','talla','color','detalle','comentario','estado'],
  Eventos:     ['fecha','sesion','evento','prenda','talla','segundos']
};
const CAT_HEAD = ['id','categoria','nombre','estilo','precio','precio_anterior','etiqueta','siluetas','colores','tela','cuidados','agotadas'];
const CAT_DATOS = [["id", "categoria", "nombre", "estilo", "precio", "precio_anterior", "etiqueta", "siluetas", "colores", "tela", "cuidados", "agotadas"], ["v1", "vestido", "Vestido envolvente", "Elegante", "139", "", "Nuevo", "reloj;pera;manzana", "Terracota;Verde;Negro;Azul", "Crepé elástico", "Lavar en frío, no usar secadora", ""], ["v2", "vestido", "Vestido línea A", "Casual", "149", "189", "-20%", "manzana;inv;rect", "Azul;Mostaza;Negro;Verde", "Crepé elástico", "Lavar en frío, no usar secadora", "L;5XL"], ["v3", "vestido", "Vestido midi fluido", "Elegante", "169", "", "", "pera;reloj;rect", "Verde;Terracota;Negro;Crema", "Crepé elástico", "Lavar en frío, no usar secadora", ""], ["b1", "top", "Blusa fluida", "Moderno", "79", "", "Nuevo", "pera;rect;reloj", "Crema;Terracota;Verde;Azul", "Algodón con elastano", "Lavar a máquina en frío", ""], ["b2", "top", "Polo básico", "Casual", "59", "", "", "reloj;pera;manzana;rect;inv", "Negro;Crema;Azul;Mostaza", "Algodón con elastano", "Lavar a máquina en frío", "5XL"], ["b3", "top", "Camisa oversize", "Clásico", "99", "129", "-20%", "manzana;inv;rect", "Crema;Azul;Negro;Verde", "Algodón con elastano", "Lavar a máquina en frío", ""], ["b4", "top", "Top cruzado", "Moderno", "89", "", "Nuevo", "reloj;pera;manzana", "Verde;Negro;Terracota;Crema", "Algodón con elastano", "Lavar a máquina en frío", ""], ["b5", "top", "Blusa manga larga", "Casual", "85", "", "", "pera;rect;inv", "Mostaza;Crema;Negro;Verde", "Algodón con elastano", "Lavar a máquina en frío", ""], ["b6", "top", "Suéter tejido", "Clásico", "129", "159", "-20%", "reloj;rect;manzana", "Terracota;Verde;Crema;Negro", "Tejido suave", "Lavar a mano", ""], ["c1", "casaca", "Casaca larga", "Urbano", "219", "", "Nuevo", "manzana;reloj;rect", "Negro;Azul;Terracota;Verde", "Mezcla de poliéster", "Lavado en seco o a mano", ""], ["c2", "casaca", "Casaca estructurada", "Clásico", "199", "", "", "pera;rect", "Negro;Crema;Mostaza;Azul", "Mezcla de poliéster", "Lavado en seco o a mano", ""], ["c3", "casaca", "Blazer largo", "Elegante", "239", "299", "-20%", "pera;reloj;manzana", "Negro;Azul;Crema;Terracota", "Mezcla de poliéster", "Lavado en seco o a mano", "2XL"], ["c4", "casaca", "Cárdigan largo", "Casual", "179", "", "Nuevo", "manzana;pera;rect", "Crema;Negro;Azul;Terracota", "Tejido suave", "Lavar a mano", ""], ["c5", "casaca", "Bomber urbana", "Urbano", "189", "", "", "inv;rect;reloj", "Verde;Negro;Azul;Mostaza", "Mezcla de poliéster", "Lavado en seco o a mano", ""], ["p1", "pantalon", "Pantalón tiro alto", "Moderno", "159", "", "", "reloj;manzana;rect", "Negro;Azul;Mostaza;Verde", "Algodón stretch", "Lavar a máquina, planchar a temperatura media", "3XL"], ["p2", "pantalon", "Pantalón recto", "Clásico", "129", "", "", "inv;pera;reloj;rect", "Negro;Crema;Azul;Terracota", "Algodón stretch", "Lavar a máquina, planchar a temperatura media", ""], ["p3", "pantalon", "Pantalón fluido", "Casual", "139", "", "Nuevo", "pera;manzana;rect", "Crema;Verde;Negro;Mostaza", "Algodón stretch", "Lavar a máquina, planchar a temperatura media", ""], ["p4", "pantalon", "Pantalón cargo", "Urbano", "149", "", "", "pera;manzana;rect", "Crema;Verde;Negro;Azul", "Algodón stretch", "Lavar a máquina, planchar a temperatura media", ""], ["p5", "pantalon", "Pantalón sastre", "Elegante", "169", "", "Nuevo", "reloj;pera;rect", "Negro;Azul;Crema;Mostaza", "Algodón stretch", "Lavar a máquina, planchar a temperatura media", ""]];
const ESTADOS_PEDIDO = ['Nuevo','Confirmado','Pagado','Listo para recoger','Entregado','Cancelado'];

function hoja_(nombre){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

// ---------- Se ejecuta UNA vez a mano ----------
function configurar(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(PESTANAS).forEach(function(n){
    const sh = hoja_(n), h = PESTANAS[n];
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,h.length).setValues([h]);
    sh.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(2,1,sh.getMaxRows()-1,1).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.autoResizeColumns(1,h.length);
  });
  // Lista desplegable de estado en Pedidos
  const p = hoja_('Pedidos');
  p.getRange(2,3,1000,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_PEDIDO,true).setAllowInvalid(true).build());
  p.getRange(2,9,1000,1).setNumberFormat('"S/ "#,##0.00');
  const reglas = [['Nuevo','#fff2cc'],['Confirmado','#cfe2f3'],['Pagado','#d9ead3'],['Listo para recoger','#d9d2e9'],['Entregado','#b6d7a8'],['Cancelado','#f4cccc']]
    .map(function(r){ return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r[0]).setBackground(r[1]).setRanges([p.getRange('C2:C1000')]).build(); });
  p.setConditionalFormatRules(reglas);
  hoja_('Contactos').getRange(2,5,1000,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pendiente','Respondido'],true).build());

  // Catálogo (editable): se carga solo si está vacío
  const c = hoja_('Catalogo');
  if (c.getLastRow() === 0) c.getRange(1,1,CAT_DATOS.length,CAT_HEAD.length).setValues(CAT_DATOS);
  c.getRange(1,1,1,CAT_HEAD.length).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  c.setFrozenRows(1);
  c.getRange(1,1,c.getMaxRows(),CAT_HEAD.length).setNumberFormat('@'); // todo como texto para no perder "L;5XL" ni ceros

  resumen_();
  const hojaInicial = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hojaInicial && ss.getSheets().length > 1 && hojaInicial.getLastRow() === 0) ss.deleteSheet(hojaInicial);
  ss.setActiveSheet(ss.getSheetByName('Resumen'));
  SpreadsheetApp.getUi().alert('Listo. Ahora: Implementar -> Nueva implementación -> Aplicación web.');
}

function resumen_(){
  const sh = hoja_('Resumen'); sh.clear();
  const f = [
    ['RESUMEN SMART FIT CUSCO',''],
    ['Pedidos totales',"=MAX(0,COUNTA(Pedidos!B:B)-1)"],
    ['Pedidos nuevos (sin atender)','=COUNTIF(Pedidos!C:C,"Nuevo")'],
    ['Ventas confirmadas (S/)','=SUM(Pedidos!I2:I)-SUMIF(Pedidos!C2:C,"Nuevo",Pedidos!I2:I)-SUMIF(Pedidos!C2:C,"Cancelado",Pedidos!I2:I)'],
    ['Pedidos por WhatsApp','=COUNTIF(Pedidos!D:D,"WhatsApp")'],
    ['Reservas en tienda','=COUNTIF(Pedidos!D:D,"Reserva en tienda")'],
    ['Tallajes realizados','=COUNTIF(Tallaje!C:C,"medidas_analizadas")'],
    ['Opiniones (tras tallaje/compra)','=MAX(0,COUNTA(Opiniones!A:A)-1)'],
    ['Estrellas promedio (opiniones)','=IFERROR(AVERAGE(Opiniones!F2:F),"–")'],
    ['Encuestas completas','=MAX(0,COUNTA(Encuesta!A:A)-1)'],
    ['Recomendación promedio 0–10 (NPS)','=IFERROR(AVERAGE(Encuesta!J2:J),"–")'],
    ['Mensajes de contacto pendientes','=COUNTIF(Contactos!E:E,"Pendiente")'],
    ['Suscriptores','=MAX(0,COUNTA(Suscriptores!B:B)-1)'],
    ['Solicitudes de cambio de talla','=MAX(0,COUNTA(Cambios!A:A)-1)']
  ];
  f.forEach(function(r,i){ sh.getRange(i+1,1).setValue(r[0]); if (r[1]) sh.getRange(i+1,2).setFormula(r[1]); });
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2:A14').setFontWeight('bold');
  sh.setColumnWidth(1,300); sh.setColumnWidth(2,140);
}

// ---------- Recibe los datos del sitio ----------
function doPost(e){
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    const f = d.fecha ? new Date(d.fecha) : new Date();
    const ev = String(d.detalle||'').split(':')[0];
    const resto = String(d.detalle||'').indexOf(':')>-1 ? String(d.detalle).slice(String(d.detalle).indexOf(':')+1).trim() : '';
    let destino, fila;
    switch (d.tipo) {
      case 'pedido':
        destino='Pedidos'; fila={fecha:f,pedido_id:d.pedido_id,estado:d.estado||'Nuevo',canal:d.canal,cliente:d.cliente,telefono:d.telefono,items:d.items,cantidad:d.cantidad,total:d.total,sesion:d.sesion,notas:''}; break;
      case 'feedback_talla': case 'feedback_compra':
        destino='Opiniones'; fila={fecha:f,sesion:d.sesion,momento:d.tipo==='feedback_talla'?'Tras el tallaje':'Tras la compra',prenda:d.prenda,talla:d.talla,estrellas:d.estrellas,gusto:d.gusto,comentario:d.comentario}; break;
      case 'encuesta':
        destino='Encuesta'; fila=Object.assign({},d,{fecha:f}); break;
      case 'contacto':
        destino='Contactos'; fila={fecha:f,cliente:d.cliente||String(d.detalle||'').split(' · ')[0],correo:d.correo||String(d.detalle||'').split(' · ')[1]||'',mensaje:d.comentario,estado:'Pendiente'}; break;
      case 'suscripcion':
        destino='Suscriptores'; fila={fecha:f,correo:d.correo||d.detalle}; break;
      case 'cambio_talla':
        destino='Cambios'; fila={fecha:f,sesion:d.sesion,prenda:d.prenda,talla:d.talla,color:d.color,detalle:d.detalle,comentario:d.comentario,estado:'Pendiente'}; break;
      default: // 'evento'
        if (['medidas_analizadas','prenda_probada','talla_lista'].indexOf(ev) > -1) {
          destino='Tallaje'; fila=Object.assign({},d,{fecha:f,paso:ev,detalle:resto});
        } else {
          destino='Eventos'; fila={fecha:f,sesion:d.sesion,evento:ev,prenda:resto||d.prenda,talla:d.talla,segundos:d.segundos};
        }
    }
    const sh = hoja_(destino), head = PESTANAS[destino];
    if (sh.getLastRow() === 0) sh.appendRow(head);
    sh.appendRow(head.map(function(k){ return fila[k]===undefined?'':fila[k]; }));
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: '+err);
  } finally { try { lock.releaseLock(); } catch(x){} }
}

// ---------- GET: estado, o catálogo en CSV para la tienda ----------
function doGet(e){
  if (e && e.parameter && e.parameter.catalogo) {
    const v = hoja_('Catalogo').getDataRange().getValues();
    const csv = v.map(function(r){ return r.map(function(c){ return '"'+String(c).replace(/"/g,'""')+'"'; }).join(','); }).join('\n');
    return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
  }
  return ContentService.createTextOutput('Smart Fit Cusco: activo');
}
