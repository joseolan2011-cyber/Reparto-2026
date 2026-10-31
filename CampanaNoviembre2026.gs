/**
 * CampanaNoviembre2026.gs
 * SuperAgua / ARYS · avisos de visita con video aprobado por Meta.
 * INDEPENDIENTE: no altera ventas, pedidos, cargas ni monitor central.
 *
 * Este archivo NO crea activadores ni envía campañas al instalarse.
 * El envío MASIVO requiere una propiedad adicional de autorización:
 *   CAMPANA_NOVIEMBRE_PERMITIR_MASIVO = ENVIAR_2026
 *
 * META_ACCESS_TOKEN contiene SOLO el token (sin "Bearer").
 */
const CN26 = Object.freeze({
  TZ: 'America/Merida',
  REPARTO_ID: '1sPVwm2CV2b4CwydTLunN_DKyoHnPLeUtfkgvPgxtPNw',
  CLIENTES: 'CLIENTES',
  ENVIOS: 'CAMPANA_NOVIEMBRE_2026',
  LOG_ERRORES: 'LOG_DESPLIEGUES',
  META_API: 'v25.0',
  META_PHONE_ID: '1041808755687636',
  PLANTILLA: 'superagua_visitas_noviembre_2026',
  VIDEO: '1063654766729292',
  IDIOMA: 'es_MX',
  LOTE: 40,
  MIN_INTERVALO_MS: 350,
  MAX_EJECUCION_MS: 245000,
  TELEFONOS_PRUEBA: ['9381771413', '9381850099'],
  COLUMNAS: [
    'Telefono','Cliente','Dias_Visita','Activo_Al_Inicio',
    'Whatsapp','Ultima_Compra','Compras_Historicas','Compras_30_Dias',
    'Vendedor','Estado','Fecha_Intento','HTTP','WAMID',
    'Error_Codigo','Error_Detalle','Intentos'
  ]
});

/** LECTURA SOLAMENTE: comprueba propiedades, tablas y audiencia. */
function CN26_DIAGNOSTICO() {
  const props = PropertiesService.getScriptProperties();
  const sh = CN26_hojaClientes_();
  const data = sh.getDataRange().getDisplayValues();
  const ix = CN26_indices_(data[0], [
    'telefono','nombre','Whatsapp','Dia_Visita','Activo','ES PUBLICO GRAL'
  ]);
  const out = { registros:0, conDias:0, conWhatsapp:0, sinWhatsapp:0,
    excluidosSinDias:0, inactivosConWhatsapp:0, telefonosInvalidos:0 };
  const vistos = {};
  for (let i=1;i<data.length;i++) {
    const r = data[i], tel = CN26_telefono_(r[ix.telefono]);
    if (!r[ix.telefono]) continue;
    out.registros++;
    const dias = String(r[ix.Dia_Visita] || '').trim();
    if (!dias) { out.excluidosSinDias++; continue; }
    out.conDias++;
    if (!tel || vistos[tel]) { out.telefonosInvalidos++; continue; }
    vistos[tel] = true;
    if (!CN26_si_(r[ix.Whatsapp])) { out.sinWhatsapp++; continue; }
    if (CN26_si_(r[ix['ES PUBLICO GRAL']])) continue;
    out.conWhatsapp++;
    if (!CN26_si_(r[ix.Activo])) out.inactivosConWhatsapp++;
  }
  return {
    ok: true,
    credencialPresente: Boolean(props.getProperty('META_ACCESS_TOKEN')),
    phoneId: props.getProperty('META_PHONE_NUMBER_ID') || CN26.META_PHONE_ID,
    plantilla: props.getProperty('META_TEMPLATE_NAME') || CN26.PLANTILLA,
    video: props.getProperty('META_VIDEO_ID') || CN26.VIDEO,
    envioMasivoHabilitado: props.getProperty('CAMPANA_NOVIEMBRE_PERMITIR_MASIVO') === 'ENVIAR_2026',
    publico: out
  };
}

/** GET de Meta: no envía WhatsApp ni escribe datos. */
function CN26_PROBAR_CREDENCIALES() {
  const cfg = CN26_config_();
  const respuesta = UrlFetchApp.fetch(
    'https://graph.facebook.com/'+cfg.api+'/'+encodeURIComponent(cfg.phoneId)+
    '?fields=id,display_phone_number',
    {method:'get', headers:{Authorization:'Bearer '+cfg.token},
      muteHttpExceptions:true}
  );
  const status=respuesta.getResponseCode();
  const body=CN26_json_(respuesta.getContentText());
  const comprobacion = {
    ok:status>=200 && status<300,
    http:status,
    id:body.id || '',
    numero:body.display_phone_number || '',
    error:body.error ? {code:body.error.code,message:body.error.message} : null
  };
  console.log('CN26_CREDENCIALES ' + JSON.stringify(comprobacion));
  return comprobacion;
}

/** ÚNICOS DOS números permitidos para la prueba. Envía dos plantillas. */
function CN26_PROBAR_VIDEO_EN_DOS_TELEFONOS() {
  const cfg=CN26_config_();
  const resultados=[];
  for (const tel of CN26.TELEFONOS_PRUEBA) {
    const r=CN26_enviar_(cfg,tel,'Lunes y Jueves');
    resultados.push({telefono:tel,ok:r.ok,http:r.http,
      wamid:r.wamid,errorCodigo:r.errorCodigo,detalle:r.detalle});
    console.log('CN26_PRUEBA_VIDEO ' + JSON.stringify(resultados[resultados.length - 1]));
    Utilities.sleep(CN26.MIN_INTERVALO_MS);
  }
  return resultados;
}

/**
 * Congela el estado de actividad ANTES de enviar. No manda WhatsApp.
 * Si la hoja existe, no vuelve a crearla ni sobrescribe avances.
 */
function CN26_PREPARAR_AUDIENCIA() {
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Otro proceso está usando el bloqueo. Reintente.');
  try {
    const ss=SpreadsheetApp.openById(CN26.REPARTO_ID);
    const existente=ss.getSheetByName(CN26.ENVIOS);
    if (existente) return {yaExistia:true,filas:Math.max(0,existente.getLastRow()-1)};
    const origen=CN26_hojaClientes_().getDataRange().getDisplayValues();
    const ix=CN26_indices_(origen[0],[
      'telefono','nombre','Whatsapp','Dia_Visita','Activo','ES PUBLICO GRAL',
      'Ultima Compra','T_Veces_Compras','Veces_30_dias','Vendedor_Nombre'
    ]);
    const lista=[], vistos={}, descartes={sinDia:0,sinWhatsapp:0,telefonoInvalido:0,duplicado:0,publico:0};
    for(let i=1;i<origen.length;i++){
      const r=origen[i],raw=r[ix.telefono];
      if(!raw)continue;
      const dias=CN26_formatearDias_(r[ix.Dia_Visita]);
      const tel=CN26_telefono_(raw);
      if(!dias){descartes.sinDia++;continue;}
      if(CN26_si_(r[ix['ES PUBLICO GRAL']])){descartes.publico++;continue;}
      if(!tel){descartes.telefonoInvalido++;continue;}
      if(vistos[tel]){descartes.duplicado++;continue;}
      vistos[tel]=true;
      if(!CN26_si_(r[ix.Whatsapp])){descartes.sinWhatsapp++;continue;}
      lista.push([
        tel,r[ix.nombre],dias,r[ix.Activo],r[ix.Whatsapp],
        r[ix['Ultima Compra']],r[ix.T_Veces_Compras],r[ix.Veces_30_dias],
        r[ix.Vendedor_Nombre],'PENDIENTE','','','','','',0
      ]);
    }
    const sh=ss.insertSheet(CN26.ENVIOS);
    sh.getRange(1,1,1,CN26.COLUMNAS.length).setValues([CN26.COLUMNAS]);
    if(lista.length)sh.getRange(2,1,lista.length,CN26.COLUMNAS.length).setValues(lista);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,CN26.COLUMNAS.length).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@');
    sh.autoResizeColumns(1,4);
    return {creada:true,preparados:lista.length,descartes:descartes,
      inactivos:lista.filter(r=>!CN26_si_(r[3])).length};
  }finally{lock.releaseLock();}
}

/** Muestra el avance sin modificar ni enviar nada. */
function CN26_ESTADO() {
  const sh=SpreadsheetApp.openById(CN26.REPARTO_ID).getSheetByName(CN26.ENVIOS);
  if(!sh)return {preparada:false};
  const rows=sh.getLastRow()>1?
    sh.getRange(2,10,sh.getLastRow()-1,1).getDisplayValues():[];
  const estados={};
  rows.forEach(r=>{const key=r[0]||'VACIO';estados[key]=(estados[key]||0)+1;});
  return {preparada:true,total:rows.length,estados:estados};
}

/**
 * Envia COMO MÁXIMO 40 destinatarios por ejecución, uno por uno.
 * Está bloqueado a menos que exista autorización expresa en propiedades.
 * Los estados EN_PROCESO no se reintentan automáticamente: evita duplicados
 * si Apps Script pierde la respuesta después de que Meta recibió el mensaje.
 */
function CN26_ENVIAR_LOTE() {
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('CAMPANA_NOVIEMBRE_PERMITIR_MASIVO')!=='ENVIAR_2026')
    throw new Error('Envío masivo bloqueado: falta autorización explícita.');
  if(props.getProperty('CAMPANA_NOVIEMBRE_PAUSADA')==='SI')
    return {pausado:true};
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000))return {otraEjecucion:true};
  try{
    const sh=SpreadsheetApp.openById(CN26.REPARTO_ID).getSheetByName(CN26.ENVIOS);
    if(!sh)throw new Error('Primero ejecute CN26_PREPARAR_AUDIENCIA.');
    const cfg=CN26_config_();
    const total=sh.getLastRow();
    if(total<=1)return {procesados:0};
    const data=sh.getRange(2,1,total-1,CN26.COLUMNAS.length).getDisplayValues();
    let realizados=0,aceptados=0,fallidos=0,pausaLimite=false;
    const start=Date.now();
    for(let i=0;i<data.length;i++){
      if(realizados>=CN26.LOTE || Date.now()-start>CN26.MAX_EJECUCION_MS)break;
      const r=data[i];
      if(r[9]!=='PENDIENTE')continue;
      const row=i+2;
      const tel=CN26_telefono_(r[0]);
      if(!tel || !r[2]){
        sh.getRange(row,10).setValue('REVISAR_DATOS');
        continue;
      }
      // Escribimos ANTES de llamar a Meta: si el proceso se interrumpe,
      // el destinatario queda en revisión, nunca se reenvía automáticamente.
      sh.getRange(row,10,1,2).setValues([['EN_PROCESO',new Date()]]);
      SpreadsheetApp.flush();
      let result;
      try{
        result=CN26_enviar_(cfg,tel,r[2]);
      }catch(e){
        result={ok:false,http:0,errorCodigo:'SIN_RESPUESTA',
          detalle:String(e).slice(0,350),wamid:''};
      }
      const estado=result.ok?'ACEPTADO_META':'RECHAZADO';
      sh.getRange(row,10).setValue(estado);
      sh.getRange(row,12,1,5).setValues([[
        result.http,result.wamid,result.errorCodigo,result.detalle,1
      ]]);
      realizados++;
      if(result.ok){aceptados++;}
      else{
        fallidos++;
        CN26_registrarFallo_(tel,r[2],result);
        if(result.http===429 || String(result.errorCodigo)==='130429'){
          pausaLimite=true;props.setProperty('CAMPANA_NOVIEMBRE_PAUSADA','SI');
        }
      }
      if(pausaLimite)break;
      Utilities.sleep(CN26.MIN_INTERVALO_MS);
    }
    // Cuando ya no hay pendientes, quita solamente el trigger de esta campaña.
    const pendientes=sh.getRange(2,10,total-1,1).getDisplayValues()
      .filter(c=>c[0]==='PENDIENTE').length;
    if(pendientes===0){
      ScriptApp.getProjectTriggers().forEach(t=>{
        if(t.getHandlerFunction()==='CN26_ENVIAR_LOTE')ScriptApp.deleteTrigger(t);
      });
    }
    return {procesados:realizados,aceptados:aceptados,
      rechazos:fallidos,pausaLimite:pausaLimite,pendientes:pendientes};
  }finally{lock.releaseLock();}
}

/** Solo instala un trigger si YA existe autorización de envío. */
function CN26_PROGRAMAR_LOTES() {
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('CAMPANA_NOVIEMBRE_PERMITIR_MASIVO')!=='ENVIAR_2026')
    throw new Error('No autorizado: no se programaron activadores.');
  if(props.getProperty('CAMPANA_NOVIEMBRE_PAUSADA')==='SI')
    throw new Error('La campaña está pausada.');
  const handler='CN26_ENVIAR_LOTE';
  if(ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()===handler))
    return {yaProgramado:true};
  ScriptApp.newTrigger(handler).timeBased().everyMinutes(1).create();
  return {programado:true,periodo:'1 minuto',lote:CN26.LOTE};
}

/** Pausa solo esta campaña y retira su trigger. No afecta otros monitores. */
function CN26_DETENER() {
  PropertiesService.getScriptProperties().setProperty('CAMPANA_NOVIEMBRE_PAUSADA','SI');
  let eliminados=0;
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()==='CN26_ENVIAR_LOTE'){
      ScriptApp.deleteTrigger(t);eliminados++;
    }
  });
  return {pausado:true,triggersRetirados:eliminados};
}

function CN26_hojaClientes_() {
  const sh=SpreadsheetApp.openById(CN26.REPARTO_ID).getSheetByName(CN26.CLIENTES);
  if(!sh)throw new Error('No existe la hoja CLIENTES de Reparto.');
  return sh;
}
function CN26_indices_(headers,required){
  const o={};
  required.forEach(name=>{
    const i=headers.findIndex(h=>String(h).trim().toLowerCase()===name.toLowerCase());
    if(i<0)throw new Error('Falta columna en CLIENTES: '+name);
    o[name]=i;
  });
  return o;
}
function CN26_formatearDias_(v){
  const texto=String(v||'').trim();
  if(!texto)return '';
  const partes=texto.split(/\s*,\s*|\s+y\s+/i).map(x=>x.trim()).filter(Boolean);
  return partes.length===2?partes.join(' y '):texto;
}
function CN26_telefono_(v){
  const clean=String(v||'').replace(/\D/g,'');
  const tel=clean.slice(-10);
  return tel.length===10 && /^[2-9]\d{9}$/.test(tel)?tel:'';
}
function CN26_si_(v){
  return ['si','sí','true','1','yes'].includes(String(v||'').trim().toLowerCase());
}
function CN26_json_(txt){
  try{return JSON.parse(txt);}catch(e){return {raw:String(txt).slice(0,250)};}
}
function CN26_config_(){
  const props=PropertiesService.getScriptProperties();
  let token=String(props.getProperty('META_ACCESS_TOKEN')||'').trim();
  token=token.replace(/^Bearer\s+/i,'');
  if(!token)throw new Error('Falta META_ACCESS_TOKEN en Propiedades del proyecto.');
  return {
    token:token,
    api:'v25.0',
    phoneId:props.getProperty('META_PHONE_NUMBER_ID')||CN26.META_PHONE_ID,
    plantilla:props.getProperty('META_TEMPLATE_NAME')||CN26.PLANTILLA,
    video:props.getProperty('META_VIDEO_ID')||CN26.VIDEO,
    idioma:props.getProperty('META_LANGUAGE')||CN26.IDIOMA
  };
}
function CN26_enviar_(cfg,tel,dias){
  const body={
    messaging_product:'whatsapp',recipient_type:'individual',
    to:'52'+tel,type:'template',
    template:{name:cfg.plantilla,language:{code:cfg.idioma},
      components:[
        {type:'header',parameters:[{type:'video',video:{id:cfg.video}}]},
        {type:'body',parameters:[{type:'text',text:CN26_formatearDias_(dias)}]}
      ]
    }
  };
  const response=UrlFetchApp.fetch(
    'https://graph.facebook.com/'+cfg.api+'/'+cfg.phoneId+'/messages',
    {method:'post', contentType:'application/json',
      headers:{Authorization:'Bearer '+cfg.token},
      payload:JSON.stringify(body),muteHttpExceptions:true}
  );
  const http=response.getResponseCode();
  const result=CN26_json_(response.getContentText());
  const ok=http>=200 && http<300 && !!(result.messages && result.messages[0] && result.messages[0].id);
  return {ok:ok,http:http,
    wamid:ok?result.messages[0].id:'',
    errorCodigo:ok?'':String(result.error && result.error.code || http),
    detalle:ok?'':String(result.error && (result.error.error_data && result.error.error_data.details || result.error.message) || 'Respuesta de Meta sin wamid').slice(0,350)};
}
function CN26_registrarFallo_(tel,dias,result){
  const ss=SpreadsheetApp.openById(CN26.REPARTO_ID);
  const sh=ss.getSheetByName(CN26.LOG_ERRORES);
  if(!sh)throw new Error('No existe LOG_DESPLIEGUES para anotar el error.');
  const now=new Date();
  // Solo se rellenan A:D y F,H:I. E y G tienen fórmulas de hoja.
  const id='APPSCRIPT-'+Utilities.getUuid();
  const mensaje=CN26_mensajeManual_(dias);
  sh.appendRow([id,now,Utilities.formatDate(now,CN26.TZ,'HH:mm:ss'),tel]);
  const row=sh.getLastRow();
  sh.getRange(row,6).setValue(result.errorCodigo||'DESCONOCIDO');
  sh.getRange(row,8).setValue(mensaje);
  sh.getRange(row,9).setFormula('=HYPERLINK("https://wa.me/52'+tel+'?text="&ENCODEURL(H'+row+'),"Abrir WhatsApp")');
}
function CN26_mensajeManual_(dias){
  return '💧 *¡En SuperAgua comenzamos una nueva etapa, pensando en usted!*\n\n'+
  'Queremos que nunca le falte agua en casa. Por eso, a partir de noviembre, nuestro repartidor lo visitará *2 veces por semana*, en días programados.\n\n'+
  '📅 *SUS DÍAS DE VISITA SON:*\n👉 *'+dias+'*\n\n'+
  '🎁 *¡Y tenemos un regalo para usted!*\n\n'+
  'Durante todo noviembre, por cada garrafón de 20 litros que compre en sus días asignados, recibirá *1 bolsa de agua ARYS GRATIS*, para que disfrute de hielo de calidad en casa. 🧊\n\n'+
  '🎬 Preparamos este video para contarle cómo estamos mejorando nuestro servicio. ¡Le invitamos a verlo!\n\n'+
  '💙 *Gracias por su confianza y por acompañarnos en esta nueva etapa.*\n\n'+
  '_Puede seguir comprando cualquier otro día, pero el obsequio solo aplica en sus días de visita. Promoción exclusiva para clientes Hogar con precio regular. No aplica con precios preferenciales, descuentos ni otras promociones._';
}
