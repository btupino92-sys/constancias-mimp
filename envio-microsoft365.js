/* Paso 5: envio delegado desde Microsoft 365 / Microsoft Graph. Sin secretos en el navegador. */
(()=>{'use strict';
const el=id=>document.getElementById(id);
let client=null,account=null,prepared=[],sending=false;
const mode=el('sendMode');const graphSettings=el('graphSettings');
function setMode(){const graph=mode.value==='m365';graphSettings.hidden=!graph;el('classicInfo').hidden=graph;}
mode.addEventListener('change',setMode);setMode();
const scopes=['User.Read','Mail.Send','Mail.Send.Shared'];
const clean=s=>String(s||'').trim();
function log(msg){el('msSendStatus').textContent=msg;}
function readyMatches(){
 const {matches}=autoMatch();
 return participants.map((p,i)=>({p,f:matches.get(i),i})).filter(x=>x.f&&validLink(x.p.link)&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean(x.p.email)));
}
function updateDisplay(){
 const parent=el('msRecipients'); parent.replaceChildren();
 if(!prepared.length){parent.textContent='No hay registros listos. Revisa los PDF, los enlaces y los correos del paso 2.';return;}
 const table=document.createElement('table');table.innerHTML='<thead><tr><th>Enviar</th><th>Participante</th><th>Correo</th><th>Estado</th></tr></thead>';
 const tbody=document.createElement('tbody');
 for(let i=0;i<prepared.length;i++){
  const row=prepared[i],tr=document.createElement('tr');
  const a=document.createElement('td'),cb=document.createElement('input');cb.type='checkbox';cb.checked=row.selected;cb.disabled=row.status==='ENVIADO';cb.addEventListener('change',()=>row.selected=cb.checked);a.append(cb);
  const b=document.createElement('td');b.textContent=row.p.name;
  const c=document.createElement('td');c.textContent=row.p.email;
  const d=document.createElement('td');d.textContent=row.status||'LISTO';
  tr.append(a,b,c,d);tbody.append(tr);
 }table.append(tbody);parent.append(table);
}
el('msCheck').addEventListener('click',()=>{
 const previous=new Map(prepared.map(x=>[x.i,x]));
 prepared=readyMatches().map(x=>({...x,selected:previous.get(x.i)?.status!=='ENVIADO',status:previous.get(x.i)?.status==='ENVIADO'?'ENVIADO':'LISTO'}));
 updateDisplay();log(`${prepared.length} constancias listas para correo; ${participants.length-prepared.length} registros pendientes (sin correo, PDF o enlace).`);
});
async function connect(){
 if(!window.msal)throw Error('No se cargó Microsoft Authentication Library (MSAL). Comprueba la conexión a Internet o las restricciones de red.');
 const clientId=clean(el('msClient').value),tenantId=clean(el('msTenant').value);
 if(!/^[0-9a-f-]{36}$/i.test(clientId)||!/^[0-9a-f-]{36}$/i.test(tenantId))throw Error('OGTI debe proporcionar el ID de aplicación y el ID de directorio válidos.');
 client=new msal.PublicClientApplication({auth:{clientId,authority:`https://login.microsoftonline.com/${tenantId}`,redirectUri:location.origin+location.pathname},cache:{cacheLocation:'sessionStorage'}});
 await client.handleRedirectPromise();
 const result=await client.loginPopup({scopes,prompt:'select_account'});account=result.account;
 el('msLoginStatus').textContent=`Conectado como ${account.username}.`;return true;
}
el('msLogin').addEventListener('click',()=>connect().catch(e=>el('msLoginStatus').textContent='No se pudo conectar: '+e.message));
async function token(){if(!client||!account)throw Error('Conecta primero tu cuenta Microsoft 365.');
 try{return (await client.acquireTokenSilent({scopes,account})).accessToken;}
 catch(e){return (await client.acquireTokenPopup({scopes,account})).accessToken;}}
function toB64(data){let out='',u=data instanceof Uint8Array?data:new Uint8Array(data);for(let i=0;i<u.length;i+=8192)out+=String.fromCharCode(...u.subarray(i,i+8192));return btoa(out);}
function emailBody(p){const v=clean(el('msBody').value);return `${v}\n\nParticipante: ${p.name}\nEnlace de verificación: ${p.link}`;}
async function sendOne(row){
 const pdf=await qrStamped(row.p,row.f);
 if(pdf.byteLength>2.8*1024*1024)throw Error('El PDF supera el límite de adjunto simple (aprox. 3 MB). Requiere envío mediante sesión de carga o reducir tamaño.');
 const from=clean(el('msFrom').value).toLowerCase(),signed=clean(account.username).toLowerCase();
 if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(from))throw Error('El correo remitente no es válido.');
 const endpoint=from===signed?'me':`users/${encodeURIComponent(from)}`;
 const mail={message:{subject:clean(el('msSubject').value)||'Constancia de participación – MIMP',body:{contentType:'Text',content:emailBody(row.p)},toRecipients:[{emailAddress:{address:row.p.email}}],attachments:[{'@odata.type':'#microsoft.graph.fileAttachment',name:safeFile(row.p.name)+'.pdf',contentType:'application/pdf',contentBytes:toB64(pdf)}]},saveToSentItems:true};
 const res=await fetch(`https://graph.microsoft.com/v1.0/${endpoint}/sendMail`,{method:'POST',headers:{Authorization:'Bearer '+await token(),'Content-Type':'application/json'},body:JSON.stringify(mail)});
 if(!res.ok){let msg=(await res.text()).slice(0,500);throw Error(`Microsoft Graph ${res.status}: ${msg}`);}
}
el('msExport').addEventListener('click',()=>{
 if(!prepared.length){log('Primero prepara la lista para descargar un reporte.');return;}
 const rows=[['Nombre completo','Correo','Enlace Drive','Estado'],...prepared.map(r=>[r.p.name,r.p.email,r.p.link,r.status])];
 const csv='\ufeff'+rows.map(a=>a.map(v=>'"'+String(v??'').replaceAll('"','""')+'"').join(';')).join('\r\n');
 downloadBlob(new Blob([csv],{type:'text/csv;charset=utf-8'}),'REPORTE_ENVIO_MIMP.csv');
});
el('msSend').addEventListener('click',async()=>{
 if(sending)return;
 if(mode.value!=='m365'){log('Para Outlook clásico descarga el paquete de borradores.');return;}
 const targets=prepared.filter(r=>r.selected&&r.status!=='ENVIADO');
 if(!targets.length){log('No hay correos seleccionados. Primero pulsa Preparar lista de envíos.');return;}
 if(!client||!account){log('Primero conecta tu cuenta Microsoft 365.');return;}
 const from=clean(el('msFrom').value);
 if(!confirm(`¿CONFIRMAS EL ENVÍO REAL de ${targets.length} correo(s) desde ${from}? Cada uno llevará su PDF final con QR adjunto. Esta acción no se puede deshacer.`))return;
 sending=true;el('msSend').disabled=true;let sent=0,errors=0;
 try{
  for(let j=0;j<targets.length;j++){
   const row=targets[j];log(`Enviando ${j+1} de ${targets.length}: ${row.p.name}...`);
   try{await sendOne(row);row.status='ENVIADO';row.selected=false;sent++;}
   catch(err){row.status='REVISAR ANTES DE REENVIAR: '+err.message;row.selected=false;errors++;}
   updateDisplay();await new Promise(r=>setTimeout(r,100));
  }
  log(`Proceso terminado. Aceptados por Microsoft 365: ${sent}. Errores: ${errors}. La aceptación no garantiza la entrega en la bandeja del destinatario.`);
 }finally{sending=false;el('msSend').disabled=false;}
});
})();
