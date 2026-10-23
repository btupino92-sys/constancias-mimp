 'use strict';
// Archivos PDF o ZIP; asociación automática mediante nombres normalizados y tokens.
let signedFiles=[],unmatchedPDFs=[];
const baseName=s=>String(s||'').replace(/\.pdf$/i,'').replace(/\s*\(\d+\)$/,'').replace(/[_\-.,]+/g,' ').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\s+/g,' ').trim();
const keyName=s=>baseName(s).replace(/[^a-z0-9]/g,'');
const tokenName=s=>baseName(s).split(' ').filter(Boolean).sort().join('|');
const validLink=s=>{try{const u=new URL(s);return u.protocol==='https:';}catch{return false}};
function autoMatch(){
 const matches=new Map(),fileUsed=new Set(),ambiguous=[];
 participants.forEach((p,i)=>{
  const exact=signedFiles.map((f,j)=>({f,j})).filter(({f})=>keyName(f.name)===keyName(p.name));
  let hits=exact.length?exact:signedFiles.map((f,j)=>({f,j})).filter(({f})=>tokenName(f.name)===tokenName(p.name));
  hits=hits.filter(({j})=>!fileUsed.has(j));
  if(hits.length===1){matches.set(i,hits[0].f);fileUsed.add(hits[0].j)}
  else if(hits.length>1)ambiguous.push(p.name);
 });
 unmatchedPDFs=signedFiles.filter((f,j)=>!fileUsed.has(j));
 return {matches,ambiguous};
}
function renderMatches(){
 const {matches,ambiguous}=autoMatch();const linked=participants.filter(p=>validLink(p.link)).length;
 const ready=[...matches.keys()].filter(i=>validLink(participants[i].link)).length;
 $('signedCount').textContent=`${matches.size} / ${participants.length} PDF asociados automáticamente (${signedFiles.length} cargados)`;
 $('linkCount').textContent=`${linked} / ${participants.length} enlaces de Drive válidos`;
 $('readyCount').textContent=`${ready} de ${participants.length} diplomas listos para QR`;
 const box=$('matchRows');box.replaceChildren();
 const problem=participants.map((p,i)=>({p,i})).filter(({p,i})=>!matches.has(i)||!validLink(p.link));
 if(problem.length||unmatchedPDFs.length||ambiguous.length){
  const d=document.createElement('div');d.className='warning';
  const heading=document.createElement('strong');heading.textContent='Revisión de registros pendientes';d.append(heading);
  const ul=document.createElement('ul');
  problem.slice(0,30).forEach(({p,i})=>{const li=document.createElement('li');li.textContent=`${p.name}: ${!matches.has(i)?'sin PDF coincidente':''}${!matches.has(i)&&!validLink(p.link)?' · ':''}${!validLink(p.link)?'sin enlace https válido':''}`;ul.append(li)});
  unmatchedPDFs.slice(0,15).forEach(f=>{const li=document.createElement('li');li.textContent='PDF sin participante: '+f.name;ul.append(li)});
  d.append(ul);box.append(d);
 } else if(participants.length){const p=document.createElement('p');p.className='success-note';p.textContent='✓ Todos los nombres, PDF y enlaces están relacionados automáticamente. No es necesario seleccionar archivos.';box.append(p)}
 return ready;
}
$('firmados').addEventListener('change',async e=>{
 const files=[...e.target.files];signedFiles=[];$('qrProgress').textContent='Leyendo PDF y ZIP...';
 try{
  for(const f of files){
   if(/\.pdf$/i.test(f.name)){signedFiles.push(f);continue}
   if(/\.zip$/i.test(f.name)){
    needLibraries('JSZip');const zip=await JSZip.loadAsync(f);
    for(const [path,item] of Object.entries(zip.files)){
     if(item.dir||! /\.pdf$/i.test(path)||path.startsWith('__MACOSX/'))continue;
     const blob=await item.async('blob');
     signedFiles.push(new File([blob],path.split('/').pop(),{type:'application/pdf'}));
    }
   }
  }
  $('qrProgress').textContent=signedFiles.length+' PDF firmados cargados correctamente.';renderMatches();
 }catch(err){$('qrProgress').textContent='Error al leer PDF o ZIP: '+err.message}
 e.target.value='';
});
$('revisar').addEventListener('click',()=>{renderMatches();$('qrProgress').textContent='Asociación automática actualizada.'});
$('archivoEnlaces').addEventListener('change',()=>setTimeout(renderMatches,500));
$('archivo').addEventListener('change',()=>setTimeout(renderMatches,500));
$('agregar').addEventListener('click',()=>setTimeout(renderMatches,0));
$('limpiar').addEventListener('click',()=>setTimeout(renderMatches,0));
async function qrStamped(p,f){
 needLibraries('PDFLib','QRCode');const {PDFDocument}=PDFLib;
 const doc=await PDFDocument.load(await f.arrayBuffer(),{ignoreEncryption:true});
 const page=doc.getPage(0),{width,height}=page.getSize();
 const url=await QRCode.toDataURL(p.link,{width:800,margin:2});
 const bytes=Uint8Array.from(atob(url.split(',')[1]),c=>c.charCodeAt(0));const qr=await doc.embedPng(bytes);
 // Superior/inferior fijado por proporción de página, QR en esquina inferior izquierda.
 const side=Math.min(width*0.128,height*0.24);
 page.drawImage(qr,{x:width*0.055,y:height*0.042,width:side,height:side});
 return doc.save();
}
async function generatePackage(){
 needLibraries('XLSX','JSZip','PDFLib','QRCode');
 if(!participants.length)throw Error('Primero importa participantes (paso 2).');
 const {matches}=autoMatch();renderMatches();
 const ready=participants.map((p,i)=>({p,f:matches.get(i)})).filter(x=>x.f&&validLink(x.p.link));
 if(!ready.length)throw Error('No hay archivos listos. Comprueba nombres de PDF, participantes y enlaces https.');
 if(ready.length<participants.length&&!confirm(`Hay ${participants.length-ready.length} pendientes. ¿Descargar los ${ready.length} diplomas listos?`))return;
 const zip=new JSZip(),out=zip.folder('DIPLOMAS_FINALES_QR'),originals=zip.folder('ORIGINALES_FIRMADOS');const rows=[],used=new Set();
 for(let i=0;i<ready.length;i++){
  const {p,f}=ready[i];let base=safeFile(p.name),filename=base+'.pdf',n=2;
  while(used.has(filename.toLowerCase()))filename=base+'_'+n+++'.pdf';used.add(filename.toLowerCase());
  $('qrProgress').textContent=`Insertando QR ${i+1} / ${ready.length}: ${p.name}`;
  out.file(filename,await qrStamped(p,f));originals.file(filename,f);
  rows.push({'Nombre completo':p.name,'Correo':p.email,'Enlace Drive':p.link,'PDF final':filename,'PDF firmado':filename,'Estado Outlook':'PENDIENTE'});
  if(i%3===0)await new Promise(r=>setTimeout(r,0));
 }
 const book=XLSX.utils.book_new(),sheet=XLSX.utils.json_to_sheet(rows);sheet['!cols']=[{wch:42},{wch:38},{wch:64},{wch:50},{wch:50},{wch:18}];XLSX.utils.book_append_sheet(book,sheet,'Envios');
 zip.file('CONTROL_ENVIO_OUTLOOK.xlsx',XLSX.write(book,{bookType:'xlsx',type:'array'}));
 zip.file('LEEME.txt','DIPLOMAS_FINALES_QR: PDFs con QR dentro del diploma.\nORIGINALES_FIRMADOS: copias firmadas sin modificar.\nAVISO: la modificación posterior a la firma puede invalidar la firma digital.\n');
 const macro=await fetch('outlook/PrepararBorradoresConstancias.bas');if(macro.ok)zip.file('outlook/PrepararBorradoresConstancias.bas',await macro.text());
 downloadBlob(await zip.generateAsync({type:'blob',compression:'DEFLATE'}),'MIMP_DIPLOMAS_QR_Y_OUTLOOK.zip');
 $('qrProgress').textContent=`Listo: ${ready.length} diplomas con QR y control de correo.`;
}
$('generarFinal').addEventListener('click',()=>generatePackage().catch(e=>$('qrProgress').textContent='Error: '+e.message));
$('paquete').addEventListener('click',()=>generatePackage().catch(e=>$('extensionStatus').textContent='Error: '+e.message));
renderMatches();
