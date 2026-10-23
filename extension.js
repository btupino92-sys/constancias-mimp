'use strict';
/* Extensión pasos 4 y 5: nunca reescribe los PDF firmados. */
let signedFiles = new Map();
const normName = s => String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\.pdf$/i,'').replace(/\s*\(\d+\)\s*$/,'').replace(/[_\-]+/g,' ').replace(/\s+/g,' ').trim();
const validDriveLink = s => {try {let u=new URL(s);return u.protocol==='https:' && /(^|\.)drive\.google\.com$/.test(u.hostname)}catch{return false}};
const matching = p => signedFiles.get(normName(safeFile(p.name))) || null;
function signedStats(){let matched=participants.filter(matching).length;let links=participants.filter(p=>validDriveLink(p.link)).length;let ready=participants.filter(p=>matching(p)&&validDriveLink(p.link)&&mailOK(p.email)).length;
 $('signedCount').textContent=matched+' / '+participants.length+' PDF firmados vinculados'; $('linkCount').textContent=links+' / '+participants.length+' enlaces de Drive válidos'; $('readyCount').textContent=ready+' participantes listos para preparar';
 let errors=participants.filter(p=>!matching(p)||!validDriveLink(p.link)||!mailOK(p.email));
 $('validation').replaceChildren();errors.slice(0,15).forEach(p=>{let li=document.createElement('li');li.textContent=p.name+': '+[!matching(p)?'sin PDF firmado':null,!validDriveLink(p.link)?'sin enlace válido de Drive':null,!mailOK(p.email)?'sin correo válido':null].filter(Boolean).join(', ');$('validation').append(li)});if(errors.length>15){let li=document.createElement('li');li.textContent='Y '+(errors.length-15)+' registros más.';$('validation').append(li)}
}
function extensionStatus(s){$('extensionStatus').textContent=s}
$('firmados').addEventListener('change',e=>{signedFiles.clear();let collisions=[];for(let f of e.target.files){if(!/\.pdf$/i.test(f.name))continue;let key=normName(f.name);if(signedFiles.has(key))collisions.push(f.name);else signedFiles.set(key,f)}signedStats();extensionStatus('Se seleccionaron '+signedFiles.size+' PDF. '+(collisions.length?'Archivos duplicados omitidos: '+collisions.join(', '):'Los archivos se vinculan por nombre completo.'));});
$('archivoEnlaces').addEventListener('change',()=>setTimeout(signedStats,250));
$('archivo').addEventListener('change',()=>setTimeout(signedStats,250));
$('agregar').addEventListener('click',()=>setTimeout(signedStats,0));
$('limpiar').addEventListener('click',()=>setTimeout(signedStats,0));
$('enlaces').addEventListener('DOMSubtreeModified',()=>{}); // las validaciones principales se ejecutan al generar paquete
async function createVerificationPdf(p){
 const {PDFDocument,StandardFonts,rgb}=PDFLib;
 const doc=await PDFDocument.create();const page=doc.addPage([595.28,841.89]);const reg=await doc.embedFont(StandardFonts.Helvetica);const bold=await doc.embedFont(StandardFonts.HelveticaBold);
 const dark=rgb(.12,.17,.25),red=rgb(.72,.07,.12);page.drawText('MIMP  |  CONSTANCIA DE PARTICIPACION',{x:54,y:774,font:bold,size:16,color:red});page.drawLine({start:{x:54,y:757},end:{x:541,y:757},thickness:1,color:red});
 page.drawText('HOJA COMPLEMENTARIA DE VERIFICACION',{x:54,y:715,font:bold,size:16,color:dark});
 const wrap=(s,max)=>{const words=String(s).split(/\s+/),lines=[];let line='';for(let word of words){let tryLine=line?line+' '+word:word;if(reg.widthOfTextAtSize(tryLine,12)<max)line=tryLine;else{if(line)lines.push(line);line=word}}if(line)lines.push(line);return lines};
 let y=667;page.drawText('Participante:',{x:54,y,font:bold,size:12,color:dark});y-=24;for(let l of wrap(p.name,480)){page.drawText(l,{x:54,y,font:reg,size:12,color:dark});y-=19}
 y-=15;page.drawText('Escanee el codigo QR para consultar el PDF firmado en Drive:',{x:54,y,font:reg,size:11,color:dark});
 const data=await QRCode.toDataURL(p.link,{width:650,margin:2,errorCorrectionLevel:'M'});const b64=data.split(',')[1];const bytes=Uint8Array.from(atob(b64),c=>c.charCodeAt(0));const qr=await doc.embedPng(bytes);page.drawImage(qr,{x:184,y:297,width:230,height:230});
 y=270;for(let line of wrap(p.link,480)){page.drawText(line,{x:54,y,font:reg,size:10,color:dark});y-=16}
 page.drawText('Esta hoja no altera el documento firmado digitalmente.',{x:54,y:130,font:bold,size:10,color:red});
 page.drawText('Verifique el acceso al enlace y la validez de la firma en el PDF original.',{x:54,y:110,font:reg,size:10,color:dark});
 return await doc.save();
}
$('paquete').addEventListener('click',async()=>{try{
 needLibraries('XLSX','JSZip','PDFLib','QRCode');if(!participants.length)throw Error('Primero importe participantes en el paso 2.');signedStats();
 let ready=participants.filter(p=>matching(p)&&validDriveLink(p.link)&&mailOK(p.email));if(!ready.length)throw Error('No hay registros completos con nombre, correo, PDF firmado y enlace de Drive.');
 if(ready.length<participants.length&&!confirm('Solo '+ready.length+' de '+participants.length+' participantes están completos. ¿Generar paquete únicamente con ellos?'))return;
 const zip=new JSZip(),signed=zip.folder('PDF_FIRMADOS'),qrDir=zip.folder('QR_COMPLEMENTARIOS'),control=[];let used=new Set();
 for(let i=0;i<ready.length;i++){
 let p=ready[i],base=safeFile(p.name),n=2,filename=base+'.pdf';while(used.has(filename.toLowerCase()))filename=base+'_'+n+++'.pdf';used.add(filename.toLowerCase());
 signed.file(filename,matching(p));let companion=base+'_QR.pdf';qrDir.file(companion,await createVerificationPdf(p));
 control.push({'Nombre completo':p.name,'Correo':p.email,'Enlace Drive':p.link,'PDF firmado':filename,'PDF QR':companion,'Estado Outlook':'PENDIENTE'});
 extensionStatus('Preparando '+(i+1)+' de '+ready.length+' participantes...');if(i%4===0)await new Promise(r=>setTimeout(r,0));
 }
 const wb=XLSX.utils.book_new();let ws=XLSX.utils.json_to_sheet(control);ws['!cols']=[{wch:45},{wch:36},{wch:65},{wch:50},{wch:55},{wch:18}];XLSX.utils.book_append_sheet(wb,ws,'Envios');let xlsx=XLSX.write(wb,{bookType:'xlsx',type:'array'});zip.file('CONTROL_ENVIO_OUTLOOK.xlsx',xlsx);
 zip.file('MACRO_OUTLOOK/PrepararBorradoresConstancias.bas',await (await fetch('outlook/PrepararBorradoresConstancias.bas')).text());
 zip.file('LEEME_ENVIO.txt',`PAQUETE DE CONSTANCIAS MIMP\n\n1. EXTRAIGA TODO EL ZIP EN UNA CARPETA LOCAL, sin modificar los nombres de carpetas.\n2. El Excel CONTROL_ENVIO_OUTLOOK.xlsx contiene los nombres, correos, enlaces de Drive y los nombres de los adjuntos.\n3. Abra Excel de escritorio con un libro habilitado para macros (.xlsm).\n4. ALT+F11 > Archivo > Importar archivo... > MACRO_OUTLOOK/PrepararBorradoresConstancias.bas.\n5. Ejecute la macro PrepararBorradoresConstancias (ALT+F8).\n6. Seleccione CONTROL_ENVIO_OUTLOOK.xlsx y luego la CARPETA RAIZ extraida.\n7. Revise uno por uno los borradores de Outlook antes de enviarlos.\n\nSe adjunta el PDF firmado ORIGINAL, SIN CAMBIOS, junto con la hoja complementaria de verificacion QR.\nOutlook clasico de escritorio para Windows es necesario; el nuevo Outlook puede no soportar esta automatizacion VBA.\nNo cargue documentos personales ni datos de participantes al repositorio PUBLICO de GitHub.\n`);
 extensionStatus('Comprimiendo paquete...');downloadBlob(await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:5}}),'CONSTANCIAS_FIRMADAS_QR_OUTLOOK.zip');extensionStatus('Paquete listo: '+ready.length+' participantes, documentos firmados originales, complementos QR y Excel para Outlook.');
 }catch(err){extensionStatus('Error: '+err.message)}});
$('revisar').addEventListener('click',signedStats);
signedStats();
