/* Les originaux sont signés localement. La conversion Office en PDF utilise le OneDrive connecté. */
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
const el = id => document.getElementById(id);
const canvas = el('canvas');
const ctx = canvas.getContext('2d');
const stamp = el('stamp');
const stampCaption = el('stampCaption');
const DEFAULT_SIGNER_NAME = 'F.DECAMPS-Directeur';
const state = {item:null,files:[],active:0,page:1,pdf:null,zoom:1,stampUrl:null,stampBytes:null,signatures:[],selectedSignatureId:null,storageError:null,position:{x:.7,y:.78},placements:new Map(),renderToken:0,selectionToken:0,busy:false};
const clamp=(value,min,max)=>Math.min(max,Math.max(min,value));
function message(text,error=false){el('status').textContent=text;el('status').classList.toggle('error',error)}
function decodeBase64(text){const binary=atob(text);const bytes=new Uint8Array(binary.length);for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);return bytes}
function encodeBase64(bytes){let result='';for(let i=0;i<bytes.length;i+=32768)result+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(result)}
function signerName(){return el('signerName').value.trim()||DEFAULT_SIGNER_NAME}
function signatureDate(value=new Date()){
  return new Intl.DateTimeFormat('fr-FR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:false}).format(value).replace(' à ', ' - ');
}
function signatureCaption(value=new Date()){return `${signerName()}\n${signatureDate(value)}`}
function loadSignerName(){
  try{el('signerName').value=localStorage.getItem('outlook-pdf-signer-name')||DEFAULT_SIGNER_NAME}catch{el('signerName').value=DEFAULT_SIGNER_NAME}
}
function openSignatureDb(){return new Promise((resolve,reject)=>{
  if(!window.indexedDB){reject(new Error('Le stockage local est indisponible.'));return}
  const request=indexedDB.open('outlook-pdf-signatures',1);
  request.onupgradeneeded=()=>request.result.createObjectStore('signatures',{keyPath:'id'});
  request.onsuccess=()=>resolve(request.result);
  request.onerror=()=>reject(request.error||new Error('Ouverture du stockage impossible.'));
})}
async function signatureStore(mode,operation){
  const db=await openSignatureDb();
  try{return await new Promise((resolve,reject)=>{
    const transaction=db.transaction('signatures',mode);const request=operation(transaction.objectStore('signatures'));
    let result;
    request.onsuccess=()=>{result=request.result};
    request.onerror=()=>reject(request.error||new Error('Enregistrement impossible.'));
    transaction.onerror=()=>reject(transaction.error||new Error('Enregistrement impossible.'));
    transaction.oncomplete=()=>resolve(result);
  })}finally{db.close()}
}
function selectedId(){try{return localStorage.getItem('outlook-pdf-selected-signature')}catch{return null}}
function rememberSelection(id){try{if(id)localStorage.setItem('outlook-pdf-selected-signature',id);else localStorage.removeItem('outlook-pdf-selected-signature')}catch{}}
function renderSignatureChoices(){
  const select=el('savedStamps');select.replaceChildren();
  const empty=document.createElement('option');empty.value='';empty.textContent=state.signatures.length?'Choisir une signature':'Aucune signature enregistrée';select.append(empty);
  for(const signature of state.signatures){const option=document.createElement('option');option.value=signature.id;option.textContent=signature.name;select.append(option)}
  select.value=state.selectedSignatureId||'';
  el('clearStamp').disabled=!state.selectedSignatureId;
}
function setStamp(signature){
  state.selectedSignatureId=signature?.id||null;
  state.stampUrl=signature?.dataUrl||null;
  state.stampBytes=signature?decodeBase64(signature.dataUrl.split(',')[1]):null;
  if(signature){stamp.src=signature.dataUrl;el('stampName').textContent=signature.name}
  else{stamp.removeAttribute('src');stamp.style.display='none';el('stampName').textContent='Aucune signature choisie'}
  rememberSelection(state.selectedSignatureId);renderSignatureChoices();placeStamp();
}
async function loadSignatures(){
  try{
    state.signatures=await signatureStore('readonly',store=>store.getAll());
    if(!state.signatures.length){
      let oldUrl=null,oldName='Signature enregistrée';
      try{oldUrl=localStorage.getItem('outlook-pdf-stamp');oldName=localStorage.getItem('outlook-pdf-stamp-name')||oldName}catch{}
      if(oldUrl?.startsWith('data:image/png;base64,')){
        const migrated={id:crypto.randomUUID(),name:oldName,dataUrl:oldUrl};
        await signatureStore('readwrite',store=>store.put(migrated));state.signatures=[migrated];
        try{localStorage.removeItem('outlook-pdf-stamp');localStorage.removeItem('outlook-pdf-stamp-name')}catch{}
      }
    }
    state.signatures.sort((a,b)=>a.name.localeCompare(b.name,'fr'));
    const preferred=state.signatures.find(s=>s.id===selectedId())||state.signatures[0];
    setStamp(preferred||null);
  }catch(error){state.storageError='Signatures non mémorisées : '+error.message;renderSignatureChoices()}
}
function getAttachment(item,id){return new Promise((resolve,reject)=>item.getAttachmentContentAsync(id,result=>result.status===Office.AsyncResultStatus.Succeeded?resolve(result.value):reject(new Error(result.error?.message||'Lecture impossible'))))}
async function initialize(){
  try{
    loadSignerName();
    await loadSignatures();
    if(typeof Office==='undefined')throw new Error('Ouvrez cette page depuis un message Outlook.');
    await Office.onReady();
    if(!Office.context?.mailbox)throw new Error('Ouvrez cette extension depuis un message Outlook.');
    if(!Office.context.requirements?.isSetSupported('Mailbox','1.15'))throw new Error('Cette version d’Outlook ne permet pas de joindre automatiquement les PDF signés.');
    const item=Office.context.mailbox.item;
    if(!item?.attachments||!item.displayReplyFormAsync)throw new Error('Ouvrez un message reçu dans Outlook.');
    state.item=item;
    const attachments=item.attachments.filter(a=>!a.isInline&&/\.(pdf|docx|xlsx)$/i.test(a.name||''));
    if(!attachments.length){message('Aucun PDF, Word (.docx) ou Excel (.xlsx) joint. Les anciens formats .doc et .xls doivent être réenregistrés dans Office.');return}
    state.files=attachments.map(a=>({id:a.id,name:a.name,kind:a.name.split('.').pop().toLowerCase(),bytes:null,pdf:null,officeInfo:null,officePlacements:[],output:'original',revision:0}));
    renderFileList();await selectFile(0);
  }catch(error){message(error.message,true)}
}
function renderFileList(){
  const root=el('attachments');root.replaceChildren();
  state.files.forEach((file,index)=>{
    const label=document.createElement('label');const input=document.createElement('input');input.type='radio';input.name='attachment';input.checked=index===state.active;
    input.addEventListener('change',()=>selectFile(index));label.append(input,document.createTextNode(file.name));root.append(label);
  });
}
function showZoom(){
  el('zoomValue').textContent=Math.round(state.zoom*100)+' %';
  el('zoomOut').disabled=state.zoom<=.5;
  el('zoomIn').disabled=state.zoom>=3;
  el('zoomReset').disabled=state.zoom===1;
}
function setBusy(value){
  if(state.busy===value)return;
  state.busy=value;
  for(const node of document.querySelectorAll('button,input,select')){
    if(value){node.dataset.wasDisabled=String(node.disabled);node.disabled=true}
    else{node.disabled=node.dataset.wasDisabled==='true';delete node.dataset.wasDisabled}
  }
  if(!value){showZoom();if(state.pdf){el('prev').disabled=state.page===1;el('next').disabled=state.page===state.pdf.numPages}}
}
function updateDocumentView(){
  const file=state.files[state.active];const native=file&&file.kind!=='pdf';
  el('officeControls').hidden=!native;el('m365Controls').hidden=!native;
  el('pdfViewer').hidden=native&&!state.pdf;
  el('pdfPlacementControls').hidden=!!native;el('pdfPlacementHint').hidden=!!native;
  el('size').parentElement.querySelector('label[for="size"]').hidden=!!native;el('size').hidden=!!native;
  if(native&&file.officeInfo){
    el('officeHint').textContent=file.kind==='docx'?'La signature sera ajoutée dans le paragraphe choisi, sous son texte. Elle peut déplacer le contenu suivant.':'La signature sera ajoutée comme une image dans la feuille choisie, à partir de la cellule indiquée.';
    el('officeTargetLabel').textContent=file.kind==='docx'?'Paragraphe où insérer la signature':'Feuille de calcul';
    el('excelCellRow').hidden=file.kind!=='xlsx';
    const select=el('officeTarget');select.replaceChildren();
    file.officeInfo.targets.forEach((target,index)=>{const o=document.createElement('option');o.value=String(target.id);o.textContent=(file.kind==='docx'?`${index+1} — `:'')+target.label+(target.protected?' (protégée)':'');o.disabled=!!target.protected;select.append(o)});
    const first=Array.from(select.options).find(o=>!o.disabled);if(first)select.value=first.value;
    el('outputFormat').value=file.output;
  }
  stamp.style.display='none';stampCaption.style.display='none';renderOfficePlacements();
}
async function selectFile(index){
  if(state.busy)return;
  const selection=++state.selectionToken;
  ++state.renderToken;state.active=index;state.page=1;state.pdf=null;state.zoom=1;showZoom();state.position={x:.7,y:.78};renderFileList();renderPlacements();
  canvas.width=0;canvas.height=0;
  const file=state.files[index];updateDocumentView();message('Ouverture de '+file.name+'…');
  try{
    if(!file.bytes){
      const content=await getAttachment(state.item,file.id);
      if(content.format!==Office.MailboxEnums.AttachmentContentFormat.Base64)throw new Error('Un fichier joint est nécessaire, pas un lien vers un document.');
      file.bytes=decodeBase64(content.content);
    }
    if(selection!==state.selectionToken)return;
    if(file.kind==='pdf'){
      if(!file.pdf)file.pdf=await pdfjsLib.getDocument({isEvalSupported:false,data:file.bytes.slice()}).promise;
      if(selection!==state.selectionToken)return;
      state.pdf=file.pdf;updateDocumentView();await renderPage();
    }else{
      if(!file.officeInfo)file.officeInfo=await OfficeFiles.inspect(file.bytes,file.kind);
      if(selection!==state.selectionToken)return;
      if(file.previewPdf&&file.previewRevision===file.revision)state.pdf=file.previewPdf;
      updateDocumentView();if(state.pdf)await renderPage();
    }
    if(selection===state.selectionToken)message(file.name+' prêt.'+(state.storageError?' '+state.storageError:''),!!state.storageError);
  }catch(error){if(selection===state.selectionToken)message('Impossible d’ouvrir '+file.name+' : '+error.message,true)}
}
async function renderPage(preserveView=false){
  if(!state.pdf)return;
  const token=++state.renderToken;const page=await state.pdf.getPage(state.page);if(token!==state.renderToken)return;
  const base=page.getViewport({scale:1});
  const target=Math.max(210,Math.min(650,el('paper').parentElement.getBoundingClientRect().width-28));
  const viewport=page.getViewport({scale:target/base.width*state.zoom});
  if(token!==state.renderToken)return;
  const rendered=document.createElement('canvas');rendered.width=Math.round(viewport.width);rendered.height=Math.round(viewport.height);
  await page.render({canvasContext:rendered.getContext('2d'),viewport}).promise;
  if(token!==state.renderToken)return;
  const wrap=el('paper').parentElement;
  const centerX=preserveView?(wrap.scrollLeft+wrap.clientWidth/2)/(canvas.width||1):0;
  const centerY=preserveView?(wrap.scrollTop+wrap.clientHeight/2)/(canvas.height||1):0;
  canvas.width=rendered.width;canvas.height=rendered.height;ctx.drawImage(rendered,0,0);
  if(preserveView){wrap.scrollLeft=centerX*canvas.width-wrap.clientWidth/2;wrap.scrollTop=centerY*canvas.height-wrap.clientHeight/2}
  else{wrap.scrollLeft=0;wrap.scrollTop=0}
  el('pageInfo').textContent=`Page ${state.page}/${state.pdf.numPages}`;
  el('prev').disabled=state.page===1;el('next').disabled=state.page===state.pdf.numPages;
  placeStamp();renderPlacements();
}
function placeStamp(){
  if(state.files[state.active]?.kind!=='pdf'||!state.stampUrl||!stamp.naturalWidth||!canvas.width){stamp.style.display='none';stampCaption.style.display='none';return}
  const width=canvas.width*Number(el('size').value)/100,height=width*stamp.naturalHeight/stamp.naturalWidth;
  stamp.style.width=width+'px';stamp.style.height=height+'px';
  stamp.style.left=(state.position.x*canvas.width-width/2)+'px';
  stamp.style.top=(state.position.y*canvas.height-height/2)+'px';
  stampCaption.textContent=signatureCaption();
  stampCaption.style.left=(state.position.x*canvas.width)+'px';
  stampCaption.style.top=(state.position.y*canvas.height+height/2+4)+'px';
  stamp.style.display='block';
  stampCaption.style.display='block';
}
stamp.onload=placeStamp;
let drag=null;
stamp.onpointerdown=event=>{event.preventDefault();const r=stamp.getBoundingClientRect();drag={x:event.clientX-r.left,y:event.clientY-r.top};stamp.setPointerCapture(event.pointerId)};
stamp.onpointermove=event=>{
  if(!drag)return;const r=canvas.getBoundingClientRect();
  const left=clamp(event.clientX-r.left-drag.x,-stamp.clientWidth/2,r.width-stamp.clientWidth/2);
  const top=clamp(event.clientY-r.top-drag.y,-stamp.clientHeight/2,r.height-stamp.clientHeight/2);
  state.position={x:(left+stamp.clientWidth/2)/r.width,y:(top+stamp.clientHeight/2)/r.height};placeStamp();
};
stamp.onpointerup=stamp.onpointercancel=()=>drag=null;
el('stampFile').onchange=async event=>{
  const file=event.target.files?.[0];if(!file)return;
  try{
    const bytes=new Uint8Array(await file.arrayBuffer());
    if(bytes.length<8||![137,80,78,71,13,10,26,10].every((value,index)=>bytes[index]===value))throw new Error('Choisissez une image PNG valide.');
    const signature={id:crypto.randomUUID(),name:file.name,dataUrl:'data:image/png;base64,'+encodeBase64(bytes)};
    await signatureStore('readwrite',store=>store.put(signature));
    state.signatures.push(signature);state.signatures.sort((a,b)=>a.name.localeCompare(b.name,'fr'));
    state.storageError=null;setStamp(signature);message('Signature « '+file.name+' » enregistrée sur cet appareil.');
  }catch(error){message('Impossible de mémoriser la signature : '+error.message,true)}
  finally{event.target.value=''}
};
el('savedStamps').onchange=event=>{const signature=state.signatures.find(s=>s.id===event.target.value);setStamp(signature||null)};
el('signerName').oninput=()=>placeStamp();
el('signerName').onchange=()=>{
  const value=signerName();el('signerName').value=value;
  try{localStorage.setItem('outlook-pdf-signer-name',value)}catch{}
  placeStamp();
};
el('clearStamp').onclick=async()=>{
  const id=state.selectedSignatureId;if(!id)return;
  try{
    await signatureStore('readwrite',store=>store.delete(id));
    state.signatures=state.signatures.filter(s=>s.id!==id);setStamp(state.signatures[0]||null);
    message('Signature supprimée de cet appareil.');
  }catch(error){message('Impossible de supprimer la signature : '+error.message,true)}
};
el('size').oninput=()=>{el('sizeValue').textContent=el('size').value+' %';placeStamp()};
el('zoomOut').onclick=()=>{state.zoom=clamp(state.zoom-.25,.5,3);showZoom();renderPage(true)};
el('zoomIn').onclick=()=>{state.zoom=clamp(state.zoom+.25,.5,3);showZoom();renderPage(true)};
el('zoomReset').onclick=()=>{state.zoom=1;showZoom();renderPage(true)};
showZoom();
el('prev').onclick=()=>{if(state.pdf&&state.page>1){state.page--;renderPage()}};
el('next').onclick=()=>{if(state.pdf&&state.page<state.pdf.numPages){state.page++;renderPage()}};
el('addPlacement').onclick=()=>{
  if(state.busy||state.files[state.active]?.kind!=='pdf'||!state.pdf||!state.stampBytes){message('Ouvrez un PDF et choisissez un tampon PNG.',true);return}
  const key=state.active+':'+state.page;
  state.placements.set(key,{file:state.active,page:state.page,x:state.position.x,y:state.position.y,size:Number(el('size').value),signature:state.stampUrl,signerName:signerName(),signedAt:new Date().toISOString()});
  renderPlacements();message('Tampon ajouté sur la page '+state.page+'.');
};
function renderPlacements(){
  const root=el('placements');root.replaceChildren();
  const selected=[...state.placements.entries()].filter(([,p])=>p.file===state.active).sort((a,b)=>a[1].page-b[1].page);
  for(const [key,p] of selected){const row=document.createElement('div');const label=document.createElement('span');label.textContent='Tampon sur la page '+p.page;
    const remove=document.createElement('button');remove.type='button';remove.textContent='Retirer';remove.onclick=()=>{state.placements.delete(key);renderPlacements()};row.append(label,remove);root.append(row)}
}
function openReply(formData){return new Promise((resolve,reject)=>state.item.displayReplyFormAsync(formData,result=>result.status===Office.AsyncResultStatus.Succeeded?resolve():reject(new Error(result.error?.message||'Réponse impossible'))))}
el('reply').onclick=async()=>{
  if(state.busy)return;
  if(state.placements.size===0&&!state.files.some(f=>f.officePlacements.length)){message('Ajoutez au moins une signature.',true);return}
  setBusy(true);message('Création des documents signés…');
  try{
    const attachments=[];
    for(let index=0;index<state.files.length;index++){
      const file=state.files[index];
      if(file.kind!=='pdf'){
        if(!file.officePlacements.length)continue;
        let bytes=await signedOffice(file),extension=file.kind;
        if(file.output==='pdf'){bytes=await officePdf(file,bytes);extension='pdf'}
        if(bytes.length>25*1024*1024)throw new Error(file.name+' dépasse 25 Mo.');
        attachments.push({type:Office.MailboxEnums.AttachmentType.Base64,name:file.name.replace(/\.[^.]+$/,'')+' - signé.'+extension,base64file:encodeBase64(bytes),inLine:false});
        continue;
      }
      const placements=[...state.placements.values()].filter(p=>p.file===index);
      if(!placements.length)continue;
      const doc=await PDFLib.PDFDocument.load(file.bytes.slice());const images=new Map();
      const captionFont=await doc.embedFont(PDFLib.StandardFonts.Helvetica);
      for(const p of placements){const page=doc.getPages()[p.page-1];if(!page)continue;
        if(!images.has(p.signature))images.set(p.signature,await doc.embedPng(decodeBase64(p.signature.split(',')[1])));
        const image=images.get(p.signature);
        const width=page.getWidth()*p.size/100,height=width*image.height/image.width;
        const imageX=p.x*page.getWidth()-width/2;
        const imageY=(1-p.y)*page.getHeight()-height/2;
        page.drawImage(image,{x:imageX,y:imageY,width,height});
        const fontSize=Math.max(7,Math.min(11,page.getWidth()*.014));
        const lines=[p.signerName||DEFAULT_SIGNER_NAME,signatureDate(new Date(p.signedAt))];
        lines.forEach((line,lineIndex)=>{
          const textWidth=captionFont.widthOfTextAtSize(line,fontSize);
          page.drawText(line,{x:p.x*page.getWidth()-textWidth/2,y:imageY-fontSize*(lineIndex+1)-3,size:fontSize,font:captionFont,color:PDFLib.rgb(.07,.07,.07)});
        });
      }
      const bytes=await doc.save();
      if(bytes.length>25*1024*1024)throw new Error(file.name+' dépasse la limite Outlook de 25 Mo.');
      attachments.push({type:Office.MailboxEnums.AttachmentType.Base64,name:file.name.replace(/\.pdf$/i,'')+' - signé.pdf',base64file:encodeBase64(bytes),inLine:false});
    }
    if(!attachments.length)throw new Error('Aucun document signé à joindre.');
    message('Ouverture de la réponse…');
    await openReply({htmlBody:attachments.length===1?'Bonjour,<br><br>Veuillez trouver ci-joint le document signé.<br><br>Cordialement,':'Bonjour,<br><br>Veuillez trouver ci-joint les documents signés.<br><br>Cordialement,',attachments});
    message('Réponse ouverte. Vérifiez les pièces jointes et les destinataires avant l’envoi.');
  }catch(error){message('Impossible de préparer la réponse : '+error.message,true)}
  finally{setBusy(false)}
};
function renderOfficePlacements(){
  const root=el('officePlacements');root.replaceChildren();
  const file=state.files[state.active];if(!file||file.kind==='pdf')return;
  for(const [index,p] of file.officePlacements.entries()){
    const row=document.createElement('div'),label=document.createElement('span'),remove=document.createElement('button');
    label.textContent=p.label;remove.type='button';remove.textContent='Retirer';
    remove.onclick=()=>{if(state.busy)return;file.officePlacements.splice(index,1);invalidateOffice(file);renderOfficePlacements()};
    row.append(label,remove);root.append(row);
  }
}
function invalidateOffice(file){
  file.revision++;file.signedBytes=null;file.convertedBytes=null;file.previewPdf=null;
  if(state.files[state.active]===file){++state.renderToken;state.pdf=null;el('pdfViewer').hidden=true;stamp.style.display='none';stampCaption.style.display='none'}
}
async function compositeStamp(widthCm){
  if(!state.stampUrl)throw new Error('Choisissez une signature PNG.');
  const img=new Image();img.src=state.stampUrl;await img.decode();
  const c=document.createElement('canvas');c.width=900;
  const imageHeight=Math.round(c.width*img.naturalHeight/img.naturalWidth);
  if(!Number.isFinite(imageHeight)||imageHeight>3000)throw new Error('Choisissez une signature plus large que haute.');
  let fontSize=Math.round(9/72*2.54/widthCm*c.width);
  const caption=signatureCaption(),lines=caption.split('\n');
  const measure=c.getContext('2d');measure.font=fontSize+'px Arial';
  const longest=Math.max(...lines.map(line=>measure.measureText(line).width));if(longest>880)fontSize=Math.floor(fontSize*880/longest);
  c.height=imageHeight+Math.round(fontSize*2.7)+12;
  const context=c.getContext('2d');context.drawImage(img,0,0,c.width,imageHeight);context.font=fontSize+'px Arial';context.fillStyle='#111';context.textAlign='center';
  lines.forEach((line,i)=>context.fillText(line,c.width/2,imageHeight+fontSize*(i+1.1)+4));
  return {image:c.toDataURL('image/png'),pixelWidth:c.width,pixelHeight:c.height,caption};
}
el('addOfficePlacement').onclick=async()=>{
  if(state.busy)return;
  const file=state.files[state.active];if(!file?.officeInfo)return;
  const target=el('officeTarget').value,widthCm=Number(el('officeWidth').value),cell=el('excelCell').value;
  try{
    if(!target||!Number.isFinite(widthCm)||widthCm<2||widthCm>12)throw new Error('Choisissez un emplacement et une largeur entre 2 et 12 cm.');
    if(file.kind==='xlsx')OfficeFiles.cellAddress(cell);
    setBusy(true);
    const composed=await compositeStamp(widthCm);
    const label=file.kind==='docx'?'Signature au paragraphe '+(Number(target)+1):'Signature — '+file.officeInfo.targets.find(t=>t.id===target).label+' · '+cell.toUpperCase();
    file.officePlacements.push({...composed,target,widthCm,cell,label});invalidateOffice(file);renderOfficePlacements();message('Signature ajoutée. Choisissez le format de renvoi ou vérifiez l’aperçu PDF.');
  }catch(error){message(error.message,true)}finally{setBusy(false)}
};
el('outputFormat').onchange=event=>{const file=state.files[state.active];if(file)file.output=event.target.value};
async function signedOffice(file){
  if(!file.signedBytes)file.signedBytes=await OfficeFiles.sign(file.bytes,file.kind,file.officePlacements);
  return file.signedBytes;
}
async function officePdf(file,bytes){
  if(!file.convertedBytes){const result=await OneDrivePDF.convert(bytes,file.kind);file.convertedBytes=result.bytes;if(result.warning)el('cleanupWarning').textContent=result.warning}
  return file.convertedBytes;
}
el('previewOffice').onclick=async()=>{
  if(state.busy)return;const file=state.files[state.active];
  if(!file?.officePlacements.length){message('Ajoutez une signature avant de demander l’aperçu.',true);return}
  setBusy(true);message('Conversion et chargement de l’aperçu…');
  try{
    const bytes=await officePdf(file,await signedOffice(file));
    file.previewPdf=await pdfjsLib.getDocument({isEvalSupported:false,data:bytes.slice()}).promise;file.previewRevision=file.revision;
    state.pdf=file.previewPdf;state.page=1;state.zoom=1;el('pdfViewer').hidden=false;await renderPage();
    message('Aperçu du document signé. Vérifiez la position de la signature et les pages avant de préparer la réponse.');
  }catch(error){message(error.message,true)}finally{setBusy(false);showZoom()}
};
function m365Status(){
  const account=OneDrivePDF.account();el('m365Status').textContent=account?'OneDrive connecté : '+account:'Non connecté.';
  el('disconnectM365').hidden=!account;el('connectM365').textContent=account?'Reconnecter / changer de compte':'Connecter mon OneDrive professionnel';
  el('m365Setup').hidden=OneDrivePDF.configured();el('connectM365').disabled=!OneDrivePDF.configured();
}
el('connectM365').onclick=async()=>{
  if(state.busy)return;setBusy(true);
  try{await OneDrivePDF.connect();const remaining=await OneDrivePDF.retryCleanup();el('cleanupWarning').textContent=remaining?remaining+' copie(s) temporaire(s) restent dans OneDrive. Réessayez la connexion.':'';message('Connexion Microsoft 365 établie.')}
  catch(error){message(error.message,true)}finally{setBusy(false);m365Status()}
};
el('disconnectM365').onclick=()=>{OneDrivePDF.disconnect();m365Status();message('Le complément est déconnecté de OneDrive.')};
m365Status();
initialize();

