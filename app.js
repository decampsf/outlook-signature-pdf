/* Les PDF et le tampon sont traités dans le navigateur. Aucune requête ne les envoie au site. */
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
const el = id => document.getElementById(id);
const canvas = el('canvas');
const ctx = canvas.getContext('2d');
const stamp = el('stamp');
const state = {item:null,files:[],active:0,page:1,pdf:null,stampUrl:null,stampBytes:null,position:{x:.7,y:.78},placements:new Map(),renderToken:0};
const clamp=(value,min,max)=>Math.min(max,Math.max(min,value));
function message(text,error=false){el('status').textContent=text;el('status').classList.toggle('error',error)}
function decodeBase64(text){const binary=atob(text);const bytes=new Uint8Array(binary.length);for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);return bytes}
function encodeBase64(bytes){let result='';for(let i=0;i<bytes.length;i+=32768)result+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(result)}
function getAttachment(item,id){return new Promise((resolve,reject)=>item.getAttachmentContentAsync(id,result=>result.status===Office.AsyncResultStatus.Succeeded?resolve(result.value):reject(new Error(result.error?.message||'Lecture impossible'))))}
async function initialize(){
  try{
    if(typeof Office==='undefined')throw new Error('Ouvrez cette page depuis un message Outlook.');
    await Office.onReady();
    if(!Office.context.requirements.isSetSupported('Mailbox','1.15'))throw new Error('Cette version d’Outlook ne permet pas de joindre automatiquement les PDF signés.');
    const item=Office.context.mailbox.item;
    if(!item?.attachments||!item.displayReplyFormAsync)throw new Error('Ouvrez un message reçu dans Outlook.');
    state.item=item;
    const attachments=item.attachments.filter(a=>!a.isInline&&/\.pdf$/i.test(a.name||''));
    if(!attachments.length){message('Aucun PDF joint à ce message.');return}
    state.files=attachments.map(a=>({id:a.id,name:a.name,bytes:null,pdf:null}));
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
async function selectFile(index){
  state.active=index;state.page=1;state.pdf=null;state.position={x:.7,y:.78};renderFileList();renderPlacements();
  const file=state.files[index];message('Ouverture de '+file.name+'…');
  try{
    if(!file.bytes){
      const content=await getAttachment(state.item,file.id);
      if(content.format!==Office.MailboxEnums.AttachmentContentFormat.Base64)throw new Error('Cette pièce jointe est un lien ou un type non pris en charge. Téléchargez le PDF joint au message.');
      file.bytes=decodeBase64(content.content);
    }
    if(!file.pdf)file.pdf=await pdfjsLib.getDocument({data:file.bytes.slice()}).promise;
    state.pdf=file.pdf;await renderPage();message(file.name+' prêt.');
  }catch(error){message('Impossible d’ouvrir '+file.name+' : '+error.message,true)}
}
async function renderPage(){
  if(!state.pdf)return;
  const token=++state.renderToken;const page=await state.pdf.getPage(state.page);const base=page.getViewport({scale:1});
  const target=Math.max(210,Math.min(650,el('paper').parentElement.clientWidth-28));
  const viewport=page.getViewport({scale:target/base.width});
  if(token!==state.renderToken)return;
  canvas.width=Math.round(viewport.width);canvas.height=Math.round(viewport.height);
  await page.render({canvasContext:ctx,viewport}).promise;
  if(token!==state.renderToken)return;
  el('pageInfo').textContent=`Page ${state.page}/${state.pdf.numPages}`;
  el('prev').disabled=state.page===1;el('next').disabled=state.page===state.pdf.numPages;
  placeStamp();renderPlacements();
}
function placeStamp(){
  if(!state.stampUrl||!stamp.naturalWidth||!canvas.width){stamp.style.display='none';return}
  const width=canvas.width*Number(el('size').value)/100,height=width*stamp.naturalHeight/stamp.naturalWidth;
  stamp.style.width=width+'px';stamp.style.height=height+'px';
  stamp.style.left=(state.position.x*canvas.width-width/2)+'px';
  stamp.style.top=(state.position.y*canvas.height-height/2)+'px';
  stamp.style.display='block';
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
  if(file.type!=='image/png'){message('Choisissez un fichier PNG.',true);return}
  const bytes=new Uint8Array(await file.arrayBuffer());
  const url='data:image/png;base64,'+encodeBase64(bytes);
  try{localStorage.setItem('outlook-pdf-stamp',url);localStorage.setItem('outlook-pdf-stamp-name',file.name)}catch{}
  setStamp(url,file.name);
};
function setStamp(url,name){state.stampUrl=url;state.stampBytes=decodeBase64(url.split(',')[1]);stamp.src=url;el('stampName').textContent=name;placeStamp()}
el('clearStamp').onclick=()=>{localStorage.removeItem('outlook-pdf-stamp');localStorage.removeItem('outlook-pdf-stamp-name');state.stampUrl=null;state.stampBytes=null;stamp.removeAttribute('src');stamp.style.display='none';el('stampName').textContent='Aucun tampon choisi'};
el('size').oninput=()=>{el('sizeValue').textContent=el('size').value+' %';placeStamp()};
el('prev').onclick=()=>{if(state.pdf&&state.page>1){state.page--;renderPage()}};
el('next').onclick=()=>{if(state.pdf&&state.page<state.pdf.numPages){state.page++;renderPage()}};
el('addPlacement').onclick=()=>{
  if(!state.pdf||!state.stampBytes){message('Ouvrez un PDF et choisissez un tampon PNG.',true);return}
  const key=state.active+':'+state.page;
  state.placements.set(key,{file:state.active,page:state.page,x:state.position.x,y:state.position.y,size:Number(el('size').value)});
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
  if(!state.stampBytes||state.placements.size===0){message('Ajoutez au moins un tampon sur une page.',true);return}
  const button=el('reply');button.disabled=true;message('Création des PDF signés…');
  try{
    const attachments=[];
    for(let index=0;index<state.files.length;index++){
      const file=state.files[index];const placements=[...state.placements.values()].filter(p=>p.file===index);
      if(!placements.length)continue;
      const doc=await PDFLib.PDFDocument.load(file.bytes.slice());const image=await doc.embedPng(state.stampBytes);
      for(const p of placements){const page=doc.getPages()[p.page-1];if(!page)continue;
        const width=page.getWidth()*p.size/100,height=width*image.height/image.width;
        page.drawImage(image,{x:p.x*page.getWidth()-width/2,y:(1-p.y)*page.getHeight()-height/2,width,height});
      }
      const bytes=await doc.save();
      if(bytes.length>25*1024*1024)throw new Error(file.name+' dépasse la limite Outlook de 25 Mo.');
      attachments.push({type:Office.MailboxEnums.AttachmentType.Base64,name:file.name.replace(/\.pdf$/i,'')+' - signé.pdf',base64file:encodeBase64(bytes),inLine:false});
    }
    if(!attachments.length)throw new Error('Aucun PDF signé à joindre.');
    message('Ouverture de la réponse…');
    await openReply({htmlBody:attachments.length===1?'Bonjour,<br><br>Veuillez trouver ci-joint le document signé.<br><br>Cordialement,':'Bonjour,<br><br>Veuillez trouver ci-joint les documents signés.<br><br>Cordialement,',attachments});
    message('Réponse ouverte. Vérifiez les pièces jointes et les destinataires avant l’envoi.');
  }catch(error){message('Impossible de préparer la réponse : '+error.message,true)}
  finally{button.disabled=false}
};
try{const saved=localStorage.getItem('outlook-pdf-stamp');if(saved)setStamp(saved,localStorage.getItem('outlook-pdf-stamp-name')||'Tampon enregistré')}catch{}
initialize();
