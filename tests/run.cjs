const fs=require('node:fs'),path=require('node:path'),http=require('node:http'),assert=require('node:assert/strict'),{execFileSync}=require('node:child_process');
const {chromium}=require('playwright');const {PDFDocument}=require('pdf-lib');
(async()=>{
  execFileSync('python',['tests/fixtures.py'],{stdio:'inherit'});
  const source={};for(const ext of ['docx','xlsx'])source[ext]=fs.readFileSync(`test-results/source.${ext}`).toString('base64');
  const stamp=fs.readFileSync('test-results/signature.png').toString('base64');
  const pdf=await PDFDocument.create();pdf.addPage([595,842]).drawText('Document de test');const pdfBytes=await pdf.save();source.pdf=Buffer.from(pdfBytes).toString('base64');
  const server=http.createServer((req,res)=>{let file=decodeURIComponent(req.url.split('?')[0]);if(file==='/empty'){res.setHeader('Content-Type','text/html');res.end('<!doctype html><html><head><meta charset="utf-8"></head><body></body></html>');return;}if(file==='/')file='/index.html';const absolute=path.resolve('.'+file);if(!absolute.startsWith(process.cwd()+path.sep)||!fs.existsSync(absolute)||fs.statSync(absolute).isDirectory()){res.writeHead(404);res.end();return;}res.setHeader('Content-Type',file.endsWith('.js')?'text/javascript; charset=utf-8':file.endsWith('.html')?'text/html; charset=utf-8':file.endsWith('.css')?'text/css; charset=utf-8':'text/plain; charset=utf-8');res.end(fs.readFileSync(absolute));});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));const base=`http://127.0.0.1:${server.address().port}`;
  let browser;let checks=0;
  try{
  browser=await chromium.launch({headless:true, args:process.env.CHROMIUM_ARGS_FILE?JSON.parse(fs.readFileSync(process.env.CHROMIUM_ARGS_FILE)):[], ...(process.env.CHROMIUM_PATH?{executablePath:process.env.CHROMIUM_PATH}:{})});

    const page=await browser.newPage({viewport:{width:420,height:1000}});const errors=[];page.on('pageerror',e=>errors.push(e.message));
    await page.goto(base+'/empty');await page.addScriptTag({url:base+'/vendor/office-files.js'});
    const result=await page.evaluate(async({source,stamp})=>{
      const bytes=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));const encode=b=>{let s='';for(const n of b)s+=String.fromCharCode(n);return btoa(s)};
      const doc=await OfficeFiles.inspect(bytes(source.docx),'docx');const book=await OfficeFiles.inspect(bytes(source.xlsx),'xlsx');
      if(book.targets.some(t=>t.label==='Archives'))throw Error('Hidden sheet exposed');
      const p={image:'data:image/png;base64,'+stamp,pixelWidth:500,pixelHeight:130,widthCm:5,caption:'Directeur & date'};
      const target=doc.targets.find(t=>t.label==='Signature ici').id;
      const signedDoc=await OfficeFiles.sign(bytes(source.docx),'docx',[{...p,target},{...p,target:1}]);
      const signedBook=await OfficeFiles.sign(bytes(source.xlsx),'xlsx',[{...p,target:book.targets[0].id,cell:'B12'},{...p,target:book.targets[0].id,cell:'B20'}]);
      let rejected=0;for(const cell of ['A0','XFE1','A1048577','A1:B2']){try{OfficeFiles.cellAddress(cell)}catch{rejected++}}
      for(const item of [{target:book.targets.find(t=>t.protected).id,cell:'A1'},{target:book.targets[0].id,cell:'A40'}]){try{await OfficeFiles.sign(bytes(source.xlsx),'xlsx',[{...p,...item}])}catch{rejected++}}
      return {doc:encode(signedDoc),book:encode(signedBook),rejected};
    },{source,stamp});
    assert.equal(result.rejected,6);fs.writeFileSync('test-results/signed.docx',Buffer.from(result.doc,'base64'));fs.writeFileSync('test-results/signed.xlsx',Buffer.from(result.book,'base64'));
    execFileSync('python',['tests/verify.py'],{stdio:'inherit'});checks+=7;
    const signedAlready=fs.readFileSync('test-results/signed-already.docx').toString('base64');assert.match(await page.evaluate(async value=>{try{await OfficeFiles.inspect(Uint8Array.from(atob(value),c=>c.charCodeAt(0)),'docx');return 'not rejected'}catch(e){return e.message}},signedAlready),/signature numérique/);checks++;
    // Outlook réel simulé uniquement à sa frontière : moteurs PDF/OOXML réels.
    await page.route('https://appsforoffice.microsoft.com/**',route=>route.fulfill({contentType:'text/javascript',body:`window.Office={onReady:()=>Promise.resolve(),AsyncResultStatus:{Succeeded:'succeeded'},MailboxEnums:{AttachmentContentFormat:{Base64:'base64'},AttachmentType:{Base64:'base64'}},context:{requirements:{isSetSupported:()=>true},mailbox:{item:{attachments:[{id:'docx',name:'Document.docx'},{id:'xlsx',name:'Budget.xlsx'},{id:'pdf',name:'Courrier.pdf'}],getAttachmentContentAsync(id,cb){setTimeout(()=>cb({status:'succeeded',value:{format:'base64',content:${JSON.stringify(source)}[id]}}),id==='docx'?80:0)},displayReplyFormAsync(data,cb){window.lastReply=data;cb({status:'succeeded'})}}}}};`}));
    await page.route('https://cdn.jsdelivr.net/**',route=>{const url=route.request().url();let local;if(url.includes('pdf-lib'))local='node_modules/pdf-lib/dist/pdf-lib.min.js';else if(url.includes('pdf.worker'))local='node_modules/pdfjs-dist/build/pdf.worker.min.js';else local='node_modules/pdfjs-dist/build/pdf.min.js';return route.fulfill({contentType:'text/javascript',body:fs.readFileSync(local)});});
    await page.goto(base);await page.waitForFunction(()=>document.getElementById('status').textContent.includes('Document.docx prêt')).catch(async e=>{console.error('Initial state:',await page.locator('#status').textContent(),errors);throw e});
    await page.locator('#stampFile').setInputFiles('test-results/signature.png');await page.waitForFunction(()=>document.getElementById('stampName').textContent==='signature.png');
    await page.locator('#officeTarget').selectOption({label:await page.locator('#officeTarget option').filter({hasText:'Signature ici'}).textContent()});
    await page.locator('#addOfficePlacement').click();await page.waitForFunction(()=>document.querySelectorAll('#officePlacements button').length===1);
    await page.locator('#reply').click();await page.waitForFunction(()=>window.lastReply);assert.equal(await page.evaluate(()=>lastReply.attachments[0].name),'Document - signé.docx');checks++;
    await page.locator('input[name=attachment]').nth(1).check();await page.waitForFunction(()=>document.getElementById('status').textContent.includes('Budget.xlsx prêt'));
    await page.locator('#addOfficePlacement').click();await page.waitForFunction(()=>document.querySelectorAll('#officePlacements button').length===1);
    await page.evaluate(()=>window.lastReply=null);await page.locator('#reply').click();await page.waitForFunction(()=>window.lastReply);assert.deepEqual(await page.evaluate(()=>lastReply.attachments.map(a=>a.name)),['Document - signé.docx','Budget - signé.xlsx']);checks++;
    assert.equal(await page.locator('header').evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(7, 90, 156)');
    await page.screenshot({path:'test-results/excel-ui.png',fullPage:true});
    await page.locator('input[name=attachment]').nth(2).check();await page.waitForFunction(()=>document.getElementById('status').textContent.includes('Courrier.pdf prêt'));
    await page.locator('#addPlacement').click();await page.evaluate(()=>window.lastReply=null);await page.locator('#reply').click();await page.waitForFunction(()=>window.lastReply);
    const pdfResult=await page.evaluate(()=>lastReply.attachments.find(a=>a.name.endsWith('.pdf')).base64file);assert.equal((await PDFDocument.load(Buffer.from(pdfResult,'base64'))).getPageCount(),1);checks++;
    await page.locator('input[name=attachment]').nth(1).check();await page.waitForFunction(()=>document.getElementById('status').textContent.includes('Budget.xlsx prêt'));
    await page.evaluate(pdf=>{OneDrivePDF.convert=async()=>({bytes:Uint8Array.from(atob(pdf),c=>c.charCodeAt(0)),warning:''})},source.pdf);
    await page.locator('#outputFormat').selectOption('pdf');await page.locator('#previewOffice').click();await page.waitForFunction(()=>document.getElementById('status').textContent.includes('Aperçu du document signé'));
    assert.equal(await page.locator('#stamp').isVisible(),false);
    await page.evaluate(()=>window.lastReply=null);await page.locator('#reply').click();await page.waitForFunction(()=>window.lastReply);
    assert.deepEqual(await page.evaluate(()=>lastReply.attachments.map(a=>a.name)),['Document - signé.docx','Budget - signé.pdf','Courrier - signé.pdf']);checks++;
    await page.locator('#officePlacements button').click();assert.equal(await page.locator('#pdfViewer').isVisible(),false);checks++;
    assert.deepEqual(errors,[]);
    // Connexion et Graph simulés : contrôle du nettoyage et des jetons.
    const graph=await browser.newPage();await graph.goto(base+'/empty');
    await graph.evaluate(()=>{window.M365_CONFIG={clientId:'11111111-1111-1111-1111-111111111111',tenantId:'22222222-2222-2222-2222-222222222222'};window.Office={AsyncResultStatus:{Succeeded:'ok'},EventType:{DialogMessageReceived:'message',DialogEventReceived:'event'},context:{ui:{displayDialogAsync(url,opts,cb){const channel=new URL(url).searchParams.get('channel');cb({status:'ok',value:{close(){},addEventHandler(type,handler){if(type==='message')setTimeout(()=>handler({origin:location.origin,message:JSON.stringify({type:'authenticated',channel,accessToken:'test-token',accountId:'test-account',username:'test@example.invalid',expiresAt:Date.now()+3600000})}),0)}}})}}}}});
    await graph.addScriptTag({url:base+'/onedrive.js'});await graph.evaluate(()=>OneDrivePDF.connect());
    let failConversion=false,failDelete=false,deleted=0,uploaded=0,downloadAuth;
    await graph.route('https://graph.microsoft.com/**',route=>{
      const req=route.request();assert.equal(req.headers().authorization,'Bearer test-token');
      const cors={'access-control-allow-origin':'*'};
      if(req.method()==='PUT'){uploaded++;return route.fulfill({headers:cors,json:{id:'temporary-item'}})}
      if(req.method()==='DELETE'){deleted++;return route.fulfill({headers:cors,status:failDelete?503:204})}
      return route.fulfill({headers:cors,json:{responses:[{id:'pdf',status:failConversion?403:302,headers:{Location:'https://example.sharepoint.com/converted.pdf'}}]}});
    });
    await graph.route('https://example.sharepoint.com/**',route=>{downloadAuth=route.request().headers().authorization;return route.fulfill({headers:{'access-control-allow-origin':'*'},contentType:'application/pdf',body:Buffer.from(pdfBytes)})});
    let converted=await graph.evaluate(async value=>{const r=await OneDrivePDF.convert(Uint8Array.from(atob(value),c=>c.charCodeAt(0)),'docx');return {size:r.bytes.length,warning:r.warning}},source.docx);assert.equal(converted.size,pdfBytes.length);assert.equal(downloadAuth,undefined);assert.equal(deleted,1);checks++;
    failConversion=true;assert.match(await graph.evaluate(async()=>{try{await OneDrivePDF.convert(new Uint8Array(10),'docx')}catch(e){return e.message}}),/refusé/);assert.equal(deleted,2);checks++;
    failConversion=false;failDelete=true;converted=await graph.evaluate(async()=>OneDrivePDF.convert(new Uint8Array(10),'xlsx'));assert.match(converted.warning,/copie temporaire/);failDelete=false;assert.equal(await graph.evaluate(()=>OneDrivePDF.retryCleanup()),0);checks++;
    const before=uploaded;assert.match(await graph.evaluate(async()=>{try{await OneDrivePDF.convert(new Uint8Array(2*1024*1024+1),'xlsx')}catch(e){return e.message}}),/2 Mo/);assert.equal(uploaded,before);checks++;
    await graph.evaluate(()=>{localStorage.setItem('outlook-signature-onedrive-cleanup-v1',JSON.stringify([{accountId:'test-account',name:'important.docx'}]))});const count=deleted;await graph.evaluate(()=>OneDrivePDF.retryCleanup());assert.equal(deleted,count);checks++;
    console.log(`${checks} integration checks passed. Real Microsoft sign-in/conversion still requires tenant configuration.`);
  }finally{await browser?.close();server.close()}
})().catch(error=>{console.error(error);process.exitCode=1});
