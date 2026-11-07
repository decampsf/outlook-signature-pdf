import JSZip from 'jszip';

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
};
const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
};
const all = (doc, ns, name) => Array.from(doc.getElementsByTagNameNS(ns, name));
const xml = text => {
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('Déclarations XML non prises en charge.');
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Document XML invalide.');
  return doc;
};
const serialize = doc => new XMLSerializer().serializeToString(doc);
const esc = value => String(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
async function read(zip, path) {
  const entry = zip.file(path);
  if (!entry) throw new Error('Élément du document introuvable : ' + path);
  const text = await entry.async('string');
  if (text.length > 15 * 1024 * 1024) throw new Error('Document trop complexe pour ce complément.');
  return xml(text);
}
function resolvePath(part, target) {
  if (!target || /[\\?#]/.test(target) || /^[a-z]+:/i.test(target)) throw new Error('Relation Office non prise en charge.');
  const stack = target.startsWith('/') ? [] : part.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '..') { if (!stack.length) throw new Error('Chemin Office invalide.'); stack.pop(); }
    else if (segment && segment !== '.') stack.push(segment);
  }
  return stack.join('/');
}
function relativePath(from, to) {
  const a = from.split('/').slice(0, -1), b = to.split('/');
  while (a.length && a[0] === b[0]) { a.shift(); b.shift(); }
  return '../'.repeat(a.length) + b.join('/');
}
const relPath = path => path.replace(/([^/]+)$/, '_rels/$1.rels');
async function relationships(zip, part) {
  const path = relPath(part);
  return {path, doc: zip.file(path) ? await read(zip, path) : xml(`<Relationships xmlns="${NS.rel}"/>`)};
}
function addRel(doc, type, target) {
  const existing = new Set(all(doc, NS.rel, 'Relationship').map(r => r.getAttribute('Id')));
  let n = 1; while (existing.has('rIdSignature' + n)) n++;
  const id = 'rIdSignature' + n;
  const node = doc.createElementNS(NS.rel, 'Relationship');
  node.setAttribute('Id', id); node.setAttribute('Type', NS.r + '/' + type); node.setAttribute('Target', target);
  doc.documentElement.append(node); return id;
}
function relation(doc, id, expected) {
  const node = all(doc, NS.rel, 'Relationship').find(r => r.getAttribute('Id') === id);
  if (!node || node.getAttribute('TargetMode') === 'External' || node.getAttribute('Type') !== NS.r + '/' + expected) throw new Error('Relation Office incompatible.');
  return node.getAttribute('Target');
}
function contentType(doc, path, type) {
  const prior = all(doc, NS.ct, 'Override').find(n => n.getAttribute('PartName') === '/' + path);
  if (prior) { if (prior.getAttribute('ContentType') !== type) throw new Error('Type Office incompatible.'); return; }
  const node = doc.createElementNS(NS.ct, 'Override'); node.setAttribute('PartName', '/' + path); node.setAttribute('ContentType', type); doc.documentElement.append(node);
}
function pngType(doc) {
  const prior = all(doc, NS.ct, 'Default').find(n => n.getAttribute('Extension')?.toLowerCase() === 'png');
  if (prior && prior.getAttribute('ContentType') !== 'image/png') throw new Error('Type PNG incompatible dans le fichier.');
  if (!prior) { const n = doc.createElementNS(NS.ct, 'Default'); n.setAttribute('Extension', 'png'); n.setAttribute('ContentType', 'image/png'); doc.documentElement.append(n); }
}
function newPart(zip, prefix, ext) { let n = 1; while (zip.file(prefix + n + ext)) n++; return prefix + n + ext; }
function paragraphs(doc) {
  return all(doc, NS.w, 'p').filter(p => {
    for (let n = p.parentElement; n; n = n.parentElement) {
      if (n.namespaceURI === NS.w && ['del','txbxContent','sdt'].includes(n.localName)) return false;
    }
    return true;
  });
}
export async function inspect(bytes, kind) {
  if (!MIME[kind]) throw new Error('Utilisez un fichier .docx ou .xlsx.');
  if (bytes.length > 25 * 1024 * 1024) throw new Error('Le fichier dépasse 25 Mo.');
  let zip;
  try { zip = await JSZip.loadAsync(bytes); } catch { throw new Error('Fichier Office illisible ou chiffré. Ouvrez-le dans Word ou Excel.'); }
  let expanded = 0;
  for (const entry of Object.values(zip.files)) {
    expanded += entry._data?.uncompressedSize || 0;
    if (expanded > 60 * 1024 * 1024) throw new Error('Document décompressé trop volumineux.');
  }
  if (Object.keys(zip.files).some(p => /^_xmlsignatures\//i.test(p))) throw new Error('Ce fichier porte déjà une signature numérique. Il ne sera pas modifié.');
  const types = await read(zip, '[Content_Types].xml');
  const path = kind === 'docx' ? 'word/document.xml' : 'xl/workbook.xml';
  if (!all(types, NS.ct, 'Override').some(n => n.getAttribute('PartName') === '/' + path && n.getAttribute('ContentType') === MIME[kind])) throw new Error('Format Office incompatible (macros ou format Strict non pris en charge).');
  const doc = await read(zip, path);
  const expectedNS = kind === 'docx' ? NS.w : NS.s;
  if (doc.documentElement.namespaceURI !== expectedNS) throw new Error('Enregistrez une copie au format Word ou Excel standard.');
  if (kind === 'docx') {
    if (zip.file('word/settings.xml')) {
      const settings = await read(zip, 'word/settings.xml');
      if (all(settings, NS.w, 'documentProtection').some(n => ['1','true','on'].includes(n.getAttributeNS(NS.w, 'enforcement')))) throw new Error('Le document Word est protégé.');
      if (all(settings, NS.w, 'trackRevisions').some(n => !['0','false','off'].includes(n.getAttributeNS(NS.w, 'val')))) throw new Error('Désactivez le suivi des modifications dans une copie avant de signer.');
    }
    const targets = paragraphs(doc).map((p, index) => ({id: index, label: all(p, NS.w, 't').map(t => t.textContent).join('').trim().slice(0, 150) || '(paragraphe vide)'}));
    if (!targets.length) throw new Error('Aucun paragraphe accessible pour la signature.');
    return {kind, targets};
  }
  const rels = await relationships(zip, path);
  const sheets = [];
  for (const sheet of all(doc, NS.s, 'sheet')) {
    if (sheet.getAttribute('state') && sheet.getAttribute('state') !== 'visible') continue;
    const id = sheet.getAttributeNS(NS.r, 'id');
    const rel = all(rels.doc, NS.rel, 'Relationship').find(r => r.getAttribute('Id') === id);
    if (rel?.getAttribute('Type') !== NS.r + '/worksheet') continue;
    const sheetPath = resolvePath(path, relation(rels.doc, id, 'worksheet'));
    const sheetDoc = await read(zip, sheetPath);
    const protectedSheet = all(sheetDoc, NS.s, 'sheetProtection').length > 0;
    sheets.push({id: sheetPath, label: sheet.getAttribute('name'), protected: protectedSheet});
  }
  if (!sheets.length) throw new Error('Aucune feuille de calcul visible.');
  return {kind, targets: sheets};
}
export function cellAddress(text) {
  const m = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(String(text).trim().toUpperCase());
  if (!m) throw new Error('Indiquez une cellule, par exemple B12.');
  let col = 0; for (const c of m[1]) col = col * 26 + c.charCodeAt(0) - 64;
  const row = Number(m[2]);
  if (col > 16384 || row > 1048576) throw new Error('Cellule hors des limites Excel.');
  return {col: col - 1, row: row - 1, name: m[1] + row};
}
function importNode(doc, text) { return doc.importNode(xml(text).documentElement, true); }
function nextDrawingId(doc, ns, name) { return Math.max(0, ...all(doc, ns, name).map(n => Number(n.getAttribute('id')) || 0)) + 1; }
export async function sign(bytes, kind, placements) {
  if (!placements.length) throw new Error('Ajoutez au moins une signature.');
  const info = await inspect(bytes, kind);
  const zip = await JSZip.loadAsync(bytes);
  const types = await read(zip, '[Content_Types].xml'); pngType(types);
  let wordDrawingId=0;
  if(kind==='docx')for(const path of Object.keys(zip.files).filter(p=>/^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(p))){
    wordDrawingId=Math.max(wordDrawingId,nextDrawingId(await read(zip,path),NS.wp,'docPr')-1);
  }
  for (const p of placements) {
    if (!info.targets.some(t => String(t.id) === String(p.target))) throw new Error('Emplacement de signature invalide.');
    const widthCm = Number(p.widthCm);
    if (!Number.isFinite(widthCm) || widthCm < 2 || widthCm > 12 || !p.pixelWidth || !p.pixelHeight) throw new Error('Dimensions de signature invalides.');
    const cx = Math.round(widthCm * 360000), cy = Math.round(cx * p.pixelHeight / p.pixelWidth);
    const folder = kind === 'docx' ? 'word' : 'xl';
    const media = newPart(zip, folder + '/media/signature-outlook-', '.png');
    if (!/^data:image\/png;base64,/.test(p.image)) throw new Error('Signature PNG attendue.');
    zip.file(media, p.image.split(',')[1], {base64: true});
    if (kind === 'docx') {
      const path = 'word/document.xml', doc = await read(zip, path), rels = await relationships(zip, path);
      const rid = addRel(rels.doc, 'image', relativePath(path, media));
      const id = ++wordDrawingId;
      const paragraph = paragraphs(doc)[Number(p.target)];
      const run = importNode(doc, `<w:r xmlns:w="${NS.w}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:r="${NS.r}">
        ${all(paragraph, NS.w, 't').length ? '<w:br/>' : ''}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Signature Outlook ${id}" descr="${esc(p.caption || 'Signature')}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="${NS.pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Signature"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`);
      paragraph.append(run); zip.file(path, serialize(doc)); zip.file(rels.path, serialize(rels.doc));
    } else {
      const path = String(p.target), doc = await read(zip, path), rels = await relationships(zip, path);
      if (all(doc, NS.s, 'sheetProtection').length) throw new Error('La feuille Excel est protégée.');
      const cell = cellAddress(p.cell);
      if (all(doc, NS.s, 'row').some(n => Number(n.getAttribute('r')) === cell.row + 1 && ['1','true'].includes(n.getAttribute('hidden'))) || all(doc, NS.s, 'col').some(n => cell.col + 1 >= Number(n.getAttribute('min')) && cell.col + 1 <= Number(n.getAttribute('max')) && ['1','true'].includes(n.getAttribute('hidden')))) throw new Error('La cellule choisie est masquée.');
      let drawing = all(doc, NS.s, 'drawing')[0], drawingPath, drawingDoc;
      if (drawing) {
        drawingPath = resolvePath(path, relation(rels.doc, drawing.getAttributeNS(NS.r, 'id'), 'drawing'));
        drawingDoc = await read(zip, drawingPath);
        if (drawingDoc.documentElement.namespaceURI !== NS.xdr) throw new Error('Dessin Excel incompatible.');
      } else {
        drawingPath = newPart(zip, 'xl/drawings/signature-outlook-', '.xml');
        drawingDoc = xml(`<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"/>`);
        const id = addRel(rels.doc, 'drawing', relativePath(path, drawingPath));
        drawing = doc.createElementNS(NS.s, 'drawing'); drawing.setAttributeNS(NS.r, 'r:id', id);
        const later = new Set(['legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst']);
        const before = Array.from(doc.documentElement.children).find(n => n.namespaceURI === NS.s && later.has(n.localName));
        doc.documentElement.insertBefore(drawing, before || null);
        contentType(types, drawingPath, 'application/vnd.openxmlformats-officedocument.drawing+xml');
      }
      const dr = await relationships(zip, drawingPath), rid = addRel(dr.doc, 'image', relativePath(drawingPath, media));
      const id = nextDrawingId(drawingDoc, NS.xdr, 'cNvPr');
      const anchor = importNode(drawingDoc, `<xdr:oneCellAnchor xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><xdr:from><xdr:col>${cell.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${cell.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Signature Outlook ${id}" descr="${esc(p.caption || 'Signature')}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`);
      const extension=Array.from(drawingDoc.documentElement.children).find(n=>n.localName==='extLst');
      drawingDoc.documentElement.insertBefore(anchor,extension||null);
      zip.file(drawingPath, serialize(drawingDoc)); zip.file(dr.path, serialize(dr.doc));
      zip.file(path, serialize(doc)); zip.file(rels.path, serialize(rels.doc));
    }
  }
  zip.file('[Content_Types].xml', serialize(types));
  return zip.generateAsync({type: 'uint8array', compression: 'DEFLATE', compressionOptions: {level: 6}});
}
