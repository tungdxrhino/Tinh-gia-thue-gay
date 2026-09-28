(function(global,document){
'use strict';
const W_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS='http://www.w3.org/XML/1998/namespace';
const MIME_DOCX='application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const MIME_XLSX='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const TEMPLATE={
  contract:'asset/templates/Hop_dong_cho_thue_gay_CLB_template.docx',
  handoverCues:'asset/templates/Bien_ban_ban_giao_gay_CLB_template.docx',
  handoverAssets:'asset/templates/Bien_ban_ban_giao_tai_san_template.docx'
};
const PREVIEW={
  contract:['asset/previews/contract/page-1.png','asset/previews/contract/page-2.png','asset/previews/contract/page-3.png'],
  handoverCues:['asset/previews/handover-cues/page-1.png','asset/previews/handover-cues/page-2.png'],
  handoverAssets:['asset/previews/handover-assets/page-1.png','asset/previews/handover-assets/page-2.png']
};
const DOC_TITLES={contract:'Hợp đồng cho thuê gậy Billiards',handoverCues:'Biên bản bàn giao gậy Billiards',handoverAssets:'Biên bản bàn giao tài sản'};
const PAGE_W=1414,PAGE_H=2000;
function num(v,d=0){const n=Number(v);return Number.isFinite(n)?n:d}
function int(v){return Math.max(0,Math.round(num(v,0)))}
function money(v){return new Intl.NumberFormat('vi-VN').format(Math.round(num(v,0)))}
function safeName(v){return String(v||'CLB').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/gi,'d').replace(/[\\/:*?"<>|]+/g,' ').replace(/\s+/g,' ').trim().replace(/ /g,'_').slice(0,80)||'CLB'}
function pad(n){return String(n).padStart(2,'0')}
function parseISO(v){const m=String(v||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return null;const d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]));return Number.isNaN(d.getTime())?null:d}
function dateVi(v){const d=parseISO(v);return d?`${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()}`:String(v||'')}
function dateLong(v){const d=parseISO(v);return d?`ngày ${pad(d.getDate())} tháng ${pad(d.getMonth()+1)} năm ${d.getFullYear()}`:String(v||'')}
function dateLongNoPrefix(v){const d=parseISO(v);return d?`${pad(d.getDate())} tháng ${pad(d.getMonth()+1)} năm ${d.getFullYear()}`:String(v||'')}
function addMonths(v,months){const d=parseISO(v);if(!d)return '';const day=d.getDate();d.setDate(1);d.setMonth(d.getMonth()+Math.max(0,int(months)));const last=new Date(d.getFullYear(),d.getMonth()+1,0).getDate();d.setDate(Math.min(day,last));return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`}
function wordsSmall(n){const a=['không','một','hai','ba','bốn','năm','sáu','bảy','tám','chín'];n=int(n);if(n<10)return a[n];if(n<20)return n===10?'mười':`mười ${n%10===5?'lăm':a[n%10]}`;if(n<100){const t=Math.floor(n/10),u=n%10;return `${a[t]} mươi${u?` ${u===1?'mốt':u===5?'lăm':a[u]}`:''}`};return String(n)}
function readThree(n,full){const a=['không','một','hai','ba','bốn','năm','sáu','bảy','tám','chín'],h=Math.floor(n/100),t=Math.floor(n%100/10),u=n%10,o=[];if(h||full)o.push(a[h],'trăm');if(t>1){o.push(a[t],'mươi');if(u)o.push(u===1?'mốt':u===4?'tư':u===5?'lăm':a[u])}else if(t===1){o.push('mười');if(u)o.push(u===5?'lăm':a[u])}else if(u){if(h||full)o.push('lẻ');o.push(a[u])}return o.join(' ')}
function numberWords(v){let n=Math.round(num(v));if(n===0)return 'Không đồng';const scales=['','nghìn','triệu','tỷ','nghìn tỷ','triệu tỷ'],g=[];while(n>0){g.push(n%1000);n=Math.floor(n/1000)}const out=[];for(let i=g.length-1;i>=0;i--){if(!g[i])continue;const s=readThree(g[i],i<g.length-1&&g[i]<10);out.push(s+(scales[i]?' '+scales[i]:''))}const s=out.join(' ').replace(/\s+/g,' ').trim();return s.charAt(0).toUpperCase()+s.slice(1)+' đồng'}
function qtyWords(v){const n=int(v),s=wordsSmall(n);return s.charAt(0).toUpperCase()+s.slice(1)+' cây'}
function cueTypes(allocation){const m=(allocation?.models||[]).filter(x=>num(x.qty)>0).map(x=>x.model);return m.length?m.join(' / '):'R68 / R88'}
function allocationNote(allocation){return (allocation?.models||[]).filter(x=>num(x.qty)>0).map(x=>`${x.model}: ${int(x.qty)} cây`).join('; ')}
function allocationCompact(allocation){return (allocation?.models||[]).filter(x=>num(x.qty)>0).map(x=>`${x.model}:${int(x.qty)}`).join(' / ')}
function planPaid(plan){const explicit=num(plan?.paid||plan?.term,0);if(explicit>0)return int(explicit);const n=num(plan?.n,0),rent=num(plan?.rent,0),unit=num(plan?.unit,0);if(n>0&&rent>0&&unit>0)return Math.max(1,Math.round(rent/(n*unit)));return 0}
function baseCommon(ctx){const info=ctx.info||{},plan=ctx.plan||{},alloc=ctx.allocation||{},extra=ctx.extra||{};return {info,plan,alloc,extra,types:cueTypes(alloc),paid:int(plan.n),backup:int(plan.backup),total:int(plan.total),unit:num(plan.unit),rent:num(plan.rent),deposit:num(plan.deposit),initial:num(plan.initial),paidMonths:planPaid(plan),useMonths:int(plan.use),gift:int(plan.gift)}}
function contractValues(ctx){const c=baseCommon(ctx),e=c.extra;return {
  contract_no:e.contractNo||'',contract_date_long:dateLong(e.contractDate),contract_place:e.contractPlace||'',club_name:c.info.club||'',club_address:c.info.address||'',tax_id:e.taxId||'',customer_phone:c.info.customerPhone||'',customer_email:c.info.customerEmail||'',representative_name:e.representativeName||c.info.customer||'',representative_title:e.representativeTitle||'Chủ hộ kinh doanh',cue_types:c.types,paid_qty:String(c.paid),total_use_months:String(c.useMonths),rental_start_date:dateVi(e.rentalStartDate),rental_end_date:dateVi(e.rentalEndDate||addMonths(e.rentalStartDate,c.useMonths)),prepaid_box:c.paidMonths>=3?'[x]':'[ ]',flexible_box:c.paidMonths<=2?'[x]':'[ ]',backup_qty:String(c.backup),total_qty:String(c.total),unit_price:money(c.unit),rent_total:money(c.rent),deposit_amount:money(c.deposit),grand_total:money(c.initial),grand_total_words:numberWords(c.initial)
}}
function handoverCuesValues(ctx){const c=baseCommon(ctx),e=c.extra,dt=e.handoverTime||'';return {handover_no:e.handoverNo||'',contract_no:e.contractNo||'',contract_no_2:e.contractNo||'',contract_date_long:dateLongNoPrefix(e.contractDate),club_name:c.info.club||'',club_name_2:c.info.club||'',handover_time:dt?dt.replace(':',' giờ ')+' phút':'',handover_date_long:dateLong(e.handoverDate),handover_place:e.handoverPlace||c.info.address||'',representative_name:e.representativeName||c.info.customer||'',representative_title:e.representativeTitle||'Chủ hộ kinh doanh',customer_phone:c.info.customerPhone||'',paid_cue_types:c.types,paid_qty:String(c.paid),paid_allocation_note:allocationCompact(c.alloc),backup_cue_types:c.types,backup_qty:String(c.backup),backup_note:'',total_qty:String(c.total),total_qty_2:String(c.total),total_qty_words:qtyWords(c.total),accessories:e.accessories||'',handover_note:e.handoverNote||''}}
function handoverAssetsValues(ctx){const c=baseCommon(ctx),e=c.extra;return {contract_no:e.contractNo||'',contract_no_2:e.contractNo||'',contract_date_long:dateLongNoPrefix(e.contractDate),handover_date_long:dateLong(e.handoverDate),handover_place:e.handoverPlace||c.info.address||'',club_name:c.info.club||'',club_address:c.info.address||'',representative_name:e.representativeName||c.info.customer||'',representative_title:e.representativeTitle||'Chủ hộ kinh doanh',cue_types:c.types,paid_qty:String(c.paid),backup_qty:String(c.backup),total_qty:String(c.total)}}
function buildValues(kind,ctx){return kind==='contract'?contractValues(ctx):kind==='handoverCues'?handoverCuesValues(ctx):handoverAssetsValues(ctx)}
function findStart(doc,name){return Array.from(doc.getElementsByTagNameNS(W_NS,'bookmarkStart')).find(n=>(n.getAttributeNS(W_NS,'name')||n.getAttribute('w:name'))===name)}
function setBookmark(doc,name,value){const start=findStart(doc,name);if(!start)return false;const id=start.getAttributeNS(W_NS,'id')||start.getAttribute('w:id');let n=start.nextSibling,run=null,end=null;while(n){if(n.nodeType===1&&n.namespaceURI===W_NS&&n.localName==='bookmarkEnd'&&(n.getAttributeNS(W_NS,'id')||n.getAttribute('w:id'))===id){end=n;break}if(!run&&n.nodeType===1&&n.namespaceURI===W_NS&&n.localName==='r')run=n;n=n.nextSibling}if(!run){run=doc.createElementNS(W_NS,'w:r');start.parentNode.insertBefore(run,end||start.nextSibling)}let t=run.getElementsByTagNameNS(W_NS,'t')[0];if(!t){t=doc.createElementNS(W_NS,'w:t');run.appendChild(t)}t.textContent=String(value??'');if(/^\s|\s$/.test(String(value??'')))t.setAttributeNS(XML_NS,'xml:space','preserve');return true}
async function createWordBlob(kind,ctx){if(!global.JSZip)throw new Error('Thiếu JSZip.');const r=await fetch(TEMPLATE[kind],{cache:'no-store'});if(!r.ok)throw new Error('Không tải được template Word.');const zip=await global.JSZip.loadAsync(await r.arrayBuffer()),f=zip.file('word/document.xml');if(!f)throw new Error('Template Word không hợp lệ.');const xml=await f.async('string'),doc=new DOMParser().parseFromString(xml,'application/xml');const vals=buildValues(kind,ctx);Object.entries(vals).forEach(([k,v])=>setBookmark(doc,k,v));zip.file('word/document.xml',new XMLSerializer().serializeToString(doc));return zip.generateAsync({type:'blob',mimeType:MIME_DOCX,compression:'DEFLATE'})}
function download(blob,name){const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),1800)}
async function exportWord(kind,ctx){const b=await createWordBlob(kind,ctx),n=`${kind==='contract'?'Hop_dong':kind==='handoverCues'?'Bien_ban_ban_giao_gay':'Bien_ban_ban_giao_tai_san'}_Rhino_${safeName(ctx.info?.club)}_${String(ctx.extra?.contractDate||ctx.extra?.handoverDate||new Date().toISOString().slice(0,10))}.docx`;download(b,n);return n}

// Preview/PDF overlay maps. Coordinates are in the 1414x2000 rendered-template coordinate space.
const OVERLAYS={
 contract:[
  {p:1,k:'contract_no',x:568,y:322,w:260},
  {p:1,k:'contract_date_long',x:290,y:583,w:150},
  {p:1,k:'contract_place',x:389,y:583,w:500},
  {p:1,k:'club_name',x:745,y:997,w:300},
  {p:1,k:'club_address',x:419,y:1069,w:500},
  {p:1,k:'tax_id',x:492,y:1104,w:150},
  {p:1,k:'customer_phone',x:399,y:1138,w:150},
  {p:1,k:'customer_email',x:556,y:1138,w:500},
  {p:1,k:'representative_name',x:558,y:1173,w:300},
  {p:1,k:'representative_title',x:751,y:1173,w:300},
  {p:1,k:'cue_types',x:437,y:1472,w:300},
  {p:1,k:'paid_qty',x:544,y:1506,w:150},
  {p:1,k:'total_use_months',x:422,y:1541,w:150},
  {p:1,k:'rental_start_date',x:665,y:1541,w:150},
  {p:1,k:'rental_end_date',x:836,y:1541,w:150},
  {p:1,k:'prepaid_box',x:342,y:1613,w:150},
  {p:1,k:'flexible_box',x:342,y:1716,w:150},
  {p:2,k:'backup_qty',x:702,y:172,w:150},
  {p:2,k:'total_qty',x:760,y:206,w:150},
  {p:2,k:'unit_price',x:433,y:355,w:150},
  {p:2,k:'rent_total',x:617,y:389,w:150},
  {p:2,k:'deposit_amount',x:778,y:424,w:150},
  {p:2,k:'grand_total',x:724,y:458,w:150},
  {p:2,k:'grand_total_words',x:980,y:458,w:394}
 ],
 handoverCues:[
  {p:1,k:'handover_no',x:226,y:310,w:260},
  {p:1,k:'contract_no',x:771,y:379,w:260},
  {p:1,k:'contract_date_long',x:924,y:379,w:150},
  {p:1,k:'club_name',x:640,y:413,w:300},
  {p:1,k:'handover_time',x:380,y:447,w:150},
  {p:1,k:'handover_date_long',x:444,y:447,w:150},
  {p:1,k:'handover_place',x:633,y:447,w:500},
  {p:1,k:'club_name_2',x:749,y:654,w:300},
  {p:1,k:'representative_name',x:558,y:723,w:300},
  {p:1,k:'representative_title',x:751,y:723,w:300},
  {p:1,k:'customer_phone',x:437,y:757,w:150},
  {p:1,k:'paid_cue_types',x:309,y:1188,w:300},
  {p:1,k:'paid_qty',x:740,y:1188,w:150},
  {p:1,k:'paid_allocation_note',x:1118,y:1188,w:256},
  {p:1,k:'backup_qty',x:740,y:1365,w:150},
  {p:1,k:'backup_note',x:1118,y:1365,w:256},
  {p:1,k:'backup_cue_types',x:309,y:1365,w:300},
  {p:1,k:'total_qty_2',x:625,y:1733,w:150},
  {p:1,k:'total_qty_words',x:393,y:1767,w:500},
  {p:1,k:'total_qty',x:740,y:1630,w:150},
  {p:2,k:'accessories',x:583,y:172,w:150},
  {p:2,k:'handover_note',x:257,y:206,w:500},
  {p:2,k:'contract_no_2',x:1138,y:551,w:236}
 ],
 handoverAssets:[
  {p:1,k:'contract_no',x:771,y:413,w:260},
  {p:1,k:'contract_date_long',x:924,y:413,w:150},
  {p:1,k:'handover_date_long',x:290,y:447,w:150},
  {p:1,k:'handover_place',x:389,y:447,w:500},
  {p:1,k:'club_name',x:887,y:689,w:300},
  {p:1,k:'club_address',x:419,y:757,w:500},
  {p:1,k:'representative_name',x:558,y:792,w:300},
  {p:1,k:'representative_title',x:753,y:792,w:300},
  {p:1,k:'cue_types',x:643,y:999,w:300},
  {p:1,k:'paid_qty',x:544,y:1033,w:150},
  {p:1,k:'backup_qty',x:702,y:1067,w:150},
  {p:1,k:'total_qty',x:760,y:1102,w:150},
  {p:1,k:'contract_no_2',x:397,y:1756,w:260}
 ]
};
function htmlEsc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function previewHtml(kind,ctx){const vals=buildValues(kind,ctx),pages=PREVIEW[kind]||[];return pages.map((src,i)=>{const ovs=(OVERLAYS[kind]||[]).filter(o=>o.p===i+1).map(o=>`<button type="button" class="doc-preview-field" data-edit-field="${htmlEsc(o.k)}" style="left:${(o.x/PAGE_W*100).toFixed(3)}%;top:${(o.y/PAGE_H*100).toFixed(3)}%;width:${(o.w/PAGE_W*100).toFixed(3)}%;font-size:clamp(7px,1.25vw,13px)">${htmlEsc(vals[o.k]||'')}</button>`).join('');return `<div class="doc-preview-page" data-page="${i+1}"><img src="${src}" alt="Trang ${i+1}"/>${ovs}</div>`}).join('')}
async function loadImage(src){return new Promise((res,rej)=>{const im=new Image();im.onload=()=>res(im);im.onerror=rej;im.src=src})}
function drawWrap(ctx,text,x,y,w,fs=21,lh=1.25){ctx.font=`600 ${fs}px Arial`;ctx.fillStyle='#111';ctx.textBaseline='top';const words=String(text||'').split(/\s+/);let line='',yy=y;for(const word of words){const test=line?line+' '+word:word;if(ctx.measureText(test).width>w&&line){ctx.fillText(line,x,yy);yy+=fs*lh;line=word}else line=test}if(line)ctx.fillText(line,x,yy)}
async function renderPageCanvas(kind,ctx,pageIndex){const img=await loadImage(PREVIEW[kind][pageIndex]);const c=document.createElement('canvas');c.width=PAGE_W;c.height=PAGE_H;const x=c.getContext('2d');x.drawImage(img,0,0,PAGE_W,PAGE_H);const vals=buildValues(kind,ctx);(OVERLAYS[kind]||[]).filter(o=>o.p===pageIndex+1).forEach(o=>{const val=vals[o.k]||'';if(!val)return;drawWrap(x,val,o.x,o.y,o.w,o.fs||21)});return c}
function ascii(s){return new TextEncoder().encode(s)}
function concat(arr){let len=0;arr.forEach(a=>len+=a.length);const out=new Uint8Array(len);let o=0;arr.forEach(a=>{out.set(a,o);o+=a.length});return out}
async function canvasJpegBytes(c){const blob=await new Promise(r=>c.toBlob(r,'image/jpeg',0.96));return new Uint8Array(await blob.arrayBuffer())}
async function createPdfBlob(kind,ctx){const pages=[];for(let i=0;i<PREVIEW[kind].length;i++)pages.push(await renderPageCanvas(kind,ctx,i));const objects=[],pageIds=[],imageIds=[],contentIds=[];let id=3;for(let i=0;i<pages.length;i++){pageIds.push(id++);imageIds.push(id++);contentIds.push(id++)}const catalogId=1,pagesId=2;objects[catalogId]=[ascii(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`)];objects[pagesId]=[ascii(`<< /Type /Pages /Kids [${pageIds.map(x=>x+' 0 R').join(' ')}] /Count ${pages.length} >>`)];for(let i=0;i<pages.length;i++){const jpg=await canvasJpegBytes(pages[i]),pw=595.28,ph=841.89;objects[pageIds[i]]=[ascii(`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${pw} ${ph}] /Resources << /XObject << /Im0 ${imageIds[i]} 0 R >> >> /Contents ${contentIds[i]} 0 R >>`)];objects[imageIds[i]]=[ascii(`<< /Type /XObject /Subtype /Image /Width ${PAGE_W} /Height ${PAGE_H} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpg.length} >>\nstream\n`),jpg,ascii('\nendstream')];const content=ascii(`q\n${pw} 0 0 ${ph} 0 0 cm\n/Im0 Do\nQ\n`);objects[contentIds[i]]=[ascii(`<< /Length ${content.length} >>\nstream\n`),content,ascii('endstream')]}
const max=id-1,parts=[ascii('%PDF-1.4\n%RhinoDocs\n')],offs=[0];let pos=parts[0].length;for(let i=1;i<=max;i++){offs[i]=pos;const h=ascii(`${i} 0 obj\n`),t=ascii('\nendobj\n');parts.push(h,...objects[i],t);pos+=h.length+objects[i].reduce((s,p)=>s+p.length,0)+t.length}const xp=pos;let xr=`xref\n0 ${max+1}\n0000000000 65535 f \n`;for(let i=1;i<=max;i++)xr+=`${String(offs[i]).padStart(10,'0')} 00000 n \n`;xr+=`trailer\n<< /Size ${max+1} /Root 1 0 R >>\nstartxref\n${xp}\n%%EOF\n`;parts.push(ascii(xr));return new Blob([concat(parts)],{type:'application/pdf'})}
async function exportPdf(kind,ctx){const b=await createPdfBlob(kind,ctx),n=`${kind==='contract'?'Hop_dong':kind==='handoverCues'?'Bien_ban_ban_giao_gay':'Bien_ban_ban_giao_tai_san'}_Rhino_${safeName(ctx.info?.club)}_${new Date().toISOString().slice(0,10)}.pdf`;download(b,n);return n}

function sx(v){return String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
function colName(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
function sheetXml(kind,ctx){const vals=buildValues(kind,ctx),rows=[[DOC_TITLES[kind],''],['Trường','Giá trị'],...Object.entries(vals).map(([k,v])=>[k,v])];let xml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="70" customWidth="1"/></cols><sheetData>';rows.forEach((r,ri)=>{xml+=`<row r="${ri+1}">`;r.forEach((v,ci)=>{const ref=colName(ci+1)+(ri+1);xml+=`<c r="${ref}" t="inlineStr"><is><t>${sx(v)}</t></is></c>`});xml+='</row>'});xml+='</sheetData></worksheet>';return xml}
async function createExcelBlob(kind,ctx){if(!global.JSZip)throw new Error('Thiếu JSZip.');const z=new global.JSZip();z.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');z.folder('_rels').file('.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');z.folder('xl').file('workbook.xml','<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="TÀI LIỆU" sheetId="1" r:id="rId1"/></sheets></workbook>');z.folder('xl').folder('_rels').file('workbook.xml.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');z.folder('xl').folder('worksheets').file('sheet1.xml',sheetXml(kind,ctx));return z.generateAsync({type:'blob',mimeType:MIME_XLSX,compression:'DEFLATE'})}
async function exportExcel(kind,ctx){const b=await createExcelBlob(kind,ctx),n=`${kind}_Rhino_${safeName(ctx.info?.club)}_${new Date().toISOString().slice(0,10)}.xlsx`;download(b,n);return n}

global.RhinoDocuments={version:'1.2.0',TEMPLATE,PREVIEW,DOC_TITLES,buildValues,previewHtml,createWordBlob,exportWord,createPdfBlob,exportPdf,createExcelBlob,exportExcel,addMonths,dateVi,dateLong,dateLongNoPrefix,numberWords};
})(window,document);
