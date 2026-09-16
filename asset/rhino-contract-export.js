/* Rhino Cue Platform - Rental contract DOCX exporter (V38)
 * Uses the existing JSZip dependency and the bookmarked Word template.
 * Keeps pricing/plan data external: this module never recalculates rental pricing.
 */
(function(global){
  'use strict';

  const W_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XML_NS='http://www.w3.org/XML/1998/namespace';
  const TEMPLATE_URL='asset/templates/Hop_dong_cho_thue_gay_Rhino_template_web.docx';

  const COMPANY_CONFIG={
    partyASignerName:'Phạm Duy Đông'
  };

  const RHINO_PRODUCT_MASTER={
    R68:{
      code:'R68',
      name:'Rhino 68',
      group:'Gậy CLB',
      unit:'Gậy',
      compensationPrice:500000,
      compensationNote:'Mất hoặc hư hỏng không thể phục hồi do nguyên nhân thuộc trách nhiệm của Bên B.'
    },
    R88:{
      code:'R88',
      name:'Rhino 88',
      group:'Gậy CLB',
      unit:'Gậy',
      compensationPrice:500000,
      compensationNote:'Mất hoặc hư hỏng không thể phục hồi do nguyên nhân thuộc trách nhiệm của Bên B.'
    }
  };

  function finiteNumber(v,fallback=0){const n=Number(v);return Number.isFinite(n)?n:fallback}
  function intNumber(v,fallback=0){const n=Math.round(finiteNumber(v,fallback));return Number.isFinite(n)?n:fallback}
  function formatVND(v){return new Intl.NumberFormat('vi-VN').format(Math.round(finiteNumber(v,0)))}
  function safeFilePart(v){
    return String(v||'CLB').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/gi,'d')
      .replace(/[\\/:*?"<>|]+/g,' ').replace(/\s+/g,' ').trim().replace(/ /g,'_').slice(0,80)||'CLB';
  }
  function isoToViDate(v){
    const s=String(v||'').trim();const m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?`${m[3]}/${m[2]}/${m[1]}`:s;
  }
  function yearFromDate(v){const m=String(v||'').match(/^(\d{4})-/);return m?m[1]:String(new Date().getFullYear())}
  function capitalizeFirst(v){const s=String(v||'').trim();return s?s.charAt(0).toLocaleUpperCase('vi')+s.slice(1):s}

  function readThreeDigits(num,full){
    const ones=['không','một','hai','ba','bốn','năm','sáu','bảy','tám','chín'];
    const h=Math.floor(num/100),t=Math.floor((num%100)/10),u=num%10;const out=[];
    if(h>0||full){out.push(ones[h],'trăm')}
    if(t>1){out.push(ones[t],'mươi');if(u===1)out.push('mốt');else if(u===4)out.push('tư');else if(u===5)out.push('lăm');else if(u>0)out.push(ones[u])}
    else if(t===1){out.push('mười');if(u===5)out.push('lăm');else if(u>0)out.push(ones[u])}
    else if(u>0){if(h>0||full)out.push('lẻ');out.push(ones[u])}
    return out.join(' ');
  }
  function numberToVietnameseWords(value){
    let n=Math.round(finiteNumber(value,0));if(n===0)return 'Không đồng';if(n<0)return 'Âm '+numberToVietnameseWords(-n).toLocaleLowerCase('vi');
    const scale=['','nghìn','triệu','tỷ','nghìn tỷ','triệu tỷ'];const groups=[];
    while(n>0){groups.push(n%1000);n=Math.floor(n/1000)}
    const parts=[];
    for(let i=groups.length-1;i>=0;i--){const g=groups[i];if(g===0)continue;const full=i<groups.length-1&&g<10;const words=readThreeDigits(g,full);if(words)parts.push(words+(scale[i]?' '+scale[i]:''))}
    return capitalizeFirst(parts.join(' ').replace(/\s+/g,' ').trim())+' đồng';
  }

  function paidMonthsFromPlan(plan){
    if(Number.isFinite(Number(plan?.paid)))return intNumber(plan.paid);
    const use=intNumber(plan?.use),gift=intNumber(plan?.gift);return Math.max(0,use-gift);
  }
  function contractPlanName(plan){
    let name=String(plan?.name||'').replace(/[★☆]/g,'').replace(/\bBEST\s*CHOICE\b/gi,'').replace(/\bTỐI ƯU CHI PHÍ\b/gi,'').replace(/\s*·\s*·\s*/g,' · ').trim();
    if(!name)name=`Gói ${String(plan?.term||paidMonthsFromPlan(plan)||'').padStart(2,'0')} tháng`;
    if(plan?.postPilot&&!/sau\s*(pilot|trải nghiệm)/i.test(name))name+=' · Sau Pilot';
    return name.replace(/\s{2,}/g,' ').replace(/^·|·$/g,'').trim();
  }
  function allocationRows(allocation){
    const rows=[];const models=Array.isArray(allocation?.models)?allocation.models:[];
    if(models.length){models.forEach(m=>{const code=String(m.model||m.code||'').toUpperCase();const qty=intNumber(m.qty);if(RHINO_PRODUCT_MASTER[code]&&qty>0)rows.push({code,qty})})}
    else{
      const r68=intNumber(allocation?.r68),r88=intNumber(allocation?.r88);if(r68>0)rows.push({code:'R68',qty:r68});if(r88>0)rows.push({code:'R88',qty:r88});
    }
    return rows.slice(0,4);
  }
  function buildContractValues(info,plan,allocation,contract){
    const rows=allocationRows(allocation);const paidMonths=paidMonthsFromPlan(plan);const unitPrice=finiteNumber(plan?.unit);const values={};
    values.contract_no=contract?.contractNo||'';
    values.contract_date=isoToViDate(contract?.contractDate||'');
    values.contract_place=contract?.contractPlace||'';
    values.contract_year=yearFromDate(contract?.contractDate||'');
    values.legal_name=contract?.legalName||'';
    values.club_name=info?.club||'';
    values.legal_address=contract?.legalAddress||'';
    values.club_address=info?.address||'';
    values.tax_id=contract?.taxId||'';
    values.representative_name=contract?.representativeName||'';
    values.representative_title=contract?.representativeTitle||'';
    values.contact_name=info?.customer||'';
    values.customer_phone=contract?.customerPhone||info?.customerPhone||'';
    values.customer_email=info?.customerEmail||'';

    for(let i=1;i<=4;i++){
      const item=rows[i-1];const master=item?RHINO_PRODUCT_MASTER[item.code]:null;const qty=item?.qty||0;
      values[`product_${i}_group`]=master?.group||'';
      values[`product_${i}_code`]=item?.code||'';
      values[`product_${i}_qty`]=item?String(qty):'';
      values[`product_${i}_unit_price`]=item?formatVND(unitPrice):'';
      values[`product_${i}_paid_months`]=item?String(paidMonths):'';
      values[`product_${i}_amount`]=item?formatVND(unitPrice*qty*paidMonths):'';
    }

    values.plan_name=contractPlanName(plan);
    values.paid_months=String(paidMonths);
    values.gift_months=String(intNumber(plan?.gift));
    values.total_use_months=String(intNumber(plan?.use));
    values.backup_percent=`${Math.round(finiteNumber(plan?.backupPct)*100)}%`;
    values.backup_qty=String(intNumber(plan?.backup));
    values.total_paid_cues=String(intNumber(plan?.n));
    values.total_received_cues=String(intNumber(plan?.total));
    values.rent_total=formatVND(plan?.rent);
    values.deposit_amount=formatVND(plan?.deposit);
    values.grand_total=formatVND(plan?.initial);
    values.grand_total_words=numberToVietnameseWords(plan?.initial);

    for(let i=1;i<=4;i++){
      const item=rows[i-1];const master=item?RHINO_PRODUCT_MASTER[item.code]:null;
      values[`comp_${i}_group`]=master?.group||'';
      values[`comp_${i}_code`]=item?.code||'';
      values[`comp_${i}_price`]=master?formatVND(master.compensationPrice):'';
      values[`comp_${i}_note`]=master?.compensationNote||'';
    }
    values.party_a_signer_name=COMPANY_CONFIG.partyASignerName;
    values.party_b_signer_name=contract?.representativeName||'';
    return values;
  }

  function findBookmarkStart(xmlDoc,name){
    const nodes=Array.from(xmlDoc.getElementsByTagNameNS(W_NS,'bookmarkStart'));
    return nodes.find(n=>n.getAttributeNS(W_NS,'name')===name||n.getAttribute('w:name')===name)||null;
  }
  function bookmarkId(node){return node?.getAttributeNS(W_NS,'id')||node?.getAttribute('w:id')||''}
  function isBookmarkEnd(node,id){return node&&node.nodeType===1&&node.namespaceURI===W_NS&&node.localName==='bookmarkEnd'&&(node.getAttributeNS(W_NS,'id')||node.getAttribute('w:id'))===id}
  function ensureTextNode(run){
    let t=run.getElementsByTagNameNS(W_NS,'t')[0];if(t)return t;
    t=run.ownerDocument.createElementNS(W_NS,'w:t');run.appendChild(t);return t;
  }
  function setBookmarkText(xmlDoc,name,value){
    const start=findBookmarkStart(xmlDoc,name);if(!start)return false;const id=bookmarkId(start);let node=start.nextSibling,targetRun=null,end=null;
    while(node){if(isBookmarkEnd(node,id)){end=node;break}if(node.nodeType===1&&node.namespaceURI===W_NS&&node.localName==='r'&&!targetRun)targetRun=node;node=node.nextSibling}
    if(!targetRun){
      targetRun=xmlDoc.createElementNS(W_NS,'w:r');const t=xmlDoc.createElementNS(W_NS,'w:t');targetRun.appendChild(t);start.parentNode.insertBefore(targetRun,end||start.nextSibling);
    }
    const text=ensureTextNode(targetRun);text.textContent=String(value??'');if(/^\s|\s$/.test(String(value??'')))text.setAttributeNS(XML_NS,'xml:space','preserve');else text.removeAttributeNS(XML_NS,'space');
    // Clear any additional text nodes that might exist inside the bookmark range.
    node=targetRun.nextSibling;while(node&&node!==end){if(node.nodeType===1){Array.from(node.getElementsByTagNameNS(W_NS,'t')).forEach(t=>{t.textContent=''})}node=node.nextSibling}
    return true;
  }

  async function fetchTemplate(){
    const res=await fetch(TEMPLATE_URL,{cache:'no-store'});if(!res.ok)throw new Error(`Không tải được template hợp đồng (${res.status}).`);return res.arrayBuffer();
  }
  async function createRentalContractBlob(info,planData,allocationData,contractInfo){
    if(typeof global.JSZip==='undefined')throw new Error('Thiếu asset/jszip.min.js');
    const buffer=await fetchTemplate();const zip=await global.JSZip.loadAsync(buffer);const file=zip.file('word/document.xml');if(!file)throw new Error('Template DOCX thiếu word/document.xml.');
    const xml=await file.async('string');const parser=new DOMParser();const doc=parser.parseFromString(xml,'application/xml');
    const parseError=doc.getElementsByTagName('parsererror')[0];if(parseError)throw new Error('Không đọc được XML của template hợp đồng.');
    const values=buildContractValues(info,planData,allocationData,contractInfo);const missing=[];
    Object.keys(values).forEach(name=>{if(!setBookmarkText(doc,name,values[name]))missing.push(name)});
    if(missing.length)console.warn('Template không có bookmark:',missing.join(', '));
    const out=new XMLSerializer().serializeToString(doc);zip.file('word/document.xml',out);
    return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',compression:'DEFLATE'});
  }
  async function exportRentalContract(info,planData,allocationData,contractInfo){
    const blob=await createRentalContractBlob(info,planData,allocationData,contractInfo);const date=String(contractInfo?.contractDate||new Date().toISOString().slice(0,10));
    const filename=`Hop_dong_thue_gay_Rhino_${safeFilePart(info?.club||info?.customer)}_${date}.docx`;const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),2000);return filename;
  }

  global.RHINO_PRODUCT_MASTER=RHINO_PRODUCT_MASTER;
  global.RHINO_COMPANY_CONFIG=COMPANY_CONFIG;
  global.RhinoContract={
    exportRentalContract,
    createRentalContractBlob,
    buildContractValues,
    numberToVietnameseWords,
    formatVND,
    PRODUCT_MASTER:RHINO_PRODUCT_MASTER,
    COMPANY_CONFIG
  };
})(typeof window!=='undefined'?window:globalThis);
