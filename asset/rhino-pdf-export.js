/*
  Rhino Cue Platform - PDF Export V32
  Creates a real one-page A4 PDF in the browser without external PDF libraries.
  Text is rasterized through Canvas, preserving Vietnamese glyphs with the browser's system font.
*/
(function(window, document){
  'use strict';

  const COMPANY='CÔNG TY TNHH CARBON BILLIARDS';
  const LOGO_PATH='asset/logo.png';
  const PDF_MIME='application/pdf';

  const safeName=(v)=>String(v||'Bao_gia').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').replace(/[^a-zA-Z0-9_-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,54)||'Bao_gia';
  const money=(v)=>new Intl.NumberFormat('vi-VN').format(Math.round(Number(v)||0));
  const text=(v)=>String(v??'').trim()||'-';

  function rentalPricing(data){
    const n=Math.max(0,Number(data.n||0)),unit=Math.max(0,Number(data.unit||0)),rent=Math.max(0,Number(data.rent||0)),use=Math.max(0,Number(data.use||0));
    let paid=Number(data.paid);if(!Number.isFinite(paid)||paid<=0)paid=n>0&&unit>0?Math.max(1,Math.round(rent/(n*unit))):use;
    const gift=Math.max(0,Number(data.gift||0),use-paid),deposit=Math.max(0,Number(data.deposit||0)),finalPay=Math.max(0,Number(data.initial??(rent+deposit)));
    return {n,unit,rent,use,paid,gift,deposit,finalPay};
  }
  function rentalPlanLabel(data,p){
    let label='';
    if(Number(data.term)===1)label='Kỳ 01 tháng';else if(Number(data.term)===2)label='Kỳ 02 tháng';else if(Number(data.term)===3)label='Gói 03 tháng';else if(Number(data.term)===6)label='Gói 06 tháng';else if(Number(data.term)===12)label='Gói 12 tháng';
    else label=String(data.name||'Phương án thuê').replace(/\s*[·-]\s*dùng\s*\d+\s*tháng.*$/i,'').replace(/\s*[·-]\s*tặng(?:\s*thêm)?\s*\d+\s*tháng(?:\s*sử dụng)?.*$/i,'').trim();
    return data.postPilot?`${label} · Sau Pilot`:label;
  }
  function rentalAllocations(data,p){
    let arr=(Array.isArray(data.models)?data.models:[]).map(x=>({model:x&&x.model==='R-88'?'R-88':'R-68',qty:Math.max(0,Math.floor(Number(x&&x.qty)||0))})).filter(x=>x.qty>0);
    if(!arr.length||arr.reduce((a,b)=>a+b.qty,0)!==p.n)arr=[{model:data.model==='R-88'?'R-88':'R-68',qty:p.n}];
    return arr;
  }
  function purchasePricing(data){
    const qty=Math.max(0,Number(data.qty||0)),baseUnit=Math.max(0,Number(data.base??data.unit??0)),finalUnit=Math.max(0,Number(data.unit??baseUnit)),discountPerUnit=Math.max(0,baseUnit-finalUnit),gross=baseUnit*qty,discount=discountPerUnit*qty,finalPay=Math.max(0,Number(data.total??(gross-discount)));
    return {qty,baseUnit,finalUnit,discountPerUnit,gross,discount,finalPay};
  }

  function roundedRect(ctx,x,y,w,h,r,fill,stroke){
    ctx.beginPath();ctx.roundRect?ctx.roundRect(x,y,w,h,r):(ctx.rect(x,y,w,h));
    if(fill){ctx.fillStyle=fill;ctx.fill()}if(stroke){ctx.strokeStyle=stroke;ctx.lineWidth=.6;ctx.stroke()}
  }
  function setFont(ctx,size,weight='400',italic=false){ctx.font=`${italic?'italic ':''}${weight} ${size}px Arial, Helvetica, sans-serif`}
  function wrapLines(ctx,value,maxW){
    const raw=String(value??'');if(!raw)return ['-'];const lines=[];
    raw.split(/\n/).forEach(par=>{const words=par.split(/\s+/);let line='';for(const word of words){const test=line?`${line} ${word}`:word;if(ctx.measureText(test).width<=maxW||!line)line=test;else{lines.push(line);line=word}}if(line)lines.push(line)});
    return lines.length?lines:['-'];
  }
  function drawWrapped(ctx,value,x,y,maxW,lineH,opts={}){
    const lines=wrapLines(ctx,value,maxW),align=opts.align||'left';ctx.textAlign=align;ctx.textBaseline='top';let tx=x;if(align==='center')tx=x+maxW/2;else if(align==='right')tx=x+maxW;
    lines.forEach((line,i)=>ctx.fillText(line,tx,y+i*lineH));return lines.length*lineH;
  }
  function drawLabelValue(ctx,label,value,x,y,w,labelW,font=8.2){
    setFont(ctx,font,'700');ctx.fillStyle='#0B3B59';drawWrapped(ctx,label,x,y,labelW,11);
    setFont(ctx,font,'400');ctx.fillStyle='#263944';return drawWrapped(ctx,text(value),x+labelW,y,w-labelW,11);
  }

  async function loadLogo(){
    if(location.protocol==='file:')return null;
    return new Promise(resolve=>{const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>resolve(null);img.src=`${LOGO_PATH}?v=${Date.now()}`});
  }

  function tableLayout(orientation,pageW){
    const portrait=orientation!=='landscape';
    const weights=portrait?[.15,.18,.085,.105,.09,.14,.09,.16]:[.16,.20,.08,.10,.09,.13,.09,.15];
    const usable=pageW-(portrait?34:42);const sum=weights.reduce((a,b)=>a+b,0);return weights.map(v=>usable*v/sum);
  }
  function drawTable(ctx,x,y,widths,headers,rows,fontSize){
    const lineH=fontSize+2.2,padX=4,padY=5;let cy=y;
    const totalW=widths.reduce((a,b)=>a+b,0);
    // header
    ctx.fillStyle='#0B3B59';ctx.fillRect(x,cy,totalW,25);setFont(ctx,fontSize,'700');ctx.fillStyle='#fff';ctx.textBaseline='middle';ctx.textAlign='center';let cx=x;
    headers.forEach((h,i)=>{ctx.fillText(h,cx+widths[i]/2,cy+12.5);cx+=widths[i]});cy+=25;
    rows.forEach(row=>{
      const isPromo=row.kind==='promo',isTotal=row.kind==='total',isPay=row.kind==='pay';
      setFont(ctx,fontSize,isTotal||isPay?'700':'400');
      let maxLines=1;row.cells.forEach((v,i)=>{maxLines=Math.max(maxLines,wrapLines(ctx,text(v),Math.max(8,widths[i]-padX*2)).length)});const rh=Math.max(24,padY*2+maxLines*lineH);
      ctx.fillStyle=isPromo?'#EEF9F6':(isPay?'#FFF6EC':isTotal?'#F2FAF7':'#F8FAFB');ctx.fillRect(x,cy,totalW,rh);
      cx=x;row.cells.forEach((v,i)=>{ctx.strokeStyle='#D8E4E7';ctx.lineWidth=.5;ctx.strokeRect(cx,cy,widths[i],rh);const numeric=i>=3;ctx.fillStyle=(i===7&&isPromo)?'#138879':(i===7&&isPay)?'#E06A1A':'#263944';setFont(ctx,fontSize,(i===7&&(isPromo||isPay))||isTotal?'700':'400');drawWrapped(ctx,text(v),cx+padX,cy+padY,widths[i]-padX*2,lineH,{align:numeric?'right':'left'});cx+=widths[i]});cy+=rh;
    });
    return cy;
  }

  function buildRentalRows(data){
    const p=rentalPricing(data),alloc=rentalAllocations(data,p),monthlyTotal=Math.round(p.unit*p.n),backupCount=Math.max(0,Number(data.backup||0)),backupPct=Math.round(Math.max(0,Number(data.backupPct||0))*100),giftValue=p.gift>0?monthlyTotal*p.gift:0,backupMonthly=p.unit*backupCount,backupValue=backupMonthly*p.use,promoTotal=giftValue+backupValue;
    const rows=[];
    alloc.forEach(part=>{const partMonthly=p.unit*part.qty;rows.push({cells:[`Thuê gậy CLB ${part.model}`,rentalPlanLabel(data,p),'Gậy',money(p.unit),String(part.qty),money(partMonthly),String(p.paid),money(partMonthly*p.paid)]})});
    if(p.gift>0)rows.push({kind:'promo',cells:['Khuyến mại',`Tặng ${p.gift} tháng sử dụng`,'Tháng',money(monthlyTotal),'-',money(monthlyTotal),String(p.gift),money(giftValue)]});
    if(backupCount>0)rows.push({kind:'promo',cells:['Ưu đãi',`Gậy dự phòng ${backupPct}%`,'Gậy',money(p.unit),String(backupCount),money(backupMonthly),String(p.use),money(backupValue)]});
    if(p.deposit>0)rows.push({cells:['Tiền cọc','Hoàn lại sau đối soát','-',money(p.deposit),'1','-','-',money(p.deposit)]});
    return {p,rows,promoTotal};
  }
  function buildPurchaseRows(data){
    const p=purchasePricing(data),rows=[{cells:[`Gậy Rhino ${data.model||''}`,`Giá theo bảng`,'Gậy',money(p.baseUnit),String(p.qty),money(p.gross)]}];
    if(p.discount>0)rows.push({kind:'promo',cells:['Khuyến mại','Ưu đãi khách đang sử dụng dịch vụ thuê','Gậy',money(p.discountPerUnit),String(p.qty),money(p.discount)]});
    return {p,rows};
  }

  async function renderQuoteCanvas(type,info,data,orientation){
    const landscape=orientation==='landscape',pageW=landscape?841.89:595.28,pageH=landscape?595.28:841.89,scale=2.1;
    const canvas=document.createElement('canvas');canvas.width=Math.round(pageW*scale);canvas.height=Math.round(pageH*scale);const ctx=canvas.getContext('2d');ctx.scale(scale,scale);ctx.fillStyle='#fff';ctx.fillRect(0,0,pageW,pageH);
    const margin=landscape?21:17,contentW=pageW-margin*2;let y=14;
    const logo=await loadLogo();if(logo){const targetH=34,targetW=Math.min(115,logo.naturalWidth*(targetH/logo.naturalHeight));ctx.drawImage(logo,margin,y,targetW,targetH)}else{setFont(ctx,13,'800');ctx.fillStyle='#25363D';ctx.fillText('RHINO',margin,y+5);setFont(ctx,7,'700');ctx.fillText('CUE PLATFORM',margin,y+20)}
    setFont(ctx,10.2,'700');ctx.fillStyle='#0B3B59';ctx.textAlign='right';ctx.textBaseline='top';ctx.fillText(COMPANY,pageW-margin,y+1);
    const title=type==='rental'?'BÁO GIÁ THUÊ GẬY CLB':'BÁO GIÁ MUA GẬY CLB';setFont(ctx,landscape?18:17,'800');ctx.textAlign='center';ctx.fillText(title,pageW/2,y+27);
    setFont(ctx,7.5,'400',true);ctx.fillStyle='#6E7E88';ctx.fillText('(Báo giá có giá trị 14 ngày kể từ ngày báo giá)',pageW/2,y+50);y+=75;

    const colGap=landscape?45:22,leftW=(contentW-colGap)/2,rightX=margin+leftW+colGap;const infoRows=[['Khách hàng',info.customer,'Thời gian lập phiếu',new Date().toLocaleString('vi-VN')],['Câu lạc bộ / Quán',info.club,'Người lập báo giá',info.seller],['Địa chỉ câu lạc bộ',info.address,'Liên hệ',info.sellerPhone],['Email khách hàng',info.customerEmail||'','Showroom',info.showroom||'']];
    let iy=y;infoRows.forEach(row=>{setFont(ctx,8.1,'700');const leftLines=Math.max(wrapLines(ctx,text(row[1]),leftW-75).length,1),rightLines=Math.max(wrapLines(ctx,text(row[3]),leftW-75).length,1),rh=Math.max(18,Math.max(leftLines,rightLines)*10.5+3);drawLabelValue(ctx,row[0],row[1],margin,iy,leftW,72,8.1);drawLabelValue(ctx,row[2],row[3],rightX,iy,leftW,72,8.1);iy+=rh});y=iy+6;

    ctx.strokeStyle='#D5E1E5';ctx.lineWidth=.7;ctx.beginPath();ctx.moveTo(margin,y);ctx.lineTo(pageW-margin,y);ctx.stroke();y+=7;
    if(type==='rental'){
      const built=buildRentalRows(data),widths=tableLayout(orientation,pageW),headers=['NỘI DUNG','THÔNG TIN','QUY CÁCH','GIÁ THUÊ','SỐ LƯỢNG','GIÁ / THÁNG','SỐ THÁNG','THÀNH TIỀN'];
      y=drawTable(ctx,margin,y,widths,headers,built.rows,landscape?8.2:7.2);
      const totalW=widths.reduce((a,b)=>a+b,0);const labelW=widths.slice(0,7).reduce((a,b)=>a+b,0),lastW=widths[7];
      function totalRow(label,value,kind){const h=26;ctx.fillStyle=kind==='pay'?'#FFF6EC':'#EEF9F6';ctx.fillRect(margin,y,totalW,h);ctx.strokeStyle='#D8E4E7';ctx.strokeRect(margin,y,labelW,h);ctx.strokeRect(margin+labelW,y,lastW,h);setFont(ctx,8.4,'800');ctx.textBaseline='middle';ctx.textAlign='left';ctx.fillStyle=kind==='pay'?'#263944':'#138879';ctx.fillText(label,margin+5,y+h/2);ctx.textAlign='right';ctx.fillStyle=kind==='pay'?'#E06A1A':'#138879';ctx.fillText(money(value),margin+totalW-5,y+h/2);y+=h}
      totalRow('TỔNG ƯU ĐÃI TƯƠNG ĐƯƠNG',built.promoTotal,'promo');totalRow('TỔNG TIỀN KHÁCH HÀNG CHI TRẢ',built.p.finalPay,'pay');
      y+=9;setFont(ctx,8.1,'700');ctx.textAlign='left';ctx.textBaseline='top';ctx.fillStyle='#138879';const note=`Chi phí thực tế sau ưu đãi: ${money(data.effective)} VND / gậy thực nhận / tháng. Tổng thực nhận ${Number(data.total||0)} gậy; thời gian sử dụng ${built.p.use} tháng.`;y+=drawWrapped(ctx,note,margin,y,contentW,11)+6;
      setFont(ctx,7.6,'400',true);ctx.fillStyle='#6E7E88';y+=drawWrapped(ctx,'Giá thuê đã gồm VAT. Các nội dung áp dụng theo chính sách và hợp đồng tại thời điểm ký kết.',margin,y,contentW,10.5)+10;
    }else{
      const built=buildPurchaseRows(data),widths=tableLayout(orientation,pageW).slice(0,6),sum=widths.reduce((a,b)=>a+b,0),factor=contentW/sum;for(let i=0;i<widths.length;i++)widths[i]*=factor;
      y=drawTable(ctx,margin,y,widths,['NỘI DUNG','THÔNG TIN','QUY CÁCH','ĐƠN GIÁ','SỐ LƯỢNG','THÀNH TIỀN'],built.rows,landscape?8.5:7.7);
      const labelW=widths.slice(0,5).reduce((a,b)=>a+b,0),lastW=widths[5],h=27,totalW=contentW;ctx.fillStyle='#FFF6EC';ctx.fillRect(margin,y,totalW,h);ctx.strokeStyle='#D8E4E7';ctx.strokeRect(margin,y,labelW,h);ctx.strokeRect(margin+labelW,y,lastW,h);setFont(ctx,8.5,'800');ctx.textBaseline='middle';ctx.textAlign='left';ctx.fillStyle='#263944';ctx.fillText('TỔNG TIỀN KHÁCH HÀNG CHI TRẢ',margin+5,y+h/2);ctx.textAlign='right';ctx.fillStyle='#E06A1A';ctx.fillText(money(built.p.finalPay),margin+totalW-5,y+h/2);y+=h+12;
    }
    setFont(ctx,9.2,'800');ctx.textAlign='left';ctx.textBaseline='top';ctx.fillStyle='#0B3B59';y+=drawWrapped(ctx,'Cảm ơn Quý khách đã quan tâm đến sản phẩm và giải pháp của Rhino Cue Platform.',margin,y,contentW,12)+14;
    setFont(ctx,8.5,'700');ctx.fillText(`Trân trọng — ${COMPANY}`,margin,y);
    // Safety: page is fixed A4; compact quote layout intentionally leaves bottom whitespace.
    return {canvas,pageW,pageH};
  }

  function asciiBytes(str){return new TextEncoder().encode(str)}
  function concatBytes(parts){let len=parts.reduce((s,p)=>s+p.length,0),out=new Uint8Array(len),off=0;parts.forEach(p=>{out.set(p,off);off+=p.length});return out}
  async function canvasToPdfBlob(canvas,pageW,pageH){
    const jpgBlob=await new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(new Error('Không thể rasterize PDF.')),'image/jpeg',0.94));
    const jpg=new Uint8Array(await jpgBlob.arrayBuffer());const objs=[];
    objs[1]=[asciiBytes('<< /Type /Catalog /Pages 2 0 R >>')];
    objs[2]=[asciiBytes('<< /Type /Pages /Kids [3 0 R] /Count 1 >>')];
    objs[3]=[asciiBytes(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageW.toFixed(2)} ${pageH.toFixed(2)}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>`)];
    objs[4]=[asciiBytes(`<< /Type /XObject /Subtype /Image /Width ${canvas.width} /Height ${canvas.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpg.length} >>\nstream\n`),jpg,asciiBytes('\nendstream')];
    const content=`q\n${pageW.toFixed(2)} 0 0 ${pageH.toFixed(2)} 0 0 cm\n/Im0 Do\nQ\n`,contentBytes=asciiBytes(content);objs[5]=[asciiBytes(`<< /Length ${contentBytes.length} >>\nstream\n`),contentBytes,asciiBytes('endstream')];
    const parts=[asciiBytes('%PDF-1.4\n%RhinoPDF\n')],offsets=[0];let pos=parts[0].length;
    for(let i=1;i<=5;i++){offsets[i]=pos;const head=asciiBytes(`${i} 0 obj\n`),tail=asciiBytes('\nendobj\n');parts.push(head,...objs[i],tail);pos+=head.length+objs[i].reduce((s,p)=>s+p.length,0)+tail.length}
    const xrefPos=pos;let xref=`xref\n0 6\n0000000000 65535 f \n`;for(let i=1;i<=5;i++)xref+=`${String(offsets[i]).padStart(10,'0')} 00000 n \n`;const trailer=`trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xrefPos}\n%%EOF\n`;parts.push(asciiBytes(xref+trailer));return new Blob([concatBytes(parts)],{type:PDF_MIME});
  }
  function downloadBlob(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1800)}

  async function createQuoteBlob(type,info,data,orientation='portrait'){
    if(!['rental','purchase'].includes(type))throw new Error('Loại báo giá không hợp lệ.');if(!info||!info.customer||!info.club||!info.address||!info.seller)throw new Error('Thiếu thông tin báo giá.');if(!data)throw new Error('Thiếu dữ liệu báo giá.');
    const rendered=await renderQuoteCanvas(type,info,data,orientation);return canvasToPdfBlob(rendered.canvas,rendered.pageW,rendered.pageH);
  }
  async function exportQuote(type,info,data,orientation='portrait'){
    const blob=await createQuoteBlob(type,info,data,orientation);let modelLabel=data.model||'';if(type==='rental'&&Array.isArray(data.models)&&data.models.length)modelLabel=data.models.filter(x=>Number(x&&x.qty)>0).map(x=>x.model).join('_');const base=type==='rental'?'Bao_gia_thue_gay_Rhino':'Bao_gia_mua_gay_Rhino';const model=modelLabel?`${safeName(modelLabel)}_`:'';const fileName=`${base}_${model}${safeName(info.club)}_${new Date().toISOString().slice(0,10)}.pdf`;downloadBlob(blob,fileName);return fileName;
  }

  window.RhinoPDF={version:'1.0.0',exportQuote,createQuoteBlob,safeName};
})(window,document);
