(function(){
'use strict';
const $id=id=>document.getElementById(id);
const FLOW_STORE='rhino_document_fields_v40';
const flow={docType:'quote',format:'pdf',payload:null,extra:{},focusField:''};
const DOC_LABEL={quote:'Báo giá thuê',contract:'Hợp đồng',handoverCues:'Biên bản bàn giao gậy',handoverAssets:'Biên bản bàn giao tài sản'};
const FIELD_SCHEMAS={
  contract:[
    {key:'contractNo',label:'Số hợp đồng',required:true,placeholder:'001/2026/HĐCT-RHINO',wide:true},
    {key:'contractDate',label:'Ngày ký hợp đồng',required:true,type:'date'},
    {key:'contractPlace',label:'Địa điểm ký',required:true,placeholder:'Hà Nội / Nghệ An...'},
    {key:'taxId',label:'Mã số thuế / CCCD',required:true,placeholder:'MST / CCCD / CMND'},
    {key:'representativeName',label:'Người đại diện Bên B',required:true,placeholder:'Họ và tên'},
    {key:'representativeTitle',label:'Chức vụ',required:true,placeholder:'Chủ hộ kinh doanh'},
    {key:'customerPhone',label:'Điện thoại liên hệ',required:true,placeholder:'0912... hoặc +84...',syncBasic:'quoteCustomerPhone'},
    {key:'rentalStartDate',label:'Ngày bắt đầu thuê',required:true,type:'date'},
    {key:'rentalEndDate',label:'Ngày kết thúc dự kiến',required:true,type:'date'}
  ],
  handoverCues:[
    {key:'handoverNo',label:'Số biên bản bàn giao',required:true,placeholder:'001/2026/BBBG-RHINO',wide:true},
    {key:'contractNo',label:'Số hợp đồng',required:true,placeholder:'001/2026/HĐCT-RHINO'},
    {key:'contractDate',label:'Ngày ký hợp đồng',required:true,type:'date'},
    {key:'handoverDate',label:'Ngày bàn giao',required:true,type:'date'},
    {key:'handoverTime',label:'Thời gian bàn giao',required:true,type:'time'},
    {key:'handoverPlace',label:'Địa điểm bàn giao',required:true,placeholder:'Địa chỉ bàn giao',wide:true},
    {key:'representativeName',label:'Người đại diện Bên B',required:true},
    {key:'representativeTitle',label:'Chức vụ',required:true},
    {key:'customerPhone',label:'Điện thoại Bên B',required:true,syncBasic:'quoteCustomerPhone'},
    {key:'accessories',label:'Vật dụng đi kèm',required:false,placeholder:'Nếu có',wide:true},
    {key:'handoverNote',label:'Ghi chú bàn giao',required:false,type:'textarea',placeholder:'Ghi chú thêm nếu có',wide:true}
  ],
  handoverAssets:[
    {key:'contractNo',label:'Số hợp đồng',required:true,placeholder:'001/2026/HĐCT-RHINO'},
    {key:'contractDate',label:'Ngày ký hợp đồng',required:true,type:'date'},
    {key:'handoverDate',label:'Ngày bàn giao',required:true,type:'date'},
    {key:'handoverPlace',label:'Địa điểm bàn giao',required:true,wide:true},
    {key:'representativeName',label:'Người đại diện Bên B',required:true},
    {key:'representativeTitle',label:'Chức vụ',required:true}
  ]
};
function today(){if(typeof todayLocalISO==='function')return todayLocalISO();const d=new Date(),p=n=>String(n).padStart(2,'0');return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`}
function nowTime(){const d=new Date();return `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`}
function clean(s){return String(s||'').trim().replace(/\s+/g,' ')}
function storeRead(){try{return JSON.parse(localStorage.getItem(FLOW_STORE)||'{}')||{}}catch(_){return {}}}
function storeWrite(data){try{localStorage.setItem(FLOW_STORE,JSON.stringify(data))}catch(_){}}
function storeKey(){return clean(flow.payload?.info?.club).toLocaleLowerCase('vi')||'__default__'}
function saveExtra(){const all=storeRead();all[storeKey()]={...(all[storeKey()]||{}),...flow.extra};storeWrite(all)}
function savedExtra(){return storeRead()[storeKey()]||{}}
function savedCustomer(){try{const list=readStore(RHINO_MEMORY_KEYS.customers);const n=clean(flow.payload?.info?.customer).toLocaleLowerCase('vi'),c=clean(flow.payload?.info?.club).toLocaleLowerCase('vi');return list.find(x=>clean(x.name).toLocaleLowerCase('vi')===n&&clean(x.club).toLocaleLowerCase('vi')===c)||{}}catch(_){return {}}}
function suggestedNo(){try{return suggestedContractNo(today())}catch(_){const y=new Date().getFullYear();return `001/${y}/HĐCT-RHINO`}}
function suggestedHandoverNo(contractNo){const y=(String(contractNo||'').match(/\/(\d{4})\//)||[])[1]||new Date().getFullYear();const seq=(String(contractNo||'').match(/^(\d+)/)||[])[1]||'001';return `${String(seq).padStart(3,'0')}/${y}/BBBG-RHINO`}
function initExtra(force=false){const info=flow.payload.info,plan=flow.payload.data,saved=savedExtra(),cust=savedCustomer();if(force)flow.extra={};const base={
  contractNo:saved.contractNo||suggestedNo(),contractDate:saved.contractDate||today(),contractPlace:saved.contractPlace||(typeof contractDefaultPlace==='function'?contractDefaultPlace():'Hà Nội'),taxId:saved.taxId||cust.taxId||'',representativeName:saved.representativeName||cust.representativeName||info.customer||'',representativeTitle:saved.representativeTitle||cust.representativeTitle||'Chủ hộ kinh doanh',customerPhone:info.customerPhone||cust.phone||'',rentalStartDate:saved.rentalStartDate||saved.contractDate||today(),rentalEndDate:saved.rentalEndDate||'',handoverNo:saved.handoverNo||'',handoverDate:saved.handoverDate||today(),handoverTime:saved.handoverTime||nowTime(),handoverPlace:saved.handoverPlace||info.address||'',accessories:saved.accessories||'',handoverNote:saved.handoverNote||''
};base.handoverNo=base.handoverNo||suggestedHandoverNo(base.contractNo);if(!base.rentalEndDate&&window.RhinoDocuments)base.rentalEndDate=RhinoDocuments.addMonths(base.rentalStartDate,plan.use);flow.extra={...base,...flow.extra};}
function formatOptions(){return flow.docType==='quote'?[['pdf','PDF (ưu tiên)'],['excel','Excel']]:[['pdf','PDF (ưu tiên)'],['word','Word'],['excel','Excel']]}
function renderCenter(){document.querySelectorAll('[data-doc-type]').forEach(b=>b.classList.toggle('active',b.dataset.docType===flow.docType));const sel=$id('docFormatSelect');sel.innerHTML=formatOptions().map(([v,t])=>`<option value="${v}">${t}</option>`).join('');if(!formatOptions().some(x=>x[0]===flow.format))flow.format='pdf';sel.value=flow.format;$id('documentCenterError').textContent=''}
function openCenter(){flow.docType='quote';flow.format='pdf';initExtra(false);renderCenter();closeModal('quoteClientModal');openModal('documentCenterModal',false)}
function currentSchema(){return FIELD_SCHEMAS[flow.docType]||[]}
function fieldId(key){return `docField_${key}`}
function renderFields(){const form=$id('documentFieldsForm'),schema=currentSchema();form.innerHTML='<div class="required-legend"><span>*</span> là bắt buộc. Các thông tin đã có được tự điền và vẫn có thể sửa.</div>'+schema.map(f=>`<div class="quote-field ${f.wide?'wide':''}" data-doc-field="${f.key}"><label for="${fieldId(f.key)}">${f.label}${f.required?' <span class="req-star" aria-hidden="true">*</span>':''}</label>${f.type==='textarea'?`<textarea id="${fieldId(f.key)}" placeholder="${f.placeholder||''}">${flow.extra[f.key]||''}</textarea>`:`<input id="${fieldId(f.key)}" type="${f.type||'text'}" value="${String(flow.extra[f.key]||'').replace(/"/g,'&quot;')}" placeholder="${f.placeholder||''}"/>`}</div>`).join('');$id('documentFieldsTitle').textContent=`${DOC_LABEL[flow.docType]} · Thông tin bổ sung`;$id('documentFieldsError').textContent='';schema.forEach(f=>{const el=$id(fieldId(f.key));el?.addEventListener('input',()=>{flow.extra[f.key]=el.value;el.closest('.quote-field')?.classList.remove('field-error');if(f.syncBasic&&$id(f.syncBasic)){$id(f.syncBasic).value=el.value;flow.payload.info.customerPhone=el.value}if(f.key==='rentalStartDate'){const end=$id(fieldId('rentalEndDate'));if(end){flow.extra.rentalEndDate=RhinoDocuments.addMonths(el.value,flow.payload.data.use);end.value=flow.extra.rentalEndDate}}if(f.key==='contractNo'&&!flow.extra.handoverNo){flow.extra.handoverNo=suggestedHandoverNo(el.value)}})});if(flow.focusField){setTimeout(()=>{$id(fieldId(flow.focusField))?.focus();$id(fieldId(flow.focusField))?.scrollIntoView({block:'center'});flow.focusField=''},100)}}
function validateField(f,v){const s=String(v||'').trim();if(!f.required&&!s)return '';if(f.required&&!s)return `${f.label} là thông tin bắt buộc.`;if(f.key==='customerPhone'&&typeof validPhone==='function'&&!validPhone(s))return 'Số điện thoại chưa hợp lệ.';if(f.key==='taxId'&&typeof validTaxOrIdentity==='function'&&!validTaxOrIdentity(s))return 'MST/CCCD chưa hợp lệ.';if(f.key==='representativeName'&&typeof validHumanName==='function'&&!validHumanName(s))return 'Tên người đại diện chưa hợp lệ.';return ''}
function validateExtra(){for(const f of currentSchema()){const v=flow.extra[f.key],e=validateField(f,v);if(e)return {ok:false,key:f.key,msg:e}}if(flow.docType!=='quote'){const a=flow.payload.allocation;if(!a||a.total!==flow.payload.data.n)return {ok:false,key:'',msg:`Tổng R68 + R88 phải bằng ${flow.payload.data.n} gậy tính phí.`}}return {ok:true}}
function openFields(){initExtra(false);renderFields();closeModal('documentCenterModal');closeModal('documentPreviewModal');openModal('documentFieldsModal',false)}
function focusBasic(id){closeModal('documentPreviewModal');openModal('quoteClientModal',false);setTimeout(()=>{$id(id)?.focus();$id(id)?.scrollIntoView({block:'center'})},100)}
const BOOKMARK_MAP={contract_no:'contractNo',contract_date_long:'contractDate',contract_place:'contractPlace',tax_id:'taxId',customer_phone:'customerPhone',representative_name:'representativeName',representative_title:'representativeTitle',rental_start_date:'rentalStartDate',rental_end_date:'rentalEndDate',handover_no:'handoverNo',handover_time:'handoverTime',handover_date_long:'handoverDate',handover_place:'handoverPlace',accessories:'accessories',handover_note:'handoverNote'};
const BASIC_MAP={club_name:'quoteClubName',club_name_2:'quoteClubName',club_address:'quoteAddress',customer_email:'quoteCustomerEmail'};
function quotePreviewHtml(){const p=flow.payload.data,i=flow.payload.info,a=flow.payload.allocation,models=(a.models||[]).filter(x=>x.qty>0);const rows=models.map(m=>{const amount=p.n>0?p.rent*(m.qty/p.n):0;return `<tr><td>Thuê gậy CLB ${m.model}</td><td>${p.name}</td><td>Gậy</td><td style="text-align:right">${fmt(p.unit)}</td><td style="text-align:right">${m.qty}</td><td style="text-align:right">${fmt(amount)}</td></tr>`}).join('');return `<div class="doc-preview-quote"><h2>BÁO GIÁ THUÊ GẬY CLB</h2><div class="preview-note">(Báo giá có giá trị 14 ngày kể từ ngày báo giá)</div><div class="preview-info"><button class="preview-edit" data-basic-field="quoteCustomerName">Khách hàng: ${xmlEsc(i.customer)}</button><button class="preview-edit" data-basic-field="quoteClubName">CLB / Quán: ${xmlEsc(i.club)}</button><button class="preview-edit" data-basic-field="quoteAddress">Địa chỉ: ${xmlEsc(i.address)}</button><button class="preview-edit" data-basic-field="quoteCustomerPhone">Điện thoại: ${xmlEsc(i.customerPhone||'-')}</button><button class="preview-edit" data-basic-field="quoteCustomerEmail">Email: ${xmlEsc(i.customerEmail||'-')}</button><span>Người lập báo giá: ${xmlEsc(i.seller||'')}</span></div><table><thead><tr><th>Nội dung</th><th>Thông tin</th><th>Quy cách</th><th>Đơn giá</th><th>Số lượng</th><th>Thành tiền</th></tr></thead><tbody>${rows}<tr><td>Khuyến mại</td><td>Tặng ${p.gift||0} tháng</td><td>Tháng</td><td>${fmt(p.unit*p.n)}</td><td>-</td><td>${fmt(p.unit*p.n*(p.gift||0))}</td></tr><tr><td>Ưu đãi</td><td>Gậy dự phòng</td><td>Gậy</td><td>${fmt(p.unit)}</td><td>${p.backup}</td><td>${fmt(p.unit*p.backup*p.use)}</td></tr></tbody></table><div class="sumline"><span>Tổng tiền khách hàng chi trả</span><strong>${fmt(p.initial)}</strong></div></div>`}
function buildDocContext(){return {info:flow.payload.info,plan:flow.payload.data,allocation:flow.payload.allocation,extra:flow.extra}}
function renderPreview(){const box=$id('documentPreviewPages');box.innerHTML=flow.docType==='quote'?quotePreviewHtml():RhinoDocuments.previewHtml(flow.docType,buildDocContext());$id('documentPreviewTitle').textContent=`Xem trước · ${DOC_LABEL[flow.docType]}`;$id('documentPreviewError').textContent='';box.querySelectorAll('[data-edit-field]').forEach(el=>el.addEventListener('click',()=>{const k=el.dataset.editField;if(BOOKMARK_MAP[k]){flow.focusField=BOOKMARK_MAP[k];openFields()}else if(BASIC_MAP[k])focusBasic(BASIC_MAP[k]);else if(['paid_qty','cue_types','paid_cue_types','paid_allocation_note'].includes(k))focusBasic('quoteUseR68');else{closeModal('documentPreviewModal');document.querySelector('.calculator')?.scrollIntoView({behavior:'smooth',block:'start'});setTimeout(()=>$id('cues')?.focus(),300)}}));box.querySelectorAll('[data-basic-field]').forEach(el=>el.addEventListener('click',()=>focusBasic(el.dataset.basicField)))}
function openPreview(){const check=flow.docType==='quote'?{ok:true}:validateExtra();if(!check.ok){$id('documentFieldsError').textContent=check.msg;const wrap=document.querySelector(`[data-doc-field="${check.key}"]`);wrap?.classList.add('field-error');$id(fieldId(check.key))?.focus();return}saveExtra();renderPreview();closeModal('documentFieldsModal');closeModal('documentCenterModal');openModal('documentPreviewModal',false)}
async function doExport(){try{if(flow.docType==='quote'){if(flow.format==='excel'){if(!window.RhinoExcel)throw new Error('Thiếu Excel exporter.');await RhinoExcel.exportQuote('rental',flow.payload.info,flow.payload.data)}else{if(!window.RhinoPDF)throw new Error('Thiếu PDF exporter.');await RhinoPDF.exportQuote('rental',flow.payload.info,flow.payload.data,'portrait')}}else{const ctx=buildDocContext();if(flow.format==='word')await RhinoDocuments.exportWord(flow.docType,ctx);else if(flow.format==='excel')await RhinoDocuments.exportExcel(flow.docType,ctx);else await RhinoDocuments.exportPdf(flow.docType,ctx)}try{saveSmartEntry('customers',{name:flow.payload.info.customer,club:flow.payload.info.club,address:flow.payload.info.address,email:flow.payload.info.customerEmail||'',phone:flow.payload.info.customerPhone||'',taxId:flow.extra.taxId||'',representativeName:flow.extra.representativeName||'',representativeTitle:flow.extra.representativeTitle||''})}catch(_){}saveExtra();return true}catch(err){console.error(err);const msg=err?.message||'Không thể xuất file.';$id('documentPreviewError').textContent=msg;$id('documentFieldsError').textContent=msg;$id('documentCenterError').textContent=msg;return false}}

// Intercept rental confirmation before legacy handler opens the old format modal.
$id('confirmQuoteExport')?.addEventListener('click',function(e){if(typeof pendingQuoteType!=='undefined'&&pendingQuoteType==='rental'){e.preventDefault();e.stopImmediatePropagation();const info=quoteInfoFromForm(),err=validateQuoteInfo(info);if(err){$id('quoteError').textContent=err;return}const ae=validateQuoteCueAllocation();if(ae){$id('quoteError').textContent=ae;return}$id('quoteError').textContent='';flow.payload={type:'rental',info,data:rentalQuoteData(),allocation:quoteCueAllocation()};initExtra(true);openCenter()}},true);

document.querySelectorAll('[data-doc-type]').forEach(btn=>btn.addEventListener('click',()=>{flow.docType=btn.dataset.docType;flow.format='pdf';initExtra(false);renderCenter()}));
$id('docFormatSelect')?.addEventListener('change',e=>{flow.format=e.target.value});
$id('documentCenterNext')?.addEventListener('click',()=>{flow.format=$id('docFormatSelect').value||'pdf';if(flow.docType==='quote')openPreview();else openFields()});
$id('documentFieldsBack')?.addEventListener('click',()=>{closeModal('documentFieldsModal');renderCenter();openModal('documentCenterModal',false)});
$id('documentFieldsPreview')?.addEventListener('click',openPreview);
$id('documentFieldsExport')?.addEventListener('click',async()=>{const v=validateExtra();if(!v.ok){$id('documentFieldsError').textContent=v.msg;document.querySelector(`[data-doc-field="${v.key}"]`)?.classList.add('field-error');$id(fieldId(v.key))?.focus();return}saveExtra();await doExport()});
$id('documentPreviewBack')?.addEventListener('click',()=>{if(flow.docType==='quote')focusBasic('quoteCustomerName');else openFields()});
$id('documentPreviewExport')?.addEventListener('click',async()=>{const b=$id('documentPreviewExport'),old=b.textContent;b.disabled=true;b.textContent='Đang xuất...';await doExport();b.disabled=false;b.textContent=old});

window.RhinoDocumentFlow=flow;
})();
