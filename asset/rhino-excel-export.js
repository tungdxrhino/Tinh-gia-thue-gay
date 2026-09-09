/*
  Rhino Cue Platform - Excel Export
  File: asset/rhino-excel-export.js
  Requirement: asset/jszip.min.js must be loaded BEFORE this file.

  Public API:
    RhinoExcel.exportQuote(type, info, data)
    RhinoExcel.exportPromoRegistration(data?)

  Compatibility bridge:
    Overrides the page functions createQuoteXlsx() and
    createPromoRegistrationXlsx() so the existing HTML buttons
    automatically use this external file without changing their logic.
*/
(function (window, document) {
  'use strict';

  const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const COMPANY = 'CÔNG TY TNHH CARBON BILLIARDS';
  const LOGO_PATH = 'asset/logo.png';

  function requireJSZip() {
    if (typeof window.JSZip === 'undefined') {
      throw new Error('Thiếu JSZip. Hãy tải asset/jszip.min.js trước asset/rhino-excel-export.js');
    }
    return window.JSZip;
  }

  const xmlEsc = (v) => String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

  const safeName = (v) => String(v || 'Bao_gia')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/[^a-zA-Z0-9_-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 54) || 'Bao_gia';

  function columnName(index) {
    let n = index + 1;
    let out = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  const a1 = (c, r) => `${columnName(c)}${r}`;

  function cell(c, r, value, style = 0, type = 's') {
    const ref = a1(c, r);
    if (type === 'n') {
      const num = Number(value);
      return `<c r="${ref}" s="${style}" t="n"><v>${Number.isFinite(num) ? num : 0}</v></c>`;
    }
    return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(value)}</t></is></c>`;
  }

  function row(r, cells, height) {
    return `<row r="${r}"${height ? ` ht="${height}" customHeight="1"` : ''}>${cells.join('')}</row>`;
  }

  function formatDateVi(date) {
    if (!date) return '';
    const d = date instanceof Date ? date : new Date(`${date}T00:00:00`);
    return Number.isNaN(d.getTime()) ? '' : d.toLocaleDateString('vi-VN');
  }

  function addDays(value, days) {
    const d = new Date(`${value}T00:00:00`);
    if (Number.isNaN(d.getTime())) return null;
    d.setDate(d.getDate() + days);
    return d;
  }

  function pad2(n) { return String(n).padStart(2, '0'); }

  function timeStamp(date = new Date()) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}_${pad2(date.getHours())}${pad2(date.getMinutes())}`;
  }

  async function tryLogoBytes() {
    try {
      const res = await fetch(LOGO_PATH, { cache: 'no-store' });
      if (res.ok) return new Uint8Array(await res.arrayBuffer());
    } catch (_) {}

    // Fallback for local/static hosting where fetch(file://...) can be restricted.
    try {
      const img = document.querySelector('.brand img');
      if (!img || !img.complete || !img.naturalWidth) return null;
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
      return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
    } catch (_) {
      return null;
    }
  }

  function workbookStyles() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="6">
<font><sz val="10"/><name val="Aptos"/><color rgb="FF20303A"/></font>
<font><b/><sz val="11"/><name val="Aptos"/><color rgb="FFFFFFFF"/></font>
<font><b/><sz val="16"/><name val="Aptos Display"/><color rgb="FF0B3B59"/></font>
<font><b/><sz val="12"/><name val="Aptos"/><color rgb="FF0B3B59"/></font>
<font><b/><sz val="11"/><name val="Aptos"/><color rgb="FF138879"/></font>
<font><i/><sz val="9"/><name val="Aptos"/><color rgb="FF6E7E88"/></font>
</fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0B3B59"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEAF7F4"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF5F8F8"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="3"><border/><border><left style="thin"><color rgb="FFD9E4E7"/></left><right style="thin"><color rgb="FFD9E4E7"/></right><top style="thin"><color rgb="FFD9E4E7"/></top><bottom style="thin"><color rgb="FFD9E4E7"/></bottom></border><border><bottom style="thin"><color rgb="FF138879"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="3" fontId="4" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="2" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  }

  function buildQuoteSheet(info, type, data, hasLogo) {
    const rows = [];
    rows.push(row(1, [cell(2, 1, COMPANY, 1)], 26));
    rows.push(row(2, [cell(2, 2, type === 'rental' ? 'BÁO GIÁ THUÊ GẬY CLB' : 'BÁO GIÁ MUA GẬY CLB', 2)], 24));
    rows.push(row(4, [
      cell(0, 4, 'Khách hàng', 2), cell(1, 4, info.customer, 0),
      cell(3, 4, 'Ngày báo giá', 2), cell(4, 4, new Date().toLocaleDateString('vi-VN'), 0)
    ], 20));
    rows.push(row(5, [cell(0, 5, 'CLB / Quán', 2), cell(1, 5, info.club, 0)], 20));
    rows.push(row(6, [cell(0, 6, 'Địa chỉ', 2), cell(1, 6, info.address, 0)], 26));

    let r = 8;
    if (type === 'rental') {
      rows.push(row(r, [cell(0, r, 'NỘI DUNG BÁO GIÁ', 3), cell(1, r, 'THÔNG TIN', 3), cell(3, r, 'GIÁ TRỊ', 3)], 24));
      r++;
      const items = [
        ['Phương án', data.name || '', ''],
        ['Số gậy tính phí', `${data.n ?? 0} gậy`, ''],
        ['Đơn giá / gậy / tháng', '', data.unit],
        ['Số tháng thanh toán', `${data.paid ?? data.use ?? 0} tháng`, ''],
        ['Tổng tiền thuê', '', data.rent],
        ['Gậy dự phòng', `${data.backup ?? 0} gậy (+${Math.round((data.backupPct || 0) * 100)}%)`, ''],
        ['Tổng gậy thực nhận', `${data.total ?? 0} gậy`, ''],
        ['Thời gian sử dụng thực tế', `${data.use ?? 0} tháng${data.gift ? ` (+${data.gift} tháng tặng)` : ''}`, ''],
        ['Tiền cọc', '', data.deposit],
        ['Thanh toán ban đầu', '', data.initial],
        ['Chi phí thực tế / gậy / tháng', '', data.effective]
      ];
      items.forEach((it, i) => {
        const money = it[2] !== '';
        rows.push(row(r, [
          cell(0, r, it[0], i === 10 ? 7 : 4),
          cell(1, r, it[1], i === 10 ? 10 : 4),
          money ? cell(3, r, it[2], i === 10 ? 6 : 5, 'n') : cell(3, r, '', 4)
        ], 22));
        r++;
      });
      rows.push(row(r + 1, [cell(0, r + 1, 'Giá thuê đã gồm VAT. Các nội dung áp dụng theo chính sách và hợp đồng tại thời điểm ký kết.', 8)], 28));
      r += 3;
    } else {
      rows.push(row(r, [cell(0, r, 'HẠNG MỤC', 3), cell(1, r, 'SỐ LƯỢNG', 3), cell(3, r, 'ĐƠN GIÁ', 3), cell(4, r, 'THÀNH TIỀN', 3)], 24));
      r++;
      rows.push(row(r, [
        cell(0, r, `Gậy Rhino ${data.model || ''}`, 4),
        cell(1, r, `${data.qty ?? 0} gậy`, 4),
        cell(3, r, data.unit, 5, 'n'),
        cell(4, r, data.total, 6, 'n')
      ], 25));
      r += 2;
    }

    rows.push(row(r, [cell(0, r, 'Cảm ơn Quý khách đã quan tâm đến sản phẩm và giải pháp của Rhino Cue Platform.', 9)], 22));
    rows.push(row(r + 1, [cell(0, r + 1, `Trân trọng — ${COMPANY}`, 8)], 22));

    const end = r + 1;
    const merges = ['C1:F1', 'C2:F2', 'B4:C4', 'E4:F4', 'B5:F5', 'B6:F6'];
    if (type === 'rental') {
      for (let rr = 8; rr <= 19; rr++) merges.push(`B${rr}:C${rr}`, `D${rr}:F${rr}`);
      merges.push('A21:F21', 'A23:F23', 'A24:F24');
    } else {
      merges.push('B8:C8', 'E8:F8', 'B9:C9', 'E9:F9', 'A11:F11', 'A12:F12');
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:F${end}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="25" customWidth="1"/><col min="2" max="3" width="16" customWidth="1"/><col min="4" max="4" width="18" customWidth="1"/><col min="5" max="6" width="17" customWidth="1"/></cols>
<sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>${hasLogo ? '<drawing r:id="rId1"/>' : ''}
<pageMargins left="0.35" right="0.35" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0" paperSize="9"/>
</worksheet>`;
  }

  function buildPromoRegistrationSheet(data, hasLogo, createdAt) {
    const rows = [];
    rows.push(row(1, [cell(2, 1, COMPANY, 1)], 26));
    rows.push(row(2, [cell(2, 2, 'PHIẾU ĐĂNG KÝ TRẢI NGHIỆM GẬY 30 NGÀY', 2)], 24));
    rows.push(row(4, [cell(0, 4, 'Thời gian lập phiếu', 2), cell(1, 4, createdAt, 0), cell(3, 4, 'Hạn đăng ký CT', 2), cell(4, 4, '15/12/2026', 0)], 21));
    rows.push(row(5, [cell(0, 5, 'Khách hàng', 2), cell(1, 5, data.customer, 0), cell(3, 5, 'Số điện thoại', 2), cell(4, 5, data.phone, 0)], 21));
    rows.push(row(6, [cell(0, 6, 'CLB / Quán', 2), cell(1, 6, data.club, 0), cell(3, 6, 'Email', 2), cell(4, 6, data.email || 'Không có', 0)], 21));
    rows.push(row(7, [cell(0, 7, 'Địa chỉ', 2), cell(1, 7, data.address, 0)], 26));
    rows.push(row(8, [cell(0, 8, 'Giao nhận dự kiến', 2), cell(1, 8, formatDateVi(data.delivery), 0), cell(3, 8, 'Kết thúc dự kiến', 2), cell(4, 8, data.endDate, 0)], 21));
    rows.push(row(10, [cell(0, 10, 'SẢN PHẨM', 3), cell(1, 10, 'SỐ LƯỢNG', 3), cell(3, 10, 'ĐƠN GIÁ CT', 3), cell(4, 10, 'THÀNH TIỀN', 3)], 24));
    rows.push(row(11, [cell(0, 11, 'Rhino R-68', 4), cell(1, 11, `${data.r68 || 0} gậy`, 4), cell(3, 11, 0, 5, 'n'), cell(4, 11, 0, 5, 'n')], 23));
    rows.push(row(12, [cell(0, 12, 'Rhino R-88', 4), cell(1, 12, `${data.r88 || 0} gậy`, 4), cell(3, 12, 0, 5, 'n'), cell(4, 12, 0, 5, 'n')], 23));
    rows.push(row(14, [cell(0, 14, 'TỔNG SỐ GẬY', 7), cell(1, 14, `${data.total || 0} gậy`, 10), cell(3, 14, 'KHUYẾN MẠI', 7), cell(4, 14, '100%', 10)], 23));
    rows.push(row(15, [cell(0, 15, 'Phí trải nghiệm', 4), cell(1, 15, '0 VND', 4), cell(3, 15, 'Tiền cọc', 4), cell(4, 15, '0 VND', 4)], 22));
    rows.push(row(16, [cell(0, 16, 'TỔNG THANH TOÁN', 7), cell(1, 16, '0 VND', 6)], 25));
    rows.push(row(18, [cell(0, 18, 'Thời hạn trải nghiệm: 30 ngày kể từ ngày giao nhận dự kiến. Chương trình áp dụng theo điều kiện xác nhận của Rhino tại thời điểm triển khai.', 8)], 34));
    rows.push(row(20, [cell(0, 20, 'Cảm ơn Quý CLB đã đăng ký trải nghiệm sản phẩm Rhino.', 9)], 22));
    rows.push(row(21, [cell(0, 21, `Trân trọng — ${COMPANY}`, 8)], 22));

    const merges = [
      'C1:F1','C2:F2','B4:C4','E4:F4','B5:C5','E5:F5','B6:C6','E6:F6','B7:F7','B8:C8','E8:F8',
      'B10:C10','E10:F10','B11:C11','E11:F11','B12:C12','E12:F12','B14:C14','E14:F14','B15:C15','E15:F15',
      'B16:F16','A18:F18','A20:F20','A21:F21'
    ];

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:F21"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="3" width="17" customWidth="1"/><col min="4" max="4" width="19" customWidth="1"/><col min="5" max="6" width="18" customWidth="1"/></cols>
<sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>${hasLogo ? '<drawing r:id="rId1"/>' : ''}
<pageMargins left="0.35" right="0.35" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0" paperSize="9"/>
</worksheet>`;
  }

  function contentTypes(hasLogo) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${hasLogo ? '<Default Extension="png" ContentType="image/png"/>' : ''}<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${hasLogo ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : ''}</Types>`;
  }

  function rootRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  }

  function workbookXml(sheetName) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="${xmlEsc(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  }

  function workbookRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  }

  function drawingXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>100000</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>100000</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Rhino Logo"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
  }

  function sheetDrawingRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`;
  }

  function drawingRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.png"/></Relationships>`;
  }

  async function createWorkbookBlob(sheetName, sheetXml) {
    const JSZip = requireJSZip();
    const zip = new JSZip();
    const logo = await tryLogoBytes();
    const hasLogo = !!logo;

    // Rebuild the sheet because it needs to know whether a drawing relationship exists.
    const finalSheetXml = typeof sheetXml === 'function' ? sheetXml(hasLogo) : sheetXml;

    zip.file('[Content_Types].xml', contentTypes(hasLogo));
    zip.folder('_rels').file('.rels', rootRels());
    zip.folder('xl').file('workbook.xml', workbookXml(sheetName));
    zip.folder('xl').folder('_rels').file('workbook.xml.rels', workbookRels());
    zip.folder('xl').file('styles.xml', workbookStyles());
    zip.folder('xl').folder('worksheets').file('sheet1.xml', finalSheetXml);

    if (hasLogo) {
      zip.folder('xl').folder('worksheets').folder('_rels').file('sheet1.xml.rels', sheetDrawingRels());
      zip.folder('xl').folder('drawings').file('drawing1.xml', drawingXml());
      zip.folder('xl').folder('drawings').folder('_rels').file('drawing1.xml.rels', drawingRels());
      zip.folder('xl').folder('media').file('logo.png', logo);
    }

    return zip.generateAsync({ type: 'blob', mimeType: MIME_XLSX, compression: 'DEFLATE' });
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  async function exportQuote(type, info, data) {
    if (!['rental', 'purchase'].includes(type)) throw new Error('Loại báo giá không hợp lệ.');
    if (!info || !info.customer || !info.club || !info.address) throw new Error('Thiếu thông tin khách hàng.');
    if (!data) throw new Error('Thiếu dữ liệu báo giá.');

    const blob = await createWorkbookBlob('BÁO GIÁ', (hasLogo) => buildQuoteSheet(info, type, data, hasLogo));
    const nameBase = type === 'rental' ? 'Bao_gia_thue_gay_Rhino' : 'Bao_gia_mua_gay_Rhino';
    const model = type === 'purchase' && data.model ? `${safeName(data.model)}_` : '';
    const fileName = `${nameBase}_${model}${safeName(info.club)}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    downloadBlob(blob, fileName);
    return fileName;
  }

  function readRegistrationFromPage() {
    const el = (id) => document.getElementById(id);
    const customer = el('regCustomer')?.value.trim() || '';
    const club = el('regClub')?.value.trim() || '';
    const address = el('regAddress')?.value.trim() || '';
    const phone = el('regPhone')?.value.trim() || '';
    const email = el('regEmail')?.value.trim() || '';
    const r68 = Math.max(0, parseInt(el('regR68')?.value || '0', 10) || 0);
    const r88 = Math.max(0, parseInt(el('regR88')?.value || '0', 10) || 0);
    const delivery = el('regDelivery')?.value || '';
    const end = delivery ? addDays(delivery, 30) : null;
    return {
      customer, club, address, phone, email, r68, r88,
      total: r68 + r88,
      delivery,
      endDate: end ? end.toLocaleDateString('vi-VN') : '',
      duration: 30
    };
  }

  async function exportPromoRegistration(data) {
    const registration = data || readRegistrationFromPage();
    if (!registration.customer || !registration.club || !registration.address || !registration.phone || !registration.delivery) {
      throw new Error('Thiếu trường bắt buộc của phiếu đăng ký trải nghiệm.');
    }

    const now = new Date();
    const createdAt = now.toLocaleString('vi-VN');
    const blob = await createWorkbookBlob('ĐĂNG KÝ TRẢI NGHIỆM', (hasLogo) => buildPromoRegistrationSheet(registration, hasLogo, createdAt));
    const fileName = `Dang_ky_trai_nghiem_Rhino_${safeName(registration.club)}_${timeStamp(now)}.xlsx`;
    downloadBlob(blob, fileName);
    return fileName;
  }

  const API = {
    version: '1.0.0',
    exportQuote,
    exportPromoRegistration,
    readRegistrationFromPage,
    safeName
  };

  window.RhinoExcel = API;

  // Compatibility with the current V17 HTML.
  // Existing buttons already call these function names.
  window.createQuoteXlsx = function (type, info, data) {
    return API.exportQuote(type, info, data);
  };

  window.createPromoRegistrationXlsx = function () {
    return API.exportPromoRegistration();
  };

})(window, document);
