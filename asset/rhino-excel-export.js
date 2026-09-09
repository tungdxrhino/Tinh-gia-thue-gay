/*
  Rhino Cue Platform - Excel Export V20
  Requirement: asset/jszip.min.js must be loaded BEFORE this file.

  Public API:
    RhinoExcel.exportQuote(type, info, data)
    RhinoExcel.exportPromoRegistration(data?)

  Quote info object:
    { customer, club, address, seller, sellerPhone }
*/
(function (window, document) {
  'use strict';

  const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const COMPANY = 'CÔNG TY TNHH CARBON BILLIARDS';
  const LOGO_PATH = 'asset/logo.png';
  const EMU_PER_INCH = 914400;

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
  const dash = '-';

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

  function textOrDash(v) {
    const s = String(v ?? '').trim();
    return s || dash;
  }

  function numOrDash(c, r, v, numberStyle = 7, dashStyle = 8) {
    if (v === null || v === undefined || v === '' || !Number.isFinite(Number(v))) return cell(c, r, dash, dashStyle);
    return cell(c, r, Number(v), numberStyle, 'n');
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

  function moneyText(v) {
    return new Intl.NumberFormat('vi-VN').format(Math.round(Number(v) || 0));
  }

  async function tryLogoBytes() {
    try {
      const res = await fetch(LOGO_PATH, { cache: 'no-store' });
      if (res.ok) return new Uint8Array(await res.arrayBuffer());
    } catch (_) {}

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
<fonts count="11">
  <font><sz val="10"/><name val="Aptos"/><color rgb="FF263944"/></font>
  <font><b/><sz val="10.5"/><name val="Aptos"/><color rgb="FFFFFFFF"/></font>
  <font><b/><sz val="18"/><name val="Aptos Display"/><color rgb="FF0B3B59"/></font>
  <font><b/><sz val="12"/><name val="Aptos"/><color rgb="FF0B3B59"/></font>
  <font><b/><sz val="10"/><name val="Aptos"/><color rgb="FF0B3B59"/></font>
  <font><b/><sz val="10"/><name val="Aptos"/><color rgb="FF138879"/></font>
  <font><i/><sz val="9.5"/><name val="Aptos"/><color rgb="FF6E7E88"/></font>
  <font><b/><sz val="10.5"/><name val="Aptos"/><color rgb="FFB55D20"/></font>
  <font><b/><sz val="11"/><name val="Aptos"/><color rgb="FFFFFFFF"/></font>
  <font><b/><sz val="10.5"/><name val="Aptos"/><color rgb="FF138879"/></font>
  <font><sz val="10"/><name val="Aptos"/><color rgb="FF4B5E67"/></font>
</fonts>
<fills count="7">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF0B3B59"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFF6F9FA"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFE9F7F4"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFF4E8"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF138879"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="3">
  <border/>
  <border><left style="thin"><color rgb="FFD9E4E7"/></left><right style="thin"><color rgb="FFD9E4E7"/></right><top style="thin"><color rgb="FFD9E4E7"/></top><bottom style="thin"><color rgb="FFD9E4E7"/></bottom></border>
  <border><bottom style="thin"><color rgb="FF138879"/></bottom></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="18">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
  <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="10" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="3" fontId="7" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="9" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="9" fillId="4" borderId="2" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="8" fillId="6" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="3" fontId="8" fillId="6" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  }

  function infoRow(r, leftLabel, leftValue, rightLabel, rightValue) {
    const cells = [
      cell(0, r, leftLabel, 17), cell(1, r, textOrDash(leftValue), 0)
    ];
    if (rightLabel) cells.push(cell(3, r, rightLabel, 17), cell(4, r, textOrDash(rightValue), rightValue ? 0 : 8));
    return row(r, cells, 21);
  }

  function tableHeader(r) {
    return row(r, [
      cell(0, r, 'NỘI DUNG', 3),
      cell(1, r, 'THÔNG TIN', 3),
      cell(2, r, 'QUY CÁCH', 3),
      cell(3, r, 'ĐƠN GIÁ', 3),
      cell(4, r, 'SỐ LƯỢNG', 3),
      cell(5, r, 'THÀNH TIỀN', 3)
    ], 25);
  }

  function tableRow(r, values, opts = {}) {
    const promo = !!opts.promo;
    const styleText = promo ? 9 : 5;
    const styleNum = promo ? 10 : 7;
    const styleDash = promo ? 11 : 8;
    return row(r, [
      cell(0, r, textOrDash(values[0]), values[0] ? styleText : styleDash),
      cell(1, r, textOrDash(values[1]), values[1] ? styleText : styleDash),
      cell(2, r, textOrDash(values[2]), values[2] ? styleText : styleDash),
      numOrDash(3, r, values[3], styleNum, styleDash),
      numOrDash(4, r, values[4], styleNum, styleDash),
      numOrDash(5, r, values[5], styleNum, styleDash)
    ], opts.height || 24);
  }

  function quoteCommonRows(info, title) {
    const now = new Date();
    return [
      row(1, [cell(3, 1, COMPANY, 1)], 22),
      row(2, [cell(0, 2, title, 2)], 30),
      infoRow(4, 'Khách hàng', info.customer, 'Ngày báo giá', now.toLocaleDateString('vi-VN')),
      infoRow(5, 'CLB / Quán', info.club, 'Người lập báo giá', info.seller),
      infoRow(6, 'Địa chỉ', info.address, 'Liên hệ', info.sellerPhone)
    ];
  }

  function rentalPricing(data) {
    const n = Math.max(0, Number(data.n || 0));
    const unit = Math.max(0, Number(data.unit || 0));
    const rent = Math.max(0, Number(data.rent || 0));
    const use = Math.max(0, Number(data.use || 0));
    let paid = Number(data.paid);
    if (!Number.isFinite(paid) || paid <= 0) {
      paid = n > 0 && unit > 0 ? Math.max(1, Math.round(rent / (n * unit))) : use;
    }
    const gift = Math.max(0, Number(data.gift || 0), use - paid);
    const discount = gift > 0 ? Math.round(unit * n * gift) : 0;
    const gross = rent + discount;
    const deposit = Math.max(0, Number(data.deposit || 0));
    const finalPay = Math.max(0, Number(data.initial ?? (rent + deposit)));
    return { n, unit, rent, use, paid, gift, discount, gross, deposit, finalPay };
  }

  function purchasePricing(data) {
    const qty = Math.max(0, Number(data.qty || 0));
    const baseUnit = Math.max(0, Number(data.base ?? data.unit ?? 0));
    const finalUnit = Math.max(0, Number(data.unit ?? baseUnit));
    const discountPerUnit = Math.max(0, baseUnit - finalUnit);
    const gross = baseUnit * qty;
    const discount = discountPerUnit * qty;
    const finalPay = Math.max(0, Number(data.total ?? (gross - discount)));
    return { qty, baseUnit, finalUnit, discountPerUnit, gross, discount, finalPay };
  }

  function purchaseTierText(qty) {
    if (qty >= 51) return 'Từ 51 gậy';
    if (qty >= 11) return '11–50 gậy';
    return '5–10 gậy';
  }

  function buildQuoteSheet(info, type, data, hasLogo) {
    const title = type === 'rental' ? 'BÁO GIÁ THUÊ GẬY CLB' : 'BÁO GIÁ MUA GẬY CLB';
    const rows = quoteCommonRows(info, title);
    const merges = ['D1:F1', 'A2:F2', 'B4:C4', 'E4:F4', 'B5:C5', 'E5:F5', 'B6:C6', 'E6:F6'];
    let r = 8;
    rows.push(tableHeader(r));
    r++;

    if (type === 'rental') {
      const p = rentalPricing(data);
      rows.push(tableRow(r, [
        'Thuê gậy CLB',
        data.name || 'Phương án thuê',
        `${p.use} tháng sử dụng`,
        p.unit,
        p.n,
        p.gross
      ]));
      r++;

      if (p.discount > 0) {
        rows.push(tableRow(r, [
          '',
          'Khuyến mại',
          `Tặng ${p.gift} tháng sử dụng`,
          p.discount,
          '',
          p.rent
        ], { promo: true }));
        r++;
      }

      rows.push(tableRow(r, [
        '',
        'Ưu đãi',
        `Gậy dự phòng ${Math.round((Number(data.backupPct || 0)) * 100)}%: ${Number(data.backup || 0)} gậy`,
        '', '', ''
      ], { promo: true }));
      r++;

      if (p.deposit > 0) {
        rows.push(tableRow(r, [
          'Tiền cọc',
          'Hoàn lại sau đối soát',
          'Theo điều kiện hợp đồng',
          p.deposit,
          1,
          p.finalPay
        ]));
        r++;
      }

      rows.push(row(r + 1, [cell(0, r + 1,
        `Chi phí thực tế sau ưu đãi: ${moneyText(data.effective)} VND / gậy thực nhận / tháng. Tổng thực nhận ${Number(data.total || 0)} gậy; thời gian sử dụng ${p.use} tháng.`, 14)], 30));
      merges.push(`A${r + 1}:F${r + 1}`);
      rows.push(row(r + 2, [cell(0, r + 2,
        'Giá thuê đã gồm VAT. Các nội dung áp dụng theo chính sách và hợp đồng tại thời điểm ký kết.', 13)], 28));
      merges.push(`A${r + 2}:F${r + 2}`);
      rows.push(row(r + 3, [cell(0, r + 3,
        'Cảm ơn Quý khách đã quan tâm đến sản phẩm và giải pháp của Rhino Cue Platform.', 13)], 24));
      merges.push(`A${r + 3}:F${r + 3}`);
      const totalRow = r + 5;
      rows.push(row(totalRow, [cell(0, totalRow, 'TỔNG GIÁ TRỊ KHÁCH HÀNG CHI TRẢ', 15), cell(5, totalRow, p.finalPay, 16, 'n')], 28));
      merges.push(`A${totalRow}:E${totalRow}`);
      r = totalRow;
    } else {
      const p = purchasePricing(data);
      rows.push(tableRow(r, [
        `Gậy Rhino ${data.model || ''}`,
        'Giá gốc theo bảng giá',
        purchaseTierText(p.qty),
        p.baseUnit,
        p.qty,
        p.gross
      ]));
      r++;

      if (p.discount > 0) {
        rows.push(tableRow(r, [
          '',
          'Khuyến mại',
          'Ưu đãi khách đang sử dụng dịch vụ thuê',
          p.discountPerUnit,
          p.qty,
          p.finalPay
        ], { promo: true }));
        r++;
      }

      rows.push(row(r + 1, [cell(0, r + 1,
        `Quy cách giá áp dụng theo số lượng ${p.qty} gậy. Đơn giá sau ưu đãi: ${moneyText(p.finalUnit)} VND / gậy.`, 14)], 28));
      merges.push(`A${r + 1}:F${r + 1}`);
      rows.push(row(r + 2, [cell(0, r + 2,
        'Cảm ơn Quý khách đã quan tâm đến sản phẩm Rhino. Báo giá áp dụng theo chính sách tại thời điểm xác nhận đơn hàng.', 13)], 28));
      merges.push(`A${r + 2}:F${r + 2}`);
      const totalRow = r + 4;
      rows.push(row(totalRow, [cell(0, totalRow, 'TỔNG GIÁ TRỊ KHÁCH HÀNG CHI TRẢ', 15), cell(5, totalRow, p.finalPay, 16, 'n')], 28));
      merges.push(`A${totalRow}:E${totalRow}`);
      r = totalRow;
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:F${r}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="19"/><cols>
<col min="1" max="1" width="21" customWidth="1"/><col min="2" max="2" width="24" customWidth="1"/><col min="3" max="3" width="26" customWidth="1"/><col min="4" max="4" width="15" customWidth="1"/><col min="5" max="5" width="13" customWidth="1"/><col min="6" max="6" width="18" customWidth="1"/>
</cols>
<sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>
<pageMargins left="0.3" right="0.3" top="0.45" bottom="0.45" header="0.2" footer="0.2"/><pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0" paperSize="9"/>${hasLogo ? '<drawing r:id="rId1"/>' : ''}
</worksheet>`;
  }

  function buildPromoRegistrationSheet(data, hasLogo, createdAt) {
    const rows = [
      row(1, [cell(3, 1, COMPANY, 1)], 22),
      row(2, [cell(0, 2, 'PHIẾU ĐĂNG KÝ TRẢI NGHIỆM GẬY 30 NGÀY', 2)], 30),
      infoRow(4, 'Thời gian lập phiếu', createdAt, 'Hạn đăng ký CT', '15/12/2026'),
      infoRow(5, 'Khách hàng', data.customer, 'Số điện thoại', data.phone),
      infoRow(6, 'CLB / Quán', data.club, 'Email', data.email || ''),
      infoRow(7, 'Địa chỉ', data.address, 'Giao nhận dự kiến', formatDateVi(data.delivery)),
      infoRow(8, 'Kết thúc dự kiến', data.endDate, 'Thời hạn', '30 ngày')
    ];
    const merges = ['D1:F1', 'A2:F2', 'B4:C4', 'E4:F4', 'B5:C5', 'E5:F5', 'B6:C6', 'E6:F6', 'B7:C7', 'E7:F7', 'B8:C8', 'E8:F8'];
    let r = 10;
    rows.push(tableHeader(r)); r++;
    rows.push(tableRow(r, ['Rhino R-68', 'Trải nghiệm 30 ngày', 'Dòng gậy mới dành cho CLB', '', Number(data.r68 || 0), 0])); r++;
    rows.push(tableRow(r, ['Rhino R-88', 'Trải nghiệm 30 ngày', 'Dòng gậy mới dành cho CLB', '', Number(data.r88 || 0), 0])); r++;
    rows.push(tableRow(r, ['', 'Khuyến mại', 'Miễn phí 100% phí trải nghiệm', '', '', 0], { promo: true })); r++;
    rows.push(tableRow(r, ['', 'Ưu đãi', 'Không đặt cọc', '', '', 0], { promo: true })); r++;
    rows.push(row(r + 1, [cell(0, r + 1, 'Chương trình áp dụng theo điều kiện xác nhận của Rhino tại thời điểm triển khai.', 13)], 28));
    merges.push(`A${r + 1}:F${r + 1}`);
    rows.push(row(r + 2, [cell(0, r + 2, 'Cảm ơn Quý CLB đã đăng ký trải nghiệm sản phẩm Rhino.', 13)], 24));
    merges.push(`A${r + 2}:F${r + 2}`);
    const totalRow = r + 4;
    rows.push(row(totalRow, [cell(0, totalRow, 'TỔNG GIÁ TRỊ KHÁCH HÀNG CHI TRẢ', 15), cell(5, totalRow, 0, 16, 'n')], 28));
    merges.push(`A${totalRow}:E${totalRow}`);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:F${totalRow}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="19"/><cols>
<col min="1" max="1" width="21" customWidth="1"/><col min="2" max="2" width="24" customWidth="1"/><col min="3" max="3" width="26" customWidth="1"/><col min="4" max="4" width="15" customWidth="1"/><col min="5" max="5" width="13" customWidth="1"/><col min="6" max="6" width="18" customWidth="1"/>
</cols>
<sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>
<pageMargins left="0.3" right="0.3" top="0.45" bottom="0.45" header="0.2" footer="0.2"/><pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0" paperSize="9"/>${hasLogo ? '<drawing r:id="rId1"/>' : ''}
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
    const width = Math.round(1.71 * EMU_PER_INCH);
    const height = Math.round(0.45 * EMU_PER_INCH);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>70000</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>50000</xdr:rowOff></xdr:from><xdr:ext cx="${width}" cy="${height}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Rhino Logo"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
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
    if (!info || !info.customer || !info.club || !info.address || !info.seller || !info.sellerPhone) {
      throw new Error('Thiếu thông tin báo giá hoặc người lập báo giá.');
    }
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
    version: '2.0.0',
    exportQuote,
    exportPromoRegistration,
    readRegistrationFromPage,
    safeName
  };

  window.RhinoExcel = API;
  window.createQuoteXlsx = function (type, info, data) { return API.exportQuote(type, info, data); };
  window.createPromoRegistrationXlsx = function () { return API.exportPromoRegistration(); };

})(window, document);
