/*
  Rhino Cue Platform - Excel Export V31
  Requirement: asset/jszip.min.js must be loaded BEFORE this file.

  Public API:
    RhinoExcel.exportQuote(type, info, data)
    RhinoExcel.exportPromoRegistration(data?)

  Quote info object:
    { customer, club, address, customerEmail?, seller, sellerPhone, showroom }
    Rental data may also include: { model: 'R-68' | 'R-88' }
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
    return `<row r="${r}"${height ? ` ht="${Number(height).toFixed(2)}" customHeight="1"` : ''}>${cells.join('')}</row>`;
  }

  // Ước lượng số dòng sau khi Wrap Text. XLSX không tự AutoFit ổn định với
  // hàng có custom height / merged cells, nên exporter chủ động tính chiều cao.
  function wrappedLineCount(value, widthChars) {
    const text = String(value ?? '');
    if (!text) return 1;
    const usable = Math.max(6, Number(widthChars || 10) * 0.86);
    return text.split(/\r?\n/).reduce((sum, paragraph) => {
      const p = paragraph.trim();
      if (!p) return sum + 1;
      // Ước lượng theo từ để tránh đánh giá thấp các dòng có nhiều khoảng trắng.
      let lines = 1, used = 0;
      for (const word of p.split(/\s+/)) {
        const w = Math.max(1, word.length);
        if (used === 0) used = w;
        else if (used + 1 + w <= usable) used += 1 + w;
        else { lines += Math.max(1, Math.ceil(w / usable)); used = w % usable; }
      }
      return sum + lines;
    }, 0);
  }

  function autoRowHeight(values, widths, minHeight = 24, lineHeight = 15, maxHeight = 120) {
    const vals = Array.isArray(values) ? values : [values];
    const ws = Array.isArray(widths) ? widths : [widths];
    let maxLines = 1;
    vals.forEach((v, i) => {
      if (v === null || v === undefined || v === '') return;
      maxLines = Math.max(maxLines, wrappedLineCount(v, ws[i] || ws[ws.length - 1] || 12));
    });
    const estimated = 8 + maxLines * lineHeight;
    return Math.min(maxHeight, Math.max(minHeight, estimated));
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
<fonts count="12">
  <font><sz val="11"/><name val="Aptos"/><color rgb="FF263944"/></font>
  <font><b/><sz val="11"/><name val="Aptos"/><color rgb="FFFFFFFF"/></font>
  <font><b/><sz val="20"/><name val="Aptos Display"/><color rgb="FF0B3B59"/></font>
  <font><b/><sz val="13"/><name val="Aptos"/><color rgb="FF0B3B59"/></font>
  <font><b/><sz val="11"/><name val="Aptos"/><color rgb="FF0B3B59"/></font>
  <font><b/><sz val="11"/><name val="Aptos"/><color rgb="FF138879"/></font>
  <font><i/><sz val="11"/><name val="Aptos"/><color rgb="FF6E7E88"/></font>
  <font><b/><sz val="11"/><name val="Aptos"/><color rgb="FFE06A1A"/></font>
  <font><b/><sz val="11"/><name val="Aptos"/><color rgb="FFFFFFFF"/></font>
  <font><b/><sz val="11"/><name val="Aptos"/><color rgb="FF138879"/></font>
  <font><sz val="11"/><name val="Aptos"/><color rgb="FF4B5E67"/></font>
  <font><b/><sz val="11"/><name val="Aptos"/><color rgb="FF0B3B59"/></font>
</fonts>
<fills count="7">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF0B3B59"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFF7F9FA"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFEEF9F6"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFF6EC"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF138879"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="3">
  <border/>
  <border><left style="thin"><color rgb="FFD9E4E7"/></left><right style="thin"><color rgb="FFD9E4E7"/></right><top style="thin"><color rgb="FFD9E4E7"/></top><bottom style="thin"><color rgb="FFD9E4E7"/></bottom></border>
  <border><bottom style="thin"><color rgb="FF138879"/></bottom></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="25">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="0"/></xf>
  <xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="10" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="3" fontId="0" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="10" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="3" fontId="7" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="10" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="9" fillId="4" borderId="2" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="3" fontId="7" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="3" fontId="0" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="0" fontId="11" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="11" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="3" fontId="9" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="3" fontId="4" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  <xf numFmtId="3" fontId="7" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  }

  function infoRow(r, leftLabel, leftValue, rightLabel, rightValue) {
    const cells = [
      cell(0, r, leftLabel, 17), cell(1, r, textOrDash(leftValue), 0)
    ];
    if (rightLabel) cells.push(cell(3, r, rightLabel, 17), cell(4, r, textOrDash(rightValue), rightValue ? 0 : 8));
    return row(r, cells, autoRowHeight([leftValue, rightValue], [43.75, 35.0], 21, 14));
  }

  function infoRowWide(r, leftLabel, leftValue, rightLabel, rightValue) {
    const cells = [cell(0, r, leftLabel, 17), cell(1, r, textOrDash(leftValue), 0)];
    if (rightLabel) cells.push(cell(5, r, rightLabel, 17), cell(6, r, textOrDash(rightValue), rightValue ? 0 : 8));
    return row(r, cells, autoRowHeight([leftValue, rightValue], [58.25, 30.75], 21, 14));
  }

  function tableHeader(r) {
    return row(r, [
      cell(0, r, 'NỘI DUNG', 3),
      cell(1, r, 'THÔNG TIN', 3),
      cell(2, r, 'QUY CÁCH', 3),
      cell(3, r, 'ĐƠN GIÁ', 3),
      cell(4, r, 'SỐ LƯỢNG', 3),
      cell(5, r, 'THÀNH TIỀN', 3)
    ], 28);
  }

  function pilotTableHeader(r) {
    return row(r, [
      cell(0, r, 'NỘI DUNG', 3),
      cell(1, r, 'THÔNG TIN', 3),
      cell(2, r, 'QUY CÁCH', 3),
      cell(3, r, 'GIÁ THUÊ', 3),
      cell(4, r, 'SỐ LƯỢNG', 3),
      cell(5, r, 'THÀNH TIỀN', 3)
    ], 27.95);
  }

  function rentalTableHeader(r) {
    return row(r, [
      cell(0, r, 'NỘI DUNG', 3),
      cell(1, r, 'THÔNG TIN', 3),
      cell(2, r, 'QUY CÁCH', 3),
      cell(3, r, 'GIÁ THUÊ', 3),
      cell(4, r, 'SỐ LƯỢNG', 3),
      cell(5, r, 'TỔNG GIÁ / THÁNG', 3),
      cell(6, r, 'SỐ THÁNG', 3),
      cell(7, r, 'THÀNH TIỀN', 3)
    ], 27.95);
  }

  function tableRow(r, values, opts = {}) {
    const promo = !!opts.promo;
    const styleText = promo ? 9 : 5;
    const styleNum = promo ? 18 : 7;
    const styleDash = promo ? 11 : 8;
    const lastStyle = promo ? 22 : styleNum;
    return row(r, [
      cell(0, r, textOrDash(values[0]), values[0] ? styleText : styleDash),
      cell(1, r, textOrDash(values[1]), values[1] ? styleText : styleDash),
      cell(2, r, textOrDash(values[2]), values[2] ? styleText : styleDash),
      numOrDash(3, r, values[3], styleNum, styleDash),
      numOrDash(4, r, values[4], styleNum, styleDash),
      numOrDash(5, r, values[5], lastStyle, styleDash)
    ], Math.max(opts.height || 24, autoRowHeight(values, [21, 43.75, 17.5, 18.125, 12.75, 18], 24, 15)));
  }

  function rentalTableRow(r, values, opts = {}) {
    const promo = !!opts.promo;
    const styleText = promo ? 9 : 5;
    const styleNum = promo ? 18 : 7;
    const styleDash = promo ? 11 : 8;
    const lastStyle = promo ? 22 : (opts.payableLast ? 10 : styleNum);
    return row(r, [
      cell(0, r, textOrDash(values[0]), values[0] ? styleText : styleDash),
      cell(1, r, textOrDash(values[1]), values[1] ? styleText : styleDash),
      cell(2, r, textOrDash(values[2]), values[2] ? styleText : styleDash),
      numOrDash(3, r, values[3], styleNum, styleDash),
      numOrDash(4, r, values[4], styleNum, styleDash),
      numOrDash(5, r, values[5], styleNum, styleDash),
      numOrDash(6, r, values[6], styleNum, styleDash),
      numOrDash(7, r, values[7], lastStyle, styleDash)
    ], Math.max(opts.height || 25, autoRowHeight(values, [25.75, 30.75, 12.75, 14.75, 12.75, 18, 12.75, 18], 25, 15)));
  }

  function quoteCommonRows(info, title, wide = true, subtitle = '') {
    const now = new Date();
    return [
      row(1, [cell(4, 1, COMPANY, 1)], 45),
      row(2, [cell(0, 2, title, 2)], 30),
      row(3, [cell(0, 3, subtitle || '(Báo giá có giá trị 14 ngày kể từ ngày báo giá)', 21)], autoRowHeight(subtitle || '(Báo giá có giá trị 14 ngày kể từ ngày báo giá)', 140, 20.1, 14)),
      infoRowWide(4, 'Khách hàng', info.customer, 'Thời gian lập phiếu', now.toLocaleString('vi-VN')),
      infoRowWide(5, 'Câu lạc bộ / Quán', info.club, 'Người lập báo giá', info.seller),
      infoRowWide(6, 'Địa chỉ câu lạc bộ', info.address, 'Liên hệ', info.sellerPhone),
      infoRowWide(7, 'Email khách hàng', info.customerEmail || '', 'Showroom', info.showroom || '')
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

  function rentalPlanLabel(data, p) {
    let label='';
    if (Number(data.term) === 1) label = 'Kỳ 01 tháng';
    else if (Number(data.term) === 2) label = 'Kỳ 02 tháng';
    else if (Number(data.term) === 3) label = 'Gói 03 tháng';
    else if (Number(data.term) === 6) label = 'Gói 06 tháng';
    else if (Number(data.term) === 12) label = 'Gói 12 tháng';
    else {
      const raw = String(data.name || 'Phương án thuê')
        .replace(/\s*[·-]\s*dùng\s*\d+\s*tháng.*$/i, '')
        .replace(/\s*[·-]\s*tặng(?:\s*thêm)?\s*\d+\s*tháng(?:\s*sử dụng)?.*$/i, '')
        .trim();
      label = raw || `Phương án ${p.paid || ''} tháng`;
    }
    return data.postPilot ? `${label} · Sau Pilot` : label;
  }

  function rentalSummaryRow(r, label, value, total = false) {
    if (total) return row(r, [cell(0, r, label, 15), cell(7, r, value, 16, 'n')], 27.95);
    return row(r, [cell(0, r, label, 14), cell(7, r, value, 22, 'n')], 27);
  }

  function purchaseSummaryRow(r, label, value, total = false) {
    if (total) return row(r, [cell(0, r, label, 15), cell(5, r, value, 16, 'n')], 28);
    return row(r, [cell(0, r, label, 14), cell(5, r, value, 22, 'n')], 27);
  }

  function buildQuoteSheet(info, type, data, hasLogo) {
    const isRental = type === 'rental';
    const title = isRental ? 'BÁO GIÁ THUÊ GẬY CLB' : 'BÁO GIÁ MUA GẬY CLB';
    const rows = quoteCommonRows(info, title, true, '(Báo giá có giá trị 14 ngày kể từ ngày báo giá)');
    const merges = ['E1:H1', 'A2:H2', 'A3:H3', 'B4:D4', 'G4:H4', 'B5:D5', 'G5:H5', 'B6:D6', 'G6:H6', 'B7:D7', 'G7:H7'];
    let r = 9;

    if (isRental) {
      rows.push(rentalTableHeader(r));
      r++;
      const p = rentalPricing(data);
      const model = data.model === 'R-88' ? 'R-88' : 'R-68';
      const monthlyTotal = Math.round(p.unit * p.n);
      const backupCount = Math.max(0, Number(data.backup || 0));
      const backupPct = Math.round(Math.max(0, Number(data.backupPct || 0)) * 100);
      const giftValue = p.gift > 0 ? Math.round(monthlyTotal * p.gift) : 0;
      const backupMonthlyValue = Math.round(p.unit * backupCount);
      const backupValue = Math.round(backupMonthlyValue * p.use);
      const promoTotal = giftValue + backupValue;

      rows.push(rentalTableRow(r, [
        `Thuê gậy CLB ${model}`,
        rentalPlanLabel(data, p),
        'Gậy',
        p.unit,
        p.n,
        monthlyTotal,
        p.paid,
        p.rent
      ], { payableLast: true, height: 24.95 }));
      r++;

      if (p.gift > 0) {
        rows.push(rentalTableRow(r, [
          'Khuyến mại',
          `Tặng ${p.gift} tháng sử dụng`,
          'Tháng',
          monthlyTotal,
          '',
          monthlyTotal,
          p.gift,
          giftValue
        ], { promo: true, height: 24.95 }));
        r++;
      }

      if (backupCount > 0) {
        rows.push(rentalTableRow(r, [
          'Ưu đãi',
          `Gậy dự phòng ${backupPct}%`,
          'Gậy',
          p.unit,
          backupCount,
          backupMonthlyValue,
          p.use,
          backupValue
        ], { promo: true, height: 24.95 }));
        r++;
      }

      if (p.deposit > 0) {
        rows.push(rentalTableRow(r, [
          'Tiền cọc',
          'Hoàn lại sau đối soát',
          '-',
          p.deposit,
          1,
          '',
          '',
          p.deposit
        ], { height: 24.95 }));
        r++;
      }

      rows.push(rentalSummaryRow(r, 'TỔNG ƯU ĐÃI TƯƠNG ĐƯƠNG', promoTotal, false));
      merges.push(`A${r}:G${r}`);
      r++;
      rows.push(rentalSummaryRow(r, 'TỔNG TIỀN KHÁCH HÀNG CHI TRẢ', p.finalPay, true));
      merges.push(`A${r}:G${r}`);
      r++;

      // Footer giữ đúng cấu trúc mẫu: 1 dòng trắng → ghi chú → điều khoản → 1 dòng trắng → cảm ơn → trân trọng.
      rows.push(row(r, [], 15));
      rows.push(row(r + 1, [cell(0, r + 1,
        `Chi phí thực tế sau ưu đãi: ${moneyText(data.effective)} VND / gậy thực nhận / tháng. Tổng thực nhận ${Number(data.total || 0)} gậy; thời gian sử dụng ${p.use} tháng.`, 14)], autoRowHeight(`Chi phí thực tế sau ưu đãi: ${moneyText(data.effective)} VND / gậy thực nhận / tháng. Tổng thực nhận ${Number(data.total || 0)} gậy; thời gian sử dụng ${p.use} tháng.`, 145, 30, 15)));
      merges.push(`A${r + 1}:H${r + 1}`);
      rows.push(row(r + 2, [cell(0, r + 2,
        'Giá thuê đã gồm VAT. Các nội dung áp dụng theo chính sách và hợp đồng tại thời điểm ký kết.', 13)], autoRowHeight('Giá thuê đã gồm VAT. Các nội dung áp dụng theo chính sách và hợp đồng tại thời điểm ký kết.', 145, 27, 14)));
      merges.push(`A${r + 2}:H${r + 2}`);
      rows.push(row(r + 3, [], 15));
      rows.push(row(r + 4, [cell(0, r + 4,
        'Cảm ơn Quý khách đã quan tâm đến sản phẩm và giải pháp của Rhino Cue Platform.', 19)], autoRowHeight('Cảm ơn Quý khách đã quan tâm đến sản phẩm và giải pháp của Rhino Cue Platform.', 145, 30, 15)));
      merges.push(`A${r + 4}:H${r + 4}`);
      rows.push(row(r + 5, [cell(0, r + 5,
        `Trân trọng — ${COMPANY}`, 20)], autoRowHeight(`Trân trọng — ${COMPANY}`, 145, 18, 14)));
      merges.push(`A${r + 5}:H${r + 5}`);
      r = r + 5;

      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:H${r}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="13.5"/><cols>
<col min="1" max="1" width="25.75" customWidth="1"/><col min="2" max="2" width="30.75" customWidth="1"/><col min="3" max="3" width="12.75" customWidth="1"/><col min="4" max="4" width="14.75" customWidth="1"/><col min="5" max="5" width="12.75" customWidth="1"/><col min="6" max="6" width="18" customWidth="1"/><col min="7" max="7" width="12.75" customWidth="1"/><col min="8" max="8" width="18" customWidth="1"/>
</cols>
<sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>
<pageMargins left="0.3" right="0.3" top="0.45" bottom="0.45" header="0.2" footer="0.2"/><pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0" paperSize="9"/>${hasLogo ? '<drawing r:id="rId1"/>' : ''}
</worksheet>`;
    }

    // Báo giá mua: dùng cùng header/footer 8 cột của mẫu, giữ nguyên nội dung bảng mua hiện tại.
    rows.push(tableHeader(r));
    r++;
    const p = purchasePricing(data);
    rows.push(tableRow(r, [
      `Gậy Rhino ${data.model || ''}`,
      `Giá theo bảng: ${purchaseTierText(p.qty)}`,
      'Gậy',
      p.baseUnit,
      p.qty,
      p.gross
    ]));
    r++;

    if (p.discount > 0) {
      rows.push(tableRow(r, [
        'Khuyến mại',
        'Ưu đãi khách đang sử dụng dịch vụ thuê',
        'Gậy',
        p.discountPerUnit,
        p.qty,
        p.discount
      ], { promo: true }));
      r++;
      rows.push(purchaseSummaryRow(r, 'TỔNG ƯU ĐÃI TƯƠNG ĐƯƠNG', p.discount, false));
      merges.push(`A${r}:E${r}`);
      r++;
    }

    rows.push(purchaseSummaryRow(r, 'TỔNG TIỀN KHÁCH HÀNG CHI TRẢ', p.finalPay, true));
    merges.push(`A${r}:E${r}`);
    r++;
    rows.push(row(r, [], 15));
    rows.push(row(r + 1, [cell(0, r + 1,
      `Đơn giá sau ưu đãi: ${moneyText(p.finalUnit)} VND / gậy. Báo giá áp dụng cho số lượng ${p.qty} gậy.`, 14)], autoRowHeight(`Đơn giá sau ưu đãi: ${moneyText(p.finalUnit)} VND / gậy. Báo giá áp dụng cho số lượng ${p.qty} gậy.`, 145, 30, 15)));
    merges.push(`A${r + 1}:H${r + 1}`);
    rows.push(row(r + 2, [cell(0, r + 2,
      'Giá bán đã gồm VAT. Các nội dung áp dụng theo chính sách tại thời điểm xác nhận.', 13)], autoRowHeight('Giá bán đã gồm VAT. Các nội dung áp dụng theo chính sách tại thời điểm xác nhận.', 145, 27, 14)));
    merges.push(`A${r + 2}:H${r + 2}`);
    rows.push(row(r + 3, [], 15));
    rows.push(row(r + 4, [cell(0, r + 4,
      'Cảm ơn Quý khách đã quan tâm đến sản phẩm Rhino.', 19)], autoRowHeight('Cảm ơn Quý khách đã quan tâm đến sản phẩm Rhino.', 145, 30, 15)));
    merges.push(`A${r + 4}:H${r + 4}`);
    rows.push(row(r + 5, [cell(0, r + 5,
      `Trân trọng — ${COMPANY}`, 20)], autoRowHeight(`Trân trọng — ${COMPANY}`, 145, 18, 14)));
    merges.push(`A${r + 5}:H${r + 5}`);
    r = r + 5;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:H${r}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="13.5"/><cols>
<col min="1" max="1" width="25.75" customWidth="1"/><col min="2" max="2" width="30.75" customWidth="1"/><col min="3" max="3" width="12.75" customWidth="1"/><col min="4" max="4" width="14.75" customWidth="1"/><col min="5" max="5" width="12.75" customWidth="1"/><col min="6" max="6" width="18" customWidth="1"/><col min="7" max="7" width="12.75" customWidth="1"/><col min="8" max="8" width="18" customWidth="1"/>
</cols>
<sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>
<pageMargins left="0.3" right="0.3" top="0.45" bottom="0.45" header="0.2" footer="0.2"/><pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0" paperSize="9"/>${hasLogo ? '<drawing r:id="rId1"/>' : ''}
</worksheet>`;
  }


  function pilotPlayingUnitPrice(qty) {
    const n = Math.max(0, Number(qty || 0));
    if (n <= 0) return 0;
    if (n <= 9) return 89000;
    if (n <= 20) return 87000;
    if (n <= 49) return 85000;
    if (n <= 99) return 83000;
    return 81000;
  }

  function pilotSummaryRow(r, label, value, kind = 'base') {
    if (kind === 'promo') return row(r, [cell(0, r, label, 14), cell(5, r, value, 22, 'n')], 27);
    if (kind === 'payable') return row(r, [cell(0, r, label, 15), cell(5, r, value, 16, 'n')], 27.95);
    return row(r, [cell(0, r, label, 4), cell(5, r, value, 24, 'n')], 27);
  }

  function buildPromoRegistrationSheet(data, hasLogo, createdAt) {
    const tables = Math.max(0, Number(data.tables || 0));
    const recommendedPlaying = Math.max(0, Number(data.recommendedPlaying || (tables * 2)));
    const requestedPlaying = Math.max(0, Number(data.requestedPlaying || ((Number(data.r68 || 0)) + (Number(data.r88 || 0)))));
    const playing = requestedPlaying > 0 ? requestedPlaying : recommendedPlaying;
    const r68 = Math.max(0, Number(data.r68 || 0));
    const r88 = Math.max(0, Number(data.r88 || 0));
    const breakCues = Math.max(0, Number(data.breakCues ?? 0));
    const jumpCues = Math.max(0, Number(data.jumpCues ?? 0));
    const legacyBreakJump = Math.max(0, Number(data.breakJump ?? data.recommendedBreakJump ?? (tables ? Math.ceil(tables / 3) : 0)));
    const breakJump = (breakCues + jumpCues) > 0 ? (breakCues + jumpCues) : legacyBreakJump;

    // Pilot 30 ngày: quy đổi giá trị theo đơn giá thuê lẻ 01 tháng.
    // Gậy đánh dùng bậc giá 01 tháng; gậy phá/nhảy = 100.000/cây/tháng.
    const playingUnit = pilotPlayingUnitPrice(playing);
    const specialUnit = 100000;
    const playingValue = Math.round(playingUnit * playing);
    const specialValue = Math.round(specialUnit * breakJump);
    const grossValue = playingValue + specialValue;
    const promoValue = grossValue; // Miễn phí 100%.
    const finalPay = 0;

    // Cấu trúc CHỐT theo file mẫu Pilot người dùng cung cấp: A:F.
    const rows = [
      row(1, [cell(3, 1, COMPANY, 1)], 45),
      row(2, [cell(0, 2, 'PHIẾU ĐĂNG KÝ PILOT 30 NGÀY – HỆ THỐNG GẬY CARBON CHO CLB', 2)], 30),
      row(3, [cell(0, 3, '(Chương trình trải nghiệm 30 ngày miễn phí · không đặt cọc)', 21)], autoRowHeight('(Chương trình trải nghiệm 30 ngày miễn phí · không đặt cọc)', 120, 20.1, 14)),
      infoRow(4, 'Khách hàng', data.customer, 'Thời gian lập phiếu', createdAt),
      infoRow(5, 'Câu lạc bộ / Quán', data.club, 'Người lập phiếu', data.seller),
      infoRow(6, 'Địa chỉ câu lạc bộ', data.address, 'Liên hệ', data.sellerPhone),
      infoRow(7, 'Số bàn', tables ? tables : '', 'Showroom', data.showroom),
      infoRow(8, 'Số điện thoại khách', data.phone, 'Giao nhận dự kiến', formatDateVi(data.delivery)),
      infoRow(9, 'Email khách hàng', data.email || '', 'Kết thúc dự kiến', data.endDate)
    ];
    const merges = [
      'D1:F1','A2:F2','A3:F3',
      'E4:F4','E5:F5','E6:F6','E7:F7','E8:F8','E9:F9'
    ];

    let r = 11;
    rows.push(pilotTableHeader(r)); r++;

    if (requestedPlaying > 0) {
      if (r68 > 0) {
        rows.push(tableRow(r, ['Gậy đánh', 'Gậy CLB R-68', 'Gậy', playingUnit, r68, playingUnit * r68], { height: 24 })); r++;
      }
      if (r88 > 0) {
        rows.push(tableRow(r, ['Gậy đánh', 'Gậy CLB R-88', 'Gậy', playingUnit, r88, playingUnit * r88], { height: 24 })); r++;
      }
    } else if (playing > 0) {
      rows.push(tableRow(r, ['Gậy đánh', 'Gậy CLB R-68 / R-88', 'Gậy', playingUnit, playing, playingValue], { height: 24 })); r++;
    }

    if ((breakCues + jumpCues) > 0) {
      if (breakCues > 0) {
        rows.push(tableRow(r, ['Gậy phá', 'Gậy phá CLB', 'Gậy', specialUnit, breakCues, specialUnit * breakCues], { height: 24 })); r++;
      }
      if (jumpCues > 0) {
        rows.push(tableRow(r, ['Gậy nhảy CLB', 'Gậy nhảy CLB', 'Gậy', specialUnit, jumpCues, specialUnit * jumpCues], { height: 24 })); r++;
      }
    } else if (breakJump > 0) {
      rows.push(tableRow(r, ['Gậy phá / nhảy', 'Cấu hình Pilot – Carbon duyệt', 'Gậy', specialUnit, breakJump, specialValue], { height: 24 })); r++;
    }

    rows.push(tableRow(r, ['Dịch vụ bảo dưỡng', 'Bao gồm trong thời gian Pilot', 'Gói', 0, 1, 0], { height: 24 })); r++;

    rows.push(pilotSummaryRow(r, 'TỔNG GIÁ TRỊ TRẢI NGHIỆM', grossValue, 'base'));
    merges.push(`A${r}:E${r}`); r++;
    rows.push(pilotSummaryRow(r, 'TỔNG ƯU ĐÃI · MIỄN PHÍ 100%', promoValue, 'promo'));
    merges.push(`A${r}:E${r}`); r++;
    rows.push(pilotSummaryRow(r, 'KHÁCH HÀNG CHI TRẢ', finalPay, 'payable'));
    merges.push(`A${r}:E${r}`); r++;

    rows.push(row(r, [], 15));
    rows.push(row(r + 1, [cell(0, r + 1,
      'Số lượng thực tế do Carbon duyệt theo quy mô và tình hình vận hành của từng CLB. Phỏng vấn online là bước xét duyệt cuối trước khi bàn giao.', 14)], autoRowHeight('Số lượng thực tế do Carbon duyệt theo quy mô và tình hình vận hành của từng CLB. Phỏng vấn online là bước xét duyệt cuối trước khi bàn giao.', 120, 33.95, 15)));
    merges.push(`A${r + 1}:F${r + 1}`);
    rows.push(row(r + 2, [cell(0, r + 2,
      'Sau khi hoàn thành Pilot, CLB đủ điều kiện chuyển đổi có thể áp dụng gói 6+2 hoặc 12+6 theo chính sách tại thời điểm xác nhận.', 13)], autoRowHeight('Sau khi hoàn thành Pilot, CLB đủ điều kiện chuyển đổi có thể áp dụng gói 6+2 hoặc 12+6 theo chính sách tại thời điểm xác nhận.', 120, 30, 15)));
    merges.push(`A${r + 2}:F${r + 2}`);
    rows.push(row(r + 3, [], 15));
    rows.push(row(r + 4, [cell(0, r + 4,
      'Cảm ơn Quý CLB đã đăng ký chương trình Pilot 30 ngày của Carbon Billiards.', 19)], autoRowHeight('Cảm ơn Quý CLB đã đăng ký chương trình Pilot 30 ngày của Carbon Billiards.', 120, 30, 15)));
    merges.push(`A${r + 4}:F${r + 4}`);
    rows.push(row(r + 5, [cell(0, r + 5,
      `Trân trọng — ${COMPANY}`, 20)], autoRowHeight(`Trân trọng — ${COMPANY}`, 120, 18, 14)));
    merges.push(`A${r + 5}:F${r + 5}`);
    r = r + 5;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:F${r}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/><cols>
<col min="1" max="1" width="21" customWidth="1"/><col min="2" max="2" width="43.75" customWidth="1"/><col min="3" max="3" width="17.5" customWidth="1"/><col min="4" max="4" width="18.125" customWidth="1"/><col min="5" max="5" width="12.75" customWidth="1"/><col min="6" max="6" width="18" customWidth="1"/>
</cols>
<sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>
<pageMargins left="0.3" right="0.3" top="0.45" bottom="0.45" header="0.2" footer="0.2"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0" paperSize="9"/>${hasLogo ? '<drawing r:id="rId1"/>' : ''}
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
    if (!info || !info.customer || !info.club || !info.address || !info.seller || !info.sellerPhone || !info.showroom) {
      throw new Error('Thiếu thông tin báo giá hoặc người lập báo giá.');
    }
    if (!data) throw new Error('Thiếu dữ liệu báo giá.');

    const blob = await createWorkbookBlob('BÁO GIÁ', (hasLogo) => buildQuoteSheet(info, type, data, hasLogo));
    const nameBase = type === 'rental' ? 'Bao_gia_thue_gay_Rhino' : 'Bao_gia_mua_gay_Rhino';
    const model = data.model ? `${safeName(data.model)}_` : '';
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
    const tables = Math.max(0, parseInt(el('regTables')?.value || '0', 10) || 0);
    const useR68 = !!el('regUseR68')?.checked, useR88 = !!el('regUseR88')?.checked;
    const useBreak = !!el('regUseBreak')?.checked, useJump = !!el('regUseJump')?.checked;
    const r68 = useR68 ? Math.max(0, parseInt(el('regR68')?.value || '0', 10) || 0) : 0;
    const r88 = useR88 ? Math.max(0, parseInt(el('regR88')?.value || '0', 10) || 0) : 0;
    const breakCues = useBreak ? Math.max(0, parseInt(el('regBreakQty')?.value || '0', 10) || 0) : 0;
    const jumpCues = useJump ? Math.max(0, parseInt(el('regJumpQty')?.value || '0', 10) || 0) : 0;
    const requestedPlaying = r68 + r88;
    const recommendedPlaying = tables * 2;
    const recommendedBreakJump = tables ? Math.ceil(tables / 3) : 0;
    const playing = requestedPlaying;
    const breakJump = breakCues + jumpCues;
    const delivery = el('regDelivery')?.value || '';
    const end = delivery ? addDays(delivery, 30) : null;
    const sellerId = el('regSellerSelect')?.value || '';
    const sellerObj = typeof window.getUserById === 'function' ? window.getUserById(sellerId) : null;
    return {
      customer, club, address, phone, email, tables, useR68, useR88, r68, r88,
      requestedPlaying, recommendedPlaying, playing, useBreak, useJump, breakCues, jumpCues,
      breakJump, recommendedBreakJump, total: playing + breakJump, delivery,
      endDate: end ? end.toLocaleDateString('vi-VN') : '', duration: 30,
      sellerId, seller: sellerObj?.name || '', sellerPhone: sellerObj?.phone || '', showroom: sellerObj?.showroom || ''
    };
  }

  async function exportPromoRegistration(data) {
    const registration = data || readRegistrationFromPage();
    if (!registration.customer || !registration.club || !registration.address || !registration.phone || !registration.tables || !registration.delivery || !registration.seller || !registration.sellerPhone || !registration.showroom) {
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
    version: '3.1.0',
    exportQuote,
    exportPromoRegistration,
    readRegistrationFromPage,
    safeName
  };

  window.RhinoExcel = API;
  window.createQuoteXlsx = function (type, info, data) { return API.exportQuote(type, info, data); };
  window.createPromoRegistrationXlsx = function () { return API.exportPromoRegistration(); };

})(window, document);
