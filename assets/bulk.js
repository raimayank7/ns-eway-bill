/* NS bulk: many invoices (or e-way bills) at once from an Excel or CSV sheet, made into the portal's JSON file by einv.js.
   <section id="bulk" data-bulk="einv"> = e-invoice site, data-bulk="ewb" = e-way bill site (with the Part B vehicle file).
   One sheet row per item line; rows with the same document number make one document. Our template's headings are English
   (it is a data file, like the printed bills). Nothing leaves the browser: the user uploads the file on the portal themselves.
   E-way bill keys and codes: NIC "EWB JSON ATTRIBUTES - Version (1.0.0621)" and "Vehicle Updation Attributes - Version (1.0.0621)",
   docs.ewaybillgst.gov.in › Bulk Generation Tools (read 5 Nov 2026). */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, num = NS.num, r2 = NS.r2, t = NS.t;
  var EWB_VERSION = '1.0.0621';
  var SUBS = [[1, 'Supply'], [3, 'Export'], [4, 'Job Work'], [5, 'For Own Use'], [8, 'Others'], [9, 'SKD/CKD/Lots'], [10, 'Line Sales'], [11, 'Recipient Not Known'], [12, 'Exhibition or Fairs']];
  /* outward sub supply type → document types the portal accepts (NIC Validations, Table 2) */
  var DOCS_FOR = { 1: ['INV', 'BIL'], 3: ['INV', 'BIL'], 4: ['CHL'], 5: ['CHL'], 8: ['CHL', 'OTH'], 9: ['INV', 'BIL', 'CHL'], 10: ['CHL'], 11: ['CHL', 'OTH'], 12: ['CHL'] };
  var TO_SELF = { 5: 1, 10: 1, 11: 1, 12: 1 };   /* goods stay with the sender: "To GSTIN" = own GSTIN */
  var MODES = { road: 1, rail: 2, air: 3, ship: 4, 'road cum ship': 4 };
  var REASONS = [4, 1, 2, 3];   /* Part B: first time, break down, transhipment, others */

  /* columns: key, heading in the template, example, aliases (headings written in other ways) */
  var COMMON = [
    ['no', 'Invoice No', 'INV/26-27/101', 'document no|doc no|bill no|invoice number|docno'],
    ['date', 'Invoice Date', '05/11/2026', 'document date|doc date|bill date|docdate'],
    ['bgstin', 'Buyer GSTIN', '27AAPFU0939F1ZV', 'customer gstin|gstin of buyer|to gstin|recipient gstin|gstin'],
    ['bname', 'Buyer Name', 'Shree Traders', 'customer name|buyer legal name|party name|to name|recipient name|legal name'],
    ['baddr', 'Buyer Address', '12 Market Road', 'customer address|to address|address'],
    ['bcity', 'Buyer City', 'Pune', 'customer city|to place|city|place'],
    ['bpin', 'Buyer PIN', '411001', 'customer pin|to pincode|pincode|pin code|pin'],
    ['bstate', 'Buyer State Code', '27', 'customer state|to state|state code|state'],
    ['pos', 'Place of Supply', '27', 'pos|place of supply state'],
    ['desc', 'Item Description', 'Cotton yarn 40s', 'item|product|product name|description|item name'],
    ['hsn', 'HSN/SAC', '52051200', 'hsn|hsn code|sac|hsn sac'],
    ['qty', 'Quantity', '100', 'qty'],
    ['unit', 'Unit', 'KGS', 'uqc|uom'],
    ['rate', 'Rate', '250', 'price|unit price|rate per unit'],
    ['disc', 'Discount %', '0', 'discount|disc'],
    ['gst', 'GST %', '5', 'gst rate|tax rate|rate of tax|igst %|gst']
  ];
  var EINV = [['sup', 'Supply Type', 'B2B', 'supply type code|suptyp'], ['rc', 'Reverse Charge', 'N', 'reverse charge y/n|regrev']];
  var EWB = [
    ['dtype', 'Document Type', 'INV', 'doc type|doctype'], ['sub', 'Sub Supply Type', 'Supply', 'sub type|subsupplytype|reason for movement'], ['subdesc', 'Sub Supply Description', '', 'sub supply desc'],
    ['dist', 'Distance km', '0', 'distance|approximate distance|km'], ['mode', 'Transport Mode', 'Road', 'mode|trans mode'], ['veh', 'Vehicle No', 'MH12AB1234', 'vehicle|vehicle number'],
    ['vtype', 'Vehicle Type', 'R', 'odc'], ['tid', 'Transporter ID', '', 'transporter gstin|transin'], ['tname', 'Transporter Name', '', 'transporter'],
    ['tdoc', 'Transport Doc No', '', 'lr no|rr no|transport document no|lr'], ['tdate', 'Transport Doc Date', '', 'lr date|transport document date'],
    ['shaddr', 'Ship To Address', '', 'delivery address|dispatch to address'], ['shcity', 'Ship To City', '', 'delivery city|ship to place'], ['shpin', 'Ship To PIN', '', 'delivery pin|ship to pincode'], ['shstate', 'Ship To State Code', '', 'delivery state|ship to state']
  ];
  /* what each column takes, for the "How to fill" sheet (English, like the template) */
  var HELP = {
    no: 'Up to 16 letters, digits, / or -. Repeat it on every row of the same invoice (one row per item).',
    date: 'DD/MM/YYYY. Not a future date.',
    sup: 'B2B (sale to a registered business), SEZWP / SEZWOP (SEZ with / without IGST), EXPWP / EXPWOP (export with / without IGST), DEXP (deemed export).',
    rc: 'Y if the buyer pays the tax under reverse charge, else N.',
    bgstin: '15 characters. Leave empty for an export (and in an e-way bill for a buyer without GSTIN).',
    bname: 'Legal name as registered.', baddr: 'Building, street, area.', bcity: 'At least 3 letters.', bpin: '6 digits. Exports: the e-invoice uses 999999 by itself.',
    bstate: 'Two-digit GST state code (03 Punjab, 07 Delhi, 27 Maharashtra...) or the state name. Exports: 96.',
    pos: 'State code of the place of supply. Empty = the buyer\'s state.',
    desc: 'What you sold.', hsn: 'HSN or SAC code. E-invoice: 6 or 8 digits. E-way bill: 4, 6 or 8 digits.',
    qty: 'Number.', unit: 'GST unit code: NOS, PCS, KGS, LTR, MTR, BOX, SET, PAC... (other words become OTH).',
    rate: 'Price per unit before GST.', disc: 'Per cent off the line, if any.', gst: 'The GST rate: 0, 0.1, 0.25, 1.5, 3, 5, 12, 18 or 40.',
    dtype: 'INV tax invoice, BIL bill of supply, CHL delivery challan, OTH others.',
    sub: 'Why the goods move: Supply, Export, Job Work, For Own Use, Others, SKD/CKD/Lots, Line Sales, Recipient Not Known, Exhibition or Fairs.',
    subdesc: 'Only for Sub Supply Type "Others": say what it is.',
    dist: 'Kilometres, up to 4000. 0 = the portal works it out from the PIN codes.',
    mode: 'Road, Rail, Air or Ship. Leave empty with no vehicle to make Part A only (add the vehicle later as Part B).',
    veh: 'Like MH12AB1234. Road needs a vehicle number or a transporter ID.', vtype: 'R regular, O over-dimensional cargo.',
    tid: '15-character GSTIN or TRANSIN of the transporter.', tname: 'Transporter\'s name.', tdoc: 'LR, RR or airway bill number (needed for rail and air).', tdate: 'DD/MM/YYYY.',
    shaddr: 'Only when the goods go to another place than the buyer address (bill to - ship to). For an export: the port.', shcity: 'City of that place.', shpin: '6-digit PIN of that place.', shstate: 'State code of that place.'
  };

  var box, mode, cols, batch = [], lastFile = '';
  function norm(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9%]/g, ''); }
  function digits(s) { return String(s || '').replace(/\D/g, ''); }
  function key(s) { return String(s || '').trim().toLowerCase(); }

  /* ---------- reading the sheet ---------- */
  function headerMap(row) {
    var map = {};
    row.forEach(function (h, i) {
      var n = norm(h);
      cols.forEach(function (c) {
        if (map[c[0]] != null) return;
        if (n === norm(c[1]) || c[3].split('|').some(function (a) { return norm(a) === n; })) map[c[0]] = i;
      });
    });
    return map;
  }
  function isoDate(v) {
    v = String(v || '').trim();
    var m;
    if (/^\d{5}(\.\d+)?$/.test(v)) return NS.xlsx.serialDate(v);
    if ((m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    if ((m = v.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/))) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    var mon = 'jan feb mar apr may jun jul aug sep oct nov dec'.split(' ');
    if ((m = v.toLowerCase().match(/^(\d{1,2})[\s\-\/]([a-z]{3})[a-z]*[\s\-\/,]+(\d{4})$/)) && mon.indexOf(m[2]) >= 0) return m[3] + '-' + ('0' + (mon.indexOf(m[2]) + 1)).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    return '';
  }
  function stateCode(v) {
    v = String(v || '').trim();
    if (!v) return '';
    if (/^\d{1,2}$/.test(v)) { var c = ('0' + v).slice(-2); return c === '99' ? '96' : NS.stateName(c) ? c : ''; }
    var m = v.match(/\((\d{2})\)/); if (m && NS.stateName(m[1])) return m[1];
    var n = norm(v), hit = NS.STATES.filter(function (s) { return norm(s[1]) === n; })[0];
    return hit ? hit[0] : /^(other|foreign|outside|export)/i.test(v) ? '96' : '';
  }
  function subCode(v) {
    v = String(v || '').trim();
    if (!v) return 1;
    if (/^\d+$/.test(v)) return +v;
    var n = norm(v), hit = SUBS.filter(function (s) { return norm(s[1]) === n || norm(s[1]).indexOf(n) === 0; })[0];
    return hit ? hit[0] : 0;
  }
  function modeCode(v) { v = String(v || '').trim().toLowerCase(); return !v ? 0 : /^\d$/.test(v) ? +v : MODES[v] || (/ship/.test(v) ? 4 : 0); }

  /* rows → documents (each with its own list of sheet problems) */
  function toDocs(rows, seller) {
    var at = 0;
    while (at < Math.min(rows.length, 10) && Object.keys(headerMap(rows[at])).length < 4) at++;
    var map = headerMap(rows[at] || []);
    if (map.no == null || map.desc == null && map.hsn == null) return { err: t('bk.nohead') };
    var get = function (r, k) { return map[k] == null ? '' : String(r[map[k]] == null ? '' : r[map[k]]).trim(); };
    var docs = [], byNo = {};
    rows.slice(at + 1).forEach(function (r, i) {
      var no = get(r, 'no').toUpperCase();
      if (!no || /^(example|e\.g\.)/i.test(no)) return;
      var line = at + i + 2, d = byNo[no];
      if (!d) {
        d = byNo[no] = { line: line, prob: [], d: {
          type: 'invoice', kind: 'tax', no: no, date: isoDate(get(r, 'date')), due: '', pos: stateCode(get(r, 'pos')), rc: /^y/i.test(get(r, 'rc')), ref: '',
          seller: JSON.parse(JSON.stringify(seller)), buyer: { name: get(r, 'bname'), addr: get(r, 'baddr'), city: get(r, 'bcity'), pin: digits(get(r, 'bpin')), state: stateCode(get(r, 'bstate')), gstin: get(r, 'bgstin').toUpperCase().replace(/\s/g, ''), phone: '' },
          ship: null, irn: null, items: [], round: box.querySelector('#bk-round').checked, comp: true, notes: '', terms: '', accent: '#0B3B5B', status: 'unpaid', copyLabel: 'orig', paper: 'a4', design: 'classic', copies: 1,
          einv: { sup: (get(r, 'sup') || 'B2B').toUpperCase(), aato: aato(), dist: get(r, 'dist') || '0', mode: '', veh: get(r, 'veh').toUpperCase().replace(/[\s-]/g, ''), vtype: (get(r, 'vtype') || 'R').toUpperCase().slice(0, 1),
            tid: get(r, 'tid').toUpperCase(), tname: get(r, 'tname'), tdoc: get(r, 'tdoc'), tdate: isoDate(get(r, 'tdate')), withEwb: false }
        } };
        docs.push(d);
        var x = d.d;
        if (!x.date) d.prob.push(t('bk.date', { r: line }));
        if (get(r, 'bstate') && !x.buyer.state) d.prob.push(t('bk.state', { r: line, v: get(r, 'bstate') }));
        if (mode === 'ewb' && x.buyer.gstin && !NS.gstinCheck(x.buyer.gstin).ok) d.prob.push(t('bk.gstin', { r: line, v: x.buyer.gstin }));
        if (mode === 'ewb') {
          var sub = subCode(get(r, 'sub')), dt = (get(r, 'dtype') || '').toUpperCase().slice(0, 3);
          if (!sub || !DOCS_FOR[sub]) d.prob.push(t('bk.sub', { r: line, v: get(r, 'sub') }));
          else {
            if (!dt) dt = DOCS_FOR[sub][0];
            if (DOCS_FOR[sub].indexOf(dt) < 0) d.prob.push(t('bk.dtype', { r: line, v: dt, ok: DOCS_FOR[sub].join(', ') }));
          }
          if (get(r, 'shaddr') || get(r, 'shpin')) x.ship = { name: x.buyer.name, addr: get(r, 'shaddr'), city: get(r, 'shcity'), pin: digits(get(r, 'shpin')), state: stateCode(get(r, 'shstate')) };
          x.einv.sub = sub; x.einv.dtype = dt; x.einv.subDesc = get(r, 'subdesc');
          if (sub === 8 && !x.einv.subDesc) d.prob.push(t('bk.subdesc', { r: line }));
          x.kind = dt === 'INV' ? 'tax' : 'supply';
          if (sub === 3) { x.buyer.gstin = ''; x.einv.sup = 'EXPWP'; }
          if (TO_SELF[sub]) { x.buyer.gstin = seller.gstin; if (!x.buyer.name) x.buyer.name = seller.name; }
          var mc = modeCode(get(r, 'mode'));
          if (get(r, 'mode') && !mc) d.prob.push(t('bk.mode', { r: line, v: get(r, 'mode') }));
          x.einv.mode = mc ? String(mc) : '';
          if (mc === 1 && !x.einv.veh && !x.einv.tid) d.prob.push(t('bk.road', { r: line }));
          if ((mc === 2 || mc === 3) && !(x.einv.tdoc && x.einv.tdate)) d.prob.push(t('bk.raildoc', { r: line }));
        } else if (['B2B', 'SEZWP', 'SEZWOP', 'EXPWP', 'EXPWOP', 'DEXP'].indexOf(x.einv.sup) < 0) d.prob.push(t('bk.sup', { r: line, v: x.einv.sup }));
        if (/^EXP/.test(x.einv.sup)) { x.buyer.gstin = ''; x.buyer.state = '96'; if (!x.pos) x.pos = '96'; }
      } else if (isoDate(get(r, 'date')) && isoDate(get(r, 'date')) !== d.d.date) d.prob.push(t('bk.twodates', { r: line }));
      var g = get(r, 'gst');
      d.d.items.push({ desc: get(r, 'desc'), hsn: digits(get(r, 'hsn')), qty: get(r, 'qty') || '1', unit: get(r, 'unit') || 'Nos', rate: get(r, 'rate').replace(/[₹,\s]/g, ''), disc: get(r, 'disc').replace('%', ''), gst: g === '' ? 0 : num(String(g).replace('%', '')) });
    });
    docs.forEach(function (x) {
      var d = x.d;
      x.tt = totals(d);
      if (mode === 'ewb' && d.kind !== 'tax' && x.tt.rows.some(function (r) { return r.g; })) x.prob.push(t('bk.zero', { v: d.einv.dtype }));
      var c = NS.einv.check(d, x.tt, mode);
      x.bad = x.prob.concat(c.bad.map(function (p) { return p[1]; }));
      x.note = c.note;
    });
    return { docs: docs, rows: rows.length - at - 1 };
  }
  /* the same arithmetic as the invoice maker (docs.js): GST per line on the value after discount, total rounded if ticked */
  function totals(d) {
    var origin = d.seller.state, dest = d.pos || d.buyer.state || origin, tax = d.kind === 'tax';
    var tt = { rows: [], taxable: 0, cgst: 0, sgst: 0, igst: 0, intra: !origin || !dest || origin === dest, sum: {} };
    d.items.forEach(function (it) {
      var q = num(it.qty), rate = num(it.rate);
      if (!it.desc && !rate) return;
      var taxable = r2(q * rate - q * rate * num(it.disc) / 100), g = tax ? num(it.gst) : num(it.gst), tx = r2(taxable * g / 100);
      tt.rows.push({ it: it, taxable: taxable, g: g, tax: tx, amount: r2(taxable + tx) });
      tt.taxable += taxable;
      if (tt.intra) { tt.cgst += tx / 2; tt.sgst += tx / 2; } else tt.igst += tx;
    });
    tt.taxable = r2(tt.taxable); tt.cgst = r2(tt.cgst); tt.sgst = r2(tt.sgst); tt.igst = r2(tt.igst);
    var total = r2(tt.taxable + tt.cgst + tt.sgst + tt.igst);
    tt.roundOff = d.round ? r2(Math.round(total) - total) : 0;
    tt.grand = r2(total + tt.roundOff);
    return tt;
  }
  function sig(d, grand) { return [String(d.no).toUpperCase(), d.date, d.seller.gstin, d.buyer.gstin, grand].join('|'); }

  /* ---------- screen ---------- */
  function aato() { var s = $('#bk-aato', box); return s ? s.value : ''; }
  function opt(v, l, cur) { return '<option value="' + esc(v) + '"' + (cur === v ? ' selected' : '') + '>' + esc(l) + '</option>'; }
  function render() {
    var ewb = mode === 'ewb', saved = ''; try { saved = localStorage.getItem('ns-aato') || ''; } catch (e) {}
    box.innerHTML =
      '<div class="bk-steps">' +
      '<div class="card"><h3>1. ' + esc(t('bk.s1')) + '</h3><div class="fg">' +
        '<div class="f wide"><label for="bk-biz">' + esc(t('bk.biz')) + '</label><select id="bk-biz"><option value="">' + esc(t('bk.loading')) + '</option></select><small class="hint">' + esc(t('bk.h.biz')) + ' <a href="' + hub() + '#biz" target="_blank" rel="noopener">' + esc(t('bk.addbiz')) + ' ›</a></small></div>' +
        (ewb ? '' : '<div class="f wide"><label for="bk-aato">' + esc(t('ei.aato')) + '</label><select id="bk-aato">' + ['', 'lt5', '5to10', '10p'].map(function (k) { return opt(k, t('ei.a.' + (k || 'none')), saved); }).join('') + '</select></div>') +
        '<div class="f wide"><label class="chk"><input id="bk-round" type="checkbox" checked> ' + esc(t('l.round')) + '</label></div>' +
      '</div></div>' +
      '<div class="card"><h3>2. ' + esc(t('bk.s2')) + '</h3><p class="hint">' + esc(t(ewb ? 'bk.h.s2e' : 'bk.h.s2')) + '</p>' +
        '<div class="row"><button type="button" class="btn" id="bk-tpl">' + esc(t('bk.tpl')) + '</button>' +
        '<label class="btn primary" for="bk-file">' + esc(t('bk.choose')) + '</label><input id="bk-file" type="file" accept=".xlsx,.csv,.txt,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv" hidden></div>' +
        '<div class="bk-out" id="bk-out" aria-live="polite"></div></div>' +
      '<div class="card"><h3>3. ' + esc(t('bk.s3')) + '</h3><p class="hint">' + esc(t(ewb ? 'bk.h.s3e' : 'bk.h.s3')) + '</p>' +
        '<div class="row"><label class="btn" for="bk-reply">' + esc(t('ei.b.load')) + '</label><input id="bk-reply" type="file" accept=".json,.xlsx,.csv,application/json" hidden></div>' +
        '<div class="bk-out" id="bk-rep" aria-live="polite"></div></div>' +
      (ewb ? partBHtml() : '') +
      '</div>';
    $('#bk-tpl', box).addEventListener('click', template);
    $('#bk-file', box).addEventListener('change', function (e) { var f = e.target.files[0]; e.target.value = ''; if (f) readSheet(f); });
    $('#bk-reply', box).addEventListener('change', function (e) { var f = e.target.files[0]; e.target.value = ''; if (f) readReply(f); });
    var a = $('#bk-aato', box); if (a) a.addEventListener('change', function () { try { localStorage.setItem('ns-aato', a.value); } catch (e) {} });
    if (ewb) wirePartB();
    fillBiz();
  }
  function hub() { return (window.NS_CFG && NS_CFG.vault ? new URL(NS_CFG.vault, location.href).origin : 'https://my.niveshsahayak.in') + '/'; }
  var bizList = [];
  function fillBiz() {
    var draft = null; try { draft = (JSON.parse(localStorage.getItem('ns-draft-invoice') || 'null') || {}).doc; } catch (e) {}
    NSStore.ready().then(function () { return NSStore.list('biz'); }).then(function (all) {
      bizList = all.filter(function (b) { return b.name; }).sort(function (a, b) { return a.name.localeCompare(b.name); });
      if (draft && draft.seller && draft.seller.name && !bizList.some(function (b) { return key(b.name) === key(draft.seller.name); })) bizList.unshift(Object.assign({ id: '_form' }, draft.seller));
      var sel = $('#bk-biz', box);
      sel.innerHTML = bizList.length ? bizList.map(function (b) { return opt(b.id, b.name + (b.gstin ? ' · ' + b.gstin : '') + (b.id === '_form' ? ' (' + t('bk.fromform') + ')' : ''), ''); }).join('') : '<option value="">' + esc(t('bk.nobiz')) + '</option>';
      var x = mode === 'ewb' && seller(), pl = $('#pb-place', box);
      if (x && pl && !pl.value) { pl.value = x.city || ''; $('#pb-state', box).value = x.state || ''; }
    }).catch(function () {});
  }
  function seller() {
    var id = $('#bk-biz', box).value, b = bizList.filter(function (x) { return x.id === id; })[0];
    if (!b) return null;
    var s = JSON.parse(JSON.stringify(b)); delete s.id; delete s.created; delete s.updated;
    ['name', 'addr', 'city', 'pin', 'state', 'gstin', 'phone', 'email'].forEach(function (k) { if (s[k] == null) s[k] = ''; });
    return s;
  }

  function template() {
    var head = cols.map(function (c) { return c[1]; });
    var ex1 = cols.map(function (c) { return c[2]; });
    var ex2 = ex1.slice(); ex2[cols.map(function (c) { return c[0]; }).indexOf('desc')] = 'Cotton yarn 60s'; ex2[cols.map(function (c) { return c[0]; }).indexOf('hsn')] = '52052400'; ex2[cols.map(function (c) { return c[0]; }).indexOf('rate')] = '310'; ex2[cols.map(function (c) { return c[0]; }).indexOf('qty')] = '40';
    var help = [['Column', 'What to type']].concat(cols.map(function (c) { return [c[1], HELP[c[0]] || '']; }),
      [['', ''], ['One row per item', 'Rows with the same ' + cols[0][1] + ' make one ' + (mode === 'ewb' ? 'e-way bill' : 'invoice') + '. Buyer details are read from its first row.'],
       ['Delete the examples', 'Replace the two example rows (INV/26-27/101) with your own.'],
       ['Save', 'Keep it as Excel (.xlsx), or File > Save As > CSV UTF-8. Then choose the file on ' + location.hostname + '.']]);
    var blob = NS.xlsx.write([[mode === 'ewb' ? 'E-way bills' : 'Invoices', [head, ex1, ex2]], ['How to fill', help, [24, 110], 'wrap']]);
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = mode === 'ewb' ? 'e-way-bill-template.xlsx' : 'e-invoice-template.xlsx';
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }

  function readSheet(f) {
    var out = $('#bk-out', box), s = seller();
    if (!s) { out.innerHTML = '<p class="warn">' + esc(t('bk.needbiz')) + '</p>'; $('#bk-biz', box).focus(); return; }
    out.innerHTML = '<p>' + esc(t('bk.reading')) + '</p>';
    NS.xlsx.read(f).then(function (rows) {
      var r = toDocs(rows, s);
      if (r.err) { out.innerHTML = '<p class="warn">' + esc(r.err) + '</p>'; return; }
      batch = r.docs; lastFile = f.name;
      showBatch();
    }).catch(function (e) { out.innerHTML = '<p class="warn">' + esc(t(e.message === 'xls' ? 'bk.xls' : 'bk.badfile')) + '</p>'; });
  }
  function showBatch() {
    var out = $('#bk-out', box), ok = batch.filter(function (x) { return !x.bad.length; }), ewb = mode === 'ewb';
    if (!batch.length) { out.innerHTML = '<p class="warn">' + esc(t('bk.none')) + '</p>'; return; }
    out.innerHTML = '<p><b>' + esc(t('bk.read', { n: batch.length, ok: ok.length, bad: batch.length - ok.length })) + '</b></p>' +
      '<div class="tbl"><table class="bk-t"><thead><tr><th>' + esc(t('bk.c.no')) + '</th><th>' + esc(t('l.date')) + '</th><th>' + esc(t('bk.c.to')) + '</th><th class="n">' + esc(t('bk.c.val')) + '</th><th>' + esc(t('bk.c.st')) + '</th></tr></thead><tbody>' +
      batch.map(function (x) {
        var d = x.d;
        return '<tr class="' + (x.bad.length ? 'bad' : 'ok') + '"><td class="mono">' + esc(d.no) + '</td><td>' + esc(d.date ? NS.fmtDate(d.date) : '—') + '</td><td>' + esc(d.buyer.name || '—') + '</td><td class="n">' + NS.money(x.tt.grand) + '</td><td>' +
          (x.bad.length ? '<details><summary>' + esc(t('bk.fixn', { n: x.bad.length })) + '</summary><ul>' + x.bad.map(function (p) { return '<li>' + esc(p) + '</li>'; }).join('') + '</ul></details>' : esc(t('bk.ready'))) +
          (x.note.length ? '<details class="soft"><summary>' + esc(t('bk.notes', { n: x.note.length })) + '</summary><ul>' + x.note.map(function (p) { return '<li>' + esc(p) + '</li>'; }).join('') + '</ul></details>' : '') + '</td></tr>';
      }).join('') + '</tbody></table></div>' +
      (ok.length ? (ewb ? '' : '<p><label class="chk"><input id="bk-save" type="checkbox" checked> ' + esc(t('bk.save')) + '</label></p>') +
        '<p class="row"><button type="button" class="btn primary" id="bk-make">' + esc(t(ewb ? 'bk.makeewb' : 'bk.make', { n: ok.length })) + '</button></p>' +
        (batch.length > ok.length ? '<p class="hint">' + esc(t('bk.leftout')) + '</p>' : '') : '<p class="hint">' + esc(t('bk.fixfirst')) + '</p>');
    var mk = $('#bk-make', box);
    if (mk) mk.addEventListener('click', function () { make(ok); });
  }
  function make(ok) {
    var data, name, ewb = mode === 'ewb';
    if (ewb) {
      data = { version: EWB_VERSION, billLists: ok.map(function (x) { return NS.einv.ewaybill(x.d, x.tt).billLists[0]; }) };
      name = 'e-way-bills-' + NS.today() + '.json';
    } else {
      data = ok.map(function (x) { return NS.einv.einvoice(x.d, x.tt)[0]; });
      name = 'e-invoices-' + NS.today() + '.json';
    }
    download(name, data);
    var msg = t(ewb ? 'bk.madeewb' : 'bk.made', { file: name, n: ok.length });
    var save = $('#bk-save', box);
    (save && save.checked ? saveAll(ok) : Promise.resolve(null)).then(function (n) {
      NS.toast(msg);
      var p = document.createElement('p'); p.className = 'ok'; p.textContent = msg + (n != null ? ' ' + t('bk.savedn', { n: n }) : '');
      $('#bk-out', box).appendChild(p);
    });
  }
  function download(name, data) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }));
    a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  /* the invoices go into My data like ones typed in the invoice maker; one already saved with the same number, year and business is updated */
  function saveAll(ok) {
    return NSStore.list('doc').then(function (all) {
      return Promise.all(ok.map(function (x) {
        var d = x.d, have = all.filter(function (r) { return (r.type || 'invoice') === 'invoice' && key(r.no) === key(d.no) && NS.fy(r.date) === NS.fy(d.date) && key(r.seller && r.seller.name) === key(d.seller.name); })[0];
        var rec = JSON.parse(JSON.stringify(d)); rec.total = x.tt.grand;
        if (have) { rec.id = have.id; rec.created = have.created; if (have.irn) rec.irn = have.irn; rec.status = have.status || rec.status; }
        return NSStore.put('doc', rec).then(function (r) { x.id = r.id; });
      }));
    }).then(function () { NS.folderWrite(false); NS.checkReminder(); return ok.length; }).catch(function (e) { NS.toast(t('nosave', { e: e.message }), 'bad'); return null; });
  }

  /* ---------- the portal's reply: JSON (any shape) or an Excel/CSV result sheet ---------- */
  function readReply(f) {
    var out = $('#bk-rep', box);
    out.innerHTML = '<p>' + esc(t('bk.reading')) + '</p>';
    var p = /\.json$/i.test(f.name) || f.type === 'application/json' ? f.text().then(function (s) { return NS.einv.findAll(JSON.parse(s.replace(/^﻿/, '')), []); })
      : NS.xlsx.read(f).then(function (rows) {
        var at = rows.findIndex(function (r) { return r.some(function (c) { return /^(irn|ewbno|ewaybillno|ewaybillnumber|ewbnumber)$/.test(norm(c)); }); });
        if (at < 0) return [];
        var h = rows[at].map(function (c) { return norm(c).replace('number', 'no').replace('date', 'dt'); });
        return NS.einv.findAll(rows.slice(at + 1).map(function (r) { var o = {}; h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o; }), []);
      });
    p.then(function (found) {
      if (!found.length) { out.innerHTML = '<p class="warn">' + esc(t('ei.noreply')) + '</p>'; return; }
      return NSStore.list('doc').then(function (all) { return attach(found, all); });
    }).catch(function () { out.innerHTML = '<p class="warn">' + esc(t('ei.badfile')) + '</p>'; });
  }
  function attach(found, all) {
    var out = $('#bk-rep', box), rows = [], jobs = [], ewb = mode === 'ewb';
    found.forEach(function (r) {
      var q = r.qr ? NS.einv.qrInfo(r.qr) : null, no = String(r.no || (q && q.DocNo) || '').toUpperCase(), sg = q && q.SellerGstin;
      var rec = no && all.filter(function (d) { return (d.type || 'invoice') === 'invoice' && String(d.no).toUpperCase() === no && (!sg || (d.seller && d.seller.gstin) === sg); })
        .sort(function (a, b) { return (b.updated || 0) - (a.updated || 0); })[0];
      var info = { no: no, irn: r.irn || (q && q.Irn) || '', ack: r.ack || '', ackDt: r.ackDt || (q && q.IrnDt) || '', ewb: r.ewb ? String(r.ewb).trim() : '', ewbDt: r.ewbDt || '', ewbTill: r.ewbTill || '', qr: r.qr || '', id: rec && rec.id };
      if (rec) {
        var irn = rec.irn || {};
        ['irn', 'ack', 'ackDt', 'qr', 'ewb', 'ewbDt', 'ewbTill'].forEach(function (k) { if (info[k]) irn[k] = String(info[k]).trim(); });
        irn.sig = sig(rec, totals(rec).grand);
        rec.irn = irn;
        jobs.push(NSStore.put('doc', rec));
      }
      rows.push(info);
    });
    Promise.all(jobs).then(function () {
      var have = rows.filter(function (r) { return r.id; }).length;
      out.innerHTML = '<p><b>' + esc(t('bk.got', { n: rows.length, k: have })) + '</b></p>' +
        '<div class="tbl"><table class="bk-t"><thead><tr><th>' + esc(t('bk.c.no')) + '</th><th>' + esc(ewb ? t('ei.ewbno') : 'IRN') + '</th><th>' + esc(ewb ? t('ei.ewbtill') : t('ei.ack')) + '</th><th></th></tr></thead><tbody>' +
        rows.map(function (r) {
          return '<tr><td class="mono">' + esc(r.no || '—') + '</td><td class="mono">' + esc(ewb ? r.ewb : (r.irn ? r.irn.slice(0, 16) + '…' : '—')) + '</td><td class="mono">' + esc(ewb ? r.ewbTill || '—' : r.ack || '—') + '</td><td>' +
            (r.id ? '<a href="?open=' + encodeURIComponent(r.id) + '#tool">' + esc(t('bk.open')) + '</a>' : '<span class="soft">' + esc(t('bk.notsaved')) + '</span>') + '</td></tr>';
        }).join('') + '</tbody></table></div>' + (rows.length > have && !ewb ? '<p class="hint">' + esc(t('bk.h.notsaved')) + '</p>' : '');
      if (ewb) rows.forEach(function (r) { if (/^\d{12}$/.test(r.ewb)) addPartB({ ewb: r.ewb }); });
      NS.folderWrite(false);
    }).catch(function (e) { NS.toast(t('nosave', { e: e.message }), 'bad'); });
  }

  /* ---------- Part B: the vehicle file (Vehicle Updation, version 1.0.0621) ---------- */
  function partBHtml() {
    return '<div class="card" id="pb"><h3>4. ' + esc(t('bk.pb')) + '</h3><p class="hint">' + esc(t('bk.h.pb')) + '</p>' +
      '<div class="fg"><div class="f"><label for="pb-place">' + esc(t('bk.pbplace')) + '</label><input id="pb-place" dir="auto"></div>' +
      '<div class="f"><label for="pb-state">' + esc(t('l.state')) + '</label><select id="pb-state"></select></div>' +
      '<div class="f"><label for="pb-reason">' + esc(t('bk.pbreason')) + '</label><select id="pb-reason">' + REASONS.map(function (k) { return opt(String(k), t('bk.r.' + k), '4'); }).join('') + '</select></div>' +
      '<div class="f"><label for="pb-remark">' + esc(t('bk.pbremark')) + '</label><input id="pb-remark" maxlength="25" dir="auto"></div></div>' +
      '<div class="tbl"><table class="bk-t pb-t"><thead><tr><th>' + esc(t('ei.ewbno')) + '</th><th>' + esc(t('ei.mode')) + '</th><th>' + esc(t('ei.veh')) + '</th><th>' + esc(t('ei.tdoc')) + '</th><th>' + esc(t('ei.tdate')) + '</th><th></th></tr></thead><tbody id="pb-rows"></tbody></table></div>' +
      '<p class="row"><button type="button" class="btn" id="pb-add">' + esc(t('bk.pbadd')) + '</button><button type="button" class="btn primary" id="pb-make">' + esc(t('bk.pbmake')) + '</button></p>' +
      '<div class="bk-out" id="pb-out" aria-live="polite"></div></div>';
  }
  function addPartB(v) {
    var tb = $('#pb-rows', box); if (!tb) return;
    if (v.ewb && $$('[data-pb="ewb"]', tb).some(function (i) { return i.value === v.ewb; })) return;
    var first = $$('tr', tb)[0];
    if (first && !$('[data-pb="ewb"]', first).value && !$$('tr', tb)[1]) first.remove();
    var tr = document.createElement('tr'), n = tb.children.length;
    tr.innerHTML = '<td><input data-pb="ewb" inputmode="numeric" maxlength="12" dir="ltr" aria-label="' + esc(t('ei.ewbno')) + ' ' + (n + 1) + '" value="' + esc(v.ewb || '') + '"></td>' +
      '<td><select data-pb="mode" aria-label="' + esc(t('ei.mode')) + '">' + [1, 2, 3, 4].map(function (k) { return opt(String(k), t('ei.m.' + k), '1'); }).join('') + '</select></td>' +
      '<td><input data-pb="veh" maxlength="15" dir="ltr" placeholder="PB10AB1234" aria-label="' + esc(t('ei.veh')) + '"></td>' +
      '<td><input data-pb="doc" maxlength="16" dir="ltr" aria-label="' + esc(t('ei.tdoc')) + '"></td>' +
      '<td><input data-pb="date" type="date" aria-label="' + esc(t('ei.tdate')) + '"></td>' +
      '<td><button type="button" class="link" data-pb-del aria-label="' + esc(t('bk.pbdel')) + '">✕</button></td>';
    tb.appendChild(tr);
  }
  function wirePartB() {
    var st = $('#pb-state', box);
    st.innerHTML = '<option value="">' + esc(t('choose.state')) + '</option>' + NS.STATES.filter(function (s) { return s[0] !== '96'; }).map(function (s) { return '<option value="' + s[0] + '">' + s[1] + ' (' + s[0] + ')</option>'; }).join('');
    var s = seller(); if (s) { st.value = s.state || ''; $('#pb-place', box).value = s.city || ''; }
    $('#bk-biz', box).addEventListener('change', function () { var x = seller(); if (x) { st.value = x.state || ''; $('#pb-place', box).value = x.city || ''; } });
    addPartB({});
    $('#pb-add', box).addEventListener('click', function () { addPartB({}); var i = $$('[data-pb="ewb"]', box); i[i.length - 1].focus(); });
    $('#pb-rows', box).addEventListener('click', function (e) { var b = e.target.closest('[data-pb-del]'); if (b) b.closest('tr').remove(); });
    $('#pb-make', box).addEventListener('click', makePartB);
  }
  function makePartB() {
    var out = $('#pb-out', box), s = seller(), bad = [], list = [];
    if (!s || !NS.gstinCheck(s.gstin).ok || !s.gstin) bad.push(t('ev.ewbseller'));
    var place = $('#pb-place', box).value.trim(), state = $('#pb-state', box).value, reason = +$('#pb-reason', box).value, remark = $('#pb-remark', box).value.trim();
    if (place.length < 3) bad.push(t('bk.v.place'));
    if (!state) bad.push(t('ev.state'));
    if (reason === 3 && !remark) bad.push(t('bk.v.remark'));
    $$('#pb-rows tr', box).forEach(function (tr, i) {
      var g = function (k) { return $('[data-pb="' + k + '"]', tr).value.trim(); };
      var ewb = digits(g('ewb')), m = +g('mode'), veh = g('veh').toUpperCase().replace(/[\s-]/g, ''), doc = g('doc'), dt = g('date');
      if (!ewb && !veh && !doc) return;
      if (!/^\d{12}$/.test(ewb)) bad.push(t('bk.v.ewb', { n: i + 1 }));
      if (m === 1 && !veh) bad.push(t('bk.v.veh', { n: i + 1 }));
      if (veh && !/^[A-Z]{2}\d{1,2}[A-Z]{0,3}\d{4}$|^[A-Z]{3}\d{4}$|^\d{2}BH\d{4}[A-Z]{1,2}$|^TR[A-Z0-9]{6,13}$|^(DF|TM|BP|NP)[A-Z0-9]{5,13}$/.test(veh)) bad.push(t('bk.v.vehf', { n: i + 1 }));
      if ((m === 2 || m === 3) && !(doc && dt)) bad.push(t('bk.v.doc', { n: i + 1 }));
      if (m === 4 && !veh && !(doc && dt)) bad.push(t('bk.v.ship', { n: i + 1 }));
      var dmy = dt ? dt.split('-').reverse().join('/') : '';
      list.push({ ewbno: +ewb, transMode: m, vehicleType: 'R', vehicleNo: veh, docNo: doc, docDate: dmy, fromPlace: place.slice(0, 50), fromState: +state, reason: reason, remark: remark.slice(0, 25) });
    });
    if (!list.length) bad.push(t('bk.v.none'));
    if (bad.length) { out.innerHTML = '<p><b>' + esc(t('ev.fix')) + '</b></p><ul class="warn">' + bad.map(function (b) { return '<li>' + esc(b) + '</li>'; }).join('') + '</ul>'; return; }
    var name = 'e-way-bill-part-b-' + NS.today() + '.json';
    download(name, { version: EWB_VERSION, userGstin: s.gstin, vehicleUpdts: list });
    out.innerHTML = '<p class="ok">' + esc(t('bk.pbmade', { file: name, n: list.length })) + '</p>';
  }

  function init() {
    box = $('[data-bulk]');
    if (!box || !NS.einv || !NS.xlsx) return;
    mode = box.getAttribute('data-bulk') === 'ewb' ? 'ewb' : 'einv';
    cols = mode === 'ewb' ? COMMON.concat(EWB) : COMMON.concat(EINV);
    if (mode === 'ewb') { cols = cols.map(function (c) { return c[0] === 'no' ? ['no', 'Document No', 'INV/26-27/101', c[3] + '|invoice no|invoice number'] : c[0] === 'date' ? ['date', 'Document Date', c[2], c[3] + '|invoice date'] : c; }); }
    render();
    document.addEventListener('ns-lang', function () { render(); if (batch.length) showBatch(); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
