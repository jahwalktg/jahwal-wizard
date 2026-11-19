/* 문서 변신 마법사 (시제품) — 자료 여러 개 → 바구니(사람마다 한 줄) → 내 양식에 붓기
   원칙: 자동으로 분류하되, 애매한 것은 전부 표시해 확인받는다. 못 읽는 자료는 바꿔 올 도구를 알려 준다. */
(function () {
  'use strict';
  var O = Office, $ = O.$, esc = O.esc, txt = XK.txt, keyOf = XK.keyOf;
  var TOOL = 'https://jahwalktg.github.io/jahwal-office/tools/';
  function pad(n) { return String(n).padStart(2, '0'); }

  /* ================= 항목 사전 ================= */
  var FIELDS = [
    { id: '이름', syn: ['성명', '이름', '성함', '참여자명', '참여자', '대상자명', '대상자', '직원명', '이용자명', '회원명'], pat: 'name' },
    { id: '생년월일', syn: ['생년월일', '생일', '출생일', '생년'], pat: 'date' },
    { id: '성별', syn: ['성별'], pat: 'sex' },
    { id: '연락처', syn: ['연락처', '전화번호', '휴대폰', '휴대전화', '핸드폰', '전화', '휴대폰번호', '핸드폰번호', 'hp', 'mobile', '본인연락처', '본인전화'], pat: 'phone' },
    { id: '비상연락처', syn: ['비상연락처', '비상연락망', '긴급연락처', '보호자연락처', '비상전화', '비상연락', '가족연락처', '보호자전화'], pat: 'phone' },
    { id: '관계', syn: ['관계', '보호자관계', '본인과의관계'] },
    { id: '주소', syn: ['주소', '거주지', '자택주소', '현주소', '도로명주소', '집주소', '실거주지'], pat: 'addr' },
    { id: '우편번호', syn: ['우편번호'], pat: 'zip' },
    { id: '소속', syn: ['소속', '부서', '팀', '사업단', '사업단명', '부서명', '소속부서', '근무지'] },
    { id: '직위', syn: ['직위', '직급', '직책'] },
    { id: '이메일', syn: ['이메일', '메일', 'email', 'e-mail', '전자우편', '메일주소'], pat: 'email' },
    { id: '비고', syn: ['비고', '메모', '특이사항', '참고'] }
  ];
  var SEQRE = /^(연번|번호|순번|no|일련번호|순서)$/;
  function fieldById(id) { return FIELDS.filter(function (F) { return F.id === id; })[0]; }
  var PAT = {
    phone: function (s) { var d = s.replace(/\D/g, ''); return /^[\d\-\s.()]+$/.test(s) && (/^0\d{8,10}$/.test(d) || /^1\d{8,9}$/.test(d)); },
    email: function (s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); },
    zip: function (s) { return /^\d{5}$/.test(s); },
    addr: function (s) { return /[가-힣]+(시|군|구)\s?[가-힣0-9]+/.test(s) && /(로|길|동|읍|면|리)\s?\d|\d+번지|\d+호/.test(s); },
    sex: function (s) { return /^(남|여|남자|여자|남성|여성|m|f)$/i.test(s); },
    date: function (s, x) { return XK.isDateCell(x) || XK.dateSerial(s) != null || /^\d{6}$/.test(s); },
    name: function (s) { return /^[가-힣]{2,4}$/.test(s.replace(/\s|\(.*?\)/g, '')); }
  };
  var PATNAME = { phone: '전화번호', email: '이메일', addr: '주소' };
  function fit(cells, pat) { var n = 0, ok = 0; cells.forEach(function (x) { var s = txt(x); if (!s) return; n++; if (PAT[pat](s, x)) ok++; }); return n ? ok / n : null; }
  /* 머리글 + 값 모양 → { sel: 항목 | '__new' | '__skip', conf: sure|ask|new|no, why } */
  function classify(head, cells) {
    var k = keyOf(head);
    if (!k || /^칸\d+$/.test(k)) return { sel: '__skip', conf: 'no', why: '머리글이 없는 칸' };
    if (SEQRE.test(k)) return { sel: '__skip', conf: 'no', why: '번호 칸' };
    var f = FIELDS.filter(function (F) { return F.syn.indexOf(k) >= 0; })[0], part = false;
    if (!f) {
      var best = 0;
      FIELDS.forEach(function (F) { F.syn.forEach(function (s) { if (s.length >= 2 && k.indexOf(s) >= 0 && s.length > best) { best = s.length; f = F; part = true; } }); });
    }
    if (f) {
      var r = f.pat ? fit(cells, f.pat) : null;
      if (r != null && r < 0.5) return { sel: f.id, conf: 'ask', why: '머리글은 「' + f.id + '」 같은데 값 모양이 다릅니다. 맞는지 봐 주세요.' };
      if (part) return { sel: f.id, conf: 'ask', why: '머리글이 정확히 같지 않아 「' + f.id + '」(으)로 짐작했습니다.' };
      return { sel: f.id, conf: 'sure' };
    }
    var g = ['phone', 'email', 'addr'].filter(function (p) { return (fit(cells, p) || 0) >= 0.8; })[0];
    if (g) { var id = { phone: '연락처', email: '이메일', addr: '주소' }[g]; return { sel: id, conf: 'ask', why: '처음 보는 머리글인데 값이 ' + PATNAME[g] + ' 모양이라 「' + id + '」(으)로 짐작했습니다.' }; }
    return { sel: '__new', conf: 'new' };
  }

  /* ================= 값 읽기 ================= */
  function fmtPhone(d) { if (/^02\d{7,8}$/.test(d)) return d.replace(/^(02)(\d{3,4})(\d{4})$/, '$1-$2-$3'); return d.replace(/^(\d{3})(\d{3,4})(\d{4})$/, '$1-$2-$3'); }
  function isPhoneF(f) { return f === '연락처' || f === '비상연락처'; }
  function cellStr(x, field) {
    if (!x) return '';
    if (XK.isDateCell(x)) { var p = XK.serialYMD(x.v); if (p) return p.y + '-' + pad(p.m) + '-' + pad(p.d); }
    var s = txt(x).replace(/\s*\n\s*/g, ' ');
    if (isPhoneF(field)) { var d = s.replace(/\D/g, ''); if (/^\d+$/.test(s) && /^1\d{8,9}$/.test(d)) s = fmtPhone('0' + d); }   // 엑셀이 앞자리 0을 지운 전화번호
    if (field === '생년월일') { var n = XK.dateSerial(s); if (n != null) { var q = XK.serialYMD(n); if (q) s = q.y + '-' + pad(q.m) + '-' + pad(q.d); } }
    return s;
  }
  function normVal(field, v) { return isPhoneF(field) ? v.replace(/\D/g, '') : v.replace(/\s/g, '').toLowerCase(); }
  function normName(s) { return s.replace(/\(.*?\)|（.*?）/g, '').replace(/\s/g, ''); }
  function dataRows(g, h) {
    return g.slice(h + 1).filter(function (r) {
      if (XK.rowEmpty(r)) return false;
      var first = r.filter(function (x) { return txt(x); })[0];
      return !(first && XK.SUMRE.test(txt(first).replace(/\s/g, '')));
    });
  }

  /* ================= 한글(hwpx) 표 읽기 ================= */
  function unx(s) { return s.replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'); }
  function topTables(xml) {
    var re = /<hp:tbl\b|<\/hp:tbl>/g, m, d = 0, s = 0, out = [], nested = false;
    while ((m = re.exec(xml))) {
      if (m[0].charAt(1) !== '/') { if (d === 0) { s = m.index; nested = false; } else nested = true; d++; }
      else { d--; if (d === 0) out.push({ s: s, e: m.index + m[0].length, xml: xml.slice(s, m.index + m[0].length), nested: nested }); }
    }
    return out;
  }
  function tblGrid(tx) {
    var rowsXml = tx.match(/<hp:tr\b[^>]*>[\s\S]*?<\/hp:tr>/g) || [], g = [], merges = [], W = 0;
    rowsXml.forEach(function (rx, ri) {
      (rx.match(/<hp:tc\b[\s\S]*?<\/hp:tc>/g) || []).forEach(function (tc) {
        var c = +(tc.match(/colAddr="(\d+)"/) || [0, 0])[1], r = +(tc.match(/rowAddr="(\d+)"/) || [0, ri])[1], cs = +(tc.match(/colSpan="(\d+)"/) || [0, 1])[1], rs = +(tc.match(/rowSpan="(\d+)"/) || [0, 1])[1];
        var t = (tc.match(/<hp:p\b[\s\S]*?<\/hp:p>/g) || []).map(function (p) { return (p.match(/<hp:t\b[^>]*>[\s\S]*?<\/hp:t>/g) || []).map(unx).join(''); }).join('\n').trim();
        (g[r] = g[r] || [])[c] = t ? { v: t, t: 's', w: t } : null; W = Math.max(W, c + cs);
        if (cs > 1 || rs > 1) merges.push({ r: r, c: c, rs: rs, cs: cs });
      });
    });
    for (var r = 0; r < rowsXml.length; r++) { g[r] = g[r] || []; for (var c = 0; c < W; c++) if (g[r][c] === undefined) g[r][c] = null; }
    return { g: g, merges: merges, rowsXml: rowsXml };
  }
  function parseHwpx(ab) {
    return JSZip.loadAsync(ab).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return /^Contents\/section\d+\.xml$/.test(n); }).sort();
      return Promise.all(names.map(function (n) { return zip.file(n).async('string'); })).then(function (xmls) {
        var tabs = [], xmlOf = {};
        xmls.forEach(function (xml, si) {
          xmlOf[names[si]] = xml;
          topTables(xml).forEach(function (t) { var gr = tblGrid(t.xml); tabs.push({ name: '표 ' + (tabs.length + 1), path: names[si], s: t.s, e: t.e, tx: t.xml, g: gr.g, merges: gr.merges, nested: t.nested, rows: gr.rowsXml }); });
        });
        return { zip: zip, tabs: tabs, xmlOf: xmlOf };
      });
    });
  }

  /* ================= ① 자료 ================= */
  var srcs = [], seq = 0;
  function mkSrc(name, kind, tabs) {
    var s = { id: ++seq, name: name, kind: kind, tabs: tabs, tab: 0, h: null, over: {}, cols: [], bad: null, notes: [] };
    for (var i = 0; i < tabs.length; i++) {   // 이름 칸이 보이는 첫 시트·표를 고름
      var h = XK.findHead(tabs[i].g), hn = XK.headNames(tabs[i].g, h);
      if (hn.some(function (n) { return classify(n, []).sel === '이름'; })) { s.tab = i; break; }
    }
    return s;
  }
  function badSrc(name, why, fixes, tail) { return { id: ++seq, name: name, bad: { why: why, fixes: fixes || [], tail: tail || '' } }; }
  function analyze(s) {
    if (s.bad) return;
    var t = s.tabs[s.tab], g = t.g; if (s.h == null) s.h = XK.findHead(g);
    var heads = XK.headNames(g, s.h), rows = dataRows(g, s.h), used = {};
    s.rows = rows; s.notes = [];
    s.cols = heads.map(function (hd, i) {
      var cells = rows.map(function (r) { return r[i]; }), c = classify(hd, cells), o = s.over[i];
      var col = { i: i, head: hd, ex: cells.map(function (x) { return txt(x); }).filter(Boolean).slice(0, 2).join(' · '), sel: c.sel, conf: c.conf, why: c.why || '', dup: false };
      if (o != null) { col.sel = o; col.conf = 'done'; col.why = ''; }
      return col;
    });
    s.cols.forEach(function (col) {
      if (col.sel.charAt(0) === '_') return;
      if (used[col.sel] != null) { if (col.conf !== 'done') { col.conf = 'ask'; col.dup = true; col.why = '「' + col.sel + '」 칸이 이 자료에 둘입니다. 이 칸은 무엇인지 골라 주세요(정하기 전에는 가져오지 않습니다).'; } }
      else used[col.sel] = col.i;
    });
    var nc = s.cols.filter(function (c) { return c.sel === '이름' && !c.dup; })[0];
    s.nameCol = nc ? nc.i : -1;
    var mg = (t.merges || []).filter(function (m) { return m.r > s.h; }).length;
    if (mg) s.notes.push({ t: '자료 줄에 합친 칸이 ' + mg + '곳 있어 값이 빠질 수 있습니다. ', a: '표 청소기', u: TOOL + 'xlclean.html', t2: '로 풀어서 다시 넣으면 안전합니다.' });
    if (!rows.length) s.notes.push({ t: '머리글 줄 아래에 자료가 없는 빈 표라 바구니에 넣지 않았습니다. 채울 양식이라면 아래 ③ 「내 양식 넣기」에 넣어 주세요. 자료가 맞다면 머리글 줄 번호를 고쳐 주세요.' });
  }
  var EXT_HELP = '원본 엑셀·한글 파일을 넣거나, 그 표를 복사해 「표를 복사해 붙여 넣기」로 넣어 주세요.';
  function addFiles(files) {
    var jobs = files.filter(function (f) { return !/^~\$/.test(f.name); }).map(function (f) {
      if (srcs.some(function (s) { return s.name === f.name && s.size === f.size; })) return null;
      var ext = (f.name.split('.').pop() || '').toLowerCase(), done = function (s) { if (s) s.size = f.size; return s; };
      if (/^(xlsx|xlsm|xls|csv)$/.test(ext)) return XK.readBook(f).then(function (wb) {
        var tabs = wb.SheetNames.map(function (n) { var sg = XK.sheetGrid(wb.Sheets[n]); return { name: n, g: sg.g, merges: sg.merges }; }).filter(function (t) { return t.g.length; });
        return done(tabs.length ? mkSrc(f.name, '엑셀', tabs) : badSrc(f.name, '내용이 없는 파일입니다.'));
      }, function () { return done(badSrc(f.name, '파일을 열지 못했습니다. 암호가 걸려 있으면 엑셀에서 암호를 푼 뒤 다시 넣어 주세요.')); });
      if (ext === 'hwpx') return O.readBuf(f).then(parseHwpx).then(function (h) {
        var tabs = h.tabs.filter(function (t) { return !t.nested && t.g.length >= 2; });
        return done(tabs.length ? mkSrc(f.name, '한글', tabs) : badSrc(f.name, '이 한글 문서에서 읽을 수 있는 표를 찾지 못했습니다(표 안에 표가 든 문서는 아직 못 읽습니다). ' + EXT_HELP));
      }, function () { return done(badSrc(f.name, '한글 문서를 열지 못했습니다. 암호가 걸렸거나 깨진 파일일 수 있습니다.')); });
      if (ext === 'hwp') return Promise.resolve(done(badSrc(f.name, '옛 한글 형식(hwp)은 바로 읽지 못합니다. ', [{ a: 'HWP → HWPX 바꾸기', u: TOOL + 'hwp2hwpx.html' }], ' 도구로 hwpx로 바꾼 뒤 그 파일을 다시 넣어 주세요.')));
      if (/^(pdf|jpg|jpeg|png|gif|bmp|heic)$/.test(ext)) return Promise.resolve(done(badSrc(f.name, 'PDF·사진 속 표는 읽지 못합니다. ' + EXT_HELP)));
      if (/^(docx|doc|pptx|ppt)$/.test(ext)) return Promise.resolve(done(badSrc(f.name, '워드·PPT 파일은 아직 읽지 못합니다. ' + EXT_HELP)));
      return Promise.resolve(done(badSrc(f.name, '읽을 수 없는 형식입니다. ' + EXT_HELP)));
    }).filter(Boolean);
    return Promise.all(jobs).then(function (list) { list.forEach(function (s) { if (s) srcs.push(s); }); refresh(); });
  }
  var pasteN = 0;
  function addPasted(m) { if (!m || !m.g.length) { O.toast('표를 읽지 못했습니다. 표를 다시 복사해 붙여 주세요.'); return; } srcs.push(mkSrc('붙여 넣은 표 ' + (++pasteN), '붙여 넣기', [{ name: '표', g: m.g, merges: m.merges }])); refresh(); }

  function fieldOpts(sel, head) {
    var h = '';
    FIELDS.forEach(function (F) { h += '<option value="' + F.id + '"' + (sel === F.id ? ' selected' : '') + '>' + F.id + '</option>'; });
    return h + '<option value="__new"' + (sel === '__new' ? ' selected' : '') + '>새 항목 「' + esc(head) + '」 그대로</option><option value="__skip"' + (sel === '__skip' ? ' selected' : '') + '>가져오지 않음</option>';
  }
  function chip(c) {
    if (c.conf === 'sure') return '<span class="st sure">✓ 알아봄</span>';
    if (c.conf === 'done') return '<span class="st done">✓ 확인함</span>';
    if (c.conf === 'ask') return '<span class="st ask">? 확인 필요</span>';
    if (c.conf === 'new') return '<span class="st new">＋ 새 항목</span>';
    return '<span class="st no">가져오지 않음</span>';
  }
  function renderSrcs() {
    var h = '';
    srcs.forEach(function (s) {
      if (s.bad) {
        h += '<div class="card bad" id="src' + s.id + '"><h3>⚠ ' + esc(s.name) + ' <small>읽지 못함</small><button type="button" class="btn mini" data-del="' + s.id + '">빼기</button></h3><p class="fix">' + esc(s.bad.why) + s.bad.fixes.map(function (x) { return '<a href="' + x.u + '" target="_blank" rel="noopener">' + esc(x.a) + ' ↗</a>'; }).join(' ') + esc(s.bad.tail) + '</p></div>';
        return;
      }
      var t = s.tabs[s.tab], ask = s.cols.filter(function (c) { return c.conf === 'ask'; }).length;
      h += '<div class="card' + (s.nameCol < 0 || ask ? ' warn' : '') + '" id="src' + s.id + '"><h3>' + (s.kind === '한글' ? '📄' : '📗') + ' ' + esc(s.name) + ' <small>' + s.kind + ' · 자료 ' + s.rows.length + '줄</small>' +
        (s.tabs.length > 1 ? '<label class="hint">' + (s.kind === '한글' ? '표' : '시트') + ' <select data-tab="' + s.id + '">' + s.tabs.map(function (x, i) { return '<option value="' + i + '"' + (i === s.tab ? ' selected' : '') + '>' + esc(x.name) + '</option>'; }).join('') + '</select></label>' : '') +
        '<label class="hint">머리글 줄 <input type="number" min="1" max="60" value="' + (s.h + 1) + '" data-head="' + s.id + '" style="width:64px;padding:3px 6px"></label>' +
        '<button type="button" class="btn mini" data-del="' + s.id + '">빼기</button></h3>';
      if (s.nameCol < 0) h += '<p class="fix" style="color:var(--red)"><b>이름 칸을 찾지 못했습니다.</b> 아래에서 이름이 든 칸을 「이름」으로 골라 주세요. 그 전에는 이 자료를 바구니에 넣지 않습니다.</p>';
      s.notes.forEach(function (n) { h += '<p class="fix">⚠ ' + esc(n.t) + (n.a ? '<a href="' + n.u + '" target="_blank" rel="noopener">' + esc(n.a) + ' ↗</a>' + esc(n.t2 || '') : '') + '</p>'; });
      h += '<div class="scroll"><table class="list"><tr><th>자료의 머리글</th><th>들어 있는 값 (예)</th><th>바구니 항목</th><th>상태</th></tr>';
      s.cols.forEach(function (c) {
        h += '<tr class="' + (c.conf === 'ask' ? 'ask' : c.sel === '__skip' ? 'skip' : '') + '"><td><b>' + esc(c.head) + '</b></td><td class="ex" title="' + esc(c.ex) + '">' + esc(c.ex || '(빈 칸)') + '</td>' +
          '<td><select data-col="' + s.id + '|' + c.i + '" aria-label="' + esc(c.head) + ' 칸의 항목">' + fieldOpts(c.sel, c.head) + '</select></td>' +
          '<td>' + chip(c) + (c.conf === 'ask' ? ' <button type="button" class="btn mini" data-ok="' + s.id + '|' + c.i + '">' + (c.dup ? '이대로 둘 다 쓰기' : '맞아요') + '</button><div class="why">' + esc(c.why) + '</div>' : '') + '</td></tr>';
      });
      h += '</table></div></div>';
    });
    $('srcList').innerHTML = h;
  }
  function srcById(id) { return srcs.filter(function (s) { return s.id === +id; })[0]; }
  $('srcList').addEventListener('change', function (e) {
    var el = e.target, d = el.dataset;
    if (d.col) { var a = d.col.split('|'), s = srcById(a[0]); s.over[+a[1]] = el.value; refresh(); }
    else if (d.tab) { var s2 = srcById(d.tab); s2.tab = +el.value; s2.h = null; s2.over = {}; refresh(); }
    else if (d.head) { var s3 = srcById(d.head); s3.h = Math.max(0, Math.min(s3.tabs[s3.tab].g.length - 1, (+el.value || 1) - 1)); s3.over = {}; refresh(); }
  });
  $('srcList').addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    if (b.dataset.del) { srcs = srcs.filter(function (s) { return s.id !== +b.dataset.del; }); refresh(); }
    else if (b.dataset.ok) { var a = b.dataset.ok.split('|'), s = srcById(a[0]), c = s.cols[+a[1]]; s.over[c.i] = c.sel; refresh(); }
  });

  /* ================= ② 바구니 ================= */
  var people = [], fieldsOrder = [], HOMO = {}, USE = [], ack = {}, pick = {}, edits = {}, off = {};
  function build() {
    people = []; HOMO = {};
    var byName = {}, fset = {}, custom = {};
    USE = srcs.filter(function (s) { return !s.bad && s.nameCol >= 0 && s.rows.length > 0; });   // 빈 표(양식)는 자료로 세지 않음
    function cname(head) { var k = keyOf(head); return custom[k] || (custom[k] = head); }
    USE.forEach(function (s) {
      var bc = s.cols.filter(function (c) { return c.sel === '생년월일' && !c.dup; })[0], bcol = bc ? bc.i : -1;
      var cols = s.cols.filter(function (c) { return c.sel !== '__skip' && !(c.dup && c.conf === 'ask'); }).map(function (c) { return { i: c.i, f: c.sel === '__new' ? cname(c.head) : c.sel }; });
      cols.forEach(function (c) { fset[c.f] = 1; });
      s.rows.forEach(function (r) {
        var disp = cellStr(r[s.nameCol], '이름').trim(), name = normName(disp); if (!name) return;
        var birth = bcol >= 0 ? cellStr(r[bcol], '생년월일') : '';
        var cands = byName[name] || (byName[name] = []), free = cands.filter(function (p) { return !p.seen[s.id]; }), m = null;
        if (birth) m = free.filter(function (p) { return p.birth === birth; })[0] || free.filter(function (p) { return !p.birth; })[0] || null;
        else m = free[0] || null;
        if (!m) { m = { name: disp, key: name, birth: birth, vals: {}, seen: {}, n: cands.length }; cands.push(m); people.push(m); }
        if (cands.length > 1) HOMO[name] = cands;
        m.seen[s.id] = 1; if (birth && !m.birth) m.birth = birth;
        cols.forEach(function (c) {
          var v = cellStr(r[c.i], c.f); if (c.f === '이름') v = v.trim(); if (!v) return;
          var a = m.vals[c.f] || (m.vals[c.f] = []), nv = normVal(c.f, v);
          if (!a.some(function (e) { return normVal(c.f, e.v) === nv; })) a.push({ v: v, src: s.id });
        });
      });
    });
    people.forEach(function (p) { p.pid = p.key + '|' + p.n; });
    fieldsOrder = Object.keys(fset).length ? ['이름'].concat(FIELDS.map(function (F) { return F.id; }).filter(function (id) { return id !== '이름' && fset[id]; })).concat(Object.keys(fset).filter(function (f) { return !fieldById(f); })) : [];
  }
  function valueOf(p, f) { var k = p.pid + '|' + f; if (edits[k] != null) return edits[k]; var a = p.vals[f]; if (!a) return ''; return (a[pick[k] || 0] || a[0]).v; }
  function included() { return people.filter(function (p) { return !off[p.pid]; }); }
  function srcNo(id) { for (var i = 0; i < USE.length; i++) if (USE[i].id === id) return i + 1; return '?'; }
  function srcNm(id) { var s = srcById(id); return s ? s.name.replace(/\.[^.]+$/, '') : ''; }
  function issues() {
    var L = [];
    srcs.forEach(function (s) {
      if (s.bad) return;
      if (s.nameCol < 0) L.push({ id: 'nokey|' + s.id, block: true, go: 'src' + s.id, t: '「' + s.name + '」에서 이름 칸을 찾지 못해 바구니에 넣지 않았습니다. 자료 카드에서 이름 칸을 골라 주세요.' });
      s.cols.forEach(function (c) { if (c.conf === 'ask') L.push({ id: 'col|' + s.id + '|' + c.i, go: 'src' + s.id, col: s.id + '|' + c.i, dup: c.dup, t: '「' + s.name + '」의 「' + c.head + '」 칸 → ' + c.why }); });
    });
    included().forEach(function (p) {
      fieldsOrder.forEach(function (f) {
        var a = p.vals[f], k = p.pid + '|' + f;
        if (a && a.length > 1 && pick[k] == null && edits[k] == null) L.push({ id: 'cf|' + k, need: true, t: p.name + ' 님의 「' + f + '」이(가) 자료마다 다릅니다 — ' + a.map(function (e) { return e.v + ' (' + srcNm(e.src) + ')'; }).join(' / ') + '. 아래 표의 노란 칸에서 골라 주세요.' });
      });
    });
    Object.keys(HOMO).forEach(function (n) { var c = HOMO[n]; L.push({ id: 'homo|' + n + '|' + c.length, ackable: true, t: '같은 이름 「' + c[0].name + '」이(가) ' + c.length + '명입니다. 서로 다른 사람이 맞는지, 각 자료의 정보가 맞는 사람에게 붙었는지 봐 주세요.' }); });
    if (USE.length > 1) USE.forEach(function (s) {
      var miss = included().filter(function (p) { return !p.seen[s.id]; });
      if (miss.length) L.push({ id: 'miss|' + s.id + '|' + miss.map(function (p) { return p.pid; }).join(','), ackable: true, t: '「' + s.name + '」에 없는 사람 ' + miss.length + '명: ' + miss.slice(0, 8).map(function (p) { return p.name; }).join(', ') + (miss.length > 8 ? ' …' : '') + ' — 이분들은 그 자료의 칸이 비어 있습니다.' });
    });
    L.forEach(function (i) { i.done = !!ack[i.id]; });
    return L;
  }
  function todoCount() { return issues().filter(function (i) { return !i.done; }).length; }
  function renderBasket() {
    var has = people.length > 0 || srcs.some(function (s) { return !s.bad; });
    $('basketPanel').hidden = !has; $('tplPanel').hidden = !people.length;
    if (!has) return;
    var L = issues(), todo = L.filter(function (i) { return !i.done; }).length, inc = included();
    $('basketSum').innerHTML = '<span class="pill">자료 ' + USE.length + '개</span><span class="pill">사람 ' + inc.length + '명' + (inc.length < people.length ? ' (뺀 사람 ' + (people.length - inc.length) + ')' : '') + '</span><span class="pill">항목 ' + fieldsOrder.length + '개</span>' +
      (todo ? '<span class="pill todo">확인할 것 ' + todo + '개</span>' : '<span class="pill ok">✓ 확인할 것 없음</span>');
    $('issueList').innerHTML = L.map(function (i) {
      return '<li class="' + (i.done ? 'done' : i.block || i.need ? 'block' : '') + '"><span class="t">' + (i.done ? '✓ ' : '') + esc(i.t) + '</span>' +
        (i.done ? '' : i.col ? '<button type="button" class="btn mini" data-colok="' + i.col + '">' + (i.dup ? '이대로 둘 다 쓰기' : '맞아요') + '</button><button type="button" class="btn mini" data-go="' + i.go + '">자료에서 고치기</button>' : i.go ? '<button type="button" class="btn mini" data-go="' + i.go + '">자료로 가기</button>' : i.ackable ? '<button type="button" class="btn mini" data-ack="' + esc(i.id) + '">확인했어요</button>' : '') + '</li>';
    }).join('');
    if (!people.length) { $('basketPv').innerHTML = '<p class="hint" style="padding:10px;margin:0">아직 바구니에 담긴 사람이 없습니다. 위 확인할 것을 먼저 풀어 주세요.</p>'; return; }
    var h = '<table><tr><th>넣기</th>' + fieldsOrder.map(function (f) { return '<th>' + esc(f) + '</th>'; }).join('') + '<th>나온 자료</th></tr>';
    people.forEach(function (p) {
      h += '<tr class="' + (off[p.pid] ? 'off ' : '') + (HOMO[p.key] ? 'flag' : '') + '"><td style="text-align:center"><input type="checkbox" data-off="' + esc(p.pid) + '"' + (off[p.pid] ? '' : ' checked') + ' aria-label="' + esc(p.name) + ' 넣기"></td>';
      fieldsOrder.forEach(function (f) {
        var a = p.vals[f], k = p.pid + '|' + f, v = valueOf(p, f);
        if (a && a.length > 1 && edits[k] == null) {
          h += '<td class="' + (pick[k] == null ? 'cf' : 'picked') + '"><select data-pick="' + esc(k) + '" aria-label="' + esc(p.name + ' ' + f) + ' 고르기">' + (pick[k] == null ? '<option value="">— 골라 주세요 —</option>' : '') + a.map(function (e, i) { return '<option value="' + i + '"' + (pick[k] === i ? ' selected' : '') + '>' + esc(e.v) + ' (' + esc(srcNm(e.src)) + ')</option>'; }).join('') + '</select></td>';
        } else h += '<td contenteditable="true" spellcheck="false" data-ed="' + esc(k) + '" class="' + (f === '이름' ? 'who ' : '') + (edits[k] != null ? 'edited' : v ? '' : 'none') + '">' + esc(v) + '</td>';
      });
      h += '<td class="src">' + USE.filter(function (s) { return p.seen[s.id]; }).map(function (s) { return srcNo(s.id); }).join('·') + ' / ' + USE.length + '</td></tr>';
    });
    $('basketPv').innerHTML = h + '</table><p class="hint" style="padding:4px 8px;margin:0">나온 자료 번호: ' + USE.map(function (s, i) { return (i + 1) + ' = ' + esc(s.name); }).join(' · ') + '</p>';
  }
  $('issueList').addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    if (b.dataset.ack) { ack[b.dataset.ack] = 1; renderBasket(); renderTpls(); }
    else if (b.dataset.colok) { var a = b.dataset.colok.split('|'), s = srcById(a[0]), c = s.cols[+a[1]]; s.over[c.i] = c.sel; refresh(); }
    else if (b.dataset.go) { var el = $(b.dataset.go); if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); el.style.outline = '3px solid var(--blue)'; setTimeout(function () { el.style.outline = ''; }, 1600); } }
  });
  $('basketPv').addEventListener('change', function (e) {
    var d = e.target.dataset;
    if (d.pick != null) { if (e.target.value === '') delete pick[d.pick]; else pick[d.pick] = +e.target.value; renderBasket(); renderTpls(); }
    else if (d.off != null) { if (e.target.checked) delete off[d.off]; else off[d.off] = 1; renderBasket(); tpls.forEach(mapTpl); renderTpls(); }
  });
  $('basketPv').addEventListener('focusout', function (e) {
    var td = e.target; if (!td.dataset || td.dataset.ed == null) return;
    var k = td.dataset.ed, now = td.textContent.replace(/\s+/g, ' ').trim(), i = k.lastIndexOf('|'), pid = k.slice(0, i), f = k.slice(i + 1);
    var p = people.filter(function (x) { return x.pid === pid; })[0]; if (!p) return;
    if (now === valueOf(p, f)) return;
    var base = p.vals[f] ? p.vals[f][pick[k] || 0].v : '';
    if (now === base) delete edits[k]; else edits[k] = now;
    renderBasket(); renderTpls();
  });
  $('basketPv').addEventListener('keydown', function (e) { if (e.key === 'Enter' && e.target.dataset && e.target.dataset.ed != null) { e.preventDefault(); e.target.blur(); } });
  $('basketXlsx').onclick = function () {
    var inc = included(); if (!inc.length) { O.toast('바구니가 비어 있습니다.'); return; }
    if (!okToGo()) return;
    var body = inc.map(function (p) { return fieldsOrder.map(function (f) { return { v: valueOf(p, f), s: 'textL' }; }); });
    O.download('바구니_합친 표.xlsx', Xlsxw.book(XK.bookSheet('바구니', fieldsOrder, body))); O.toast(inc.length + '명을 엑셀로 내려받았습니다.');
  };
  function okToGo(extra) {
    var n = todoCount() + (extra || 0);
    return !n || confirm('아직 확인하지 않은 것이 ' + n + '개 있습니다.\n확인하지 않은 채로 받으면 틀린 값이 들어갈 수 있습니다. 그래도 받을까요?');
  }

  /* ================= ③ 양식 ================= */
  var tpls = [], tseq = 0;
  function headScore(g) {   // 양식에서 머리글 줄 찾기: 아는 항목(또는 번호)이 가장 많은 줄
    var best = -1, bs = 0;
    for (var r = 0; r < Math.min(g.length, 40); r++) {
      var sc = 0; (g[r] || []).forEach(function (x) { var t = txt(x); if (!t) return; var k = keyOf(t), c = classify(t, []); if (SEQRE.test(k) || (c.sel.charAt(0) !== '_' && c.conf === 'sure')) sc++; });
      if (sc > bs) { bs = sc; best = r; }
    }
    return { h: best, score: bs };
  }
  function shape(t) {   // 머리글 아래 빈 줄 수(cap), 그 아래 다른 내용(footer)
    var tab = t.tabs[t.tab], g = tab.g, h = t.h, heads = XK.headNames(g, h);
    var seqC = heads.map(function (n, i) { return SEQRE.test(keyOf(n)) ? i : -1; }).filter(function (i) { return i >= 0; });
    function blank(r) { return !(g[r] || []).some(function (x, i) { var s = txt(x); return s && !(seqC.indexOf(i) >= 0 && /^\d+\.?$/.test(s)); }); }
    var start = h + 1; while (start < g.length && !blank(start)) start++;
    var cap = 0; while (start + cap < g.length && blank(start + cap)) cap++;
    var footer = false; for (var r = start + cap; r < g.length; r++) if (!blank(r)) { footer = true; break; }
    t.heads = heads; t.start = start; t.cap = cap; t.footer = footer; t.pre = start - (h + 1);
  }
  function autoMap(hd) {
    var k = keyOf(hd), c = classify(hd, []);
    if (SEQRE.test(k)) return { sel: '__seq', conf: 'sure' };
    if (c.sel === '__skip') return { sel: '__none', conf: 'no' };
    var f = c.sel === '__new' ? fieldsOrder.filter(function (x) { return keyOf(x) === k; })[0] : c.sel;
    if (f && fieldsOrder.indexOf(f) >= 0) return { sel: f, conf: c.conf === 'ask' ? 'ask' : 'sure', why: c.why || '' };
    if (f === '비고') return { sel: '__none', conf: 'new', why: '' };   // 비고 칸은 비워 두는 것이 보통
    if (f) return { sel: '__none', conf: 'ask', why: '바구니에 「' + f + '」 항목이 없어 비워 둡니다. 다른 항목을 넣으려면 골라 주세요.' };
    return { sel: '__none', conf: 'new', why: '' };
  }
  function mapTpl(t) {
    if (t.bad) return;
    shape(t);
    t.map = t.heads.map(function (hd, j) {
      var m = autoMap(hd), o = t.over[j];
      if (o != null) return { head: hd, sel: o, conf: 'done', why: '' };
      return { head: hd, sel: m.sel, conf: m.conf, why: m.why || '' };
    });
    var need = included().length; t.notes = []; t.limit = need;
    if (t.pre) t.notes.push({ warn: true, t: '양식에 이미 내용이 ' + t.pre + '줄 적혀 있습니다. 그 아래 빈 줄부터 채웁니다(빈 양식이 맞는지 확인해 주세요).' });
    if (!t.map.some(function (m) { return m.sel === '이름'; })) t.notes.push({ warn: true, t: '이 양식에는 이름이 들어갈 칸이 정해지지 않았습니다. 아래에서 이름 칸을 골라 주세요.' });
    if (need > t.cap) {
      if (t.footer && t.kind === 'xlsx') { t.limit = t.cap; t.notes.push({ warn: true, t: '빈 줄은 ' + t.cap + '개인데 사람은 ' + need + '명입니다. 표 아래에 다른 내용이 있어 줄을 늘리지 못하니 ' + t.cap + '명만 채웁니다. 양식의 빈 줄을 늘려 다시 넣어 주세요.' }); }
      else if (t.cap === 0 && t.kind === 'hwpx') { t.limit = 0; t.notes.push({ warn: true, t: '머리글 아래에 빈 줄이 하나도 없어 본뜰 줄이 없습니다. 양식에 빈 줄을 하나 이상 남겨 다시 넣어 주세요.' }); }
      else t.notes.push({ t: '빈 줄은 ' + t.cap + '개, 사람은 ' + need + '명 → 모자란 ' + (need - t.cap) + '줄은 마지막 빈 줄과 같은 모양으로 늘립니다.' });
    }
  }
  function tplRows(t) {
    return included().slice(0, t.limit).map(function (p, i) { return t.map.map(function (m) { return m.sel === '__seq' ? i + 1 : m.sel === '__none' ? undefined : valueOf(p, m.sel); }); });
  }
  function badTpl(name, why, fixes, tail) { return { id: ++tseq, name: name, bad: { why: why, fixes: fixes || [], tail: tail || '' } }; }
  var NOLIST = '머리글 줄 아래에 빈 줄이 있는 목록 모양 표를 찾지 못했습니다. (사람마다 한 장씩 뽑는 양식은 다음 단계에서 만듭니다)';
  function pickTab(t) {
    var best = -1, bs = 1, bh = 0;
    t.tabs.forEach(function (tab, i) { if (tab.nested) return; var s = headScore(tab.g); if (s.score > bs) { bs = s.score; best = i; bh = s.h; } });
    if (best < 0) return false;
    t.tab = best; t.h = bh; return true;
  }
  function addTpls(files) {
    var jobs = files.filter(function (f) { return !/^~\$/.test(f.name); }).map(function (f) {
      var ext = (f.name.split('.').pop() || '').toLowerCase();
      if (ext === 'xlsx' || ext === 'xlsm') return O.readBuf(f).then(function (ab) {
        var wb = XLSX.read(new Uint8Array(ab), { type: 'array', cellNF: true, sheetStubs: true });
        var tabs = wb.SheetNames.map(function (n) { var ws = wb.Sheets[n], sg = XK.sheetGrid(ws), rg = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 } }; return { name: n, g: sg.g, merges: sg.merges, r0: rg.s.r, c0: rg.s.c }; });
        var t = { id: ++tseq, name: f.name, kind: 'xlsx', ab: ab, tabs: tabs, over: {} };
        return pickTab(t) ? t : badTpl(f.name, NOLIST);
      }, function () { return badTpl(f.name, '양식을 열지 못했습니다. 암호가 걸려 있으면 풀고 다시 넣어 주세요.'); });
      if (ext === 'hwpx') return O.readBuf(f).then(function (ab) { return parseHwpx(ab).then(function (h) {
        var t = { id: ++tseq, name: f.name, kind: 'hwpx', ab: ab, tabs: h.tabs, xmlOf: h.xmlOf, over: {} };
        return pickTab(t) ? t : badTpl(f.name, NOLIST);
      }); }, function () { return badTpl(f.name, '한글 양식을 열지 못했습니다.'); });
      if (ext === 'hwp') return Promise.resolve(badTpl(f.name, '옛 한글 형식(hwp) 양식은 바로 채우지 못합니다. ', [{ a: 'HWP → HWPX 바꾸기', u: TOOL + 'hwp2hwpx.html' }], ' 도구로 hwpx로 바꾼 뒤 다시 넣어 주세요. (한글에서 「다른 이름으로 저장 → hwpx」도 됩니다)'));
      if (ext === 'xls') return Promise.resolve(badTpl(f.name, '옛 엑셀 형식(xls) 양식은 서식을 지킨 채 채우지 못합니다. 엑셀에서 「다른 이름으로 저장 → xlsx」로 바꿔 다시 넣어 주세요.'));
      return Promise.resolve(badTpl(f.name, '양식은 엑셀(xlsx)이나 한글(hwpx) 파일이어야 합니다.'));
    });
    return Promise.all(jobs).then(function (list) { list.forEach(function (t) { tpls.push(t); }); tpls.forEach(mapTpl); renderTpls(); });
  }
  function tplOpts(sel) {
    return '<option value="__none"' + (sel === '__none' ? ' selected' : '') + '>(비워 둠)</option><option value="__seq"' + (sel === '__seq' ? ' selected' : '') + '>(번호 1, 2, 3 … 자동)</option>' +
      fieldsOrder.map(function (f) { return '<option value="' + esc(f) + '"' + (sel === f ? ' selected' : '') + '>' + esc(f) + '</option>'; }).join('');
  }
  function tplAsk(t) { return t.bad ? 0 : t.map.filter(function (m) { return m.conf === 'ask'; }).length; }
  function renderTpls() {
    var h = '';
    tpls.forEach(function (t) {
      if (t.bad) { h += '<div class="card bad"><h3>⚠ ' + esc(t.name) + ' <small>채울 수 없음</small><button type="button" class="btn mini" data-tdel="' + t.id + '">빼기</button></h3><p class="fix">' + esc(t.bad.why) + t.bad.fixes.map(function (x) { return '<a href="' + x.u + '" target="_blank" rel="noopener">' + esc(x.a) + ' ↗</a>'; }).join(' ') + esc(t.bad.tail) + '</p></div>'; return; }
      var rows = tplRows(t), ask = tplAsk(t), tab = t.tabs[t.tab];
      h += '<div class="card' + (ask || t.notes.some(function (n) { return n.warn; }) ? ' warn' : '') + '"><h3>' + (t.kind === 'hwpx' ? '📄' : '📗') + ' ' + esc(t.name) + ' <small>' + (t.kind === 'hwpx' ? '한글' : '엑셀') + ' · ' + esc(tab.name) + ' · 머리글 ' + (t.h + 1) + '번째 줄 · 빈 줄 ' + t.cap + '개</small>' +
        '<label class="hint">머리글 줄 <input type="number" min="1" max="60" value="' + (t.h + 1) + '" data-thead="' + t.id + '" style="width:64px;padding:3px 6px"></label><button type="button" class="btn mini" data-tdel="' + t.id + '">빼기</button></h3>';
      t.notes.forEach(function (n) { h += '<p class="fix"' + (n.warn ? ' style="color:var(--amber)"' : '') + '>' + (n.warn ? '⚠ ' : 'ℹ ') + esc(n.t) + '</p>'; });
      h += '<div class="scroll"><table class="list"><tr><th>양식의 칸</th><th>여기에 넣을 것</th><th>상태</th></tr>';
      t.map.forEach(function (m, j) {
        var st = m.conf === 'sure' ? '<span class="st sure">✓ 짝 맞음</span>' : m.conf === 'done' ? '<span class="st done">✓ 확인함</span>' : m.conf === 'ask' ? '<span class="st ask">? 확인 필요</span> <button type="button" class="btn mini" data-tok="' + t.id + '|' + j + '">맞아요</button><div class="why">' + esc(m.why) + '</div>' : '<span class="st new">비워 둠</span>';
        h += '<tr class="' + (m.conf === 'ask' ? 'ask' : '') + '"><td><b>' + esc(m.head) + '</b></td><td><select data-tmap="' + t.id + '|' + j + '" aria-label="' + esc(m.head) + ' 칸에 넣을 것">' + tplOpts(m.sel) + '</select></td><td>' + st + '</td></tr>';
      });
      h += '</table></div><p class="hint" style="margin:10px 0 4px">채워질 모습 (앞 6줄)</p><div class="pv" style="max-height:230px"><table><tr>' + t.map.map(function (m) { return '<th>' + esc(m.head) + '</th>'; }).join('') + '</tr>' +
        rows.slice(0, 6).map(function (r) { return '<tr>' + r.map(function (v) { return '<td' + (v === undefined || v === '' ? ' class="none"' : '') + '>' + esc(v === undefined ? '' : v) + '</td>'; }).join('') + '</tr>'; }).join('') + '</table></div>' +
        '<div class="row" style="margin-top:10px"><button type="button" class="btn main" data-fill="' + t.id + '"' + (rows.length ? '' : ' disabled') + '>' + rows.length + '명 채워서 받기</button><span class="hint">원래 양식 파일을 그대로 열어 칸만 채웁니다 → 「' + esc(outName(t)) + '」</span></div></div>';
    });
    $('tplList').innerHTML = h;
  }
  function outName(t) { return t.name.replace(/\.[^.]+$/, '') + '_채움.' + t.kind; }
  function tplById(id) { return tpls.filter(function (t) { return t.id === +id; })[0]; }
  $('tplList').addEventListener('change', function (e) {
    var d = e.target.dataset;
    if (d.tmap) { var a = d.tmap.split('|'), t = tplById(a[0]); t.over[+a[1]] = e.target.value; mapTpl(t); renderTpls(); }
    else if (d.thead) { var t2 = tplById(d.thead); t2.h = Math.max(0, Math.min(t2.tabs[t2.tab].g.length - 1, (+e.target.value || 1) - 1)); t2.over = {}; mapTpl(t2); renderTpls(); }
  });
  $('tplList').addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    if (b.dataset.tdel) { tpls = tpls.filter(function (t) { return t.id !== +b.dataset.tdel; }); renderTpls(); }
    else if (b.dataset.tok) { var a = b.dataset.tok.split('|'), t = tplById(a[0]); t.over[+a[1]] = t.map[+a[1]].sel; mapTpl(t); renderTpls(); }
    else if (b.dataset.fill) { var t3 = tplById(b.dataset.fill); if (!okToGo(tplAsk(t3))) return; fill(t3).then(function (blob) { O.download(outName(t3), blob); O.toast(tplRows(t3).length + '명을 채운 양식을 내려받았습니다.'); }, function (err) { O.toast('채우지 못했습니다: ' + (err && err.message || err), 5000); }); }
  });

  /* ---------- 채우기: 한글 ---------- */
  function xe(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function setCell(tc, v) {   // 칸의 첫 문단 모양을 본떠 글자만 바꿈 (기관판 후원금 보고서에서 검증한 방식)
    if (v === undefined) return tc;
    var m = tc.match(/(<hp:subList\b[^>]*>)([\s\S]*?)(<\/hp:subList>)/); if (!m) return tc;
    var p = (m[2].match(/<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/) || [''])[0]; if (!p) return tc;
    var open = p.match(/^<hp:p\b[^>]*>/)[0], cp = (p.match(/charPrIDRef="(\d+)"/) || [0, '0'])[1];
    var paras = String(v).split('\n').map(function (t) { return open + '<hp:run charPrIDRef="' + cp + '">' + (t ? '<hp:t>' + xe(t) + '</hp:t>' : '') + '</hp:run></hp:p>'; }).join('');
    return tc.replace(m[0], m[1] + paras + m[3]);
  }
  function setRow(row, vals) { return row.replace(/<hp:tc\b[\s\S]*?<\/hp:tc>/g, function (tc) { var c = +(tc.match(/colAddr="(\d+)"/) || [0, 0])[1]; return setCell(tc, vals[c]); }); }
  function fillHwpx(t, rowsOut) {
    var tab = t.tabs[t.tab], rows = tab.rows, H = rows.slice(0, t.start), B = rows.slice(t.start, t.start + t.cap), T = rows.slice(t.start + t.cap);
    if (!B.length) return Promise.reject(new Error('양식에 빈 줄이 없습니다.'));
    var body = rowsOut.map(function (vals, i) { return setRow(B[Math.min(i, B.length - 1)], vals); });
    var all = H.concat(body, B.slice(rowsOut.length), T).map(function (r, i) { return r.replace(/rowAddr="\d+"/g, 'rowAddr="' + i + '"'); });
    var pre = tab.tx.slice(0, tab.tx.search(/<hp:tr\b/));
    if (all.length !== rows.length) {
      var hgt = all.reduce(function (s, r) {
        var hs = (r.match(/<hp:tc\b[\s\S]*?<\/hp:tc>/g) || []).filter(function (c) { return /rowSpan="1"/.test(c); }).map(function (c) { return +(c.match(/<hp:cellSz width="\d+" height="(\d+)"/) || [0, 0])[1]; });
        return s + (hs.length ? Math.min.apply(null, hs) : 0);
      }, 0);
      pre = pre.replace(/rowCnt="\d+"/, 'rowCnt="' + all.length + '"');
      if (hgt) pre = pre.replace(/(<hp:sz width="\d+" widthRelTo="\w+" height=")\d+/, '$1' + hgt);
    }
    var xml = t.xmlOf[tab.path], out = xml.slice(0, tab.s) + pre + all.join('') + '</hp:tbl>' + xml.slice(tab.e);
    return JSZip.loadAsync(t.ab).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return !zip.files[n].dir; });
      names.sort(function (a, b) { return (a === 'mimetype' ? 0 : 1) - (b === 'mimetype' ? 0 : 1); });   // mimetype 맨 앞
      return Promise.all(names.map(function (n) { return n === tab.path ? out : zip.file(n).async('uint8array'); })).then(function (datas) {
        var nz = new JSZip();
        names.forEach(function (n, i) { nz.file(n, datas[i], n === 'mimetype' ? { compression: 'STORE' } : { compression: 'DEFLATE' }); });
        return nz.generateAsync({ type: 'blob', mimeType: 'application/hwp+zip' });
      });
    });
  }
  /* ---------- 채우기: 엑셀 (시트 XML의 칸만 고침 → 글꼴·테두리·열 너비 그대로) ---------- */
  function sheetPath(zip, sheetName) {
    return Promise.all([zip.file('xl/workbook.xml').async('string'), zip.file('xl/_rels/workbook.xml.rels').async('string')]).then(function (a) {
      var wb = new DOMParser().parseFromString(a[0], 'application/xml'), rel = new DOMParser().parseFromString(a[1], 'application/xml'), rid = null;
      [].forEach.call(wb.getElementsByTagName('sheet'), function (s) { if (s.getAttribute('name') === sheetName) rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'); });
      var target = null;
      [].forEach.call(rel.getElementsByTagName('Relationship'), function (r) { if (r.getAttribute('Id') === rid) target = r.getAttribute('Target'); });
      if (!target) throw new Error('시트를 찾지 못했습니다.');
      return target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target;
    });
  }
  function fillXlsx(t, rowsOut) {
    return JSZip.loadAsync(t.ab).then(function (zip) {
      var tab = t.tabs[t.tab];
      return sheetPath(zip, tab.name).then(function (path) {
        return zip.file(path).async('string').then(function (xml) {
          var doc = new DOMParser().parseFromString(xml, 'application/xml'), ns = doc.documentElement.namespaceURI, sd = doc.getElementsByTagNameNS(ns, 'sheetData')[0];
          if (!sd) throw new Error('시트 내용을 읽지 못했습니다.');
          var colL = XLSX.utils.encode_col, r0 = tab.r0, c0 = tab.c0;
          function rowEl(R) { var kids = sd.children, i; for (i = 0; i < kids.length; i++) { var n = +kids[i].getAttribute('r'); if (n === R) return kids[i]; if (n > R) break; } var el = doc.createElementNS(ns, 'row'); el.setAttribute('r', R); sd.insertBefore(el, kids[i] || null); return el; }
          function cellEl(row, ci, R) { var ref = colL(ci) + R, kids = row.children, i; for (i = 0; i < kids.length; i++) { var a = kids[i].getAttribute('r'); if (a === ref) return kids[i]; if (XLSX.utils.decode_cell(a).c > ci) break; } var el = doc.createElementNS(ns, 'c'); el.setAttribute('r', ref); row.insertBefore(el, kids[i] || null); return el; }
          var model = t.cap > 0 ? rowEl(r0 + t.start + t.cap) : null, maxR = 0;   // 마지막 빈 줄(1부터 세는 줄 번호)
          var mAttr = {}, mCells = [];
          if (model) { ['ht', 'customHeight', 's', 'customFormat', 'spans'].forEach(function (a) { if (model.hasAttribute(a)) mAttr[a] = model.getAttribute(a); }); [].forEach.call(model.children, function (mc) { mCells.push({ c: XLSX.utils.decode_cell(mc.getAttribute('r')).c, s: mc.getAttribute('s') }); }); }
          rowsOut.forEach(function (vals, i) {
            var R = r0 + t.start + i + 1, row = rowEl(R); maxR = R;
            if (i >= t.cap && model) { Object.keys(mAttr).forEach(function (a) { row.setAttribute(a, mAttr[a]); }); mCells.forEach(function (mc) { var el = cellEl(row, mc.c, R); if (mc.s != null) el.setAttribute('s', mc.s); }); }
            vals.forEach(function (v, j) {
              if (v === undefined) return;
              var el = cellEl(row, c0 + j, R); while (el.firstChild) el.removeChild(el.firstChild);
              if (typeof v === 'number') { el.removeAttribute('t'); var ve = doc.createElementNS(ns, 'v'); ve.textContent = String(v); el.appendChild(ve); }
              else if (v === '') el.removeAttribute('t');
              else { el.setAttribute('t', 'inlineStr'); var is = doc.createElementNS(ns, 'is'), tt = doc.createElementNS(ns, 't'); tt.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); tt.textContent = v; is.appendChild(tt); el.appendChild(is); }
            });
          });
          var dim = doc.getElementsByTagNameNS(ns, 'dimension')[0];
          if (dim && dim.getAttribute('ref') && /:/.test(dim.getAttribute('ref'))) { var rg = XLSX.utils.decode_range(dim.getAttribute('ref')); if (maxR - 1 > rg.e.r) { rg.e.r = maxR - 1; dim.setAttribute('ref', XLSX.utils.encode_range(rg)); } }
          zip.file(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(doc.documentElement));
          return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        });
      });
    });
  }
  function fill(t) { var rows = tplRows(t); return t.kind === 'hwpx' ? fillHwpx(t, rows) : fillXlsx(t, rows); }

  /* ================= 전체 다시 그리기 ================= */
  function refresh() { srcs.forEach(analyze); build(); renderSrcs(); renderBasket(); tpls.forEach(mapTpl); renderTpls(); }

  /* ================= 넣기 연결 ================= */
  O.bindDrop($('dropIn'), $('fileIn'), addFiles);
  O.bindDrop($('dropTpl'), $('fileTpl'), addTpls);
  $('pasteBtn').onclick = function () { $('pasteBox').hidden = !$('pasteBox').hidden; if (!$('pasteBox').hidden) $('pasteArea').focus(); };
  $('pasteArea').addEventListener('paste', function (e) {
    var cd = e.clipboardData; if (!cd) return; e.preventDefault();
    var html = cd.getData('text/html'), m = html && /<table/i.test(html) ? XK.fromHtml(html) : XK.fromText(cd.getData('text/plain') || '');
    addPasted(m); $('pasteArea').value = '';
  });

  /* ================= 예시 (모두 가상 자료) ================= */
  var NAMES = ['김하늘', '이바다', '박구름', '최별', '정나무', '한들꽃', '오솔길', '윤햇살', '서은빛', '장마루'];
  var TEAMS = ['가나팀', '다라팀', '마바팀'], REL = ['배우자', '자녀', '부모', '형제', '자녀', '배우자', '친구', '자녀', '부모', '배우자'];
  var ROADS = ['나래로', '새솔길', '한빛로', '도담길', '온누리로'];
  function ph(i, base) { return '010-' + (base + i) + '-' + String(1000 + i * 37).slice(-4); }
  function sampleSrc() {
    var a = [['2026년 직원 명단(예시)'], [], ['연번', '부서', '성명', '휴대폰']];
    NAMES.forEach(function (n, i) { a.push([i + 1, TEAMS[i % 3], n, i === 2 ? 1020021074 : ph(i, 2000)]); });   // 한 명은 엑셀이 앞자리 0을 지운 번호
    var b = [['이름', '주소', '우편번호']];
    NAMES.forEach(function (n, i) { if (i === 8) return; b.push([n, '서울시 가람구 ' + ROADS[i % 5] + ' ' + (12 + i * 3) + ', ' + (101 + i) + '동 ' + (i + 2) + '02호', String(4500 + i * 11).padStart(5, '0')]); });
    var c = [['비상연락 조사(예시)'], ['성 명', '비상연락처', '관계', '본인 번호', '작성일']];
    NAMES.forEach(function (n, i) { if (i === 3) return; c.push([n, ph(i, 7000), REL[i], i === 4 ? '010-2004-9999' : ph(i, 2000), '2026-11-1' + (i % 5)]); });
    c.push(['강새벽', '010-7010-1370', '부모', '010-2010-1370', '2026-11-12']);
    return [XK.aoaFile('직원 명단(예시).xlsx', a), XK.aoaFile('주소록(예시).xlsx', b), XK.aoaFile('비상연락 조사(예시).xlsx', c)];
  }
  $('sampleIn').onclick = function () { addFiles(sampleSrc()); };
  function sampleTplX() {
    var head = ['연번', '부서', '성명', 'H.P', '주소', '비상연락처', '관계', '확인'].map(function (v) { return { v: v, s: 'head' }; }), rows = [[{ v: '비 상 연 락 망', s: 'title' }], [{ v: '(2026년 11월 기준)', s: 'plain' }], head];
    for (var i = 0; i < 8; i++) rows.push(head.map(function (_, j) { return { v: '', s: j === 4 ? 'textL' : 'text' }; }));
    return new File([Xlsxw.book({ name: '비상연락망', cols: [6, 11, 10, 16, 42, 16, 9, 9], rows: rows, merges: ['A1:H1'], heights: { 1: 30 } })], '비상연락망 양식(예시).xlsx');
  }
  function sampleTplH() {
    var head = ['연번', '성명', '연락처', '주소', '비상연락처', '비고'].map(function (v) { return { t: v, gray: true }; }), rows = [head];
    for (var i = 0; i < 5; i++) rows.push(head.map(function () { return { t: '', align: 'center' }; }));
    return new File([Hwpx.textDoc([{ k: 'p', t: '비 상 연 락 망', align: 'center', style: 'title' }, { k: 'p', t: '(2026년 11월 기준)', align: 'right', style: 'text' }, { k: 'table', widths: [12, 22, 30, 62, 30, 14], rows: rows, h: 8 }, { k: 'p', t: '' }, { k: 'p', t: '○○기관', align: 'center', style: 'sub' }], '비상연락망')], '비상연락망 양식(예시).hwpx');
  }
  $('sampleTplX').onclick = function () { addTpls([sampleTplX()]); };
  $('sampleTplH').onclick = function () { addTpls([sampleTplH()]); };

  /* 시험 입구 */
  window._mm = { addFiles: addFiles, addTpls: addTpls, sampleSrc: sampleSrc, sampleTplX: sampleTplX, sampleTplH: sampleTplH, fill: fill, tplRows: tplRows, issues: issues,
    state: function () { return { srcs: srcs, people: people, fields: fieldsOrder, tpls: tpls, pick: pick, ack: ack, edits: edits, off: off }; }, refresh: refresh, valueOf: valueOf };
})();
