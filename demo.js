/* 문서 변신 마법사 — 화면 안에서 돌아가는 시연(영상 대신)
   실제 도구 위에서 커서가 움직이며 예시 재료·그릇을 넣고, 자막으로 설명한다. 내려받기는 누르지 않는다.
   멈추기: 「그만 보기」 또는 Esc. 화면이 바뀌어도 다시 찍을 필요가 없다(실제 화면을 그대로 쓰므로). */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
  var running = false, stopFlag = false, cur, cap, ring;
  var css = '.dm-cur{position:absolute;left:0;top:0;width:28px;height:28px;z-index:90;pointer-events:none;transition:transform .8s cubic-bezier(.4,0,.2,1);filter:drop-shadow(0 2px 3px rgba(0,0,0,.35))}' +
    '.dm-cur.fast{transition:none}.dm-click{position:absolute;width:36px;height:36px;margin:-18px 0 0 -18px;border:3px solid #FFB800;border-radius:50%;z-index:89;pointer-events:none;animation:dmc .55s ease-out forwards}' +
    '@keyframes dmc{from{transform:scale(.3);opacity:1}to{transform:scale(1.6);opacity:0}}' +
    '.dm-ring{position:absolute;z-index:88;pointer-events:none;border:3px solid #FFB800;border-radius:12px;box-shadow:0 0 0 9999px rgba(10,20,45,.28);transition:all .45s ease}' +
    '.dm-cap{position:fixed;left:50%;bottom:18px;transform:translateX(-50%);z-index:95;width:min(720px,calc(100vw - 24px));background:#1A2B50;color:#fff;border-radius:14px;padding:12px 16px 10px;box-shadow:0 8px 28px rgba(0,0,0,.35);font-size:1rem;line-height:1.5}' +
    '.dm-cap b{color:#FFD166}.dm-cap .dm-row{display:flex;gap:8px;align-items:center;justify-content:space-between;margin-top:8px}' +
    '.dm-cap .dm-dots{display:flex;gap:5px}.dm-cap .dm-dots i{width:8px;height:8px;border-radius:50%;background:#56688f}.dm-cap .dm-dots i.on{background:#FFD166}' +
    '.dm-cap button{background:#fff;color:#1A2B50;border:0;border-radius:8px;padding:6px 12px;font:inherit;font-size:.86rem;font-weight:700;cursor:pointer}' +
    '.dm-cap button.ghost{background:transparent;color:#cfd8ef;border:1px solid #56688f}' +
    '@media (prefers-reduced-motion: reduce){.dm-cur{transition:none}.dm-ring{transition:none}}';
  function ensure() {
    if (cur) return;
    var st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
    cur = document.createElement('div'); cur.className = 'dm-cur'; cur.setAttribute('aria-hidden', 'true');
    cur.innerHTML = '<svg viewBox="0 0 24 24" width="28" height="28"><path d="M4 2l15 9-6.5 1.6L9.6 19z" fill="#fff" stroke="#1A2B50" stroke-width="1.6" stroke-linejoin="round"/></svg>';
    ring = document.createElement('div'); ring.className = 'dm-ring'; ring.hidden = true;
    cap = document.createElement('div'); cap.className = 'dm-cap'; cap.setAttribute('role', 'status'); cap.setAttribute('aria-live', 'polite');
    document.body.appendChild(ring); document.body.appendChild(cur); document.body.appendChild(cap);
  }
  function wait(ms) { return new Promise(function (ok, no) { setTimeout(function () { stopFlag ? no('stop') : ok(); }, reduce ? Math.min(ms, 400) : ms); }); }
  var STEPS = 7, stepNo = 0;
  function say(html, n) {
    if (n) stepNo = n;
    var dots = ''; for (var i = 1; i <= STEPS; i++) dots += '<i class="' + (i <= stepNo ? 'on' : '') + '"></i>';
    cap.innerHTML = '<div>' + html + '</div><div class="dm-row"><span class="dm-dots">' + dots + '</span><button type="button" class="ghost" id="dmStop">그만 보기 (Esc)</button></div>';
    $('#dmStop').onclick = stop;
  }
  function box(el) { var r = el.getBoundingClientRect(); return { x: r.left + scrollX, y: r.top + scrollY, w: r.width, h: r.height }; }
  function show(el) { if (!el) return wait(0); el.scrollIntoView({ behavior: reduce ? 'auto' : 'smooth', block: 'center' }); return wait(750); }
  function moveTo(el) { if (!el) return wait(0); var b = box(el); cur.style.transform = 'translate(' + Math.round(b.x + Math.min(b.w * 0.5, 60)) + 'px,' + Math.round(b.y + b.h * 0.55) + 'px)'; return wait(850); }
  function tap(el) {
    if (!el) return wait(0);
    var b = box(el), c = document.createElement('div'); c.className = 'dm-click'; c.style.left = (b.x + Math.min(b.w * 0.5, 60)) + 'px'; c.style.top = (b.y + b.h * 0.55) + 'px';
    document.body.appendChild(c); setTimeout(function () { c.remove(); }, 600);
    return wait(250);
  }
  function hl(el, pad) { if (!el) { ring.hidden = true; return; } var b = box(el), p = pad == null ? 6 : pad; ring.hidden = false; ring.style.left = (b.x - p) + 'px'; ring.style.top = (b.y - p) + 'px'; ring.style.width = (b.w + p * 2) + 'px'; ring.style.height = (b.h + p * 2) + 'px'; }
  function noHl() { ring.hidden = true; }
  function press(sel) { var el = typeof sel === 'string' ? $(sel) : sel; return show(el).then(function () { return moveTo(el); }).then(function () { return tap(el); }).then(function () { if (el) el.click(); return wait(900); }); }

  function run() {
    if (running) return; running = true; stopFlag = false; ensure();
    document.addEventListener('keydown', onKey);
    var start = $('#demoBtn'); var b0 = start ? box(start) : { x: 40, y: 40, w: 0, h: 0 };
    cur.classList.add('fast'); cur.style.transform = 'translate(' + Math.round(b0.x + 30) + 'px,' + Math.round(b0.y + 20) + 'px)'; cur.hidden = false;
    setTimeout(function () { cur.classList.remove('fast'); }, 50);
    Promise.resolve()
      .then(function () { say('<b>문서 변신 마법사</b> — 흩어진 재료를 모아, 내가 쓰는 양식(그릇)에 맞춰 담아 줍니다. 40초만 같이 보세요.', 1); return wait(2600); })
      .then(function () { say('<b>① 재료 넣기</b> — 엑셀·한글 파일을 한꺼번에 넣습니다. 지금은 예시로 <b>사람 재료 3개</b>(명단·주소록·비상연락 조사)와 <b>물품 재료 3개</b>를 넣어 볼게요.', 2); return press('#sampleIn'); })
      .then(function () { return press('#sampleIn2'); })
      .then(function () {
        var row = $('#srcList tr.ask'); var card = row && row.closest('.card');
        say('도구가 칸마다 무엇인지 알아봅니다. 확실한 칸은 <b>✓ 알아봄</b>, 애매한 칸은 <b>노랗게</b> 표시해 물어봅니다 — 「맞아요」 한 번이면 끝.', 3);
        return show(row || card).then(function () { hl(row || card); return wait(2200); }).then(function () { var ok = row && row.querySelector('[data-ok]'); return ok ? moveTo(ok).then(function () { return tap(ok); }).then(function () { ok.click(); return wait(900); }) : null; });
      })
      .then(function () {
        noHl(); var list = $('#basketList');
        say('<b>② 재료를 구분해 바구니에</b> — 종류가 다른 재료는 바구니를 나누고, 같은 사람·같은 물품은 <b>이름·관리번호</b>로 한 줄에 모읍니다. 칸 모양이 같은 재료는 아래로 이어 붙입니다.', 4);
        return show($('#basketPanel .step')).then(function () { var cards = list.querySelectorAll(':scope > .card'); hl(cards[0] && cards[0].querySelector('h3'), 4); return wait(1500).then(function () { if (cards[1]) { show(cards[1]); } return wait(900); }).then(function () { if (cards[1]) hl(cards[1].querySelector('h3'), 4); return wait(1700); }); });
      })
      .then(function () {
        noHl(); var sel = $('select[data-pick]');
        say('재료마다 값이 다르면(예: 연락처가 두 개) 도구가 정하지 않고 <b>노란 칸에서 골라 달라고</b> 합니다. 칸을 눌러 바로 고칠 수도 있어요.', 5);
        if (!sel) return wait(1500);
        return show(sel).then(function () { hl(sel.closest('td'), 3); return moveTo(sel); }).then(function () { return tap(sel); }).then(function () { sel.value = '0'; sel.dispatchEvent(new Event('change', { bubbles: true })); return wait(1300); });
      })
      .then(function () {
        noHl();
        say('<b>③ 그릇(내 양식) 넣기</b> — 우리 기관이 쓰는 엑셀·한글 양식을 그대로 넣습니다. 목록 모양(비상연락망)도, <b>한 사람에 한 장</b> 모양(개인별 카드)도 알아서 알아봅니다.', 6);
        return press('#sampleTplX').then(function () { return press('#sampleTplF'); }).then(function () { var c = $('#tplList .card'); return show(c && c.querySelector('table.list')).then(function () { hl(c && c.querySelector('table.list'), 4); return wait(2200); }); });
      })
      .then(function () {
        noHl(); var btn = $('#tplList [data-fill]') || $('#tplList [data-fzip]');
        say('<b>④ 담아서 받기</b> — 원래 양식 파일을 그대로 열어 칸만 채웁니다. 글꼴·테두리·제목은 그대로, 빈 줄이 모자라면 같은 모양으로 늘립니다.', 7);
        return show(btn).then(function () { hl(btn, 5); return moveTo(btn); }).then(function () { return wait(1600); }).then(function () {
          var b2 = $('#tplList [data-fzip]'); if (!b2) return null; noHl(); return show(b2).then(function () { hl(b2, 5); return moveTo(b2); }).then(function () { return wait(1600); });
        });
      })
      .then(function () {
        noHl(); cur.hidden = true;
        cap.innerHTML = '<div><b>여기까지가 전부입니다.</b> 이제 내 재료와 내 양식으로 해 보세요. 넣은 파일은 이 컴퓨터 밖으로 나가지 않습니다.</div><div class="dm-row"><span></span><span style="display:flex;gap:8px"><button type="button" class="ghost" id="dmAgain">처음부터 다시 보기</button><button type="button" id="dmClear">비우고 내 파일로 시작</button></span></div>';
        $('#dmAgain').onclick = function () { location.replace(location.pathname + '?demo'); };
        $('#dmClear').onclick = function () { location.replace(location.pathname); };
        running = false; document.removeEventListener('keydown', onKey);
      })
      .catch(function () { finish(); });
  }
  function onKey(e) { if (e.key === 'Escape') stop(); }
  function stop() { stopFlag = true; finish(); }
  function finish() {
    running = false; document.removeEventListener('keydown', onKey);
    if (ring) ring.hidden = true; if (cur) cur.hidden = true;
    if (cap) { cap.innerHTML = '<div>시연을 멈췄습니다. 넣어 둔 예시는 그대로 두었으니 이어서 눌러 보셔도 됩니다.</div><div class="dm-row"><span></span><span style="display:flex;gap:8px"><button type="button" class="ghost" id="dmClose">닫기</button><button type="button" id="dmClear2">비우고 내 파일로 시작</button></span></div>'; $('#dmClose').onclick = function () { cap.remove(); cap = null; cur.remove(); cur = null; ring.remove(); ring = null; }; $('#dmClear2').onclick = function () { location.replace(location.pathname); }; }
  }
  function begin() {
    var used = window._mm && window._mm.state().srcs.length;
    if (used) { if (!confirm('시연은 빈 화면에서 시작합니다. 지금 넣은 재료와 그릇을 비울까요?')) return; location.replace(location.pathname + '?demo'); return; }
    window.scrollTo(0, 0); run();
  }
  var btn = document.getElementById('demoBtn'); if (btn) btn.onclick = begin;
  if (/[?&]demo\b/.test(location.search)) { history.replaceState(null, '', location.pathname); setTimeout(run, 600); }
  window._demo = { run: run, stop: stop };
})();
