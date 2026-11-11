/* Ассистент — заставка главной страницы: строгая анимация по теме работы с закупками и строительной документацией.
   Чертёж объекта прорисовывается линиями, на листе договора появляется текст и таблица, затем ставится круглая печать
   и росчерк подписи. Рисунок собственный. Показывается раз за сессию; закрывается кликом, Enter или Esc. */
(function () {
  'use strict';
  var KEY = 'assistant.splash.seen';
  var seen = false; try { seen = sessionStorage.getItem(KEY) === '1'; } catch (e) {}
  if (seen && location.hash !== '#splash') return;
  var NS = '';
  function ln(x1, y1, x2, y2, d, w) { return '<line class="sp-d" pathLength="1" style="animation-delay:' + d + 's" x1="' + x1 + '" y1="' + y1 + '" x2="' + x2 + '" y2="' + y2 + '"' + (w ? ' stroke-width="' + w + '"' : '') + '/>'; }
  var grid = '';
  for (var gx = 0; gx <= 800; gx += 40) grid += '<line x1="' + gx + '" y1="0" x2="' + gx + '" y2="450"/>';
  for (var gy = 0; gy <= 450; gy += 40) grid += '<line x1="0" y1="' + gy + '" x2="800" y2="' + gy + '"/>';
  // строки текста на листе
  var rows = '';
  for (var i = 0; i < 9; i++) rows += ln(318, 150 + i * 13, 318 + (i % 3 === 2 ? 120 : 164), 150 + i * 13, (1.1 + i * 0.12).toFixed(2));
  var svg = '<svg viewBox="0 0 800 450" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Чертёж, договор и печать: анимация заставки">' +
    '<defs><radialGradient id="spBg" cx="50%" cy="45%" r="75%"><stop offset="0" stop-color="#22345f"/><stop offset="1" stop-color="#0c1429"/></radialGradient>' +
    '<path id="spRing" d="M400 232 m-52 0 a52 52 0 1 1 104 0 a52 52 0 1 1 -104 0"/></defs>' +
    '<rect width="800" height="450" fill="url(#spBg)"/>' +
    '<g class="sp-grid" stroke="#c8a45a" stroke-width=".5" opacity=".10">' + grid + '</g>' +
    // левый чертёж: башенный кран
    '<g class="sp-bp" fill="none" stroke="#d8b96e" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">' +
    ln(90, 392, 90, 120, .2) + ln(110, 392, 110, 120, .25) + ln(90, 120, 110, 120, .3) +
    ln(90, 360, 110, 340, .5) + ln(110, 360, 90, 340, .5) + ln(90, 320, 110, 300, .6) + ln(110, 320, 90, 300, .6) + ln(90, 280, 110, 260, .7) + ln(110, 280, 90, 260, .7) + ln(90, 240, 110, 220, .8) + ln(110, 240, 90, 220, .8) + ln(90, 200, 110, 180, .9) + ln(110, 200, 90, 180, .9) +
    ln(40, 120, 250, 120, 1.0) + ln(90, 100, 40, 120, 1.05) + ln(90, 100, 230, 120, 1.1) + ln(60, 120, 60, 160, 1.15) +
    ln(215, 120, 215, 190, 1.25) + ln(205, 190, 225, 190, 1.3) + ln(70, 392, 130, 392, .15, 2.2) + '</g>' +
    // правый чертёж: каркас здания
    '<g class="sp-bp" fill="none" stroke="#d8b96e" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">' +
    ln(580, 392, 580, 170, .3) + ln(650, 392, 650, 170, .35) + ln(720, 392, 720, 170, .4) +
    ln(570, 170, 730, 170, .6) + ln(570, 230, 730, 230, .7) + ln(570, 290, 730, 290, .8) + ln(570, 350, 730, 350, .9) +
    ln(580, 350, 650, 290, 1.0) + ln(650, 350, 580, 290, 1.0) + ln(650, 230, 720, 170, 1.1) + ln(720, 230, 650, 170, 1.1) +
    ln(560, 392, 740, 392, .25, 2.2) + ln(565, 170, 650, 130, 1.3) + ln(650, 130, 735, 170, 1.35) + '</g>' +
    // лист договора
    '<g class="sp-doc"><rect x="290" y="62" width="220" height="326" rx="3" fill="#f7f4ea"/><rect x="290" y="62" width="220" height="326" rx="3" fill="none" stroke="#c8a45a" stroke-width="1.5"/>' +
    '<g stroke="#1c2a4d" stroke-width="2.2" stroke-linecap="round">' + ln(345, 94, 455, 94, .9, 3) + ln(370, 106, 430, 106, 1.0, 2) + '</g>' +
    '<g stroke="#8a93a8" stroke-width="1.6" stroke-linecap="round">' + rows + '</g>' +
    '<g class="sp-tbl" fill="none" stroke="#8a93a8" stroke-width="1.2"><rect x="318" y="275" width="164" height="42"/><line x1="318" y1="289" x2="482" y2="289"/><line x1="318" y1="303" x2="482" y2="303"/><line x1="372" y1="275" x2="372" y2="317"/><line x1="432" y1="275" x2="432" y2="317"/></g>' +
    '<path class="sp-sign sp-d" pathLength="1" d="M330 358 q10 -22 16 -4 t14 -2 q8 -16 14 0 t16 -6 q10 -10 14 4 t16 -2" fill="none" stroke="#1d3f94" stroke-width="2" stroke-linecap="round"/><line x1="322" y1="364" x2="480" y2="364" stroke="#8a93a8" stroke-width="1"/>' +
    '</g>' +
    // печать
    '<g class="sp-seal"><g transform="rotate(-12 400 232)" opacity=".92"><circle cx="400" cy="232" r="64" fill="none" stroke="#1d3f94" stroke-width="3.2"/><circle cx="400" cy="232" r="58" fill="none" stroke="#1d3f94" stroke-width="1"/><circle cx="400" cy="232" r="36" fill="none" stroke="#1d3f94" stroke-width="1.4"/>' +
    '<text font-family="Cormorant Garamond,Georgia,serif" font-size="11" font-weight="700" letter-spacing="1.6" fill="#1d3f94"><textPath href="#spRing" startOffset="0">АССИСТЕНТ · ДОКУМЕНТЫ · ЗАКУПКИ · </textPath></text>' +
    '<path d="M400 214 l5.6 11.4 12.6 1.8 -9.1 8.9 2.2 12.5 -11.3 -5.9 -11.3 5.9 2.2 -12.5 -9.1 -8.9 12.6 -1.8z" fill="#1d3f94"/></g></g>' +
    '<circle class="sp-ping" cx="400" cy="232" r="64" fill="none" stroke="#e3c783" stroke-width="2"/>' +
    '</svg>';
  var box = document.createElement('div'); box.className = 'splash'; box.setAttribute('role', 'dialog'); box.setAttribute('aria-label', 'Заставка');
  box.innerHTML = '<div class="splash-art">' + svg + '</div><div class="splash-title"><b>АССИСТЕНТ</b><small>творческое пространство</small></div><div class="splash-bar"><i></i></div><button type="button" class="splash-go">Войти →</button>';
  document.body.appendChild(box); document.body.classList.add('splash-on');
  var closed = false, timer;
  function close() { if (closed) return; closed = true; clearTimeout(timer); try { sessionStorage.setItem(KEY, '1'); } catch (e) {} box.classList.add('out'); document.body.classList.remove('splash-on'); setTimeout(function () { box.remove(); }, 800); }
  box.addEventListener('click', close); document.addEventListener('keydown', function (e) { if (e.key === 'Escape' || e.key === 'Enter') close(); });
  timer = setTimeout(close, 7500);
})();
