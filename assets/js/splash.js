/* Ассистент — заставка главной страницы: авторская анимация по мотиву «Сотворения Адама»
   (Адам за офисным столом, Бог тянется к нему с ноутбуком). Рисунок собственный, в плоском стиле. */
(function () {
  'use strict';
  var KEY = 'assistant.splash.seen';
  var seen = false; try { seen = sessionStorage.getItem(KEY) === '1'; } catch (e) {}
  if (seen && location.hash !== '#splash') return;
  var svg = '<svg viewBox="0 0 800 450" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Адам за рабочим столом и Бог с ноутбуком тянутся друг к другу">' +
    '<defs><linearGradient id="spSky" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1c2a4d"/><stop offset=".55" stop-color="#7b6a8f"/><stop offset="1" stop-color="#e3b88a"/></linearGradient>' +
    '<radialGradient id="spGlow"><stop offset="0" stop-color="#fff6c8" stop-opacity="1"/><stop offset="1" stop-color="#ffd36b" stop-opacity="0"/></radialGradient>' +
    '<linearGradient id="spScr" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#bfe3ff"/><stop offset="1" stop-color="#5aa0e8"/></linearGradient></defs>' +
    '<rect width="800" height="450" fill="url(#spSky)"/>' +
    '<g class="sp-stars" fill="#fff"><circle cx="90" cy="50" r="1.6"/><circle cx="210" cy="90" r="1.2"/><circle cx="330" cy="40" r="1.5"/><circle cx="470" cy="70" r="1.2"/><circle cx="560" cy="30" r="1.6"/><circle cx="730" cy="80" r="1.3"/></g>' +
    // облака
    '<g fill="#f6e9dc" opacity=".92"><ellipse cx="610" cy="340" rx="190" ry="34"/><ellipse cx="540" cy="318" rx="90" ry="30"/><ellipse cx="690" cy="316" rx="80" ry="28"/></g>' +
    '<g fill="#fff" opacity=".35"><ellipse cx="150" cy="130" rx="90" ry="14"/><ellipse cx="420" cy="110" rx="70" ry="10"/></g>' +
    // офис Адама: пол, стол, стул
    '<rect x="0" y="372" width="420" height="78" fill="#3b2e2a"/><rect x="0" y="372" width="420" height="4" fill="#5b463d"/>' +
    '<g><rect x="60" y="292" width="270" height="12" rx="2" fill="#8a5a34"/><rect x="72" y="304" width="10" height="68" fill="#6e4527"/><rect x="308" y="304" width="10" height="68" fill="#6e4527"/>' +
    '<rect x="95" y="262" width="62" height="30" rx="3" fill="#2f3948"/><rect x="101" y="267" width="50" height="20" fill="#7fb6ee"/><rect x="118" y="292" width="16" height="3" fill="#2f3948"/>' +
    '<rect x="196" y="272" width="30" height="20" rx="2" fill="#e9e2d0"/><rect x="236" y="276" width="12" height="16" rx="2" fill="#c0584a"/></g>' +
    '<rect x="372" y="330" width="8" height="42" fill="#222"/><rect x="342" y="330" width="68" height="8" rx="3" fill="#2a2f3a"/>' +
    // Адам
    '<g class="sp-adam"><path d="M350 330 L350 232 Q350 214 370 210 L396 210 Q410 214 408 236 L402 330 Z" fill="#e7e9ee"/>' +
    '<path d="M365 214 L380 252 L392 214 Z" fill="#c9ccd6"/><path d="M378 220 L382 224 L380 256 L376 224 Z" fill="#1f3a6b"/>' +
    '<path d="M352 332 L352 372 L372 372 L372 336 Z" fill="#2b3550"/><path d="M386 332 L388 372 L410 372 L402 332 Z" fill="#2b3550"/>' +
    '<ellipse cx="364" cy="374" rx="22" ry="5" fill="#161616"/><ellipse cx="402" cy="374" rx="22" ry="5" fill="#161616"/>' +
    '<circle cx="380" cy="190" r="22" fill="#e8b997"/><path d="M358 188 Q360 160 384 162 Q404 164 402 190 Q392 172 378 174 Q364 176 358 188Z" fill="#4a3321"/>' +
    '<circle cx="390" cy="192" r="1.8" fill="#2a1a10"/><path d="M386 202 Q392 205 396 201" stroke="#9a5c46" stroke-width="1.6" fill="none"/>' +
    '<path class="sp-arm-a" d="M396 226 Q430 222 462 206" stroke="#e7e9ee" stroke-width="15" stroke-linecap="round" fill="none"/>' +
    '<path d="M456 208 L478 203" stroke="#e8b997" stroke-width="7" stroke-linecap="round"/><path d="M470 202 L494 199" stroke="#e8b997" stroke-width="3.5" stroke-linecap="round"/></g>' +
    // Бог
    '<g class="sp-god"><g class="sp-bob">' +
    '<path d="M560 150 Q520 210 560 300 Q650 318 740 300 Q770 220 720 140 Q640 100 560 150Z" fill="#b9444f"/>' +
    '<path d="M590 170 Q580 240 600 296 L650 300 Q640 230 650 170Z" fill="#a03642" opacity=".7"/>' +
    '<path d="M610 170 Q640 150 676 168 L690 250 Q650 276 612 250Z" fill="#f3e8d6"/>' +
    '<circle cx="600" cy="132" r="26" fill="#efc9a8"/><path d="M572 130 Q566 92 604 90 Q640 92 628 134 Q628 108 604 108 Q584 110 580 140Z" fill="#fafafa"/>' +
    '<path d="M578 146 Q574 196 604 202 Q634 190 628 148 Q612 170 596 164 Q584 160 578 146Z" fill="#fafafa"/><circle cx="590" cy="130" r="1.8" fill="#2a1a10"/><path d="M580 124 L592 122" stroke="#d9d9d9" stroke-width="2.4"/>' +
    '<path class="sp-arm-g" d="M612 182 Q560 190 520 206" stroke="#efc9a8" stroke-width="13" stroke-linecap="round" fill="none"/>' +
    '<path d="M516 207 L496 202" stroke="#efc9a8" stroke-width="6" stroke-linecap="round"/><path d="M510 205 L498 200" stroke="#efc9a8" stroke-width="3" stroke-linecap="round"/>' +
    // ноутбук в другой руке
    '<path d="M690 190 Q716 170 724 126" stroke="#efc9a8" stroke-width="13" stroke-linecap="round" fill="none"/>' +
    '<g class="sp-laptop" transform="rotate(-8 735 104)"><path d="M700 104 L770 104 L778 70 L708 70Z" fill="#2b2f3a"/><path d="M706 100 L764 100 L771 74 L713 74Z" fill="url(#spScr)" class="sp-screen"/><path d="M696 106 L784 106 L788 114 L692 114Z" fill="#9aa1ad"/><circle cx="742" cy="88" r="5" fill="#fff" opacity=".85"/></g>' +
    '</g></g>' +
    // искра между пальцами
    '<g class="sp-spark"><circle cx="490" cy="201" r="46" fill="url(#spGlow)"/><path d="M490 180 L493 198 L510 201 L493 204 L490 222 L487 204 L470 201 L487 198Z" fill="#fff8d6"/></g>' +
    '</svg>';
  var box = document.createElement('div'); box.className = 'splash'; box.setAttribute('role', 'dialog'); box.setAttribute('aria-label', 'Заставка');
  box.innerHTML = '<div class="splash-art">' + svg + '</div><div class="splash-title"><b>АССИСТЕНТ</b><small>творческое пространство</small></div><button type="button" class="splash-go">Войти →</button>';
  document.body.appendChild(box); document.body.classList.add('splash-on');
  var closed = false, timer;
  function close() { if (closed) return; closed = true; clearTimeout(timer); try { sessionStorage.setItem(KEY, '1'); } catch (e) {} box.classList.add('out'); document.body.classList.remove('splash-on'); setTimeout(function () { box.remove(); }, 700); }
  box.addEventListener('click', close); document.addEventListener('keydown', function (e) { if (e.key === 'Escape' || e.key === 'Enter') close(); });
  timer = setTimeout(close, 9000);
})();
