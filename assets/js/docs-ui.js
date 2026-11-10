/* Автозаполнение: навигация по блокам в верхней части экрана (вместо плавающего меню «Инструменты»).
   Панель закреплена сверху: нажмите на название блока — он раскроется и окажется на экране;
   подсвечивается блок, который открыт сейчас. Кнопки печатей и правки теперь стоят прямо в блоке, рядом с документом. */
(function () {
  'use strict';
  var blocks = Array.prototype.slice.call(document.querySelectorAll('.blk'));
  if (!blocks.length) return;
  var nav = document.createElement('nav');
  nav.className = 'pnav'; nav.setAttribute('aria-label', 'Блоки программы');
  var lib = document.querySelector('#libPanel .lib-toggle');
  var html = '<div class="pnav-in">';
  if (lib) html += '<button type="button" class="pnav-b" data-nav="lib"><span class="pn-n">★</span><span>Моя библиотека</span></button>';
  blocks.forEach(function (b, i) {
    var t = b.querySelector('.blk-t b'), n = b.querySelector('.blk-n');
    html += '<button type="button" class="pnav-b" data-nav="' + b.id + '"><span class="pn-n">' + (n ? n.textContent : i + 1) + '</span><span>' + (t ? t.textContent : b.id) + '</span></button>';
  });
  html += '</div><button type="button" class="pnav-b pnav-fold" data-nav="fold" title="Свернуть все блоки">Свернуть всё</button>';
  nav.innerHTML = html;
  var wrap = document.querySelector('main.wrap'), head = wrap.querySelector('.page-head');
  wrap.insertBefore(nav, head ? head.nextSibling : wrap.firstChild);

  function isOpen(b) { var h = b.querySelector('.blk-head'); return h && h.getAttribute('aria-expanded') === 'true'; }
  function sync() {
    nav.querySelectorAll('[data-nav]').forEach(function (btn) {
      var id = btn.dataset.nav, on = false;
      if (id === 'lib') on = !!lib && lib.getAttribute('aria-expanded') === 'true';
      else if (id !== 'fold') { var b = document.getElementById(id); on = !!b && isOpen(b); }
      btn.classList.toggle('on', on);
    });
  }
  function scrollTo(el) {
    var r = el.getBoundingClientRect(), z = (window.UIZoom && window.UIZoom.k && window.UIZoom.k()) || 1;
    window.scrollTo({ top: window.pageYOffset + r.top - (nav.offsetHeight + 10) * 1, behavior: 'smooth' });
  }
  nav.addEventListener('click', function (e) {
    var btn = e.target.closest('[data-nav]'); if (!btn) return; var id = btn.dataset.nav;
    if (id === 'fold') { blocks.forEach(function (b) { if (isOpen(b)) b.querySelector('.blk-head').click(); }); if (lib && lib.getAttribute('aria-expanded') === 'true') lib.click(); setTimeout(sync, 60); return; }
    if (id === 'lib') { if (lib.getAttribute('aria-expanded') !== 'true') lib.click(); setTimeout(function () { sync(); scrollTo(document.getElementById('libPanel')); }, 80); return; }
    var b = document.getElementById(id); if (!b) return;
    if (!isOpen(b)) b.querySelector('.blk-head').click();
    setTimeout(function () { sync(); scrollTo(b); }, 120);
  });
  document.addEventListener('click', function (e) { if (e.target.closest && e.target.closest('.blk-head, .lib-toggle')) setTimeout(sync, 80); }, true);
  sync();
})();
