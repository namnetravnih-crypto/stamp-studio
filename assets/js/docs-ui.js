/* Автозаполнение: инструменты наложения печати и правки (показать документ, редактировать, убрать жёлтый маркер,
   убрать печати, забыть печати) каждого блока переезжают в плавающую иконку «Инструменты» (барабан).
   Барабан показывает инструменты того блока, который сейчас раскрыт. */
(function () {
  'use strict';
  if (!window.ToolWheel) return;
  var SEL = ['.ov-show', '.ov-edit', '.ov-unmark', '.ov-clear', '.ov-forget'];
  var wheel = window.ToolWheel.create(), timer = 0;
  function scan() {
    document.querySelectorAll('.blk').forEach(function (blk) {
      if (blk._twDone || !blk.querySelector('.ov-box')) return;
      var nodes = [];
      SEL.forEach(function (s) { var n = blk.querySelector(s); if (n && !n.closest('.tw-item')) nodes.push(n); });
      if (nodes.length < SEL.length - 1) return;                     // коробка ещё не достроена
      blk._twDone = 1;
      var tag = blk.id || ('blk' + Math.random());
      blk.id = tag;
      var box = blk.querySelector('.ov-box'); box._moved = {};
      SEL.forEach(function (s) { var n = blk.querySelector(s); if (n) box._moved[s] = n; });
      wheel.addGroup(tag, nodes);
      var hint = document.createElement('div'); hint.textContent = 'Выберите инструмент: он применится к документу этого блока.'; wheel.setHint(hint);
      sync();
    });
  }
  function sync() {
    var open = document.querySelector('.blk.open') || (document.querySelector('.blk-head[aria-expanded="true"]') || {}).parentNode;
    wheel.setActive(open && open.id ? open.id : '__none');
  }
  function later() { clearTimeout(timer); timer = setTimeout(function () { scan(); sync(); }, 80); }
  document.addEventListener('click', function (e) { if (e.target.closest && e.target.closest('.blk-head')) later(); }, true);
  new MutationObserver(later).observe(document.body, { childList: true, subtree: true });
  later(); setTimeout(later, 600);
})();
