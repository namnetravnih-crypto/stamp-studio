/* Ассистент — «Инструменты»: плавающая иконка в правом верхнем углу рабочей зоны.
   При наведении мыши раскрывается барабан с инструментами (листается колесом мыши или стрелками ↑ ↓),
   выбор инструмента сворачивает барабан обратно в иконку. Иконка стоит на одном месте на всех страницах;
   чтобы переместить, нужно зажать ПРАВУЮ кнопку мыши и перетащить (место запоминается).
   Сами кнопки не копируются, а переносятся в барабан целиком — все обработчики остаются рабочими. */
(function () {
  'use strict';
  var POS_KEY = 'assistant.toolwheel.pos';
  var ITEM_H = 54, STEP = 21, VISIBLE_SPAN = 3.4;     // высота строки, угол между строками, сколько строк видно по обе стороны
  var R = ITEM_H / (STEP * Math.PI / 180);
  var K = function () { return (window.UIZoom && window.UIZoom.k()) || 1; };
  var ICON = '<svg viewBox="0 0 64 64" width="34" height="34" aria-hidden="true"><g fill="none" stroke="currentColor" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round">' +
    '<circle cx="24" cy="24" r="7"/><path d="M24 8v5M24 35v5M8 24h5M35 24h5M12.7 12.7l3.5 3.5M31.8 31.8l3.5 3.5M35.3 12.7l-3.5 3.5M16.2 31.8l-3.5 3.5"/>' +
    '<path d="M38 52l12-12a7 7 0 0 0 9-9l-5 5-5-1-1-5 5-5a7 7 0 0 0-9 9L28 50a3.4 3.4 0 0 0 5 5z" transform="translate(-4 2)"/></g></svg>';
  var inst = null;

  function create(opts) {
    if (inst) return inst;
    opts = opts || {};
    var root = document.createElement('div'); root.className = 'tw'; root.id = 'toolWheel';
    root.innerHTML = '<button type="button" class="tw-icon" aria-haspopup="true" aria-expanded="false" title="Инструменты. Наведите мышь — раскроется. Правой кнопкой мыши зажмите и перетащите, чтобы переместить.">' + ICON + '<span>Инструменты</span></button>' +
      '<div class="tw-panel" hidden><div class="tw-head">Инструменты<small>листайте колесом мыши</small></div><div class="tw-drum"><div class="tw-fade tw-fade-t"></div><div class="tw-fade tw-fade-b"></div><div class="tw-mid"></div><div class="tw-lists"></div></div><div class="tw-hint"></div></div>';
    document.body.appendChild(root);
    var icon = root.querySelector('.tw-icon'), panel = root.querySelector('.tw-panel'), lists = root.querySelector('.tw-lists'), hintBox = root.querySelector('.tw-hint');
    var groups = {}, active = null, pos = 0, target = 0, raf = 0, open = false, closeT = 0, sliding = false;

    function items() { return active && groups[active] ? groups[active].items : []; }
    function render() {
      var arr = items(), n = arr.length;
      for (var i = 0; i < n; i++) {
        var d = i - pos, a = d * STEP, ad = Math.abs(d), w = arr[i].wrap;
        if (ad > VISIBLE_SPAN) { w.style.opacity = '0'; w.style.pointerEvents = 'none'; w.style.visibility = 'hidden'; continue; }
        var rad = a * Math.PI / 180, y = R * Math.sin(rad), s = 0.82 + 0.18 * Math.cos(rad);
        w.style.visibility = 'visible'; w.style.pointerEvents = 'auto';
        w.style.opacity = String(Math.max(0.35, Math.cos(rad)));
        w.style.transform = 'translateY(' + y.toFixed(1) + 'px) rotateX(' + (-a).toFixed(1) + 'deg) scale(' + s.toFixed(3) + ')';
        w.style.zIndex = String(100 - Math.round(ad * 10));
        w.classList.toggle('tw-cur', ad < 0.5);
      }
    }
    function tick() {
      var diff = target - pos;
      if (Math.abs(diff) < 0.004) { pos = target; raf = 0; render(); return; }
      pos += diff * 0.22; render(); raf = requestAnimationFrame(tick);
    }
    function go(t) { var n = items().length; target = Math.max(0, Math.min(Math.max(0, n - 1), t)); if (!raf) raf = requestAnimationFrame(tick); }

    function build(id) {
      var g = groups[id]; if (!g) return;
      g.items = []; g.box.innerHTML = '';
      g.nodes.forEach(function (node) {
        if (!node || node.nodeType !== 1 || node.hidden || node.classList.contains('tb-sep')) return;
        var w = document.createElement('div'); w.className = 'tw-item'; w.appendChild(node); g.box.appendChild(w);
        g.items.push({ wrap: w, node: node });
      });
    }
    /* Группа инструментов: nodes — элементы, которые переносятся в барабан. */
    function addGroup(id, nodes) {
      var box = document.createElement('div'); box.className = 'tw-list'; box.dataset.g = id; box.hidden = true; lists.appendChild(box);
      groups[id] = { nodes: nodes, box: box, items: [] }; build(id);
      if (!active) setActive(id);
      return groups[id];
    }
    function setNodes(id, nodes) { if (!groups[id]) return addGroup(id, nodes); groups[id].nodes = nodes; build(id); if (active === id) { pos = target = Math.min(target, Math.max(0, groups[id].items.length - 1)); render(); } syncVisible(); return groups[id]; }
    function setActive(id) {
      active = id; Object.keys(groups).forEach(function (k) { groups[k].box.hidden = k !== id; });
      pos = target = 0; render(); syncVisible();
    }
    function syncVisible() { root.hidden = !(active && groups[active] && groups[active].items.length); }
    function setHint(node) { hintBox.innerHTML = ''; if (node) hintBox.appendChild(node); }

    /* --- открыть / свернуть --- */
    function place() {
      var r = root.getBoundingClientRect(), up = (r.top + r.height / 2) / K() > (window.innerHeight / K()) / 2;
      root.classList.toggle('tw-up', up);
    }
    function openMenu() {
      clearTimeout(closeT); if (open) return; place(); open = true; panel.hidden = false; root.classList.add('tw-open'); icon.setAttribute('aria-expanded', 'true');
      pos = target = Math.min(Math.round(pos), Math.max(0, items().length - 1)); render();
    }
    function closeMenu() { clearTimeout(closeT); if (!open) return; open = false; root.classList.remove('tw-open'); icon.setAttribute('aria-expanded', 'false'); setTimeout(function () { if (!open) panel.hidden = true; }, 260); }
    function softClose() { clearTimeout(closeT); closeT = setTimeout(function () { if (!sliding && !dragging) closeMenu(); }, 450); }

    root.addEventListener('mouseenter', function () { if (!dragging) openMenu(); });
    root.addEventListener('mouseleave', softClose);
    icon.addEventListener('click', function () { open ? closeMenu() : openMenu(); });
    root.addEventListener('wheel', function (e) {
      if (!open) return;
      e.preventDefault(); e.stopPropagation();
      var dy = e.deltaY; if (e.deltaMode === 1) dy *= 33;
      var step = Math.abs(dy) < 40 ? dy / 90 : (dy > 0 ? 1 : -1);        // тачпад — плавно, колесо — по одному пункту
      go(Math.round(target * 1000) / 1000 + step);
    }, { passive: false });
    document.addEventListener('keydown', function (e) {
      if (!open) return;
      if (e.key === 'Escape') closeMenu();
      else if (e.key === 'ArrowDown') { e.preventDefault(); go(Math.round(target) + 1); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); go(Math.round(target) - 1); }
    });
    /* выбор инструмента сворачивает барабан; движение ползунков — нет */
    panel.addEventListener('pointerdown', function (e) { if (e.target.closest && e.target.closest('input[type=range]')) { sliding = true; var up = function () { sliding = false; window.removeEventListener('pointerup', up); window.removeEventListener('pointercancel', up); }; window.addEventListener('pointerup', up); window.addEventListener('pointercancel', up); } });
    panel.addEventListener('click', function (e) {
      var t = e.target; if (!t.closest || t.closest('.tw-head') || t.closest('.tw-hint')) return;
      if (t.closest('input[type=range]')) return;
      var w = t.closest('.tw-item'); if (!w) return;
      if (t.closest('input[type=text],input[type=number],select,textarea')) return;
      setTimeout(closeMenu, 120);
    });

    /* --- перетаскивание правой кнопкой --- */
    var dragging = false, sx = 0, sy = 0, startL = 0, startT = 0;
    function applyPos(left, top) {
      var vw = window.innerWidth / K(), vh = window.innerHeight / K(), w = icon.offsetWidth || 70, h = icon.offsetHeight || 70;
      left = Math.max(4, Math.min(vw - w - 4, left)); top = Math.max(4, Math.min(vh - h - 4, top));
      root.style.left = left + 'px'; root.style.top = top + 'px'; root.style.right = 'auto';
      return { rx: vw - left - w, ty: top };
    }
    function defaultPos() {
      var tb = document.querySelector('.topbar'), top = tb ? tb.getBoundingClientRect().bottom / K() + 12 : 80;
      return { rx: 22, ty: top };
    }
    function restore() {
      var p = null; try { p = JSON.parse(localStorage.getItem(POS_KEY) || 'null'); } catch (e) {}
      if (!p || typeof p.rx !== 'number' || typeof p.ty !== 'number') p = defaultPos();
      var vw = window.innerWidth / K(), w = icon.offsetWidth || 70;
      applyPos(vw - p.rx - w, p.ty);
    }
    root.addEventListener('contextmenu', function (e) { e.preventDefault(); });
    root.addEventListener('mousedown', function (e) {
      if (e.button !== 2) return;
      e.preventDefault(); dragging = true; closeMenu(); root.classList.add('tw-drag');
      sx = e.clientX; sy = e.clientY; var r = root.getBoundingClientRect(); startL = r.left / K(); startT = r.top / K();
    });
    window.addEventListener('mousemove', function (e) { if (!dragging) return; applyPos(startL + (e.clientX - sx) / K(), startT + (e.clientY - sy) / K()); });
    window.addEventListener('mouseup', function (e) {
      if (!dragging) return; dragging = false; root.classList.remove('tw-drag');
      var r = root.getBoundingClientRect(), vw = window.innerWidth / K();
      try { localStorage.setItem(POS_KEY, JSON.stringify({ rx: vw - r.left / K() - r.width / K(), ty: r.top / K() })); } catch (_) {}
      place();
    });
    window.addEventListener('resize', restore);
    restore(); setTimeout(restore, 300);

    inst = { addGroup: addGroup, setNodes: setNodes, setActive: setActive, setHint: setHint, open: openMenu, close: closeMenu, root: root, refresh: function () { build(active); render(); syncVisible(); }, has: function (id) { return !!groups[id]; }, active: function () { return active; } };
    return inst;
  }
  window.ToolWheel = { create: create, get: function () { return inst; } };
})();
