/* Анализ документации: параметры поиска, «Очистить» и кнопки скачивания форм собраны в плавающей иконке «Инструменты». */
(function () {
  'use strict';
  if (!window.ToolWheel) return;
  var ids = ['optDelims', 'optTitle', 'anClear', 'selAll', 'dlSel', 'dlZip'], nodes = [];
  ids.forEach(function (id) { var n = document.getElementById(id); if (!n) return; nodes.push(n.tagName === 'INPUT' ? n.closest('label') : n); });
  nodes = nodes.filter(Boolean);
  if (!nodes.length) return;
  var w = window.ToolWheel.create(); w.addGroup('analysis', nodes);
  var h = document.createElement('div'); h.textContent = 'Скачивание форм доступно после анализа файлов.'; w.setHint(h);
  var bar = document.getElementById('formBar'); if (bar) bar.style.display = 'none';
})();
