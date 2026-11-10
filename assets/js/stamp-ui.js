/* Stamp Studio: основная группа инструментов переезжает в плавающую иконку «Инструменты» (барабан). */
(function () {
  'use strict';
  var tb = document.querySelector('.toolbar'), row = tb && tb.querySelector('.tb-row');
  if (!tb || !row || !window.ToolWheel) return;
  var nodes = Array.prototype.slice.call(row.children), hint = document.getElementById('toolHint');
  var w = window.ToolWheel.create();
  w.addGroup('stamp', nodes);
  if (hint) w.setHint(hint);
  tb.style.display = 'none';
  window.dispatchEvent(new Event('resize'));
})();
