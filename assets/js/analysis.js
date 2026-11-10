/* Ассистент — программа «Анализ документации»: интерфейс. Ядро — analysis-core.js */
(function () {
  'use strict';
  const A = window.AnalysisCore, E = window.AssistantEngine;
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const extOf = n => (n.split('.').pop() || '').toLowerCase();
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const state = { files: [], res: null, blobs: new Map(), sel: new Set(), seq: 0 };

  function setStatus(t, cls) { const el = $('anStatus'); el.textContent = t || ''; el.className = 'status' + (cls ? ' ' + cls : ''); }
  function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); }

  /* ---------- чтение файлов ---------- */
  async function pdfToText(buf) {
    const pdfjs = window.pdfjsLib; if (!pdfjs) throw new Error('Не загрузилась библиотека PDF');
    pdfjs.GlobalWorkerOptions.workerSrc = 'assets/vendor/pdfjs/pdf.worker.min.js';
    const pdf = await pdfjs.getDocument({ data: buf }).promise, lines = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const content = await (await pdf.getPage(i)).getTextContent(), rows = [];
      for (const it of content.items) { const y = Math.round(it.transform[5] / 3), x = it.transform[4]; let row = rows.find(r => r.y === y); if (!row) rows.push(row = { y, parts: [] }); row.parts.push({ x, w: it.width, s: it.str }); }
      rows.sort((a, b) => b.y - a.y).forEach(r => {
        r.parts.sort((a, b) => a.x - b.x); let out = '', prevEnd = null;
        for (const p of r.parts) { if (!p.s.trim()) continue; out += (prevEnd !== null ? (p.x - prevEnd > 12 ? ' | ' : ' ') : '') + p.s; prevEnd = p.x + p.w; }
        if (out.trim()) lines.push(out.trim());
      });
    }
    return lines.join('\n');
  }
  async function readSource(file) {
    const ext = extOf(file.name), buf = await file.arrayBuffer();
    if (ext === 'docx') return { src: await A.loadDocx(buf, file.name), note: '' };
    if (ext === 'doc' || ext === 'dot') {
      if (!window.Doc97) throw new Error('Не загрузился модуль чтения .doc');
      const out = await window.Doc97.toDocx(buf);
      return { src: await A.loadDocx(out, file.name), note: 'Старый формат .doc прочитан упрощённо: текст и таблицы сохраняются, сноски и сложное оформление могут не перенестись. Если важно оформление, пересохраните файл в .docx.' };
    }
    if (ext === 'xlsx') return { src: A.loadText(await E.xlsxToText(buf), file.name, 'xlsx'), note: 'Из Excel формы вырезаются как текст и таблицы без исходного оформления.' };
    if (ext === 'xls') return { src: A.loadText(await E.xlsxToText((await E.xlsToXlsx(buf)).data), file.name, 'xlsx'), note: 'Из Excel формы вырезаются как текст и таблицы без исходного оформления.' };
    if (ext === 'pdf') return { src: A.loadText(await pdfToText(buf), file.name, 'pdf'), note: 'Из PDF формы вырезаются только как текст: исходное оформление и сноски в Word не переносятся.' };
    if (ext === 'txt') return { src: A.loadText(new TextDecoder('utf-8').decode(buf), file.name, 'text'), note: '' };
    throw new Error('Формат .' + ext + ' не поддерживается. Нужны .doc, .docx, .xls, .xlsx, .pdf');
  }

  async function addFiles(list) {
    const files = Array.from(list || []); if (!files.length) return;
    for (const f of files) {
      if (state.files.some(x => x.name === f.name && x.size === f.size)) continue;
      const rec = { id: ++state.seq, name: f.name, size: f.size, status: 'read', src: null, note: '', err: '' };
      state.files.push(rec); renderFiles();
      try { const r = await readSource(f); rec.src = r.src; rec.note = r.note; rec.status = 'ok'; }
      catch (e) { rec.status = 'err'; rec.err = e && e.message || String(e); }
      renderFiles();
    }
    run();
  }
  function renderFiles() {
    $('anFiles').innerHTML = state.files.map(f => `<li data-id="${f.id}"><span class="name" title="${esc(f.name)}">${esc(f.name)}</span>` +
      (f.status === 'read' ? '<span class="pill hl">читаю…</span>' : f.status === 'err' ? `<span class="pill err" title="${esc(f.err)}">ошибка: ${esc(f.err)}</span>` : `<span class="pill ok">${f.src.items.length} эл.</span>`) +
      `<button class="x" type="button" data-rm="${f.id}" aria-label="Убрать файл">×</button></li>`).join('');
  }

  /* ---------- анализ ---------- */
  function opts() { return { includeDelims: $('optDelims').checked, titleForms: $('optTitles').checked }; }
  function run() {
    state.blobs.clear();
    const srcs = state.files.filter(f => f.status === 'ok').map(f => f.src);
    if (!srcs.length) { state.res = null; $('anRes').hidden = true; setStatus(state.files.length ? 'Нет файлов, которые удалось прочитать.' : '', state.files.length ? 'err' : ''); return; }
    try { state.res = A.analyze(srcs, opts()); } catch (e) { console.error(e); setStatus('Не удалось проанализировать: ' + (e && e.message || e), 'err'); return; }
    state.sel = new Set(state.res.forms.map((_, i) => i));
    render();
    const r = state.res;
    setStatus(`Готово. Состав заявки — ${r.composition.length}, требования к участнику — ${r.requirements.length}, формы — ${r.forms.length}.`, 'ok');
  }

  function previewOf(q) { return q.text.replace(/\s+/g, ' ').slice(0, 260); }
  function quoteRow(q, kind, i) {
    return `<div class="an-item"><div class="info"><div class="ttl">${q.num ? '<span class="tag e">п. ' + esc(q.num) + '</span>' : ''}${esc(q.title)}</div>` +
      `<div class="meta"><span class="tag c">${esc(q.file)}</span><span class="tag">найдено по: ${esc(q.kw)}</span>${q.inTable ? '<span class="tag">строка таблицы</span>' : ''}${q.alt && q.alt.length ? '<span class="tag">также: ' + esc(q.alt.join('; ')) + '</span>' : ''}</div>` +
      `<div class="snip">${esc(previewOf(q))}</div></div><div class="acts"><button class="btn small" type="button" data-q="${kind}:${i}" data-tip="Открыть полный текст найденного раздела или списка.">Показать цитату</button></div></div>`;
  }
  function formRow(f, i) {
    const warn = f.srcRef.kind !== 'docx' ? 'Из этого типа файла форма вырезается только как текст.' : '';
    return `<div class="an-item form"><input type="checkbox" data-sel="${i}" ${state.sel.has(i) ? 'checked' : ''} aria-label="Выбрать форму"><div class="info"><div class="ttl">${esc(A.formFileName(f, i).replace(/\.docx$/, ''))}</div>` +
      `<div class="meta"><span class="tag c">${esc(f.file)}</span><span class="tag e">${esc(f.label)}</span><span class="tag">элементов: ${f.items}</span>${f.tables ? '<span class="tag">таблиц: ' + f.tables + '</span>' : ''}${f.footnotes ? '<span class="tag">сносок: ' + f.footnotes + '</span>' : ''}` +
      `<span class="tag">конец: ${({ marker: 'по маркеру «Конец формы»', next: 'перед следующей формой', section: 'перед новым разделом', signature: 'после блока подписи', eof: 'конец файла' })[f.endBy] || ''}</span><span class="tag">${esc(f.why)}</span></div>${warn ? '<div class="warn-note">' + esc(warn) + '</div>' : ''}</div>` +
      `<div class="acts"><button class="btn small" type="button" data-pv="${i}" data-tip="Посмотреть, что попадёт в файл.">Просмотр</button><button class="btn small primary" type="button" data-dl="${i}" data-tip="Скачать эту форму отдельным файлом Word.">Скачать .docx</button></div></div>`;
  }
  function render() {
    const r = state.res; $('anRes').hidden = false;
    const set = (cntId, listId, arr, fn, empty) => { $(cntId).textContent = arr.length ? 'найдено: ' + arr.length : 'не найдено'; $(cntId).classList.toggle('has', !!arr.length); $(listId).innerHTML = arr.length ? arr.map(fn).join('') : `<div class="an-empty">${empty}</div>`; };
    set('cntComp', 'listComp', r.composition, (q, i) => quoteRow(q, 'c', i), 'Раздел о составе заявки не найден. Попробуйте загрузить основной файл документации (информационную карту, извещение).');
    set('cntReq', 'listReq', r.requirements, (q, i) => quoteRow(q, 'r', i), 'Требования к участнику не найдены. Попробуйте загрузить основной файл документации.');
    set('cntForms', 'listForms', r.forms, formRow, 'Формы не найдены. Формы определяются по словам «Форма №…», «Приложение №…», «Образец», «Конец формы» и по названиям вроде «ЗАЯВКА НА УЧАСТИЕ».');
    $('formBar').hidden = !r.forms.length; $('selAll').checked = state.sel.size === r.forms.length;
    $('dlSel').disabled = $('dlZip').disabled = !state.sel.size;
  }

  /* ---------- окна ---------- */
  let modalText = '';
  function openModal(title, meta, bodyHtml, text, docv, note) {
    $('anMT').textContent = title; $('anMM').textContent = meta || ''; $('anMB').textContent = note || '';
    const b = $('anBody'); b.className = 'an-body' + (docv ? ' docv' : ''); b.innerHTML = bodyHtml; modalText = text || '';
    $('anCopy').hidden = !text; $('anModal').hidden = false; b.scrollTop = 0;
  }
  const closeModal = () => { $('anModal').hidden = true; $('anBody').innerHTML = ''; };
  function quoteHtml(q) {
    const lines = q.text.split('\n'); let out = '', i = 0;
    while (i < lines.length) {
      if (/ \| /.test(lines[i]) && q.inTable !== true) {
        out += '<div class="qr">'; while (i < lines.length && / \| /.test(lines[i])) out += '<div>' + lines[i++].split(' | ').map(c => '<span>' + esc(c) + '</span>').join('') + '</div>'; out += '</div>';
      } else out += '<p class="ql">' + esc(lines[i++]) + '</p>';
    }
    return out;
  }
  function showQuote(kind, i) {
    const q = (kind === 'c' ? state.res.composition : state.res.requirements)[i]; if (!q) return;
    openModal((kind === 'c' ? 'Состав заявки' : 'Требования к участнику') + (q.num ? ' — п. ' + q.num : ''), 'Файл: ' + q.file + ' · найдено по: «' + q.kw + '»',
      quoteHtml(q), q.text, false, q.truncated ? 'Цитата очень длинная и показана не полностью — откройте исходный файл.' : '');
  }

  async function blobOf(i) {
    if (state.blobs.has(i)) return state.blobs.get(i);
    const f = state.res.forms[i]; const bytes = await A.exportForm(f);
    const blob = new Blob([bytes], { type: DOCX_MIME }); state.blobs.set(i, blob); return blob;
  }
  async function showForm(i) {
    const f = state.res.forms[i];
    openModal(A.formFileName(f, i).replace(/\.docx$/, ''), 'Файл: ' + f.file + ' · ' + f.label, '<p class="ql">Готовлю просмотр…</p>', '', true);
    try {
      const blob = await blobOf(i); const body = $('anBody'); body.innerHTML = '';
      if (window.docx && window.docx.renderAsync) { const host = document.createElement('div'); body.append(host); await window.docx.renderAsync(blob, host, null, { className: 'docx', inWrapper: true, breakPages: true, useBase64URL: true }); }
      else throw new Error('нет docx-preview');
    } catch (e) { $('anBody').className = 'an-body'; $('anBody').innerHTML = f.preview.split('\n').map(l => '<p class="ql">' + esc(l) + '</p>').join(''); }
  }
  async function saveOne(i) { const f = state.res.forms[i]; setStatus('Готовлю файл…'); try { download(await blobOf(i), A.formFileName(f, i)); setStatus('Файл «' + A.formFileName(f, i) + '» скачан.', 'ok'); } catch (e) { setStatus('Не удалось вырезать форму: ' + (e && e.message || e), 'err'); } }
  async function saveSelected() {
    const ids = Array.from(state.sel).sort((a, b) => a - b);
    for (const i of ids) { await saveOne(i); await new Promise(r => setTimeout(r, 350)); }
  }
  async function saveZip() {
    const ids = Array.from(state.sel).sort((a, b) => a - b); if (!ids.length) return;
    setStatus('Собираю архив…');
    try {
      const zip = new JSZip(), used = new Set(), multi = new Set(state.res.forms.map(f => f.file)).size > 1;
      for (const i of ids) {
        const f = state.res.forms[i]; let name = A.formFileName(f, i); const dir = multi ? f.file.replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]/g, ' ').trim() + '/' : '';
        let path = dir + name, k = 2; while (used.has(path)) path = dir + name.replace(/\.docx$/, ' (' + (k++) + ').docx'); used.add(path);
        zip.file(path, await blobOf(i));
      }
      download(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }), 'Найденные_формы.zip'); setStatus('Архив скачан: форм — ' + ids.length + '.', 'ok');
    } catch (e) { setStatus('Не удалось собрать архив: ' + (e && e.message || e), 'err'); }
  }

  /* ---------- события ---------- */
  function init() {
    const drop = $('anDrop'), inp = $('anFile');
    inp.addEventListener('change', () => { addFiles(inp.files); inp.value = ''; });
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => addFiles(e.dataTransfer.files));
    $('anFiles').addEventListener('click', e => { const b = e.target.closest('[data-rm]'); if (!b) return; state.files = state.files.filter(f => f.id !== +b.dataset.rm); renderFiles(); run(); });
    $('anClear').addEventListener('click', () => { state.files = []; state.res = null; state.blobs.clear(); renderFiles(); $('anRes').hidden = true; setStatus(''); });
    $('optDelims').addEventListener('change', run); $('optTitles').addEventListener('change', run);
    $('anRes').addEventListener('click', e => {
      const q = e.target.closest('[data-q]'); if (q) { const [k, i] = q.dataset.q.split(':'); return showQuote(k, +i); }
      const pv = e.target.closest('[data-pv]'); if (pv) return showForm(+pv.dataset.pv);
      const dl = e.target.closest('[data-dl]'); if (dl) return saveOne(+dl.dataset.dl);
    });
    $('anRes').addEventListener('change', e => {
      const s = e.target.closest('[data-sel]'); if (s) { const i = +s.dataset.sel; s.checked ? state.sel.add(i) : state.sel.delete(i); $('selAll').checked = state.sel.size === state.res.forms.length; $('dlSel').disabled = $('dlZip').disabled = !state.sel.size; }
    });
    $('selAll').addEventListener('change', e => { state.sel = e.target.checked ? new Set(state.res.forms.map((_, i) => i)) : new Set(); document.querySelectorAll('[data-sel]').forEach(c => c.checked = e.target.checked); $('dlSel').disabled = $('dlZip').disabled = !state.sel.size; });
    $('dlSel').addEventListener('click', saveSelected); $('dlZip').addEventListener('click', saveZip);
    $('anOk').addEventListener('click', closeModal); $('anMX').addEventListener('click', closeModal);
    $('anModal').addEventListener('mousedown', e => { if (e.target === $('anModal')) closeModal(); });
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('anModal').hidden) closeModal(); });
    $('anCopy').addEventListener('click', async () => { try { await navigator.clipboard.writeText(modalText); $('anCopy').textContent = 'Скопировано'; } catch (e) { const t = document.createElement('textarea'); t.value = modalText; document.body.append(t); t.select(); document.execCommand('copy'); t.remove(); $('anCopy').textContent = 'Скопировано'; } setTimeout(() => $('anCopy').textContent = 'Копировать текст', 1600); });
    window.__analysis = { state, addFiles };
  }
  init();
})();
