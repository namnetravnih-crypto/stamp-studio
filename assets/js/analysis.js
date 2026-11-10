/* Ассистент — программа «Анализ документации»: интерфейс. Ядро — analysis-core.js */
(function () {
  'use strict';
  const A = window.AnalysisCore, X = window.AnalysisExtract, E = window.AssistantEngine;
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const extOf = n => (n.split('.').pop() || '').toLowerCase();
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const state = { files: [], res: null, sum: null, blobs: new Map(), sel: new Set(), seq: 0, tab: 'sum' };

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
    if (!srcs.length) { state.res = null; state.sum = null; $('anRes').hidden = true; $('anTabs').hidden = true; setStatus(state.files.length ? 'Нет файлов, которые удалось прочитать.' : '', state.files.length ? 'err' : ''); return; }
    try { state.res = A.analyze(srcs, opts()); state.sum = X.summarize(srcs, state.res); state.res.composition.concat(state.res.requirements).forEach(q => { q.points = X.pointsOf(q); }); } catch (e) { console.error(e); setStatus('Не удалось проанализировать: ' + (e && e.message || e), 'err'); return; }
    state.sel = new Set(state.res.forms.map((_, i) => i));
    render();
    const r = state.res;
    setStatus(`Готово. Состав заявки — ${r.composition.length}, требования к участнику — ${r.requirements.length}, «требование не установлено» — ${state.sum.notSet.length}, формы — ${r.forms.length}.`, 'ok');
  }

  /* ---------- карточки ---------- */
  const NOTSET_LABEL = 'требование не установлено';
  function pointsHtml(q) {
    const pts = q.points || []; if (!pts.length) return '';
    const miss = pts.filter(p => p.notSet).length;
    return `<details class="pts" open><summary>Пункты: ${pts.length}${miss ? ' · <span class="miss-c">не установлено: ' + miss + '</span>' : ''}</summary><div class="pts-in">` +
      pts.map(p => `<div class="pt${p.notSet ? ' miss' : ''}${p.head ? ' head' : ''}">${p.num ? '<b class="pn">' + esc(p.num) + '</b>' : ''}<span class="pt-t">${esc(p.text)}</span>${p.notSet ? '<span class="tag miss" title="' + esc(p.mark || '') + '">' + NOTSET_LABEL + '</span>' : ''}</div>`).join('') + '</div></details>';
  }
  function quoteRow(q, kind, i) {
    const miss = (q.points || []).filter(p => p.notSet).length;
    return `<div class="an-item"><div class="info"><div class="ttl">${q.num ? '<span class="tag e">п. ' + esc(q.num) + '</span>' : ''}${esc(q.title)}</div>` +
      `<div class="meta"><span class="tag c">${esc(q.file)}</span><span class="tag">найдено по: ${esc(q.kw)}</span>${q.inTable ? '<span class="tag">таблица</span>' : ''}${miss ? '<span class="tag miss">' + NOTSET_LABEL + ': ' + miss + '</span>' : ''}${q.alt && q.alt.length ? '<span class="tag">также: ' + esc(q.alt.join('; ')) + '</span>' : ''}</div>` +
      pointsHtml(q) + `</div><div class="acts"><button class="btn small" type="button" data-q="${kind}:${i}" data-tip="Открыть документ: найденный раздел выделен жёлтым, пометки «не установлено» — оранжевым.">Показать цитату</button></div></div>`;
  }
  function missRow(m, i) {
    return `<div class="an-item miss"><div class="info"><div class="ttl">${m.num ? '<span class="tag e">п. ' + esc(m.num) + '</span>' : ''}${esc(m.text)}</div>` +
      `<div class="meta"><span class="tag miss">${NOTSET_LABEL}</span><span class="tag">в документе: «${esc(m.marker)}»</span><span class="tag c">${esc(m.file)}</span>${m.section ? '<span class="tag">' + esc(m.section) + '</span>' : ''}</div></div>` +
      `<div class="acts"><button class="btn small" type="button" data-m="${i}" data-tip="Открыть документ: пункт выделен жёлтым, пометка — оранжевым.">Показать цитату</button></div></div>`;
  }
  const endLabel = f => ({ marker: 'по маркеру «Конец формы»', next: 'перед следующей формой', section: 'перед новым разделом', signature: 'после блока подписи', eof: 'конец файла' })[f.endBy] || '';
  function formRow(f, i) {
    const warn = f.srcRef.kind !== 'docx' ? 'Из этого типа файла форма вырезается только как текст.' : '';
    return `<div class="an-item form"><input type="checkbox" data-sel="${i}" ${state.sel.has(i) ? 'checked' : ''} aria-label="Выбрать форму"><div class="info"><div class="ttl">${esc(A.formFileName(f, i).replace(/\.docx$/, ''))}</div>` +
      `<div class="meta"><span class="tag c">${esc(f.file)}</span><span class="tag e">${esc(f.label)}</span><span class="tag">элементов: ${f.items}</span>${f.tables ? '<span class="tag">таблиц: ' + f.tables + '</span>' : ''}${f.footnotes ? '<span class="tag">сносок: ' + f.footnotes + '</span>' : ''}` +
      `<span class="tag">конец: ${endLabel(f)}</span><span class="tag">${esc(f.why)}</span></div>${warn ? '<div class="warn-note">' + esc(warn) + '</div>' : ''}</div>` +
      `<div class="acts"><button class="btn small" type="button" data-pv="${i}" data-tip="Посмотреть, что попадёт в файл.">Просмотр</button><button class="btn small primary" type="button" data-dl="${i}" data-tip="Скачать эту форму отдельным файлом Word.">Скачать .docx</button></div></div>`;
  }

  /* ---------- общая информация ---------- */
  function citeBtns(f) {
    if (f.key === 'requirements') return f.count ? `<button class="btn small" type="button" data-cite="requirements" data-tip="Открыть документ с первым найденным разделом требований.">Показать цитату</button><button class="btn small ghost" type="button" data-go="req">К разделу ниже</button>` : '';
    if (f.key === 'forms') {
      const files = Array.from(new Set(f.forms.map(x => x.file)));
      return files.map(n => `<button class="btn small" type="button" data-cite="forms:${esc(n)}" data-tip="Открыть документ: места, где начинаются формы, выделены жёлтым.">Показать цитату${files.length > 1 ? ' — ' + esc(n.length > 28 ? n.slice(0, 26) + '…' : n) : ''}</button>`).join('') + (f.forms.length ? '<button class="btn small ghost" type="button" data-go="forms">К формам ниже</button>' : '');
    }
    return f.hit && f.hit.length && f.src ? `<button class="btn small" type="button" data-cite="${f.key}" data-tip="Открыть документ: найденная информация выделена жёлтым.">Показать цитату</button>` : '';
  }
  function sumValue(f) {
    if (f.key === 'forms') return f.forms.length ? `<div class="v-forms">${f.forms.map((x, i) => '<div>' + esc(A.formFileName(x, i).replace(/\.docx$/, '')) + '</div>').join('')}</div>` : '<span class="v-nf">информация не найдена</span>';
    let h = '';
    if (f.kind === 'value' || f.kind === 'flag') h = `<span class="v-ok">${esc(f.value)}</span>`;
    else if (f.kind === 'none') h = `<span class="v-none">${esc(f.value)}</span>`;
    else h = `<span class="v-nf">${esc(f.value)}</span>`;
    if (f.note && f.kind !== 'value') h += `<div class="v-note">в документации: «${esc(f.note)}»</div>`;
    if (f.flag) h += `<div class="flag-row"><span class="flag">${esc(f.flag.value)}</span><button class="btn small" type="button" data-cite="flag" data-tip="Открыть документ: найденная фраза выделена жёлтым.">Показать цитату</button></div>`;
    return h;
  }
  function renderSum() {
    const s = state.sum, order = ['endDate', 'bid', 'exec', 'warranty', 'advance', 'requirements', 'criteria', 'forms'];
    $('listSum').innerHTML = order.map((k, n) => `<div class="sum-row"><div class="sum-k"><span class="sn">${n + 1}</span>${esc(s[k].title)}</div><div class="sum-v">${sumValue(s[k])}</div><div class="sum-a">${citeBtns(s[k])}</div></div>`).join('');
  }

  function showMissing() { return $('optMissing').checked; }
  function syncTabs() {
    const r = state.res; if (!r) return;
    $('tabMiss').hidden = !showMissing();
    if (!showMissing() && state.tab === 'miss') state.tab = 'sum';
    document.querySelectorAll('.an-tab').forEach(b => b.classList.toggle('on', b.dataset.tab === state.tab));
    document.querySelectorAll('.an-pane').forEach(p => { p.hidden = p.dataset.pane !== state.tab; });
  }
  function setTab(t, scroll) { state.tab = t; syncTabs(); if (scroll) { const tb = $('anTabs'); if (tb) window.scrollTo({ top: Math.max(0, tb.offsetTop - 4), behavior: 'smooth' }); } }

  function render() {
    const r = state.res; $('anRes').hidden = false; $('anTabs').hidden = false;
    const set = (cntId, tcId, listId, arr, fn, empty) => { $(cntId).textContent = arr.length ? 'найдено: ' + arr.length : 'не найдено'; $(cntId).classList.toggle('has', !!arr.length); $(tcId).textContent = arr.length; $(listId).innerHTML = arr.length ? arr.map(fn).join('') : `<div class="an-empty">${empty}</div>`; };
    renderSum();
    set('cntComp', 'tcComp', 'listComp', r.composition, (q, i) => quoteRow(q, 'c', i), 'Раздел о составе заявки не найден. Попробуйте загрузить основной файл документации (информационную карту, извещение).');
    set('cntReq', 'tcReq', 'listReq', r.requirements, (q, i) => quoteRow(q, 'r', i), 'Требования к участнику не найдены. Попробуйте загрузить основной файл документации.');
    set('cntMiss', 'tcMiss', 'listMiss', state.sum.notSet, missRow, 'Пометок «Требование не установлено», «требование отсутствует», «не установлено» в документации не найдено.');
    set('cntForms', 'tcForms', 'listForms', r.forms, formRow, 'Формы не найдены. Формы определяются по словам «Форма №…», «Приложение №…», «Образец», «Конец формы» и по названиям вроде «ЗАЯВКА НА УЧАСТИЕ».');
    $('formBar').hidden = !r.forms.length; $('selAll').checked = state.sel.size === r.forms.length;
    $('dlSel').disabled = $('dlZip').disabled = !state.sel.size;
    syncTabs();
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
  /** Окно просмотра: весь файл целиком, найденное выделено жёлтым, пометки «не установлено» — оранжевым. */
  function bgOf(el) { return (el.style && el.style.backgroundColor) || ''; }
  async function openMarked(title, meta, src, marks, text, trunc) {
    if (src && src.kind === 'docx' && window.docx && window.docx.renderAsync) {
      openModal(title, meta, '<p class="ql">Готовлю просмотр всего файла…</p>', text, true);
      try {
        const bytes = await A.highlightLines(src, marks);
        const body = $('anBody'); body.innerHTML = '';
        const host = document.createElement('div'); body.append(host);
        await window.docx.renderAsync(new Blob([bytes], { type: DOCX_MIME }), host, null, { className: 'docx', inWrapper: true, breakPages: true, useBase64URL: true });
        const all = Array.from(host.querySelectorAll('span, p, td'));
        const hit = all.find(el => /yellow|255,\s*255,\s*0\b/i.test(bgOf(el))) || all.find(el => /255,\s*180,\s*84|ffb454/i.test(bgOf(el)));
        if (hit) { hit.classList.add('an-first-hit'); const go = () => hit.scrollIntoView({ block: 'center', inline: 'nearest' }); go(); setTimeout(go, 250); setTimeout(go, 900); }
        $('anMB').textContent = trunc ? 'Цитата очень длинная: выделена не полностью.' : '';
        return;
      } catch (e) { console.warn('предпросмотр файла не получился', e); }
    }
    if (src && src.items) {
      const kindOf = new Map(); marks.forEach(m => { if (m.kind === 'miss' || !kindOf.has(m.i)) kindOf.set(m.i, m.kind); });
      const rows = src.items.map((x, k) => '<p class="ql' + (kindOf.get(k) === 'hit' ? ' qhit' : kindOf.get(k) === 'miss' ? ' qmiss' : '') + '" data-k="' + k + '">' + esc(x.tbl ? x.rows.map(r => r.join(' | ')).join('\n') : x.s) + '</p>');
      openModal(title, meta, rows.join(''), text, false, src.kind === 'docx' ? '' : 'Показан весь текст файла, извлечённый программой (оформление исходного файла не воспроизводится).');
      const b = $('anBody'), first = b.querySelector('.qhit, .qmiss'); if (first) b.scrollTop = Math.max(0, first.offsetTop - Math.round(b.clientHeight / 3));
      return;
    }
    openModal(title, meta, '<p class="ql">' + esc(text) + '</p>', text, false, '');
  }
  const asHit = arr => (arr || []).map(m => ({ i: m.i, row: m.row, kind: 'hit' }));
  /** Пометки «не установлено», попавшие внутрь найденного раздела q. */
  function missInside(q) {
    return (state.sum ? state.sum.notSet : []).filter(m => m.srcRef === q.srcRef && (q.inTable ? (m.i === q.from && m.row >= q.row && m.row < q.rowTo) : (m.i >= q.from && m.i < q.to)))
      .reduce((a, m) => a.concat(m.miss.map(x => ({ i: x.i, row: x.row, kind: 'miss' }))), []);
  }
  function quoteMarks(q) {
    const hit = [];
    if (q.inTable) for (let r = q.row; r < q.rowTo; r++) hit.push({ i: q.from, row: r, kind: 'hit' });
    else for (let k = q.from; k < q.to; k++) hit.push({ i: k, row: null, kind: 'hit' });
    return hit.concat(missInside(q));
  }
  async function showQuote(kind, i) {
    const q = (kind === 'c' ? state.res.composition : state.res.requirements)[i]; if (!q) return;
    const title = (kind === 'c' ? 'Состав заявки' : 'Требования к участнику') + (q.num ? ' — п. ' + q.num : '');
    const miss = (q.points || []).filter(p => p.notSet).length;
    const meta = 'Файл: ' + q.file + ' · найдено по: «' + q.kw + '» · цитата выделена жёлтым' + (miss ? ', «требование не установлено» — оранжевым' : '');
    return openMarked(title, meta, q.srcRef, quoteMarks(q), q.text, q.truncated);
  }
  async function showMissing1(i) {
    const m = state.sum.notSet[i]; if (!m) return;
    return openMarked('Требование не установлено' + (m.num ? ' — п. ' + m.num : ''), 'Файл: ' + m.file + (m.section ? ' · ' + m.section : '') + ' · пункт — жёлтым, пометка — оранжевым', m.srcRef,
      asHit(m.hit).concat(m.miss.map(x => ({ i: x.i, row: x.row, kind: 'miss' }))), m.text + '\n' + m.marker, false);
  }
  async function showField(key) {
    const s = state.sum;
    if (key === 'requirements') { const q = state.res.requirements[0]; return q ? showQuote('r', 0) : null; }
    if (key.indexOf('forms:') === 0) {
      const name = key.slice(6), fs = state.res.forms.filter(f => f.file === name); if (!fs.length) return;
      return openMarked('Список форм', 'Файл: ' + name + ' · начало каждой формы выделено жёлтым', fs[0].srcRef, fs.map(f => ({ i: f.start, row: null, kind: 'hit' })), fs.map((f, k) => A.formFileName(f, k)).join('\n'), false);
    }
    if (key === 'flag') { const f = s.criteria.flag; if (!f) return; return openMarked(f.value, 'Файл: ' + f.src.name + ' · найденная фраза выделена жёлтым', f.src, asHit(f.hit), f.value, false); }
    const f = s[key]; if (!f || !f.src) return;
    const txt = (f.note ? f.note + '\n' : '') + f.value;
    return openMarked(f.title, 'Файл: ' + f.src.name + ' · найденная информация выделена жёлтым', f.src, asHit(f.hit), txt, false);
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

  async function saveSummary() {
    if (!state.sum) return;
    setStatus('Готовлю документ…');
    try {
      const names = state.files.filter(f => f.status === 'ok').map(f => f.name);
      const bytes = await X.summaryDocx(state.sum, { files: names, formName: (f) => { const k = state.res.forms.indexOf(f); return A.formFileName(f, k).replace(/\.docx$/, ''); }, missing: state.sum.notSet, withMissing: showMissing() });
      const base = (names[0] || 'закупка').replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]/g, ' ').trim().slice(0, 60);
      download(new Blob([bytes], { type: DOCX_MIME }), 'Общая информация — ' + base + '.docx'); setStatus('Список скачан.', 'ok');
    } catch (e) { console.error(e); setStatus('Не удалось собрать документ: ' + (e && e.message || e), 'err'); }
  }

  /* ---------- события ---------- */
  function init() {
    const drop = $('anDrop'), inp = $('anFile');
    inp.addEventListener('change', () => { addFiles(inp.files); inp.value = ''; });
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => addFiles(e.dataTransfer.files));
    $('anFiles').addEventListener('click', e => { const b = e.target.closest('[data-rm]'); if (!b) return; state.files = state.files.filter(f => f.id !== +b.dataset.rm); renderFiles(); run(); });
    $('anClear').addEventListener('click', () => { state.files = []; state.res = null; state.sum = null; state.blobs.clear(); renderFiles(); $('anRes').hidden = true; $('anTabs').hidden = true; setStatus(''); });
    $('optDelims').addEventListener('change', run); $('optTitles').addEventListener('change', run);
    $('optMissing').addEventListener('change', () => { try { localStorage.setItem('assistant.analysis.showMissing', showMissing() ? '1' : '0'); } catch (e) {} syncTabs(); });
    try { $('optMissing').checked = localStorage.getItem('assistant.analysis.showMissing') === '1'; } catch (e) {}
    $('anTabs').addEventListener('click', e => { const b = e.target.closest('.an-tab'); if (b) setTab(b.dataset.tab, true); });
    $('dlSummary').addEventListener('click', saveSummary);
    $('anRes').addEventListener('click', e => {
      const q = e.target.closest('[data-q]'); if (q) { const [k, i] = q.dataset.q.split(':'); return showQuote(k, +i); }
      const mi = e.target.closest('[data-m]'); if (mi) return showMissing1(+mi.dataset.m);
      const ci = e.target.closest('[data-cite]'); if (ci) return showField(ci.dataset.cite);
      const go = e.target.closest('[data-go]'); if (go) return setTab(go.dataset.go, true);
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
