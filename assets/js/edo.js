/* Ассистент — «Штамп ЭДО»: синие отметки электронного документооборота внизу страниц Word-документа.
   Варианты: РТС-тендер (рамка на каждой странице), Тензор (полоса на каждой странице),
   Контур и Диадок (лист подписания в конце документа). Все данные вводит пользователь:
   идентификаторы и номера сертификатов программа сама не придумывает. */
(function () {
  'use strict';
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const VARIANTS = [
    { id: 'rts', title: 'Рамка на каждой странице', sub: 'Как на «РТС-тендер»: «Документ подписан на ЭП», номер договора, «Страница N из M» — справа внизу.', color: '7B7BD0', where: 'footer', op: 'РТС-тендер' },
    { id: 'tensor', title: 'Полоса «Подписано / Согласовано»', sub: 'Как у оператора «Тензор»: полоса на всю ширину внизу каждой страницы, две строки подписантов.', color: '3F3FE0', where: 'footer', op: 'ООО «Компания «Тензор»' },
    { id: 'kontur', title: 'Лист подписания (Контур)', sub: 'Синяя рамка «Документ подписан и передан через оператора ЭДО» с таблицей подписей — внизу последней страницы.', color: '1565C0', where: 'end', op: 'АО «ПФ «СКБ Контур»' },
    { id: 'diadoc', title: 'Лист подписания с доверенностью (Диадок)', sub: 'Тот же лист, но с колонкой «Доверенность» и пометкой тестового режима.', color: '1565C0', where: 'end', op: 'АО «ПФ «СКБ Контур»' }
  ];

  /* ---------- XML-кирпичики ---------- */
  const rpr = (o = {}) => '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' + (o.b ? '<w:b/>' : '') + '<w:color w:val="' + (o.c || '000000') + '"/><w:sz w:val="' + (o.sz || 14) + '"/><w:szCs w:val="' + (o.sz || 14) + '"/></w:rPr>';
  const run = (t, o) => '<w:r>' + rpr(o) + '<w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
  const fld = (code, o) => '<w:r>' + rpr(o) + '<w:fldChar w:fldCharType="begin"/></w:r><w:r>' + rpr(o) + '<w:instrText xml:space="preserve"> ' + code + ' </w:instrText></w:r><w:r>' + rpr(o) + '<w:fldChar w:fldCharType="separate"/></w:r><w:r>' + rpr(o) + '<w:t>1</w:t></w:r><w:r>' + rpr(o) + '<w:fldChar w:fldCharType="end"/></w:r>';
  const para = (inner, o = {}) => '<w:p><w:pPr>' + (o.keep ? '<w:keepNext/>' : '') + '<w:spacing w:before="' + (o.before || 0) + '" w:after="' + (o.after || 0) + '" w:line="240" w:lineRule="auto"/>' + (o.jc ? '<w:jc w:val="' + o.jc + '"/>' : '') + '</w:pPr>' + inner + '</w:p>';
  const border = (side, color, sz) => '<w:' + side + ' w:val="single" w:sz="' + (sz || 6) + '" w:space="0" w:color="' + color + '"/>';
  const cell = (w, inner, o = {}) => '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' + (o.span ? '<w:gridSpan w:val="' + o.span + '"/>' : '') + (o.borders || '') + '<w:tcMar><w:top w:w="' + (o.mt == null ? 30 : o.mt) + '" w:type="dxa"/><w:left w:w="70" w:type="dxa"/><w:bottom w:w="' + (o.mb == null ? 30 : o.mb) + '" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tcMar></w:tcPr>' + inner + '</w:tc>';
  const table = (cols, rows, o = {}) => {
    const tot = cols.reduce((a, b) => a + b, 0), c = o.color || '000000';
    return '<w:tbl><w:tblPr><w:tblW w:w="' + tot + '" w:type="dxa"/>' + (o.jc ? '<w:jc w:val="' + o.jc + '"/>' : '') + '<w:tblBorders>' + (o.noInner ? border('top', c, o.sz) + border('left', c, o.sz) + border('bottom', c, o.sz) + border('right', c, o.sz) : border('top', c, o.sz) + border('left', c, o.sz) + border('bottom', c, o.sz) + border('right', c, o.sz) + border('insideH', c, 4)) + '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' + cols.map(w => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>' + rows.map(r => '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + r + '</w:tr>').join('') + '</w:tbl>';
  };

  /* ---------- сборка вариантов ---------- */
  function build(v, d) {
    const c = v.color, T = (t, o) => run(t, Object.assign({ c }, o));
    const s1 = d.s1 || {}, s2 = d.s2 || {};
    const who = s => [s.org, s.fio, s.pos].filter(Boolean).join(', ');
    const certLine = s => [s.cert ? 'Сертификат ' + s.cert : '', s.valid ? 'действителен ' + s.valid : ''].filter(Boolean).join(', ');
    if (v.id === 'rts') {
      const lines = para(T('Документ подписан на ЭП «' + (d.op || '') + '»', { sz: 13 }), { jc: 'left' }) + (d.docname ? para(T(d.docname, { sz: 13 })) : '') + para(T('Страница ', { sz: 13 }) + fld('PAGE', { c, sz: 13 }) + T(' из ', { sz: 13 }) + fld('NUMPAGES', { c, sz: 13 }));
      return { footer: table([3100], ['' + cell(3100, lines, { mt: 40, mb: 40 })], { jc: 'right', color: c, sz: 8 }) + para('', {}), end: '' };
    }
    if (v.id === 'tensor') {
      const row = (lab, s) => '' + cell(1250, para(T(lab, { sz: 13 })), { mt: 20, mb: 20 }) + cell(4300, para(T(s.org ? s.org + ', ' : '', { b: 1, sz: 13 }) + T(s.fio || '', { sz: 13 })), { mt: 20, mb: 20 }) + cell(1900, para(T(s.date || '', { b: 1, sz: 13 })), { mt: 20, mb: 20 }) + cell(2200, para(T('Сертификат ' + (s.cert || ''), { sz: 13 })), { mt: 20, mb: 20 });
      const head = cell(5550, para(T('ДОКУМЕНТ ПОДПИСАН ЭЛЕКТРОННОЙ ПОДПИСЬЮ', { sz: 14 })), { span: 2 }) + cell(4100, para(T('Оператор ЭДО ' + (d.op || '') + (d.docid ? '   Идентификатор: ' + d.docid : ''), { sz: 12 }), { jc: 'right' }), { span: 2 });
      const rows = [head, row('ПОДПИСАНО', s1), row('СОГЛАСОВАНО', s2)];
      return { footer: table([1240, 4300, 1900, 2190], rows, { color: c, sz: 6 }) + para('', {}), end: '' };
    }
    // Контур / Диадок: лист подписания в конец
    const dia = v.id === 'diadoc', cols = dia ? [1250, 2750, 1650, 2200, 1800] : [1250, 3300, 2600, 2500];
    const H = (t, w) => cell(w, para(T(t, { sz: 13 })), { mt: 60 });
    const ctext = s => para(T(s.cert || '', { sz: 13 })) + (s.valid ? para(T('с ' + s.valid, { sz: 13 })) : '');
    const dtext = s => para(T(s.date || '', { sz: 13 })) + (s.date ? para(T('Подпись соответствует файлу документа', { sz: 13 })) : '');
    const sigRow = (lab, s) => cell(cols[0], para(T(lab, { sz: 13 })), { mt: 80 }) + cell(cols[1], para(T(s.org || '', { sz: 13 })) + para(T([s.fio, s.pos].filter(Boolean).join(', '), { sz: 13 })), { mt: 80 }) + (dia ? cell(cols[2], para(T('Не требуется для подписания', { sz: 13 })), { mt: 80 }) : '') + cell(cols[dia ? 3 : 2], ctext(s), { mt: 80 }) + cell(cols[dia ? 4 : 3], dtext(s), { mt: 80 });
    const title = 'Документ подписан и передан ' + (dia ? 'в тестовом режиме ' : '') + 'через оператора ЭДО ' + (d.op || '');
    const idline = d.docid ? 'Идентификатор документа ' + d.docid : '';
    const top = cell(cols.reduce((a, b) => a + b, 0), para(T(idline, { sz: 12 }), { jc: 'right' }) + para(T(title, { sz: 22 }), { before: 60, after: 120 }), { span: cols.length, mt: 40 });
    const head = H('', cols[0]) + H(dia ? 'Организация, сотрудник' : 'Владелец сертификата: организация, сотрудник', cols[1]) + (dia ? H('Доверенность: рег. номер, период действия и статус', cols[2]) : '') + H('Сертификат: серийный номер, период действия', cols[dia ? 3 : 2]) + H('Дата и время подписания', cols[dia ? 4 : 3]);
    const rows = [top, head, sigRow('Подписи отправителя:', s1), sigRow('Подписи получателя:', s2)];
    return { footer: '', end: para('', { before: 240 }) + table(cols, rows, { color: c, sz: 8 }) + para(T('Страница ', { sz: 13 }) + fld('PAGE', { c, sz: 13 }) + T(' из ', { sz: 13 }) + fld('NUMPAGES', { c, sz: 13 }), { jc: 'right', before: 40 }) };
  }

  /* лист «Сведения о подписях» для варианта РТС (по желанию, в конец) */
  function rtsSheet(d, c) {
    const T = (t, o) => run(t, Object.assign({ c, sz: 13 }, o)), s1 = d.s1 || {}, s2 = d.s2 || {};
    const line = (a, b) => '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cell(2800, para(T(a)), { mt: 10, mb: 10 }) + cell(6000, para(T(b)), { mt: 10, mb: 10 }) + '</w:tr>';
    const blockOf = (t, s) => line(t, '').replace(/<w:t xml:space="preserve">([^<]*)<\/w:t>/, '<w:t xml:space="preserve">$1</w:t>') + line('Дата подписания:', s.date || '') + line('Организация:', s.org || '') + line('ФИО:', s.fio || '') + line('Должность:', s.pos || '') + line('Сертификат ЭП действителен:', s.valid || '') + line('Серийный номер сертификата ЭП:', s.cert || '');
    const tbl = '<w:tbl><w:tblPr><w:tblW w:w="8800" w:type="dxa"/><w:tblBorders>' + border('top', c, 6) + border('left', c, 6) + border('bottom', c, 6) + border('right', c, 6) + '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2800"/><w:gridCol w:w="6000"/></w:tblGrid>' + line('Документ:', d.docname || '') + blockOf('Исполнитель', s1) + blockOf('Заказчик', s2) + '</w:tbl>';
    return para('', { before: 240 }) + tbl;
  }

  /* ---------- вставка в docx ---------- */
  async function apply(buf, variantId, data, opts) {
    opts = opts || {};
    const v = VARIANTS.find(x => x.id === variantId) || VARIANTS[0];
    const zip = await JSZip.loadAsync(buf);
    const docFile = zip.file('word/document.xml'); if (!docFile) throw new Error('Это не документ Word (.docx).');
    let xml = await docFile.async('string');
    const parts = build(v, data);
    let end = parts.end;
    if (v.id === 'rts' && opts.sheet) end += rtsSheet(data, v.color);
    // на «последней странице» вместо колонтитула
    if (opts.lastOnly && parts.footer) { end = parts.footer + end; parts.footer = ''; }
    if (end) {
      let i = xml.lastIndexOf('<w:sectPr'), j = xml.lastIndexOf('</w:body>');
      if (i < 0 || i < xml.lastIndexOf('</w:p>')) i = j;
      if (i < 0) throw new Error('Не удалось найти конец документа.');
      xml = xml.slice(0, i) + end + xml.slice(i);
    }
    if (parts.footer) {
      let rels = await zip.file('word/_rels/document.xml.rels').async('string');
      let ct = await zip.file('[Content_Types].xml').async('string');
      const done = {}; let newRid = null, relsChanged = false;
      const addToPart = async target => {
        const path = 'word/' + target.replace(/^\/?(word\/)?/, '');
        if (done[path]) return; done[path] = 1;
        const f = zip.file(path); if (!f) return;
        let fx = await f.async('string');
        const k = fx.lastIndexOf('</w:ftr>'); if (k < 0) return;
        zip.file(path, fx.slice(0, k) + parts.footer + fx.slice(k));
      };
      const relTarget = rid => { const m = rels.match(new RegExp('<Relationship[^>]*Id="' + rid + '"[^>]*>')); if (!m) return null; const t = m[0].match(/Target="([^"]+)"/); return t ? t[1] : null; };
      const sect = /<w:sectPr\b[^>]*?(?:\/>|>[\s\S]*?<\/w:sectPr>)/g;
      const cnt = (xml.match(sect) || []).length;
      const out = []; let last = 0, m, idx = 0;
      while ((m = sect.exec(xml))) {
        let s = m[0]; idx++;
        if (s.endsWith('/>')) s = s.replace(/\/>$/, '></w:sectPr>');
        const refs = s.match(/<w:footerReference\b[^>]*>/g) || [];
        let hasDefault = false;
        for (const r of refs) {
          const rid = (r.match(/r:id="([^"]+)"/) || [])[1], ty = (r.match(/w:type="([^"]+)"/) || [])[1];
          if (ty === 'default') hasDefault = true;
          const tg = rid && relTarget(rid); if (tg) await addToPart(tg);
        }
        if (!hasDefault) {
          if (!newRid) { newRid = 'rIdEdo1'; relsChanged = true; }
          s = s.replace(/^(<w:sectPr\b[^>]*>)/, '$1<w:footerReference w:type="default" r:id="' + newRid + '"/>');
        }
        out.push(xml.slice(last, m.index), s); last = m.index + m[0].length;
      }
      out.push(xml.slice(last)); xml = out.join('');
      if (!cnt) throw new Error('В документе не найдены параметры страницы.');
      if (newRid) {
        const ftr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '">' + parts.footer + '</w:ftr>';
        zip.file('word/footerEdo1.xml', ftr);
        rels = rels.replace('</Relationships>', '<Relationship Id="' + newRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerEdo1.xml"/></Relationships>');
        zip.file('word/_rels/document.xml.rels', rels);
        if (ct.indexOf('/word/footerEdo1.xml') < 0) ct = ct.replace('</Types>', '<Override PartName="/word/footerEdo1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
        zip.file('[Content_Types].xml', ct);
      }
    }
    zip.file('word/document.xml', xml);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  /* ---------- интерфейс ---------- */
  function mount() {
    const host = document.querySelector('#blk-edo .blk-body'); if (!host) return;
    const F = [['op', 'Оператор / площадка'], ['docid', 'Идентификатор документа'], ['docname', 'Строка о документе (например: Договор №…)']];
    const S = [['org', 'Организация'], ['fio', 'ФИО сотрудника'], ['pos', 'Должность'], ['date', 'Дата и время подписания'], ['cert', 'Серийный номер сертификата'], ['valid', 'Срок действия сертификата (с … по …)']];
    host.innerHTML = '<div class="edo">' +
      '<div class="guide"><div><b>1. Документ</b>Загрузите договор в формате Word (.docx).</div><div><b>2. Вариант штампа</b>Синий штамп встанет внизу каждой страницы или листом подписания в конце документа.</div><div><b>3. Данные</b>Введите сведения о подписях и скачайте результат.</div></div>' +
      '<div class="drop edo-drop" tabindex="0"><b>Перетащите договор сюда</b> или нажмите, чтобы выбрать<small>.docx</small><input type="file" accept=".docx" hidden></div>' +
      '<div class="edo-file"></div>' +
      '<div class="tpl-head">Вариант штампа</div><div class="edo-vars"></div>' +
      '<div class="edo-grid"><div class="edo-col edo-common"></div><div class="edo-col edo-s1"><h4>Подписант 1 (отправитель)</h4></div><div class="edo-col edo-s2"><h4>Подписант 2 (получатель)</h4></div></div>' +
      '<div class="edo-opts"><label class="check"><input type="checkbox" class="edo-last"> Поставить только на последней странице, а не на каждой</label><label class="check edo-sheet-l"><input type="checkbox" class="edo-sheet"> Добавить в конец сведения о подписях (для рамки на каждой странице)</label></div>' +
      '<div class="edo-prev-h">Так будет выглядеть штамп</div><div class="edo-prev"></div>' +
      '<div class="row edo-actions"><button type="button" class="btn small ghost edo-demo">Подставить образец данных</button><button type="button" class="btn primary edo-run" disabled>Поставить штамп и скачать</button></div>' +
      '<div class="edo-msg" role="status"></div>' +
      '<div class="edo-note">Штамп — только графическая отметка в документе. Юридической силы подписи он не имеет, поэтому программа сама не придумывает идентификаторы и номера сертификатов: впишите их из настоящего документа ЭДО.</div></div>';
    const q = s => host.querySelector(s), st = { buf: null, name: '', v: 'rts', d: { op: VARIANTS[0].op, s1: {}, s2: {} } };
    const inp = {};
    const mk = (box, arr, bag, prefix) => arr.forEach(([k, lab]) => { const l = document.createElement('label'); l.innerHTML = '<span>' + lab + '</span><input type="text">'; const i = l.querySelector('input'); inp[prefix + k] = i; if (prefix === '' && st.d[k]) i.value = st.d[k]; i.addEventListener('input', () => { bag()[k] = i.value; prev(); }); box.append(l); });
    mk(q('.edo-common'), F, () => st.d, ''); mk(q('.edo-s1'), S, () => st.d.s1, 's1.'); mk(q('.edo-s2'), S, () => st.d.s2, 's2.');
    const vars = q('.edo-vars');
    VARIANTS.forEach(v => { const b = document.createElement('button'); b.type = 'button'; b.className = 'edo-var'; b.dataset.v = v.id; b.innerHTML = '<i style="border-color:#' + v.color + '"></i><b>' + v.title + '</b><small>' + v.sub + '</small>'; b.addEventListener('click', () => { st.v = v.id; if (!inp.op.value || VARIANTS.some(x => x.op === inp.op.value)) { inp.op.value = v.op; st.d.op = v.op; } sync(); }); vars.append(b); });
    function sync() {
      vars.querySelectorAll('.edo-var').forEach(b => b.classList.toggle('on', b.dataset.v === st.v));
      const v = VARIANTS.find(x => x.id === st.v), foot = v.where === 'footer';
      q('.edo-last').closest('label').style.display = foot ? '' : 'none'; q('.edo-sheet-l').style.display = st.v === 'rts' ? '' : 'none';
      q('.edo-run').disabled = !st.buf; prev();
    }
    function prev() {
      const v = VARIANTS.find(x => x.id === st.v), d = st.d, s1 = d.s1 || {}, s2 = d.s2 || {}, c = '#' + v.color, e = esc;
      let h;
      if (v.id === 'rts') h = '<div class="pv pv-r" style="border-color:' + c + ';color:' + c + '">Документ подписан на ЭП «' + e(d.op) + '»<br>' + e(d.docname) + '<br>Страница 1 из N</div>';
      else if (v.id === 'tensor') h = '<div class="pv pv-t" style="border-color:' + c + ';color:' + c + '"><div class="pv-th"><span>ДОКУМЕНТ ПОДПИСАН ЭЛЕКТРОННОЙ ПОДПИСЬЮ</span><span>Оператор ЭДО ' + e(d.op) + (d.docid ? ' Идентификатор: ' + e(d.docid) : '') + '</span></div>' + [['ПОДПИСАНО', s1], ['СОГЛАСОВАНО', s2]].map(([l, s]) => '<div class="pv-tr"><span>' + l + '</span><span><b>' + e(s.org || '') + '</b> ' + e(s.fio || '') + '</span><span><b>' + e(s.date || '') + '</b></span><span>Сертификат ' + e(s.cert || '') + '</span></div>').join('') + '</div>';
      else { const dia = v.id === 'diadoc'; h = '<div class="pv pv-k" style="border-color:' + c + ';color:' + c + '"><div class="pv-id">' + (d.docid ? 'Идентификатор документа ' + e(d.docid) : '&nbsp;') + '</div><div class="pv-title">Документ подписан и передан ' + (dia ? 'в тестовом режиме ' : '') + 'через оператора ЭДО ' + e(d.op) + '</div>' + [['Подписи отправителя:', s1], ['Подписи получателя:', s2]].map(([l, s]) => '<div class="pv-kr"><span>' + l + '</span><span>' + e(s.org || '') + '<br>' + e([s.fio, s.pos].filter(Boolean).join(', ')) + '</span>' + (dia ? '<span>Не требуется для подписания</span>' : '') + '<span>' + e(s.cert || '') + '<br>' + (s.valid ? 'с ' + e(s.valid) : '') + '</span><span>' + e(s.date || '') + '</span></div>').join('') + '</div>'; }
      q('.edo-prev').innerHTML = h;
    }
    const fileUp = q('.edo-drop input'), drop = q('.edo-drop');
    const take = async f => { if (!f) return; if (!/\.docx$/i.test(f.name)) { q('.edo-msg').textContent = 'Нужен файл Word (.docx). Старый .doc сначала сохраните как .docx.'; return; } st.buf = await f.arrayBuffer(); st.name = f.name; q('.edo-file').textContent = 'Документ: ' + f.name; q('.edo-msg').textContent = ''; sync(); };
    fileUp.addEventListener('change', () => take(fileUp.files[0])); drop.addEventListener('click', e => { if (e.target !== fileUp) fileUp.click(); });
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); }); drop.addEventListener('dragleave', () => drop.classList.remove('over'));
    drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); take(e.dataTransfer.files[0]); });
    q('.edo-demo').addEventListener('click', () => {
      const demo = { docname: 'Договор № 123/45', s1: { org: 'ООО «Тестовая организация»', fio: 'Иванов Иван Иванович, директор', pos: '', date: '01.01.2026 10:00 GMT+03:00', cert: '', valid: '' }, s2: { org: 'ООО «Образец»', fio: 'Петров Пётр Петрович, директор', pos: '', date: '02.01.2026 11:00 GMT+03:00', cert: '', valid: '' } };
      st.d.docname = demo.docname; inp.docname.value = demo.docname;
      ['s1', 's2'].forEach(k => Object.keys(demo[k]).forEach(f => { st.d[k][f] = demo[k][f]; inp[k + '.' + f].value = demo[k][f]; })); prev();
    });
    q('.edo-run').addEventListener('click', async () => {
      const msg = q('.edo-msg'); msg.textContent = 'Готовлю документ…';
      try {
        const out = await apply(st.buf, st.v, st.d, { lastOnly: q('.edo-last').checked, sheet: q('.edo-sheet').checked });
        const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })); a.download = st.name.replace(/\.docx$/i, '') + ' (ЭДО).docx'; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
        msg.textContent = 'Готово: ' + a.download + '. Открывайте в Word — номера страниц «N из M» обновятся сами.';
      } catch (e) { msg.textContent = 'Не получилось: ' + (e.message || e); }
    });
    sync();
  }
  window.AssistantEdo = { apply, build, VARIANTS };
  if (typeof document !== 'undefined') { if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount); else mount(); }
})();
