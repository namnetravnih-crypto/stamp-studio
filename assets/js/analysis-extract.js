/* =====================================================================
   Ассистент — «Анализ документации»: общая сводка и пометки «требование не установлено».
   1) summarize(): дата окончания подачи заявок, обеспечение заявки / исполнения / гарантийных обязательств,
      аванс, критерии оценки, «реестровый опыт», список форм — у каждого поля есть место в документе (для цитаты);
   2) findNotSet(): все места, где в закупке стоит «Требование не установлено» / «требование отсутствует» / «не установлено»;
   3) pointsOf(): пункты внутри найденного раздела (номер, текст, отметка «не установлено»);
   4) summaryDocx(): сводка одним документом Word.
   Работает и в браузере, и в Node (для проверки).
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./analysis-core.js'), require('jszip'));
  else root.AnalysisExtract = factory(root.AnalysisCore, root.JSZip);
})(typeof self !== 'undefined' ? self : this, function (A, JSZip) {
  'use strict';
  const nl = A.nl, norm = A.normText;

  /* ---------- ФРАЗЫ ДЛЯ ПОИСКА «только реестровый опыт» ----------
     Ищется ПОЛНАЯ фраза целиком (регистр, кавычки, пробелы, «№» / «N» не важны), а не отдельные слова.
     Чтобы добавить фразу — допишите ещё одну строку в список (звёздочки по краям можно оставить — они отбрасываются). */
  const REGISTRY_FLAG = 'Внимание! Принимается только реестровый опыт';
  const REGISTRY_PHRASES = [
    '*Установлено положение о принятии к оценке исключительно контракта, заключенного и исполненного в соответствии с Федеральным законом от 5 апреля 2013 г. N 44-ФЗ "О контрактной системе в сфере закупок товаров, работ, услуг для обеспечения государственных и муниципальных нужд", и договора, заключенного и исполненного в соответствии с Федеральным законом от 18 июля 2011 г. N 223-ФЗ "О закупках товаров, работ, услуг отдельными видами юридических лиц"*',
    '*к оценке принимаются исключительно контракты, заключенные и исполненные в соответствии с Законом № 44-ФЗ, и договоры, заключенные и исполненные в соответствии с Федеральным законом от 18.07.2011 № 223-ФЗ «О закупках товаров, работ, услуг отдельными видами юридических лиц»*'
  ];
  /** Ключ для сравнения фраз: только буквы и цифры, «№» = «N», ё = е, без регистра. */
  const keyOf = s => String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/№/g, 'n').replace(/[^a-zа-я0-9]+/g, '');
  const stripStars = s => String(s).trim().replace(/^\*+|\*+$/g, '').trim();

  /* ---------- метка «не установлено» ---------- */
  const MARK_RE = /(?:требовани[а-я]*\s+(?:не\s+установлен[а-я]*|отсутству[а-я]*)|не\s+установлен[а-я]*)/;
  /** Короткая пометка: вся ячейка / абзац — это «Требование не установлено» (допускается пара слов рядом). */
  function isMarkText(t) {
    const n = nl(t); if (!n || n.length > 90) return false;
    const m = MARK_RE.exec(n); if (!m) return false;
    const rest = (n.slice(0, m.index) + ' ' + n.slice(m.index + m[0].length)).replace(/[^a-zа-я0-9]/g, '');
    return rest.length <= 20;
  }
  /** Пометка внутри фразы: «Требование о наличии СРО: не установлено.», «…статьи 31: требование не установлено». */
  function longMark(t) {
    const n = nl(t); if (!n || n.length > 420) return '';
    const m = MARK_RE.exec(n); if (!m) return '';
    const tail = n.slice(m.index + m[0].length).replace(/[^a-zа-я0-9]/g, '');
    if (tail.length > 12 && !/требовани|допуск|квалификац|членств|лиценз|\bсро\b|опыт|деловой\s+репутац/.test(n)) return '';
    return m[0];
  }
  function markText(line) {
    if (line.tbl) {
      for (const c of line.cells) if (c && isMarkText(c)) return { mark: norm(c), self: false };
      for (const c of line.cells) { const lm = c && !splitEnum(c).length && longMark(c); if (lm) return { mark: lm, self: true }; }
      return { mark: '', self: false };
    }
    if (isMarkText(line.text)) return { mark: norm(line.text), self: false };
    const lm = !splitEnum(line.text).length && longMark(line.text); return lm ? { mark: lm, self: true } : { mark: '', self: false };
  }

  /* ---------- плоские строки документа ---------- */
  function buildLines(src) {
    if (src._lines) return src._lines;
    const L = [];
    src.items.forEach((it, i) => {
      if (it.toc) return;
      if (it.tbl) {
        it.rows.forEach((cells, r) => {
          const ri = A.rowInfo(cells); const text = ((ri.num ? ri.num + ' ' : '') + ri.rest.join(' | ')).trim(); if (!text) return;
          L.push({ i, row: r, tbl: true, cells, num: ri.num, depth: ri.depth, head: ri.heading, label: ri.label, value: ri.value, text, n: nl(text), k: L.length });
        });
      } else if (it.s) {
        const t = A.textNumber(it.s);
        L.push({ i, row: null, tbl: false, cells: [it.s], num: t.num, depth: t.depth, head: (!t.num && it.hl > 0 && it.hl <= 2) || (!!t.num && A.capsShort(t.body)) || /^(?:раздел|глава|часть)\s+(?:[ivxlc]+|\d+)\b/i.test(it.s), label: t.body, value: '', text: it.s, n: it.n, k: L.length });
      }
    });
    L.forEach(l => { const mt = markText(l); l.mark = mt.mark; l.self = mt.self; });
    src._lines = L; return L;
  }
  const mk = l => ({ i: l.i, row: l.row });

  /** Окно после строки k: продолжение таблицы (строки без номера и вложенные) или несколько абзацев. */
  function windowOf(lines, k, max) {
    const me = lines[k], out = [me]; max = max || 8;
    for (let j = k + 1; j < lines.length && out.length < max; j++) {
      const x = lines[j];
      if (x.head) break;
      if (me.tbl) {
        if (!x.tbl || x.i !== me.i) break;
        if (x.depth && me.depth && x.depth <= me.depth) break;
        if (x.depth && !me.depth) break;
      } else {
        if (x.depth && x.depth <= (me.depth || 1)) break;
        if (/^(?:размер|обеспечени|аванс|критери|срок|порядок|дата|место|форма)[а-я]*\b/.test(x.n)) break;
        if (out.length >= 4) break;
      }
      out.push(x);
    }
    return out;
  }

  /* ---------- суммы и проценты ---------- */
  const PCT_RE = /\d+(?:[.,]\d+)?\s*(?:%|процент)/i;
  const MONEY_RE = /\d[\d\s ]*(?:[.,]\d+)?\s*(?:руб|₽|р\.)/i;
  const hasAmount = t => PCT_RE.test(t) || MONEY_RE.test(t);
  const sentences = t => norm(t).split(/(?<=[.;!?])\s+(?=[^\s])/).map(x => x.trim()).filter(Boolean);
  /** Предложение, где встретилась сумма или процент (с учётом, что в таблице подпись и значение — разные ячейки). */
  function amountSentence(win, must) {
    for (const l of win) {
      const text = l.tbl && l.label && l.value && l === win[0] ? l.value : l.text;
      for (const s of text.split(' | ').reduce((a, part) => a.concat(sentences(part)), [])) if (hasAmount(s) && (!must || must.test(nl(s)) || (text.indexOf(' | ') >= 0 && must.test(nl(text))))) return { text: clamp(s), line: l };
    }
    return null;
  }
  const trunc = (s, n) => { s = norm(s).replace(/^\d+(?:\.\d+)*[.)]\s+/, ''); n = n || 240; return s.length > n ? s.slice(0, n - 1).replace(/\s+\S*$/, '') + '…' : s; };
  const clamp = (s, n) => { s = norm(s).replace(/^\d+[.)]\s+/, '').replace(/^(?:размер|сумма|величина|авансирование|аванс)[^:]{0,80}:\s*(?=\S)/i, ''); n = n || 240; return s.length > n ? s.slice(0, n - 1).replace(/\s+\S*$/, '') + '…' : s; };

  const NONE_VAL = /^(?:не\s+(?:требуется|установлен[а-я]*|предусмотрен[а-я]*|применя[а-я]*|вносится)|отсутству[а-я]*|нет|-|—|–)[\s.;,]*$/;
  function noneIn(win, near) {
    for (const l of win) { if (l.mark) return { text: l.mark, line: l }; }
    const first = win[0];
    if (first.tbl && first.value && first.value.split(' | ').some(v => NONE_VAL.test(nl(v)))) return { text: clamp(first.value), line: first };
    for (const l of win) {
      for (const s of sentences(l.tbl && l === first && l.value ? l.value : l.text)) {
        const n = nl(s);
        if (near.test(n)) return { text: clamp(s), line: l };
      }
    }
    return null;
  }
  const NEG = '(?:не\\s+(?:требуется|установлен[а-я]*|предусмотрен[а-я]*|вносится|применя[а-я]*|осуществля[а-я]*|выплачива[а-я]*|производ[а-я]*)|отсутству[а-я]*|без\\s+обеспечени[а-я]*)';

  /* ---------- обеспечение ---------- */
  const SEC = {
    bid: {
      title: 'Обеспечение заявки',
      size: /размер\s+обеспечени[а-я]*\s+(?:заявк[а-я]*|заявок|участия)/,
      any: /обеспечени[а-я]*\s+(?:заявк[а-я]*|заявок|участия\s+в)|(?:заявк[а-я]*|заявок)\s+(?:на\s+участие\s+в\s+[а-я]+\s+)?обеспечени/,
      excl: null
    },
    exec: {
      title: 'Обеспечение исполнения контракта',
      size: /размер\s+обеспечени[а-я]*\s+исполнени[а-я]*\s+(?:контракт|договор)/,
      any: /обеспечени[а-я]*\s+исполнени[а-я]*\s+(?:контракт[а-я]*|договор[а-я]*|обязательств)/,
      excl: /гарантийн/
    },
    warranty: {
      title: 'Обеспечение гарантийных обязательств',
      size: /размер\s+обеспечени[а-я]*\s+гарантийн/,
      any: /обеспечени[а-я]*\s+(?:исполнени[а-я]*\s+)?гарантийн[а-я]*\s+обязательств|гарантийн[а-я]*\s+обязательств[а-я]*[^.;]{0,60}обеспечени/,
      excl: null
    }
  };
  const NEAR_OBESP = new RegExp('обеспечени[а-я]*[^.;]{0,90}?' + NEG + '|' + NEG + '[^.;]{0,60}?обеспечени');

  function evalSecurity(spec, src) {
    const lines = buildLines(src); const cands = [];
    lines.forEach((l, k) => {
      const lab = nl(l.tbl ? l.label : l.text);
      if (spec.excl && spec.excl.test(lab) && !spec.size.test(lab)) return;
      if (spec.size.test(lab)) cands.push({ k, p: 0 });
      else if (spec.any.test(lab) && (l.tbl ? true : true) && !(l.tbl && /^(?:порядок|реквизит|форма|срок)/.test(lab))) cands.push({ k, p: 1 });
    });
    cands.sort((a, b) => a.p - b.p || a.k - b.k);
    let none = null;
    for (const c of cands) {
      const win = windowOf(lines, c.k, 8);
      const own = noneIn([win[0]], NEAR_OBESP);
      if (own && !amountSentence([win[0]], null)) { if (!none) none = { kind: 'none', value: 'не установлено / отсутствует', note: own.text, hit: [mk(win[0])], src }; if (c.p === 0) break; continue; }
      const v = amountSentence(win, c.p === 0 ? null : /размер|составля|в\s+размере|сумм|не\s+менее|не\s+более|процент|%/);
      if (v) return { kind: 'value', value: v.text, hit: win.slice(0, win.indexOf(v.line) + 1).map(mk), src };
      if (!none && (c.p === 0 || windowOf(lines, c.k, 3).length)) {
        const nn = noneIn(win, NEAR_OBESP);
        if (nn) none = { kind: 'none', value: 'не установлено / отсутствует', note: nn.text, hit: win.slice(0, win.indexOf(nn.line) + 1).map(mk), src };
      }
      if (c.p === 0 && none) break;
    }
    return none;
  }

  /* ---------- аванс ---------- */
  function evalAdvance(src) {
    const lines = buildLines(src); let note = null;
    for (let k = 0; k < lines.length; k++) {
      const l = lines[k]; if (!/аванс/.test(nl(l.text))) continue;
      const win = windowOf(lines, k, 5);
      const v = amountSentence(win, /аванс/);
      if (v) return { kind: 'value', value: v.text, hit: win.slice(0, win.indexOf(v.line) + 1).map(mk), src };
      if (!note) {
        const nn = noneIn(win, new RegExp('аванс[а-я]*[^.;]{0,60}?' + NEG + '|' + NEG + '[^.;]{0,40}?аванс'));
        if (nn && /аванс/.test(nl(nn.line.text))) note = { kind: 'notfound', value: 'информация не найдена', note: nn.text, hit: [mk(nn.line)], src };
      }
    }
    return note;
  }

  /* ---------- дата окончания подачи заявок ---------- */
  const MONTHS = { января: 1, февраля: 2, марта: 3, апреля: 4, мая: 5, июня: 6, июля: 7, августа: 8, сентября: 9, октября: 10, ноября: 11, декабря: 12 };
  const DATE_RES = [
    /(\d{1,2})\.(\d{1,2})\.(\d{4})/,
    /[«"„]?\s*(\d{1,2})\s*[»"“]?\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})/i
  ];
  function allDates(t) {
    const out = [];
    for (const re of [new RegExp(DATE_RES[0].source, 'g'), new RegExp(DATE_RES[1].source, 'gi')]) {
      let m; while ((m = re.exec(t))) {
        const mode = re === undefined ? 1 : (m[2] && /^\d+$/.test(m[2]) ? 1 : 2);
        const dd = String(+m[1]).padStart(2, '0'), mm = String(mode === 1 ? +m[2] : MONTHS[m[2].toLowerCase()]).padStart(2, '0');
        out.push({ date: dd + '.' + mm + '.' + m[3], idx: m.index, len: m[0].length });
      }
    }
    return out.sort((x, y) => x.idx - y.idx);
  }
  /** Из текста выбираем именно дату окончания: перед ней «окончания / до», а не «начала». */
  function findDate(t) {
    const ds = allDates(t); if (!ds.length) return null;
    let prevEnd = 0, pick = null, fallback = null;
    for (const d of ds) {
      const ctx = nl(t.slice(Math.max(prevEnd, d.idx - 40), d.idx)); prevEnd = d.idx + d.len;
      if (/окончани|заканчива|завершени|прекраща|(?:^|\s)до\s*$|(?:^|\s)по\s*$/.test(ctx)) { pick = pick || d; }
      else if (!/начал|с\s*$|открыт/.test(ctx) && !fallback) fallback = d;
    }
    return pick || fallback;
  }
  const TIME_RE = /(?:^|[^\d.:])(\d{1,2}):(\d{2})(?!\d)(\s*\((?:время\s+)?[^)]{2,30}\))?(\s*(?:мск|по\s+московскому\s+времени|по\s+местному\s+времени))?/i;
  const END_LABEL = /(?:дата(?:\s+и\s+время)?\s+окончани[а-я]*\s+(?:срока\s+)?(?:подачи|приема|приёма)|окончани[а-я]*\s+(?:срока\s+)?(?:подачи|приема|приёма)\s+(?:заявок|заявки|предложений|оферт)|(?:прием|приём|подача)\s+(?:заявок|предложений)\s+(?:заканчивается|оканчивается|прекращается|завершается)|срок\s+окончания\s+(?:подачи|приема|приёма)|заявки\s+принимаются\s+до|принимаются\s+до|дата\s+окончания\s+(?:подачи|приема|приёма)|срок\s+(?:и\s+порядок\s+)?(?:подачи|приема|приёма)\s+(?:заявок|заявки|предложений|оферт)|(?:крайний|предельный)\s+срок\s+(?:подачи|приема|приёма)|дата\s+(?:и\s+время\s+)?(?:подачи|приема|приёма)\s+(?:заявок|заявки)|заявки\s+принимаются|подача\s+заявок)/;
  function evalEndDate(src) {
    const lines = buildLines(src);
    for (let k = 0; k < lines.length; k++) {
      const l = lines[k]; const lab = nl(l.tbl ? l.label : l.text);
      if (!END_LABEL.test(lab)) continue;
      const win = windowOf(lines, k, 6); let date = null, time = '', used = [];
      for (const w of win) {
        const seg = (w.tbl && w === win[0] && w.value) ? w.value : (w === win[0] && !w.tbl ? w.text : w.text);
        const n = nl(seg);
        const d = findDate(seg); if (!d) continue;
        if (/начал/.test(n) && !/окончани|заканчива|до\s/.test(n)) continue;
        date = d; used.push(w);
        const t = TIME_RE.exec(seg.slice(d.idx + d.len)) || TIME_RE.exec(seg.slice(0, d.idx));
        if (t) time = t[1].padStart(2, '0') + ':' + t[2] + (t[3] ? ' ' + norm(t[3]) : '') + (t[4] ? ' ' + norm(t[4]) : '');
        break;
      }
      if (!date) continue;
      if (!time) for (const w of win) {
        if (!/врем|час|мск/.test(nl(w.text))) continue; const tm = TIME_RE.exec(w.text);
        if (tm) { time = tm[1].padStart(2, '0') + ':' + tm[2] + (tm[3] ? ' ' + norm(tm[3]) : '') + (tm[4] ? ' ' + norm(tm[4]) : ''); used.push(w); break; }
      }
      return { kind: 'value', value: date.date + (time ? ' ' + time : ''), hit: used.map(mk), src };
    }
    return null;
  }

  /* ---------- критерии оценки ---------- */
  const CRIT_LABEL = /критери[а-я]*\s+(?:и\s+порядок\s+)?оценк|порядок\s+оценки|значимост[а-я]*\s+критери|оценк[а-я]*\s+заявок\s+(?:по|на\s+основании)\s+критери/;
  const CRIT_NONE = /(?:критери[а-я]*|оценк[а-я]*)[^.;]{0,80}?(?:не\s+(?:предусмотрен[а-я]*|установлен[а-я]*|применя[а-я]*|проводится|используются)|отсутству[а-я]*)|(?:не\s+(?:предусмотрен[а-я]*|установлен[а-я]*|применя[а-я]*)|отсутству[а-я]*)[^.;]{0,50}?критери/;
  const WEIGHT = '(?:значимост[а-я]*|вес[а-я]*|доля|удельн[а-я]*\\s+вес|весов[а-я]*\\s+коэффициент[а-я]*|коэффициент[а-я]*\\s+значимости)';
  const cleanName = s => norm(s).replace(/^(?:\d+(?:\.\d+)*[.)]?|[а-я]\))\s+/i, '').replace(/^[«"(\-–—:,;.\s]+|[»"\-–—:,;.\s]+$/g, '').replace(/\)$/, m => /\(/.test(s) ? m : '').replace(/\s*(?:значимость|вес|доля)\s*(?:критерия)?\s*$/i, '').trim();
  const pctStr = s => s.replace(',', '.').replace(/\.0+$/, '').replace(/(\.\d*?)0+$/, '$1').replace('.', ',');
  function critEntries(win) {
    const out = [], seen = new Set();
    const add = (name, pct, line) => {
      name = cleanName(name); if (!name || name.length < 3 || name.length > 130 || /^(?:итого|всего|сумма)\b/i.test(name)) return;
      if (/критерий|значимость|наименование|№|показатель/i.test(name) && name.length < 25 && !/цен|опыт|квалиф|срок|качеств/i.test(name)) return;
      const key = nl(name); if (seen.has(key)) return; seen.add(key);
      out.push({ name, pct: pctStr(pct), line });
    };
    for (const l of win) {
      if (l.tbl) {
        const cells = l.cells.filter((c, i, a) => c && a.indexOf(c) === i);
        const pc = cells.find(c => /^\d+(?:[.,]\d+)?\s*%?$/.test(c.trim()) && !/^\d+[.)]?$/.test(c.trim()) || /^\d+(?:[.,]\d+)?\s*%$/.test(c.trim()));
        if (pc) {
          const names = cells.filter(c => c !== pc && !/^\d+(?:\.\d+)*[.)]?$/.test(c.trim())).sort((a, b) => b.length - a.length);
          if (names[0]) { add(names[0], pc.replace(/[^\d.,]/g, ''), l); continue; }
        }
      }
      const segs = l.tbl ? l.cells.filter((c, i, a) => c && a.indexOf(c) === i && !A.numDepth(c)) : [l.text];
      for (const s of segs.reduce((a, seg) => a.concat(sentences(seg)), [])) {
        let found = 0, m, re;
        re = /[«"“]([^»"”]{3,130})[»"”][^\d%«»]{0,40}?(\d+(?:[.,]\d+)?)\s*(?:%|процент)/g;
        while ((m = re.exec(s))) { if (!/^(?:критерий|значимость)$/i.test(m[1].trim())) { add(m[1], m[2], l); found++; } }
        if (found) continue;
        re = new RegExp('([^:;%]{4,130}?)\\s*[,(]?\\s*' + WEIGHT + '[^\\d%]{0,30}?(\\d+(?:[.,]\\d+)?)\\s*(?:%|процент)', 'ig');
        while ((m = re.exec(s))) { add(m[1], m[2], l); found++; }
        if (found) continue;
        re = /([А-ЯЁA-Zа-яёa-z«»"()][^:;–—%]{2,120}?)\s*[–—:\-]\s*(\d+(?:[.,]\d+)?)\s*(?:%|процент)/g;
        while ((m = re.exec(s))) { add(m[1], m[2], l); found++; }
        if (found) continue;
        re = /(\d+(?:[.,]\d+)?)\s*%\s*[–—:\-]\s*([^;%]{4,120}?)(?=\s+\d+(?:[.,]\d+)?\s*%|;|$)/g;
        while ((m = re.exec(s))) { add(m[2], m[1], l); }
      }
    }
    return out;
  }
  function evalCriteria(src) {
    const lines = buildLines(src); let none = null, best = null;
    for (let k = 0; k < lines.length; k++) {
      const l = lines[k]; const lab = nl(l.tbl ? l.label : l.text);
      if (!CRIT_LABEL.test(lab) || (l.text.length > 260 && !l.tbl && !l.head)) continue;
      let win = windowOf(lines, k, l.head ? 60 : 30);
      if (!l.tbl) { win = [l]; for (let j = k + 1; j < lines.length && win.length < 45; j++) { if (lines[j].head && !CRIT_LABEL.test(lines[j].n)) break; win.push(lines[j]); } }
      const en = critEntries(win);
      if (en.length && (!best || en.length > best.en.length)) best = { en, win, k };
      if (!none) { const nn = noneIn(win.slice(0, 4), CRIT_NONE); if (nn && CRIT_NONE.test(nl(nn.line.text)) || (nn && nn.line.mark)) none = { kind: 'none', value: 'не установлено / отсутствует', note: nn.text, hit: [mk(nn.line)], src }; }
    }
    if (best) return { kind: 'value', value: best.en.map(e => e.pct + '% – ' + e.name).join(', '), hit: best.en.map(e => mk(e.line)), src };
    return none;
  }

  /* ---------- реестровый опыт (поиск полной фразы) ---------- */
  function findPhrases(src, phrases) {
    const lines = buildLines(src); if (!lines.length) return null;
    // склеиваем ключи строк, помним границы
    let big = '', spans = [];
    lines.forEach(l => { const kk = keyOf(l.text); spans.push([big.length, big.length + kk.length]); big += kk; });
    const hits = [];
    for (const ph of phrases) {
      const key = keyOf(stripStars(ph)); if (key.length < 20) continue;
      let from = 0, at;
      while ((at = big.indexOf(key, from)) >= 0) {
        const endAt = at + key.length; const hit = [];
        spans.forEach((sp, k) => { if (sp[1] > sp[0] && sp[0] < endAt && sp[1] > at) hit.push(mk(lines[k])); });
        hits.push({ phrase: stripStars(ph), hit }); from = endAt;
      }
    }
    return hits.length ? { kind: 'flag', value: REGISTRY_FLAG, hit: hits.reduce((a, h) => a.concat(h.hit), []), src } : null;
  }

  /* ---------- «Требование не установлено» ---------- */
  const SECTION_RE = /^(?:раздел|глава|часть|section|chapter)\s+(?:[ivxlc]+|\d+)\b/i;
  function findNotSet(sources) {
    const out = [];
    for (const src of sources) {
      const lines = buildLines(src); const seen = new Set();
      lines.forEach((l, k) => {
        if (!l.mark) return;
        let lab = null;
        if (l.self && !l.tbl) lab = { text: l.label || l.text, num: l.num, line: l, k };
        if (!lab && l.tbl) {
          const others = l.cells.filter((c, i, a) => c && a.indexOf(c) === i && !isMarkText(c) && !A.numDepth(c));
          if (others.length) lab = { text: others.join(' | '), num: l.num, line: l, k };
        }
        if (!lab) for (let j = k - 1; j >= 0 && k - j <= 14; j--) {
          const x = lines[j]; if (x.mark || x.head) { if (x.head) break; continue; }
          const body = x.tbl ? x.label : x.text; if (norm(body).length < 8) continue;
          lab = { text: norm(x.tbl ? x.rest0 || x.label : x.label), num: x.num, line: x, k: j }; break;
        }
        if (!lab) lab = { text: '(пункт не определён)', num: '', line: l, k };
        const id = src.name + ':' + lab.k; if (seen.has(id)) return; seen.add(id);
        let section = '';
        for (let j = k; j >= 0; j--) { const x = lines[j]; if (x.head || SECTION_RE.test(x.text)) { section = norm(x.text).slice(0, 140); break; } }
        out.push({ file: src.name, srcRef: src, num: lab.num, text: trunc(lab.text, 320), marker: l.mark, section, hit: [mk(lab.line)], miss: [mk(l)], k: l.k, i: l.i, row: l.row });
      });
      // пометки внутри перечня одной ячейки / абзаца: «7.1 … 7.2 … не установлено …»
      lines.forEach((l, k) => {
        if (l.mark) return;
        const sub = splitEnum(l.tbl ? (l.value || l.text) : l.text); if (!sub.length) return;
        sub.forEach(sp => {
          const n = nl(sp.text), m = MARK_RE.exec(n); if (!m || n.length > 420) return;
          if (!(isMarkText(sp.text) || /требовани|допуск|квалификац|членств|лиценз|\bсро\b|опыт|репутац/.test(n))) return;
          let section = '';
          for (let j = k; j >= 0; j--) { const x = lines[j]; if (x.head || SECTION_RE.test(x.text)) { section = norm(x.text).slice(0, 140); break; } }
          out.push({ file: src.name, srcRef: src, num: (l.num ? l.num + ' · ' : '') + sp.num, text: trunc(sp.text, 320), marker: m[0], section, hit: [mk(l)], miss: [mk(l)], k: l.k, i: l.i, row: l.row });
        });
      });
    }
    return out;
  }

  /** Перечень, записанный внутри одной ячейки или абзаца: «7.1. … 7.2. … 7.3. …», «1) … 2) …», «а) … б) …». */
  function splitEnum(text) {
    const t = norm(text), re = /(?:^|[\s;:])((?:\d{1,2}(?:\.\d{1,2}){0,3}[.)])|(?:[а-я]\)))\s+(?=[А-ЯЁA-Z«"(а-яa-z0-9])/g;
    const pos = []; let m;
    while ((m = re.exec(t))) { const at = m.index + m[0].indexOf(m[1]); pos.push({ at, num: m[1].replace(/\.$/, ''), end: m.index + m[0].length }); }
    if (pos.length < 2) return [];
    const out = [];
    for (let i = 0; i < pos.length; i++) {
      const body = t.slice(pos[i].end, i + 1 < pos.length ? pos[i + 1].at : t.length).replace(/[\s;]+$/, '').trim();
      if (body.length >= 3) out.push({ num: pos[i].num, text: body });
    }
    return out.length >= 2 ? out : [];
  }

  /** Пункты внутри найденного раздела: номер, текст, пометка «не установлено». */
  function pointsOf(q) {
    const lines = buildLines(q.srcRef); const pts = [];
    const inRange = l => q.inTable ? (l.i === q.from && l.row >= q.row && l.row < q.rowTo) : (l.i >= q.from && l.i < q.to);
    let cur = null;
    for (const l of lines) {
      if (!inRange(l)) continue;
      if (l.mark) {
        if (l.tbl) {
          const others = l.cells.filter((c, i, a) => c && a.indexOf(c) === i && !isMarkText(c) && !A.numDepth(c));
          if (others.length) { pts.push({ num: l.num, text: trunc(others.join(' | '), 300), notSet: true, mark: l.mark, line: l }); cur = null; continue; }
        }
        if (l.self) { pts.push({ num: l.num, text: trunc(l.label || l.text, 300), notSet: true, mark: l.mark, line: l }); cur = null; continue; }
        if (cur) { cur.notSet = true; cur.mark = l.mark; cur.missLine = l; } else pts.push({ num: '', text: l.mark, notSet: true, mark: l.mark, line: l });
        continue;
      }
      if (l.num || l.head) {
        const sub = splitEnum(l.tbl ? l.value : (l.label || l.text));
        if (sub.length) {
          const lead = l.tbl ? l.label : (l.label || l.text).slice(0, Math.max(0, (l.label || l.text).indexOf(sub[0].num)));
          cur = { num: l.num, text: trunc(lead || l.label || l.text, 300), notSet: false, line: l, head: true };
          pts.push(cur);
          sub.forEach(sp => pts.push({ num: (l.num ? l.num + ' · ' : '') + sp.num, text: trunc(sp.text, 300), notSet: MARK_RE.test(nl(sp.text)) && sp.text.length < 160, mark: (MARK_RE.exec(nl(sp.text)) || [''])[0], line: l, sub: true }));
          continue;
        }
        cur = { num: l.num, text: trunc(l.tbl ? [l.label].concat(l.value ? [l.value] : []).join(' — ') : l.label || l.text, 300), notSet: false, line: l, head: l.head };
        pts.push(cur);
      }
    }
    return pts.filter(p => p.text && p.text.length > 2);
  }

  /* ---------- общая сводка ---------- */
  function pickBest(results) {
    const order = { value: 0, flag: 0, none: 1, notfound: 2 };
    const ok = results.filter(Boolean).sort((a, b) => order[a.kind] - order[b.kind]);
    return ok[0] || null;
  }
  const NOT_FOUND = { kind: 'notfound', value: 'информация не найдена', note: '', hit: [], src: null };

  /** Запасная сводка: если разбор сводки сорвался, остальные блоки всё равно показываются. */
  function emptySummary(res) {
    const nf = (key, title) => Object.assign({ key, title }, NOT_FOUND);
    const sum = { endDate: nf('endDate', 'Дата окончания подачи заявки'), bid: nf('bid', SEC.bid.title), exec: nf('exec', SEC.exec.title), warranty: nf('warranty', SEC.warranty.title), advance: nf('advance', 'Аванс'), criteria: Object.assign(nf('criteria', 'Критерии оценки'), { flag: null }), notSet: [] };
    const rq = (res && res.requirements) || [], fm = (res && res.forms) || [];
    sum.requirements = { key: 'requirements', title: 'Требования', kind: rq.length ? 'value' : 'notfound', count: rq.length, points: 0, notSet: 0, value: rq.length ? 'найдено разделов: ' + rq.length : 'информация не найдена', hits: rq.map((q, i) => ({ q, i })), src: null, hit: [] };
    sum.forms = { key: 'forms', title: 'Список форм', kind: fm.length ? 'value' : 'notfound', forms: fm, value: fm.length ? 'найдено форм: ' + fm.length : 'информация не найдена', src: null, hit: [] };
    return sum;
  }

  function summarize(sources, res) {
    const per = fn => sources.map(s => { try { return fn(s); } catch (e) { console.warn('сводка:', e); return null; } });
    const get = fn => pickBest(per(fn)) || Object.assign({}, NOT_FOUND);
    const sum = {};
    sum.endDate = Object.assign({ key: 'endDate', title: 'Дата окончания подачи заявки' }, get(evalEndDate));
    sum.bid = Object.assign({ key: 'bid', title: SEC.bid.title }, get(s => evalSecurity(SEC.bid, s)));
    sum.exec = Object.assign({ key: 'exec', title: SEC.exec.title }, get(s => evalSecurity(SEC.exec, s)));
    sum.warranty = Object.assign({ key: 'warranty', title: SEC.warranty.title }, get(s => evalSecurity(SEC.warranty, s)));
    sum.advance = Object.assign({ key: 'advance', title: 'Аванс' }, get(evalAdvance));
    if (sum.advance.kind === 'none') sum.advance.kind = 'notfound';
    const crit = Object.assign({ key: 'criteria', title: 'Критерии оценки' }, get(evalCriteria));
    const flags = per(s => findPhrases(s, REGISTRY_PHRASES)).filter(Boolean);
    crit.flag = flags[0] || null;
    sum.criteria = crit;
    // требования — как раньше: найденные разделы + пометки «не установлено»
    const notSet = findNotSet(sources);
    const reqPts = (res.requirements || []).reduce((a, q) => a + pointsOf(q).length, 0);
    sum.requirements = { key: 'requirements', title: 'Требования', kind: res.requirements.length ? 'value' : 'notfound', count: res.requirements.length, points: reqPts, notSet: notSet.length,
      value: res.requirements.length ? 'найдено разделов: ' + res.requirements.length + ', пунктов: ' + reqPts + (notSet.length ? ', из них «требование не установлено»: ' + notSet.length : '') : 'информация не найдена',
      hits: (res.requirements || []).map((q, i) => ({ q, i })), src: null, hit: [] };
    sum.forms = { key: 'forms', title: 'Список форм', kind: res.forms.length ? 'value' : 'notfound', forms: res.forms, value: res.forms.length ? 'найдено форм: ' + res.forms.length : 'информация не найдена', src: null, hit: [] };
    sum.notSet = notSet.sort((a, b) => a.file === b.file ? (a.i - b.i || (a.row || 0) - (b.row || 0) || String(a.num).localeCompare(String(b.num), 'ru', { numeric: true })) : 0);
    return sum;
  }

  /* ---------- документ Word со сводкой ---------- */
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const run = (t, o) => '<w:r><w:rPr>' + (o && o.b ? '<w:b/>' : '') + (o && o.color ? '<w:color w:val="' + o.color + '"/>' : '') + (o && o.sz ? '<w:sz w:val="' + o.sz + '"/>' : '') + '</w:rPr><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
  const para = (t, o) => '<w:p><w:pPr><w:spacing w:before="' + ((o && o.before) || 0) + '" w:after="' + ((o && o.after) || 60) + '"/></w:pPr>' + run(t, o) + '</w:p>';
  function table(widths, rows, head) {
    let x = '<w:tbl><w:tblPr><w:tblW w:w="' + widths.reduce((a, b) => a + b, 0) + '" w:type="dxa"/><w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="8A8A8A"/>').join('') + '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' + widths.map(w => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>';
    rows.forEach((r, ri) => {
      const hd = head && ri === 0;
      x += '<w:tr>' + (hd ? '<w:trPr><w:tblHeader/></w:trPr>' : '');
      r.forEach((c, ci) => {
        const cell = typeof c === 'object' && c !== null ? c : { t: c };
        x += '<w:tc><w:tcPr><w:tcW w:w="' + widths[ci] + '" w:type="dxa"/>' + (hd ? '<w:shd w:val="clear" w:color="auto" w:fill="E3E8F2"/>' : cell.fill ? '<w:shd w:val="clear" w:color="auto" w:fill="' + cell.fill + '"/>' : '') + '</w:tcPr>';
        String(cell.t == null ? '' : cell.t).split('\n').forEach(line => { x += '<w:p><w:pPr><w:spacing w:before="0" w:after="20"/></w:pPr>' + run(line, { b: hd || cell.b, color: cell.color, sz: 22 }) + '</w:p>'; });
        x += '</w:tc>';
      });
      x += '</w:tr>';
    });
    return x + '</w:tbl>';
  }
  const shortFile = s => String(s || '').replace(/\.[^.]+$/, '');

  async function summaryDocx(sum, opts) {
    opts = opts || {}; const files = opts.files || [];
    const rows = [['Показатель', 'Значение', 'Где найдено']];
    const fileOf = f => f.src ? f.src.name : '';
    const val = f => { let v = f.value; if (f.note && f.kind !== 'value') v += '\n(в документации: «' + f.note + '»)'; if (f.flag) v += '\n' + f.flag.value; return v; };
    rows.push(['Дата окончания подачи заявки', val(sum.endDate), fileOf(sum.endDate)]);
    ['bid', 'exec', 'warranty', 'advance', 'criteria'].forEach(k => rows.push([sum[k].title, val(sum[k]), fileOf(sum[k])]));
    rows.push(['Требования', sum.requirements.value, '']);
    rows.push(['Список форм', sum.forms.forms.length ? sum.forms.forms.map(opts.formName).join('\n') : sum.forms.value, '']);
    let body = para('Анализ закупочной документации — общая информация', { b: true, sz: 32, after: 80 });
    if (files.length) body += para('Файлы: ' + files.join('; '), { sz: 20, after: 160 });
    body += table([2900, 5200, 2300], rows, true);
    const miss = opts.missing || [];
    if (opts.withMissing) {
      body += para('Требования, отмеченные в документации как «не установлено»', { b: true, sz: 26, before: 320, after: 100 });
      if (miss.length) body += table([800, 5600, 2000, 2000], [['№', 'Требование', 'Пометка в документации', 'Раздел']].concat(miss.map(m => [m.num || '', m.text, m.marker, m.section])), true);
      else body += para('Таких отметок в документации не найдено.');
    }
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    zip.file('word/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="' + NS_W + '"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>');
    zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="' + NS_W + '"><w:body>' + body + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1134" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>');
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  return { summarize, emptySummary, findNotSet, pointsOf, splitEnum, summaryDocx, buildLines, findPhrases, REGISTRY_FLAG, REGISTRY_PHRASES, isMarkText, shortFile };
});
