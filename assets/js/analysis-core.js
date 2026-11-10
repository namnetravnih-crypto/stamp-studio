/* =====================================================================
   Ассистент — ядро «Анализа документации».
   1) читает Word (.docx; .doc — через упрощённое преобразование), Excel, PDF как список «элементов»;
   2) находит «Состав заявки» и «Требования к участнику» (полная цитата раздела);
   3) находит формы (заявка, техническое предложение, анкета…), определяет их начало и конец
      по маркерам и вырезает форму в отдельный .docx БЕЗ изменения содержимого:
      абзацы, таблицы, стили, нумерация, колонтитулы и сноски остаются как в исходнике.
   Работает и в браузере, и в Node (для проверки).
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('jszip'), require('@xmldom/xmldom'));
  else root.AnalysisCore = factory(root.JSZip, { DOMParser: root.DOMParser, XMLSerializer: root.XMLSerializer });
})(typeof self !== 'undefined' ? self : this, function (JSZip, X) {
  'use strict';
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /* ---------- XML ---------- */
  const parseXml = s => new X.DOMParser().parseFromString(s, 'application/xml');
  const serialize = d => { let s = new X.XMLSerializer().serializeToString(d); if (!/^<\?xml/.test(s)) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s; return s; };
  const lname = n => n.localName || String(n.nodeName).replace(/^.*:/, '');
  const kids = (n, name) => { const o = []; for (let c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && (!name || lname(c) === name)) o.push(c); return o; };
  const kid = (n, name) => kids(n, name)[0] || null;
  const wval = el => el ? (el.getAttributeNS && el.getAttributeNS(NS_W, 'val')) || el.getAttribute('w:val') || '' : '';
  function walk(n, fn) { for (let c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) { if (fn(c) !== false) walk(c, fn); } }
  const attrsOf = el => Array.from(el.attributes || []);

  /* ---------- текст ---------- */
  const normText = s => String(s || '').replace(/[\u00a0\u2000-\u200b\u202f]/g, ' ').replace(/\s+/g, ' ').trim();
  const nl = s => normText(s).toLowerCase().replace(/ё/g, 'е');

  const INLINE = new Set(['hyperlink', 'ins', 'smartTag', 'fldSimple', 'sdt', 'sdtContent', 'customXml']);
  function paraInfo(p) {
    let t = '', pb = false, fn = 0, instr = '', bold = null;
    const run = r => {
      let rb = false; const rpr = kid(r, 'rPr'); if (rpr) { const b = kid(rpr, 'b'); if (b && wval(b) !== '0' && wval(b) !== 'false') rb = true; }
      let had = false;
      for (const c of kids(r)) {
        const n = lname(c);
        if (n === 't') { t += c.textContent; if (c.textContent.trim()) had = true; }
        else if (n === 'tab') t += '\t';
        else if (n === 'br') { if ((c.getAttributeNS && c.getAttributeNS(NS_W, 'type') || c.getAttribute('w:type')) === 'page') pb = true; else t += ' '; }
        else if (n === 'cr') t += ' ';
        else if (n === 'noBreakHyphen') t += '-';
        else if (n === 'instrText') instr += c.textContent;
        else if (n === 'footnoteReference' || n === 'endnoteReference') fn++;
      }
      if (had) bold = bold === null ? rb : (bold && rb);
    };
    const rec = n => { for (const c of kids(n)) { const k = lname(c); if (k === 'r') run(c); else if (INLINE.has(k)) rec(c); } };
    rec(p);
    return { t: t.replace(/[ \t]+/g, m => m.indexOf('\t') >= 0 ? '\t' : ' '), pb, fn, instr, bold: !!bold };
  }
  function tableRows(tbl) {
    const rows = [];
    for (const tr of kids(tbl, 'tr')) {
      const cells = [];
      for (const tc of kids(tr, 'tc')) {
        const parts = []; walk(tc, n => { if (lname(n) === 'p') { const s = normText(paraInfo(n).t); if (s) parts.push(s); return false; } });
        cells.push(parts.join(' '));
      }
      rows.push(cells);
    }
    return rows;
  }

  /* ---------- нумерация / уровни ---------- */
  const NUM_RE = /^\s*(?:(раздел|глава|часть|статья)\s+([ivxlc]+|\d+)\b|(\d+(?:\.\d+)+)\.?\s+(?=\S)|(\d{1,3})[.)]\s+(?=\S)|([ivxlc]{1,5})[.)]\s+(?=\S))/i;
  function manualNumber(t) {
    const m = NUM_RE.exec(t); if (!m) return null;
    if (m[1]) return { num: m[1] + ' ' + m[2], depth: 1 };
    if (m[3]) return { num: m[3], depth: m[3].split('.').length };
    if (m[4]) return { num: m[4], depth: 1 };
    return { num: m[5], depth: 1 };
  }
  const LIST_RE = /^\s*(?:[-–—•·●▪■□○*✓]|\d{1,3}[.)]|[а-яa-z][.)]|\(\d{1,3}\)|\(?[а-я]\)|\d+(?:\.\d+)+\.?)\s+/i;

  async function readStyles(zip) {
    const f = zip.file('word/styles.xml'); const map = {}; if (!f) return map;
    const d = parseXml(await f.async('string'));
    walk(d.documentElement, n => {
      if (lname(n) !== 'style') return;
      const id = n.getAttributeNS(NS_W, 'styleId') || n.getAttribute('w:styleId'); if (!id) return false;
      const name = wval(kid(n, 'name')), based = wval(kid(n, 'basedOn'));
      const ppr = kid(n, 'pPr'); const ol = ppr && kid(ppr, 'outlineLvl');
      map[id] = { name, based, outline: ol ? (parseInt(wval(ol), 10) + 1) : 0 };
      return false;
    });
    return map;
  }
  function styleLevel(styles, id) {
    for (let k = 0; id && k < 6; k++) {
      const s = styles[id]; if (!s) break;
      const m = /^(?:heading|заголовок)\s*(\d)/i.exec(s.name || '') || /^(?:heading|заголовок)(\d)/i.exec(id);
      if (m) return +m[1]; if (s.outline) return s.outline; id = s.based;
    }
    return 0;
  }
  const isTocStyle = (styles, id) => { for (let k = 0; id && k < 5; k++) { const s = styles[id]; const nm = (s && s.name || '') + ' ' + id; if (/(^|\s)(toc\s*\d?|оглавление|содержание|table of contents)/i.test(nm)) return true; id = s && s.based; } return false; };

  const capsShort = t => { const L = t.replace(/[^A-Za-zА-Яа-яЁё]/g, ''); if (L.length < 3 || t.length > 150) return false; const U = L.replace(/[^A-ZА-ЯЁ]/g, ''); return U.length / L.length >= 0.85; };

  /** Один и тот же набор признаков для всех форматов. */
  function decorate(it) {
    const t = it.t, s = normText(t);
    it.s = s; it.n = nl(s); it.empty = !s && !it.tbl;
    const mn = manualNumber(s);
    it.num = mn ? mn.num : ''; it.mdepth = mn ? mn.depth : 0;
    it.depth = it.mdepth || (it.adepth || 0);
    it.list = !it.tbl && (LIST_RE.test(s) || (!!it.adepth && !it.hstyle));
    const short = s.length <= 140 && !/[;,]$/.test(s) && !(s.length > 70 && /\.$/.test(s));
    let hl = 0;
    if (!it.tbl && s) {
      if (it.hstyle && s.length <= 250) hl = it.hstyle;
      else if (/^(?:раздел|глава|часть|section|chapter)\s+(?:[ivxlc]+|\d+)\b/i.test(s) && s.length <= 160) hl = 1;
      else if (mn && short && it.mdepth && !/^\s*\d{1,3}\)/.test(s) && (it.mdepth > 1 || s.length <= 70)) hl = it.mdepth;
      else if (it.adepth && short && s.length <= 120 && !LIST_RE.test(s)) hl = it.adepth;
      else if (capsShort(s)) hl = 2;
      else if (it.bold && s.length <= 120 && !/[.;,]$/.test(s)) hl = 3;
    }
    it.hl = hl;
    return it;
  }

  /* ---------- загрузка источников ---------- */
  async function loadDocx(buf, name) {
    const zip = await JSZip.loadAsync(buf);
    const f = zip.file('word/document.xml'); if (!f) throw new Error('В файле нет word/document.xml — это не .docx');
    const xml = await f.async('string'); const doc = parseXml(xml);
    const styles = await readStyles(zip);
    const body = kid(doc.documentElement, 'body'); const items = [], idxOf = [];
    let i = 0;
    for (const el of kids(body)) {
      const n = lname(el); let it = null;
      if (n === 'p') {
        const pi = paraInfo(el); const ppr = kid(el, 'pPr'); let sid = '', ilvl = -1, numId = 0, outline = 0;
        if (ppr) {
          sid = wval(kid(ppr, 'pStyle'));
          const np = kid(ppr, 'numPr'); if (np) { ilvl = parseInt(wval(kid(np, 'ilvl')) || '0', 10); numId = parseInt(wval(kid(np, 'numId')) || '0', 10); }
          const ol = kid(ppr, 'outlineLvl'); if (ol) outline = parseInt(wval(ol), 10) + 1;
          if (!pi.bold && !pi.t.trim()) pi.bold = false;
        }
        const toc = /PAGEREF|TOC\s|\\o\s/.test(pi.instr) || isTocStyle(styles, sid) || /\.{4,}\s*\d+\s*$/.test(pi.t) || /\t\d{1,4}\s*$/.test(pi.t) && /\.{2,}|…/.test(pi.t);
        it = { t: pi.t, pb: pi.pb, fn: pi.fn, bold: pi.bold, toc, hstyle: styleLevel(styles, sid) || outline, adepth: numId > 0 && ilvl >= 0 ? ilvl + 1 : 0 };
      } else if (n === 'tbl') {
        const rows = tableRows(el);
        it = { t: rows.map(r => r.join(' | ')).join('\n'), rows, tbl: true, fn: 0 };
      } else if (n === 'sdt') {
        let t = '', toc = false; const pr = kid(el, 'sdtPr');
        if (pr) walk(pr, x => { if (lname(x) === 'docPartGallery' && /table of contents|оглавление|содержание/i.test(wval(x))) toc = true; });
        walk(el, x => { if (lname(x) === 'p') { const s = paraInfo(x).t; if (s.trim()) t += (t ? '\n' : '') + s; return false; } });
        it = { t, toc, fn: 0 };
      } else continue;
      it.node = el; it.i = i++; decorate(it); items.push(it); idxOf.push(el);
    }
    return { name, kind: 'docx', buf, items, canCut: true };
  }
  /** Текстовые источники (PDF, Excel, txt): построчно. */
  function loadText(text, name, kind) {
    const items = String(text || '').split(/\r?\n/).map((l, i) => decorate({ t: l.replace(/\t/g, ' '), i, fn: 0, bold: false, toc: /\.{4,}\s*\d+\s*$/.test(l) }));
    return { name, kind: kind || 'text', items, canCut: false };
  }

  /* ---------- «Состав заявки» и «Требования к участнику» ---------- */
  const W = '[а-я]*';   // окончание
  const COMPOSITION = [
    ['состав заявки', new RegExp('состав' + W + '\\s+(?:единой\\s+|первой\\s+|второй\\s+)?(?:част' + W + '\\s+)?(?:заявк' + W + '|предложени' + W + '|документов\\s+(?:и\\s+сведений\\s+)?заявк' + W + ')')],
    ['содержание заявки', new RegExp('(?:требовани' + W + '\\s+к\\s+)?(?:содержани' + W + '|оформлени' + W + '|составлени' + W + '|подготовк' + W + ')\\s+(?:и\\s+(?:состав' + W + '|содержани' + W + '|оформлени' + W + ')\\s+)?заявк' + W)],
    ['заявка должна состоять из…', new RegExp('заявк' + W + '(?:\\s+на\\s+участи' + W + ')?(?:\\s+в\\s+[а-я]+)?\\s+(?:должн' + W + '\\s+)?(?:состо' + W + '|содерж' + W + '|включ' + W + '|быть\\s+(?:составлен' + W + '|оформлен' + W + '|сформирован' + W + ')|состоит)')],
    ['перечень документов', new RegExp('перечень\\s+(?:' + W + '\\s+){0,4}?(?:документ' + W + '|сведени' + W + ')')],
    ['необходимые для участия документы', new RegExp('(?:необходим' + W + '|требуем' + W + ')\\s+(?:для\\s+участи' + W + '\\s+)?(?:в\\s+[а-я\\s]{0,30}?)?документ' + W)],
    ['документы, входящие в состав заявки', new RegExp('(?:документ' + W + '|сведени' + W + '|материал' + W + ')\\s*,?\\s+(?:вход' + W + '|включа' + W + '|включенн' + W + '|прилага' + W + '|прикладыва' + W + '|подава' + W + '|представля' + W + '|предоставля' + W + '|необходим' + W + '|требу' + W + '|подтвержда' + W + ')(?:\\s+[а-я]+){0,4}?\\s+(?:в\\s+состав' + W + '|для\\s+участи' + W + '|заявк' + W + '|участник' + W + '|соответстви' + W + ')')],
    ['к заявке должны быть приложены', new RegExp('к\\s+заявк' + W + '\\s+(?:на\\s+участи' + W + '\\s+)?(?:должн' + W + '\\s+быть\\s+)?(?:прилага' + W + '|прикладыва' + W + '|приложен' + W + '|приобщ' + W + ')')],
    ['в составе заявки', new RegExp('в\\s+состав' + W + '\\s+(?:первой\\s+|второй\\s+)?(?:част' + W + '\\s+)?(?:заявк' + W + '|предложени' + W + ')')],
    ['часть заявки', new RegExp('(?:перв|втор)' + W + '\\s+част' + W + '\\s+заявк' + W)],
    ['опись документов', /опись\s+(?:документов|вложени)/],
    ['документы и сведения', /документы\s+и\s+сведения\s*,?\s+(?:входящие|представляемые|подаваемые|предоставляемые)/]
  ];
  const AGENT = '(?:участник' + W + '|претендент' + W + '|подрядчик' + W + '|исполнител' + W + '|поставщик' + W + '|организаци' + W + '|соискател' + W + '|контрагент' + W + '|заявител' + W + '|кандидат' + W + '|оферент' + W + '|субподрядчик' + W + '|генеральн' + W + '\\s+подрядчик' + W + '|лиц' + W + ',?\\s+(?:желающ' + W + '|подающ' + W + '|выступающ' + W + ')|стороны?\\s+договора)';
  const REQUIREMENTS = [
    ['требования к участнику', new RegExp('требовани' + W + '\\s+(?:' + W + '\\s+){0,3}?к\\s+(?:' + W + '\\s+){0,3}?' + AGENT)],
    ['требования, предъявляемые к участнику', new RegExp('требовани' + W + '\\s*,?\\s+(?:предъявляем' + W + '|устанавливаем' + W + '|установленн' + W + '|налагаем' + W + ')\\s+(?:к|в\\s+отношении)\\s+(?:' + W + '\\s+){0,2}?' + AGENT)],
    ['участник должен соответствовать…', new RegExp(AGENT + '(?:\\s+закупки)?(?:\\s+процедуры\\s+закупки)?(?:\\s+[а-я]+){0,2}?\\s+(?:должен|должны|обязан' + W + ')\\s+(?:' + W + '\\s+){0,2}?(?:соответствовать|отвечать|удовлетворять|иметь|обладать|являться|быть|следующ' + W + '|нижеследующ' + W + ')')],
    ['единые / дополнительные требования', /(?:единые|дополнительные|обязательные|специальные|общие|основные)\s+(?:и\s+дополнительные\s+)?требовани/],
    ['квалификационные требования', /квалификационн[а-яa-z0-9]*\s+(?:требовани|критери|отбор)/],
    ['требования к квалификации / опыту', /требовани[а-яa-z0-9]*\s+к\s+(?:квалификаци|опыт|деловой\s+репутаци|материально|наличию|членству|финансов|кадров|трудов)/],
    ['критерии / условия допуска', /(?:критери[а-яa-z0-9]*\s+допуска|условия\s+допуска|допуск[а-яa-z0-9]*\s+к\s+участию|соответстви[а-яa-z0-9]*\s+участник[а-яa-z0-9]*\s+(?:закупки\s+)?требованиям)/]
  ];

  function matchAtStart(s, rules, maxPos) {
    for (const [label, re] of rules) {
      const m = re.exec(s); if (m && m.index <= maxPos) return { label, index: m.index, text: m[0] };
    }
    return null;
  }
  const LABEL_AFTER_NUM = /^\s*(?:\d+(?:\.\d+)*[.)]?\s+|[а-я][.)]\s+|[-–—•]\s+)/;
  function matchItem(it, rules) {
    const n = it.n; if (!n) return null;
    const len = n.length;
    const pre = (LABEL_AFTER_NUM.exec(n) || [''])[0].length;
    const maxPos = pre + (len > 220 ? 55 : len > 120 ? 80 : 110);
    return matchAtStart(n, rules, maxPos);
  }

  /* ---------- маркеры форм ---------- */
  const VERB_RE = /(?:заполня|представл|предостав|прилага|подает|подаёт|должн|являет|использ|составл|оформл|приведен|содержит|утвержд|прикладыв|может|можно|осуществ|устанавлив|указыва|направля|подписыва)/;
  const FORM_TITLE = /^(?:заявк|анкет|техническ[а-яa-z0-9]+\s+предложени|ценов[а-яa-z0-9]+\s+предложени|коммерческ[а-яa-z0-9]+\s+предложени|предложение\s+(?:о|об|участника)|справк|сведени|декларац|письм|согласи|опись|график|смет|расч[её]т|гарантийн|доверенност|реестр|таблиц|перечень|характеристик|обоснован|заявление|уведомлени|подтвержден|форма\s+заявки)/;
  const NOT_FORM = /(?:проект\s+(?:контракта|договора)|техническое\s+задание|информационн[а-яa-z0-9]+\s+карт|описание\s+объекта|инструкци[а-яa-z0-9]+\s+(?:по|для)|извещение|аукционная\s+документация|документация\s+о\s+закупке|положение\s+о|требования\s+к\s+(?:содержанию|оформлению)\s+заявки)/;
  const SIGN_RE = /(?:м\.\s?п\.|\(подпись\)|\(должность\)|\(ф\.?\s?и\.?\s?о\.?\)|наименование\s+участника|фирменн[а-яa-z0-9]+\s+бланк|«\s*_{1,}\s*»|_{5,}|\(дата\)|подпись\s+руководителя|\(наименование\s+организации\))/;

  function markerOf(it) {
    if (it.tbl || it.toc || !it.s || it.s.length > 220) return null;
    let s = it.n.replace(/^[\s(\[«"*_#]+/, '').replace(/^\d+(?:\.\d+)*[.)]?\s+(?=(?:приложение|форма|образец|бланк))/, '');
    let m = /^(.{3,150}?)\s*\(\s*(?:форма|образец)\s*(?:№|n°?|no\.?)?\s*(\d+[а-яa-z]?)?\s*\)\s*\.?$/.exec(s);
    if (m && !VERB_RE.test(m[1]) && it.s.length <= 160) return { kind: 'form', no: m[2] || '', rest: it.s.slice(0, m[1].length), end: false, suffix: true };
    m = /^начал[оа]\s+формы\s*$/.exec(s.replace(/[\s.\-–—_*)\]»"]+$/, '')); if (m) return { kind: 'delim', no: '', rest: '', end: false };
    m = /^форма(?:\s*(?:№|n°?|no\.?|n)\s*|\s+)?(\d+(?:[.\-]\d+)*[а-яa-z]?)?(?=$|[\s.,:;)\]»"–—-])(.*)$/.exec(s);
    if (m) {
      const rest = it.s.slice(it.s.length - m[2].length).replace(/^[\s.,:;)\]»"–—-]+/, '');
      if (!m[1] && /^(?:[а-яa-z0-9]+\s+){0,1}(?:заполня|представл)/.test(rest)) return null;
      if (it.s.length > 150 || (VERB_RE.test(rest) && rest.length > 60)) return null;
      return { kind: 'form', no: m[1] || '', rest, end: false };
    }
    m = /^формы\s*(?:№|n°?|no\.?)\s*(\d+[а-яa-z]?)\b(.*)$/.exec(s);
    if (m && it.s.length <= 150 && !VERB_RE.test(m[2])) return { kind: 'form', no: m[1], rest: it.s.slice(it.s.length - m[2].length).replace(/^[\s.,:;]+/, ''), end: false };
    m = /^приложение\s*(?:№|n°?|no\.?|n)?\s*(\d+[а-яa-z]?(?:[.\-]\d+)*)(?=$|[\s.,:;)\]»"–—-])(.*)$/.exec(s);
    if (m) return { kind: 'annex', no: m[1], rest: it.s.slice(it.s.length - m[2].length).replace(/^[\s.,:;)\]»"–—-]+/, ''), end: false };
    m = /^(?:образец|бланк)\s+(.*)$/.exec(s);
    if (m && it.s.length <= 130 && !VERB_RE.test(m[1])) return { kind: 'sample', no: '', rest: it.s.slice(it.s.length - m[1].length), end: false };
    return null;
  }
  function endMarkerOf(it) {
    if (it.tbl || !it.s) return false;
    const s = it.n.replace(/^[\s(\[«"*_#\-–—]+/, '').replace(/[\s.)\]»"*_\-–—]+$/, '');
    return /^(?:конец|окончание)\s+(?:формы|бланка|образца|приложения)(?:\s*(?:№|n)?\s*\d+[а-я]?)?$/.test(s) || /^(?:форма|бланк|образец)\s+(?:№\s*\d+\s+)?(?:закончен[а-яa-z0-9]*|завершен[а-яa-z0-9]*)$/.test(s);
  }
  const isSectionStart = it => !it.tbl && !it.toc && it.s && it.s.length <= 160 && /^(?:раздел|глава|часть|section|chapter)\s+(?:[ivxlc]+|\d+)\b/i.test(it.s) && !/форм/i.test(it.s);

  function formLabel(txt) {
    const t = nl(txt);
    const L = [
      [/техническ[а-яa-z0-9]+\s+предложени/, 'Техническое предложение'], [/ценов[а-яa-z0-9]+\s+предложени/, 'Ценовое предложение'], [/коммерческ[а-яa-z0-9]+\s+предложени/, 'Коммерческое предложение'],
      [/(?:предложение\s+о\s+цене|предложение\s+о\s+цене\s+контракта)/, 'Предложение о цене'],
      [/заявк[а-яa-z0-9]+\s+на\s+участи/, 'Заявка на участие'], [/заявк/, 'Заявка'], [/анкет/, 'Анкета'], [/справк/, 'Справка'], [/декларац/, 'Декларация'],
      [/согласи/, 'Согласие'], [/опись/, 'Опись документов'], [/доверенност/, 'Доверенность'], [/гарантийн/, 'Гарантийное письмо'], [/письм/, 'Письмо'],
      [/сведени/, 'Сведения'], [/график/, 'График'], [/смет|расч[её]т/, 'Смета / расчёт'], [/реестр|перечень/, 'Перечень / реестр'], [/предложени/, 'Предложение']
    ];
    for (const [re, name] of L) if (re.test(t)) return name;
    return 'Форма';
  }

  /* ---------- общий анализ ---------- */
  function findFormClusters(items, opts) {
    opts = opts || {};
    const raw = [];
    items.forEach((it, i) => { const m = markerOf(it); if (m) raw.push(Object.assign({ i }, m)); });
    if (opts.titleForms !== false) {
      const have = new Set(raw.map(r => r.i));
      items.forEach((it, i) => {
        if (have.has(i) || it.tbl || it.toc || !it.s || !capsShort(it.s)) return;
        const n = it.n; if (!FORM_TITLE.test(n) || NOT_FORM.test(n) || n.length < 8) return;
        raw.push({ i, kind: 'title', no: '', rest: it.s, end: false });
      });
    }
    raw.sort((a, b) => a.i - b.i);
    const cl = [];
    for (const r of raw) {
      const last = cl[cl.length - 1]; let join = false;
      if (last) {
        const gap = r.i - last.lastI, ks = last.markers.map(m => m.kind);
        if (r.kind === 'title') join = gap <= 5;
        else if (r.kind === 'form' || r.kind === 'sample') join = gap <= 1 || (gap <= 4 && ks.every(k => k === 'annex' || k === 'delim'));
        else if (r.kind === 'annex') join = gap <= 2 && ks.indexOf('annex') < 0;
        else join = gap <= 3;
        for (let k = last.lastI + 1; join && k < r.i; k++) { const x = items[k]; if (x.tbl || x.s.length > 220 || (x.list && x.s.length > 80)) join = false; }
        if (join && r.kind === 'form' && last.markers.some(m => m.kind === 'form' && m.no && r.no && m.no !== r.no)) join = false;
      }
      if (join) { last.markers.push(r); last.lastI = r.i; if (r.no && !last.no) last.no = r.no; }
      else cl.push({ start: r.i, lastI: r.i, no: r.no, markers: [r] });
    }
    const out = [];
    for (const c of cl) {
      const prev = out[out.length - 1];
      if (prev && c.markers.every(m => m.kind === 'title') && prev.markers.some(m => m.kind === 'form' || m.kind === 'sample') && c.start - prev.lastI <= 40) {
        let signed = false; for (let k = prev.lastI + 1; k < c.start; k++) if (/(?:м\.\s?п\.|\(подпись\)|подпись\s+руководителя)/.test(items[k].n || '') || endMarkerOf(items[k])) { signed = true; break; }
        if (!signed) continue;
      }
      out.push(c);
    }
    return out;
  }

  function trimRange(items, s, e, incDelims) {
    const blank = k => items[k].empty || (!items[k].s && !items[k].tbl);
    const delim = k => { const m = markerOf(items[k]); return (m && m.kind === 'delim') || endMarkerOf(items[k]); };
    for (let go = true; go;) { go = false; while (s < e && blank(s)) { s++; go = true; } while (e > s && blank(e - 1)) { e--; go = true; } if (!incDelims) { while (s < e && delim(s)) { s++; go = true; } while (e > s && delim(e - 1)) { e--; go = true; } } }
    return [s, e];
  }

  function findForms(src, opts) {
    opts = opts || {}; const items = src.items, n = items.length, out = [];
    const cl = findFormClusters(items, opts);
    const starts = cl.map(c => c.start);
    cl.forEach((c, k) => {
      const next = k + 1 < cl.length ? starts[k + 1] : n;
      let end = next, endBy = k + 1 < cl.length ? 'next' : 'eof', endMarkerAt = -1;
      for (let j = c.lastI + 1; j < next; j++) {
        if (endMarkerOf(items[j])) { end = j + 1; endBy = 'marker'; endMarkerAt = j; break; }
        if (j > c.lastI + 2 && (isSectionStart(items[j]) || (items[j].hstyle === 1 && !FORM_TITLE.test(items[j].n) && j > c.lastI + 3 && !items[j].toc && items[j].s))) { end = j; endBy = 'section'; break; }
      }
      // слишком длинная форма без явного конца — режем после первого блока подписи
      if ((endBy === 'eof' || endBy === 'next') && end - c.start > 140) {
        for (let j = c.lastI + 1; j < end; j++) if (SIGN_RE.test(items[j].n || '') || (items[j].tbl && SIGN_RE.test(nl(items[j].t)))) { const cut = Math.min(end, j + 7); if (cut < end) { end = cut; endBy = 'signature'; } break; }
      }
      const [s0, e0] = trimRange(items, c.start, end, !!opts.includeDelims);
      // название
      let title = '', after = c.lastI + 1;
      const own = c.markers.filter(m => (m.kind === 'form' || m.kind === 'sample') && m.rest && m.rest.length > 2 && m.rest.length <= 200)[0] || c.markers.filter(m => m.kind === 'title')[0];
      if (own) title = own.kind === 'title' ? items[own.i].s : own.rest;
      if (!title) for (let j = after; j < Math.min(e0, after + 6); j++) {
        const x = items[j]; if (!x.s || x.tbl) continue; const mo = markerOf(x); if (mo && mo.kind !== 'delim') continue;
        if (x.s.length <= 200) { title = x.s; break; }
      }
      const mk = c.markers.filter(m => m.kind !== 'delim');
      const mText = mk.map(m => items[m.i].s).join(' · ');
      const restTxt = mk.map(m => m.rest).filter(Boolean).join(' ');
      const comb = (restTxt + ' ' + title).trim();
      let label = formLabel(comb || mText);
      const body = [];
      for (let j = s0; j < Math.min(e0, s0 + 45); j++) body.push(items[j]);
      const sign = body.filter(x => SIGN_RE.test(x.n || '') || (x.tbl && SIGN_RE.test(nl(x.t)))).length;
      const kinds = new Set(mk.map(m => m.kind));
      let isForm = false, why = '';
      if (kinds.has('form')) { isForm = true; why = 'маркер «Форма»'; }
      else if (kinds.has('sample')) { isForm = true; why = 'маркер «Образец / бланк»'; }
      else {
        const probe = nl(mText + ' ' + comb);
        if (NOT_FORM.test(probe) && !/форм|образец|бланк/.test(probe)) isForm = false;
        else if (FORM_TITLE.test(nl(comb)) || /заявк|анкет|предложени|справк|сведени|декларац|согласи|опись|письм|доверенн|гарант|график|смет|расч[её]т|реестр|форм|образец|бланк/.test(probe)) { isForm = sign >= 1 || /форм|образец|бланк/.test(probe); why = isForm ? 'приложение с признаками формы' : ''; }
        else { isForm = sign >= 2; why = isForm ? 'признаки формы (подпись, М.П.)' : ''; }
        if (!isForm && sign >= 2 && !NOT_FORM.test(probe)) { isForm = true; why = 'признаки формы (подпись, М.П.)'; }
      }
      const nonEmpty = items.slice(s0, e0).filter(x => x.s || x.tbl).length;
      const contentOnly = items.slice(s0, e0).filter(x => (x.s || x.tbl) && !c.markers.some(m => m.i === x.i)).length;
      if (!isForm || contentOnly < (kinds.has('form') || kinds.has('sample') ? 2 : 3) || e0 <= s0) return;
      const no = c.no || '';
      const text = items.slice(s0, e0).map(x => x.t).join('\n');
      out.push({
        file: src.name, srcRef: src, start: s0, end: e0, no, label, title: title || restTxt || mText,
        marker: mText, endBy, why, items: nonEmpty, footnotes: items.slice(s0, e0).reduce((a, x) => a + (x.fn || 0), 0),
        tables: items.slice(s0, e0).filter(x => x.tbl).length, preview: text.slice(0, 1200), words: text.split(/\s+/).length
      });
    });
    return out;
  }

  /* ---------- цитаты ---------- */
  const CAP_ITEMS = 160, CAP_CHARS = 40000;
  function quoteRange(items, i, formStartSet) {
    const it = items[i]; const n = items.length; let j = i + 1, chars = it.s.length, truncated = false;
    const stop = k => formStartSet.has(k) || endMarkerOf(items[k]);
    const heading = it.hl > 0 && it.s.length <= 160 && !(it.num && /[.;]$/.test(it.s));
    const intro = /[:：]\s*$/.test(it.s);
    const lastNonEmpty = k => { for (; k >= i; k--) if (items[k].s || items[k].tbl) return items[k]; return it; };
    let seenList = false;
    while (j < n) {
      const x = items[j];
      if (x.toc) { j++; continue; }
      if (stop(j)) break;
      if (j - i > CAP_ITEMS || chars > CAP_CHARS) { truncated = true; break; }
      if (!x.s && !x.tbl) { j++; continue; }
      if (heading) {
        if (x.hl > 0 && x.hl <= it.hl && !(x.hl === 3 && it.hl < 3 && false)) { if (!(x.hl >= 3 && it.hl >= 3 && false)) break; }
        if (it.hl === 3 && x.mdepth === 1) break;
      } else {
        if (x.hl > 0 && x.hl <= 2 && !x.list) break;
        if (it.depth && x.depth && x.depth <= it.depth && !(x.adepth && !x.mdepth && x.list && it.adepth && x.adepth > it.adepth)) break;
        if (!x.tbl && !x.list && !(x.depth > it.depth && it.depth)) {
          const prev = lastNonEmpty(j - 1), pe = /[:;,]\s*$/.test(prev.s);
          if (!(pe || (it.depth && !x.depth && x.s.length > 0 && x.s.length < 200 && (seenList || intro) && /^[а-яa-z(«"-]/.test(x.s)))) { if (seenList || !intro) break; }
        }
        if (x.list || x.tbl) seenList = true;
      }
      chars += x.s.length + 1; j++;
    }
    let e = j; while (e > i + 1 && !items[e - 1].s && !items[e - 1].tbl) e--;
    return { from: i, to: e, truncated };
  }
  function quoteText(items, from, to) {
    const out = [];
    for (let k = from; k < to; k++) {
      const x = items[k]; if (x.toc) continue;
      if (x.tbl) { out.push(x.rows.map(r => r.join(' | ')).join('\n')); } else if (x.s) out.push(x.s);
    }
    return out.join('\n');
  }
  function findQuotes(src, rules, cat, formStartSet, inForm) {
    const items = src.items, res = [];
    items.forEach((it, i) => {
      if (it.toc || (inForm && inForm(i))) return;
      if (it.tbl) {
        it.rows.forEach((cells, r) => {
          if (!cells.length) return;
          const label = nl((cells[0].length < 8 && cells[1]) ? cells[1] : cells[0]); if (!label) return;
          const mm = matchAtStart(label, rules, (LABEL_AFTER_NUM.exec(label) || [''])[0].length + 110); if (!mm) return;
          const uniq = cells.filter((c, k) => c && cells.indexOf(c) === k);
          res.push({ file: src.name, cat, kw: mm.label, matched: mm.text, from: i, to: i + 1, row: r, num: '', title: uniq[0] || '', text: uniq.join('\n'), truncated: false, inTable: true });
        });
        return;
      }
      if (/^(?:в\s+соответстви|согласно|на\s+основании|во\s+исполнение|с\s+учетом|подтвержда|настоящим|мы\s|нами\s|заявляем|гарантиру)/.test(it.n)) return;
      const m = matchItem(it, rules); if (!m) return;
      const q = quoteRange(items, i, formStartSet);
      res.push({ file: src.name, cat, kw: m.label, matched: m.text, from: q.from, to: q.to, num: it.num || '', title: it.s.slice(0, 220), text: quoteText(items, q.from, q.to), truncated: q.truncated });
    });
    // убираем вложенные и повторы
    res.sort((a, b) => a.from - b.from || (b.to - b.from) - (a.to - a.from));
    const keep = [];
    for (const r of res) {
      const dup = keep.find(k => !r.inTable && !k.inTable && r.from >= k.from && r.from < k.to) || keep.find(k => r.inTable && k.inTable && k.from === r.from && k.row === r.row);
      if (dup) { if (!dup.alt) dup.alt = []; if (dup.alt.indexOf(r.kw) < 0 && r.kw !== dup.kw) dup.alt.push(r.kw); continue; }
      keep.push(r);
    }
    return keep.filter(r => r.text.replace(/\s+/g, ' ').length >= 12 && !(r.to - r.from <= 1 && !r.inTable && r.text.length < 70 && src.items[r.from].hl > 0));
  }

  function analyze(sources, opts) {
    opts = opts || {}; const r = { composition: [], requirements: [], forms: [] };
    for (const src of sources) {
      const cl = findFormClusters(src.items, opts); const set = new Set(cl.map(c => c.start));
      const forms = findForms(src, opts); r.forms.push(...forms);
      const inForm = k => forms.some(f => k >= f.start && k < f.end);
      r.composition.push(...findQuotes(src, COMPOSITION, 'composition', set, inForm));
      r.requirements.push(...findQuotes(src, REQUIREMENTS, 'requirements', set, inForm));
    }
    return r;
  }

  /* ---------- вырезание формы в отдельный .docx ---------- */
  function resolvePart(base, target) {
    const parts = (base + target).split('/'), o = [];
    for (const p of parts) { if (p === '..') o.pop(); else if (p && p !== '.') o.push(p); }
    return o.join('/');
  }
  const PRUNE_TYPES = /\/(image|hyperlink|oleObject|package|chart|audio|video|diagramData|diagramLayout|diagramQuickStyle|diagramColors|slide|drawing)$/;

  /** Возвращает Uint8Array с отдельным .docx, где оставлены элементы тела [start, end). */
  async function cutDocx(src, start, end) {
    const zip = await JSZip.loadAsync(src.buf);
    const doc = parseXml(await zip.file('word/document.xml').async('string'));
    const body = kid(doc.documentElement, 'body');
    const els = kids(body).filter(e => ['p', 'tbl', 'sdt'].indexOf(lname(e)) >= 0);
    const bodySect = kid(body, 'sectPr');
    // какое оформление страницы действует на конец формы
    let sect = null, moveFromLast = false;
    for (let k = Math.max(end - 1, 0); k < els.length; k++) {
      const e = els[k]; if (lname(e) !== 'p') continue; const ppr = kid(e, 'pPr'); const sp = ppr && kid(ppr, 'sectPr');
      if (sp) { sect = sp.cloneNode(true); if (k === end - 1) moveFromLast = true; break; }
    }
    if (!sect && bodySect) sect = bodySect.cloneNode(true);
    const keep = new Set(els.slice(start, end));
    for (const e of kids(body)) if (e !== bodySect && !keep.has(e)) body.removeChild(e);
    if (bodySect) body.removeChild(bodySect);
    if (moveFromLast) { const last = els[end - 1]; const ppr = kid(last, 'pPr'); const sp = ppr && kid(ppr, 'sectPr'); if (sp) ppr.removeChild(sp); }
    if (sect) body.appendChild(sect);
    // хвостовой разрыв страницы в последнем абзаце убираем, чтобы не было пустой страницы
    const kept = els.slice(start, end);
    for (let k = kept.length - 1; k >= 0; k--) {
      const p = kept[k]; if (lname(p) !== 'p') break;
      const runs = kids(p, 'r'); const lr = runs[runs.length - 1];
      if (lr) { const cs = kids(lr).filter(c => lname(c) !== 'rPr'); if (cs.length === 1 && lname(cs[0]) === 'br' && (cs[0].getAttributeNS(NS_W, 'type') || cs[0].getAttribute('w:type')) === 'page') p.removeChild(lr); }
      if (!paraInfo(p).t.trim() && !p.getElementsByTagName('w:drawing').length) continue; break;
    }
    // сноски: оставляем те, на которые есть ссылки (и служебные разделители)
    const fnIds = new Set(), enIds = new Set(); const bmS = {}, bmE = {}, cmS = {}, cmE = {}; const usedR = new Set();
    walk(body, n => {
      const k = lname(n);
      if (k === 'footnoteReference') fnIds.add(n.getAttributeNS(NS_W, 'id') || n.getAttribute('w:id'));
      else if (k === 'endnoteReference') enIds.add(n.getAttributeNS(NS_W, 'id') || n.getAttribute('w:id'));
      else if (k === 'bookmarkStart') bmS[n.getAttributeNS(NS_W, 'id') || n.getAttribute('w:id')] = n;
      else if (k === 'bookmarkEnd') bmE[n.getAttributeNS(NS_W, 'id') || n.getAttribute('w:id')] = n;
      else if (k === 'commentRangeStart') cmS[n.getAttributeNS(NS_W, 'id') || n.getAttribute('w:id')] = n;
      else if (k === 'commentRangeEnd') cmE[n.getAttributeNS(NS_W, 'id') || n.getAttribute('w:id')] = n;
      for (const a of attrsOf(n)) if (a.namespaceURI === NS_R || /^r:/.test(a.name)) usedR.add(a.value);
    });
    const orphan = (S, E) => { for (const id in S) if (!E[id]) S[id].parentNode.removeChild(S[id]); for (const id in E) if (!S[id]) E[id].parentNode.removeChild(E[id]); };
    orphan(bmS, bmE); orphan(cmS, cmE);
    for (const [path, ids, tag] of [['word/footnotes.xml', fnIds, 'footnote'], ['word/endnotes.xml', enIds, 'endnote']]) {
      const f = zip.file(path); if (!f) continue;
      const d = parseXml(await f.async('string'));
      for (const e of kids(d.documentElement, tag)) {
        const type = e.getAttributeNS(NS_W, 'type') || e.getAttribute('w:type'); const id = e.getAttributeNS(NS_W, 'id') || e.getAttribute('w:id');
        if (!type && !ids.has(id)) d.documentElement.removeChild(e);
      }
      zip.file(path, serialize(d));
    }
    // связи: убираем рисунки, ссылки и т.п., которых в вырезанной форме нет
    const relPath = 'word/_rels/document.xml.rels'; const rf = zip.file(relPath);
    if (rf) {
      const rd = parseXml(await rf.async('string')); const drop = [];
      for (const r of kids(rd.documentElement, 'Relationship')) {
        const id = r.getAttribute('Id'), type = r.getAttribute('Type') || '';
        if (PRUNE_TYPES.test(type) && !usedR.has(id)) drop.push(r);
      }
      const dropTargets = [];
      for (const r of drop) { if (r.getAttribute('TargetMode') !== 'External') dropTargets.push(resolvePart('word/', r.getAttribute('Target'))); rd.documentElement.removeChild(r); }
      zip.file(relPath, serialize(rd));
      if (dropTargets.length) {
        const still = new Set();
        for (const name of Object.keys(zip.files)) {
          if (!/\.rels$/.test(name) || zip.files[name].dir) continue;
          const dir = name.replace(/_rels\/[^/]*$/, '');
          const d = parseXml(await zip.file(name).async('string'));
          for (const r of kids(d.documentElement, 'Relationship')) if (r.getAttribute('TargetMode') !== 'External') still.add(resolvePart(dir, r.getAttribute('Target')));
        }
        for (const t of dropTargets) if (!still.has(t) && zip.file(t)) zip.remove(t);
      }
    }
    zip.file('word/document.xml', serialize(doc));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  /* ---------- .docx из текста (для PDF / Excel / txt) ---------- */
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  async function docxFromLines(lines, title) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    zip.file('word/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="' + NS_W + '"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>');
    let b = '', i = 0;
    const para = t => '<w:p><w:r><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r></w:p>';
    while (i < lines.length) {
      if (/ \| /.test(lines[i])) {
        const rows = []; while (i < lines.length && / \| /.test(lines[i])) rows.push(lines[i++].split(' | ')); const cols = Math.max(...rows.map(r => r.length));
        b += '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="000000"/>').join('') + '</w:tblBorders></w:tblPr><w:tblGrid>' + '<w:gridCol w:w="' + Math.floor(9300 / cols) + '"/>'.repeat(1) + '</w:tblGrid>';
        b = b.replace(/<w:tblGrid>.*$/, '<w:tblGrid>' + ('<w:gridCol w:w="' + Math.floor(9300 / cols) + '"/>').repeat(cols) + '</w:tblGrid>');
        for (const r of rows) { b += '<w:tr>'; for (let c = 0; c < cols; c++) b += '<w:tc><w:tcPr><w:tcW w:w="' + Math.floor(9300 / cols) + '" w:type="dxa"/></w:tcPr>' + para(r[c] || '') + '</w:tc>'; b += '</w:tr>'; }
        b += '</w:tbl><w:p/>';
      } else b += lines[i].trim() ? para(lines[i]) : '<w:p/>', i++;
    }
    zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="' + NS_W + '"><w:body>' + b + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>');
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async function exportForm(form) {
    const src = form.srcRef;
    if (src.canCut) return cutDocx(src, form.start, form.end);
    const lines = src.items.slice(form.start, form.end).map(x => x.t);
    return docxFromLines(lines, form.title);
  }

  const safeName = s => String(s || '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/[«»“”]/g, '').replace(/\s+/g, ' ').trim().replace(/[. ]+$/, '');
  function formFileName(f, k) {
    const head = f.no ? 'Форма ' + f.no : 'Форма ' + (k + 1);
    let tail = safeName(f.title && f.title.length > 3 ? f.title : f.label);
    if (tail.toLowerCase().indexOf(head.toLowerCase()) === 0) tail = tail.slice(head.length).replace(/^[\s.\-–—:]+/, '');
    let name = safeName(head + (tail ? ' — ' + tail : ''));
    if (name.length > 90) name = name.slice(0, 90).replace(/\s+\S*$/, '').trim();
    return name + '.docx';
  }

  return { loadDocx, loadText, analyze, findForms, findFormClusters, cutDocx, exportForm, docxFromLines, formFileName, quoteText, markerOf, endMarkerOf, nl, normText, COMPOSITION, REQUIREMENTS };
});
