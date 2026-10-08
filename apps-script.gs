/**
 * Журнал рейсів у Google-таблиці — серверна частина для калькулятора
 * https://escoreodessa.github.io/logistica/
 *
 * Як підключити:
 * 1. Створіть Google-таблицю → Розширення → Apps Script.
 * 2. Видаліть усе в редакторі, вставте цей код, збережіть.
 * 3. (Необов'язково) впишіть секретне слово в KEY і те саме — у калькуляторі.
 * 4. Розгорнути → Нове розгортання → тип «Веб-застосунок»:
 *      Виконувати як: «Я»;  Хто має доступ: «Будь-хто».
 *    Дозвольте доступ до таблиці, скопіюйте URL веб-застосунку.
 * 5. У калькуляторі: вкладка «Журнал рейсів» → «Google-таблиця» → вставте URL.
 */
const KEY = '';            // секретне слово (необов'язково)
const SHEET = 'Рейси';     // назва аркуша з журналом
const ID_COL = 'ID рейсу';

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function headers_(sh) {
  return sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0] : [];
}

function doGet() {
  return out_({ ok: true, app: 'logistica', message: 'Веб-застосунок працює' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (KEY && d.key !== KEY) return out_({ ok: false, error: 'Невірний ключ' });
    const sh = sheet_();

    if (d.action === 'list') {
      const v = sh.getDataRange().getDisplayValues();
      let rows = [];
      if (v.length >= 2) {
        const h = v[0];
        rows = v.slice(1).filter(r => r.join('') !== '')
          .map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
      }
      return out_({ ok: true, rows: rows, items: listItems_(), v: 2 });
    }

    if (d.action === 'sync') {
      (d.items || []).forEach(it => upsertItem_(it));
      (d.deleted || []).forEach(k => deleteItem_(k));
      return out_({ ok: true });
    }

    if (d.action === 'add') {
      const rows = d.rows || [];
      let h = headers_(sh);
      // нові колонки додаються в кінець
      rows.forEach(r => Object.keys(r).forEach(k => { if (h.indexOf(k) < 0) h.push(k); }));
      if (h.length) {
        sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
        sh.setFrozenRows(1);
      }
      const idIdx = h.indexOf(ID_COL);
      const existing = new Set();
      if (idIdx >= 0 && sh.getLastRow() > 1)
        sh.getRange(2, idIdx + 1, sh.getLastRow() - 1, 1).getDisplayValues().forEach(x => existing.add(x[0]));
      let added = 0;
      rows.forEach(r => {
        if (idIdx >= 0 && existing.has(String(r[ID_COL]))) return;   // вже є — не дублюємо
        const vals = h.map(k => (r[k] === undefined || r[k] === null) ? '' : r[k]);
        const fmts = vals.map(v => typeof v === 'number' ? 'General' : '@');
        const row = sh.getLastRow() + 1;
        sh.getRange(row, 1, 1, h.length).setNumberFormats([fmts]).setValues([vals]);
        existing.add(String(r[ID_COL])); added++;
      });
      return out_({ ok: true, added: added });
    }

    if (d.action === 'delete') {
      const h = headers_(sh), idIdx = h.indexOf(ID_COL);
      if (idIdx < 0 || sh.getLastRow() < 2) return out_({ ok: true, deleted: 0 });
      const ids = sh.getRange(2, idIdx + 1, sh.getLastRow() - 1, 1).getDisplayValues();
      let deleted = 0;
      for (let i = ids.length - 1; i >= 0; i--) {
        if (ids[i][0] === String(d.id)) { sh.deleteRow(i + 2); deleted++; }
      }
      return out_({ ok: true, deleted: deleted });
    }

    return out_({ ok: false, error: 'Невідома дія' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Разове прибирання ТО/ремонту (запустити вручну один раз:
 * вибрати cleanupRepair у списку функцій угорі → Run).
 * Видаляє колонки «ТО/ремонт…» і перераховує «Разом, грн» = пальне + амортизація та «Грн/км».
 */
function cleanupRepair() {
  const sh = sheet_();
  let h = headers_(sh);
  for (let i = h.length - 1; i >= 0; i--) {
    if (/^ТО\/ремонт/.test(h[i])) sh.deleteColumn(i + 1);
  }
  h = headers_(sh);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const col = k => h.indexOf(k) + 1;
  const km = sh.getRange(2, col('Км'), n, 1).getValues();
  const fuel = sh.getRange(2, col('Пальне, грн'), n, 1).getValues();
  const am = sh.getRange(2, col('Амортизація, грн'), n, 1).getValues();
  const tot = [], per = [];
  for (let i = 0; i < n; i++) {
    const t = Math.round((Number(fuel[i][0]) + Number(am[i][0])) * 100) / 100;
    tot.push([t]);
    per.push([Number(km[i][0]) ? Math.round(t / Number(km[i][0]) * 100) / 100 : '']);
  }
  sh.getRange(2, col('Разом, грн'), n, 1).setValues(tot);
  sh.getRange(2, col('Грн/км'), n, 1).setValues(per);
}


/* ===== Спільні налаштування (машини, ціни, склад, підприємство) і подорожні листи ===== */
const SET_SHEET = 'Налаштування';
const WB_SHEET = 'Подорожні листи';
const SET_HEAD = ['Ключ', 'Значення (JSON)', 'Оновлено', 'Ким'];
const WB_HEAD = ['Ключ', 'Дата', 'Машина', '№ листа', 'Водій', 'Держ. номер', 'Спідометр виїзд', 'Спідометр повернення',
  'Пальне виїзд', 'Заправлено', 'Чек №', 'Пальне повернення', 'Механік', 'Медпрацівник', 'Бухгалтер', 'Оновлено', 'Ким', 'JSON'];
function kvSheet_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function findRow_(sh, key) {
  if (sh.getLastRow() < 2) return 0;
  const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < keys.length; i++) if (keys[i][0] === key) return i + 2;
  return 0;
}
function upsertItem_(it) {
  if (!it || !it.key) return;
  const isWb = String(it.key).indexOf('wb:') === 0;
  const sh = isWb ? kvSheet_(WB_SHEET, WB_HEAD) : kvSheet_(SET_SHEET, SET_HEAD);
  const row = findRow_(sh, it.key);
  if (row) {   // новіший запис перемагає
    const atCol = isWb ? 16 : 3;
    const prev = Number(sh.getRange(row, atCol).getValue()) || 0;
    if (prev > Number(it.at || 0)) return;
  }
  const json = JSON.stringify(it.value);
  let vals;
  if (isWb) {
    const w = it.value || {}, parts = String(it.key).slice(3).split('|');
    vals = [it.key, parts[0] || '', parts[1] || '', w.no || '', w.driver || '', w.plate || '', w.odoOut || '', w.odoIn || '',
      w.fuelOut || '', w.fueled || '', w.receipt || '', w.fuelIn || '', w.mechanic || '', w.medic || '', w.accountant || '',
      Number(it.at || 0), it.by || '', json];
  } else {
    vals = [it.key, json, Number(it.at || 0), it.by || ''];
  }
  const r = row || sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, vals.length).setNumberFormats([vals.map(v => typeof v === 'number' ? '0' : '@')]).setValues([vals]);
}
function deleteItem_(key) {
  [kvSheet_(SET_SHEET, SET_HEAD), kvSheet_(WB_SHEET, WB_HEAD)].forEach(sh => {
    const row = findRow_(sh, key);
    if (row) sh.deleteRow(row);
  });
}
function listItems_() {
  const items = [];
  [[SET_SHEET, SET_HEAD, 1, 2, 3], [WB_SHEET, WB_HEAD, 17, 15, 16]].forEach(([name, head, jsonIdx, atIdx, byIdx]) => {
    const sh = kvSheet_(name, head);
    if (sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getDisplayValues().forEach(r => {
      if (!r[0]) return;
      try { items.push({ key: r[0], value: JSON.parse(r[jsonIdx]), at: Number(r[atIdx]) || 0, by: r[byIdx] }); } catch (e) {}
    });
  });
  return items;
}
