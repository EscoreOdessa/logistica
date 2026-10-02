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
      if (v.length < 2) return out_({ ok: true, rows: [] });
      const h = v[0];
      const rows = v.slice(1).filter(r => r.join('') !== '')
        .map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
      return out_({ ok: true, rows: rows });
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
