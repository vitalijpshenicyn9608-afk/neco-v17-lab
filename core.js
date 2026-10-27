// NECO planner: чистая логика без вызовов Office и без записи. Проверяется тестами в Node.
(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.NecoPlanner = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  // Листы, в которые разрешена запись. 'Old outports' намеренно отсутствует.
  const SHEETS = ['QINGDAO, NANSHA, OUTPORTS', 'TIANJIN, DALIAN', 'SHANGHAI, NINGBO'];
  const PORTS = {
    'QINGDAO, NANSHA, OUTPORTS': ['QINGDAO', 'NANSHA', 'RIZHAO', 'LIANYUNGANG', 'XIAMEN'],
    'TIANJIN, DALIAN': ['TIANJIN', 'DALIAN'],
    'SHANGHAI, NINGBO': ['SHANGHAI', 'NINGBO']
  };
  const DUP_DAYS = 5;            // сегодня и 4 предыдущих календарных дня по столбцу A
  const MAX_AGE_HOURS = 12;      // заявка из буфера старше этого срока отклоняется
  const MAX_ROWS = 60;
  const MAX_SHEET_ROW = 200000;
  // Столбцы, по которым строка считается занятой (всё, кроме A и G): B C D E F H I J K L M
  const BUSY_COLS = [2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13];

  const norm = v => String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toUpperCase();
  const isEmpty = v => v === null || v === undefined || String(v).trim() === '';

  function validDate(y, m, d) {
    const x = new Date(y, m - 1, d);
    return x.getFullYear() === y && x.getMonth() === m - 1 && x.getDate() === d ? x : null;
  }
  function parseDate(value) {
    if (typeof value === 'number' && Number.isFinite(value) && value > 30000 && value < 80000) {
      const d = new Date(Math.round((value - 25569) * 86400000));
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    const text = String(value == null ? '' : value).trim();
    let m = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})/);
    if (m) return validDate(+m[3], +m[2], +m[1]);
    m = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return validDate(+m[1], +m[2], +m[3]);
    return null;
  }
  function excelSerial(d) {
    return Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000) + 25569;
  }
  const key = o => [norm(o.booking), norm(o.port), String(Number(o.quantity)), norm(o.type)].join('|');

  function validatePayload(p, now) {
    if (!p || p.schema !== 2 || p.app !== 'NECO') throw Error('Это не заявка NECO (в буфере другие данные).');
    if (!SHEETS.includes(p.sheet)) throw Error('Неизвестный лист «' + p.sheet + '». Запись разрешена только в три рабочих листа.');
    if (!Array.isArray(p.rows) || p.rows.length < 1 || p.rows.length > MAX_ROWS) throw Error('В заявке должно быть от 1 до ' + MAX_ROWS + ' строк.');
    const created = Date.parse(p.createdAt);
    if (!Number.isFinite(created)) throw Error('В заявке нет времени создания.');
    const ageH = ((now || new Date()).getTime() - created) / 3600000;
    if (ageH > MAX_AGE_HOURS) throw Error('Заявка создана более ' + MAX_AGE_HOURS + ' ч назад. Сформируйте её в программе заново.');
    if (ageH < -1) throw Error('Время создания заявки в будущем: проверьте часы компьютера.');
    for (const r of p.rows) {
      if (!/^LED\d{5,10}$/.test(norm(r.booking))) throw Error('Неверный номер букинга: ' + r.booking);
      if (!PORTS[p.sheet].includes(norm(r.port))) throw Error('Порт «' + r.port + '» не относится к листу ' + p.sheet);
      const q = Number(r.quantity);
      if (!Number.isInteger(q) || q < 1 || q > 1500) throw Error('Некорректное количество у ' + r.booking);
      if (!/^(20|40|45)[A-Z]{2,3}$/.test(norm(r.type))) throw Error('Нераспознанный тип контейнера у ' + r.booking + ': ' + r.type);
    }
  }

  // values: массив строк листа, начиная со строки firstRow (1-based), по 13 столбцов A:M.
  function plan(values, firstRow, payload, today, now) {
    validatePayload(payload, now);
    if (!Array.isArray(values) || !values.length || !Array.isArray(values[0]) || values[0].length < 13) throw Error('Не удалось прочитать столбцы A:M.');
    const day = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const cutoff = new Date(day); cutoff.setDate(cutoff.getDate() - (DUP_DAYS - 1));
    const at = (r, c) => { const row = values[r - firstRow]; return row ? row[c - 1] : ''; };
    const lastRow = firstRow + values.length - 1;

    const existing = new Set();
    let lastData = 0, gTemplate = 0;
    for (let r = firstRow; r <= lastRow; r++) {
      if (BUSY_COLS.some(c => !isEmpty(at(r, c)))) lastData = r;
      const booking = norm(at(r, 3));
      if (!booking) continue;
      if (norm(at(r, 7)) === 'PENDING') gTemplate = r;
      const date = parseDate(at(r, 1));
      if (date && date >= cutoff && date <= day) {
        existing.add(key({ booking, port: at(r, 11), quantity: at(r, 4), type: at(r, 5) }));
      }
    }
    if (!lastData) throw Error('В прочитанной части листа нет данных: запись остановлена.');

    let cursor = lastData + 1;
    const seen = new Set(), results = [];
    for (const item of payload.rows) {
      const k = key(item);
      if (seen.has(k)) { results.push({ item, kind: 'duplicate', why: 'повтор внутри заявки', row: null }); continue; }
      seen.add(k);
      if (existing.has(k)) { results.push({ item, kind: 'duplicate', why: 'уже есть в таблице за ' + DUP_DAYS + ' дн.', row: null }); continue; }
      let found = false;
      for (let tries = 0; tries < 400; tries++, cursor++) {
        if (cursor > MAX_SHEET_ROW) throw Error('Слишком большой номер строки: запись остановлена.');
        const busy = BUSY_COLS.some(c => !isEmpty(at(cursor, c)));
        const g = norm(at(cursor, 7));
        if (!busy && (g === '' || g === 'PENDING')) { found = true; break; }
      }
      if (!found) throw Error('Не удалось найти свободную строку после занятых данных.');
      results.push({ item, kind: 'new', row: cursor });
      cursor++;
    }
    return {
      results, lastData, gTemplate,
      added: results.filter(r => r.kind === 'new').length,
      duplicates: results.filter(r => r.kind === 'duplicate').length
    };
  }

  return { SHEETS, PORTS, DUP_DAYS, BUSY_COLS, norm, key, plan, parseDate, excelSerial, validatePayload };
});
