/* NECO Booking add-in v1.7 beta. Запись только в пустые строки, с повторной проверкой и перечитыванием.
   Два входа: ручной (вставка из буфера) и автоматический (сообщение SUBMIT от расширения-курьера). */
'use strict';

// ---- Настройки ----------------------------------------------------------
// Часть имени файла, в который разрешена запись. Для тестов — клон «ПРОБА».
const EXPECTED_FILE_PART = 'ПРОБА';
// Сколько последних строк листа читать для проверок (дубли за 5 дней и поиск свободной строки).
const WINDOW_ROWS = 2000;
// Формат даты в столбце A.
const DATE_FORMAT = 'dd.mm.yyyy';
// Переносить оформление (границы, заливки, форматы) с последней строки со статусом pending.
const COPY_ROW_FORMATS = true;
// ------------------------------------------------------------------------

const P = NecoPlanner;
const el = id => document.getElementById(id);
let payload = null;
let shownPlanKey = '';
let busy = false;
const handledJobs = new Set();

const showErr = m => { el('error').textContent = m || ''; };
function planKey(pl) { return pl.results.map(r => r.kind + ':' + (r.row || '-') + ':' + P.key(r.item)).join(';'); }
function labelPlan(pl) {
  const lines = pl.results.map(r => {
    const it = r.item, base = it.booking + ' · ' + it.port + ' · ' + it.quantity + ' × ' + it.type;
    return r.kind === 'duplicate' ? base + ': ПРОПУЩЕН (' + r.why + ')' : base + ': строка ' + r.row;
  });
  lines.push('', 'Новых: ' + pl.added + '; дублей: ' + pl.duplicates);
  return lines.join('\n');
}
function docUrl() { try { return decodeURIComponent(String(Office.context.document.url || '')); } catch (e) { return String(Office.context.document.url || ''); } }

function necoBookIdentity(raw) {
  try {
    const u = new URL(String(raw || '').trim());
    const q = u.searchParams;
    // Only strong document identifiers; URL paths alone are not reliable in Excel Online.
    const source = q.get('sourcedoc') || q.get('sourceDoc');
    if (source) return 'guid:' + source.replace(/[{}]/g, '').trim().toLowerCase();
    const resid = q.get('resid') || q.get('id');
    if (resid && /^[a-z0-9!._-]+$/i.test(resid.trim())) return 'resid:' + resid.trim().toLowerCase();
    const share = u.pathname.match(/\/([a-f0-9]{8}-[a-f0-9-]{20,})\/?$/i);
    if (share) return 'guid:' + share[1].toLowerCase();
    return '';
  } catch (_) { return ''; }
}
function necoRequireSameBook(actualUrl, expectedUrl) {
  const actual = necoBookIdentity(actualUrl), expected = necoBookIdentity(expectedUrl);
  if (!actual || !expected) throw Error('Не удалось подтвердить ID книги Excel (sourcedoc/resid). Запись запрещена. Открыто: ' + actualUrl + '; ожидается: ' + expectedUrl);
  if (actual !== expected) throw Error('Открыта другая книга Excel. Запись запрещена. ID открытой: ' + actual + '; ID ожидаемой: ' + expected);
}
function checkEnv() {
  const url = docUrl();
  const expected = payload && String(payload.expectedWorkbookUrl || '').trim();
  if (!url || !expected) throw Error('Для записи требуется привязка к конкретной книге Excel. Запись запрещена.');
  necoRequireSameBook(url, expected);
}

// Shared, read-only verifier for the Broker POD module.
window.docUrl = docUrl;
window.necoBookIdentity = necoBookIdentity;
window.necoRequireSameBook = necoRequireSameBook;

async function readSheet(ctx, name) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
  ws.load('isNullObject,name');
  await ctx.sync();
  if (ws.isNullObject) throw Error('В книге нет листа «' + name + '».');
  ws.protection.load('protected,isPasswordProtected,isPaused');
  const used = ws.getUsedRange(true);
  used.load('rowIndex,rowCount');
  const head = ws.getRange('A1:M1');
  head.load('values');
  await ctx.sync();

  const h = head.values[0].map(v => P.norm(v));
  const okHead = h[0].includes('ДАТА ЗАЯВКИ') && h[2].includes('НОМЕР БУКИНГА') && h[3].includes('КОЛИЧЕСТВО') &&
                 h[4].includes('ТИП') && h[6].includes('РЕЛИЗ') && h[10] === 'POL';
  if (!okHead) throw Error('Заголовки листа не совпадают со схемой A:M. Запись заблокирована.');
  // Чтение и предварительная проверка разрешены и на защищённых листах.
  // На время ЗАПИСИ защита приостанавливается отдельно, без пароля, если это возможно.

  const last = used.rowIndex + used.rowCount;
  if (last < 3) throw Error('На листе почти нет данных: запись заблокирована.');
  const first = Math.max(2, last - WINDOW_ROWS + 1);
  const win = ws.getRange('A' + first + ':M' + last);
  win.load('values');
  await ctx.sync();
  return { ws, first, last, values: win.values };
}

// Число правил условного форматирования в столбце C (для контроля после копирования оформления).
async function cfCountC(name) {
  try {
    return await Excel.run(async ctx => {
      const c = ctx.workbook.worksheets.getItem(name).getRange('C:C').conditionalFormats.getCount();
      await ctx.sync();
      return c.value;
    });
  } catch (e) { return null; }
}

async function buildPlan() {
  checkEnv();
  return Excel.run(async ctx => {
    const st = await readSheet(ctx, payload.sheet);
    return P.plan(st.values, st.first, payload, new Date(), new Date());
  });
}

async function check() {
  showErr(''); el('result').textContent = ''; el('write').disabled = true; el('recheck').disabled = !payload;
  if (!payload) return;
  const pl = await buildPlan();
  shownPlanKey = planKey(pl);
  el('plan').textContent = 'Лист: ' + payload.sheet + '\n' + labelPlan(pl);
  el('write').disabled = pl.added === 0;
  if (pl.added === 0) el('result').textContent = 'Все строки уже есть в таблице. Записывать нечего.';
}

class StaleError extends Error {}

// Общая запись. expectedKey === null — автоматический режим (без сверки с показанным планом).
async function performWrite(expectedKey) {
  checkEnv();
  const today = new Date();
  const serial = P.excelSerial(today);
  const cfBefore = await cfCountC(payload.sheet);
  let out = null;

  await Excel.run(async ctx => {
    // Читаем заново прямо перед записью: за это время таблицу мог изменить коллега.
    const st = await readSheet(ctx, payload.sheet);
    const pl = P.plan(st.values, st.first, payload, today, today);
    if (expectedKey !== null && planKey(pl) !== expectedKey) {
      shownPlanKey = planKey(pl);
      el('plan').textContent = 'Лист: ' + payload.sheet + '\n' + labelPlan(pl);
      throw new StaleError('Таблица изменилась с момента проверки. Ничего не записано: проверьте новый план и нажмите «Записать» ещё раз.');
    }
    if (!pl.added) { out = { pl, warnings: [], written: 0 }; return; }

    const ws = st.ws;
    // Office.js / Excel Online: временно приостанавливаем защиту ТОЛЬКО в своей сессии.
    // Никакого глобального unprotect/protect, чтобы не открывать лист для других сотрудников.
    let pausedHere = false;
    let operationError = null;
    try {
      if (ws.protection.protected && !ws.protection.isPaused) {
        if (ws.protection.isPasswordProtected) {
          throw new Error('Лист защищён паролем. Автоматическая запись без пароля запрещена.');
        }
        if (typeof ws.protection.pauseProtection !== 'function') {
          throw new Error('Текущий Excel не поддерживает временную приостановку защиты листа. Запись не начата.');
        }
        ws.protection.pauseProtection();
        await ctx.sync();
        pausedHere = true;
      }
    const tpl = pl.gTemplate || pl.lastData;      // строка-образец оформления
    let gTpl = null;
    if (pl.gTemplate) {
      gTpl = ws.getRange('G' + pl.gTemplate);
      gTpl.format.fill.load('color');
      gTpl.format.load('horizontalAlignment');
      await ctx.sync();
    }
    const gFill = gTpl ? gTpl.format.fill.color : null;
    const gAlign = gTpl ? gTpl.format.horizontalAlignment : null;

    ws.activate();
    for (const r of pl.results) {
      if (r.kind !== 'new') continue;
      const n = r.row, it = r.item;
      if (COPY_ROW_FORMATS) {
        ws.getRange('A' + n + ':M' + n).copyFrom('A' + tpl + ':M' + tpl, Excel.RangeCopyType.formats);
      }
      const a = ws.getRange('A' + n);
      a.values = [[serial]];
      a.numberFormat = [[DATE_FORMAT]];
      ws.getRange('C' + n + ':F' + n).values = [[P.norm(it.booking), Number(it.quantity), P.norm(it.type), it.comment ? String(it.comment) : '']];
      const g = ws.getRange('G' + n);
      g.values = [['pending']];
      // Always use the book's standard pending fill, regardless of template/release color.
      g.format.fill.color = '#FFF2CC';
      if (gAlign && gAlign !== 'General') g.format.horizontalAlignment = gAlign;
      if (it.voyage) ws.getRange('H' + n).values = [[String(it.voyage)]];
      ws.getRange('K' + n).values = [[String(it.port).trim()]];
    }
    await ctx.sync();

    // Перечитываем записанное и сверяем.
    const checks = [];
    for (const r of pl.results) {
      if (r.kind !== 'new') continue;
      const rg = ws.getRange('A' + r.row + ':M' + r.row);
      rg.load('values');
      checks.push({ r, rg });
    }
    await ctx.sync();
    const problems = [];
    for (const { r, rg } of checks) {
      const v = rg.values[0], it = r.item;
      const ok = Number(v[0]) === serial && P.norm(v[2]) === P.norm(it.booking) && Number(v[3]) === Number(it.quantity) &&
                 P.norm(v[4]) === P.norm(it.type) && P.norm(v[6]) === 'PENDING' && P.norm(v[10]) === P.norm(it.port);
      if (!ok) problems.push(r.row);
    }
    if (problems.length) throw Error('Повторное чтение не совпало на строках: ' + problems.join(', ') + '. Возможно, одновременно писал коллега. Проверьте эти строки вручную.');
    out = { pl, warnings: [], written: pl.added };
    } catch (error) {
      operationError = error;
    } finally {
      if (pausedHere) {
        try {
          ws.protection.resumeProtection();
          await ctx.sync();
        } catch (restoreError) {
          const prefix = operationError ? (operationError.message + ' ') : '';
          operationError = new Error(prefix + 'Не удалось возобновить защиту листа. Часть строк МОГЛА записаться — проверьте книгу вручную. ' + restoreError.message);
        }
      }
    }
    if (operationError) throw operationError;
  });

  if (out && out.written && COPY_ROW_FORMATS) {
    const cfAfter = await cfCountC(payload.sheet);
    if (cfBefore !== null && cfAfter !== null && cfBefore !== cfAfter) {
      out.warnings.push('Изменилось число правил условного форматирования в столбце C: было ' + cfBefore + ', стало ' + cfAfter + '. Проверьте книгу и сообщите разработчику.');
    }
  }
  return out;
}

async function onPaste() {
  showErr(''); payload = null; shownPlanKey = '';
  el('write').disabled = true; el('recheck').disabled = true; el('result').textContent = '';
  const text = el('paste').value.trim();
  if (!text) { el('plan').textContent = 'Заявка не вставлена.'; return; }
  try {
    payload = JSON.parse(text);
    P.validatePayload(payload, new Date());
  } catch (e) {
    payload = null;
    el('plan').textContent = '';
    showErr(e instanceof SyntaxError ? 'В буфере не заявка NECO. Нажмите в программе «Копировать для Excel Online».' : e.message);
    return;
  }
  el('plan').textContent = 'Проверяю таблицу...';
  try { await check(); } catch (e) { el('plan').textContent = ''; showErr(e.message); }
}

async function write() {
  if (busy) return;
  busy = true; el('write').disabled = true; showErr('');
  try {
    if (!payload || !shownPlanKey) throw Error('Сначала вставьте заявку и дождитесь проверки.');
    const out = await performWrite(shownPlanKey);
    if (!out.written) throw Error('Все строки уже есть в таблице. Записывать нечего.');
    el('result').className = 'ok';
    el('result').textContent = 'Записано и перечитано.\n' + labelPlan(out.pl) +
      (out.warnings.length ? '\n\nВНИМАНИЕ:\n' + out.warnings.join('\n') : '') +
      '\n\nПроверьте в заголовке Excel, что стоит «Сохранено».';
    el('paste').value = ''; payload = null; shownPlanKey = '';
    el('recheck').disabled = true;
  } catch (e) {
    el('result').className = '';
    showErr(e.message);
  } finally {
    busy = false;
    el('write').disabled = !(payload && shownPlanKey);
  }
}

// ---- Автоматический режим: сообщения от расширения-курьера (content script в этом же фрейме) ----
function postResult(msg) { window.postMessage(Object.assign({ neco: 'RESULT' }, msg), location.origin); }

async function autoSubmit(m) {
  const id = String(m.id || '');
  if (handledJobs.has(id)) { postResult({id, ok:false, code:'REPLAY', message:'Повторная доставка одной и той же заявки. Проверьте таблицу.'}); return; }
  if (busy) { postResult({ id, ok: false, code: 'BUSY', message: 'Панель занята другой записью. Повторите через минуту.' }); return; }
  busy = true; showErr('');
  try {
    if (!/^[0-9a-f-]{36}$/i.test(id)) throw Error('Некорректный ID заявки');
    handledJobs.add(id);
    if (m.kind === 'broker-pod') {
      if (!window.NECO_BROKER_AUTO) throw Error('Broker POD module missing');
      const outcome=await window.NECO_BROKER_AUTO(m.payload);
      postResult({id,ok:true,written:outcome.written,rows:outcome.rows});
      return;
    }
    payload = m.payload; shownPlanKey = '';
    P.validatePayload(payload, new Date());
    el('paste').value = ''; el('write').disabled = true; el('recheck').disabled = true;
    el('plan').textContent = 'Получена заявка от программы...';
    const out = await performWrite(null);
    el('plan').textContent = 'Лист: ' + payload.sheet + '\n' + labelPlan(out.pl);
    el('result').className = out.written ? 'ok' : '';
    el('result').textContent = (out.written ? 'Записано автоматически и перечитано.' : 'Все строки уже есть в таблице.') +
      (out.warnings.length ? '\n\nВНИМАНИЕ:\n' + out.warnings.join('\n') : '');
    postResult({
      id, ok: true, written: out.written, duplicates: out.pl.duplicates, warnings: out.warnings,
      rows: out.pl.results.map(r => ({ booking: r.item.booking, port: r.item.port, quantity: Number(r.item.quantity), type: r.item.type, status: r.kind === 'new' ? 'written' : 'duplicate', row: r.row, why: r.why || '' }))
    });
  } catch (e) {
    showErr(e.message);
    postResult({ id, ok: false, code: 'ERROR', message: e.message });
  } finally {
    payload = null; shownPlanKey = ''; busy = false;
  }
}

let officeIsReady = false;
window.addEventListener('message', ev => {
  if (ev.source !== window || ev.origin !== location.origin) return;   // только от content script этого же фрейма
  const m = ev.data;
  if (!m || typeof m !== 'object') return;
  if (!officeIsReady) return;
  if (m.neco === 'PING') {
    let url = ''; try { url = docUrl(); } catch (e) { /* ignore */ }
    window.postMessage({ neco: 'PONG', id: String(m.id || ''), url, busy }, location.origin);
  } else if (m.neco === 'SUBMIT') {
    autoSubmit(m);
  }
});

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) { showErr('Откройте надстройку внутри Excel.'); return; }
  officeIsReady = true;
  el('env').textContent = 'Книга: ' + (docUrl() || '(адрес не получен)');
  el('paste').addEventListener('input', () => onPaste());
  el('recheck').onclick = () => check().catch(e => showErr(e.message));
  el('write').onclick = () => write();
  const hideButton = el('hide-panel');
  if (hideButton) hideButton.onclick = async () => {
    const label = el('hide-status');
    if (busy) { label.textContent = 'Идёт запись. Дождитесь подтверждения и повторите.'; return; }
    hideButton.disabled = true;
    try {
      if (!Office.context.requirements.isSetSupported('SharedRuntime', '1.1')) {
        throw new Error('В этом Excel не поддерживается SharedRuntime 1.1. Оставьте панель открытой.');
      }
      if (!Office.addin || typeof Office.addin.hide !== 'function') {
        throw new Error('Office.addin.hide() недоступен. Оставьте панель открытой.');
      }
      label.textContent = 'Скрываю только панель. Проверьте связь в расширении NECO Courier...';
      await Office.addin.hide();
    } catch (error) {
      hideButton.disabled = false;
      label.textContent = 'Не удалось скрыть панель: ' + (error.message || String(error));
    }
  };
});
