/* Independent Broker POD updater. Existing booking transport is untouched. */
(function () {
'use strict';
const $ = id => document.getElementById(id);
const norm = v => String(v == null ? '' : v).trim().replace(/\s+/g,' ').toUpperCase();
const port = v => {
  const s = norm(v).replace(/\s/g,'');
  const aliases = {QINGDAO:['QINGDAO','ЦИНДАО'],NANSHA:['NANSHA','НАНЬША'],NINGBO:['NINGBO','НИНБО','НИНГБО'],SHANGHAI:['SHANGHAI','ШАНХАЙ'],TIANJIN:['TIANJIN','ТЯНЬЦЗИНЬ'],DALIAN:['DALIAN','ДАЛЯНЬ'],RIZHAO:['RIZHAO','ЖИЧЖАО']};
  return Object.keys(aliases).find(k => aliases[k].includes(s)) || s;
};
const containerType = v => norm(v).replace(/[\s'’′`“”"]/g,'');
const key = (b,p,t) => [norm(b),port(p),containerType(t)].join('|');
let current = null, preview = null, running = false;
function display(s, bad=false) { $('broker-status').textContent = s; $('broker-status').className = bad?'err':''; }
function parse() {
  const data = JSON.parse($('broker-paste').value);
  if (data.app !== 'NECO_BROKER_POD' || data.schema !== 1 || !Array.isArray(data.rows)) throw Error('Неверный формат файла Broker POD.');
  if (!data.sheet || !Array.isArray(data.rows) || data.rows.length > 20000) throw Error('Не указан лист или слишком много строк.');
  if (!['QINGDAO, NANSHA, OUTPORTS','TIANJIN, DALIAN','SHANGHAI, NINGBO','QINGDAO, NANSHA','SHANGHAI, NINGBO, OUTPORTS'].includes(data.sheet)) throw Error('Неизвестный лист.');
  for (const r of data.rows) {
    if (!r.booking || !r.pol || !r.type || !Number.isInteger(Number(r.qty)) || Number(r.qty)<1 || !['COC','SOC'].includes(norm(r.soc))) throw Error('Неполная или неверная запись Брокера: '+JSON.stringify(r));
  }
  return data;
}
function calc(data, values, first) {
  const groups = new Map();
  for (const r of data.rows) {
    const k = key(r.booking,r.pol,r.type);
    if (!groups.has(k)) groups.set(k,[]);
    groups.get(k).push(r);
  }
  const counts = new Map();
  for (const v of values) {
    if (!norm(v[2])) continue;
    const k = [norm(v[2]),containerType(v[4])].join('|');
    counts.set(k,(counts.get(k)||0)+1);
  }
  const result=[];
  for (let i=values.length-1;i>=0;i--) {
    const v=values[i], booking=norm(v[2]);
    if (!booking || norm(v[6])!=='PENDING' || norm(v[11])) continue;
    const type=containerType(v[4]);
    const records=groups.get(key(booking,v[10],type));
    if (!records) continue;
    const item={row:first+i, booking, type, city:'', status:'SKIP', why:''};
    if ((counts.get([booking,type].join('|'))||0)!==1) item.why='Несколько строк букинга и типа в Excel';
    else if (records.some(r=>norm(r.soc)!=='COC')) {item.status='SOC';item.why='SOC или смешанные SOC/COC';}
    else if (records.some(r=>!norm(r.city))) item.why='Нет города Drop Off';
    else if (new Set(records.map(r=>norm(r.city))).size!==records.length) {item.status='REVIEW';item.why='Возможный дубль города';}
    else if (records.reduce((n,r)=>n+Number(r.qty),0)!==Number(v[3])) item.why='Количество не совпало';
    else {item.status='READY';item.city=records.map(r=>String(r.city).trim()).join(' / ');}
    result.push(item);
  }
  return result;
}
function validateBook(data) {
  // V17 TEST ONLY. Reuse the exact existing Outlook->Excel test check.
  // The shared function is defined by pane-hidden.js and is invoked at runtime,
  // after both deferred scripts have loaded. Never fall back to URL identity.
  if (!data || data.workbookMode === 'main') {
    throw Error('Broker POD: разрешён только тестовый режим ПРОБА. Запись запрещена.');
  }
  if (typeof window.checkEnv !== 'function' || typeof window.docUrl !== 'function') {
    throw Error('Broker POD: общая проверка книги недоступна. Запись запрещена.');
  }
  // The original checkEnv() uses its own payload for mode selection.
  // For Broker we explicitly enforce test mode, and check the same book
  // URL using the existing original test criterion without changing it.
  const url = window.docUrl();
  if (!url || !url.toUpperCase().includes('ПРОБА')) {
    throw Error('Broker POD: тестовая запись разрешена только в книге ПРОБА. Открыто: ' + (url || '(URL отсутствует)'));
  }
  // The original Outlook checkEnv uses the same predicate in test mode.
  // Require the book name itself to contain ПРОБА (not just its folder).
  let pathname;
  try { pathname = decodeURIComponent(new URL(url).pathname); }
  catch (_) { throw Error('Broker POD: некорректный адрес открытой книги. Запись запрещена.'); }
  const filename = pathname.split('/').filter(Boolean).pop() || '';
  if (!/ПРОБА/i.test(filename) || !/\.xlsx?$/i.test(filename)) {
    throw Error('Broker POD: в имени открытой книги нет ПРОБА.xlsx. Запись запрещена. Имя: ' + filename);
  }
}

async function snapshot(data) {
  validateBook(data);
  return Excel.run(async ctx=>{
    const ws=ctx.workbook.worksheets.getItemOrNullObject(data.sheet);
    ws.load('isNullObject'); await ctx.sync();
    if(ws.isNullObject) throw Error('Не найден лист '+data.sheet);
    const header=ws.getRange('A1:M1'); header.load('values');
    const used=ws.getUsedRange(true);used.load('rowIndex,rowCount');
    await ctx.sync();
    const h=header.values[0].map(norm);
    if (!h[2].includes('БУКИНГ') || !h[3].includes('КОЛИЧЕСТВО') || !h[4].includes('ТИП') || !h[6].includes('РЕЛИЗ') || h[10]!=='POL' || h[11]!=='POD') throw Error('Заголовки Excel A:M не соответствуют схеме.');
    const last=used.rowIndex+used.rowCount, first=Math.max(2,last-499);
    const range=ws.getRange('A'+first+':M'+last);range.load('values');await ctx.sync();
    return {first,values:range.values};
  });
}
async function check() {
  if(running) return;
  running=true; $('broker-write').disabled=true;
  try {
    current=parse(); const snap=await snapshot(current);
    preview=calc(current,snap.values,snap.first);
    const ready=preview.filter(x=>x.status==='READY');
    display('Готово к записи: '+ready.length+'; пропущено/проверить: '+(preview.length-ready.length)+'\n'+preview.map(x=>x.booking+' / '+x.type+' / '+x.row+': '+(x.city||x.why)).join('\n'));
    $('broker-write').disabled=!ready.length;
  } catch(e) {preview=null;display(e.message,true);} finally {running=false;}
}
async function write() {
  if(running || !preview || !current) return;
  running=true;$('broker-write').disabled=true;
  try {
    const snap=await snapshot(current);
    const latest=calc(current,snap.values,snap.first);
    const approved=latest.filter(x=>x.status==='READY');
    // Revalidate every target cell in a fresh Excel.run immediately before writing.
    const output=await Excel.run(async ctx=>{
      validateBook(current);
      const ws=ctx.workbook.worksheets.getItem(current.sheet);
      ws.protection.load('protected,isPasswordProtected,isPaused');await ctx.sync();
      let paused=false, err=null, written=0;
      try {
        if(ws.protection.protected&&!ws.protection.isPaused) {
          if(ws.protection.isPasswordProtected) throw Error('Лист защищён паролем. Запись запрещена.');
          if(typeof ws.protection.pauseProtection !== 'function') throw Error('Excel не поддерживает приостановку защиты.');
          ws.protection.pauseProtection(); await ctx.sync(); paused=true;
        }
        for(const item of approved) {
          const rg=ws.getRange('C'+item.row+':L'+item.row);rg.load('values');await ctx.sync();
          const v=rg.values[0];
          // C=0,D=1,E=2,G=4,K=8,L=9
          if(norm(v[0])!==item.booking || containerType(v[2])!==item.type || norm(v[4])!=='PENDING' || norm(v[9])) continue;
          const group=current.rows.filter(r=>key(r.booking,r.pol,r.type)===key(item.booking,v[8],item.type));
          if(group.reduce((n,r)=>n+Number(r.qty),0)!==Number(v[1])) continue;
          ws.getRange('L'+item.row).values=[[item.city]];
          await ctx.sync();
          const verify=ws.getRange('L'+item.row);verify.load('values');await ctx.sync();
          if(verify.values[0][0]!==item.city) throw Error('Не удалось подтвердить запись L'+item.row);
          written++;
        }
      } catch(e) {err=e;} finally {
        if(paused) try {ws.protection.resumeProtection();await ctx.sync();} catch(e){err=Error('Не удалось восстановить защиту листа. Проверьте Excel: '+e.message);}
      }
      if(err) throw err;
      return written;
    });
    display('Записано и проверено: '+output+'. Остальные строки не изменены.');preview=null;
  } catch(e) {display('Ошибка: '+e.message+'; проверьте уже записанные ячейки.',true);} finally {running=false;}
}
async function automatic(data) {
  if (running) throw Error('Broker POD занят');
  running = true;
  try {
    if (data.app !== 'NECO_BROKER_POD' || data.schema !== 1 || !Array.isArray(data.rows) || data.rows.length > 20000) throw Error('Неверный Broker payload');
    const snap = await snapshot(data);
    const rows = calc(data, snap.values, snap.first);
    const approved = rows.filter(x=>x.status==='READY');
    const written=[];
    await Excel.run(async ctx=>{
      validateBook(data);
      const ws=ctx.workbook.worksheets.getItem(data.sheet);
      ws.protection.load('protected,isPasswordProtected,isPaused'); await ctx.sync();
      let pausedHere=false, opErr=null;
      try {
        if(ws.protection.protected && !ws.protection.isPaused) {
          if(ws.protection.isPasswordProtected) throw Error('Лист защищён паролем. Запись запрещена.');
          if(typeof ws.protection.pauseProtection !== 'function') throw Error('Excel не поддерживает приостановку защиты.');
          ws.protection.pauseProtection(); await ctx.sync(); pausedHere=true;
        }
      for(const item of approved) {
        const rg=ws.getRange('C'+item.row+':L'+item.row);rg.load('values');await ctx.sync();
        const v=rg.values[0];
        if(norm(v[0])!==item.booking || containerType(v[2])!==item.type || norm(v[4])!=='PENDING' || norm(v[9])) continue;
        const group=data.rows.filter(r=>key(r.booking,r.pol,r.type)===key(item.booking,v[8],item.type));
        if(group.length===0 || group.some(r=>norm(r.soc)!=='COC') || group.reduce((n,r)=>n+Number(r.qty),0)!==Number(v[1])) continue;
        ws.getRange('L'+item.row).values=[[item.city]];await ctx.sync();
        const verify=ws.getRange('L'+item.row);verify.load('values');await ctx.sync();
        if(String(verify.values[0][0])!==item.city) throw Error('Не подтверждена запись L'+item.row);
        written.push(item);
      }
      } catch(e) {opErr=e;} finally {
        if(pausedHere) {
          try {ws.protection.resumeProtection(); await ctx.sync();}
          catch(e) {opErr=Error((opErr ? opErr.message+'; ' : '')+'Не удалось восстановить защиту: '+e.message);}
        }
      }
      if(opErr) throw opErr;
    });
    if (!written.length) {
      const reasons = rows.length ? rows.slice(0,20).map(r => r.booking+' '+r.status+' '+(r.why||'')).join('; ') : 'Совпадений по booking + POL + type среди последних 500 строк не найдено';
      throw Error('Broker POD: записано 0. Кандидатов: '+rows.length+', готовых: '+approved.length+'. '+reasons+'; Excel: '+data.sheet+'. Проверьте POL, тип, количество, pending и пустой L.');
    }
    return {written:written.length, rows:rows.map(r=>({booking:r.booking,type:r.type,row:r.row,status:written.some(w=>w.row===r.row)?'written':r.status.toLowerCase(),why:r.why||''})), diagnostic: rows.length ? ('Найдено кандидатов: '+rows.length+'; готовых: '+approved.length+'; записано: '+written.length+'; примеры: '+rows.slice(0,12).map(r=>r.booking+' '+r.status+' '+(r.why||'')).join('; ')) : 'Не найдено совпадений по букингу + POL + типу среди последних 500 строк Excel. Проверьте лист, тип и POL.'};
  } finally {running=false;}
}
window.NECO_BROKER_AUTO = automatic;
Office.onReady(info=>{
  if(info.host!==Office.HostType.Excel) return;
  $('broker-check').onclick=check;
  $('broker-write').onclick=write;
  $('broker-paste').addEventListener('input',()=>{$('broker-write').disabled=true;preview=null;});
});
})();
