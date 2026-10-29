/* Independent Broker POD updater. Existing booking transport is untouched. */
(function () {
'use strict';
const $ = id => document.getElementById(id);
const norm = v => String(v == null ? '' : v).trim().replace(/\s+/g,' ').toUpperCase();
const port = v => {
  const s = norm(v).replace(/\s/g,'');
  const aliases = {QINGDAO:['QINGDAO','ЦИНДАО'],NANSHA:['NANSHA','НАНЬША'],NINGBO:['NINGBO','НИНБО','НИНГБО'],SHANGHAI:['SHANGHAI','ШАНХАЙ'],TIANJIN:['TIANJIN','ТЯНЬЦЗИНЬ'],DALIAN:['DALIAN','ДАЛЯНЬ'],RIZHAO:['RIZHAO','ЖИЧЖАО'],LIANYUNGANG:['LIANYUNGANG','ЛЯНЬЮНЬГАН','ЛЯНЬЮНЬГАНЬ']};
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
// Accept plain codes and names with a code after a slash, e.g. "JING YE / 126JY".
// Never guess when one source field contains multiple different voyage codes.
function extractVoyage(raw) {
  const text = String(raw == null ? '' : raw).normalize('NFKC').toUpperCase();
  const codes = [...text.matchAll(/(?:^|[^A-Z0-9])([0-9]{2,4}[A-Z]{1,4})(?=$|[^A-Z0-9])/g)].map(m=>m[1]);
  const unique = [...new Set(codes)];
  return {code:unique.length === 1 ? unique[0] : '', ambiguous:unique.length > 1};
}
function groupVoyage(records) {
  const codes=[];
  for (const r of records) {
    const actual=extractVoyage(r.voyage);
    if (actual.ambiguous) return {code:'',why:'Несколько рейсов в фактическом поле Брокера'};
    const planned=extractVoyage(r.plannedVoyage);
    if (planned.ambiguous && !actual.code) return {code:'',why:'Несколько плановых рейсов Брокера'};
    const chosen=actual.code || planned.code;
    if(chosen) codes.push(chosen);
  }
  const unique=[...new Set(codes)];
  if(unique.length > 1) return {code:'',why:'Разные рейсы у строк одного букинга'};
  return {code:unique[0]||'',why:unique.length ? '' : 'Рейс и плановый рейс отсутствуют'};
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
    if (!booking) continue;
    const podEligible=norm(v[6])==='PENDING' && !norm(v[11]);
    const voyageEligible=!norm(v[7]);
    if(!podEligible && !voyageEligible) continue;
    const type=containerType(v[4]);
    const records=groups.get(key(booking,v[10],type));
    if (!records) continue;
    const item={row:first+i, booking, type, city:'', voyage:'', podReady:false, voyageReady:false, status:'SKIP', why:''};
    if ((counts.get([booking,type].join('|'))||0)!==1) item.why='Несколько строк букинга и типа в Excel';
    else if (records.some(r=>norm(r.soc)!=='COC')) {item.status='SOC';item.why='SOC или смешанные SOC/COC';}
    else if (records.reduce((n,r)=>n+Number(r.qty),0)!==Number(v[3])) item.why='Количество не совпало';
    else {
      if (podEligible) {
        if (records.some(r=>!norm(r.city))) item.why='Нет города Drop Off';
        else if (new Set(records.map(r=>norm(r.city))).size!==records.length) {item.why='Возможный дубль города';}
        else {item.podReady=true; item.city=records.map(r=>String(r.city).trim()).join(' / ');}
      }
      if (voyageEligible) {
        const chosen=groupVoyage(records);
        if(chosen.code){item.voyageReady=true;item.voyage=chosen.code;}
        else if(!item.why) item.why=chosen.why;
      }
      if(item.podReady || item.voyageReady) item.status='READY';
      else if (item.why && (/Разные рейсы|Несколько.*рейсов/.test(item.why))) item.status='REVIEW';
    }
    result.push(item);
  }
  return result;
}
// Re-check the row just before the write; H and L are independent targets.
async function writeOne(ctx, ws, item, data) {
  const rg=ws.getRange('C'+item.row+':L'+item.row);rg.load('values');await ctx.sync();
  const v=rg.values[0];
  // C=0,D=1,E=2,G=4,H=5,K=8,L=9
  if(norm(v[0])!==item.booking || containerType(v[2])!==item.type) return {pod:false,voyage:false};
  const group=data.rows.filter(r=>key(r.booking,r.pol,r.type)===key(item.booking,v[8],item.type));
  if(!group.length || group.some(r=>norm(r.soc)!=='COC') || group.reduce((n,r)=>n+Number(r.qty),0)!==Number(v[1])) return {pod:false,voyage:false};
  const pod= item.podReady && norm(v[4])==='PENDING' && !norm(v[9]);
  const voyage= item.voyageReady && !norm(v[5]);
  if (!pod && !voyage) return {pod:false,voyage:false};
  if (pod) ws.getRange('L'+item.row).values=[[item.city]];
  if (voyage) ws.getRange('H'+item.row).values=[[item.voyage]];
  await ctx.sync();
  const checks=[];
  if (pod) {const r=ws.getRange('L'+item.row);r.load('values');checks.push(['L',r,item.city]);}
  if (voyage) {const r=ws.getRange('H'+item.row);r.load('values');checks.push(['H',r,item.voyage]);}
  await ctx.sync();
  for (const [col,r,wanted] of checks) if (String(r.values[0][0])!==wanted) throw Error('Не подтверждена запись '+col+item.row);
  return {pod,voyage};
}
function validateBook(data) {
  if (!data || !['main','test'].includes(data.workbookMode)) {
    throw Error('Broker POD: режим книги не указан. Запись запрещена.');
  }
  const url = String(Office.context.document.url || '').trim();
  if (!url) throw Error('Broker POD: Excel не сообщил адрес книги. Запись запрещена.');
  const canonical = raw => {
    const u = new URL(raw);
    return u.origin.toLowerCase() + decodeURIComponent(u.pathname)
      .normalize('NFC').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ')
      .replace(/\/$/, '').toUpperCase();
  };
  if (data.workbookMode === 'main') {
    const expected = String(data.expectedWorkbookUrl || '').trim();
    if (!expected.startsWith('https://') || !url.startsWith('https://'))
      throw Error('Broker POD: отсутствует защищённая привязка основной книги.');
    if (canonical(url) !== canonical(expected))
      throw Error('Broker POD: открыта другая книга. Ожидалась: ' + expected + '; открыта: ' + url);
    if (/ПРОБА/i.test(new URL(url).pathname))
      throw Error('Broker POD: тестовая книга не разрешена в основном режиме.');
  } else {
    const filename = decodeURIComponent(new URL(url).pathname).split('/').filter(Boolean).pop() || '';
    if (!/ПРОБА/i.test(filename) || !/\.xlsx?$/i.test(filename))
      throw Error('Broker POD: тестовая запись разрешена только в ПРОБА.xlsx. Открыто: ' + url);
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
    if (!h[2].includes('БУКИНГ') || !h[3].includes('КОЛИЧЕСТВО') || !h[4].includes('ТИП') || !h[6].includes('РЕЛИЗ') || !h[7].includes('РЕЙС') || h[10]!=='POL' || h[11]!=='POD') throw Error('Заголовки Excel A:M не соответствуют схеме.');
    const last=used.rowIndex+used.rowCount, first=Math.max(2,last-499);
    const range=ws.getRange('A'+first+':M'+last);range.load('values');await ctx.sync();
    return {first,values:range.values};
  });
}
function resultText(rows) {
  return rows.map(x=>x.booking+' / '+x.type+' / '+x.row+': '
    +(x.podReady ? 'POD='+x.city+' ' : '')+(x.voyageReady ? 'H='+x.voyage : '')
    +((!x.podReady&&!x.voyageReady) ? x.why : '')).join('\n');
}
async function check() {
  if(running) return;
  running=true; $('broker-write').disabled=true;
  try {
    current=parse(); const snap=await snapshot(current);
    preview=calc(current,snap.values,snap.first);
    const ready=preview.filter(x=>x.status==='READY');
    const podCount=ready.filter(x=>x.podReady).length, voyageCount=ready.filter(x=>x.voyageReady).length;
    display('Готово к записи POD (L): '+podCount+'; рейсов (H): '+voyageCount+'; пропущено: '+(preview.length-ready.length)+'\n'+resultText(preview));
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
    const output=await Excel.run(async ctx=>{
      validateBook(current);
      const ws=ctx.workbook.worksheets.getItem(current.sheet);
      ws.protection.load('protected,isPasswordProtected,isPaused');await ctx.sync();
      let paused=false, err=null, pods=0, voyages=0;
      try {
        if(ws.protection.protected&&!ws.protection.isPaused) {
          if(ws.protection.isPasswordProtected) throw Error('Лист защищён паролем. Запись запрещена.');
          if(typeof ws.protection.pauseProtection !== 'function') throw Error('Excel не поддерживает приостановку защиты.');
          ws.protection.pauseProtection(); await ctx.sync(); paused=true;
        }
        for(const item of approved) {
          const wrote=await writeOne(ctx,ws,item,current);
          if(wrote.pod) pods++;
          if(wrote.voyage) voyages++;
        }
      } catch(e) {err=e;} finally {
        if(paused) try {ws.protection.resumeProtection();await ctx.sync();} catch(e){err=Error('Не удалось восстановить защиту листа. Проверьте Excel: '+e.message);}
      }
      if(err) throw err;
      return {pods,voyages};
    });
    display('Записано и проверено: POD (L) '+output.pods+'; рейсов (H) '+output.voyages+'. Остальные ячейки не изменены.');preview=null;
  } catch(e) {display('Ошибка: '+e.message+'; проверьте уже записанные ячейки.',true);} finally {running=false;}
}
async function automatic(data) {
  if (running) throw Error('Broker POD занят');
  running = true;
  try {
    if (data.app !== 'NECO_BROKER_POD' || data.schema !== 1 || !Array.isArray(data.rows) || data.rows.length > 20000) throw Error('Неверный Broker payload');
    const snap=await snapshot(data);
    const rows=calc(data,snap.values,snap.first);
    const approved=rows.filter(x=>x.status==='READY');
    const written=[];let podCount=0, voyageCount=0;
    await Excel.run(async ctx=>{
      validateBook(data);
      const ws=ctx.workbook.worksheets.getItem(data.sheet);
      ws.protection.load('protected,isPasswordProtected,isPaused'); await ctx.sync();
      let pausedHere=false,opErr=null;
      try {
        if(ws.protection.protected && !ws.protection.isPaused) {
          if(ws.protection.isPasswordProtected) throw Error('Лист защищён паролем. Запись запрещена.');
          if(typeof ws.protection.pauseProtection !== 'function') throw Error('Excel не поддерживает приостановку защиты.');
          ws.protection.pauseProtection(); await ctx.sync(); pausedHere=true;
        }
        for(const item of approved) {
          const done=await writeOne(ctx,ws,item,data);
          if(done.pod || done.voyage) written.push({...item,podWritten:done.pod,voyageWritten:done.voyage});
          if(done.pod) podCount++;
          if(done.voyage) voyageCount++;
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
      const reasons=rows.length ? rows.slice(0,20).map(r=>r.booking+' '+r.status+' '+(r.why||'')).join('; ') : 'Совпадений по booking + POL + type среди последних 500 строк не найдено';
      throw Error('Broker POD/Voyage: записано 0. Кандидатов: '+rows.length+', готовых: '+approved.length+'. '+reasons+'; Excel: '+data.sheet+'. Проверьте POL, тип, количество, H и L.');
    }
    return {written:written.length,podWritten:podCount,voyageWritten:voyageCount,
      rows:rows.map(r=>({booking:r.booking,type:r.type,row:r.row,status:written.some(w=>w.row===r.row)?'written':r.status.toLowerCase(),why:r.why||''})),
      diagnostic:'POD (L): '+podCount+'; рейсов (H): '+voyageCount+'; кандидатов: '+rows.length+'; готовых: '+approved.length};
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
