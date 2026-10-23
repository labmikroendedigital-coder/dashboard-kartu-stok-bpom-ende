const CFG={CACHE:'dashboard_v10',TTL:300,CHUNK:20000};

function doGet(e){
  // API JSONP untuk dashboard GitHub Pages; tanpa parameter api tetap tampilkan Web App lama.
  if (e && e.parameter && e.parameter.api === '1') {
    const p = e.parameter || {};
    const callback = String(p.callback || '');
    if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
      return ContentService.createTextOutput('/* callback JSONP tidak valid */')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    let result;
    try {
      const action = String(p.action || 'dashboard');
      if (action === 'dashboard') {
        result = getDashboardData();
      } else if (action === 'refresh') {
        result = refreshDashboardData();
      } else if (action === 'transactions') {
        result = getTransactions(String(p.sheetName || ''));
      } else {
        throw new Error('Aksi API tidak dikenal: ' + action);
      }
    } catch (err) {
      result = { error: err && err.message ? err.message : String(err) };
    }
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dashboard Kartu Stok Media dan Reagen');
}

function getDashboardData(){
  const cached=cacheRead_();
  if(cached) return cached;
  const data=buildDashboardData_();
  cacheWrite_(data);
  return data;
}

function buildDashboardData_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error('Spreadsheet tidak ditemukan.');
  const today=new Date(); today.setHours(0,0,0,0);
  const items=[];
  ss.getSheets().forEach(sh=>{
    const last=sh.getLastRow();
    if(last<13) return;
    const sourceRange=sh.getRange(9,1,Math.max(last-8,5),12);
    const v=sourceRange.getValues();
    const display=sourceRange.getDisplayValues();
    if(!String(v[0][4]||'').trim() || String(v[4][2]||'').toLowerCase().indexOf('tanggal')<0) return;

    const rows=v.slice(5);
    const displayRows=display.slice(5);
    let stock=0,lastDate=null,hasTxn=false;
    let activeBatches=[];

    rows.forEach((r,i)=>{
      const d=parseDate_(r[2]);
      if(d&&(!lastDate||d>lastDate)) lastDate=d;
      if(r.some(x=>String(x??'').trim()!=='')){stock=number_(r[7]);hasTxn=true;}

      const shown=displayRows[i]||[];
      const batchText=String(shown[4]??'').trim();
      const edText=String(shown[8]??'').trim();
      const batchEntries=parseBatchEntries_(batchText);
      if(batchEntries.length){
        const edEntries=parseEdEntries_(edText);
        activeBatches=batchEntries.map(b=>{
          const ed=edEntries.find(e=>e.number===b.number) || edEntries[batchEntries.indexOf(b)] || (edEntries.length===1 && edEntries[0].value==='-' ? edEntries[0] : null);
          return {number:b.number,batch:b.value,expiry:ed?ed.value:'',rowNumber:14+i};
        });
      }
    });

    const activeDatePairs=activeBatches.map(b=>({batch:b,date:parseDate_(b.expiry)})).filter(x=>x.date);
    const futurePairs=activeDatePairs.filter(x=>x.date.getTime()>=today.getTime()).sort((a,b)=>a.date-b.date);
    const expiredPairs=activeDatePairs.filter(x=>x.date.getTime()<today.getTime()).sort((a,b)=>b.date-a.date);
    const orderedDates=futurePairs.concat(expiredPairs).map(x=>x.date);
    let status='Tidak ada data ED';
    if(futurePairs.length){
      const daysLeft=days_(today,futurePairs[0].date);
      status=daysLeft<=90?'Segera kedaluwarsa':'Masih berlaku';
    } else if(expiredPairs.length) status='Kedaluwarsa';
    items.push({sheetName:sh.getName(),name:String(v[0][4]||'').trim(),
      unit:String(v[1][4]||'').trim(),manufacturer:String(v[2][4]||'').trim(),
      stock:hasTxn?stock:0,expiry:orderedDates.length ? orderedDates.map(formatDate_).join('\n') : (activeBatches.some(b=>String(b.expiry||'').trim()==='-') ? '-' : ''),
      lastTransaction:lastDate?formatDate_(lastDate):'-',status});
  });
  items.sort((a,b)=>a.name.localeCompare(b.name));
  return {spreadsheetName:ss.getName(),generatedAt:new Date().toISOString(),items,
    summary:{total:items.length,
      expired:items.filter(x=>x.status==='Kedaluwarsa'||x.status==='Sebagian kedaluwarsa').length,
      nearExpiry:items.filter(x=>x.status==='Segera kedaluwarsa').length,
      zeroStock:items.filter(x=>x.stock<=0).length}};
}
function getTransactions(sheetName){
  if(!sheetName) throw new Error('Nama sheet tidak diberikan.');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName(sheetName);
  if(!sh) throw new Error('Kartu stok tidak ditemukan: '+sheetName);
  const info=sh.getRange('E9:E11').getDisplayValues();
  const out={sheetName:sh.getName(),name:info[0][0]||sheetName,unit:info[1][0]||'',manufacturer:info[2][0]||'',batches:[],transactions:[]};
  const last=sh.getLastRow(); if(last<14) return out;
  const values=sh.getRange(14,1,last-13,12).getDisplayValues();
  let latestBatchRow=null;
  values.forEach((r,i)=>{
    const rowNumber=14+i;
    const detail=String(r[4]||'').trim();
    const edText=String(r[8]||'').trim();
    const batchEntries=parseBatchEntries_(detail);
    const edEntries=parseEdEntries_(edText);
    if(batchEntries.length){
      const mapped=batchEntries.map(b=>{
        const ed=edEntries.find(e=>e.number===b.number) || edEntries[batchEntries.indexOf(b)] || (edEntries.length===1 && edEntries[0].value==='-' ? edEntries[0] : null);
        return {number:b.number,batch:b.value,expiry:ed?ed.value:'',rowNumber:rowNumber};
      });
      latestBatchRow=mapped;
    }
    if(r.some(x=>String(x||'').trim()!=='')) out.transactions.push({
      rowNumber:rowNumber,date:r[2],batch:detail,description:detail,
      outgoing:r[5],incoming:r[6],balance:r[7],expiry:edText,signature:r[11]
    });
  });
  out.batches=(latestBatchRow||[]).sort((a,b)=>a.number-b.number);
  return out;
}
function parseBatchEntries_(text){
  const raw=String(text||'').trim();
  if(!raw || raw==='-' ) return [];
  const lines=raw.split(/[\n\r;]+/).map(x=>x.trim()).filter(Boolean);
  const result=[];
  let nextNumber=1;
  lines.forEach(line=>{
    let number=null, body=line;
    let m=body.match(/^\s*(\d{1,3})\s*[.)\-:,]\s*(.*?)\s*$/);
    if(m){number=Number(m[1]);body=m[2].trim();}
    else {
      m=body.match(/^\s*(?:batch|lot)\s*(\d{1,3})\s*[:.)\-]\s*(.*?)\s*$/i);
      if(m){number=Number(m[1]);body=m[2].trim();}
    }
    if(!body || body==='-') return;
    let placeholderBatch=false;
    if(body.includes('/')) {
      body=body.slice(body.indexOf('/')+1).trim();
      placeholderBatch=(body==='-');
    } else if(/^kat\s*[:.]?/i.test(body)) {
      return;
    }
    if(!body || (body==='-' && !placeholderBatch)) return;
    if(/^kat\s*[:.]?/i.test(body)) return;
    const n=number || nextNumber;
    result.push({number:n,value:body});
    nextNumber=Math.max(nextNumber,n+1);
  });
  const seen={};
  return result.map((e,i)=>{
    if(seen[e.number]) e.number=Math.max(nextNumber,i+1);
    seen[e.number]=true; return e;
  });
}
function parseEdEntries_(text){
  if(text===null || text===undefined) return [];
  if(text instanceof Date && !isNaN(text)) return [{number:1,value:formatDate_(text)}];
  const raw=String(text).trim();
  if(!raw) return [];
  if(raw==='-') return [{number:1,value:'-'}];
  const numbered=parseNumberedLines_(raw,false);
  if(numbered.length) return numbered.map(e=>{
    const d=parseDate_(e.value);
    return {number:e.number,value:d?formatDate_(d):e.value};
  });
  const direct=parseDate_(raw);
  if(direct) return [{number:1,value:formatDate_(direct)}];
  const parts=raw.split(/[\n\r;]+/).map(x=>x.trim()).filter(Boolean);
  const dates=[];
  parts.forEach(part=>{
    const cleaned=part.replace(/^\s*(?:(?:ed|batch)\s*)?\d{1,3}\s*[:.)\-]\s*/i,'').trim();
    const d=parseDate_(cleaned) || parseDate_(part);
    if(d) dates.push({number:dates.length+1,value:formatDate_(d)});
  });
  return dates;
}
function parseNumberedLines_(text, isBatch){
  if(!text) return [];
  const result=[];
  const lines=String(text).split(/[\n\r;]+/).map(s=>s.trim()).filter(Boolean);
  lines.forEach(line=>{
    if(isBatch && /^kat\.?\s*/i.test(line)) return;
    let m=line.match(/^\s*(\d{1,3})\s*[.)\-:]\s*(.*?)\s*$/);
    if(!m) m=line.match(/^\s*(?:batch|ed)\s*(\d{1,3})\s*[:.)\-]\s*(.*?)\s*$/i);
    if(!m) return;
    const number=Number(m[1]);
    let value=String(m[2]||'').trim();
    if(!number || !value) return;
    if(isBatch){
      if(/^kat\.?\s*/i.test(value)) return;
      value=value.replace(/^\/+\s*/, '').trim();
    }
    result.push({number:number,value:value});
  });
  return result;
}
function cacheRead_(){
  const c=CacheService.getScriptCache(),p=CFG.CACHE;
  try{
    const n=Number(c.get(p+'_n')||0); if(!Number.isInteger(n)||n<1||n>100)return null;
    const keys=Array.from({length:n},(_,i)=>p+'_'+i),parts=c.getAll(keys);
    let s=''; for(const k of keys){if(parts[k]===undefined)return null;s+=parts[k];}
    return JSON.parse(s);
  }catch(e){console.warn(e);return null;}
}
function cacheWrite_(data){
  const c=CacheService.getScriptCache(),p=CFG.CACHE;
  try{
    const s=JSON.stringify(data),n=Math.ceil(s.length/CFG.CHUNK);
    if(n<1||n>100)return;
    c.remove(p+'_n');
    const obj={};for(let i=0;i<n;i++)obj[p+'_'+i]=s.slice(i*CFG.CHUNK,(i+1)*CFG.CHUNK);
    c.putAll(obj,CFG.TTL);c.put(p+'_n',String(n),CFG.TTL);
  }catch(e){console.warn('Cache gagal disimpan: '+e.message);}
}
function onEdit(e){
  try { clearDashboardCache(); } catch(err) { console.warn(err); }
}
function clearDashboardCache(){
  const c=CacheService.getScriptCache(),p=CFG.CACHE,n=Number(c.get(p+'_n')||0);
  c.remove(p+'_n');
  if(n>0&&n<=100)c.removeAll(Array.from({length:n},(_,i)=>p+'_'+i));
}
function refreshDashboardData(){
  clearDashboardCache();const d=buildDashboardData_();cacheWrite_(d);return d;
}
function refreshDashboardCache(){
  const d=refreshDashboardData();return 'Cache diperbarui. Total kartu: '+d.items.length;
}
function testDashboard(){const d=buildDashboardData_();console.log(d.summary);return d.summary;}
function number_(v){if(typeof v==='number')return v;const n=Number(String(v??'').replace(',','.'));return Number.isFinite(n)?n:0;}
function parseDate_(v){
  if(!v || v==='-')return null;
  if(v instanceof Date&&!isNaN(v)){const d=new Date(v);d.setHours(0,0,0,0);return d;}
  const s=String(v).trim();let m=s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if(m)return validDate_(+m[1],+m[2],+m[3]);
  m=s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if(m)return validDate_(+m[3],+m[2],+m[1]);
  const months={januari:1,jan:1,februari:2,feb:2,maret:3,mar:3,april:4,apr:4,mei:5,juni:6,jun:6,juli:7,jul:7,agustus:8,agu:8,agt:8,september:9,sep:9,oktober:10,okt:10,november:11,nov:11,desember:12,des:12};
  m=s.toLowerCase().match(/^(\d{1,2})\s+([a-z]+)\s+(\d{4})$/);
  if(m&&months[m[2]])return validDate_(+m[3],months[m[2]],+m[1]);
  m=s.toLowerCase().replace(/\.$/,'').match(/^([a-z]+)\s+(\d{4})$/);
  if(m&&months[m[1]]){const d=new Date(+m[2],months[m[1]],0);d.setHours(0,0,0,0);return d;}
  return null;
}
function validDate_(y,m,d){const x=new Date(y,m-1,d);if(x.getFullYear()!==y||x.getMonth()!==m-1||x.getDate()!==d)return null;x.setHours(0,0,0,0);return x;}
function formatDate_(d){const m=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];return d.getDate()+' '+m[d.getMonth()]+' '+d.getFullYear();}
function splitNumberedText_(value){
  if(value===null || value===undefined || value==='') return [];
  if(value instanceof Date) return [{number:1,value:value}];
  const text=String(value).trim();
  if(!text) return [];
  const direct=parseDate_(text);
  if(direct) return [{number:1,value:direct}];
  const parts=text.split(/[\n;]+/).map(x=>x.trim()).filter(Boolean);
  const result=[];
  parts.forEach((part,index)=>{
    const cleaned=part.replace(/^(?:(?:batch|ed)\s*)?\d+\s*[:.)\-]\s*/i,'').trim();
    const date=parseDate_(cleaned) || parseDate_(part);
    if(date) result.push({number:index+1,value:date});
  });
  return result;
}
function uniqueDates_(a){const seen={};return a.filter(d=>{const k=d.getFullYear()+'-'+d.getMonth()+'-'+d.getDate();if(seen[k])return false;seen[k]=true;return true;});}
function days_(a,b){return Math.floor((b-a)/86400000);}
