import * as XLSX from 'xlsx';

const n = v => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const x = Number(String(v).replace(/[$,]/g,'').trim());
  return Number.isFinite(x) ? x : 0;
};
const norm = v => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase();
const fileDate = name => {
  const m = String(name).match(/^(\d{2})(\d{2})(\d{4})_/);
  return m ? `${m[3]}-${m[2]}-${m[1]}` : new Date().toISOString().slice(0,10);
};
const toISO = v => {
  if (!v) return null;
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toISOString().slice(0,10);
  if (typeof v === 'number' && v > 20000 && v < 100000) {
    const d = XLSX.SSF.parse_date_code(v);
    return d ? `${String(d.y).padStart(4,'0')}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}` : null;
  }
  const s = String(v).trim();
  let m=s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if(m)return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;
  m=s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if(m)return `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  return null;
};
const sheets = wb => wb.SheetNames.map(name => ({ name, rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, raw:true, defval:null }) }));
const catPush = (map, category, value) => {
  const c = norm(category).replace(/\s+/g,' ');
  if (!c || !value) return;
  map.set(c, (map.get(c)||0)+value);
};
const mapCats = m => [...m.entries()].map(([category,value])=>({category,value})).sort((a,b)=>b.value-a.value);

function importAS(wb, source) {
  let best=null;
  for(const sh of sheets(wb)){
    let score=0;
    for(const r of sh.rows){if(norm(r[0])==='ASW')score++;if(norm(r[7])==='ASF')score++;}
    if(!best||score>best.score)best={...sh,score};
  }
  if(!best||best.score<10)throw new Error('AS: no se encontró la tabla ASW/ASF.');
  const matrix={NETCOL:{WG:{total:0,count:0,categories:new Map()},FC:{total:0,count:0,categories:new Map()}},INTEEGRA:{WG:{total:0,count:0,categories:new Map()},FC:{total:0,count:0,categories:new Map()}}};
  for(const r of best.rows){
    if(norm(r[0])==='ASW'){
      const company=norm(r[5]), value=n(r[3]);
      if(matrix[company]&&value!==0){matrix[company].WG.total+=value;matrix[company].WG.count++;catPush(matrix[company].WG.categories,r[1],value);}
    }
    if(norm(r[7])==='ASF'){
      const company=norm(r[12]), value=n(r[10]);
      if(matrix[company]&&value!==0){matrix[company].FC.total+=value;matrix[company].FC.count++;catPush(matrix[company].FC.categories,r[8],value);}
    }
  }
  for(const c of Object.keys(matrix))for(const p of ['WG','FC'])matrix[c][p].categories=mapCats(matrix[c][p].categories);
  const wg=matrix.NETCOL.WG.total+matrix.INTEEGRA.WG.total, fc=matrix.NETCOL.FC.total+matrix.INTEEGRA.FC.total;
  if(!(wg>0&&fc>0))throw new Error('AS: totales WG/FC inválidos.');
  return {summary:{wg_total:wg,fc_total:fc,difference:Math.abs(wg-fc),source,updated:fileDate(source),companies:['NETCOL','INTEEGRA'],partners:['WG','FC']},matrix};
}

function rowByLabel(rows,label){
  const target=norm(label);
  return rows.find(r=>r.some(v=>norm(v)===target));
}
function valuesAfterLabel(row,label,count=5){
  const idx=row.findIndex(v=>norm(v)===norm(label));
  if(idx<0)return [];
  return row.slice(idx+1).map(n).filter((v,i)=>v!==0||row.slice(idx+1)[i]===0).slice(0,count);
}
function importColegios(wb, source){
  let chosen=null;
  for(const sh of sheets(wb)){
    if(rowByLabel(sh.rows,'GIMNASIO FONTANA')&&rowByLabel(sh.rows,'COLEGIO VICTORIA SCHOOL')){chosen=sh;break;}
  }
  if(!chosen)throw new Error('Colegios: no se encontró la hoja resumen.');
  const g=valuesAfterLabel(rowByLabel(chosen.rows,'GIMNASIO FONTANA'),'GIMNASIO FONTANA');
  const v=valuesAfterLabel(rowByLabel(chosen.rows,'COLEGIO VICTORIA SCHOOL'),'COLEGIO VICTORIA SCHOOL');
  const krow=rowByLabel(chosen.rows,'KAJUYALI');
  const k=krow?valuesAfterLabel(krow,'KAJUYALI'):[0,0,0,0,0];
  if(g.length<4||v.length<4)throw new Error('Colegios: faltan valores anuales.');
  const years=['2023','2024','2025','2026'];
  const detail={WG:[],FC:[]};
  years.forEach((year,i)=>{
    if(g[i])detail.WG.push({year:Number(year),beneficiary:'GIMNASIO FONTANA',value:g[i]});
    if(v[i])detail.FC.push({year:Number(year),beneficiary:'COLEGIO VICTORIA SCHOOL',value:v[i]});
    if(k[i])detail.FC.push({year:Number(year),beneficiary:'KAJUYALI',value:k[i]});
  });
  const wg=detail.WG.reduce((a,x)=>a+x.value,0),fc=detail.FC.reduce((a,x)=>a+x.value,0);
  const yr=years.map((year,i)=>({year,WG:g[i]||0,FC:(v[i]||0)+(k[i]||0)}));
  return {summary:{wg_total:wg,fc_total:fc,difference:Math.abs(wg-fc),lower_partner:wg<fc?'WG':'FC',available:Math.abs(wg-fc),company:'INTEEGRA',source,updated:fileDate(source)},years:yr,detail};
}

function importLiquitty(wb, source){
  let best=null;
  for(const sh of sheets(wb)){
    const rows=sh.rows, q=rows.slice(24,130).filter(r=>n(r[16])!==0).length;
    if(!best||q>best.q)best={...sh,q};
  }
  if(!best||best.q<80)throw new Error('Liquitty: no se encontró la proyección de 96 meses en columna Q.');
  const rows=best.rows;
  const projection=rows.slice(24,120).map(r=>n(r[16]));
  if(projection.length!==96)throw new Error('Liquitty: la proyección no contiene 96 periodos.');
  const actual=[];
  for(let i=24;i<rows.length;i++){
    const gross=n(rows[i][17]),net=n(rows[i][18]),commission=n(rows[i][23]),vat=n(rows[i][24]);
    if(gross||net||commission||vat)actual.push({month:i-23,gross,net,commission,vat});
  }
  const investment=295898912, actualGross=actual.reduce((a,x)=>a+x.gross,0), actualNet=actual.reduce((a,x)=>a+x.net,0), commission=actual.reduce((a,x)=>a+x.commission,0), vat=actual.reduce((a,x)=>a+x.vat,0);
  const projected=projection.reduce((a,x)=>a+x,0), periods=actual.length, projectedToDate=projection.slice(0,periods).reduce((a,x)=>a+x,0);
  return {summary:{investment,investment_date:'2025-10-30',months:96,projected_net_total:projected,projected_profit:projected-investment,actual_gross:actualGross,commission_cost:commission,commission_vat:vat,actual_net:actualNet,remaining_capital:Math.max(0,investment-actualNet),recovery_pct:actualNet/investment,execution_vs_projection_to_date:projectedToDate?actualNet/projectedToDate:0,ahead_vs_projection_to_date:actualNet-projectedToDate,actual_periods:periods,owner:'INVERSIONES EON',source,updated:fileDate(source)},projection,actual};
}

const hasAlias=(row,aliases)=>{
  const cells=row.map(norm);
  return aliases.some(a=>cells.some(c=>c===a||c.includes(a)));
};
function importBoracay(wb, source){
  const wgAliases=['ASW','WRGB','WGB','WILLIAM GUTIERREZ'];
  const fcAliases=['ASF','FACR','FCR','FABIO CUELLAR'];
  let wg=0,fc=0,found=0;
  for(const sh of sheets(wb)){
    let valueIdx=10;
    for(let i=0;i<Math.min(sh.rows.length,30);i++){
      const idx=sh.rows[i].findIndex(v=>norm(v).includes('VALOR NETO A ABONAR A TARJETAS DAVIVIENDA'));
      if(idx>=0)valueIdx=idx;
    }
    for(const row of sh.rows){
      const value=n(row[valueIdx]); if(!value)continue;
      const isWG=hasAlias(row,wgAliases), isFC=hasAlias(row,fcAliases);
      if(isWG&&!isFC){wg+=value;found++;}
      if(isFC&&!isWG){fc+=value;found++;}
    }
  }
  if(!(wg>0&&fc>0&&found>1))throw new Error('Boracay: no se pudieron identificar movimientos WG/FC y valor neto.');
  return {summary:{wg_total:wg,fc_total:fc,difference:Math.abs(wg-fc),lower_partner:wg<fc?'WG':'FC',source,updated:fileDate(source)}};
}

const investorNorm=v=>{
  const s=norm(v);
  if(s.includes('INTEEGRA'))return 'INTEEGRA';
  if(s.includes('EON')||s.includes('EÓN'))return 'EON';
  if(s==='NS'||s.includes('NEWSTAR'))return 'NEWSTAR';
  return s||null;
};
const participationPct=v=>{const x=n(v);return x>0&&x<=1?x*100:x||null};
function importSentencias(wb, source){
  const name=wb.SheetNames.find(nm=>norm(nm).includes('VALORACION'));
  if(!name)throw new Error('Sentencias: no existe hoja VALORACIÓN.');
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:true,defval:null});
  const out=[];
  for(let i=1;i<rows.length;i++){
    const r=rows[i], stateRaw=norm(r[2]), caseName=String(r[3]||'').trim();
    if(!caseName||!stateRaw)continue;
    const state=stateRaw.includes('VIGENTE')?'Vigente':'Cancelado';
    const inv=state==='Cancelado'?(r[34]||r[9]):r[9];
    const investment=state==='Cancelado'?n(r[35]):n(r[19]);
    const current=state==='Cancelado'?n(r[36]):n(r[24]);
    if(!investment&&!current)continue;
    const exec=toISO(r[12]),entity=String(r[8]||'').trim(),turno=r[4]?String(r[4]).trim():null;
    const group_key=`${turno||caseName.toUpperCase()}|${entity.toUpperCase()}|${exec||''}`;
    const tir=state==='Cancelado'?n(r[39]):n(r[25]);
    out.push({
      row:i+1,state,case:caseName,turno,entity,investor:String(inv||''),participation:n(r[10]),regime:String(r[11]||'').trim(),
      fecha_ejecutoria:exec,fecha_cobro:toISO(r[13]),total_sentencia:n(r[16]),investment,current_or_return:current,
      invest_date:state==='Cancelado'?toISO(r[37]):toISO(r[20]),valuation_date:state==='Cancelado'?toISO(r[38]):toISO(r[23]),
      tir,valoracion_actual:n(r[24]),valor_pago:n(r[31]),fecha_pago:toISO(r[29]),valoracion_pago:n(r[27]),
      expectation_date:toISO(r[26]),investor_norm:investorNorm(inv),group_key,profit:current-investment,participation_pct:participationPct(r[10])
    });
  }
  if(!out.length)throw new Error('Sentencias: no se obtuvieron registros válidos.');
  return {summary:{rows:out.length,vigente_rows:out.filter(x=>x.state==='Vigente').length,cancelado_rows:out.filter(x=>x.state==='Cancelado').length,unique_cases:new Set(out.map(x=>x.group_key)).size,source,updated:fileDate(source)},rows:out};
}

export function importWorkbook(project, buffer, source) {
  const wb=XLSX.read(buffer,{type:'buffer',cellDates:true});
  if(project==='as')return importAS(wb,source);
  if(project==='colegios')return importColegios(wb,source);
  if(project==='liquitty')return importLiquitty(wb,source);
  if(project==='boracay')return importBoracay(wb,source);
  if(project==='sentencias')return importSentencias(wb,source);
  throw new Error(`Proyecto sin importador: ${project}`);
}
