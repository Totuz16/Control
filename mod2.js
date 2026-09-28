// ===== Módulo 2: Excel (leer/escribir), editar, todos los conceptos, filtros =====
const num=s=>{s=String(s).trim();if(s.includes(','))s=s.replace(/\./g,'').replace(',','.');else if(/^\d{1,3}(\.\d{3})+$/.test(s))s=s.replace(/\./g,'');return parseFloat(s)};
const dec=s=>s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
const colN=s=>{let n=0;for(const c of s)n=n*26+c.charCodeAt(0)-64;return n};
const colL=i=>{let s='';for(i++;i>0;i=Math.floor((i-1)/26))s=String.fromCharCode(65+(i-1)%26)+s;return s};
const inflar=async u=>new Uint8Array(await new Response(new Blob([u]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
async function unzip(buf){const u=new Uint8Array(buf),v=new DataView(buf),Z={};let e=u.length-22;while(v.getUint32(e,true)!==0x06054b50)e--;
  let n=v.getUint16(e+10,true),p=v.getUint32(e+16,true);
  for(let i=0;i<n;i++){const m=v.getUint16(p+10,true),cs=v.getUint32(p+20,true),nl=v.getUint16(p+28,true),xl=v.getUint16(p+30,true),cl=v.getUint16(p+32,true),o=v.getUint32(p+42,true),nm=new TextDecoder().decode(u.subarray(p+46,p+46+nl));p+=46+nl+xl+cl;
    const a=o+30+v.getUint16(o+26,true)+v.getUint16(o+28,true),d=u.subarray(a,a+cs);Z[nm]=()=>m===0?d:inflar(d)}
  return Z}
const txt=async(Z,n)=>Z[n]?new TextDecoder().decode(await Z[n]()):'';
function parseHoja(x,ss){const R=[];
  for(const r of x.matchAll(/<row\s[^>]*?\br="(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g)){const row=[];
    if(r[2])for(const c of r[2].matchAll(/<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
      const a=c[1],ref=(a.match(/\br="([A-Z]+)/)||[])[1],t=(a.match(/\bt="(\w+)"/)||[])[1],v=c[2]&&(c[2].match(/<v>([\s\S]*?)<\/v>/)||[])[1];let val=null;
      if(v!=null)val=t==='s'?ss[+v]:(t==='str'||t==='e')?dec(v):t==='b'?v==='1':+v;
      else if(t==='inlineStr')val=dec((c[2].match(/<t[^>]*>([\s\S]*?)<\/t>/)||[])[1]||'');
      if(ref)row[colN(ref)-1]=val}
    R[+r[1]-1]=row}
  return R}
async function leerXlsx(file){
  const Z=await unzip(await file.arrayBuffer());
  const ss=[...(await txt(Z,'xl/sharedStrings.xml')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>dec([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x=>x[1]).join('')));
  const wb=await txt(Z,'xl/workbook.xml'),rel=await txt(Z,'xl/_rels/workbook.xml.rels'),sh={},TB={},cache={};
  for(const m of wb.matchAll(/<sheet\s[^>]*?name="([^"]*)"[^>]*?r:id="([^"]*)"/g)){
    const el=(rel.match(new RegExp('<Relationship\\s[^>]*Id="'+m[2]+'"[^>]*>'))||[''])[0],t=(el.match(/Target="([^"]*)"/)||[])[1];
    if(t)sh[dec(m[1])]=t.startsWith('/')?t.slice(1):'xl/'+t}
  for(const [n,p] of Object.entries(sh))for(const m of (await txt(Z,p.replace(/([^/]+)$/,'_rels/$1.rels'))).matchAll(/Target="[^"]*?tables\/([^"]+)"/g)){
    const tx=await txt(Z,'xl/tables/'+m[1]),nm=(tx.match(/<table\s[^>]*?\bname="([^"]*)"/)||[])[1],rf=(tx.match(/<table\s[^>]*?\bref="([^"]*)"/)||[])[1];
    if(nm&&rf)TB[dec(nm)]={hoja:n,ref:rf}}
  const tabla=async nm=>{const t=TB[nm];if(!t)return null;const R=cache[t.hoja]||(cache[t.hoja]=parseHoja(await txt(Z,sh[t.hoja]),ss));
    const [a,b]=t.ref.split(':'),c1=colN(a.match(/[A-Z]+/)[0])-1,r1=+a.match(/\d+/)[0]-1,c2=colN(b.match(/[A-Z]+/)[0])-1,r2=+b.match(/\d+/)[0]-1,rows=[];
    for(let r=r1;r<=r2;r++){const w=[];for(let c=c1;c<=c2;c++)w.push((R[r]||[])[c]??null);rows.push(w)}
    return {h:rows[0].map(x=>String(x??'').trim().toUpperCase()),r:rows.slice(1)}};
  return {tabla,TB}}
async function excelADatos(file){
  const X_=await leerXlsx(file),reg=await X_.tabla('REGISTROS'),cl=await X_.tabla('CLASE'),ho=await X_.tabla('HOMOLOGACION');
  if(!reg||!cl||!ho)throw new Error('No encontré las tablas REGISTROS, CLASE u HOMOLOGACION en el archivo');
  const cuentas=cl.r.filter(r=>r[0]).map(r=>({medio:r[0],clase:r[1],grupo:r[2],moneda:r[3]})),signos={},cats={},subs={},movimientos=[];
  ho.r.filter(r=>r[0]&&r[1]).forEach(r=>signos[r[0]+'|'+r[1]]=r[4]);
  for(const [n,t] of Object.entries(X_.TB)){const h=t.hoja.normalize('NFC');
    if(h==='Listas Categorías'||h==='Listas Subcategorías')(h==='Listas Categorías'?cats:subs)[n]=(await X_.tabla(n)).r.flat().filter(x=>x!=null&&x!=='').map(String)}
  const I={};['FECHA','MEDIO','CONCEPTO','MONTO','CATEGORÍA','SUBCATEGORÍA','ESTABLECIMIENTO','OBSERVACIÓN','MONEDA'].forEach(k=>I[k]=reg.h.indexOf(k));
  const s=v=>v==null||v===''?null:String(v);
  for(const r of reg.r){const conc=r[I['CONCEPTO']];if(!conc)continue;const medio=r[I['MEDIO']]==='CXC'?'CxC':r[I['MEDIO']],f=r[I['FECHA']];
    movimientos.push({fecha:typeof f==='number'&&f>20000?new Date(Math.round((f-25569)*864e5)).toISOString().slice(0,10):(typeof f==='string'&&/^\d{4}-\d\d-\d\d/.test(f)?f.slice(0,10):null),
      medio,concepto:conc,monto:typeof r[I['MONTO']]==='number'?r[I['MONTO']]:0,moneda:r[I['MONEDA']]||(cuentas.find(c=>c.medio===medio)||{}).moneda||'COP',
      categoria:s(r[I['CATEGORÍA']]),subcategoria:s(r[I['SUBCATEGORÍA']]),establecimiento:s(r[I['ESTABLECIMIENTO']]),observacion:s(r[I['OBSERVACIÓN']])})}
  return {app:'finanzas-pwa',version:1,cuentas,signos,categorias:cats,subcategorias:subs,movimientos}}

// ---- Escritor .xlsx (zip sin compresión) ----
const crcT=(()=>{const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
const crc=u=>{let c=-1;for(const b of u)c=crcT[(c^b)&255]^(c>>>8);return (~c)>>>0};
function zipar(files){const te=new TextEncoder(),P=[],D=[],ks=Object.keys(files);let off=0;
  for(const n of ks){const nb=te.encode(n),d=te.encode(files[n]),c=crc(d),h=new DataView(new ArrayBuffer(30)),e=new DataView(new ArrayBuffer(46));
    h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint32(14,c,true);h.setUint32(18,d.length,true);h.setUint32(22,d.length,true);h.setUint16(26,nb.length,true);P.push(h,nb,d);
    e.setUint32(0,0x02014b50,true);e.setUint16(4,20,true);e.setUint16(6,20,true);e.setUint32(16,c,true);e.setUint32(20,d.length,true);e.setUint32(24,d.length,true);e.setUint16(28,nb.length,true);e.setUint32(42,off,true);D.push(e,nb);off+=30+nb.length+d.length}
  const cl=D.reduce((a,x)=>a+x.byteLength,0),z=new DataView(new ArrayBuffer(22));z.setUint32(0,0x06054b50,true);z.setUint16(8,ks.length,true);z.setUint16(10,ks.length,true);z.setUint32(12,cl,true);z.setUint32(16,off,true);
  return new Blob([...P,...D,z],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
const xe=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const serial=f=>Math.round(Date.parse(f+'T00:00:00Z')/864e5)+25569;
function hojaXml(rows){return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+rows.map((r,i)=>`<row r="${i+1}">`+r.map((v,j)=>{const a=colL(j)+(i+1);
  return v==null||v===''?'':typeof v==='number'?`<c r="${a}"><v>${v}</v></c>`:v.d?`<c r="${a}" s="1"><v>${v.d}</v></c>`:`<c r="${a}" t="inlineStr"><is><t>${xe(v)}</t></is></c>`}).join('')+'</row>').join('')+'</sheetData></worksheet>'}
function construirXlsx(){
  const S={};M.forEach(m=>S[m.medio]=(S[m.medio]||0)+val(m));
  const H=[['Registro Diario',[['FECHA','MEDIO','CONCEPTO','MONTO','CATEGORÍA','SUBCATEGORÍA','ESTABLECIMIENTO','OBSERVACIÓN','MONEDA','CLASE','MEDIO AGRUP.','AÑO-MES','VALOR']].concat(M.slice().sort((a,b)=>(a.fecha||'').localeCompare(b.fecha||'')||a.id-b.id).map(m=>{const c=CM[m.medio]||{};return [m.fecha?{d:serial(m.fecha)}:null,m.medio,m.concepto,m.monto,m.categoria,m.subcategoria,m.establecimiento,m.observacion,m.moneda,c.clase,c.grupo,m.fecha?m.fecha.slice(0,7):null,val(m)]}))],
    ['Cuentas (tabla CLASE)',[['MEDIO','CLASE','MEDIO AGRUP.','MONEDA']].concat(C.cuentas.map(c=>[c.medio,c.clase,c.grupo,c.moneda]))],
    ['Homologaciones',[['CLASE','CONCEPTO','SIGNO']].concat(Object.entries(C.signos).map(([k,v])=>[...k.split('|'),v]))],
    ['Categorías',[['LISTA','CATEGORÍA']].concat(Object.entries(C.categorias).flatMap(([l,a])=>a.map(x=>[l,x])))],
    ['Subcategorías',[['CATEGORÍA','SUBCATEGORÍA']].concat(Object.entries(C.subcategorias).flatMap(([l,a])=>a.map(x=>[l,x])))],
    ['Saldos',[['MEDIO','MEDIO AGRUP.','CLASE','MONEDA','SALDO']].concat(C.cuentas.map(c=>[c.medio,c.grupo,c.clase,c.moneda,Math.round((S[c.medio]||0)*100)/100]))]];
  const F={'[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+H.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')+'</Types>',
    '_rels/.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml':'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+H.map((h,i)=>`<sheet name="${xe(h[0])}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')+'</sheets></workbook>',
    'xl/_rels/workbook.xml.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+H.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')+`<Relationship Id="rId${H.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml':'<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'};
  H.forEach((h,i)=>F[`xl/worksheets/sheet${i+1}.xml`]=hojaXml(h[1]));
  return zipar(F)}
async function entregar(f){try{if(navigator.canShare&&navigator.canShare({files:[f]}))await navigator.share({files:[f],title:f.name});else{const a=document.createElement('a');a.href=URL.createObjectURL(f);a.download=f.name;a.click()}return true}catch(e){if(e.name!=='AbortError')alert('No se pudo exportar: '+e.message);return false}}

// ---- Importar / restaurar (Excel o copia JSON) ----
restaurar=async function(file){let d;
  try{d=/\.xlsx$/i.test(file.name)?await excelADatos(file):JSON.parse(await file.text())}catch(e){alert('No se pudo leer el archivo: '+e.message);return}
  if(d.app!=='finanzas-pwa'||!Array.isArray(d.movimientos)||!d.cuentas||!d.signos){alert('Archivo no reconocido.');return}
  const fs=d.movimientos.map(m=>m.fecha).filter(Boolean).sort();
  if(!confirm(`Leí ${d.movimientos.length} movimientos (${fs[0]} a ${fs[fs.length-1]}) y ${d.cuentas.length} cuentas.${M.length?` Esto reemplaza los ${M.length} actuales.`:''} ¿Continuar?`))return;
  const t=db.transaction(['mov','meta'],'readwrite');t.objectStore('mov').clear();d.movimientos.forEach(m=>t.objectStore('mov').add(m));
  t.objectStore('meta').put({k:'cfg',v:{cuentas:d.cuentas,signos:d.signos,categorias:d.categorias,subcategorias:d.subcategorias}});
  await done(t);await load();render();alert(`Listo: ${M.length} movimientos cargados.`)};

// ---- Crear / editar cualquier movimiento ----
const PAIR={'TRANS.SALIDA':'TRANS.ENTRADA','RETIRO':'RETIRO.INGRESO','DESEMBOLSO':'DESEMBOLSO.INGRESO','AHORRO':'AHORRO.FUENTE','DEPÓSITO':'DEPÓSITO.FUENTE','AVANCE-TRANSFERENCIA':'AVANCE-TRANSFERENCIA.INGRESO','PAGO.CxC':'PAGO.CxC.INGRESO','ADM.SALIDA':'ADM.ENTRADA','AMORTIZACIÓN.PAGO':'PAGO.PASIVO','AHORRO.RETIRO':'AHORRO.RETIRO.INGRESO','AVANCE':'AVANCE.INGRESO'};
let E={};
const concs=c=>Object.keys(C.signos).filter(k=>k.startsWith((CM[c]||{}).clase+'|')).map(k=>k.slice(k.indexOf('|')+1));
const CATS=()=>[...new Set(Object.values(C.categorias).flat())].sort(),SUBS=c=>C.subcategorias[c]||[...new Set(Object.values(C.subcategorias).flat())].sort();
const leer=()=>['fecha','cta','conc','monto','cat','sub','est','obs','cta2'].forEach(k=>{const e=$('#e-'+k);if(e)E[k]=e.value});
function formMov(m){E=m?{id:m.id,raw:m,fecha:m.fecha||'',cta:m.medio,conc:m.concepto,monto:String(m.monto),cat:m.categoria||'',sub:m.subcategoria||'',est:m.establecimiento||'',obs:m.observacion||''}:{fecha:today(),cta:(X.ult||{}).gasto||C.cuentas[0].medio,conc:'',monto:'',cat:'',sub:'',est:'',obs:''};pintar()}
function pintar(){
  const cs=concs(E.cta);if(!cs.includes(E.conc))E.conc=cs[0];
  const par=!E.id&&PAIR[E.conc],c2=par?C.cuentas.filter(c=>C.signos[c.clase+'|'+par]!=null).map(c=>c.medio):[];
  if(par&&!c2.includes(E.cta2))E.cta2=c2[0];
  sheet(`<div class="top"><h1 style="margin:0">${E.id?'Editar':'Nuevo'} movimiento</h1><button onclick="close()">Cancelar</button></div><div style="height:12px"></div>
  <input id="e-fecha" type="date" value="${E.fecha}"><select id="e-cta">${opts(C.cuentas.map(c=>c.medio),E.cta)}</select><select id="e-conc">${opts(cs,E.conc)}</select>
  <input id="e-monto" inputmode="decimal" placeholder="Valor" value="${esc(E.monto)}">
  ${par?`<div class="msg">${esc(E.conc)} en ${esc(E.cta)} + ${esc(par)} en la cuenta contraparte:</div><select id="e-cta2">${opts(c2,E.cta2)}</select>`:''}
  <input id="e-cat" list="dl-c" placeholder="Categoría" value="${esc(E.cat)}"><datalist id="dl-c">${CATS().map(x=>`<option value="${esc(x)}">`).join('')}</datalist>
  <input id="e-sub" list="dl-s" placeholder="Subcategoría" value="${esc(E.sub)}"><datalist id="dl-s">${SUBS(E.cat).map(x=>`<option value="${esc(x)}">`).join('')}</datalist>
  <input id="e-est" placeholder="Establecimiento" value="${esc(E.est)}"><input id="e-obs" placeholder="Observación" value="${esc(E.obs)}"><button class="btn" id="e-ok">Guardar</button>`);
  $('#e-cta').onchange=$('#e-conc').onchange=()=>{leer();pintar()};$('#e-cat').onchange=()=>{leer();pintar()};$('#e-ok').onclick=guardarMov}
async function guardarMov(){leer();const monto=num(E.monto);if(!(monto>=0)){alert('Escribe un valor válido.');return}
  const b={fecha:E.fecha||null,medio:E.cta,concepto:E.conc,monto,moneda:CM[E.cta].moneda,categoria:E.cat||null,subcategoria:E.sub||null,establecimiento:E.est||null,observacion:E.obs||null};
  if(E.id)await put('mov',[{...E.raw,...b,id:E.id}]);
  else{const rows=[b];if(PAIR[E.conc]&&E.cta2)rows.push({...b,medio:E.cta2,concepto:PAIR[E.conc],moneda:CM[E.cta2].moneda});await put('mov',rows)}
  await load();close();render()}
detalle=function(id){const m=M.find(x=>x.id===id);if(!m)return;
  sheet(`<h1>${money(m.monto,m.moneda)}</h1><div class="msg">${esc(m.concepto)} · ${esc(m.medio)}<br>${m.fecha||'Saldo inicial'}<br>${esc([m.categoria,m.subcategoria,m.establecimiento,m.observacion].filter(Boolean).join(' · '))}</div>
  <button class="btn" id="ed">Editar</button><button class="btn del" id="del">Eliminar</button><button class="btn sec" onclick="close()">Cerrar</button>`);
  $('#ed').onclick=()=>formMov(m);
  $('#del').onclick=async()=>{if(!confirm('¿Eliminar este movimiento? Si es parte de una pareja (transferencia, pago de deuda…), elimina también la otra fila.'))return;const t=db.transaction('mov','readwrite');t.objectStore('mov').delete(id);await done(t);await load();close();render()}};

// ---- Filtros ----
let FL={campo:''};
const CAMPOS=[['','Todos los campos'],['medio','Cuenta'],['concepto','Concepto'],['categoria','Categoría'],['subcategoria','Subcategoría'],['establecimiento','Establecimiento'],['observacion','Observación']];
const hoyF=m=>{const t=q.toLowerCase();if(!t)return true;return (FL.campo?[m[FL.campo]]:[m.medio,m.concepto,m.categoria,m.subcategoria,m.establecimiento,m.observacion]).some(x=>x&&String(x).toLowerCase().includes(t))};
movs=function(){const n=['desde','hasta','medio','concepto','categoria'].filter(k=>FL[k]).length;
  const L=M.filter(m=>(!FL.desde||(m.fecha||'')>=FL.desde)&&(!FL.hasta||(m.fecha||'')<=FL.hasta)&&(!FL.medio||m.medio===FL.medio)&&(!FL.concepto||m.concepto===FL.concepto)&&(!FL.categoria||m.categoria===FL.categoria)&&hoyF(m)).sort((a,b)=>(b.fecha||'0').localeCompare(a.fecha||'0')||b.id-a.id);
  const tot=L.filter(m=>m.moneda==='COP').reduce((a,m)=>a+m.monto,0);
  return `<h1>Movimientos</h1><input id="q" type="search" placeholder="Buscar…" value="${esc(q)}"><button class="btn sec" id="flt">Filtros${n||FL.campo?` (${n+(FL.campo?1:0)} activos)`:''}</button>
  <div class="mute" style="margin:0 2px 8px;font-size:14px">${L.length.toLocaleString('es-CO')} resultados · suma de montos ${money(tot)}${L.length>200?' · mostrando 200':''}</div>
  <div class="card">${L.slice(0,200).map(m=>{const v=val(m);return `<div class="row" data-id="${m.id}"><span>${esc(m.categoria||m.concepto)}${m.subcategoria?' · '+esc(m.subcategoria):''}<small>${m.fecha||'Saldo inicial'} · ${esc(m.concepto)} · ${esc(m.medio)}${m.establecimiento?' · '+esc(m.establecimiento):''}${m.observacion?' · '+esc(m.observacion):''}</small></span><b class="${v<0?'neg':'pos'}">${money(m.monto,m.moneda)}</b></div>`}).join('')||'<div class="row mute">Nada encontrado</div>'}</div>`};
function filtros(){const cs=[...new Set(M.map(m=>m.concepto))].sort(),ca=CATS(),o=(a,s)=>`<option value="">Todos</option>`+a.map(x=>`<option ${x===s?'selected':''}>${esc(x)}</option>`).join('');
  sheet(`<div class="top"><h1 style="margin:0">Filtros</h1><button onclick="close()">Cerrar</button></div><div style="height:12px"></div>
  <div class="mute">Desde</div><input id="g-d" type="date" value="${FL.desde||''}"><div class="mute">Hasta</div><input id="g-h" type="date" value="${FL.hasta||''}">
  <select id="g-m">${o(C.cuentas.map(c=>c.medio),FL.medio)}</select><select id="g-c">${o(cs,FL.concepto)}</select><select id="g-k">${o(ca,FL.categoria)}</select>
  <div class="mute">El texto de búsqueda se aplica en:</div><select id="g-f">${CAMPOS.map(([v,t])=>`<option value="${v}" ${v===FL.campo?'selected':''}>${t}</option>`).join('')}</select>
  <button class="btn" id="g-ok">Aplicar</button><button class="btn sec" id="g-x">Limpiar filtros</button>`);
  $('#g-ok').onclick=()=>{FL={desde:$('#g-d').value,hasta:$('#g-h').value,medio:$('#g-m').value,concepto:$('#g-c').value,categoria:$('#g-k').value,campo:$('#g-f').value};close();render()};
  $('#g-x').onclick=()=>{FL={campo:''};q='';close();render()}}
if(typeof document!=='undefined'){
  $('#file').accept='.xlsx,.json,application/json';
  document.addEventListener('click',async e=>{const i=e.target.id;
    if(i==='flt')filtros();
    else if(i==='xlsx'){if(await entregar(new File([construirXlsx()],`finanzas-${today()}.xlsx`,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}))){await setX('copia',Date.now());render()}}});
  $('#add').onclick=()=>{if(!C.cuentas.length){alert('Primero importa tu Excel en Más → Importar.');return}R={tipo:'gasto'};form()};
}
