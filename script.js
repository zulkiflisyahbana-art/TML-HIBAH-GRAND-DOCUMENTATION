/* TML HIBAH – Grand Documentation · logika aplikasi (statis, berjalan di browser) */
'use strict';
/* ============ Mesin Excel: mengisi template resmi TML tanpa mengubah format ============ */
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const num=v=>{const n=parseFloat(String(v??'').replace(/[^0-9.\-]/g,''));return isNaN(n)?0:n};
const fmt=n=>Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
function setCell(xml,ref,val,formula){
  const m=xml.match(new RegExp('<c r="'+ref+'"([^>]*?)(?:/>|>[\\s\\S]*?</c>)'));
  if(!m) throw new Error('Sel '+ref+' tidak ada di template');
  const s=(m[1].match(/\ss="(\d+)"/)||[])[1], a=s?' s="'+s+'"':'';
  let n;
  if(formula!==undefined) n='<c r="'+ref+'"'+a+'><f>'+formula+'</f><v>'+val+'</v></c>';
  else if(typeof val==='number') n='<c r="'+ref+'"'+a+'><v>'+val+'</v></c>';
  else if(val===''||val==null) n='<c r="'+ref+'"'+a+'/>';
  else n='<c r="'+ref+'"'+a+' t="inlineStr"><is><t xml:space="preserve">'+esc(val)+'</t></is></c>';
  return xml.replace(m[0],()=>n);
}
function cellsFor(kind,s){
  const ci=kind==='CI', c=ci?'C':'E', f=ci?'L':'N', out=[], P=(r,v,fx)=>out.push([r,v,fx]);
  P(ci?'S5':'T5', s.invNo+' '.repeat(28)+s.invDate);
  P(c+'20', s.depart); P(c+'22', s.vessel);
  P(c+'21','④ VESSEL/FLIGHT'+' '.repeat(21)+'⑤ FROM: '+s.from);
  P(c+'23','⑥ TO: '+s.to);
  [5,6,7,8,9].forEach((r,i)=>P(c+r,s.shipper[i]||''));
  [11,12,13,14,15,16,17,18].forEach((r,i)=>P(c+r,s.consignee[i]||''));
  let tot=0,tp=0; const amt=[];
  s.units.forEach((u,i)=>{
    const b=27+i*7, k=num(u.kg), p=num(u.price); tot+=k; tp+=p; amt.push('X'+b);
    [u.brand,u.model,u.vin,/^\d+$/.test(String(u.year).trim())?Number(u.year):u.year,u.cyl,u.eng,u.color].forEach((v,j)=>P(f+(b+j),v));
    if(ci){P('V'+b,p);P('X'+b,p,'V'+b);} else {P('W'+b,fmt(k)+' Kgs');P('Y'+b,fmt(k)+' Kgs');}
  });
  const T=27+7*s.units.length;
  if(ci) P('X'+T,tp,amt.join('+')); else {P('W'+T,fmt(tot)+' Kgs');P('Y'+T,fmt(tot)+' Kgs');}
  return out;
}
const MR=/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/;
function restructure(xml,N){
  const delta=7*(N-2);
  const sh=(str,o)=>str.replace(/<row r="(\d+)"/,(m,r)=>'<row r="'+(+r+o)+'"').replace(/<c r="([A-Z]+)(\d+)"/g,(m,c,r)=>'<c r="'+c+(+r+o)+'"');
  const sd=xml.match(/<sheetData>([\s\S]*)<\/sheetData>/)[1];
  const rows=[...sd.matchAll(/<row r="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)].map(m=>[+m[1],m[0]]);
  const A=rows.filter(r=>r[0]<=33), B=rows.filter(r=>r[0]>=34&&r[0]<=40), C=rows.filter(r=>r[0]>=41);
  let out=A.map(r=>r[1]).join('');
  for(let k=0;k<N-1;k++) out+=B.map(r=>sh(r[1],7*k)).join('');
  out+=C.map(r=>sh(r[1],delta)).join('');
  xml=xml.replace(sd,()=>out);
  const mm=[...xml.matchAll(/<mergeCell ref="([^"]+)"\/>/g)].map(m=>m[1]);
  const mv=(ref,o)=>{const m=ref.match(MR);return m[1]+(+m[2]+o)+(m[3]?':'+m[3]+(+m[4]+o):'');};
  const rr=ref=>{const m=ref.match(MR);return [+m[2],+(m[4]||m[2])];};
  const nm=[];
  mm.forEach(ref=>{const [a,b]=rr(ref); if(b<=33)nm.push(ref); else if(a>=41)nm.push(mv(ref,delta));});
  for(let k=0;k<N-1;k++) mm.forEach(ref=>{const [a,b]=rr(ref); if(a>=34&&b<=40)nm.push(mv(ref,7*k));});
  xml=xml.replace(/<mergeCells[\s\S]*?<\/mergeCells>/,'<mergeCells count="'+nm.length+'">'+nm.map(r=>'<mergeCell ref="'+r+'"/>').join('')+'</mergeCells>');
  xml=xml.replace(/(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)"/,(m,a,r)=>a+(+r+delta)+'"');
  if(N>=3){
    if(!/<pageSetUpPr/.test(xml)) xml=/<sheetPr[^>]*\/>/.test(xml)?xml.replace(/<sheetPr([^>]*)\/>/,'<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>'):/<sheetPr/.test(xml)?xml.replace(/<\/sheetPr>/,'<pageSetUpPr fitToPage="1"/></sheetPr>'):xml.replace(/(<worksheet[^>]*>)/,'$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    else xml=xml.replace(/<pageSetUpPr[^>]*\/>/,'<pageSetUpPr fitToPage="1"/>');
    xml=xml.replace(/<pageSetup([^>]*?)\/>/,(m,a)=>'<pageSetup'+a.replace(/\s(fitToWidth|fitToHeight)="\d+"/g,'')+' fitToWidth="1" fitToHeight="'+(N<=5?1:0)+'"/>');
  }
  return xml;
}
async function buildXlsx(b64,kind,s,assets){
  const N=s.units.length; if(N<1) throw new Error('Minimal 1 unit');
  const z=await JSZip.loadAsync(b64,{base64:true});
  if(assets){const mp=kind==='CI'?{'xl/media/image1.png':'sig','xl/media/image2.png':'stamp'}:{'xl/media/image2.png':'sig','xl/media/image3.png':'stamp'};for(const [p,r] of Object.entries(mp))if(assets[r]&&z.file(p))z.file(p,assets[r]);}
  let x=await z.file('xl/worksheets/sheet1.xml').async('string');
  if(N!==2) x=restructure(x,N);
  for(const [r,v,fx] of cellsFor(kind,s)) x=setCell(x,r,v,fx);
  z.file('xl/worksheets/sheet1.xml',x);
  let w=await z.file('xl/workbook.xml').async('string');
  w=w.replace(/<calcPr([^>]*)\/>/,(m,a)=>a.includes('fullCalcOnLoad')?m:'<calcPr'+a+' fullCalcOnLoad="1"/>');
  if(N!==2) w=w.replace(/(\$[A-Z]+\$\d+:\$[A-Z]+\$)(\d+)(<\/definedName>)/,(m,a,r,e)=>a+(+r+7*(N-2))+e);
  z.file('xl/workbook.xml',w);
  if(N!==2){const dn='xl/drawings/drawing1.xml'; if(z.file(dn)){let d=await z.file(dn).async('string'); d=d.replace(/<xdr:row>(\d+)<\/xdr:row>/g,(m,r)=>'<xdr:row>'+(+r+7*(N-2))+'</xdr:row>'); z.file(dn,d);}}
  if(z.file('xl/calcChain.xml')){
    z.remove('xl/calcChain.xml');
    for(const [p,re] of [['[Content_Types].xml',/<Override[^>]*calcChain[^>]*\/>/],['xl/_rels/workbook.xml.rels',/<Relationship [^>]*calcChain[^>]*\/>/]]){const t=await z.file(p).async('string'); z.file(p,t.replace(re,''));}
  }
  return z.generateAsync({type:'arraybuffer',compression:'DEFLATE'});
}
/* ---- Impor dari Excel CI/PL buatan aplikasi ini ---- */
async function readXlsx(buf){
  const z=await JSZip.loadAsync(buf);
  const wb=await z.file('xl/workbook.xml').async('string'), sh=await z.file('xl/worksheets/sheet1.xml').async('string');
  const sst=z.file('xl/sharedStrings.xml')?[...(await z.file('xl/sharedStrings.xml').async('string')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>[...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t=>t[1]).join('')):[];
  const un=t=>t.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
  const cells={};
  for(const m of sh.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
    const body=m[3]||''; let v=null;
    if(/t="s"/.test(m[2])){const x=body.match(/<v>(\d+)<\/v>/); if(x)v=sst[+x[1]];}
    else if(/t="inlineStr"/.test(m[2])){const x=body.match(/<t[^>]*>([\s\S]*?)<\/t>/); if(x)v=x[1];}
    else {const x=body.match(/<v>([^<]*)<\/v>/); if(x)v=x[1];}
    if(v!=null&&v!=='') cells[m[1]]=un(String(v));
  }
  return {kind:/Invoice/i.test(wb.match(/<sheet [^>]*name="([^"]*)"/)[1])?'CI':'PL',cells};
}
function parseSheet({kind,cells}){
  const ci=kind==='CI', c=ci?'C':'E', f=ci?'L':'N', g=r=>(cells[r]||''), out={units:[]};
  const inv=g(ci?'S5':'T5').split(/\s{2,}/); out.invNo=(inv[0]||'').trim(); out.invDate=(inv[1]||'').trim();
  out.depart=g(c+'20'); out.vessel=g(c+'22');
  const fr=g(c+'21').match(/FROM:\s*(.*)$/); if(fr) out.from=fr[1].trim();
  const to=g(c+'23').match(/TO:\s*(.*)$/); if(to) out.to=to[1].trim();
  const lines=rows=>{const a=rows.map(r=>g(c+r)); while(a.length&&!a[a.length-1])a.pop(); return a;};
  out.shipper=lines([5,6,7,8,9]); out.consignee=lines([11,12,13,14,15,16,17,18]);
  for(let i=0;i<60;i++){ const b=27+7*i; if(!g(f+(b+2)))break;
    const u={brand:g(f+b),model:g(f+(b+1)),vin:g(f+(b+2)),year:g(f+(b+3)),cyl:g(f+(b+4)),eng:g(f+(b+5)),color:g(f+(b+6))};
    if(ci) u.price=g('V'+b); else u.kg=String(num(g('W'+b)));
    out.units.push(u); }
  return out;
}
/* ============ Aplikasi ============ */
let D=null; // data contoh dimuat dari vault terenkripsi setelah kode akses benar

const $=id=>document.getElementById(id);
const HF=[['name','Nama proyek (untuk nama file)'],['invNo','No. invoice'],['invDate','Tanggal invoice'],['depart','Departure date (boleh kosong)'],['vessel','Vessel (boleh kosong)'],['from','From'],['to','To'],['blTotal','Berat total di B/L (kg)'],['blQty','Jumlah unit di B/L']];
const UF=[['brand','Brand'],['model','Model'],['vin','VIN / Chassis'],['year','Year'],['cyl','Cylinder'],['eng','Engine (jenis)'],['color','Color'],['kg','Berat (kg), net = gross'],['price','Harga satuan (₩)'],['engNo','Nomor mesin (opsional, hanya laporan)'],['dim','Dimensi P×L×T (opsional, hanya laporan)']];
const CP='<svg width="14" height="14" viewBox="0 0 14 14"><path d="M3 5l4 4 4-4" fill="none" stroke="currentColor" stroke-width="2"/></svg>';
const LS={get(k,d){try{const v=localStorage.getItem(k);return v==null?d:JSON.parse(v);}catch(e){return d;}},set(k,v){try{localStorage.setItem(k,JSON.stringify(v));}catch(e){}}};
let S, OPT={}, SAVED=[], UID=0, PK='CI', FT=new Set();
const rid=()=>Math.random().toString(36).slice(2,10);
const cfg=()=>LS.get('tml_cfg',{url:''});
const note=t=>document.querySelectorAll('.note').forEach(e=>e.textContent=t);
const blank=(p={})=>({id:++UID,brand:'',model:'',vin:'',year:'',cyl:'',eng:'',color:'',kg:'',price:'',engNo:'',dim:'',edit:true,...p});
const addOpt=(k,v)=>{v=String(v??'').trim();if(v)(OPT[k]=OPT[k]||new Set()).add(v);};
function learn(s){HF.forEach(([k])=>addOpt(k,s[k]));(s.units||[]).forEach(u=>UF.forEach(([k])=>addOpt(k,u[k])));addOpt('shipper',(s.shipper||[]).join('\n'));addOpt('consignee',(s.consignee||[]).join('\n'));}
function persistOpts(){const o={};for(const k in OPT)o[k]=[...OPT[k]];LS.set('tml_opts',o);}
const fresh=()=>({key:rid(),confirmed:null,confirmedBy:'',log:[],rev:0,savedRev:null,savedAt:null,done:{},ext:null,src:{},marks:{},readAt:null,name:'',invNo:'',invDate:'',depart:'',vessel:'',from:D.from,to:D.to,blTotal:'',blQty:'',blChassis:'',mou:'',shipper:[...D.shipper],consignee:[...D.consignee],units:[blank()]});
const example=()=>({...D,key:(S&&S.key)||rid(),confirmed:null,confirmedBy:'',log:[],rev:0,savedRev:null,savedAt:null,done:{},ext:null,src:{},marks:{},readAt:null,shipper:[...D.shipper],consignee:[...D.consignee],units:D.units.map(u=>blank({...u,edit:false}))});
const clean=s=>({...s,units:s.units.map(({id,edit,sure,ai,tp,...u})=>u)});
const hydrate=d=>{const u=(d.units||[]).map(x=>blank({...x,edit:false}));return {...fresh(),...d,key:d.key||rid(),units:u.length?u:[blank()]};};
let dT;function draft(){clearTimeout(dT);dT=setTimeout(()=>LS.set('tml_draft',clean(read())),250);}
function touch(){ if(S.confirmed){S.confirmed=null;logA('Data diubah setelah konfirmasi — konfirmasi dibatalkan');} S.rev=(S.rev||0)+1; S.done={}; draft();updConf();}
const cb=(id,k,attrs,val)=>'<div class="cb"><input id="'+id+'" '+attrs+' value="'+String(val??'').replace(/&/g,'&amp;').replace(/"/g,'&quot;')+'" autocomplete="off"><button type="button" class="dd" data-k="'+k+'" data-t="'+id+'" aria-label="Pilih">'+CP+'</button></div>';
function hdr(){ $('hdr').innerHTML=HF.map(([k,l])=>'<div class="w"><label>'+l+'</label>'+cb('h_'+k,k,'data-h="'+k+'"','')+'</div>').join(''); }
function render(){ HF.forEach(([k])=>$('h_'+k).value=S[k]??''); $('shipper').value=S.shipper.join('\n'); $('consignee').value=S.consignee.join('\n'); $('blChassis').value=S.blChassis||''; $('mou').value=S.mou||''; document.querySelectorAll('[data-h],#shipper,#consignee,#blChassis,#mou').forEach(e=>e.classList.remove('ai','tpl')); Object.entries(S.marks||{}).forEach(([id,k])=>{const e=$(id); if(e)e.classList.add(k);});
  renderUnits(); renderFiles(); renderLog(); updConf(); }
const sum=u=>[u.brand,u.model].filter(Boolean).join(' ')+(u.vin?' · '+u.vin:'')+(u.kg?' · '+fmt(num(u.kg))+' kg':'')+(num(u.price)?' · ₩'+fmt(num(u.price)).replace('.00',''):'');
function renderUnits(){
  $('units').innerHTML=S.units.map((u,i)=>{
    const okU=u.brand&&u.model&&VINRE.test(u.vin.trim())&&u.year&&num(u.kg)>0&&num(u.price)>0, ub=u.edit?'<span class="ub ed">✎ Sedang diedit</span>':(okU?'<span class="ub ok">✓ Lengkap</span>':'<span class="ub wn">⚠ Belum lengkap</span>');
    const head='<div class="uh"><b>Unit '+(i+1)+'</b>'+ub+(u.edit?'<span class="s"></span>':'<span class="s">'+(sum(u)||'(belum diisi)')+'</span>')+'<span class="row">'
      +(u.edit?'<button class="sm p" data-a="done" data-id="'+u.id+'">Selesai</button>':'<button class="sm" data-a="edit" data-id="'+u.id+'">Edit</button>')
      +'<button class="sm d" data-a="del" data-id="'+u.id+'">'+(u.sure?'Yakin hapus?':'Hapus')+'</button></span></div>';
    const form=u.edit?'<div class="g" style="margin-top:12px">'+UF.map(([k,l])=>'<div class="w"><label>'+l+'</label>'+cb('u'+u.id+k,k,'class="'+((u.ai||{})[k]?'ai':((u.tp||{})[k]?'tpl':''))+'" data-u="'+u.id+'" data-k="'+k+'"',u[k])+'</div>').join('')+'</div>':'';
    return '<div class="card">'+head+form+'</div>'; }).join('');
  $('uCount').textContent=S.units.length+' unit'; check(); }
function read(){ const s=clean(S); s.shipper=s.shipper.slice(0,5); s.consignee=s.consignee.slice(0,8); return s; }
/* ---- Pemeriksaan dan laporan verifikasi ---- */
const VINRE=/^[A-HJ-NPR-Z0-9]{17}$/i;
function analyze(s){
  const R=[], A=(info,data,status,note,go,src)=>R.push({info,data,status,note:note||'',go,src:src||''}), sr=k=>(S.src&&S.src[k])||'', X=S.ext||null;
  const n=s.units.length, bl=num(s.blTotal), blq=num(s.blQty);
  const list=(s.blChassis||'').split('\n').map(x=>x.trim().toUpperCase()).filter(Boolean);
  const it=x=>/^\d{1,3}([.,]\d{3})+$/.test(x)?Number(x.replace(/[.,]/g,'')):num(x);
  const mou={}; (s.mou||'').split('\n').forEach(l=>{const m=l.match(/([A-HJ-NPR-Z0-9]{17})\s+([\d.,]+)(?:\s+[₩W]?\s*([\d.,]+))?/i); if(m)mou[m[1].toUpperCase()]={kg:it(m[2]),price:m[3]?it(m[3]):null};});
  [['B/L','B/L'],['MoU / Perjanjian Hibah','MoU'],['NPWP','NPWP']].forEach(([t,nm])=>A('Dokumen sumber · '+nm,FT.has(t)?'Terlampir':'Belum diunggah',FT.has(t)?'VERIFIED':'MISSING',FT.has(t)?'':'Unggah di tab Isi Data, bagian Unggah dokumen',{id:'fType'}));
  A('Jumlah unit',n+' unit',!n?'MISSING':blq?(n===blq?'VERIFIED':'CONFLICT'):'NEED CONFIRMATION',blq?'B/L: '+blq+' unit':'Isi "Jumlah unit di B/L" untuk mencocokkan',{id:'h_blQty'},sr('blQty'));
  A('No. invoice',s.invNo||'-',s.invNo.trim()?'NEED CONFIRMATION':'MISSING','Cocokkan dengan keterangan invoice di B/L',{id:'h_invNo'},sr('invNo'));
  A('Shipper',(s.shipper[0]||'-'),s.shipper.join('').trim()?'NEED CONFIRMATION':'MISSING','Cocokkan dengan B/L',{id:'shipper'},sr('shipper'));
  A('Consignee',(s.consignee[0]||'-'),s.consignee.join('').trim()?'NEED CONFIRMATION':'MISSING','Harus sesuai B/L dan NPWP',{id:'consignee'},sr('consignee'));
  A('From / To',(s.from||'-')+' → '+(s.to||'-'),(s.from.trim()&&s.to.trim())?'NEED CONFIRMATION':'MISSING','Port of loading di B/L dapat berbeda dari template',{id:'h_from'},[sr('from'),sr('to')].filter(Boolean).join(' · '));
  A('Departure date / Vessel',(s.depart||'-')+' / '+(s.vessel||'-'),(s.depart.trim()&&s.vessel.trim())?'NEED CONFIRMATION':'MISSING','Opsional; boleh dikosongkan sesuai keputusan TML',{id:'h_depart'});
  const tw=s.units.reduce((a,u)=>a+num(u.kg),0);
  A('Total berat',fmt(tw)+' kg',!bl?'NEED CONFIRMATION':(tw===bl?'VERIFIED':'CONFLICT'),bl?'B/L: '+fmt(bl)+' kg':'Isi berat total B/L untuk mencocokkan',{id:'h_blTotal'},sr('blTotal'));
  const seen={};
  S.units.forEach((su,i)=>{ const u=s.units[i], L='Unit '+(i+1), v=u.vin.trim().toUpperCase(); seen[v]=(seen[v]||0)+1;
    const G=k=>({u:su.id,k}), su_=k=>sr('u:'+v+':'+k);
    let st,nt=''; if(!v){st='MISSING';} else if(!VINRE.test(v)){st='CONFLICT';nt='VIN harus 17 karakter alfanumerik tanpa I, O, Q ('+v.length+' karakter)';} else if(list.length){st=list.includes(v)?'VERIFIED':'CONFLICT';nt=list.includes(v)?'Ada di daftar chassis B/L':'Tidak ada di daftar chassis B/L';} else {st='NEED CONFIRMATION';nt='Format valid; cocokkan dengan B/L';}
    A(L+' · Chassis / VIN',v||'-',st,nt,G('vin'),su_('vin'));
    A(L+' · Brand / Model / Year',[u.brand,u.model,u.year].filter(Boolean).join(' / ')||'-',(u.brand&&u.model&&u.year)?((/^\d{4}$/.test(String(u.year).trim())&&+u.year>=1980&&+u.year<=2027)?'NEED CONFIRMATION':'CONFLICT'):'MISSING',(u.brand&&u.model&&u.year)?'Cocokkan dengan MoU dan spesifikasi':'',G('brand'),su_('brand'));
    A(L+' · Cylinder / Engine / Color',[u.cyl,u.eng,u.color].filter(Boolean).join(' / ')||'-',(u.cyl&&u.eng&&u.color)?'NEED CONFIRMATION':'MISSING','Belum ada di B/L dan MoU; ikuti template atau spesifikasi kendaraan',G('cyl'),su_('cyl'));
    A(L+' · Nomor mesin',u.engNo||'-',u.engNo.trim()?'NEED CONFIRMATION':'MISSING','Opsional; tidak tercetak di template',G('engNo'),su_('engNo'));
    A(L+' · Dimensi',u.dim||'-',u.dim.trim()?'NEED CONFIRMATION':'MISSING','Opsional; tidak tercetak di template',G('dim'),su_('dim'));
    const mm=mou[v], mw=mm?mm.kg:null; A(L+' · Berat',fmt(num(u.kg))+' kg',!num(u.kg)?'MISSING':(mw==null?'NEED CONFIRMATION':(mw===num(u.kg)?'VERIFIED':'CONFLICT')),mw==null?'Isi data MoU untuk mencocokkan':'MoU: '+fmt(mw)+' kg',G('kg'),su_('kg'));
    const pr=num(u.price), mp=mm?mm.price:null, same=s.units.findIndex((x,j)=>j!==i&&num(x.price)>0&&num(x.price)===pr);
    A(L+' · Harga satuan','₩'+fmt(pr).replace('.00',''),pr<=0?'MISSING':(mp==null?'NEED CONFIRMATION':(mp===pr?'VERIFIED':'CONFLICT')),mp!=null?'MoU: ₩'+fmt(mp).replace('.00',''):(same>=0?'Sama dengan Unit '+(same+1)+'; pastikan benar sesuai lampiran MoU':'Cocokkan dengan lampiran MoU; atau isi harga di data MoU'),G('price'),su_('price'));
  });
  if(X){
    if(X.bl&&X.npwp&&X.bl.npwp&&X.npwp.npwp15){ const same=X.bl.npwp===X.npwp.npwp15; A('NPWP consignee: B/L ↔ kartu NPWP',X.bl.npwp+' ↔ '+X.npwp.npwp15,same?'VERIFIED':'CONFLICT',same?'Nomor sama pada B/L dan kartu NPWP':'Nomor berbeda; periksa dokumen','','B/L hlm.1 · NPWP hlm.1 (OCR)'); }
    else if(X.bl&&X.bl.npwp) A('NPWP consignee',X.bl.npwp,'NEED CONFIRMATION','Kartu NPWP belum dibaca atau diunggah','','B/L hlm.1');
    if(X.npwp&&X.npwp.name){ const t=BACA.norm(s.consignee.join('')), ws=X.npwp.name.toUpperCase().split(/[^A-Z]+/).filter(w=>w.length>=4), hit=ws.filter(w=>t.includes(w)).length, rt=ws.length?hit/ws.length:0;
      A('Nama consignee ↔ kartu NPWP',Math.round(rt*100)+'% kata cocok',rt>=0.7?'VERIFIED':'NEED CONFIRMATION','OCR kartu NPWP dapat berisi salah baca kecil','','NPWP hlm.1 (OCR)'); }
    if(X.mou&&X.bl){ const bv=new Set(X.bl.chassis), out=X.mou.units.filter(u=>!bv.has(u.vin)); if(out.length)A('Unit di MoU yang tidak ada di B/L ini',out.length+' unit: '+out.map(u=>'…'+u.vin.slice(-6)).join(', '),'NEED CONFIRMATION','Dikirim terpisah pada pengiriman lain?','','MoU hlm.2 · B/L hlm.1');
      A('Jumlah chassis B/L = jumlah unit B/L',X.bl.chassis.length+' chassis, '+X.bl.qty+' unit',X.bl.chassis.length===X.bl.qty?'VERIFIED':'CONFLICT','','',X.bl.src); }
  }
  Object.keys(seen).filter(k=>k&&seen[k]>1).forEach(k=>A('VIN kembar',k,'CONFLICT','Dipakai oleh '+seen[k]+' unit'));
  list.filter(v=>!s.units.some(u=>u.vin.trim().toUpperCase()===v)).forEach(v=>A('Chassis di B/L belum ada di form',v,'CONFLICT','Tambahkan unit'));
  return R;
}
function check(){ const s=read(), r=[], v=s.units.map(u=>u.vin.trim());
  r.push([s.units.length>0,'Minimal 1 unit ('+s.units.length+' unit)']);
  v.forEach((x,i)=>r.push([VINRE.test(x),'Unit '+(i+1)+': VIN 17 karakter valid ('+x.length+')']));
  r.push([new Set(v).size===v.length,'Tidak ada VIN kembar']);
  const t=s.units.reduce((a,u)=>a+num(u.kg),0), bl=num(s.blTotal);
  r.push([bl>0&&t===bl,'Total berat unit '+fmt(t)+' kg'+(bl?(t===bl?' = berat B/L':' ≠ berat B/L '+fmt(bl)):' (isi berat B/L untuk mencocokkan)')]);
  if(num(s.blQty))r.push([num(s.blQty)===s.units.length,'Jumlah unit '+s.units.length+(num(s.blQty)===s.units.length?' = B/L':' ≠ B/L '+num(s.blQty))]);
  r.push([s.units.every(u=>num(u.price)>0),'Harga satuan semua unit terisi']);
  r.push([S.shipper.length<=5&&S.consignee.length<=8,'Shipper maks. 5 baris, consignee maks. 8 baris']);
  r.push([!!s.invNo.trim(),'No. invoice terisi']);
  $('checks').innerHTML=r.map(([ok,t])=>'<div class="chk '+(ok?'ok':'wn')+'">'+(ok?'✓ ':'⚠ ')+t+'</div>').join(''); renderRep(); }
const stc=s=>s.split(' ')[0];
function repHTML(act){ const R=analyze(read());
  return '<thead><tr><th>No</th><th>Informasi</th><th>Data pada form</th><th>Sumber dokumen</th><th>Status</th><th>Catatan</th>'+(act?'<th></th>':'')+'</tr></thead><tbody>'+R.map((r,i)=>'<tr class="r-'+stc(r.status)+'"><td>'+(i+1)+'</td><td>'+esc(r.info)+'</td><td>'+esc(r.data)+'</td><td>'+esc(r.src||'')+'</td><td class="st s-'+stc(r.status)+'">'+r.status+'</td><td>'+esc(r.note)+'</td>'+(act?'<td>'+(r.go?'<button class="sm" data-edit="'+esc(JSON.stringify(r.go)).replace(/"/g,'&quot;')+'">Edit</button>':'')+'</td>':'')+'</tr>').join('')+'</tbody>'; }
function renderRep(){ $('rep').innerHTML=repHTML(true); const R=analyze(read()),c={};R.forEach(r=>c[stc(r.status)]=(c[stc(r.status)]||0)+1);
  $('repSum').innerHTML=['VERIFIED','NEED','CONFLICT','MISSING'].map(k=>'<span class="chip s-'+k+'">'+(k==='NEED'?'NEED CONFIRMATION':k)+': '+(c[k]||0)+'</span>').join(''); }
function fname(p,ext){ return 'TML_'+p+'_'+((S.name||'').trim().toUpperCase().replace(/[^A-Z0-9]+/g,'_').replace(/^_|_$/g,'')||'PROYEK')+'.'+(ext||'xlsx'); }
const KN={CI:'COMMERCIAL_INVOICE',PL:'PACKING_LIST'};
function saveFile(name,data,type){ const b=new Blob([data],{type:type||'application/octet-stream'}), u=URL.createObjectURL(b), a=document.createElement('a'); a.href=u; a.download=name; document.body.append(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(u),4000); }
async function dl(kind){
  if(!S.units.length){note('Tambahkan minimal 1 unit.');return;}
  try{ note('Menyiapkan file…'); const s=read(), buf=await buildXlsx(window.TML_TEMPLATES[kind],kind,s);
    saveFile(fname(KN[kind]),buf,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); learn(s); persistOpts(); logA('Excel '+KN[kind]+' diunduh'); markDone(kind); note('Selesai: '+fname(KN[kind])+'. Periksa file di Excel.');
  }catch(e){ note('Gagal: '+(e.message||e)); } }
/* ---- Penyimpanan: browser (default) + Google Apps Script (opsional) ---- */
function jsonp(params){ const url=cfg().url; params={...params,code:cfg().code||''}; return new Promise((res,rej)=>{ const cbn='tml_cb_'+Date.now()+Math.floor(Math.random()*1e4), sc=document.createElement('script');
  const done=()=>{clearTimeout(t);delete window[cbn];sc.remove();}; const t=setTimeout(()=>{done();rej(new Error('Waktu habis'));},20000);
  window[cbn]=d=>{done();res(d);}; sc.onerror=()=>{done();rej(new Error('Gagal terhubung ke Apps Script'));};
  sc.src=url+(url.includes('?')?'&':'?')+new URLSearchParams({...params,callback:cbn}); document.head.append(sc); }); }
async function loadSaved(){ let L=Object.values(LS.get('tml_projects',{}));
  if(cfg().url){ try{ const r=await jsonp({action:'list'}); if(r&&r.ok) L=r.items.map(x=>({id:x.id,name:x.name,updated:+x.updated,data:x.data})); }catch(e){ note('Penyimpanan bersama tidak terjangkau: '+e.message+'. Menampilkan proyek di perangkat ini.'); } }
  SAVED=L.sort((a,b)=>b.updated-a.updated);
  SAVED.forEach(p=>{try{learn(JSON.parse(p.data));addOpt('name',p.name);}catch(e){}});
  const cur=$('projSel').value;
  $('projSel').innerHTML='<option value="">— Proyek tersimpan ('+SAVED.length+') —</option>'+SAVED.map(p=>'<option value="'+esc(p.id)+'">'+esc(p.name)+'</option>').join(''); $('projSel').value=cur; }
const pid=n=>(n.trim().toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'')||'proyek');
function load(p){ S=hydrate(JSON.parse(p.data)); render(); draft(); }
async function save(){ if(!S.name.trim()){note('Isi nama proyek dulu.');goEdit({id:'h_name'});return false;}
  logA(S.confirmed?'Proyek disimpan (terkonfirmasi)':'Proyek disimpan'); S.savedRev=S.rev; S.savedAt=Date.now();
  const d=clean(read()); learn(d); persistOpts(); const rec={id:pid(S.name),name:S.name,updated:Date.now(),data:JSON.stringify(d)};
  const all=LS.get('tml_projects',{}); all[rec.id]=rec; LS.set('tml_projects',all); let t='✓ Tersimpan di perangkat ini: '+S.name;
  if(cfg().url){ try{ const r=await jsonp({action:'save',id:rec.id,name:rec.name,updated:rec.updated,data:rec.data}); t=r&&r.ok?'✓ Tersimpan di perangkat ini dan di penyimpanan bersama: '+S.name:'Tersimpan di perangkat ini. Penyimpanan bersama menolak: '+((r&&r.error)||'tidak diketahui'); }catch(e){ t='Tersimpan di perangkat ini. Penyimpanan bersama gagal: '+e.message; } }
  note(t); draft(); await loadSaved(); $('projSel').value=rec.id; $('bDelP').hidden=false; return true; }
/* ---- Impor ---- */
async function importBuf(buf){ const sh=parseSheet(await readXlsx(buf)); let added=0,upd=0;
  ['invNo','invDate','depart','vessel','from','to'].forEach(k=>{if(sh[k]!==undefined&&sh[k]!=='')S[k]=sh[k];});
  if(sh.shipper.length)S.shipper=sh.shipper; if(sh.consignee.length)S.consignee=sh.consignee;
  sh.units.forEach(x=>{ let u=S.units.find(y=>y.vin.trim()&&y.vin.trim()===x.vin.trim());
    if(!u){ const e=S.units.length===1&&!S.units[0].vin.trim()&&!S.units[0].brand.trim(); if(e){u=S.units[0];} else {u=blank();S.units.push(u);} u.edit=false; added++; } else upd++;
    Object.keys(x).forEach(k=>{if(x[k]!==''&&x[k]!=='0')u[k]=x[k];}); });
  render(); touch(); return {kind:sh.kind,added,upd}; }
/* ---- Lampiran dokumen (disimpan di browser perangkat ini lewat IndexedDB) ---- */
const IDB={db:null,open(){return this.db?Promise.resolve(this.db):new Promise((res,rej)=>{if(!window.indexedDB)return rej(new Error('IndexedDB tidak tersedia'));const r=indexedDB.open('tml_hibah',1);r.onupgradeneeded=()=>r.result.createObjectStore('files',{keyPath:'id',autoIncrement:true}).createIndex('pid','pid');r.onsuccess=()=>{this.db=r.result;res(this.db);};r.onerror=()=>rej(r.error);});},
  async tx(mode,fn){const db=await this.open();return new Promise((res,rej)=>{const t=db.transaction('files',mode),q=fn(t.objectStore('files'));t.oncomplete=()=>res(q&&q.result);t.onerror=()=>rej(t.error);});},
  add(r){return this.tx('readwrite',s=>s.add(r));},list(p){return this.tx('readonly',s=>s.index('pid').getAll(p));},get(i){return this.tx('readonly',s=>s.get(i));},del(i){return this.tx('readwrite',s=>s.delete(i));}};
const kb=n=>n>1048576?(n/1048576).toFixed(1)+' MB':Math.max(1,Math.round(n/1024))+' KB';
async function renderFiles(){ let L=[]; try{L=await IDB.list(S.key)||[];}catch(e){} FT=new Set(L.map(f=>f.type));
  const h=L.length?L.map(f=>'<div class="fr"><span><b>'+esc(f.type)+'</b> · '+esc(f.name)+' <i>'+kb(f.size)+'</i></span><span class="row"><button class="sm" data-f="view" data-id="'+f.id+'">Lihat</button>'+(/\.xlsx$/i.test(f.name)?'<button class="sm" data-f="imp" data-id="'+f.id+'">Impor ke form</button>':'')+'<button class="sm d" data-f="del" data-id="'+f.id+'">Hapus</button></span></div>').join(''):'<div class="sub">Belum ada lampiran.</div>';
  $('files').innerHTML=h; $('filesV').innerHTML=h;
  $('docChk').innerHTML=[['B/L','B/L'],['MoU / Perjanjian Hibah','MoU'],['NPWP','NPWP']].map(([t,n])=>'<span class="dchip '+(FT.has(t)?'ok':'miss')+'">'+(FT.has(t)?'✓ ':'⚠ ')+n+(FT.has(t)?' terlampir':' belum diunggah')+'</span>').join(''); if(S)renderRep(); }
let viewURL=null;
async function viewFile(id){ const f=await IDB.get(id); if(!f)return; if(viewURL)URL.revokeObjectURL(viewURL); viewURL=URL.createObjectURL(f.blob);
  const pdf=/pdf/i.test(f.mime)||/\.pdf$/i.test(f.name), img=/^image\//.test(f.mime);
  $('vTitle').textContent=f.type+' · '+f.name; $('vBody').innerHTML=pdf?'<iframe src="'+viewURL+'" title="Dokumen"></iframe>':img?'<img src="'+viewURL+'" alt="Dokumen">':'<p class="sub">Format ini tidak bisa ditampilkan di sini.</p>';
  $('vOpen').href=viewURL; $('vOpen').download=f.name; $('viewer').hidden=false; }
$('fFiles').onchange=async e=>{ const fs=[...e.target.files]; e.target.value=''; let ok=0;
  for(const f of fs){ try{ await IDB.add({pid:S.key,type:$('fType').value,name:f.name,size:f.size,mime:f.type,blob:f,added:Date.now()}); ok++; }catch(x){ note('Gagal menyimpan lampiran: '+(x.message||x)); } }
  if(ok)note(ok+' lampiran ditambahkan. Lampiran tersimpan di browser perangkat ini dan tidak ikut penyimpanan online.'); renderFiles(); };
$('vClose').onclick=()=>{$('viewer').hidden=true;$('vBody').innerHTML='';};

/* ---- Baca dokumen otomatis: teks asli PDF + OCR di browser, hasil ditinjau dulu sebelum diisi ---- */
let EXT=null, PATCH=null; const REQT=['B/L','MoU / Perjanjian Hibah','NPWP'];
const fmtPort=(t,sep)=>{ const p=String(t||'').split(',').map(x=>x.trim()).filter(Boolean); if(!p.length)return ''; let c=p[1]?BACA.title(p[1]):''; if(/^korea$/i.test(c))c='Republic of Korea'; return c?BACA.title(p[0])+sep+c:BACA.title(p[0]); };
const wrapTxt=(t,n)=>{ const o=[]; let l=''; String(t).split(/\s+/).filter(Boolean).forEach(w=>{ if((l+' '+w).trim().length>n&&l){o.push(l);l=w;} else l=(l+' '+w).trim(); }); if(l)o.push(l); return o; };
function joinWrapped(lines,vocab){ if(!lines.length)return ''; let out=lines[0]; for(let i=1;i<lines.length;i++){ const last=(out.split(/\s+/).pop()||'').replace(/[^A-Za-z0-9]/g,''), first=(lines[i].split(/\s+/)[0]||'').replace(/[^A-Za-z0-9]/g,''), c=(last+first).toUpperCase(); out+=(vocab&&vocab.has(c)&&!vocab.has(last.toUpperCase()))?lines[i]:' '+lines[i]; } return out; }
function buildPatch(x){
  const bl=x.bl, mo=x.mou, np=x.npwp, R=[], W=[], OK=[], tp=(D.units&&D.units[0])||{};
  const row=(group,label,value,src,set,o={})=>R.push({group,label,value:Array.isArray(value)?value.join(' / '):String(value),src,set,use:o.use!==false,kind:o.kind||'doc',note:o.note||''});
  const put=(s,k,v,src,kind)=>{ s[k]=v; s.src[k]=src; s.marks['h_'+k]=kind||'ai'; }, putTA=(s,k,v,src,kind)=>{ s[k]=v; s.src[k]=src; s.marks[k]=kind||'ai'; };
  let UT=[];
  if(!bl) W.push('B/L belum diunggah atau tidak terbaca. Unggah B/L dengan jenis "B/L" (PDF asli yang masih berisi teks paling akurat).');
  else { const sB=bl.src;
    if(!S.name.trim()&&bl.no) row('Data proyek','Nama proyek (saran)','HIBAH_'+bl.no,sB,s=>{ s.name='HIBAH_'+bl.no; s.marks.h_name='ai'; });
    if(bl.invNo) row('Data proyek','No. invoice',bl.invNo,sB,s=>put(s,'invNo',bl.invNo,sB));
    if(bl.invDate) row('Data proyek','Tanggal invoice',bl.invDate,'Diturunkan dari No. invoice di B/L',s=>put(s,'invDate',bl.invDate,'Diturunkan dari No. invoice di B/L'));
    const from=fmtPort(bl.pol,','), to=fmtPort(bl.pod,'-');
    if(from) row('Data proyek','From',from,sB+' (Port of loading)',s=>put(s,'from',from,sB+' (Port of loading)'),{note:(D.from&&BACA.norm(D.from)!==BACA.norm(from))?'Template contoh memakai: '+D.from+'. Dipakai: sesuai B/L.':''});
    if(to) row('Data proyek','To',to,sB+' (Port of discharge)',s=>put(s,'to',to,sB+' (Port of discharge)'));
    if(bl.depart) row('Data proyek','Departure date',bl.depart,sB+' (Laden on board)',s=>put(s,'depart',bl.depart,sB+' (Laden on board)'),{use:false,note:'Keputusan TML sebelumnya: dikosongkan. Centang bila ingin diisi.'});
    if(bl.vessel) row('Data proyek','Vessel',bl.vessel,sB,s=>put(s,'vessel',bl.vessel,sB),{use:false,note:'Keputusan TML sebelumnya: dikosongkan. Centang bila ingin diisi.'});
    if(bl.grossKg) row('Pembanding','Berat total B/L (kg)',bl.grossKg,sB,s=>put(s,'blTotal',String(bl.grossKg),sB));
    if(bl.qty) row('Pembanding','Jumlah unit di B/L',bl.qty,sB,s=>put(s,'blQty',String(bl.qty),sB));
    if(bl.chassis.length) row('Pembanding','Daftar chassis B/L',bl.chassis,sB,s=>putTA(s,'blChassis',bl.chassis.join('\n'),sB));
    const isT=l=>/^(TEL|FAX)\b/i.test(l), shName=bl.shipperRaw.filter(l=>!isT(l)), shTel=(bl.shipperRaw.find(l=>/^TEL\b/i.test(l))||'').replace(/^TEL\s*:\s*/i,'').trim();
    if(shName.length){ const shMatch=D.shipper&&D.shipper.length>=3&&BACA.norm(shName.join(' '))===BACA.norm(D.shipper.slice(0,3).join(''));
      let sh=shMatch?D.shipper.slice(0,3):wrapTxt(shName.join(' '),58); if(shTel)sh=sh.concat(['TEL : '+shTel+' ']); if(bl.shipFax)sh=sh.concat(['FAX : '+bl.shipFax]); while(sh.length>5){ sh.splice(0,2,sh[0]+' '+sh[1]); }
      row('Shipper dan Consignee','Shipper',sh,sB+(shMatch?' (nama dan alamat sama dengan format TML)':''),s=>putTA(s,'shipper',sh,sB+(shMatch?' (cocok format TML)':''))); }
    if(bl.consigneeRaw.length){ const cText=joinWrapped(bl.consigneeRaw,x.vocab)+' '+(bl.consAddr2||''), cMatch=D.consignee&&D.consignee.length>=4&&BACA.norm(cText)===BACA.norm(D.consignee.slice(0,4).join(''));
      let cons=cMatch?D.consignee.slice(0,4):wrapTxt(cText,60); const ct=[]; if(bl.tel)ct.push('Tel : '+bl.tel); if(bl.emails.length)ct.push('Email : '+bl.emails.join(', ')); if(bl.attn)ct.push('ATTN: '+bl.attn); cons=cons.concat(ct); while(cons.length>8){ cons.splice(0,2,cons[0]+' '+cons[1]); }
      const src=sB+(np?' · NPWP hlm.1':'')+(cMatch?' (nama dan alamat sama dengan format TML)':'');
      row('Shipper dan Consignee','Consignee',cons,src,s=>putTA(s,'consignee',cons,src),{note:(cMatch&&D.consignee.slice(4).join('')&&BACA.norm(D.consignee.slice(4).join(''))!==BACA.norm(ct.join('')))?'Telepon/email/ATTN memakai data B/L (berbeda dari template contoh).':''}); }
    UT=bl.chassis.map(v=>({vin:v,brand:bl.brand||'',model:bl.model||'',year:bl.year||'',cyl:'',eng:'',color:'',kg:'',price:'',engNo:'',dim:'',_ai:{},_tp:{},_src:{}}));
    const per=(bl.grossKg&&bl.qty&&bl.qty===bl.chassis.length)?bl.grossKg/bl.qty:0;
    UT.forEach((u,i)=>{ const m=mo&&mo.units.find(z=>z.vin===u.vin), L='Unit '+(i+1), y=(m&&m.year)||u.year, sI=sB+(m?' · MoU hlm.2 (OCR)':'');
      row(L,'Identitas',[u.vin,[u.brand,u.model].filter(Boolean).join(' '),y].join(' · '),sI,()=>{ u._use=true; u.year=y; ['vin','brand','model','year'].forEach(k=>u._ai[k]=1); u._src.vin=sI; u._src.brand=sB; u._src.year=sI; });
      if(per) row(L,'Berat (kg)',fmt(per)+' kg',sB+' (total ÷ jumlah unit, aturan TML)',()=>{ u.kg=String(per%1?per.toFixed(2):per); u._ai.kg=1; u._src.kg=sB+' (total ÷ jumlah unit)'; });
      if(m&&m.priceOk&&m.price) row(L,'Harga satuan (₩)',fmt(m.price).replace('.00',''),'MoU lampiran (OCR, lolos cek jumlah total)',()=>{ u.price=String(m.price); u._ai.price=1; u._src.price='MoU lampiran (OCR, lolos cek jumlah total)'; });
      if(m&&m.dim) row(L,'Dimensi (mm)',m.dim.replace(/x/g,' × '),'MoU lampiran (OCR)',()=>{ u.dim=m.dim.replace(/x/g,' x ')+' mm'; u._ai.dim=1; u._src.dim='MoU lampiran (OCR)'; });
      if(tp.cyl||tp.eng||tp.color) row(L,'Cylinder / Engine / Color',[tp.cyl,tp.eng,tp.color].filter(Boolean).join(' / '),'Contoh template TML (bukan dari dokumen)',()=>{ u.cyl=tp.cyl||''; u.eng=tp.eng||''; u.color=tp.color||''; ['cyl','eng','color'].forEach(k=>u._tp[k]=1); u._src.cyl='Contoh template TML (bukan dari dokumen)'; },{kind:'tpl',note:'Tidak ada di B/L, MoU, maupun NPWP; periksa dengan spesifikasi kendaraan.'}); });
    if(bl.chassis.length!==bl.qty) W.push('Jumlah chassis di B/L ('+bl.chassis.length+') tidak sama dengan jumlah unit ('+bl.qty+'); periksa B/L.');
    if(per) OK.push('Berat per unit = total B/L ('+fmt(bl.grossKg)+' kg) ÷ '+bl.qty+' unit = '+fmt(per)+' kg (aturan TML).'); }
  if(mo){ const out=mo.units.filter(u=>!bl||!bl.chassis.includes(u.vin)); if(out.length)W.push('MoU memuat '+out.length+' unit yang tidak ada di B/L ini ('+out.map(u=>'…'+u.vin.slice(-6)).join(', ')+'): tidak diisi; dikirim terpisah?');
    if(!mo.priceValidated) W.push('Harga, berat, dan dimensi per unit di tabel lampiran MoU tidak terbaca otomatis (scan kecil). Isi manual sambil membuka MoU (tombol Lihat). Total MoU terbaca: ₩'+fmt(mo.totalValue||0).replace('.00','')+' untuk '+(mo.totalUnits||'?')+' unit.');
    if(mo.units.some(u=>u.snapped))W.push('Sebagian nomor chassis MoU dikoreksi 1 karakter agar cocok dengan B/L; periksa.');
    OK.push('MoU terbaca: '+mo.units.length+' chassis'+(mo.no?', nomor '+mo.no:'')+(mo.date?', tanggal '+mo.date:'')+'.'); }
  else W.push('MoU belum diunggah atau tidak terbaca: harga dan berat per unit harus diisi manual.');
  if(np&&bl&&bl.npwp&&np.npwp15){ if(bl.npwp===np.npwp15)OK.push('NPWP consignee sama pada B/L dan kartu NPWP ('+np.npwp15+').'); else W.push('NPWP B/L ('+bl.npwp+') berbeda dari kartu NPWP ('+np.npwp15+').'); }
  else if(!np) W.push('Kartu NPWP belum diunggah atau tidak terbaca: kecocokan NPWP consignee tidak bisa diperiksa.');
  (x.errors||[]).forEach(e=>W.push('Gagal membaca: '+e));
  return {rows:R,warns:W,oks:OK,finish:s=>{ const used=UT.filter(u=>u._use); if(!used.length)return; s.units=used.map(u=>{ const {_use,_ai,_tp,_src,...f}=u; Object.keys(_src).forEach(k=>{s.src['u:'+u.vin+':'+k]=_src[k];}); return blank({...f,edit:true,ai:_ai,tp:_tp}); }); }}; }
const pick=(o,ks)=>{const r={};ks.forEach(k=>{if(o[k]!==undefined)r[k]=o[k];});return r;};
const compactExt=x=>({at:Date.now(),bl:x.bl?pick(x.bl,['no','vessel','pol','pod','grossKg','qty','chassis','brand','model','year','invNo','invDate','npwp','hs','depart','containers','src','file']):null,mou:x.mou?{...pick(x.mou,['no','date','totalUnits','totalValue','priceValidated','src','file']),units:x.mou.units.map(u=>pick(u,['vin','km','year','brand']))}:null,npwp:x.npwp?pick(x.npwp,['npwp15','npwp16','name','registered','kpp','src','file']):null});
function showPatch(){ const P=PATCH, g={}, chip=(ok,t)=>'<span class="dchip '+(ok?'ok':'miss')+'">'+(ok?'✓ ':'⚠ ')+t+'</span>';
  let h='<div class="chips">'+chip(EXT.bl,'B/L'+(EXT.bl?' terbaca (teks asli PDF)':' tidak terbaca'))+chip(EXT.mou,'MoU'+(EXT.mou?' terbaca (OCR)':' tidak terbaca'))+chip(EXT.npwp,'NPWP'+(EXT.npwp?' terbaca (OCR)':' tidak terbaca'))+'</div>';
  if(P.oks.length)h+='<ul class="oklist">'+P.oks.map(w=>'<li>✓ '+esc(w)+'</li>').join('')+'</ul>';
  if(P.warns.length)h+='<div class="warnbox"><b>Perlu perhatian</b><ul>'+P.warns.map(w=>'<li>'+esc(w)+'</li>').join('')+'</ul></div>';
  h+='<div class="tw"><table class="rt"><thead><tr><th></th><th>Isian</th><th>Nilai terbaca</th><th>Sumber</th></tr></thead><tbody>'; let last='';
  P.rows.forEach((r,i)=>{ if(r.group!==last){ h+='<tr class="grp"><td colspan="4">'+esc(r.group)+'</td></tr>'; last=r.group; }
    h+='<tr><td><input type="checkbox" data-pi="'+i+'"'+(r.use?' checked':'')+'></td><td>'+esc(r.label)+(r.kind==='tpl'?' <span class="ub wn">dari template</span>':'')+'</td><td>'+esc(r.value)+(r.note?'<div class="sub">'+esc(r.note)+'</div>':'')+'</td><td>'+esc(r.src)+'</td></tr>'; });
  h+='</tbody></table></div><div class="row" style="margin-top:12px"><button class="p" id="bApply">✔ Isi ke Data Proyek &amp; Unit</button><button class="sm" id="bReadX">Tutup</button></div><p class="sub">Mengisi akan menggantikan isian unit saat ini. Hasilnya tetap bisa diedit. Kuning = dari dokumen, biru = dari template contoh.</p>';
  $('readOut').innerHTML=h; $('readOut').hidden=false; $('bApply').onclick=applyPatch; $('bReadX').onclick=()=>{$('readOut').hidden=true;}; }
function applyPatch(){ const ch=[...document.querySelectorAll('#readOut [data-pi]')].map(c=>c.checked); PATCH.rows.forEach((r,i)=>{r.use=ch[i];}); S.src=S.src||{}; S.marks=S.marks||{};
  PATCH.rows.forEach(r=>{ if(r.use)r.set(S); }); PATCH.finish(S); S.ext=compactExt(EXT); S.readAt=Date.now(); logA('Data proyek dan unit diisi otomatis dari dokumen'); render(); touch(); $('readOut').hidden=true;
  note('✓ Data terisi dari dokumen. Periksa isian berwarna (kuning = dari dokumen, biru = template contoh), edit bila perlu, lalu Simpan.'); window.scrollTo({top:Math.max(0,$('hdr').getBoundingClientRect().top+scrollY-130),behavior:'smooth'}); }
async function readDocs(){ if(typeof BACA==='undefined'||!window.pdfjsLib||!window.Tesseract){ note('Pembaca dokumen belum terpasang: unggah folder "baca" dan berkas baca.js ke repository.'); return; }
  pdfjsLib.GlobalWorkerOptions.workerSrc='baca/pdf.worker.min.js';
  const fs=(await IDB.list(S.key))||[], docs=fs.filter(f=>REQT.includes(f.type)).map(f=>({id:f.id,name:f.name,type:f.type,file:new File([f.blob],f.name,{type:f.mime||(/\.pdf$/i.test(f.name)?'application/pdf':'')})}));
  if(!docs.length){ note('Unggah dulu dokumen dengan jenis B/L, MoU / Perjanjian Hibah, atau NPWP.'); return; }
  const t0=Date.now(); $('readLog').textContent='Membaca…'; const x=await BACA.readAll(docs,m=>{$('readLog').textContent=m;});
  EXT=x; PATCH=buildPatch(x); showPatch(); $('readLog').textContent='Selesai dalam '+Math.round((Date.now()-t0)/1000)+' detik. Periksa hasil di bawah, lalu klik "Isi ke Data Proyek & Unit".'; logA('Dokumen dibaca otomatis ('+docs.length+' berkas)'); }
$('bRead').onclick=()=>busy('bRead',readDocs);

/* ---- Menu ▾ dan event ---- */
function closeMenus(){document.querySelectorAll('.menu').forEach(m=>m.remove());}
function openMenu(btn){ closeMenus(); const k=btn.dataset.k, tgt=$(btn.dataset.t), w=btn.closest('.w'); const L=[...(OPT[k]||[])].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  const m=document.createElement('div'); m.className='menu'; const cbx=btn.closest('.cb'); m.style.top=((cbx?cbx.offsetTop+cbx.offsetHeight:btn.offsetTop+btn.offsetHeight)+2)+'px';
  if(!L.length){const e=document.createElement('div');e.className='em';e.textContent='Belum ada pilihan, ketik saja.';m.append(e);}
  L.forEach(v=>{const d=document.createElement('div');d.textContent=v.replace(/\n/g,' ⏎ ');d.title=v;d.onclick=()=>{tgt.value=v;tgt.dispatchEvent(new Event('input',{bubbles:true}));closeMenus();};m.append(d);});
  w.append(m); }
function goEdit(g){ tab('isi'); let id=g.id; if(g.u){const u=S.units.find(x=>x.id===g.u); if(u){u.edit=true;renderUnits();} id='u'+g.u+g.k;}
  setTimeout(()=>{const el=$(id); if(el){el.scrollIntoView({block:'center',behavior:'smooth'});el.focus({preventScroll:true});el.classList.add('flash');setTimeout(()=>el.classList.remove('flash'),1900);}},60); }
document.addEventListener('click',async e=>{
  const b=e.target.closest('button'); if(!b){ if(!e.target.closest('.menu'))closeMenus(); return; }
  if(b.dataset.t){ e.stopPropagation(); openMenu(b); return; }
  closeMenus();
  if(b.dataset.edit){ goEdit(JSON.parse(b.dataset.edit)); return; }
  if(b.dataset.f){ const id=+b.dataset.id;
    if(b.dataset.f==='view')viewFile(id);
    else if(b.dataset.f==='imp'){ const f=await IDB.get(id); try{const r=await importBuf(await f.blob.arrayBuffer());note('Impor '+(r.kind==='CI'?'Commercial Invoice':'Packing List')+': '+r.added+' unit baru, '+r.upd+' diperbarui. Periksa data di tab Isi Data.');}catch(x){note('Gagal membaca Excel: '+x.message);} }
    else if(b.dataset.f==='del'){ if(!b.dataset.sure){b.dataset.sure=1;b.textContent='Yakin hapus?';setTimeout(()=>{if(b.isConnected){delete b.dataset.sure;b.textContent='Hapus';}},3000);} else {await IDB.del(id);renderFiles();} }
    return; }
  const a=b.dataset.a; if(!a)return; const u=S.units.find(x=>x.id===+b.dataset.id); if(!u)return;
  if(a==='edit'){u.edit=true;renderUnits();} else if(a==='done'){u.edit=false;renderUnits();}
  else if(a==='del'){ if(u.sure){S.units=S.units.filter(x=>x!==u);renderUnits();touch();} else {u.sure=true;renderUnits();setTimeout(()=>{if(S.units.includes(u)&&u.sure){u.sure=false;renderUnits();}},3000);} } });
document.addEventListener('input',e=>{ const t=e.target, d=t.dataset; if(!(d.h||d.u||d.ml||d.s))return;
  if(d.h)S[d.h]=t.value; else if(d.u){const u=S.units.find(x=>x.id===+d.u);u[d.k]=t.value;if(u.ai)delete u.ai[d.k];if(u.tp)delete u.tp[d.k];}
  else if(d.ml)S[d.ml]=t.value.split('\n'); else if(d.s)S[d.s]=t.value;
  t.classList.remove('ai','tpl'); if(S.marks&&t.id)delete S.marks[t.id]; check(); touch(); });
$('bAdd').onclick=()=>{S.units.forEach(u=>u.edit=false);const p=S.units[S.units.length-1]||{};S.units.push(blank({brand:p.brand||'',model:p.model||'',year:p.year||'',cyl:p.cyl||'',eng:p.eng||'',color:p.color||'',kg:p.kg||''}));renderUnits();touch();};
$('projSel').onchange=e=>{const p=SAVED.find(x=>x.id===e.target.value);$('bDelP').hidden=!p;if(p)load(p);};
$('bNew').onclick=()=>{S=fresh();$('projSel').value='';$('bDelP').hidden=true;render();draft();note('');};
$('bEx').onclick=()=>{S=example();render();draft();};
let sureP=false;
$('bDelP').onclick=async()=>{ if(!sureP){sureP=true;$('bDelP').textContent='Yakin hapus proyek ini?';setTimeout(()=>{sureP=false;$('bDelP').textContent='Hapus proyek tersimpan';},3000);return;}
  sureP=false;$('bDelP').textContent='Hapus proyek tersimpan'; const id=$('projSel').value; if(!id)return;
  const all=LS.get('tml_projects',{}); delete all[id]; LS.set('tml_projects',all); let t='Proyek dihapus dari perangkat ini.';
  if(cfg().url){try{const r=await jsonp({action:'delete',id});t=r&&r.ok?'Proyek dihapus dari perangkat ini dan penyimpanan bersama.':'Dihapus di perangkat ini; penyimpanan bersama menolak.';}catch(e){t='Dihapus di perangkat ini; penyimpanan bersama gagal: '+e.message;}}
  note(t); await loadSaved(); $('bDelP').hidden=true; };
/* ---- Tahap 1 -> 2 -> 3 ---- */
function tab(n){ document.querySelectorAll('.tab').forEach(t=>t.hidden=t.dataset.tab!==n); document.querySelectorAll('.tabs button').forEach(b=>b.classList.toggle('on',b.dataset.go===n)); window.scrollTo(0,0);
  if(n==='ver'){renderRep();renderFiles();renderLog();} if(n==='isi')renderFiles(); updConf(); if(n==='xls'&&S.confirmed)showPreview(); }
const LBL={bSaveGo:'💾 Simpan & Lanjut ke Verifikasi',bSaveV:'💾 Simpan',bSaveX:'💾 Simpan',bConf:'✅ Telah Dikonfirmasi',bXlsx:'⬇ Unduh Excel (.xlsx)',bBoth:'Unduh CI + PL',bPrint:'🖨 Cetak',bPdf:'📄 Ekspor PDF',bRead:'🔍 Baca dokumen otomatis',bZip:'📦 Unduh Paket Arsip (.zip)',bFolder:'💾 Simpan ke Folder TML HIBAH',bDrive:'☁ Kirim ke Google Drive (eksperimental)'};
const hm=t=>t?new Date(t).toLocaleTimeString('id-ID',{hour:'2-digit',minute:'2-digit'}):'';
function setBtn(id,on,txt){ const b=$(id); if(!b)return; b.classList.toggle('done',!!on); b.textContent=on?txt:LBL[id]; }
function tk(id,st,txt,title){ const e=$(id); if(!e)return; e.className='tk '+st; e.textContent=txt; e.title=title; e.parentNode.classList.toggle('st-ok',st==='ok'); }
function applyStates(){ if(!S)return; const saved=S.rev===S.savedRev&&!!S.savedAt, d=S.done||{};
  ['bSaveGo','bSaveV','bSaveX'].forEach(id=>setBtn(id,saved,'✓ Tersimpan '+hm(S.savedAt)));
  setBtn('bRead',!!S.readAt,'✓ Dokumen dibaca '+hm(S.readAt));
  setBtn('bConf',!!S.confirmed,'✓ Terkonfirmasi '+hm(S.confirmed));
  setBtn('bXlsx',d[PK],'✓ '+PK+' terunduh'); setBtn('bBoth',d.CI&&d.PL,'✓ CI + PL terunduh'); setBtn('bPdf',d['pdf'+PK],'✓ PDF '+PK+' dibuat'); setBtn('bPrint',d.print,'✓ Cetak dibuka');
  setBtn('bZip',d.zip,'✓ Arsip ZIP terunduh'); setBtn('bFolder',d.folder,'✓ Tersimpan di folder'); setBtn('bDrive',d.drive,'✓ Terkirim ke Drive');
  tk('tkIsi',saved?'ok':'warn',saved?'✓':'●',saved?'Data tersimpan':'Ada perubahan yang belum disimpan');
  tk('tkVer',S.confirmed?'ok':'idle',S.confirmed?'✓':'○',S.confirmed?'Telah dikonfirmasi':'Belum dikonfirmasi');
  const ex=Object.keys(d).length>0; tk('tkXls',!S.confirmed?'lk':(ex?'ok':'idle'),!S.confirmed?'🔒':(ex?'✓':'○'),!S.confirmed?'Terkunci sampai dikonfirmasi':(ex?'Sudah ada hasil yang diunduh':'Belum diunduh'));
  document.querySelectorAll('.sstate').forEach(e=>{ e.className='sstate '+(saved?'ok':'wn'); e.textContent=saved?'✓ Semua perubahan tersimpan ('+hm(S.savedAt)+')':'● Ada perubahan yang belum disimpan'; }); }
async function busy(id,fn){ const b=$(id); if(!b||b.classList.contains('busy'))return; b.classList.add('busy'); b.disabled=true;
  try{ return await fn(); }catch(e){ note('Gagal: '+(e.message||e)); } finally{ b.classList.remove('busy'); b.disabled=false; if(!armed)applyStates(); } }
const markDone=k=>{ (S.done=S.done||{})[k]=Date.now(); applyStates(); };
function updConf(){ const c=S&&S.confirmed; $('confBadge').className='badge '+(c?'okb':'wnb');
  $('confBadge').textContent=c?'✅ TELAH DIKONFIRMASI · '+new Date(c).toLocaleString('id-ID')+(S.confirmedBy?' · oleh '+S.confirmedBy:''):'⚠ Belum dikonfirmasi — periksa data lalu klik "Telah Dikonfirmasi"';
  $('xlsLock').hidden=!!c; $('xlsBody').hidden=!c; $('bDrive').hidden=!cfg().url; applyStates(); }
let armed=false;
$('bSaveGo').onclick=()=>busy('bSaveGo',async()=>{ if(S.rev===S.savedRev&&S.savedAt){tab('ver');return;} if(await save())tab('ver'); });
$('bEdit').onclick=()=>tab('isi');
$('bSaveV').onclick=()=>busy('bSaveV',()=>save());
const confirmFlow=async()=>{ if(S.confirmed){tab('xls');return;} if(!S.units.length){note('Tambahkan minimal 1 unit.');return;} if(!S.name.trim()){note('Isi nama proyek dulu.');goEdit({id:'h_name'});return;}
  const R0=analyze(read()), nc=R0.filter(r=>r.status==='CONFLICT').length, nd=R0.filter(r=>r.info.startsWith('Dokumen sumber')&&r.status==='MISSING').length;
  if((nc||nd)&&!armed){armed=true;$('bConf').textContent='Masih ada '+[nc?nc+' CONFLICT':'',nd?nd+' dokumen sumber belum diunggah':''].filter(Boolean).join(' dan ')+'. Klik lagi untuk tetap mengonfirmasi';setTimeout(()=>{armed=false;updConf();},5000);return;}
  armed=false; S.confirmed=Date.now(); S.confirmedBy=opName(); logA('Telah dikonfirmasi (CONFLICT: '+nc+', dokumen sumber belum diunggah: '+nd+')'); if(await save()){updConf();tab('xls');} };
$('bConf').onclick=()=>busy('bConf',confirmFlow);
$('goVer').onclick=()=>tab('ver');
document.querySelector('.tabs').addEventListener('click',e=>{const b=e.target.closest('button');if(b)tab(b.dataset.go);});
/* ---- Tahap 3: pratinjau, unduh, cetak, PDF ---- */
async function showPreview(kind){ if(kind)PK=kind; applyStates(); $('swCI').classList.toggle('on',PK==='CI'); $('swPL').classList.toggle('on',PK==='PL');
  try{ $('prev').innerHTML='<p class="sub">Menyiapkan pratinjau…</p>'; const buf=await buildXlsx(window.TML_TEMPLATES[PK],PK,read()), res=await renderSheet(buf); window.PREV=res; mountPreview(); note(''); }
  catch(e){ $('prev').innerHTML=''; note('Pratinjau gagal: '+(e.message||e)); } }
function mountPreview(){ if(!window.PREV)return; const host=$('prev'); host.innerHTML=''; const pg=pageEl(window.PREV), sc=Math.min(1,(host.clientWidth-4)/pg.pw), box=document.createElement('div');
  box.className='pvbox'; box.style.cssText='width:'+pg.pw*sc+'px;height:'+pg.ph*sc+'px'; pg.el.style.transformOrigin='0 0'; pg.el.style.transform='scale('+sc+')'; box.append(pg.el); host.append(box); }
let rT; window.addEventListener('resize',()=>{clearTimeout(rT);rT=setTimeout(mountPreview,150);});
function printEl(el,css){ let h=$('printHost'); if(!h){h=document.createElement('div');h.id='printHost';document.body.append(h);} h.innerHTML='';h.append(el);
  let st=$('pgStyle'); if(!st){st=document.createElement('style');st.id='pgStyle';document.head.append(st);} st.textContent=css;
  const done=()=>{h.innerHTML='';window.removeEventListener('afterprint',done);}; window.addEventListener('afterprint',done); window.print(); }
async function printDoc(){ if(!window.PREV)await showPreview(); const pg=pageEl(window.PREV); printEl(pg.el,'@page{size:'+pg.wmm+'mm '+pg.hmm+'mm;margin:0}'); markDone('print'); }
async function pdfBlob(kind){ const buf=await buildXlsx(window.TML_TEMPLATES[kind],kind,read()), res=await renderSheet(buf), pg=pageEl(res), h=document.createElement('div');
  h.style.cssText='position:fixed;left:0;top:0;z-index:-9999;pointer-events:none'; h.append(pg.el); document.body.append(h);
  try{ const cv=await html2canvas(pg.el,{scale:2,backgroundColor:'#ffffff',logging:false}); const pdf=new jspdf.jsPDF({unit:'mm',format:[pg.wmm,pg.hmm],orientation:'portrait'}); pdf.addImage(cv.toDataURL('image/jpeg',0.92),'JPEG',0,0,pg.wmm,pg.hmm); return pdf.output('blob'); } finally{ h.remove(); } }
async function pdfDoc(){ try{ note('Membuat PDF…'); const b=await pdfBlob(PK); saveFile(fname(KN[PK],'pdf'),b,'application/pdf'); logA('PDF '+KN[PK]+' dibuat'); markDone('pdf'+PK); note('✓ PDF dibuat: '+fname(KN[PK],'pdf')); }catch(e){ note('Ekspor PDF gagal: '+(e.message||e)); } }
$('swCI').onclick=()=>showPreview('CI'); $('swPL').onclick=()=>showPreview('PL');
$('bXlsx').onclick=()=>busy('bXlsx',()=>dl(PK)); $('bBoth').onclick=()=>busy('bBoth',async()=>{await dl('CI');await new Promise(r=>setTimeout(r,600));await dl('PL');});
$('bPrint').onclick=()=>busy('bPrint',printDoc); $('bPdf').onclick=()=>busy('bPdf',pdfDoc); $('bSaveX').onclick=()=>busy('bSaveX',()=>save());
function reportHtml(){ return '<!doctype html><meta charset="utf-8"><title>Laporan Verifikasi</title><style>body{font:14px Arial;margin:24px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:6px 8px;text-align:left;vertical-align:top}th{background:#eee}</style><h2>Laporan Verifikasi – '+esc(S.name||'Proyek')+'</h2><p>PT Tri Mandiri Logistik · dibuat '+new Date().toLocaleString('id-ID')+'. Status dihitung dari data yang diinput, bukan pencocokan otomatis ke dokumen asli.</p><table>'+repHTML(false)+'</table>'; }
$('bRepDl').onclick=()=>saveFile(fname('LAPORAN_VERIFIKASI','html'),reportHtml(),'text/html');
$('bRepPr').onclick=()=>{ const d=document.createElement('div'); d.className='rptp'; d.innerHTML='<h2>Laporan Verifikasi – '+esc(S.name||'Proyek')+'</h2><p>PT Tri Mandiri Logistik · '+new Date().toLocaleString('id-ID')+'</p><table>'+repHTML(false)+'</table>'; printEl(d,'@page{size:A4;margin:10mm}'); };
$('bExp').onclick=()=>saveFile(fname('PROYEK','json'),JSON.stringify(clean(read()),null,1),'application/json');
$('fJson').onchange=async e=>{ const f=e.target.files[0]; e.target.value=''; if(!f)return; try{ const d=JSON.parse(await f.text()); if(!Array.isArray(d.units))throw new Error('Format bukan proyek TML'); S=hydrate({...d,key:rid(),confirmed:null}); render(); draft(); note('Proyek diimpor: '+(S.name||'(tanpa nama)')); }catch(x){ note('Gagal impor: '+x.message); } };
$('fXls').onchange=async e=>{ const f=e.target.files[0]; e.target.value=''; if(!f)return; try{ const r=await importBuf(await f.arrayBuffer()); note('Impor '+(r.kind==='CI'?'Commercial Invoice':'Packing List')+' selesai: '+r.added+' unit baru, '+r.upd+' diperbarui. Periksa data di tab Isi Data.'); }catch(x){ note('Gagal membaca Excel: '+x.message+' (hanya file CI/PL buatan aplikasi ini).'); } };
function openCfg(force){ $('cfg').hidden=force?false:!$('cfg').hidden; $('cfgUrl').value=cfg().url||''; $('cfgCode').value=cfg().code||''; $('opName').value=LS.get('tml_op',''); if(!$('cfg').hidden)window.scrollTo(0,0); }
$('bCfg').onclick=()=>openCfg();
function pill(){$('pillT').textContent=cfg().url?'Tersimpan online':'Disimpan di perangkat';}
$('bCfgSave').onclick=async()=>{ LS.set('tml_cfg',{url:$('cfgUrl').value.trim(),code:$('cfgCode').value.trim()}); LS.set('tml_op',$('opName').value.trim()); $('cfgMsg').textContent=cfg().url?'Tersimpan. Memuat daftar proyek bersama…':'Pengaturan tersimpan (penyimpanan bersama dimatikan).'; pill(); updConf(); await loadSaved(); };

/* ---- Jejak aktivitas ---- */
const opName=()=>LS.get('tml_op','')||'(tanpa nama)';
function logA(a){ if(!S)return; S.log=S.log||[]; S.log.push({t:Date.now(),u:opName(),a}); if(S.log.length>300)S.log=S.log.slice(-300); renderLog(); draft(); }
function renderLog(){ const L=((S&&S.log)||[]).slice(-15).reverse(); $('logBox').innerHTML=L.length?L.map(x=>'<div class="fr"><span>'+esc(x.a)+'</span><span class="sub nm">'+esc(x.u)+' · '+new Date(x.t).toLocaleString('id-ID')+'</span></div>').join(''):'<div class="sub">Belum ada aktivitas.</div>'; }
const logText=()=>((S.log||[]).map(x=>new Date(x.t).toLocaleString('id-ID')+' | '+x.u+' | '+x.a).join('\n'))||'(belum ada aktivitas)';
/* ---- Paket arsip: ZIP, folder komputer, Google Drive ---- */
const slug=s=>((s||'').toUpperCase().replace(/[^A-Z0-9]+/g,'_').replace(/^_|_$/g,''))||'PROYEK';
function summaryText(files){ const s=read(), R=analyze(s), c={}, tw=s.units.reduce((a,u)=>a+num(u.kg),0), tp=s.units.reduce((a,u)=>a+num(u.price),0); R.forEach(r=>{c[stc(r.status)]=(c[stc(r.status)]||0)+1;});
  return ['TML HIBAH – GRAND DOCUMENTATION · PT Tri Mandiri Logistik','Ringkasan arsip proyek','',
  'Proyek        : '+s.name,'No. invoice   : '+s.invNo+'  ('+s.invDate+')','Jumlah unit   : '+s.units.length,'Total berat    : '+fmt(tw)+' kg','Total nilai    : ₩'+fmt(tp).replace('.00',''),
  'Konfirmasi    : '+(S.confirmed?new Date(S.confirmed).toLocaleString('id-ID'):'BELUM dikonfirmasi'),'Operator      : '+opName(),
  'Laporan       : VERIFIED '+(c.VERIFIED||0)+' · NEED CONFIRMATION '+(c.NEED||0)+' · CONFLICT '+(c.CONFLICT||0)+' · MISSING '+(c.MISSING||0),
  'Lampiran      : '+(files.length?files.map(f=>f.type+' ('+f.name+')').join('; '):'tidak ada'),
  'Dibuat         : '+new Date().toLocaleString('id-ID'),'',
  'Catatan: status laporan dihitung dari data yang diinput operator, bukan pencocokan otomatis ke dokumen asli. Penentuan akhir isi dokumen tetap tanggung jawab pihak yang menyetujui.'].join('\n'); }
async function archiveEntries(){ const s=read(), yr=new Date(S.confirmed||Date.now()).getFullYear(), base='TML HIBAH/'+yr+'/'+slug(s.name)+'/', E=[];
  let files=[]; try{ files=await IDB.list(S.key)||[]; }catch(e){}
  files.forEach(f=>E.push({path:base+'01_Dokumen_Sumber/'+f.type.replace(/[^\w]+/g,'_')+'__'+f.name.replace(/[\\/:*?"<>|]+/g,'_'),data:f.blob,mime:f.mime}));
  for(const k of ['CI','PL']){ E.push({path:base+'02_Excel/'+fname(KN[k]),data:await buildXlsx(window.TML_TEMPLATES[k],k,s),mime:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}); E.push({path:base+'03_PDF/'+fname(KN[k],'pdf'),data:await pdfBlob(k),mime:'application/pdf'}); }
  E.push({path:base+'04_Laporan_Verifikasi/'+fname('LAPORAN_VERIFIKASI','html'),data:reportHtml(),mime:'text/html'});
  E.push({path:base+'data-proyek.json',data:JSON.stringify(clean(s),null,1),mime:'application/json'});
  E.push({path:base+'jejak-aktivitas.txt',data:logText(),mime:'text/plain'});
  E.push({path:base+'00_RINGKASAN.txt',data:summaryText(files),mime:'text/plain'});
  return {base,entries:E}; }
async function archiveZip(){ try{ note('Menyiapkan paket arsip…'); logA('Paket arsip (.zip) diunduh'); const {entries}=await archiveEntries(), z=new JSZip(); entries.forEach(e=>z.file(e.path,e.data)); const b=await z.generateAsync({type:'blob',compression:'DEFLATE'}); saveFile('TML_HIBAH_'+slug(S.name)+'_ARSIP.zip',b,'application/zip'); markDone('zip'); note('✓ Paket arsip dibuat: '+entries.length+' berkas. Struktur: TML HIBAH / tahun / nama proyek /'); }catch(e){ note('Paket arsip gagal: '+(e.message||e)); } }
const KV={open(){return new Promise((res,rej)=>{const r=indexedDB.open('tml_kv',1);r.onupgradeneeded=()=>r.result.createObjectStore('kv');r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);});},async get(k){const d=await this.open();return new Promise((res,rej)=>{const q=d.transaction('kv').objectStore('kv').get(k);q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error);});},async set(k,v){const d=await this.open();return new Promise((res,rej)=>{const t=d.transaction('kv','readwrite');t.objectStore('kv').put(v,k);t.oncomplete=()=>res();t.onerror=()=>rej(t.error);});}};
async function pickFolder(){ if(!window.showDirectoryPicker){ note('Penyimpanan ke folder hanya tersedia di Chrome/Edge pada PC atau laptop. Gunakan "Unduh Paket Arsip (.zip)".'); return null; }
  try{ const h=await window.showDirectoryPicker({mode:'readwrite',id:'tml-hibah'}); await KV.set('dir',h); $('dirName').textContent='Folder arsip: '+h.name; note('Folder arsip dipilih: '+h.name); return h; }catch(e){ if(e&&e.name!=='AbortError')note('Gagal memilih folder: '+e.message); return null; } }
async function saveFolder(){ try{ let h=await KV.get('dir').catch(()=>null); if(!h)h=await pickFolder(); if(!h)return;
    let p='granted'; if(h.queryPermission){ p=await h.queryPermission({mode:'readwrite'}); if(p!=='granted')p=await h.requestPermission({mode:'readwrite'}); }
    if(p!=='granted'){ note('Izin akses folder ditolak.'); return; }
    note('Menyimpan ke folder…'); logA('Arsip disimpan ke folder komputer'); const {entries}=await archiveEntries(), skip=/^TML HIBAH$/i.test(h.name)?'TML HIBAH/':'';
    for(const e of entries){ const parts=e.path.slice(skip.length).split('/'); let d=h; for(const x of parts.slice(0,-1))d=await d.getDirectoryHandle(x,{create:true}); const fh=await d.getFileHandle(parts[parts.length-1],{create:true}), w=await fh.createWritable(); await w.write(e.data); await w.close(); }
    markDone('folder'); note('✓ '+entries.length+' berkas disimpan di folder "'+h.name+'" → '+entries[0].path.split('/').slice(skip?1:0,-2).join('/')+'/'); }catch(e){ note('Simpan ke folder gagal: '+(e.message||e)); } }
const b64of=d=>new Promise((res,rej)=>{const fr=new FileReader();fr.onload=()=>res(String(fr.result).split(',')[1]||'');fr.onerror=()=>rej(fr.error);fr.readAsDataURL(d instanceof Blob?d:new Blob([d]));});
async function driveSend(){ if(!cfg().url){note('Isi URL Apps Script di ⚙ Pengaturan.');return;}
  try{ note('Menyiapkan arsip untuk Google Drive…'); logA('Arsip dikirim ke Google Drive (eksperimental)'); const {base,entries}=await archiveEntries(); let n=0;
    for(const e of entries){ note('Mengirim '+(++n)+'/'+entries.length+': '+e.path.split('/').pop()); await fetch(cfg().url,{method:'POST',mode:'no-cors',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'putfile',code:cfg().code||'',path:e.path,mime:e.mime||'application/octet-stream',b64:await b64of(e.data)})}); }
    note('Memverifikasi di Google Drive…'); let r=null; try{ r=await jsonp({action:'archive_check',path:base}); }catch(x){}
    if(r&&r.ok&&r.count>=entries.length){markDone('drive');} if(r&&r.ok&&r.count>=entries.length)note('✓ '+r.count+' berkas terverifikasi di Google Drive'+(r.url?' → '+r.url:'')+'.'); else note('Berkas terkirim, tetapi BELUM terverifikasi di Google Drive'+(r&&r.error?' ('+r.error+')':'')+'. Pastikan Code.gs sudah diperbarui, dideploy ulang, dan kode akses sama.'); }
  catch(e){ note('Kirim ke Google Drive gagal: '+(e.message||e)); } }
$('bZip').onclick=()=>busy('bZip',archiveZip); $('bDir').onclick=pickFolder; $('bFolder').onclick=()=>busy('bFolder',saveFolder); $('bDrive').onclick=()=>busy('bDrive',driveSend);
KV.get('dir').then(h=>{ if(h)$('dirName').textContent='Folder arsip: '+h.name; }).catch(()=>{});

/* ---- Kunci akses: template + stempel + data contoh terenkripsi (AES-256-GCM, kunci dari PBKDF2) ---- */
const b2u=b=>Uint8Array.from(atob(b),c=>c.charCodeAt(0));
async function openVault(pass){ const V=window.TML_VAULT; if(!V)throw new Error('Berkas vault.enc.js belum terunggah'); if(!(window.crypto&&crypto.subtle))throw new Error('Browser tidak mendukung enkripsi; buka lewat alamat https');
  const km=await crypto.subtle.importKey('raw',new TextEncoder().encode(pass),'PBKDF2',false,['deriveKey']);
  const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt:b2u(V.salt),iterations:V.iter,hash:'SHA-256'},km,{name:'AES-GCM',length:256},false,['decrypt']);
  return JSON.parse(new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:b2u(V.iv)},key,b2u(V.ct)))); }
function startApp(){
  ['brand','model','year','cyl','eng','color','kg','price','vin','from','to','vessel','invNo','invDate','depart','blTotal','blQty','name','engNo','dim','shipper','consignee'].forEach(k=>OPT[k]=new Set());
  learn(D); addOpt('name',D.name); for(const [k,v] of Object.entries(LS.get('tml_opts',{}))) v.forEach(x=>addOpt(k,x));
  $('bEx').hidden=!(D.units&&D.units.length);
  hdr(); const dr=LS.get('tml_draft',null); S=dr&&dr.units?hydrate(dr):fresh(); render(); pill(); loadSaved(); tab('isi'); }
const DDEF=()=>({name:'',invNo:'',invDate:'',depart:'',vessel:'',from:'',to:'',blTotal:'',blQty:'',blChassis:'',mou:'',shipper:[],consignee:[],units:[]});
async function tryUnlock(pass,remember){ const o=await openVault(pass); window.TML_TEMPLATES={CI:o.CI,PL:o.PL};
  D=Object.assign(DDEF(),o.D||{});
  try{ sessionStorage.setItem('tml_pass',pass); if(remember)localStorage.setItem('tml_pass',pass); }catch(e){}
  $('lock').hidden=true; startApp(); }
const forget=()=>{try{sessionStorage.removeItem('tml_pass');localStorage.removeItem('tml_pass');}catch(e){}};
$('bUnlock').onclick=async()=>{ const p=$('lockPass').value; if(!p)return; $('lockMsg').textContent='Membuka…'; try{ await tryUnlock(p,$('lockRem').checked); }catch(e){ $('lockMsg').textContent=(e&&e.name==='OperationError')?'Kode akses salah.':'Gagal: '+(e.message||e); } };
$('lockPass').addEventListener('keydown',e=>{if(e.key==='Enter')$('bUnlock').click();});
$('bLock').onclick=()=>{ forget(); location.reload(); };
(async()=>{ const need=[['JSZip','jszip.min.js'],['html2canvas','html2canvas.min.js'],['jspdf','jspdf.umd.min.js'],['renderSheet','preview.js']].filter(([g])=>typeof window[g]==='undefined');
  if(!window.TML_VAULT&&!window.TML_TEMPLATES)need.push(['','vault.enc.js atau templates-data.js']);
  const opt=[]; if(typeof window.pdfjsLib==='undefined')opt.push('baca/pdf.min.js'); if(typeof window.Tesseract==='undefined')opt.push('baca/tesseract.min.js'); if(typeof BACA==='undefined')opt.push('baca.js');
  if(need.length||opt.length){ $('warn').hidden=false; $('warn').innerHTML='<b>Berkas pendukung belum terunggah ke repository:</b> '+need.map(n=>'<code>'+n[1]+'</code>').concat(opt.map(n=>'<code>'+n+'</code>')).join(', ')+'. Unggah file tersebut ke repository <b>di level yang sama dengan index.html</b> (untuk folder <code>baca</code>, unggah foldernya juga), tunggu 1–3 menit, lalu muat ulang dengan Ctrl+F5.'+(need.length?'':' <i>Aplikasi tetap berjalan; hanya fitur Baca dokumen otomatis yang belum aktif.</i>'); }
  $('bLock').hidden=!window.TML_VAULT;
  if(!need.length&&!window.TML_VAULT){ D=Object.assign(DDEF(),window.TML_D||{}); startApp(); return; }   // mode terbuka
  let p=null; try{ p=sessionStorage.getItem('tml_pass')||localStorage.getItem('tml_pass'); }catch(e){}
  if(p&&!need.length){ try{ await tryUnlock(p,false); return; }catch(e){ forget(); } }
  if(window.TML_VAULT){ $('lock').hidden=false; $('lockPass').focus(); } })();
