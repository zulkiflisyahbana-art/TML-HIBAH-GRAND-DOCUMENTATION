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
async function buildXlsx(b64,kind,s){
  const N=s.units.length; if(N<1) throw new Error('Minimal 1 unit');
  const z=await JSZip.loadAsync(b64,{base64:true});
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
const D={name:'PARTY_2_OF_4',invNo:'SWC-OCTOBER-13 2026',invDate:'Oct 13-2026',depart:'',vessel:'',from:'Incheon,Republic of Korea',to:'Jakarta-Indonesia',blTotal:'21620',blQty:'2',blChassis:'KMFLA18TPFC092868\nKMFLA18TPFC092108',mou:'KMFLA18TPFC092868 10955\nKMFLA18TPFC092108 10665',
shipper:['SOCIAL WELFARE CORPORATION GO AND DO','80-22, Hanyuji-gil, Poseung-eup, Pyeongtaek-si, Gyeonggi-do,','Republic of Korea','TEL : (+82) 031-683-6491 ','FAX : (+82) 031-601-6417'],
consignee:['DIREKTORAT JENDERAL BINA ADMINISTRASI KEWILAYAHAN','DITJEN BINA ADMINISTRASI KEWILAYAHAN,KEMENTERIAN DALAM NEGERI','Medan Merdeka Utara 7,Medan Merdeka Utara ,Gambir,Gambir,','Kota ADM,Jakarta Pusat ,DKI Jakarta ,10110','Tel : +62 8122566520 +62 81313106941,+62 81284088281','Email : subditssarpras.infodamkar@gmail.com. Danang.ins@gmail.com ','vitus.nd15@gmail.com','ATTN: Danang Insita Putra, ST., M.Si(Han),Ph.d.'],
units:[{brand:'HYUNDAI',model:'FIRE TRUCK',vin:'KMFLA18TPFC092868',year:'2015',cyl:'5,899 CC',eng:'DIESEL',color:'RED&WHITE',kg:'10810',price:'210500000'},{brand:'HYUNDAI',model:'FIRE TRUCK',vin:'KMFLA18TPFC092108',year:'2015',cyl:'5,899 CC',eng:'DIESEL',color:'RED&WHITE',kg:'10810',price:'212272000'}]};
const $=id=>document.getElementById(id);
const HF=[['name','Nama proyek (untuk nama file)'],['invNo','No. invoice'],['invDate','Tanggal invoice'],['depart','Departure date (boleh kosong)'],['vessel','Vessel (boleh kosong)'],['from','From'],['to','To'],['blTotal','Berat total di B/L (kg)'],['blQty','Jumlah unit di B/L']];
const UF=[['brand','Brand'],['model','Model'],['vin','VIN / Chassis'],['year','Year'],['cyl','Cylinder'],['eng','Engine (jenis)'],['color','Color'],['kg','Berat (kg), net = gross'],['price','Harga satuan (₩)'],['engNo','Nomor mesin (opsional, hanya laporan)'],['dim','Dimensi P×L×T (opsional, hanya laporan)']];
const CP='<svg width="14" height="14" viewBox="0 0 14 14"><path d="M3 5l4 4 4-4" fill="none" stroke="currentColor" stroke-width="2"/></svg>';
const LS={get(k,d){try{const v=localStorage.getItem(k);return v==null?d:JSON.parse(v);}catch(e){return d;}},set(k,v){try{localStorage.setItem(k,JSON.stringify(v));}catch(e){}}};
let S, OPT={}, SAVED=[], UID=0;
const cfg=()=>LS.get('tml_cfg',{url:''});
const blank=(p={})=>({id:++UID,brand:'',model:'',vin:'',year:'',cyl:'',eng:'',color:'',kg:'',price:'',engNo:'',dim:'',edit:true,...p});
const addOpt=(k,v)=>{v=String(v??'').trim();if(v)(OPT[k]=OPT[k]||new Set()).add(v);};
function learn(s){HF.forEach(([k])=>addOpt(k,s[k]));(s.units||[]).forEach(u=>UF.forEach(([k])=>addOpt(k,u[k])));addOpt('shipper',(s.shipper||[]).join('\n'));addOpt('consignee',(s.consignee||[]).join('\n'));}
function persistOpts(){const o={};for(const k in OPT)o[k]=[...OPT[k]];LS.set('tml_opts',o);}
const fresh=()=>({name:'',invNo:'',invDate:'',depart:'',vessel:'',from:D.from,to:D.to,blTotal:'',blQty:'',blChassis:'',mou:'',shipper:[...D.shipper],consignee:[...D.consignee],units:[blank()]});
const example=()=>({...D,shipper:[...D.shipper],consignee:[...D.consignee],units:D.units.map(u=>blank({...u,edit:false}))});
const clean=s=>({...s,units:s.units.map(({id,edit,sure,ai,...u})=>u)});
const cb=(id,k,attrs,val)=>'<div class="cb"><input id="'+id+'" '+attrs+' value="'+String(val??'').replace(/&/g,'&amp;').replace(/"/g,'&quot;')+'" autocomplete="off"><button type="button" class="dd" data-k="'+k+'" data-t="'+id+'" aria-label="Pilih">'+CP+'</button></div>';
function hdr(){ $('hdr').innerHTML=HF.map(([k,l])=>'<div class="w"><label>'+l+'</label>'+cb('h_'+k,k,'data-h="'+k+'"','')+'</div>').join(''); }
function render(){ HF.forEach(([k])=>$('h_'+k).value=S[k]??''); $('shipper').value=S.shipper.join('\n'); $('consignee').value=S.consignee.join('\n'); $('blChassis').value=S.blChassis||''; $('mou').value=S.mou||''; renderUnits(); }
const sum=u=>[u.brand,u.model].filter(Boolean).join(' ')+(u.vin?' · '+u.vin:'')+(u.kg?' · '+fmt(num(u.kg))+' kg':'')+(num(u.price)?' · ₩'+fmt(num(u.price)).replace('.00',''):'');
function renderUnits(){
  $('units').innerHTML=S.units.map((u,i)=>{
    const head='<div class="uh"><b>Unit '+(i+1)+'</b>'+(u.edit?'<span class="s"></span>':'<span class="s">'+(sum(u)||'(belum diisi)')+'</span>')+'<span class="row">'
      +(u.edit?'<button class="sm p" data-a="done" data-id="'+u.id+'">Selesai</button>':'<button class="sm" data-a="edit" data-id="'+u.id+'">Edit</button>')
      +'<button class="sm d" data-a="del" data-id="'+u.id+'">'+(u.sure?'Yakin hapus?':'Hapus')+'</button></span></div>';
    const form=u.edit?'<div class="g" style="margin-top:12px">'+UF.map(([k,l])=>'<div class="w"><label>'+l+'</label>'+cb('u'+u.id+k,k,'class="'+((u.ai||{})[k]?'ai':'')+'" data-u="'+u.id+'" data-k="'+k+'"',u[k])+'</div>').join('')+'</div>':'';
    return '<div class="card">'+head+form+'</div>'; }).join('');
  $('uCount').textContent=S.units.length+' unit'; check(); }
function read(){ const s=clean(S); s.shipper=s.shipper.slice(0,5); s.consignee=s.consignee.slice(0,8); return s; }
/* ---- Pemeriksaan dan laporan verifikasi ---- */
const VINRE=/^[A-HJ-NPR-Z0-9]{17}$/i;
function analyze(s){
  const R=[], A=(info,data,status,note)=>R.push({info,data,status,note:note||''});
  const n=s.units.length, bl=num(s.blTotal), blq=num(s.blQty);
  const list=(s.blChassis||'').split('\n').map(x=>x.trim().toUpperCase()).filter(Boolean);
  const mou={}; (s.mou||'').split('\n').forEach(l=>{const m=l.match(/([A-HJ-NPR-Z0-9]{17})\D+([\d.,]+)/i); if(m)mou[m[1].toUpperCase()]=num(m[2]);});
  A('Jumlah unit',n+' unit',!n?'MISSING':blq?(n===blq?'VERIFIED':'CONFLICT'):'NEED CONFIRMATION',blq?'B/L: '+blq+' unit':'Isi "Jumlah unit di B/L" untuk mencocokkan');
  A('No. invoice',s.invNo||'-',s.invNo.trim()?'NEED CONFIRMATION':'MISSING','Cocokkan dengan keterangan invoice di B/L');
  A('Shipper',(s.shipper[0]||'-'),s.shipper.join('').trim()?'NEED CONFIRMATION':'MISSING','Cocokkan dengan B/L');
  A('Consignee',(s.consignee[0]||'-'),s.consignee.join('').trim()?'NEED CONFIRMATION':'MISSING','Harus sesuai B/L dan NPWP');
  A('From / To',(s.from||'-')+' → '+(s.to||'-'),(s.from.trim()&&s.to.trim())?'NEED CONFIRMATION':'MISSING','Port of loading di B/L dapat berbeda dari template');
  A('Departure date / Vessel',(s.depart||'-')+' / '+(s.vessel||'-'),(s.depart.trim()&&s.vessel.trim())?'NEED CONFIRMATION':'MISSING','Opsional; boleh dikosongkan sesuai keputusan TML');
  const tw=s.units.reduce((a,u)=>a+num(u.kg),0);
  A('Total berat',fmt(tw)+' kg',!bl?'NEED CONFIRMATION':(tw===bl?'VERIFIED':'CONFLICT'),bl?'B/L: '+fmt(bl)+' kg':'Isi berat total B/L untuk mencocokkan');
  const seen={};
  s.units.forEach((u,i)=>{ const L='Unit '+(i+1), v=u.vin.trim().toUpperCase(); seen[v]=(seen[v]||0)+1;
    let st,nt=''; if(!v){st='MISSING';} else if(!VINRE.test(v)){st='CONFLICT';nt='VIN harus 17 karakter alfanumerik tanpa I, O, Q ('+v.length+' karakter)';} else if(list.length){st=list.includes(v)?'VERIFIED':'CONFLICT';nt=list.includes(v)?'Ada di daftar chassis B/L':'Tidak ada di daftar chassis B/L';} else {st='NEED CONFIRMATION';nt='Format valid; cocokkan dengan B/L';}
    A(L+' · Chassis / VIN',v||'-',st,nt);
    A(L+' · Brand / Model / Year',[u.brand,u.model,u.year].filter(Boolean).join(' / ')||'-',(u.brand&&u.model&&u.year)?((/^\d{4}$/.test(String(u.year).trim())&&+u.year>=1980&&+u.year<=2027)?'NEED CONFIRMATION':'CONFLICT'):'MISSING',(u.brand&&u.model&&u.year)?'Cocokkan dengan MoU dan spesifikasi':'');
    A(L+' · Cylinder / Engine / Color',[u.cyl,u.eng,u.color].filter(Boolean).join(' / ')||'-',(u.cyl&&u.eng&&u.color)?'NEED CONFIRMATION':'MISSING','Belum ada di B/L dan MoU; ikuti template atau spesifikasi kendaraan');
    A(L+' · Nomor mesin',u.engNo||'-',u.engNo.trim()?'NEED CONFIRMATION':'MISSING','Opsional; tidak tercetak di template');
    A(L+' · Dimensi',u.dim||'-',u.dim.trim()?'NEED CONFIRMATION':'MISSING','Opsional; tidak tercetak di template');
    const mw=mou[v]; A(L+' · Berat',fmt(num(u.kg))+' kg',!num(u.kg)?'MISSING':(mw==null?'NEED CONFIRMATION':(mw===num(u.kg)?'VERIFIED':'CONFLICT')),mw==null?'Isi data MoU untuk mencocokkan':'MoU: '+fmt(mw)+' kg');
    A(L+' · Harga satuan','₩'+fmt(num(u.price)).replace('.00',''),num(u.price)>0?'NEED CONFIRMATION':'MISSING','Cocokkan dengan lampiran MoU');
  });
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
  $('checks').innerHTML=r.map(([ok,t])=>'<div class="chk '+(ok?'ok':'wn')+'">'+(ok?'✓ ':'⚠ ')+t+'</div>').join(''); }
function repHTML(){ const R=analyze(read());
  return '<thead><tr><th>No</th><th>Informasi</th><th>Data pada form</th><th>Status</th><th>Catatan</th></tr></thead><tbody>'+R.map((r,i)=>'<tr><td>'+(i+1)+'</td><td>'+esc(r.info)+'</td><td>'+esc(r.data)+'</td><td class="st s-'+r.status.split(' ')[0]+'">'+r.status+'</td><td>'+esc(r.note)+'</td></tr>').join('')+'</tbody>'; }
function fname(p){ return 'TML_'+p+'_'+((S.name||'').trim().toUpperCase().replace(/[^A-Z0-9]+/g,'_').replace(/^_|_$/g,'')||'PROYEK')+'.xlsx'; }
const msg=(t,c)=>{$('msg').textContent=t;$('msg').className=c||'';};
function saveFile(name,data,type){ const b=new Blob([data],{type:type||'application/octet-stream'}), u=URL.createObjectURL(b), a=document.createElement('a'); a.href=u; a.download=name; document.body.append(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(u),4000); }
async function dl(kind){
  if(!S.units.length){msg('Tambahkan minimal 1 unit.','wn');return;}
  try{ msg('Menyiapkan file…'); const s=read(), buf=await buildXlsx(window.TML_TEMPLATES[kind],kind,s);
    saveFile(fname(kind==='CI'?'COMMERCIAL_INVOICE':'PACKING_LIST'),buf,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); learn(s); persistOpts(); msg('Selesai. Periksa file di Excel.','ok');
  }catch(e){ msg('Gagal: '+(e.message||e),'wn'); } }
/* ---- Penyimpanan: browser (default) + Google Apps Script (opsional) ---- */
function jsonp(params){ const url=cfg().url; return new Promise((res,rej)=>{ const cbn='tml_cb_'+Date.now()+Math.floor(Math.random()*1e4), sc=document.createElement('script');
  const done=()=>{clearTimeout(t);delete window[cbn];sc.remove();}; const t=setTimeout(()=>{done();rej(new Error('Waktu habis'));},20000);
  window[cbn]=d=>{done();res(d);}; sc.onerror=()=>{done();rej(new Error('Gagal terhubung ke Apps Script'));};
  sc.src=url+(url.includes('?')?'&':'?')+new URLSearchParams({...params,callback:cbn}); document.head.append(sc); }); }
async function loadSaved(){ let L=Object.values(LS.get('tml_projects',{}));
  if(cfg().url){ try{ const r=await jsonp({action:'list'}); if(r&&r.ok) L=r.items.map(x=>({id:x.id,name:x.name,updated:+x.updated,data:x.data})); }catch(e){ $('plist').textContent='Penyimpanan bersama tidak terjangkau: '+e.message+'. Menampilkan proyek di perangkat ini.'; } }
  SAVED=L.sort((a,b)=>b.updated-a.updated);
  SAVED.forEach(p=>{try{learn(JSON.parse(p.data));addOpt('name',p.name);}catch(e){}});
  $('projSel').innerHTML='<option value="">— Proyek tersimpan ('+SAVED.length+') —</option>'+SAVED.map(p=>'<option value="'+esc(p.id)+'">'+esc(p.name)+'</option>').join(''); }
const pid=n=>(n.trim().toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'')||'proyek');
function load(p){ const d=JSON.parse(p.data); d.units=(d.units||[]).map(u=>blank({...u,edit:false})); S={...fresh(),...d}; render(); }
async function save(){ if(!S.name.trim()){$('plist').textContent='Isi nama proyek dulu.';return;}
  const d=clean(read()); learn(d); persistOpts(); const rec={id:pid(S.name),name:S.name,updated:Date.now(),data:JSON.stringify(d)};
  const all=LS.get('tml_projects',{}); all[rec.id]=rec; LS.set('tml_projects',all); let t='Tersimpan di perangkat ini: '+S.name;
  if(cfg().url){ try{ const r=await jsonp({action:'save',id:rec.id,name:rec.name,updated:rec.updated,data:rec.data}); t=r&&r.ok?'Tersimpan di perangkat ini dan di penyimpanan bersama: '+S.name:'Tersimpan di perangkat ini. Penyimpanan bersama menolak: '+((r&&r.error)||'tidak diketahui'); }catch(e){ t='Tersimpan di perangkat ini. Penyimpanan bersama gagal: '+e.message; } }
  $('plist').textContent=t; await loadSaved(); $('projSel').value=rec.id; $('bDelP').hidden=false; }
/* ---- Impor ---- */
async function importBuf(buf){ const sh=parseSheet(await readXlsx(buf)); let added=0,upd=0;
  ['invNo','invDate','depart','vessel','from','to'].forEach(k=>{if(sh[k]!==undefined&&sh[k]!=='')S[k]=sh[k];});
  if(sh.shipper.length)S.shipper=sh.shipper; if(sh.consignee.length)S.consignee=sh.consignee;
  sh.units.forEach(x=>{ let u=S.units.find(y=>y.vin.trim()&&y.vin.trim()===x.vin.trim());
    if(!u){ const e=S.units.length===1&&!S.units[0].vin.trim()&&!S.units[0].brand.trim(); if(e){u=S.units[0];} else {u=blank();S.units.push(u);} u.edit=false; added++; } else upd++;
    Object.keys(x).forEach(k=>{if(x[k]!==''&&x[k]!=='0')u[k]=x[k];}); });
  render(); return {kind:sh.kind,added,upd}; }
/* ---- Menu ▾ dan event ---- */
function closeMenus(){document.querySelectorAll('.menu').forEach(m=>m.remove());}
function openMenu(btn){ closeMenus(); const k=btn.dataset.k, tgt=$(btn.dataset.t), w=btn.closest('.w'); const L=[...(OPT[k]||[])].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  const m=document.createElement('div'); m.className='menu'; const cbx=btn.closest('.cb'); m.style.top=((cbx?cbx.offsetTop+cbx.offsetHeight:btn.offsetTop+btn.offsetHeight)+2)+'px';
  if(!L.length){const e=document.createElement('div');e.className='em';e.textContent='Belum ada pilihan, ketik saja.';m.append(e);}
  L.forEach(v=>{const d=document.createElement('div');d.textContent=v.replace(/\n/g,' ⏎ ');d.title=v;d.onclick=()=>{tgt.value=v;tgt.dispatchEvent(new Event('input',{bubbles:true}));closeMenus();};m.append(d);});
  w.append(m); }
document.addEventListener('click',e=>{
  const b=e.target.closest('button'); if(!b){ if(!e.target.closest('.menu'))closeMenus(); return; }
  if(b.dataset.t){ e.stopPropagation(); openMenu(b); return; }
  closeMenus(); const a=b.dataset.a; if(!a)return; const u=S.units.find(x=>x.id===+b.dataset.id); if(!u)return;
  if(a==='edit'){u.edit=true;renderUnits();} else if(a==='done'){u.edit=false;renderUnits();}
  else if(a==='del'){ if(u.sure){S.units=S.units.filter(x=>x!==u);renderUnits();} else {u.sure=true;renderUnits();setTimeout(()=>{if(S.units.includes(u)&&u.sure){u.sure=false;renderUnits();}},3000);} } });
document.addEventListener('input',e=>{ const t=e.target;
  if(t.dataset.h)S[t.dataset.h]=t.value; else if(t.dataset.u){const u=S.units.find(x=>x.id===+t.dataset.u);u[t.dataset.k]=t.value;if(u.ai)delete u.ai[t.dataset.k];}
  else if(t.dataset.ml)S[t.dataset.ml]=t.value.split('\n'); else if(t.dataset.s)S[t.dataset.s]=t.value;
  t.classList.remove('ai'); check(); });
$('bAdd').onclick=()=>{S.units.forEach(u=>u.edit=false);const p=S.units[S.units.length-1]||{};S.units.push(blank({brand:p.brand||'',model:p.model||'',year:p.year||'',cyl:p.cyl||'',eng:p.eng||'',color:p.color||'',kg:p.kg||''}));renderUnits();};
$('projSel').onchange=e=>{const p=SAVED.find(x=>x.id===e.target.value);$('bDelP').hidden=!p;if(p)load(p);};
$('bNew').onclick=()=>{S=fresh();$('projSel').value='';$('bDelP').hidden=true;render();};
$('bEx').onclick=()=>{S=example();render();};
let sureP=false;
$('bDelP').onclick=async()=>{ if(!sureP){sureP=true;$('bDelP').textContent='Yakin hapus proyek ini?';setTimeout(()=>{sureP=false;$('bDelP').textContent='Hapus proyek tersimpan';},3000);return;}
  sureP=false;$('bDelP').textContent='Hapus proyek tersimpan'; const id=$('projSel').value; if(!id)return;
  const all=LS.get('tml_projects',{}); delete all[id]; LS.set('tml_projects',all); let t='Proyek dihapus dari perangkat ini.';
  if(cfg().url){try{const r=await jsonp({action:'delete',id});t=r&&r.ok?'Proyek dihapus dari perangkat ini dan penyimpanan bersama.':'Dihapus di perangkat ini; penyimpanan bersama menolak.';}catch(e){t='Dihapus di perangkat ini; penyimpanan bersama gagal: '+e.message;}}
  $('plist').textContent=t; await loadSaved(); $('bDelP').hidden=true; };
$('bSave').onclick=save; $('bCI').onclick=()=>dl('CI'); $('bPL').onclick=()=>dl('PL');
$('bRep').onclick=()=>{$('rep').innerHTML=repHTML();$('repCard').hidden=false;$('repCard').scrollIntoView({behavior:'smooth'});};
$('bRepPr').onclick=()=>window.print();
$('bRepDl').onclick=()=>{ const html='<!doctype html><meta charset="utf-8"><title>Laporan Verifikasi</title><style>body{font:14px Arial;margin:24px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:6px 8px;text-align:left;vertical-align:top}th{background:#eee}</style><h2>Laporan Verifikasi – '+esc(S.name||'Proyek')+'</h2><p>PT Tri Mandiri Logistik · dibuat '+new Date().toLocaleString('id-ID')+'. Status dihitung dari data yang diinput, bukan pencocokan otomatis ke dokumen asli.</p><table>'+repHTML()+'</table>'; saveFile('TML_LAPORAN_VERIFIKASI_'+((S.name||'PROYEK').toUpperCase().replace(/[^A-Z0-9]+/g,'_'))+'.html',html,'text/html'); };
$('bExp').onclick=()=>saveFile('TML_PROYEK_'+((S.name||'PROYEK').toUpperCase().replace(/[^A-Z0-9]+/g,'_'))+'.json',JSON.stringify(clean(read()),null,1),'application/json');
$('fJson').onchange=async e=>{ const f=e.target.files[0]; e.target.value=''; if(!f)return; try{ const d=JSON.parse(await f.text()); if(!Array.isArray(d.units))throw new Error('Format bukan proyek TML'); d.units=d.units.map(u=>blank({...u,edit:false})); S={...fresh(),...d}; render(); $('plist').textContent='Proyek diimpor: '+(S.name||'(tanpa nama)'); }catch(x){ $('plist').textContent='Gagal impor: '+x.message; } };
$('fXls').onchange=async e=>{ const f=e.target.files[0]; e.target.value=''; if(!f)return; try{ const r=await importBuf(await f.arrayBuffer()); $('plist').textContent='Impor '+(r.kind==='CI'?'Commercial Invoice':'Packing List')+' selesai: '+r.added+' unit baru, '+r.upd+' unit diperbarui. Periksa data sebelum diunduh.'; }catch(x){ $('plist').textContent='Gagal membaca Excel: '+x.message+' (hanya file CI/PL buatan aplikasi ini).'; } };
$('bCfg').onclick=()=>{$('cfg').hidden=!$('cfg').hidden;$('cfgUrl').value=cfg().url||'';};
$('bCfgSave').onclick=async()=>{ LS.set('tml_cfg',{url:$('cfgUrl').value.trim()}); $('cfgMsg').textContent=cfg().url?'Tersimpan. Memuat daftar proyek bersama…':'Penyimpanan bersama dimatikan.'; pill(); await loadSaved(); };
/* ---- Tab dan status penyimpanan ---- */
function tab(n){document.querySelectorAll('.tab').forEach(t=>t.hidden=t.dataset.tab!==n);document.querySelectorAll('.tabs button').forEach(b=>b.classList.toggle('on',b.dataset.go===n));window.scrollTo(0,0);}
document.querySelector('.tabs').addEventListener('click',e=>{const b=e.target.closest('button');if(b)tab(b.dataset.go);});
function pill(){$('pillT').textContent=cfg().url?'Tersimpan online':'Disimpan di perangkat';}
/* ---- Init ---- */
['brand','model','year','cyl','eng','color','kg','price','vin','from','to','vessel','invNo','invDate','depart','blTotal','blQty','name','engNo','dim','shipper','consignee'].forEach(k=>OPT[k]=new Set());
learn(D); addOpt('name',D.name); for(const [k,v] of Object.entries(LS.get('tml_opts',{}))) v.forEach(x=>addOpt(k,x));
hdr(); S=fresh(); render(); pill(); loadSaved();
if(typeof module!=='undefined') module.exports={buildXlsx,readXlsx,parseSheet,analyze};
