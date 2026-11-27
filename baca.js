/* TML HIBAH – pembaca dokumen sumber: teks asli PDF (pdf.js) dan OCR di browser (tesseract.js). Semua lokal, tanpa API berbayar. */
'use strict';
const BACA=(function(){
const norm=s=>String(s||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
const fixVin=t=>String(t||'').toUpperCase().replace(/[^A-Z0-9]/g,'').replace(/[IOQ]/g,m=>m==='I'?'1':'0');
const lev=(a,b)=>{const m=a.length,n=b.length,d=Array.from({length:m+1},(_,i)=>[i]);for(let j=1;j<=n;j++)d[0][j]=j;for(let i=1;i<=m;i++)for(let j=1;j<=n;j++)d[i][j]=Math.min(d[i-1][j]+1,d[i][j-1]+1,d[i-1][j-1]+(a[i-1]===b[j-1]?0:1));return d[m][n];};
const money=s=>{const t=String(s).replace(/[^0-9.,]/g,'');return t?Number(t.replace(/[.,]/g,'')):0;};
const MON={JAN:'Jan',FEB:'Feb',MAR:'Mar',APR:'Apr',MAY:'May',JUN:'Jun',JUL:'Jul',AUG:'Aug',SEP:'Sep',OCT:'Oct',NOV:'Nov',DEC:'Dec',JANUARY:'Jan',FEBRUARY:'Feb',MARCH:'Mar',APRIL:'Apr',JUNE:'Jun',JULY:'Jul',AUGUST:'Aug',SEPTEMBER:'Sep',OCTOBER:'Oct',NOVEMBER:'Nov',DECEMBER:'Dec'};
const title=s=>String(s||'').toLowerCase().replace(/\b([a-z])/g,m=>m.toUpperCase());
let worker=null;
async function getWorker(log){ if(worker)return worker; const Q=window.BACA_PATHS||{worker:'baca/worker.min.js',core:'baca/',lang:'baca/'}; worker=await Tesseract.createWorker('eng',1,{workerPath:Q.worker,corePath:Q.core,langPath:Q.lang,gzip:true,logger:log||(()=>{})}); return worker; }
async function ocr(canvas,psm,log){ const w=await getWorker(log); await w.setParameters({tessedit_pageseg_mode:String(psm),preserve_interword_spaces:'1'}); const r=await w.recognize(canvas);
  return {text:r.data.text,lines:r.data.lines.map(l=>({t:l.text.trim(),x:l.bbox.x0,y:l.bbox.y0,y1:l.bbox.y1}))}; }
async function renderCrop(pg,scale,rect){ const vp=pg.getViewport({scale}), x0=Math.round(rect[0]*vp.width), y0=Math.round(rect[1]*vp.height), w=Math.round((rect[2]-rect[0])*vp.width), h=Math.round((rect[3]-rect[1])*vp.height);
  const c=document.createElement('canvas'); c.width=w; c.height=h; const g=c.getContext('2d'); g.fillStyle='#fff'; g.fillRect(0,0,w,h);
  await pg.render({canvasContext:g,viewport:vp,transform:[1,0,0,1,-x0,-y0]}).promise; return c; }
function linesOf(items){ const rows=[]; items.slice().sort((a,b)=>b.y-a.y||a.x-b.x).forEach(it=>{ let r=rows.find(r=>Math.abs(r.y-it.y)<=2.5); if(!r){r={y:it.y,items:[]};rows.push(r);} r.items.push(it); });
  rows.sort((a,b)=>b.y-a.y); rows.forEach(r=>{r.items.sort((a,b)=>a.x-b.x); r.text=r.items.map(i=>i.s).join(' ').replace(/\s+/g,' ').trim();}); return rows; }
/* ---------- membaca berkas ---------- */
async function readFile(file,type,log){
  const pages=[], isPdf=/pdf/i.test(file.type)||/\.pdf$/i.test(file.name);
  if(!isPdf){ const bmp=await createImageBitmap(file), c=document.createElement('canvas'), sc=Math.min(3,4000/Math.max(bmp.width,bmp.height)); c.width=bmp.width*sc; c.height=bmp.height*sc; c.getContext('2d').drawImage(bmp,0,0,c.width,c.height);
    log('Membaca gambar (OCR)…'); const r=await ocr(c,6,null), r2=await ocr(c,11,null); pages.push({n:1,method:'ocr',text:r.text,text2:r2.text,lines:r.lines}); return pages; }
  const doc=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;
  for(let i=1;i<=doc.numPages;i++){ const pg=await doc.getPage(i), tc=await pg.getTextContent(), vp=pg.getViewport({scale:1});
    const items=tc.items.filter(x=>x.str.trim()).map(x=>({x:x.transform[4],y:x.transform[5],s:x.str})), chars=items.reduce((a,x)=>a+x.s.length,0), P={n:i,W:vp.width,H:vp.height,items};
    if(chars>=40){ P.method='teks'; P.text=linesOf(items).map(l=>l.text).join('\n'); }
    else { P.method='ocr'; log('OCR halaman '+i+' dari '+doc.numPages+'…'); const r=await ocr(await renderCrop(pg,3,[0,0,1,1]),6,null); P.text=r.text; P.lines=r.lines;
      if(type==='NPWP'||/NPWP/i.test(r.text)){ log('OCR tambahan (kartu) hal. '+i+'…'); P.text2=(await ocr(await renderCrop(pg,6,[0,0,1,1]),11,null)).text; }
      if(/NOMOR\s*CHASSIS|DAFTAR\s*BARANG/i.test(r.text)){ log('OCR tabel hal. '+i+'…'); const hl=(r.lines.find(l=>/DAFTAR\s*BARANG/i.test(l.t))||{y:0}), top=Math.max(0,hl.y/(pg.getViewport({scale:3}).height)-0.01), rect=[0.03,top,0.95,Math.min(1,top+0.17)];
        P.tbl=[]; const tc2=await renderCrop(pg,8,rect); P.tbl.push((await ocr(tc2,4,null)).text); P.tbl.push((await ocr(tc2,6,null)).text); } }
    pages.push(P); }
  return pages; }
/* ---------- B/L (teks asli) ---------- */
function parseBL(pages){ const P=pages.find(p=>p.items&&/BILL OF LADING/i.test(p.text||''))||pages[0]; if(!P||!P.items)return null; const it=P.items, L=linesOf(it), T=L.map(l=>l.text).join('\n'), src='B/L hlm.'+P.n+' (teks asli PDF)';
  const yOf=(re,xm=140)=>{const i=it.find(i=>re.test(i.s)&&i.x<xm);return i?i.y:null;};
  const ys=yOf(/^Shipper$/), yc=yOf(/^Consignee$/), yn=yOf(/^Notify Party$/), yv=yOf(/^Ocean Vessel/), yd=yOf(/^Port of discharge/);
  const reg=(x0,x1,yT,yB)=>linesOf(it.filter(i=>i.x>=x0&&i.x<x1&&i.y<yT&&i.y>yB)).map(l=>l.text);
  const o={src,pages:[P.n]};
  o.no=((T.match(/\b([A-Z]{3,6}\d{9,14})\b/)||[])[1])||''; o.nox=o.no;
  if(yv!=null){ const r=linesOf(it.filter(i=>i.y<yv-1&&i.y>yv-14)); if(r[0]){ o.vessel=r[0].items.filter(i=>i.x<190).map(i=>i.s).join(' ').trim(); o.pol=r[0].items.filter(i=>i.x>=190&&i.x<420).map(i=>i.s).join(' ').trim(); } }
  if(yd!=null){ const r=linesOf(it.filter(i=>i.y<yd-1&&i.y>yd-14)); if(r[0]) o.pod=r[0].items.filter(i=>i.x<190).map(i=>i.s).join(' ').trim(); }
  const g=T.match(/([\d.,]+)\s*KGS/i); o.grossKg=g?Math.round(Number(g[1].replace(/,/g,''))):0;
  const q=T.match(/(\d+)\s+UNIT\b/i); o.qty=q?+q[1]:0;
  o.chassis=[...T.matchAll(/CHASSIS\s*NO\.?\s*:?\s*([A-Z0-9]{17})/gi)].map(m=>m[1].toUpperCase());
  const d=T.match(/([A-Z]+)\s+([A-Z ]+?)\s*\/\s*(\d{4})\s*YEAR/i); if(d){o.brand=d[1];o.model=d[2].trim();o.year=d[3];}
  const iv=T.match(/INVOICE\s*NO\.?\s*([A-Z0-9][A-Z0-9\- ]*?\s\d{4})/i); o.invNo=iv?iv[1].trim():'';
  const dm=o.invNo.match(/([A-Za-z]+)[- ](\d{1,2})[ -](\d{4})/); o.invDate=dm&&MON[dm[1].toUpperCase()]?MON[dm[1].toUpperCase()]+' '+dm[2]+'-'+dm[3]:'';
  const np=T.match(/NPWP\s*:?\s*([\d.\-]{12,})/i); o.npwp=np?np[1].replace(/\D/g,''):'';
  const hs=T.match(/H\.?S\.?\s*CODE\s*:?\s*([\d.]+)/i); o.hs=hs?hs[1]:'';
  const ld=T.match(/Laden on board[\s\S]*?\b([A-Z]{3})\.(\d{1,2})\.(\d{4})/i)||[...T.matchAll(/\b([A-Z]{3})\.(\d{1,2})\.(\d{4})\b/g)].pop(); o.depart=ld?(MON[ld[1].toUpperCase()]||ld[1])+' '+ld[2]+'-'+ld[3]:'';
  const cont=[...T.matchAll(/\b(\d{2}[A-Z]{2})\s+([A-Z]{4}\d{7})\b/g)].map(m=>m[2]); o.containers=cont;
  o.shipperRaw=reg(0,340,ys,yc+1).map(s=>s.replace(/\*+/g,'').trim()).filter(Boolean);
  o.consigneeRaw=reg(0,340,yc,yn+1).map(s=>s.replace(/\*+/g,'').trim()).filter(Boolean);
  const blk=reg(225,600,(it.find(i=>/INVOICE\s*NO/i.test(i.s))||{y:500}).y+3,(it.find(i=>/^SAY\s*:/i.test(i.s))||{y:280}).y);
  const B=blk.join('\n'); o.block=B;
  o.shipFax=((B.match(/\*\*\s*FAX\s*:\s*([^\n]+)/i)||[])[1]||'').trim(); o.consAddr2=((B.match(/\*\*\*\s*([^\n]+)/)||[])[1]||'').trim();
  o.tel=((B.match(/(?:^|\n)TEL\s*:\s*([^\n]+)/i)||[])[1]||'').trim(); o.emails=[...B.matchAll(/[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}/gi)].map(m=>m[0].toLowerCase());
  o.attn=((B.match(/ATTN\s*:\s*([^\n]+)/i)||[])[1]||'').trim();
  return o; }
/* ---------- MoU (OCR) ---------- */
function parseMoU(pages,knownVins){ const P=pages.filter(p=>p.method==='ocr'||p.text); const all=P.map(p=>p.text||'').join('\n'); const o={src:'MoU',pages:P.map(p=>p.n),units:[],notes:[]};
  o.no=((all.match(/NOMOR\s*:?\s*([0-9.]+\/\d+\/[A-Z]+)/i)||[])[1])||''; o.date=((all.match(/TANGGAL\s*:?\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})/i)||[])[1])||'';
  const tu=all.match(/menghibahkan\s*(\d+)\s*\(/i); o.totalUnits=tu?+tu[1]:0; const tv=all.match(/nilai\s*sebesar\s*[^\d]{0,4}([\d.,]{9,})/i); o.totalValue=tv?money(tv[1]):0;
  const known=(knownVins||[]).map(fixVin); const rawToks=[];
  for(const m of all.matchAll(/Chassis\s*\n?\s*([A-Z0-9]{14,19})([\s\S]{0,140}?)(?=Chassis|\n\s*\n[a-d]\.|Pasal|$)/gi)) rawToks.push({raw:m[1],vin:fixVin(m[1]),rest:m[2]});
  const exact=new Set(rawToks.map(t=>t.vin).filter(v=>known.includes(v)));
  const snap=v=>{ if(!v||known.includes(v))return null; for(const k of known){ if(!exact.has(k)&&lev(v,k)===1)return {vin:k,d:1}; } return null; };
  const vinTok=txt=>[...String(txt).matchAll(/[A-Z0-9]{3}[A-Z0-9]{3,}[A-Z0-9]{3}/gi)].map(m=>m[0]).filter(t=>t.length>=15&&t.length<=19&&/\d{4}/.test(t)&&/^[A-Za-z]/.test(t));
  const seen={};
  for(const t of rawToks){ const sn=snap(t.vin), vin=sn?sn.vin:t.vin; if(seen[vin])continue; seen[vin]=1;
    const rest=t.rest, km=rest.match(/Tempuh\s*([\d.,]+)\s*KM/i), yr=rest.match(/produksi\s*\n?\s*(\S{3,5})/i), mk=rest.match(/Merk\s*([A-Za-z]+)/i);
    let year=yr?yr[1].replace(/[sS]/g,'5').replace(/[oO]/g,'0').replace(/[lI]/g,'1').replace(/\D/g,'').slice(0,4):''; if(!(+year>=1980&&+year<=2035))year='';
    o.units.push({vin,vinRaw:t.raw,snapped:!!sn,km:km?money(km[1]):0,year,brand:mk?mk[1]:'',pages:[2]}); }
  const cf=t=>String(t).toUpperCase().replace(/[OD]/g,'0').replace(/B/g,'8').replace(/S/g,'5').replace(/[IL]/g,'1').replace(/Z/g,'2').replace(/G/g,'6'); const tail=t=>cf(t).replace(/[^A-Z0-9]/g,'').slice(-6);
  const snapT=v=>{ const f=fixVin(v); if(o.units.some(u=>u.vin===f))return {vin:f}; const tt=tail(v); let best=null,sc=0,tie=false; o.units.forEach(u=>{ const ut=u.vin.slice(-6); let n=0; for(let i=0;i<6;i++)if(tt[i]===ut[i])n++; if(n>sc){sc=n;best=u.vin;tie=false;} else if(n===sc)tie=true; }); return (sc>=5&&!tie)?{vin:best}:null; };
  // tabel lampiran: dimensi, berat, harga per chassis (angka hanya diisi bila lolos pengecekan aritmetika)
  const tb=[]; pages.forEach(p=>{ if(p.rows)p.rows.forEach(t=>tb.push(t)); if(p.tbl)p.tbl.forEach(t=>tb.push(t)); if(p.text&&/NOMOR\s*CHASSIS/i.test(p.text))tb.push(p.text); });
  const cand={}; o.units.forEach(u=>cand[u.vin]={price:new Set(),kg:new Set(),dim:new Set()});
  tb.forEach(txt=>String(txt).split('\n').forEach(line=>{ const toks=vinTok(line); toks.forEach(tk=>{ const sn=snapT(tk); if(!sn||!cand[sn.vin])return; const rest=line.slice(line.toUpperCase().indexOf(tk.toUpperCase())+tk.length);
      [...rest.matchAll(/(\d{3})[.,](\d{3})[.,](\d{3})/g)].forEach(m=>cand[sn.vin].price.add(+(m[1]+m[2]+m[3])));
      [...rest.matchAll(/(\d{1,2})[.,]?(\d{3})\s*Kg/gi)].forEach(m=>cand[sn.vin].kg.add(+(m[1]+m[2])));
      const dm=[...rest.matchAll(/(\d)[.,]?(\d{3})\s*mm/gi)].map(m=>+(m[1]+m[2])); if(dm.length>=3)cand[sn.vin].dim.add(dm.slice(0,3).join('x')); }); }));
  o.cand=cand; o.tableTried=tb.length>0; o.priceValidated=false;
  const vins=o.units.map(u=>u.vin);
  if(o.totalValue&&vins.length&&vins.every(v=>cand[v].price.size)){ // cari kombinasi harga yang berjumlah sama dengan total MoU
    const opts=vins.map(v=>[...cand[v].price]); let found=null; (function rec(i,acc){ if(found)return; if(i===opts.length){ if(acc.reduce((a,b)=>a+b,0)===o.totalValue)found=acc.slice(); return;} for(const p of opts[i]){acc.push(p);rec(i+1,acc);acc.pop();} })(0,[]);
    if(found){ o.units.forEach((u,i)=>{u.price=found[i];u.priceOk=true;}); o.priceValidated=true; } }
  o.units.forEach(u=>{ if(u.price==null&&cand[u.vin].price.size===1){ u.priceCand=[...cand[u.vin].price][0]; } const dm=[...cand[u.vin].dim]; if(dm.length===1)u.dim=dm[0]; });
  return o; }
/* ---------- NPWP (OCR) ---------- */
function parseNPWP(pages){ const t=pages.map(p=>(p.text2||'')+'\n'+(p.text||'')).join('\n'); const o={src:'NPWP',pages:pages.map(p=>p.n)};
  const m=t.match(/(\d{2})[.,](\d{3})[.,](\d{3})[.,](\d)\s*[-–]\s*(\d{3})[.,](\d{3})/); o.npwp15=m?m.slice(1).join(''):'';
  const m16=t.match(/NPWP\s*16\s*:?\s*([\d ]{16,22})/i); o.npwp16=m16?m16[1].replace(/\D/g,''):''; if(!o.npwp15&&o.npwp16)o.npwp15=o.npwp16.slice(1);
  const nm=t.match(/(DIREKTORAT[\s\S]{0,160}?NEGERI)/i); o.name=nm?nm[1].replace(/\s+/g,' ').trim():''; const dt=t.match(/Terdaftar\s*(\d{2}\/\d{2}\/\d{4})/i); o.registered=dt?dt[1]:'';
  const kpp=t.match(/(KPP\s+PRATAMA[^\n]+)/i); o.kpp=kpp?kpp[1].trim():''; return o; }
/* ---------- gabungkan ---------- */
async function readAll(docs,log){ const out={bl:null,mou:null,npwp:null,errors:[]}; const rd={};
  for(const d of docs){ try{ log('Membaca '+d.name+'…'); rd[d.id]=await readFile(d.file,d.type,log); }catch(e){ out.errors.push(d.name+': '+(e.message||e)); } }
  const of=t=>docs.filter(d=>d.type===t&&rd[d.id]);
  for(const d of of('B/L')){ try{ const r=parseBL(rd[d.id]); if(r&&(r.no||r.chassis.length)){ r.file=d.name; out.bl=r; break; } }catch(e){ out.errors.push('B/L: '+e.message); } }
  const known=out.bl?out.bl.chassis:[];
  for(const d of of('MoU / Perjanjian Hibah')){ try{ const r=parseMoU(rd[d.id],known); r.file=d.name; out.mou=r; break; }catch(e){ out.errors.push('MoU: '+e.message); } }
  for(const d of of('NPWP')){ try{ const r=parseNPWP(rd[d.id]); r.file=d.name; out.npwp=r; break; }catch(e){ out.errors.push('NPWP: '+e.message); } }
  out.vocab=new Set(Object.values(rd).flat().map(p=>[p.text,p.text2].join(' ')).join(' ').toUpperCase().split(/[^A-Z0-9]+/).filter(w=>w.length>=4));
  if(worker){ try{ await worker.terminate(); }catch(e){} worker=null; }
  return out; }
return {readAll,readFile,parseBL,parseMoU,parseNPWP,norm,fixVin,title,lev,money};
})();
