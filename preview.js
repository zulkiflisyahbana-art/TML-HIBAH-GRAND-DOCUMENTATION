/* TML HIBAH – pratinjau Excel: menggambar ulang file .xlsx (merge, border, font, gambar) sebagai halaman HTML */
'use strict';
const PAPER={5:[215.9,355.6],9:[210,297],1:[215.9,279.4]};
function colNum(s){let n=0;for(const ch of s)n=n*26+ch.charCodeAt(0)-64;return n;}
function parseRef(r){const m=r.match(/^([A-Z]+)(\d+)$/);return [+m[2],colNum(m[1])];}
async function renderSheet(buf){
  const z=await JSZip.loadAsync(buf), rd=async p=>z.file(p)?z.file(p).async('string'):'';
  const P=s=>new DOMParser().parseFromString(s,'application/xml');
  const wb=P(await rd('xl/workbook.xml')), st=P(await rd('xl/styles.xml')), sh=P(await rd('xl/worksheets/sheet1.xml'));
  const kids=(el,tag)=>el?[...el.children].filter(c=>c.localName===tag):[];
  const g1=(el,t)=>el.getElementsByTagName(t)[0];
  const sstXml=await rd('xl/sharedStrings.xml');
  const sst=sstXml?[...P(sstXml).getElementsByTagName('si')].map(si=>[...si.getElementsByTagName('t')].filter(t=>t.parentNode.nodeName!=='rPh').map(t=>t.textContent).join('')):[];
  const color=c=>{if(!c)return null;const rgb=c.getAttribute('rgb');if(rgb)return '#'+rgb.slice(-6);const th=c.getAttribute('theme');if(th!=null)return th==='0'?'#ffffff':th==='1'?'#000000':null;const ix=c.getAttribute('indexed');return ix==='8'?'#000000':ix==='9'?'#ffffff':null;};
  const numFmts={};kids(g1(st,'numFmts'),'numFmt').forEach(n=>numFmts[n.getAttribute('numFmtId')]=n.getAttribute('formatCode'));
  const fonts=kids(g1(st,'fonts'),'font').map(f=>({n:(g1(f,'name')||{getAttribute:()=>null}).getAttribute('val'),sz:+((g1(f,'sz')||{getAttribute:()=>11}).getAttribute('val')),b:!!g1(f,'b'),i:!!g1(f,'i'),u:!!g1(f,'u'),c:color(g1(f,'color'))}));
  const fills=kids(g1(st,'fills'),'fill').map(f=>{const p=g1(f,'patternFill');return p&&p.getAttribute('patternType')==='solid'?color(g1(p,'fgColor')):null;});
  const bdef={thin:'1px solid',hair:'1px dotted',dotted:'1px dotted',dashed:'1px dashed',medium:'2px solid',thick:'3px solid',double:'3px double',mediumDashed:'2px dashed'};
  const borders=kids(g1(st,'borders'),'border').map(b=>{const o={};['left','right','top','bottom'].forEach(s=>{const e=kids(b,s)[0];o[s]=e&&e.getAttribute('style')?(bdef[e.getAttribute('style')]||'1px solid')+' '+(color(g1(e,'color'))||'#000'):null;});return o;});
  const xfs=kids(g1(st,'cellXfs'),'xf').map(x=>{const a=kids(x,'alignment')[0];return{n:x.getAttribute('numFmtId'),f:+x.getAttribute('fontId'),fl:+x.getAttribute('fillId'),b:+x.getAttribute('borderId'),h:a&&a.getAttribute('horizontal'),v:a&&a.getAttribute('vertical'),w:!!a&&a.getAttribute('wrapText')==='1',ind:a?+(a.getAttribute('indent')||0):0};});
  const fp=g1(sh,'sheetFormatPr'), dW=+((fp&&(fp.getAttribute('defaultColWidth')||fp.getAttribute('baseColWidth')))||8.43), dH=+((fp&&fp.getAttribute('defaultRowHeight'))||15);
  const colW={};kids(g1(sh,'cols'),'col').forEach(c=>{for(let k=+c.getAttribute('min');k<=Math.min(+c.getAttribute('max'),300);k++)colW[k]=c.getAttribute('hidden')==='1'?0:+c.getAttribute('width');});
  const rowH={};kids(g1(sh,'sheetData'),'row').forEach(r=>{rowH[+r.getAttribute('r')]=r.getAttribute('hidden')==='1'?0:(r.getAttribute('ht')?+r.getAttribute('ht'):dH);});
  const cw=c=>Math.round((colW[c]!=null?colW[c]:dW)*7+5), rh=r=>Math.round((rowH[r]!=null?rowH[r]:dH)*96/72);
  const cells={};
  kids(g1(sh,'sheetData'),'row').forEach(r=>kids(r,'c').forEach(c=>{const ref=c.getAttribute('r'),t=c.getAttribute('t'),v=g1(c,'v');let val=null;
    if(t==='s'&&v)val=sst[+v.textContent];else if(t==='inlineStr')val=[...c.getElementsByTagName('t')].map(x=>x.textContent).join('');else if(v)val=v.textContent;
    const [rr,cc]=parseRef(ref);cells[rr+','+cc]={s:+(c.getAttribute('s')||0),v:val,t};}));
  const merges={},covered={};
  kids(g1(sh,'mergeCells'),'mergeCell').forEach(m=>{const [a,b]=m.getAttribute('ref').split(':');const [r1,c1]=parseRef(a),[r2,c2]=b?parseRef(b):[r1,c1];merges[r1+','+c1]={r2,c2};for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)if(r!==r1||c!==c1)covered[r+','+c]=1;});
  let area=null;[...wb.getElementsByTagName('definedName')].forEach(d=>{if(d.getAttribute('name')==='_xlnm.Print_Area'){const m=d.textContent.match(/\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)/);if(m)area=[+m[2],colNum(m[1]),+m[4],colNum(m[3])];}});
  if(!area){const d=g1(sh,'dimension').getAttribute('ref').split(':');const a=parseRef(d[0]),b=parseRef(d[1]||d[0]);area=[a[0],a[1],b[0],b[1]];}
  const [R1,C1,R2,C2]=area, X={},Y={};let x=0;for(let c=C1;c<=C2+1;c++){X[c]=x;if(c<=C2)x+=cw(c);}let y=0;for(let r=R1;r<=R2+1;r++){Y[r]=y;if(r<=R2)y+=rh(r);}const W=x,H=y;
  const xAt=c=>{if(c>=C1&&c<=C2+1)return X[c];let s=0;if(c<C1){for(let k=c;k<C1;k++)s-=cw(k);}else{s=W;for(let k=C2+1;k<c;k++)s+=cw(k);}return s;};
  const yAt=r=>{if(r>=R1&&r<=R2+1)return Y[r];let s=0;if(r<R1){for(let k=r;k<R1;k++)s-=rh(k);}else{s=H;for(let k=R2+1;k<r;k++)s+=rh(k);}return s;};
  const fmtNum=(v,n)=>{if(v==null||v==='')return '';const f=numFmts[n]||({'3':'#,##0','4':'#,##0.00','2':'0.00','1':'0'})[n];const x=parseFloat(v);if(isNaN(x)||!f||!/[0#]/.test(f))return v;const d=(f.match(/\.(0+)/)||['',''])[1].length;return x.toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d,useGrouping:/,/.test(f)});};
  const root=document.createElement('div');root.style.cssText='position:relative;width:'+W+'px;height:'+H+'px;background:#fff;overflow:hidden;color:#000;';
  const cellAt=(r,c)=>cells[r+','+c];
  Object.keys(cells).sort().forEach(k=>{
    if(covered[k])return;const [r,c]=k.split(',').map(Number);if(r<R1||r>R2||c<C1||c>C2)return;
    const ce=cells[k],xf=xfs[ce.s]||xfs[0],m=merges[k],r2=Math.min(m?m.r2:r,R2),c2=Math.min(m?m.c2:c,C2);
    const bo=borders[xf.b]||{},bR=(m&&cellAt(r,c2)?borders[(xfs[cellAt(r,c2).s]||xf).b]||{}:bo).right,bB=(m&&cellAt(r2,c)?borders[(xfs[cellAt(r2,c).s]||xf).b]||{}:bo).bottom;
    const fill=fills[xf.fl],txt=ce.v==null||ce.v===''?'':(ce.t==='s'||ce.t==='inlineStr'||ce.t==='str'?ce.v:(ce.t==='b'||ce.t==='e'?ce.v:fmtNum(ce.v,xf.n)));
    if(!txt&&!fill&&!bo.left&&!bo.top&&!bR&&!bB)return;
    const fo=fonts[xf.f]||fonts[0]||{},isNum=ce.v!=null&&!ce.t&&txt!=='';
    const h=xf.h||(isNum?'right':'left'),jc=h==='center'||h==='centerContinuous'?'center':h==='right'?'flex-end':'flex-start',ai=xf.v==='top'?'flex-start':xf.v==='center'?'center':'flex-end';
    const d=document.createElement('div'),w=X[c2+1]-X[c]+1,hh=Y[r2+1]-Y[r]+1;
    d.style.cssText='position:absolute;box-sizing:border-box;left:'+X[c]+'px;top:'+Y[r]+'px;width:'+w+'px;height:'+hh+'px;display:flex;justify-content:'+jc+';align-items:'+ai+';padding:0 3px;padding-left:'+(3+xf.ind*9)+'px;overflow:'+(xf.w||m?'hidden':'visible')+';'+(fill?'background:'+fill+';':'')+(bo.left?'border-left:'+bo.left+';':'')+(bo.top?'border-top:'+bo.top+';':'')+(bR?'border-right:'+bR+';':'')+(bB?'border-bottom:'+bB+';':'');
    if(txt!==''){const s=document.createElement('span');s.textContent=txt;s.style.cssText='white-space:'+(xf.w?'pre-wrap':'pre')+';text-align:'+(h==='center'?'center':h==='right'?'right':'left')+';line-height:1.15;font-size:'+(fo.sz||11)+'pt;font-family:"'+(fo.n||'Arial')+'","Malgun Gothic","Noto Sans KR",Arial,sans-serif;'+(fo.b?'font-weight:700;':'')+(fo.i?'font-style:italic;':'')+(fo.u?'text-decoration:underline;':'')+'color:'+(fo.c||'#000')+';';d.append(s);}
    root.append(d);});
  const dr=await rd('xl/drawings/drawing1.xml');
  if(dr){const rels={};[...P(await rd('xl/drawings/_rels/drawing1.xml.rels')).getElementsByTagName('Relationship')].forEach(r=>rels[r.getAttribute('Id')]=r.getAttribute('Target'));
    for(const a of [...P(dr).getElementsByTagName('xdr:twoCellAnchor')]){
      const pic=a.getElementsByTagName('xdr:pic')[0];if(!pic)continue;const blip=pic.getElementsByTagName('a:blip')[0],id=blip.getAttribute('r:embed');
      const n=(el,t)=>+el.getElementsByTagName('xdr:'+t)[0].textContent,f=a.getElementsByTagName('xdr:from')[0],t=a.getElementsByTagName('xdr:to')[0];
      const x1=xAt(n(f,'col')+1)+n(f,'colOff')/9525,y1=yAt(n(f,'row')+1)+n(f,'rowOff')/9525,x2=xAt(n(t,'col')+1)+n(t,'colOff')/9525,y2=yAt(n(t,'row')+1)+n(t,'rowOff')/9525;
      const path='xl/'+String(rels[id]).replace(/^\.\.\//,''),file=z.file(path);if(!file)continue;
      const ext=path.split('.').pop().toLowerCase(),img=document.createElement('img');img.src='data:image/'+(ext==='jpg'?'jpeg':ext)+';base64,'+await file.async('base64');
      img.style.cssText='position:absolute;left:'+x1+'px;top:'+y1+'px;width:'+(x2-x1)+'px;height:'+(y2-y1)+'px;';root.append(img);}}
  const ps=g1(sh,'pageSetup'),pm=g1(sh,'pageMargins'),po=g1(sh,'printOptions'),sp=g1(sh,'pageSetUpPr');
  const paper=PAPER[ps&&ps.getAttribute('paperSize')]||PAPER[9];
  const mg=a=>pm&&pm.getAttribute(a)!=null?+pm.getAttribute(a):0.5;
  return {el:root,w:W,h:H,wmm:paper[0],hmm:paper[1],ml:mg('left'),mr:mg('right'),mt:mg('top'),mb:mg('bottom'),center:!!po&&po.getAttribute('horizontalCentered')==='1'};
}
function pageEl(res){
  const pw=Math.round(res.wmm/25.4*96),ph=Math.round(res.hmm/25.4*96),aw=pw-(res.ml+res.mr)*96,ah=ph-(res.mt+res.mb)*96,sc=Math.min(aw/res.w,ah/res.h,1);
  const pg=document.createElement('div');pg.className='page';pg.style.cssText='position:relative;background:#fff;overflow:hidden;width:'+pw+'px;height:'+ph+'px;';
  const box=res.el.cloneNode(true);const left=res.ml*96+(res.center?(aw-res.w*sc)/2:0);
  box.style.position='absolute';box.style.left=left+'px';box.style.top=(res.mt*96)+'px';box.style.transformOrigin='0 0';box.style.transform='scale('+sc+')';pg.append(box);
  return {el:pg,pw,ph,wmm:res.wmm,hmm:res.hmm};
}
