function colLetter(n){let s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

function buildStyles(){
  return `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="7"><font><sz val="10"/><name val="Arial"/></font><font><sz val="14"/><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><sz val="9"/><color rgb="FFAAAAAA"/><name val="Arial"/></font><font><sz val="7"/><b/><color rgb="FF888888"/><name val="Arial"/></font><font><sz val="10"/><b/><color rgb="FF1A1A1A"/><name val="Arial"/></font><font><sz val="9"/><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><sz val="8"/><i/><color rgb="FF888888"/><name val="Arial"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1A1A1A"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF5F5F3"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF8F8F6"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1A1A1A"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/></border><border><left style="thin"><color rgb="FFCCCCCC"/></left><right style="thin"><color rgb="FFCCCCCC"/></right><top style="thin"><color rgb="FFCCCCCC"/></top><bottom style="thin"><color rgb="FFCCCCCC"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="left" vertical="center" indent="1"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"><alignment horizontal="left" vertical="center" indent="1"/></xf><xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0"><alignment horizontal="left" vertical="bottom" indent="1"/></xf><xf numFmtId="0" fontId="4" fillId="3" borderId="0" xfId="0"><alignment horizontal="left" vertical="center" indent="1"/></xf><xf numFmtId="0" fontId="5" fillId="6" borderId="0" xfId="0"><alignment horizontal="left" vertical="center" indent="1" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" indent="1"/></xf><xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" indent="1"/></xf><xf numFmtId="0" fontId="6" fillId="3" borderId="0" xfId="0"><alignment horizontal="left" vertical="center" indent="1"/></xf></cellXfs></styleSheet>`;
}

function buildSheet(rows,colWidths){
  let xml=`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>`;
  colWidths.forEach((w,i)=>{xml+=`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;});
  xml+='</cols><sheetData>';
  const allMerges=[...new Set(rows.flatMap(r=>r.merges||[]))];
  rows.forEach(r=>{
    if(!r.cells){xml+=`<row ht="${r.ht||15}" customHeight="1"/>`;return;}
    xml+=`<row ht="${r.ht||18}" customHeight="1">`;
    r.cells.forEach((cell,ci)=>{
      if(!cell)return;
      const ref=colLetter(ci+1)+(r.ri||1);
      const s=cell.s||0,v=cell.v;
      if(v===null||v===undefined||v===''){xml+=`<c r="${ref}" s="${s}"/>`;return;}
      if(typeof v==='number'){xml+=`<c r="${ref}" t="n" s="${s}"><v>${v}</v></c>`;}
      else{xml+=`<c r="${ref}" t="inlineStr" s="${s}"><is><t>${esc(String(v))}</t></is></c>`;}
    });
    xml+='</row>';
  });
  xml+='</sheetData>';
  if(allMerges.length){xml+=`<mergeCells count="${allMerges.length}">${allMerges.map(m=>`<mergeCell ref="${m}"/>`).join('')}</mergeCells>`;}
  xml+='</worksheet>';
  return xml;
}

function zipFiles(files){
  const enc=new TextEncoder();
  const entries=[],centralDir=[];
  let offset=0;
  function crc32(buf){let c=0xFFFFFFFF;const t=new Uint32Array(256);for(let i=0;i<256;i++){let x=i;for(let j=0;j<8;j++)x=x&1?0xEDB88320^(x>>>1):x>>>1;t[i]=x;}for(let i=0;i<buf.length;i++)c=t[(c^buf[i])&0xFF]^(c>>>8);return(c^0xFFFFFFFF)>>>0;}
  function u16(v){const b=new Uint8Array(2);new DataView(b.buffer).setUint16(0,v,true);return b;}
  function u32(v){const b=new Uint8Array(4);new DataView(b.buffer).setUint32(0,v,true);return b;}
  for(const[name,content]of Object.entries(files)){
    const nb=enc.encode(name),db=typeof content==='string'?enc.encode(content):content;
    const crc=crc32(db);
    const lh=new Uint8Array([0x50,0x4B,0x03,0x04,20,0,0,0,0,0,0,0,0,0,...u32(crc),...u32(db.length),...u32(db.length),...u16(nb.length),...u16(0),...nb]);
    centralDir.push({name:nb,crc,size:db.length,offset});
    entries.push(lh,db);
    offset+=lh.length+db.length;
  }
  const cdOffset=offset;
  const cde=centralDir.map(({name,crc,size,offset})=>new Uint8Array([0x50,0x4B,0x01,0x02,20,0,20,0,0,0,0,0,0,0,0,0,...u32(crc),...u32(size),...u32(size),...u16(name.length),...u16(0),...u16(0),0,0,0,0,0,0,0,0,...u32(offset),...name]));
  const cdSize=cde.reduce((s,e)=>s+e.length,0);
  const eocd=new Uint8Array([0x50,0x4B,0x05,0x06,0,0,0,0,...u16(centralDir.length),...u16(centralDir.length),...u32(cdSize),...u32(cdOffset),0,0]);
  const all=[...entries,...cde,eocd];
  const total=all.reduce((s,a)=>s+a.length,0);
  const out=new Uint8Array(total);let pos=0;
  all.forEach(a=>{out.set(a,pos);pos+=a.length;});
  return out;
}

function buildXlsx(klient,zakazka,adresa,datum,technik,items){
  const NCOLS=11;
  const S_DEF=0,S_TITLE=1,S_SUBTITLE=2,S_INFO_LBL=3,S_INFO_VAL=4,S_HDR=5,S_DATA=6,S_DATA_ALT=7,S_FOOTER=8;
  let ri=1;
  const rows=[],merges=[];
  function addMerge(c1,r1,c2,r2){merges.push(`${colLetter(c1)}${r1}:${colLetter(c2)}${r2}`);}
  function blankRow(ht,s){rows.push({ri:ri++,ht,cells:Array(NCOLS).fill({v:'',s:s||S_TITLE}),merges});}
  function titleRow(text,s,ht){const cells=Array(NCOLS).fill({v:'',s});cells[0]={v:text,s};addMerge(1,ri,NCOLS,ri);rows.push({ri:ri++,ht,cells,merges});}

  blankRow(8,S_TITLE);
  titleRow(`Výrobní dokumentace — Plisé standard / střešní  ·  Zakázka ${zakazka}`,S_TITLE,28);
  titleRow(`Klient: ${klient}    ·    Adresa: ${adresa}    ·    Technik: ${technik}    ·    Datum: ${datum}`,S_SUBTITLE,18);
  blankRow(8,S_INFO_LBL);blankRow(6,S_INFO_LBL);

  const lblC=Array(NCOLS).fill({v:'',s:S_INFO_LBL});
  lblC[0]={v:'KLIENT',s:S_INFO_LBL};lblC[3]={v:'ADRESA',s:S_INFO_LBL};lblC[6]={v:'ZAKÁZKA',s:S_INFO_LBL};lblC[9]={v:'TECHNIK',s:S_INFO_LBL};
  addMerge(1,ri,3,ri);addMerge(4,ri,6,ri);addMerge(7,ri,9,ri);addMerge(10,ri,NCOLS,ri);
  rows.push({ri:ri++,ht:16,cells:lblC,merges});

  const valC=Array(NCOLS).fill({v:'',s:S_INFO_VAL});
  valC[0]={v:klient,s:S_INFO_VAL};valC[3]={v:adresa,s:S_INFO_VAL};valC[6]={v:zakazka,s:S_INFO_VAL};valC[9]={v:technik,s:S_INFO_VAL};
  addMerge(1,ri,3,ri);addMerge(4,ri,6,ri);addMerge(7,ri,9,ri);addMerge(10,ri,NCOLS,ri);
  rows.push({ri:ri++,ht:20,cells:valC,merges});

  blankRow(10,S_INFO_LBL);blankRow(6,S_DEF);
  rows.push({ri:ri++,ht:22,cells:['Místnost','Šířka (mm)','Výška (mm)','m²','Barva profilu','Kód látky','Vodící lišta','Typ uchycení','Tyč (mm)','Ks','Poznámka'].map(h=>({v:h,s:S_HDR})),merges});

  items.forEach((o,i)=>{
    const s=i%2===0?S_DATA:S_DATA_ALT;
    const m2v=Math.round((o.s/1000)*(o.v/1000)*100)/100;
    rows.push({ri:ri++,ht:20,cells:[{v:o.mis,s},{v:o.s,s},{v:o.v,s},{v:m2v,s},{v:o.pro,s},{v:o.lat,s},{v:o.lis,s},{v:o.uch,s},{v:o.tyc||'',s},{v:o.p,s},{v:o.poz,s}],merges});
  });

  blankRow(6,S_FOOTER);
  const totalKs=items.reduce((s,o)=>s+o.p,0);
  const ftC=Array(NCOLS).fill({v:'',s:S_FOOTER});
  ftC[0]={v:`Celkem položek: ${items.length}   ·   Celkem ks: ${totalKs}`,s:S_FOOTER};
  ftC[7]={v:'Podpis technika: ___________________________',s:S_FOOTER};
  addMerge(1,ri,5,ri);addMerge(8,ri,NCOLS,ri);
  rows.push({ri:ri++,ht:18,cells:ftC,merges});

  const colWidths=[20,12,12,8,16,14,16,16,12,6,28];
  const sheetXml=buildSheet(rows,colWidths);
  const styleXml=buildStyles();
  const files={
    '[Content_Types].xml':`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels':`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml':`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Výrobní dokumentace" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels':`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml':sheetXml,
    'xl/styles.xml':styleXml
  };
  return zipFiles(files);
}

// --- Paměť posledních hodnot ---
const REMEMBER_FIELDS = ['f-lat','f-pro','f-lis','f-tyc','f-uch'];

function saveLastValues(){
  const last={};
  REMEMBER_FIELDS.forEach(id=>{last[id]=document.getElementById(id).value;});
  try{localStorage.setItem('plise-last',JSON.stringify(last));}catch(e){}
}

function loadLastValues(){
  try{
    const last=JSON.parse(localStorage.getItem('plise-last')||'{}');
    REMEMBER_FIELDS.forEach(id=>{
      if(last[id]!==undefined&&last[id]!==''){
        document.getElementById(id).value=last[id];
      }
    });
  }catch(e){}
}

// --- App ---
let items=[];
document.getElementById('datum').value=new Date().toISOString().split('T')[0];
loadLastValues();

function plural(n){return n===1?'1 položka':n>=2&&n<=4?n+' položky':n+' položek';}

function render(){
  const list=document.getElementById('list');
  document.getElementById('badge').textContent=plural(items.length);
  document.getElementById('btn-exp').disabled=items.length===0;
  if(!items.length){list.innerHTML='<div class="empty">Zatím žádná roleta.</div>';return;}
  list.innerHTML=items.map((o,i)=>`
    <div class="card">
      <div>
        <div class="card-name">${o.mis} <span style="font-weight:400;color:#888">— ${o.uch||'—'}</span></div>
        <div class="card-meta">
          ${o.s} × ${o.v} mm · ${o.p} ks · ${Math.round((o.s/1000)*(o.v/1000)*100)/100} m²<br>
          Látka: ${o.lat||'—'} · Profil: ${o.pro||'—'}${o.lis?' · Lišta: '+o.lis:''}${o.tyc?' · Tyč: '+o.tyc+' mm':''}${o.poz?'<br>'+o.poz:''}
        </div>
      </div>
      <button class="del" onclick="del(${i})">✕</button>
    </div>`).join('');
}

window.del=function(i){items.splice(i,1);render();};

document.getElementById('btn-add').addEventListener('click',function(){
  const mis=document.getElementById('f-mis').value.trim();
  const s=document.getElementById('f-s').value;
  const v=document.getElementById('f-v').value;
  if(!mis||!s||!v){alert('Vyplňte místnost a rozměry.');return;}
  items.push({
    mis,s:+s,v:+v,
    p:+(document.getElementById('f-p').value||1),
    lat:document.getElementById('f-lat').value.trim(),
    pro:document.getElementById('f-pro').value.trim(),
    lis:document.getElementById('f-lis').value.trim(),
    tyc:document.getElementById('f-tyc').value?+document.getElementById('f-tyc').value:'',
    uch:document.getElementById('f-uch').value,
    poz:document.getElementById('f-poz').value.trim()
  });
  saveLastValues();
  // vyčistit jen místnost, rozměry a poznámku — zbytek zůstane
  ['f-mis','f-s','f-v','f-poz'].forEach(function(id){document.getElementById(id).value='';});
  document.getElementById('f-p').value='1';
  render();
});

document.getElementById('btn-exp').addEventListener('click',function(){
  const klient=document.getElementById('klient').value||'—';
  const zakazka=document.getElementById('zakazka').value||'—';
  const adresa=document.getElementById('adresa').value||'—';
  const datum=document.getElementById('datum').value||'—';
  const technik=document.getElementById('technik').value||'—';
  const xlsx=buildXlsx(klient,zakazka,adresa,datum,technik,items);
  const blob=new Blob([xlsx],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  a.download=zakazka+'_'+klient.replace(/\s+/g,'-')+'.xlsx';
  document.body.appendChild(a);a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});
