// ===================== XLSX ENGINE =====================
function colLetter(n){let s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

function buildStyles(){
  // Styles:
  // 0=default, 1=title(white bold 16 center dark bg), 2=subtitle(gray 12 center dark bg)
  // 3=header(white bold 11 center dark bg), 4=data(black bold 10 center white bg)
  // 5=data-alt(black bold 10 center gray bg), 6=footer-left, 7=footer-right
  return `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="8">
<font><sz val="10"/><name val="Arial"/></font>
<font><sz val="16"/><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
<font><sz val="12"/><color rgb="FFAAAAAA"/><name val="Arial"/></font>
<font><sz val="11"/><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
<font><sz val="10"/><b/><color rgb="FF1A1A1A"/><name val="Arial"/></font>
<font><sz val="8"/><i/><color rgb="FF888888"/><name val="Arial"/></font>
<font><sz val="8"/><i/><color rgb="FF888888"/><name val="Arial"/></font>
<font><sz val="10"/><b/><color rgb="FF1A1A1A"/><name val="Arial"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1A1A1A"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF5F5F3"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/></border>
<border>
<left style="thin"><color rgb="FFCCCCCC"/></left>
<right style="thin"><color rgb="FFCCCCCC"/></right>
<top style="thin"><color rgb="FFCCCCCC"/></top>
<bottom style="thin"><color rgb="FFCCCCCC"/></bottom>
</border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="4" borderId="0" xfId="0"><alignment horizontal="left" vertical="center" indent="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="4" borderId="0" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
</cellXfs>
</styleSheet>`;
}

function buildSheetXml(title, subtitle, headers, colWidths, dataRows, totalM2){
  const NCOLS = headers.length;
  const S_TITLE=1, S_SUB=2, S_HDR=3, S_DATA=4, S_DATA_ALT=5, S_FOOT_L=6, S_FOOT_R=7;
  let ri=1; const rows=[], merges=[];
  function addM(c1,r1,c2,r2){merges.push(`${colLetter(c1)}${r1}:${colLetter(c2)}${r2}`);}

  // Row 1: Title
  addM(1,ri,NCOLS,ri);
  const titleCells=Array(NCOLS).fill({v:'',s:S_TITLE});
  titleCells[0]={v:title,s:S_TITLE};
  rows.push({ri:ri++,ht:36,cells:titleCells,merges});

  // Row 2: Subtitle
  addM(1,ri,NCOLS,ri);
  const subCells=Array(NCOLS).fill({v:'',s:S_SUB});
  subCells[0]={v:subtitle,s:S_SUB};
  rows.push({ri:ri++,ht:26,cells:subCells,merges});

  // Row 3: Spacer
  rows.push({ri:ri++,ht:8,cells:Array(NCOLS).fill({v:'',s:S_TITLE}),merges});

  // Row 4: Headers
  rows.push({ri:ri++,ht:34,cells:headers.map(h=>({v:h,s:S_HDR})),merges});

  // Data rows
  dataRows.forEach((dr,i)=>{
    const s=i%2===0?S_DATA:S_DATA_ALT;
    rows.push({ri:ri++,ht:33,cells:dr.map(v=>({v:v==null?'':v,s})),merges});
  });

  // Spacer before footer
  rows.push({ri:ri++,ht:6,cells:Array(NCOLS).fill({v:'',s:S_FOOT_L}),merges});

  // Footer row
  const half=Math.max(1,Math.floor(NCOLS/2));
  const ftCells=Array(NCOLS).fill({v:'',s:S_FOOT_L});
  const m2Text=totalM2!=null?`   ·   Celkem m²: ${totalM2}`:'';
  ftCells[0]={v:`Celkem položek: ${dataRows.length}${m2Text}`,s:S_FOOT_L};
  ftCells[half]={v:'Podpis technika: ___________________________',s:S_FOOT_R};
  if(half>1)addM(1,ri,half-1,ri);
  addM(half,ri,NCOLS,ri);
  rows.push({ri:ri++,ht:18,cells:ftCells,merges});

  // Build XML
  let xml=`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>`;
  colWidths.forEach((w,i)=>{xml+=`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;});
  xml+='</cols><sheetData>';
  const allMerges=[...new Set(merges)];
  rows.forEach(r=>{
    xml+=`<row ht="${r.ht||18}" customHeight="1">`;
    r.cells.forEach((cell,ci)=>{
      const ref=colLetter(ci+1)+(r.ri||1),s=cell.s||0,v=cell.v;
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

function buildMultiSheetXlsx(sheets){
  const styleXml=buildStyles();
  const sheetFiles={},sheetRels=[],sheetList=[],overrides=[];
  sheets.forEach((s,i)=>{
    const name=`sheet${i+1}`,rid=`rId${i+1}`;
    sheetFiles[`xl/worksheets/${name}.xml`]=s.xml;
    sheetRels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${name}.xml"/>`);
    const sname=s.name.replace(/[:\\\/\?\*\[\]]/g,'').substring(0,31);
    sheetList.push(`<sheet name="${esc(sname)}" sheetId="${i+1}" r:id="${rid}"/>`);
    overrides.push(`<Override PartName="/xl/worksheets/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
  });
  const files={
    '[Content_Types].xml':`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides.join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels':`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml':`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetList.join('')}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels':`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRels.join('')}<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml':styleXml,
    ...sheetFiles
  };
  return zipFiles(files);
}

function zipFiles(files){
  const enc=new TextEncoder(),entries=[],centralDir=[];
  let offset=0;
  function crc32(buf){let c=0xFFFFFFFF;const t=new Uint32Array(256);for(let i=0;i<256;i++){let x=i;for(let j=0;j<8;j++)x=x&1?0xEDB88320^(x>>>1):x>>>1;t[i]=x;}for(let i=0;i<buf.length;i++)c=t[(c^buf[i])&0xFF]^(c>>>8);return(c^0xFFFFFFFF)>>>0;}
  function u16(v){const b=new Uint8Array(2);new DataView(b.buffer).setUint16(0,v,true);return b;}
  function u32(v){const b=new Uint8Array(4);new DataView(b.buffer).setUint32(0,v,true);return b;}
  for(const[name,content]of Object.entries(files)){
    const nb=enc.encode(name),db=typeof content==='string'?enc.encode(content):content;
    const crc=crc32(db);
    const lh=new Uint8Array([0x50,0x4B,0x03,0x04,20,0,0,0,0,0,0,0,0,0,...u32(crc),...u32(db.length),...u32(db.length),...u16(nb.length),...u16(0),...nb]);
    centralDir.push({name:nb,crc,size:db.length,offset});
    entries.push(lh,db);offset+=lh.length+db.length;
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

function downloadBlob(data,filename){
  const blob=new Blob([data],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download=filename;
  document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
}

function formatDatum(d){
  if(!d)return '—';
  const parts=d.split('-');
  if(parts.length===3)return `${parts[2]}-${parts[1]}-${parts[0]}`;
  return d;
}

// ===================== HELPERS =====================
function m2(s,v){return Math.round((+s/1000)*(+v/1000)*100)/100;}

// Products with m2 field injected automatically where applicable
// m2Field: index where m2 should appear in toRow output (after height)
// hasM2: true if product uses width/height and should auto-calc m2

const PRODUCTS={
  stresni:{label:'Látkové střešní rolety',hasM2:true,
    headers:['Místnost','Šířka (mm)','Výška (mm)','m²','Barva profilu','Typ látky','Kód látky','Počet ks','Poznámka'],
    colWidths:[20,12,12,8,14,14,14,8,28],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.bprofilu,o.typ_latky,o.kod_latky,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  aluvegasdn:{label:'Den a noc — Alu Vegas',hasM2:true,
    headers:['Místnost','Šířka kazety C (mm)','Šířka látky A (mm)','Celk. výška D (mm)','Výška látky J (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Vodící lišta','Počet ks','Poznámka'],
    colWidths:[18,16,16,16,16,8,10,14,12,12,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirkac',label:'Šířka kazety C (mm)',type:'number',required:true},{id:'sirkaa',label:'Šířka látky A (mm)',type:'number',required:true,hint:'Rozdíl C−A musí být 35–50 mm',validate:(v,all)=>{const d=+all.sirkac-(+v);if(d<35||d>50)return `Rozdíl C−A = ${d} mm. Musí být 35–50 mm!`;return null;}},{id:'vyskad',label:'Celková výška D (mm)',type:'number',required:true},{id:'vyskaj',label:'Výška látky J (mm)',type:'number'},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'vod_lista',label:'Vodící lišta',type:'select',options:['plochá','radius']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirkac,+o.sirkaa,+o.vyskad,+o.vyskaj||'',m2(o.sirkac,o.vyskad),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.vod_lista,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirkac,o.vyskad)},
  aluvegastx:{label:'Textilní roleta — Alu Classic',hasM2:true,
    headers:['Místnost','Šířka kazety C (mm)','Šířka látky A (mm)','Celk. výška D (mm)','m²','Počet ks','Ovládání','Barva profilu','Typ látky','Kód látky','Vodící lišta','Poznámka'],
    colWidths:[18,16,16,16,8,8,10,14,12,12,14,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirkac',label:'Šířka kazety C (mm)',type:'number',required:true},{id:'sirkaa',label:'Šířka látky A (mm)',type:'number',required:true,hint:'Rozdíl C−A musí být 35–50 mm',validate:(v,all)=>{const d=+all.sirkac-(+v);if(d<35||d>50)return `Rozdíl C−A = ${d} mm. Musí být 35–50 mm!`;return null;}},{id:'vyskad',label:'Celková výška D (mm)',type:'number',required:true},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'vod_lista',label:'Vodící lišta',type:'select',options:['plochá','radius']},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirkac,+o.sirkaa,+o.vyskad,m2(o.sirkac,o.vyskad),+o.pocet||1,o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.vod_lista,o.poznamka],
    getM2:o=>m2(o.sirkac,o.vyskad)},
  maxidn:{label:'Den a noc — Maxi kazeta',hasM2:true,
    headers:['Místnost','Šířka (mm)','Výška (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Typ uchycení','Boční vod. lišta','Počet ks','Poznámka'],
    colWidths:[18,12,12,8,10,14,12,12,28,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'uchyceni',label:'Typ uchycení',type:'select',options:['Do zdi','Do stropu místnosti','Do stropu výklenku','Na okno']},{id:'boc_lista',label:'Boční vodící lišta',type:'select',options:['ANO','NE']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.uchyceni,o.boc_lista,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  maxitx:{label:'Textilní — Maxi kazeta',hasM2:true,
    headers:['Místnost','Šířka (mm)','Výška (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Typ uchycení','Boční vod. lišta','Počet ks','Poznámka'],
    colWidths:[18,12,12,8,10,14,12,12,28,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'select',options:['Bílá','Hnědá']},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'uchyceni',label:'Typ uchycení',type:'select',options:['Do zdi','Do stropu místnosti','Do stropu výklenku','Na okno']},{id:'boc_lista',label:'Boční vodící lišta',type:'select',options:['ANO','NE']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.uchyceni,o.boc_lista,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  otevrenadn:{label:'Den a noc — Otevřená kazeta',hasM2:true,
    headers:['Místnost','Šířka (mm)','Výška (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Typ uchycení','Vedení','Počet ks','Poznámka'],
    colWidths:[18,12,12,8,10,14,12,12,30,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'uchyceni',label:'Typ uchycení',type:'select',options:['Zavěšení — bez vrtání','Vrtání do zdi','Vrtání do stropu','Vrtání na okno']},{id:'vedeni',label:'Vedení',type:'select',options:['Silon','U lišta']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.uchyceni,o.vedeni,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  otevrenatx:{label:'Textilní — Otevřená kazeta',hasM2:true,
    headers:['Místnost','Šířka (mm)','Výška (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Typ uchycení','Vedení','Počet ks','Poznámka'],
    colWidths:[18,12,12,8,10,14,12,12,30,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'select',options:['Bílá','Hnědá']},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'uchyceni',label:'Typ uchycení',type:'select',options:['Zavěšení — bez vrtání','Vrtání do zdi','Vrtání do stropu','Vrtání na okno']},{id:'vedeni',label:'Vedení',type:'select',options:['NE','Silon','U lišta']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.uchyceni,o.vedeni,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  polodn:{label:'Den a noc — Polo kazeta',hasM2:true,
    headers:['Místnost','Šířka (mm)','Výška (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Typ uchycení','Vedení','Počet ks','Poznámka'],
    colWidths:[18,12,12,8,10,14,12,12,30,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'uchyceni',label:'Typ uchycení',type:'select',options:['Vrtání do zdi','Vrtání do stropu','Vrtání na okno']},{id:'vedeni',label:'Vedení',type:'select',options:['NE','U lišta']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.uchyceni,o.vedeni,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  polotx:{label:'Textilní — Polo kazeta',hasM2:true,
    headers:['Místnost','Šířka (mm)','Výška (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Typ uchycení','Vedení','Počet ks','Poznámka'],
    colWidths:[18,12,12,8,10,14,12,12,30,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'select',options:['Bílá','Hnědá']},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'uchyceni',label:'Typ uchycení',type:'select',options:['Vrtání do zdi','Vrtání do stropu místnosti','Vrtání do stropu výklenku','Vrtání na okno']},{id:'vedeni',label:'Vedení',type:'select',options:['NE','Silon','U lišta']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.uchyceni,o.vedeni,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  sonosdn:{label:'Den a noc — SONO S',hasM2:true,
    headers:['Místnost','Šířka kazety C (mm)','Celk. výška I (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Boční vod. lišta','Počet ks','Poznámka'],
    colWidths:[18,16,16,8,10,14,12,12,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirkac',label:'Šířka kazety C (mm)',type:'number',required:true},{id:'vyskai',label:'Celková výška I (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'boc_lista',label:'Boční vodící lišta',type:'select',options:['ANO','NE']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirkac,+o.vyskai,m2(o.sirkac,o.vyskai),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.boc_lista,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirkac,o.vyskai)},
  sonostx:{label:'Textilní — SONO S',hasM2:true,
    headers:['Místnost','Šířka kazety C (mm)','Celk. výška I (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Boční vod. lišta','Spodní vod. lišta','Počet ks','Poznámka'],
    colWidths:[18,16,16,8,10,14,12,12,14,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirkac',label:'Šířka kazety C (mm)',type:'number',required:true},{id:'vyskai',label:'Celková výška I (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'boc_lista',label:'Boční vodící lišta',type:'select',options:['ANO','NE']},{id:'spod_lista',label:'Spodní vodící lišta',type:'select',options:['ANO','NE']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirkac,+o.vyskai,m2(o.sirkac,o.vyskai),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.boc_lista,o.spod_lista,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirkac,o.vyskai)},
  sonolldn:{label:'Den a noc — SONO L',hasM2:true,
    headers:['Místnost','Šířka kazety C (mm)','Celk. výška I (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Boční vod. lišta','Počet ks','Poznámka'],
    colWidths:[18,16,16,8,10,14,12,12,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirkac',label:'Šířka kazety C (mm)',type:'number',required:true},{id:'vyskai',label:'Celková výška I (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'boc_lista',label:'Boční vodící lišta',type:'select',options:['ANO','NE']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirkac,+o.vyskai,m2(o.sirkac,o.vyskai),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.boc_lista,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirkac,o.vyskai)},
  sonolltx:{label:'Textilní — SONO L',hasM2:true,
    headers:['Místnost','Šířka kazety C (mm)','Celk. výška I (mm)','m²','Ovládání','Barva profilu','Typ látky','Kód látky','Boční vod. lišta','Spodní vod. lišta','Počet ks','Poznámka'],
    colWidths:[18,16,16,8,10,14,12,12,14,14,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirkac',label:'Šířka kazety C (mm)',type:'number',required:true},{id:'vyskai',label:'Celková výška I (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'typ_latky',label:'Typ látky',type:'text'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'boc_lista',label:'Boční vodící lišta',type:'select',options:['ANO','NE']},{id:'spod_lista',label:'Spodní vodící lišta',type:'select',options:['ANO','NE']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirkac,+o.vyskai,m2(o.sirkac,o.vyskai),o.ovladani,o.bprofilu,o.typ_latky,o.kod_latky,o.boc_lista,o.spod_lista,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirkac,o.vyskai)},
  horizont:{label:'Horizontální žaluzie',hasM2:true,
    headers:['Pozice','Šířka (mm)','Výška (mm)','m²','Ovládání','Délka řetízku (mm)','Barva profilu','Typ lamely','Barva lamely','Domyk. provedení','Délka ovládání (mm)','Distanční podložky','Bar. sladění žebřík+texband','Bezpečnostní prvek','Počet ks','Poznámka'],
    colWidths:[14,12,12,8,10,16,14,12,14,14,14,16,20,16,8,24],
    fields:[{id:'mis',label:'Pozice / místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['P — pravé','L — levé']},{id:'delka_retizku',label:'Délka ovládacího řetízku (mm)',type:'number'},{id:'bprofilu',label:'Barva profilu',type:'text',hint:'např. RAL 9010'},{id:'typ_lamely',label:'Typ lamely',type:'text',hint:'Isoline / Loco / Prim / Eco'},{id:'barva_lamely',label:'Barva lamely',type:'text'},{id:'domyk',label:'Domykací provedení',type:'select',options:['ANO','NE']},{id:'delka_ovl',label:'Délka ovládání — jiná (mm)',type:'number'},{id:'dist_podlozky',label:'Distanční podložky (ks)',type:'number'},{id:'bar_sladeni',label:'Bar. sladění žebřík+texband',type:'select',options:['ANO','NE']},{id:'bezpec',label:'Bezpečnostní prvek',type:'select',options:['ANO','NE']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.ovladani,+o.delka_retizku||'',o.bprofilu,o.typ_lamely,o.barva_lamely,o.domyk,+o.delka_ovl||'',+o.dist_podlozky||'',o.bar_sladeni,o.bezpec,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  vertikal:{label:'Vertikální žaluzie',hasM2:true,
    headers:['Pozice','Provedení typ','Šířka látky','Počet ks','Šířka (mm)','Výška (mm)','m²','Typ stahování','Počet barev','Barva','Uchycení','Uchycení navíc (ks)','Délka ovládání (mm)','Bezpečnostní prvek','Omezení typ','Poznámka'],
    colWidths:[12,14,12,8,12,12,8,16,12,12,12,14,16,16,14,24],
    fields:[{id:'mis',label:'Pozice / místnost',type:'text',required:true},{id:'provedeni',label:'Provedení typ',type:'select',options:['1 — standard','2 — lux']},{id:'sirka_latky',label:'Šířka látky',type:'select',options:['89','127']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'typ_stah',label:'Typ stahování',type:'select',options:['1 — k ovladači','2 — od ovladače','3 — od středu','4 — do středu','5 — oboustranné','8/1','8/2','8/3','8/4']},{id:'pocet_barev',label:'Počet barev',type:'number',defaultVal:'1'},{id:'barva',label:'Barva (kód)',type:'text'},{id:'uchyceni',label:'Uchycení',type:'select',options:['Strop','Stěna']},{id:'uchyceni_navic',label:'Uchycení navíc (ks)',type:'number'},{id:'delka_ovl',label:'Délka ovládání (mm)',type:'number'},{id:'bezpec',label:'Bezpečnostní prvek',type:'select',options:['ANO','NE']},{id:'omezeni',label:'Omezení typ',type:'select',options:['bez omezení','pouze profil','pouze látka']},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,o.provedeni,o.sirka_latky,+o.pocet||1,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.typ_stah,+o.pocet_barev||1,o.barva,o.uchyceni,+o.uchyceni_navic||'',+o.delka_ovl||'',o.bezpec,o.omezeni,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  plise:{label:'Plisé roleta',hasM2:true,
    headers:['Místnost','Šířka (mm)','Výška (mm)','m²','Počet ks','Kód látky','Barva profilu','Vodící lišta (střešní)','Typ uchycení','Ovládací tyč (mm)','Poznámka'],
    colWidths:[18,12,12,8,8,14,14,18,22,14,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'kod_latky',label:'Kód látky',type:'text'},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'vod_lista',label:'Barva vodící lišty (střešní)',type:'text'},{id:'uchyceni',label:'Typ uchycení',type:'select',options:['Zasklívačka','Rám okna','Zavěšení na rám okna']},{id:'tyc',label:'Ovládací tyč — délka (mm)',type:'number'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),+o.pocet||1,o.kod_latky,o.bprofilu,o.vod_lista,o.uchyceni,+o.tyc||'',o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  sit_dverpant:{label:'Dveřní síť — pantová',hasM2:true,
    headers:['Místnost','Šířka sítě (mm)','Výška sítě (mm)','m²','Strana pantů','Barva profilu','Barva sítě','Typ pantů','Magnet','Horiz. zpevňující profil','Okopová lišta','Průchod pro zvířata','Počet ks','Poznámka'],
    colWidths:[16,14,14,8,12,12,20,16,22,24,12,22,8,24],
    fields:[{id:'mis',label:'Místnost',type:'text',required:true},{id:'sirka',label:'Šířka sítě (mm)',type:'number',required:true},{id:'vyska',label:'Výška sítě (mm)',type:'number',required:true},{id:'strana',label:'Strana pantů (z pohledu z venku)',type:'select',options:['L','P']},{id:'bprofilu',label:'Barva profilu',type:'text'},{id:'bsite',label:'Barva sítě',type:'select',options:['černá','šedá','anti-alergení černá','anti-alergení šedá','proti hlodavcům černá']},{id:'typ_pantu',label:'Typ pantů',type:'select',options:['Klasické','Samozavírací']},{id:'magnet',label:'Magnet',type:'select',options:['Standard','Magnetická páska hnědá','Magnetická páska bílá']},{id:'horiz_profil',label:'Horiz. zpevňující profil',type:'select',options:['Standard (1/3 výšky)','Vlastní']},{id:'horiz_vlastni',label:'Vlastní výška profilu (cm)',type:'number',hint:'Vyplnit jen při Vlastní'},{id:'okopova',label:'Okopová lišta',type:'select',options:['10 cm','30 cm']},{id:'pruchod',label:'Průchod pro zvířata',type:'select',options:['NE','15×15 cm','23,3×27,5 cm','29,7×34,5 cm']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.strana,o.bprofilu,o.bsite,o.typ_pantu,o.magnet,o.horiz_profil+(o.horiz_vlastni?' ('+o.horiz_vlastni+' cm)':''),o.okopova,o.pruchod,+o.pocet||1,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  sit_plise:{label:'Plisé síť proti hmyzu',hasM2:true,
    headers:['Pozice','Počet ks','Šířka (mm)','Výška (mm)','m²','Barva profilu','Typ sítě','Práh','Barva síťoviny','Krycí lišta','Vynášecí profil','Montážní L-profil','Poznámka'],
    colWidths:[14,8,12,12,8,16,16,16,14,12,14,16,24],
    fields:[{id:'mis',label:'Pozice / místnost',type:'text',required:true},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'bprofilu',label:'Barva profilu',type:'select',options:['1 — bílá mat','2 — hnědá mat','3 — antracit mat','4 — DB 703','5 — antracit str.','6 — nástřik zlatý dub']},{id:'typ_site',label:'Typ sítě',type:'select',options:['a — Stellar','b — Stellar opona','c — Stellar Lux','d — Stellar Lux opona','e — Stellar Mini']},{id:'prah',label:'Práh',type:'select',options:['1a — standard','1b — zešikmený','2a — standard','2b — s praporkem']},{id:'bsitoviny',label:'Barva síťoviny',type:'select',options:['černá','šedá (jen Stellar Mini)']},{id:'kryci_lista',label:'Krycí lišta',type:'select',options:['ANO','NE']},{id:'vynaseci',label:'Vynášecí profil',type:'select',options:['ANO','NE']},{id:'mont_lprofil',label:'Montážní L-profil',type:'select',options:['40×20 — standard','60×20']},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,+o.pocet||1,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.bprofilu,o.typ_site,o.prah,o.bsitoviny,o.kryci_lista,o.vynaseci,o.mont_lprofil,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  sit_posuvram:{label:'Posuvná síť — v rámu',hasM2:true,
    headers:['Pozice','Typ profilu','Počet ks','Šířka (mm)','Výška (mm)','m²','Šířka křídla','Poloha příčky','Barva profilu','Barva síťoviny','Montáž rám','Montáž ostění','Poznámka'],
    colWidths:[14,14,8,12,12,8,16,16,14,12,12,12,24],
    fields:[{id:'mis',label:'Pozice / místnost',type:'text',required:true},{id:'typ_profil',label:'Typ profilu',type:'select',options:['PSR1','PSR1 ECO','PSR2','PSR2 ECO','PSR1 T','PSR2 T']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'sirka_kridla',label:'Šířka křídla',type:'select',options:['v 1/2 (standard)','v 1/3','vlastní']},{id:'sirka_kridla_mm',label:'Vlastní šířka křídla (mm)',type:'number',hint:'Jen při vlastní'},{id:'poloha_pricka',label:'Poloha příčky',type:'select',options:['v 1/3 (standard)','v 1/2','vlastní']},{id:'pricka_mm',label:'Vlastní poloha příčky (mm)',type:'number',hint:'Jen při vlastní'},{id:'bprofilu',label:'Barva profilu',type:'select',options:['bílá','hnědá','zlatý dub — renolit','RAL — vlastní']},{id:'bsitoviny',label:'Barva síťoviny',type:'select',options:['Š — šedá','Č — černá']},{id:'montaz_ram',label:'Montáž — rám',type:'select',options:['ANO','NE']},{id:'montaz_osteni',label:'Montáž — ostění',type:'select',options:['ANO','NE']},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,o.typ_profil,+o.pocet||1,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.sirka_kridla+(o.sirka_kridla_mm?' ('+o.sirka_kridla_mm+' mm)':''),o.poloha_pricka+(o.pricka_mm?' ('+o.pricka_mm+' mm)':''),o.bprofilu,o.bsitoviny,o.montaz_ram,o.montaz_osteni,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  sit_posuvlist:{label:'Posuvná síť — v lištách',hasM2:false,
    headers:['Pozice','Typ profilu','Počet ks','Šířka křídla (mm)','Výška vč. vod. lišt (mm)','Délka vod. lišt (mm)','Poloha příčky','Barva profilu','Barva síťoviny','Poznámka'],
    colWidths:[14,14,8,16,20,18,18,14,12,24],
    fields:[{id:'mis',label:'Pozice / místnost',type:'text',required:true},{id:'typ_profil',label:'Typ profilu',type:'select',options:['PS1','PS1 ECO s příčkou','PS1 Z','PS1 ECO Z s příčkou']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'sirka_kridla',label:'Šířka křídla (mm)',type:'number',required:true},{id:'vyska_vod',label:'Výška vč. vod. lišt (mm)',type:'number',required:true},{id:'delka_list',label:'Délka vod. lišt (mm)',type:'number'},{id:'poloha_pricka',label:'Poloha příčky',type:'select',options:['v 1/3 (standard)','vlastní']},{id:'pricka_mm',label:'Vlastní poloha příčky (mm)',type:'number',hint:'Jen při vlastní'},{id:'bprofilu',label:'Barva profilu',type:'select',options:['bílá','hnědá','zlatý dub — renolit','RAL — vlastní']},{id:'bsitoviny',label:'Barva síťoviny',type:'select',options:['Š — šedá','Č — černá']},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,o.typ_profil,+o.pocet||1,+o.sirka_kridla,+o.vyska_vod,+o.delka_list||'',o.poloha_pricka+(o.pricka_mm?' ('+o.pricka_mm+' mm)':''),o.bprofilu,o.bsitoviny,o.poznamka]},
  sit_pevna:{label:'Pevná okenní síť',hasM2:true,
    headers:['Pozice','Profil','Počet ks','Šířka (mm)','Výška (mm)','m²','Barva profilu','Kartáček','Výška kartáčku','Síťovina','Způsob uchycení','Výška OD držáku','Nýtování','Provedení rohů','Příčka — počet','Příčka — výška 1','Příčka — výška 2','Poznámka'],
    colWidths:[12,14,8,12,12,8,14,14,14,18,16,14,10,14,12,14,14,22],
    fields:[{id:'mis',label:'Pozice / místnost',type:'text',required:true},{id:'profil',label:'Profil',type:'select',options:['OV 25x10','OE 24x24','ISSO OV 19x8','ISSO OV 25x10','ISSO OE 34x9 LUX','ISSO OE 25x10','ISSO OE 19x8','OE 32x11 LUX','ISSO 37x10']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'bprofilu',label:'Barva profilu',type:'text',hint:'Viz katalog'},{id:'kartacek',label:'Kartáček',type:'select',options:['1 — bez kartáčku','2 — na výšku','3 — na šířku','4 — po celém obvodu']},{id:'vys_kartacku',label:'Výška kartáčku (mm)',type:'select',options:['3','5','8','12','15','18']},{id:'sitovina',label:'Síťovina',type:'select',options:['Š — šedá','Č — černá','P — protipylová černá','N — nanovlákno černá','PSČ — pet screen černá','PSŠ — pet screen šedá']},{id:'uchyceni',label:'Způsob uchycení',type:'select',options:['OD — otočný držák','Z — Z držák','O — obrtlík','P — pružinový kolík']},{id:'vys_od',label:'Výška OD držáku',type:'number'},{id:'nytovani',label:'Nýtování',type:'select',options:['ANO','NE']},{id:'rohy',label:'Provedení rohů',type:'select',options:['vnější','vnitřní']},{id:'pricka_pocet',label:'Příčka — počet',type:'number'},{id:'pricka_vys1',label:'Příčka — výška 1 (mm)',type:'number'},{id:'pricka_vys2',label:'Příčka — výška 2 (mm)',type:'number'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,o.profil,+o.pocet||1,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.bprofilu,o.kartacek,o.vys_kartacku,o.sitovina,o.uchyceni,+o.vys_od||'',o.nytovani,o.rohy,+o.pricka_pocet||'',+o.pricka_vys1||'',+o.pricka_vys2||'',o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
  sit_rolovaci:{label:'Rolovací síť',hasM2:true,
    headers:['Pozice','Typ','Počet ks','Šířka (mm)','Výška (mm)','m²','Barva box+vod.lišty','Barva síťoviny','Montáž','Typ montáže','Typ dorazů','Brzda','Poznámka'],
    colWidths:[14,14,8,12,12,8,20,12,14,22,18,10,24],
    fields:[{id:'mis',label:'Pozice / místnost',type:'text',required:true},{id:'typ',label:'Typ',type:'select',options:['okenní','střešní VERSA','dveřní DAROS']},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'barva_box',label:'Barva box + vod. lišty',type:'select',options:['B — bílá','H — hnědá','zlatý dub — renolit 21','tmavý dub — 24','sapeli — 26','tmavý ořech — 28','RAL — vlastní']},{id:'bsitoviny',label:'Barva síťoviny',type:'select',options:['Š — šedá','Č — černá']},{id:'montaz',label:'Montáž',type:'select',options:['ostění','rám okna','rám dveří','střešní okno']},{id:'typ_montaze',label:'Typ montáže',type:'select',options:['1 — šrouby + klipsy','2 — plastový mont. úchyt','3 — plast. úchyt + klipsy','4 — střešní okno']},{id:'typ_dorazu',label:'Typ dorazů',type:'select',options:['1 — koncový doraz','2 — záchytná lišta','3 — klik-klak']},{id:'brzda',label:'Brzda',type:'select',options:['ANO','NE']},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,o.typ,+o.pocet||1,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.barva_box,o.bsitoviny,o.montaz,o.typ_montaze,o.typ_dorazu,o.brzda,o.poznamka],
    getM2:o=>m2(o.sirka,o.vyska)},
};

// ===================== STATE =====================
const STORE='zamereni-stav';
let items=[],activTyp=null;
function saveState(){try{localStorage.setItem(STORE,JSON.stringify({klient:g('klient'),zakazka:g('zakazka'),adresa:g('adresa'),datum:g('datum'),technik:g('technik'),items,activTyp}));}catch(e){}}
function loadState(){try{const raw=localStorage.getItem(STORE);if(!raw)return;const s=JSON.parse(raw);['klient','zakazka','adresa','datum','technik'].forEach(id=>{if(s[id])document.getElementById(id).value=s[id];});if(s.items)items=s.items;if(s.activTyp)activTyp=s.activTyp;}catch(e){}}
window.resetState=function(){if(!confirm('Smazat celou zakázku a začít znovu?'))return;try{localStorage.removeItem(STORE);}catch(e){}items=[];activTyp=null;['klient','zakazka','adresa','technik'].forEach(id=>document.getElementById(id).value='');document.getElementById('datum').value=new Date().toISOString().split('T')[0];document.querySelectorAll('.typ-btn').forEach(b=>b.classList.remove('active'));document.getElementById('form-box').style.display='none';renderList();};
function g(id){const el=document.getElementById(id);return el?el.value:'';}
function fv(id){const el=document.getElementById('f_'+id);return el?el.value:'';}

// ===================== RENDER FORM =====================
function renderForm(typ){
  activTyp=typ;const prod=PRODUCTS[typ];
  document.getElementById('form-title').textContent=prod.label;
  document.getElementById('form-box').style.display='block';
  let lastVals={};try{lastVals=JSON.parse(localStorage.getItem('last_'+typ)||'{}');}catch(e){}
  const REMEMBER=['ovladani','bprofilu','typ_latky','kod_latky','uchyceni','vedeni','boc_lista','spod_lista','vod_lista','vynaseci','mont_lprofil','bsitoviny','barva_box','typ_stah','provedeni','sirka_latky','kartacek','sitovina','typ_pantu','magnet','bsite','typ_site','prah','montaz','typ_montaze','typ_dorazu','brzda','profil','typ_profil'];
  const container=document.getElementById('form-fields');
  let html='',i=0;
  while(i<prod.fields.length){const f=prod.fields[i],f2=prod.fields[i+1];if(f2&&prod.fields.length>3){html+=`<div class="row2">${fieldHtml(f,lastVals)}${fieldHtml(f2,lastVals)}</div>`;i+=2;}else{html+=fieldHtml(f,lastVals);i++;}}
  container.innerHTML=html;
  const sirkaaEl=document.getElementById('f_sirkaa'),sirkacEl=document.getElementById('f_sirkac');
  if(sirkaaEl&&sirkacEl){function checkCA(){const d=+(sirkacEl.value)-(+(sirkaaEl.value));const err=document.getElementById('err_sirkaa');if(!sirkaaEl.value||!sirkacEl.value){if(err){err.textContent='';err.classList.remove('show');}sirkaaEl.classList.remove('error');return;}if(d<35||d>50){sirkaaEl.classList.add('error');if(err){err.textContent=`Rozdíl C−A = ${d} mm. Musí být 35–50 mm!`;err.classList.add('show');}}else{sirkaaEl.classList.remove('error');if(err){err.textContent='';err.classList.remove('show');}}}sirkaaEl.addEventListener('input',checkCA);sirkacEl.addEventListener('input',checkCA);}
}
function fieldHtml(f,lastVals){
  const val=lastVals[f.id]!==undefined?lastVals[f.id]:(f.defaultVal||'');
  let input='';
  if(f.type==='select'){input=`<select id="f_${f.id}">`;if(!f.required)input+=`<option value="">—</option>`;f.options.forEach(o=>{input+=`<option value="${o}"${val===o?' selected':''}>${o}</option>`;});input+=`</select>`;}
  else{input=`<input type="${f.type==='number'?'number':'text'}" id="f_${f.id}" value="${val}" ${f.hint&&f.type!=='number'?`placeholder="${f.hint}"`:''}${f.type==='number'?' min="0"':''}>`;  }
  return `<div class="field"><label>${f.label}</label>${input}${f.hint&&f.type!=='number'?`<div class="hint">${f.hint}</div>`:''}<div class="err-msg" id="err_${f.id}"></div></div>`;
}

// ===================== ADD ITEM =====================
document.getElementById('btn-pridat').addEventListener('click',function(){
  if(!activTyp)return;
  const prod=PRODUCTS[activTyp];const obj={_typ:activTyp};let hasError=false;
  prod.fields.forEach(f=>{const val=fv(f.id);obj[f.id]=val;const errEl=document.getElementById('err_'+f.id);const inputEl=document.getElementById('f_'+f.id);if(f.required&&!val){if(inputEl)inputEl.classList.add('error');if(errEl){errEl.textContent='Povinné pole';errEl.classList.add('show');}hasError=true;}else{if(inputEl)inputEl.classList.remove('error');if(f.validate&&val){const err=f.validate(val,obj);if(err){if(inputEl)inputEl.classList.add('error');if(errEl){errEl.textContent=err;errEl.classList.add('show');}hasError=true;}else{if(errEl){errEl.textContent='';errEl.classList.remove('show');}}}}});
  if(obj.sirkac&&obj.sirkaa){const d=+(obj.sirkac)-(+(obj.sirkaa));if(d<35||d>50)hasError=true;}
  if(hasError)return;
  items.push(obj);
  const REMEMBER=['ovladani','bprofilu','typ_latky','kod_latky','uchyceni','vedeni','boc_lista','spod_lista','vod_lista','vynaseci','mont_lprofil','bsitoviny','barva_box','typ_stah','provedeni','sirka_latky','kartacek','sitovina','typ_pantu','magnet','bsite','typ_site','prah','montaz','typ_montaze','typ_dorazu','brzda','profil','typ_profil'];
  const last={};prod.fields.forEach(f=>{if(REMEMBER.includes(f.id))last[f.id]=obj[f.id];});
  try{localStorage.setItem('last_'+activTyp,JSON.stringify(last));}catch(e){}
  if(editingIdx!==null){
    // update existing item
    items[editingIdx]=obj;
    editingIdx=null;
    var addBtn=document.getElementById('btn-pridat');
    addBtn.textContent='+ přidat do zakázky';
    addBtn.style.cssText='';
  } else {
    // clear non-remembered fields for next entry
    prod.fields.forEach(f=>{if(!REMEMBER.includes(f.id)){const el=document.getElementById('f_'+f.id);if(el)el.value=f.defaultVal||'';}});
  }
  saveState();renderList();
});

// ===================== RENDER LIST =====================
function plural(n){return n===1?'1 položka':n>=2&&n<=4?n+' položky':n+' položek';}
let editingIdx=null;

let editingIdx=null;

function cardSummary(o){
  const parts=[];
  if(o.sirka&&o.vyska) parts.push('Rozměr: '+o.sirka+' × '+o.vyska+' mm');
  else if(o.sirkac&&o.vyskad) parts.push('Rozměr: '+o.sirkac+' × '+o.vyskad+' mm');
  else if(o.sirkac&&o.vyskai) parts.push('Rozměr: '+o.sirkac+' × '+o.vyskai+' mm');
  else if(o.sirka_kridla&&o.vyska_vod) parts.push('Křídlo: '+o.sirka_kridla+' × '+o.vyska_vod+' mm');
  if(o.pocet&&+o.pocet>1) parts.push(o.pocet+' ks');
  if(o.ovladani) parts.push('Ovl.: '+o.ovladani);
  if(o.typ_latky) parts.push('Látka: '+o.typ_latky);
  if(o.kod_latky) parts.push('Kód: '+o.kod_latky);
  if(o.bprofilu) parts.push('Profil: '+o.bprofilu);
  if(o.uchyceni) parts.push(o.uchyceni);
  if(o.vedeni) parts.push('Vedení: '+o.vedeni);
  if(o.vod_lista) parts.push('Lišta: '+o.vod_lista);
  if(o.boc_lista) parts.push('Boč. lišta: '+o.boc_lista);
  if(o.typ_pantu) parts.push(o.typ_pantu);
  if(o.bsite) parts.push(o.bsite);
  if(o.typ_lamely) parts.push('Lamela: '+o.typ_lamely);
  if(o.barva_lamely) parts.push('Barva: '+o.barva_lamely);
  if(o.typ) parts.push(o.typ);
  if(o.profil) parts.push(o.profil);
  if(o.typ_profil) parts.push(o.typ_profil);
  if(o.bsitoviny) parts.push(o.bsitoviny);
  if(o.poznamka) parts.push('Pozn.: '+o.poznamka);
  return parts.join('  ·  ');
}

function renderList(){
  const list=document.getElementById('list');
  document.getElementById('badge').textContent=plural(items.length);
  document.getElementById('btn-exp').disabled=items.length===0;
  if(!items.length){list.innerHTML='<div class="empty">Zatím žádná položka.</div>';return;}
  list.innerHTML=items.map((o,i)=>{
    const prod=PRODUCTS[o._typ];
    const label=prod?prod.label:o._typ;
    const summary=cardSummary(o);
    const isEditing=editingIdx===i;
    return '<div class="card'+(isEditing?' card-editing':'')+'">'
      +'<div style="flex:1;min-width:0">'
      +'<span class="card-typ">'+label+'</span>'
      +'<div class="card-name">'+(o.mis||'—')+'</div>'
      +'<div class="card-meta">'+summary+'</div>'
      +'</div>'
      +'<div style="display:flex;gap:4px;flex-shrink:0;margin-left:8px;align-items:flex-start">'
      +'<button class="edit-btn" data-edit="'+i+'" title="Upravit">✎</button>'
      +'<button class="del" data-idx="'+i+'" title="Smazat">✕</button>'
      +'</div></div>';
  }).join('');
}

window.delItem=function(i){
  if(editingIdx===i)editingIdx=null;
  else if(editingIdx!==null&&editingIdx>i)editingIdx--;
  items.splice(i,1);saveState();renderList();
};

function startEdit(i){
  const o=items[i];
  editingIdx=i;
  document.querySelectorAll('.typ-btn').forEach(b=>b.classList.remove('active'));
  var btn=document.querySelector('.typ-btn[data-typ="'+o._typ+'"]');
  if(btn)btn.classList.add('active');
  renderForm(o._typ);
  var prod=PRODUCTS[o._typ];
  prod.fields.forEach(function(f){
    var el=document.getElementById('f_'+f.id);
    if(el)el.value=o[f.id]||f.defaultVal||'';
  });
  var addBtn=document.getElementById('btn-pridat');
  addBtn.textContent='✔ uložit změny';
  addBtn.style.cssText='background:#185FA5;color:#fff;border:none;width:100%;padding:10px;border-radius:8px;font-size:13px;cursor:pointer;margin-top:4px';
  document.getElementById('form-box').scrollIntoView({behavior:'smooth',block:'start'});
  renderList();
}

document.getElementById('list').addEventListener('click',function(e){
  var delBtn=e.target.closest('[data-idx]');
  if(delBtn){var idx=parseInt(delBtn.getAttribute('data-idx'));if(!isNaN(idx))delItem(idx);return;}
  var editBtn=e.target.closest('[data-edit]');
  if(editBtn){var idx=parseInt(editBtn.getAttribute('data-edit'));if(!isNaN(idx))startEdit(idx);}
});

// ===================== TYP BUTTONS =====================
document.querySelectorAll('.typ-btn').forEach(function(btn){btn.addEventListener('click',function(){document.querySelectorAll('.typ-btn').forEach(b=>b.classList.remove('active'));btn.classList.add('active');renderForm(btn.dataset.typ);saveState();});});

// ===================== EXPORT =====================
document.getElementById('btn-exp').addEventListener('click',function(){
  const klient=g('klient')||'—',zakazka=g('zakazka')||'—',adresa=g('adresa')||'—';
  const datum=formatDatum(g('datum')),technik=g('technik')||'—';
  const subtitle=`Klient: ${klient}    ·    Adresa: ${adresa}    ·    Technik: ${technik}    ·    Datum: ${datum}`;
  const order=[],groups={};
  items.forEach(o=>{if(!groups[o._typ]){groups[o._typ]=[];order.push(o._typ);}groups[o._typ].push(o);});
  const sheets=order.map(typ=>{
    const prod=PRODUCTS[typ];
    const title=`Výrobní dokumentace — ${prod.label}  ·  Zakázka ${zakazka}`;
    const dataRows=groups[typ].map(o=>prod.toRow(o));
    let totalM2=null;
    if(prod.hasM2&&prod.getM2){totalM2=Math.round(groups[typ].reduce((s,o)=>s+(prod.getM2(o)||0),0)*100)/100;}
    const xml=buildSheetXml(title,subtitle,prod.headers,prod.colWidths,dataRows,totalM2);
    return {name:prod.label,xml};
  });
  const xlsx=buildMultiSheetXlsx(sheets);
  downloadBlob(xlsx,`${zakazka}_${klient.replace(/\s+/g,'-')}.xlsx`);
});

// ===================== INIT =====================
document.querySelectorAll('#klient,#zakazka,#adresa,#datum,#technik').forEach(function(el){el.addEventListener('input',saveState);});
document.getElementById('datum').value=new Date().toISOString().split('T')[0];
loadState();
if(activTyp){var btn=document.querySelector(`.typ-btn[data-typ="${activTyp}"]`);if(btn){btn.classList.add('active');renderForm(activTyp);}}
renderList();
var btnReset=document.getElementById('btn-reset');
if(btnReset)btnReset.addEventListener('click',resetState);
