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
    headers:['Pozice','Typ žaluzie','Šířka (mm)','Výška (mm)','m²','Ovládání','Délka řetízku (mm)','Barva profilu','Typ lamely','Barva lamely','Domyk. provedení','Délka ovládání (mm)','Distanční podložky','Bar. sladění žebřík+texband','Bezpečnostní prvek','Počet ks','Poznámka'],
    colWidths:[14,16,12,12,8,10,16,14,14,14,14,14,16,20,16,8,24],
    fields:[{id:'mis',label:'Pozice / místnost',type:'text',required:true},{id:'typ_zal',label:'Typ žaluzie',type:'select',required:true,options:['ISOLINE LOCO','PRIM','ECO']},{id:'sirka',label:'Šířka (mm)',type:'number',required:true},{id:'vyska',label:'Výška (mm)',type:'number',required:true},{id:'ovladani',label:'Ovládání',type:'select',options:['P — pravé','L — levé']},{id:'delka_retizku',label:'Délka ovládacího řetízku (mm)',type:'number'},{id:'bprofilu',label:'Barva profilu',type:'text',hint:'např. RAL 9010'},{id:'typ_lamely',label:'Typ lamely',type:'select',options:['25 x 0,18','25 x 0,21','16 x 0,21']},{id:'barva_lamely',label:'Barva lamely',type:'text'},{id:'domyk',label:'Domykací provedení',type:'select',options:['ANO','NE'],defaultVal:'ANO'},{id:'delka_ovl',label:'Délka ovládání — jiná (mm)',type:'number'},{id:'dist_podlozky',label:'Distanční podložky (ks)',type:'number'},{id:'bar_sladeni',label:'Bar. sladění žebřík+texband',type:'select',options:['ANO','NE']},{id:'bezpec',label:'Bezpečnostní prvek',type:'select',options:['ANO','NE'],defaultVal:'NE'},{id:'pocet',label:'Počet ks',type:'number',defaultVal:'1'},{id:'poznamka',label:'Poznámka',type:'text'}],
    toRow:o=>[o.mis,o.typ_zal,+o.sirka,+o.vyska,m2(o.sirka,o.vyska),o.ovladani,+o.delka_retizku||'',o.bprofilu,o.typ_lamely,o.barva_lamely,o.domyk,+o.delka_ovl||'',+o.dist_podlozky||'',o.bar_sladeni,o.bezpec,+o.pocet||1,o.poznamka],
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
window.resetState=function(){if(!confirm('Smazat celou zakázku a začít znovu?'))return;try{localStorage.removeItem(STORE);}catch(e){}items=[];activTyp=null;['klient','zakazka','adresa','technik'].forEach(id=>document.getElementById(id).value='');document.getElementById('datum').value=new Date().toISOString().split('T')[0];var sel=document.getElementById('typ-select');if(sel)sel.value='';document.getElementById('form-box').style.display='none';renderList();};
function g(id){const el=document.getElementById(id);return el?el.value:'';}
function fv(id){const el=document.getElementById('f_'+id);return el?el.value:'';}

// ===================== RENDER FORM =====================
function renderForm(typ){
  activTyp=typ;const prod=PRODUCTS[typ];
  document.getElementById('form-title').textContent=prod.label;
  document.getElementById('form-box').style.display='block';
  let lastVals={};try{lastVals=JSON.parse(localStorage.getItem('last_'+typ)||'{}');}catch(e){}
  const REMEMBER=['mis','ovladani','bprofilu','typ_latky','kod_latky','uchyceni','vedeni','boc_lista','spod_lista','vod_lista','vynaseci','mont_lprofil','bsitoviny','barva_box','typ_stah','provedeni','sirka_latky','kartacek','sitovina','typ_pantu','magnet','bsite','typ_site','prah','montaz','typ_montaze','typ_dorazu','brzda','profil','typ_profil'];
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
  const REMEMBER=['mis','ovladani','bprofilu','typ_latky','kod_latky','uchyceni','vedeni','boc_lista','spod_lista','vod_lista','vynaseci','mont_lprofil','bsitoviny','barva_box','typ_stah','provedeni','sirka_latky','kartacek','sitovina','typ_pantu','magnet','bsite','typ_site','prah','montaz','typ_montaze','typ_dorazu','brzda','profil','typ_profil'];
  const last={};prod.fields.forEach(f=>{if(REMEMBER.includes(f.id))last[f.id]=obj[f.id];});
  try{localStorage.setItem('last_'+activTyp,JSON.stringify(last));}catch(e){}
  if(editingIdx!==null){
    items[editingIdx]=obj;
    editingIdx=null;
    var addBtn=document.getElementById('btn-pridat');
    addBtn.textContent='+ přidat do zakázky';
    addBtn.style.cssText='';
  } else {
    items.push(obj);
    prod.fields.forEach(f=>{if(!REMEMBER.includes(f.id)){const el=document.getElementById('f_'+f.id);if(el)el.value=f.defaultVal||'';}});
  }
  saveState();renderList();
});

// ===================== RENDER LIST =====================
function plural(n){return n===1?'1 položka':n>=2&&n<=4?n+' položky':n+' položek';}
let editingIdx=null;


function cardSummary(o){
  const prod=PRODUCTS[o._typ];
  let dim='';
  if(o.sirka&&o.vyska) dim=o.sirka+' × '+o.vyska+' mm';
  else if(o.sirkac&&o.vyskad) dim=o.sirkac+' × '+o.vyskad+' mm';
  else if(o.sirkac&&o.vyskai) dim=o.sirkac+' × '+o.vyskai+' mm';
  else if(o.sirka_kridla&&o.vyska_vod) dim=o.sirka_kridla+' × '+o.vyska_vod+' mm';
  const pocet=+o.pocet||1;
  let m2val='';
  if(prod&&prod.getM2){m2val=Math.round(prod.getM2(o)*pocet*100)/100+' m²';}
  const hlavni=[];
  if(dim) hlavni.push(dim);
  hlavni.push(pocet+' ks');
  if(m2val) hlavni.push(m2val);
  const detail=[];
  if(o.ovladani) detail.push('Ovl.: '+o.ovladani);
  if(o.typ_latky) detail.push('Látka: '+o.typ_latky);
  if(o.kod_latky) detail.push('Kód: '+o.kod_latky);
  if(o.bprofilu) detail.push('Profil: '+o.bprofilu);
  if(o.uchyceni) detail.push(o.uchyceni);
  if(o.vedeni) detail.push('Vedení: '+o.vedeni);
  if(o.vod_lista) detail.push('Lišta: '+o.vod_lista);
  if(o.boc_lista) detail.push('Boč. lišta: '+o.boc_lista);
  if(o.typ_pantu) detail.push(o.typ_pantu);
  if(o.bsite) detail.push(o.bsite);
  if(o.typ_lamely) detail.push('Lamela: '+o.typ_lamely);
  if(o.barva_lamely) detail.push('Barva: '+o.barva_lamely);
  if(o.typ) detail.push(o.typ);
  if(o.profil) detail.push(o.profil);
  if(o.typ_profil) detail.push(o.typ_profil);
  if(o.bsitoviny) detail.push(o.bsitoviny);
  if(o.poznamka) detail.push('Pozn.: '+o.poznamka);
  return {hlavni:hlavni.join('  ·  '), detail:detail.join('  ·  ')};
}

function renderList(){
  const list=document.getElementById('list');
  document.getElementById('badge').textContent=plural(items.length);
  document.getElementById('btn-exp').disabled=items.length===0;
  var btnPdf=document.getElementById('btn-pdf');if(btnPdf)btnPdf.disabled=items.length===0;
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
      +(summary.hlavni?'<div class="card-hlavni">'+summary.hlavni+'</div>':'')
      +(summary.detail?'<div class="card-meta">'+summary.detail+'</div>':'')
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
  var sel=document.getElementById('typ-select');
  if(sel)sel.value=o._typ;
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

// ===================== TYP SELECT =====================
document.getElementById('typ-select').addEventListener('change',function(){
  var typ=this.value;
  if(!typ)return;
  renderForm(typ);
  saveState();
});

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
    if(prod.hasM2&&prod.getM2){totalM2=Math.round(groups[typ].reduce((s,o)=>s+(prod.getM2(o)||0)*(+o.pocet||1),0)*100)/100;}
    const xml=buildSheetXml(title,subtitle,prod.headers,prod.colWidths,dataRows,totalM2);
    return {name:prod.label,xml};
  });
  const xlsx=buildMultiSheetXlsx(sheets);
  downloadBlob(xlsx,`${zakazka}_${klient.replace(/\s+/g,'-')}.xlsx`);
});

// ===================== PDF EXPORT =====================
document.getElementById('btn-pdf').addEventListener('click',function(){
  const klient=g('klient')||'—',zakazka=g('zakazka')||'—',adresa=g('adresa')||'—';
  const datum=formatDatum(g('datum')),technik=g('technik')||'—';

  // Seskup položky podle typu
  const order=[],groups={};
  items.forEach(o=>{if(!groups[o._typ]){groups[o._typ]=[];order.push(o._typ);}groups[o._typ].push(o);});

  // Vytvoř řádky tabulky — jen místnost + typ produktu, BEZ rozměrů
  let tableRows='';
  let radek=1;
  order.forEach(typ=>{
    const prod=PRODUCTS[typ];
    groups[typ].forEach(o=>{
      const pocet=+o.pocet||1;
      const bg=radek%2===0?'#fdecea':'#ffffff';
      tableRows+=`<tr style="background:${bg}">
        <td style="padding:10px 14px;border-bottom:1px solid #f0e0e0;font-size:13px;color:#1a1a1a">${radek}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #f0e0e0;font-size:13px;font-weight:600;color:#1a1a1a">${o.mis||'—'}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #f0e0e0;font-size:13px;color:#1a1a1a">${prod.label}</td>
        <td style="padding:10px 14px;border-bottom:1px solid #f0e0e0;font-size:13px;text-align:center;color:#1a1a1a">${pocet} ks</td>
      </tr>`;
      radek++;
    });
  });

  const logoSvg=`<svg xmlns="http://www.w3.org/2000/svg" width="130" height="38" viewBox="0 0 165 48" fill="none"><g clip-path="url(#cp)"><path d="M47.2241 15.0551C47.2241 15.9356 47.2526 16.7026 47.2241 17.4412C47.1957 17.8388 47.3948 17.8956 47.7078 17.8956C47.9638 17.8956 48.2198 18.0093 48.4474 18.1229C48.6181 18.2081 51.4345 22.0997 51.5767 22.3554C51.8328 22.8667 51.6336 23.2359 51.0931 23.2643C50.2966 23.2927 49.4716 23.2643 48.675 23.2643C48.2198 23.2643 47.7362 23.2643 47.2241 23.2643C47.2241 24.0029 47.2241 24.6562 47.2241 25.3096C47.2241 25.9913 47.2241 26.0197 47.9353 26.0765C48.2767 26.1049 48.4474 26.3038 48.6181 26.5026C49.5569 27.7241 50.3819 29.0307 51.4345 30.1669C51.4629 30.1954 51.4629 30.2522 51.4914 30.2806C51.9466 31.0475 51.7474 31.4736 50.8655 31.502C50.581 31.502 50.2966 31.502 50.0121 31.502C49.1302 31.502 48.2198 31.502 47.2241 31.502C47.2241 32.1269 47.2241 32.7519 47.2241 33.3484C47.2241 33.6325 47.1103 33.9733 47.2526 34.1722C47.3664 34.3142 47.7647 34.2574 47.9922 34.3426C48.2483 34.4562 48.5043 34.5983 48.6466 34.7971C49.2724 35.6493 49.8414 36.5014 50.5526 37.2684C50.9224 37.6661 51.1784 38.1206 51.4629 38.6035C51.8897 39.3136 51.6905 39.7113 50.894 39.7397C49.9552 39.7681 48.9879 39.7397 48.0491 39.7397C47.8216 39.7397 47.5655 39.7397 47.2241 39.7397C47.1957 40.0806 47.1672 40.4498 47.1672 40.8191C47.1672 41.9553 47.1672 43.0916 47.1672 44.2278C47.1672 44.7959 47.1672 44.7959 47.7931 45.364C48.3052 45.8469 48.6466 46.4151 48.5897 47.182C48.5328 47.7217 48.5897 48.3182 48.4474 48.8295C48.1345 49.909 47.1388 50.3066 46.2 50.5339C44.8345 50.8464 43.5828 50.1078 42.8716 48.858C42.3026 47.8069 42.6155 46.0458 43.5259 45.2504C44.0948 44.7675 44.0664 44.7675 44.0095 44.0006C43.981 43.5745 44.0095 43.1484 44.0095 42.7223C44.0095 41.7849 44.0095 40.8475 44.0095 39.7681C43.5828 39.7681 43.156 39.7681 42.7578 39.7681C29.0457 39.7681 15.3621 39.7681 1.65001 39.7681C1.42242 39.7681 1.16639 39.7397 0.938803 39.7681C0.540527 39.7965 0.0853548 39.7965 9.95025e-06 39.3704C-0.0568866 39.0864 0.113803 38.6887 0.284493 38.4046C1.16639 37.1832 2.07673 36.0185 2.98708 34.8255C3.18622 34.5983 3.35691 34.2574 3.78363 34.3426C4.3526 34.4562 4.52328 34.229 4.46639 33.6609C4.40949 32.9507 4.46639 32.269 4.46639 31.502C3.27156 31.502 2.13363 31.502 0.995699 31.502C0.768113 31.502 0.512079 31.502 0.312941 31.4452C0.0284582 31.3884 -0.113783 30.9339 0.0284582 30.6783C0.369837 30.1101 0.739665 29.542 1.13794 29.0023C1.73535 28.2069 2.33277 27.4116 2.95863 26.6162C3.21466 26.3038 3.49915 25.9629 4.01122 26.0765C4.38104 26.1617 4.49484 25.9629 4.46639 25.5936C4.55173 24.9119 4.55173 24.1449 4.55173 23.2927C3.30001 23.2927 2.07673 23.2927 0.853458 23.2927C0.0569065 23.2927 -0.142231 22.8951 0.284493 22.1849C0.65432 21.6168 1.08104 21.0771 1.47932 20.5374C1.96294 19.8841 2.38966 19.2023 2.90173 18.5774C3.12932 18.2933 3.44225 18.0661 3.75518 17.9241C3.98277 17.8104 4.40949 17.8956 4.52328 17.7536C4.66553 17.5548 4.55173 17.2139 4.55173 16.9298C4.55173 16.3617 4.55173 15.7936 4.55173 15.1403C3.92587 14.9983 3.27156 15.0835 2.6457 15.0835C2.01984 15.0551 1.42242 15.0551 0.796562 15.0835C0.398286 15.1119 0.142251 14.8846 0.0284582 14.6006C-0.0284383 14.4017 0.0853548 14.0609 0.227596 13.862C1.0526 12.7826 1.8776 11.7316 2.6457 10.5954C2.98708 10.0841 3.4138 9.48753 4.2388 9.65797C4.32415 9.68637 4.52328 9.43072 4.52328 9.3171C4.55173 8.52174 4.52328 7.72637 4.52328 6.81739C3.4138 6.81739 2.33277 6.81739 1.25173 6.81739C0.455182 6.81739 9.95025e-06 6.3629 9.95025e-06 5.53913C9.95025e-06 4.26087 9.95025e-06 2.98261 9.95025e-06 1.70435C9.95025e-06 0.624927 0.483631 0.113623 1.53622 0C1.79225 0 2.01984 0 2.27587 0C18.1216 0 33.9388 0 49.7845 0C49.8414 0 49.8698 0 49.9267 0C51.3491 0 51.7759 0.426087 51.8043 1.81797C51.8043 2.9542 51.8043 4.09043 51.8043 5.22667C51.8043 6.53333 51.5198 6.81739 50.1543 6.81739C49.244 6.81739 48.3621 6.81739 47.4517 6.81739C47.4233 6.81739 47.3664 6.8742 47.281 6.93101C47.281 7.58435 47.3379 8.29449 47.281 8.97623C47.2241 9.57275 47.4517 9.71478 47.9922 9.65797C48.5612 9.60116 48.7888 10.1125 49.0733 10.4533C49.8983 11.4759 50.6379 12.527 51.3776 13.6064C51.5767 13.9188 51.8328 14.3449 51.7474 14.6574C51.6336 15.1687 51.0647 15.0267 50.6664 15.0267C49.5285 15.0551 48.419 15.0551 47.2241 15.0551ZM7.70949 17.8672C7.88018 17.8956 7.99397 17.9241 8.13622 17.9241C19.8285 17.9241 31.5207 17.9241 43.2129 17.9241C43.4405 17.9241 43.725 18.0093 43.8957 17.8956C44.0948 17.782 44.3793 17.3275 44.3224 17.2707C43.7535 16.6174 44.2655 15.822 43.981 15.1119C31.8621 15.1119 19.7716 15.1119 7.68104 15.1119C7.70949 16.0493 7.70949 16.9298 7.70949 17.8672ZM7.70949 26.0481C7.79484 26.1049 7.82329 26.1333 7.85173 26.1333C19.6862 26.1333 31.5207 26.1333 43.3552 26.1333C43.5828 26.1333 43.8103 26.1049 44.0095 26.0197C44.1233 25.9629 44.294 25.7072 44.2655 25.6504C43.7535 24.9403 44.294 24.0881 44.0095 23.3496C31.8905 23.3496 19.8 23.3496 7.70949 23.3496C7.70949 24.3154 7.70949 25.1959 7.70949 26.0481ZM7.70949 9.62956C19.8285 9.62956 31.9474 9.62956 44.0664 9.62956C44.0664 8.69217 44.0664 7.81159 44.0664 6.8742C31.9474 6.8742 19.8285 6.8742 7.70949 6.8742C7.70949 7.81159 7.70949 8.69217 7.70949 9.62956ZM44.0664 31.5872C31.8621 31.5872 19.8 31.5872 7.68104 31.5872C7.68104 32.0985 7.68104 32.553 7.68104 33.0359C7.68104 33.4904 7.70949 33.9449 7.73794 34.3426C19.9138 34.3426 31.9759 34.3426 44.0664 34.3426C44.0664 33.4052 44.0664 32.5246 44.0664 31.5872Z" fill="#CB1419"/><path d="M143.038 31.0598C144.262 31.0598 145.343 31.0598 146.452 31.0598C146.452 36.5741 146.452 42.06 146.452 47.6591C145.912 47.7722 145.314 47.6874 144.717 47.7157C144.091 47.7439 143.493 47.7439 142.868 47.7157C142.668 47.7157 142.412 47.546 142.327 47.3763C141.644 46.2735 140.962 45.1989 140.279 44.1243C139.511 42.9366 138.771 41.749 138.003 40.533C137.149 39.2039 136.467 37.79 135.67 36.4044C135.585 36.263 135.471 36.1216 135.357 35.9802C135.329 35.9519 135.3 35.9237 135.243 35.8671C135.187 35.9237 135.101 35.9802 135.101 36.065C135.073 36.263 135.016 36.4892 135.101 36.6306C135.528 37.4507 135.357 38.3556 135.357 39.2039C135.386 42.0317 135.357 44.8313 135.357 47.7157C134.134 47.7157 132.968 47.7157 131.716 47.7157C131.716 42.2297 131.716 36.6872 131.716 31.0881C132.683 31.0881 133.679 31.1446 134.646 31.0598C135.386 31.0032 135.841 31.286 136.211 31.9081C137.121 33.4069 138.117 34.8774 139.027 36.4044C140.336 38.5535 141.644 40.7309 142.953 42.9084C143.038 38.9777 143.038 35.0753 143.038 31.0598Z" fill="#CB1419"/><path d="M80.2812 31.0312C81.4476 31.0312 82.5856 31.0312 83.7235 31.0312C84.0364 31.0312 84.1218 31.1444 84.264 31.4554C84.6623 32.2755 85.2312 33.039 85.7433 33.8025C86.995 35.6688 88.2183 37.5917 89.3562 39.5146C89.9821 40.6174 90.608 41.7203 91.2338 42.8231C91.2054 42.8231 91.2338 42.8231 91.2623 42.7948C91.2907 42.7666 91.3192 42.71 91.3192 42.6817C91.3192 38.8359 91.3192 34.9619 91.3192 31.0595C92.4571 31.0595 93.5666 31.0595 94.7045 31.0595C94.7045 36.5737 94.7045 42.0879 94.7045 47.6869C93.5666 47.6869 92.4571 47.6586 91.3192 47.6869C90.8924 47.6869 90.6364 47.5172 90.4373 47.1496C89.3562 45.1702 88.1045 43.3038 86.8812 41.4092C85.8002 39.7126 84.7476 37.9876 83.8942 36.1495C83.8657 36.0647 83.8942 35.9233 83.8373 35.895C83.7235 35.8385 83.5812 35.8102 83.4675 35.8102C83.4106 35.8102 83.3252 35.9799 83.3537 36.0364C83.9226 37.2807 83.5528 38.6097 83.6097 39.8822C83.6666 41.1547 83.6381 42.4272 83.6381 43.6997C83.6381 45.0005 83.6381 46.3013 83.6381 47.6586C82.4718 47.6586 81.3907 47.6586 80.2812 47.6586C80.2812 42.1727 80.2812 36.6868 80.2812 31.0312Z" fill="#CB1419"/><path d="M99.0566 1.16699C100.479 1.16699 101.816 1.16699 103.239 1.16699C103.239 8.30382 103.239 15.4123 103.239 22.6058C102.613 22.7191 101.901 22.6341 101.219 22.6625C100.507 22.6908 99.8247 22.6625 99.0566 22.6625C99.0566 15.5256 99.0566 8.41711 99.0566 1.16699Z" fill="black"/><path d="M142.014 6.86133C143.408 6.86133 144.773 6.86133 146.167 6.86133C146.167 12.119 146.167 17.3767 146.167 22.6345C144.773 22.6345 143.408 22.6345 142.014 22.6345C142.014 17.3767 142.014 12.1473 142.014 6.86133Z" fill="black"/></g><defs><clipPath id="cp"><rect width="165" height="48" fill="white"/></clipPath></defs></svg>`;

  const html=`<!DOCTYPE html><html lang="cs"><head><meta charset="UTF-8">
<title>Nabídka — ${klient}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;color:#1a1a1a;background:#fff;padding:0}
.page{max-width:800px;margin:0 auto;padding:40px}
.header{background:linear-gradient(135deg,#7a0c10 0%,#cb1419 100%);border-radius:12px;padding:28px 32px;margin-bottom:28px;display:flex;align-items:center;justify-content:space-between}
.header-left h1{font-size:22px;font-weight:700;color:#fff;margin-bottom:4px}
.header-left p{font-size:13px;color:rgba(255,255,255,.75)}
.info-box{background:#f7f4f4;border-radius:10px;padding:18px 22px;margin-bottom:22px;display:grid;grid-template-columns:1fr 1fr;gap:10px 24px}
.info-row{display:flex;flex-direction:column;gap:2px}
.info-label{font-size:10px;font-weight:700;color:#cb1419;text-transform:uppercase;letter-spacing:.08em}
.info-val{font-size:14px;color:#1a1a1a;font-weight:500}
.section-title{font-size:11px;font-weight:700;color:#cb1419;text-transform:uppercase;letter-spacing:.1em;margin-bottom:12px;display:flex;align-items:center;gap:6px}
.section-title::before{content:'';width:3px;height:14px;background:#cb1419;border-radius:2px;display:inline-block;flex-shrink:0}
table{width:100%;border-collapse:collapse;border-radius:10px;overflow:hidden}
thead tr{background:linear-gradient(135deg,#7a0c10,#cb1419)}
thead th{padding:11px 14px;font-size:11px;font-weight:700;color:#fff;text-align:left;text-transform:uppercase;letter-spacing:.06em}
thead th:last-child{text-align:center}
tbody td:last-child{text-align:center}
.footer{margin-top:32px;border-top:2px solid #cb1419;padding-top:20px;display:flex;justify-content:space-between;align-items:flex-start}
.footer-left{font-size:12px;color:#888;line-height:1.7}
.footer-right{text-align:right;font-size:12px;color:#888;line-height:1.7}
.footer-brand{font-size:13px;font-weight:700;color:#cb1419;margin-bottom:2px}
.sign-box{margin-top:28px;display:flex;justify-content:flex-end}
.sign-line{border-top:1px solid #999;width:200px;text-align:center;padding-top:6px;font-size:11px;color:#888}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}.page{padding:20px}}
</style></head><body>
<div class="page">
  <div class="header">
    <div class="header-left">
      <h1>Nabídka stínění</h1>
      <p>Záměření ze dne ${datum}</p>
    </div>
    <img src="data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4KPHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxNjUiIGhlaWdodD0iNDgiIHZpZXdCb3g9IjAgMCAxNjUgNDgiIGZpbGw9Im5vbmUiPgogIDxnIGNsaXAtcGF0aD0idXJsKCNjbGlwMF8xODAwMV8xNTY5MSkiPgogICAgPHBhdGggZD0iTTQ3LjIyNDEgMTUuMDU1MUM0Ny4yMjQxIDE1LjkzNTYgNDcuMjUyNiAxNi43MDI2IDQ3LjIyNDEgMTcuNDQxMkM0Ny4xOTU3IDE3LjgzODggNDcuMzk0OCAxNy44OTU2IDQ3LjcwNzggMTcuODk1NkM0Ny45NjM4IDE3Ljg5NTYgNDguMjE5OCAxOC4wMDkzIDQ4LjQ0NzQgMTguMTIyOUM0OC42MTgxIDE4LjIwODEgNTEuNDM0NSAyMi4wOTk3IDUxLjU3NjcgMjIuMzU1NEM1MS44MzI4IDIyLjg2NjcgNTEuNjMzNiAyMy4yMzU5IDUxLjA5MzEgMjMuMjY0M0M1MC4yOTY2IDIzLjI5MjcgNDkuNDcxNiAyMy4yNjQzIDQ4LjY3NSAyMy4yNjQzQzQ4LjIxOTggMjMuMjY0MyA0Ny43MzYyIDIzLjI2NDMgNDcuMjI0MSAyMy4yNjQzQzQ3LjIyNDEgMjQuMDAyOSA0Ny4yMjQxIDI0LjY1NjIgNDcuMjI0MSAyNS4zMDk2QzQ3LjIyNDEgMjUuOTkxMyA0Ny4yMjQxIDI2LjAxOTcgNDcuOTM1MyAyNi4wNzY1QzQ4LjI3NjcgMjYuMTA0OSA0OC40NDc0IDI2LjMwMzggNDguNjE4MSAyNi41MDI2QzQ5LjU1NjkgMjcuNzI0MSA1MC4zODE5IDI5LjAzMDcgNTEuNDM0NSAzMC4xNjY5QzUxLjQ2MjkgMzAuMTk1NCA1MS40NjI5IDMwLjI1MjIgNTEuNDkxNCAzMC4yODA2QzUxLjk0NjYgMzEuMDQ3NSA1MS43NDc0IDMxLjQ3MzYgNTAuODY1NSAzMS41MDJDNTAuNTgxIDMxLjUwMiA1MC4yOTY2IDMxLjUwMiA1MC4wMTIxIDMxLjUwMkM0OS4xMzAyIDMxLjUwMiA0OC4yMTk4IDMxLjUwMiA0Ny4yMjQxIDMxLjUwMkM0Ny4yMjQxIDMyLjEyNjkgNDcuMjI0MSAzMi43NTE5IDQ3LjIyNDEgMzMuMzQ4NEM0Ny4yMjQxIDMzLjYzMjUgNDcuMTEwMyAzMy45NzMzIDQ3LjI1MjYgMzQuMTcyMkM0Ny4zNjY0IDM0LjMxNDIgNDcuNzY0NyAzNC4yNTc0IDQ3Ljk5MjIgMzQuMzQyNkM0OC4yNDgzIDM0LjQ1NjIgNDguNTA0MyAzNC41OTgzIDQ4LjY0NjYgMzQuNzk3MUM0OS4yNzI0IDM1LjY0OTMgNDkuODQxNCAzNi41MDE0IDUwLjU1MjYgMzcuMjY4NEM1MC45MjI0IDM3LjY2NjEgNTEuMTc4NCAzOC4xMjA2IDUxLjQ2MjkgMzguNjAzNUM1MS44ODk3IDM5LjMxMzYgNTEuNjkwNSAzOS43MTEzIDUwLjg5NCAzOS43Mzk3QzQ5Ljk1NTIgMzkuNzY4MSA0OC45ODc5IDM5LjczOTcgNDguMDQ5MSAzOS43Mzk3QzQ3LjgyMTYgMzkuNzM5NyA0Ny41NjU1IDM5LjczOTcgNDcuMjI0MSAzOS43Mzk3QzQ3LjE5NTcgNDAuMDgwNiA0Ny4xNjcyIDQwLjQ0OTggNDcuMTY3MiA0MC44MTkxQzQ3LjE2NzIgNDEuOTU1MyA0Ny4xNjcyIDQzLjA5MTYgNDcuMTY3MiA0NC4yMjc4QzQ3LjE2NzIgNDQuNzk1OSA0Ny4xNjcyIDQ0Ljc5NTkgNDcuNzkzMSA0NS4zNjRDNDguMzA1MiA0NS44NDY5IDQ4LjY0NjYgNDYuNDE1MSA0OC41ODk3IDQ3LjE4MkM0OC41MzI4IDQ3LjcyMTcgNDguNTg5NyA0OC4zMTgyIDQ4LjQ0NzQgNDguODI5NUM0OC4xMzQ1IDQ5LjkwOSA0Ny4xMzg4IDUwLjMwNjYgNDYuMiA1MC41MzM5QzQ0LjgzNDUgNTAuODQ2NCA0My41ODI4IDUwLjEwNzggNDIuODcxNiA0OC44NThDNDIuMzAyNiA0Ny44MDY5IDQyLjYxNTUgNDYuMDQ1OCA0My41MjU5IDQ1LjI1MDRDNDQuMDk0OCA0NC43Njc1IDQ0LjA2NjQgNDQuNzY3NSA0NC4wMDk1IDQ0LjAwMDZDNDMuOTgxIDQzLjU3NDUgNDQuMDA5NSA0My4xNDg0IDQ0LjAwOTUgNDIuNzIyM0M0NC4wMDk1IDQxLjc4NDkgNDQuMDA5NSA0MC44NDc1IDQ0LjAwOTUgMzkuNzY4MUM0My41ODI4IDM5Ljc2ODEgNDMuMTU2IDM5Ljc2ODEgNDIuNzU3OCAzOS43NjgxQzI5LjA0NTcgMzkuNzY4MSAxNS4zNjIxIDM5Ljc2ODEgMS42NTAwMSAzOS43NjgxQzEuNDIyNDIgMzkuNzY4MSAxLjE2NjM5IDM5LjczOTcgMC45Mzg4MDMgMzkuNzY4MUMwLjU0MDUyNyAzOS43OTY1IDAuMDg1MzU0OCAzOS43OTY1IDkuOTUwMjVlLTA2IDM5LjM3MDRDLTAuMDU2ODg2NiAzOS4wODY0IDAuMTEzODAzIDM4LjY4ODcgMC4yODQ0OTMgMzguNDA0NkMxLjE2NjM5IDM3LjE4MzIgMi4wNzY3MyAzNi4wMTg1IDIuOTg3MDggMzQuODI1NUMzLjE4NjIyIDM0LjU5ODMgMy4zNTY5MSAzNC4yNTc0IDMuNzgzNjMgMzQuMzQyNkM0LjM1MjYgMzQuNDU2MiA0LjUyMzI4IDM0LjIyOSA0LjQ2NjM5IDMzLjY2MDlDNC40MDk0OSAzMi45NTA3IDQuNDY2MzkgMzIuMjY5IDQuNDY2MzkgMzEuNTAyQzMuMjcxNTYgMzEuNTAyIDIuMTMzNjMgMzEuNTAyIDAuOTk1Njk5IDMxLjUwMkMwLjc2ODExMyAzMS41MDIgMC41MTIwNzkgMzEuNTAyIDAuMzEyOTQxIDMxLjQ0NTJDMC4wMjg0NTgyIDMxLjM4ODQgLTAuMTEzNzgzIDMwLjkzMzkgMC4wMjg0NTgyIDMwLjY3ODNDMC4zNjk4MzcgMzAuMTEwMSAwLjczOTY2NSAyOS41NDIgMS4xMzc5NCAyOS4wMDIzQzEuNzM1MzUgMjguMjA2OSAyLjMzMjc3IDI3LjQxMTYgMi45NTg2MyAyNi42MTYyQzMuMjE0NjYgMjYuMzAzOCAzLjQ5OTE1IDI1Ljk2MjkgNC4wMTEyMiAyNi4wNzY1QzQuMzgxMDQgMjYuMTYxNyA0LjQ5NDg0IDI1Ljk2MjkgNC40NjYzOSAyNS41OTM2QzQuNTUxNzMgMjQuOTExOSA0LjU1MTczIDI0LjE0NDkgNC41NTE3MyAyMy4yOTI3QzMuMzAwMDEgMjMuMjkyNyAyLjA3NjczIDIzLjI5MjcgMC44NTM0NTggMjMuMjkyN0MwLjA1NjkwNjUgMjMuMjkyNyAtMC4xNDIyMzEgMjIuODk1MSAwLjI4NDQ5MyAyMi4xODQ5QzAuNjU0MzIgMjEuNjE2OCAxLjA4MTA0IDIxLjA3NzEgMS40NzkzMiAyMC41Mzc0QzEuOTYyOTQgMTkuODg0MSAyLjM4OTY2IDE5LjIwMjMgMi45MDE3MyAxOC41Nzc0QzMuMTI5MzIgMTguMjkzMyAzLjQ0MjI1IDE4LjA2NjEgMy43NTUxOCAxNy45MjQxQzMuOTgyNzcgMTcuODEwNCA0LjQwOTQ5IDE3Ljg5NTYgNC41MjMyOCAxNy43NTM2QzQuNjY1NTMgMTcuNTU0OCA0LjU1MTczIDE3LjIxMzkgNC41NTE3MyAxNi45Mjk4QzQuNTUxNzMgMTYuMzYxNyA0LjU1MTczIDE1Ljc5MzYgNC41NTE3MyAxNS4xNDAzQzMuOTI1ODcgMTQuOTk4MyAzLjI3MTU2IDE1LjA4MzUgMi42NDU3IDE1LjA4MzVDMi4wMTk4NCAxNS4wNTUxIDEuNDIyNDIgMTUuMDU1MSAwLjc5NjU2MiAxNS4wODM1QzAuMzk4Mjg2IDE1LjExMTkgMC4xNDIyNTEgMTQuODg0NiAwLjAyODQ1ODIgMTQuNjAwNkMtMC4wMjg0MzgzIDE0LjQwMTcgMC4wODUzNTQ4IDE0LjA2MDkgMC4yMjc1OTYgMTMuODYyQzEuMDUyNiAxMi43ODI2IDEuODc3NiAxMS43MzE2IDIuNjQ1NyAxMC41OTU0QzIuOTg3MDggMTAuMDg0MSAzLjQxMzggOS40ODc1MyA0LjIzODggOS42NTc5N0M0LjMyNDE1IDkuNjg2MzcgNC41MjMyOCA5LjQzMDcyIDQuNTIzMjggOS4zMTcxQzQuNTUxNzMgOC41MjE3NCA0LjUyMzI4IDcuNzI2MzcgNC41MjMyOCA2LjgxNzM5QzMuNDEzOCA2LjgxNzM5IDIuMzMyNzcgNi44MTczOSAxLjI1MTczIDYuODE3MzlDMC40NTUxODIgNi44MTczOSA5Ljk1MDI1ZS0wNiA2LjM2MjkgOS45NTAyNWUtMDYgNS41MzkxM0M5Ljk1MDI1ZS0wNiA0LjI2MDg3IDkuOTUwMjVlLTA2IDIuOTgyNjEgOS45NTAyNWUtMDYgMS43MDQzNUM5Ljk1MDI1ZS0wNiAwLjYyNDkyNyAwLjQ4MzYzMSAwLjExMzYyMyAxLjUzNjIyIDBDMS43OTIyNSAwIDIuMDE5ODQgMCAyLjI3NTg3IDBDMTguMTIxNiAwIDMzLjkzODggMCA0OS43ODQ1IDBDNDkuODQxNCAwIDQ5Ljg2OTggMCA0OS45MjY3IDBDNTEuMzQ5MSAwIDUxLjc3NTkgMC40MjYwODcgNTEuODA0MyAxLjgxNzk3QzUxLjgwNDMgMi45NTQyIDUxLjgwNDMgNC4wOTA0MyA1MS44MDQzIDUuMjI2NjdDNTEuODA0MyA2LjUzMzMzIDUxLjUxOTggNi44MTczOSA1MC4xNTQzIDYuODE3MzlDNDkuMjQ0IDYuODE3MzkgNDguMzYyMSA2LjgxNzM5IDQ3LjQ1MTcgNi44MTczOUM0Ny40MjMzIDYuODE3MzkgNDcuMzY2NCA2Ljg3NDIgNDcuMjgxIDYuOTMxMDFDNDcuMjgxIDcuNTg0MzUgNDcuMzM3OSA4LjI5NDQ5IDQ3LjI4MSA4Ljk3NjIzQzQ3LjIyNDEgOS41NzI3NSA0Ny40NTE3IDkuNzE0NzggNDcuOTkyMiA5LjY1Nzk3QzQ4LjU2MTIgOS42MDExNiA0OC43ODg4IDEwLjExMjUgNDkuMDczMyAxMC40NTMzQzQ5Ljg5ODMgMTEuNDc1OSA1MC42Mzc5IDEyLjUyNyA1MS4zNzc2IDEzLjYwNjRDNTEuNTc2NyAxMy45MTg4IDUxLjgzMjggMTQuMzQ0OSA1MS43NDc0IDE0LjY1NzRDNTEuNjMzNiAxNS4xNjg3IDUxLjA2NDcgMTUuMDI2NyA1MC42NjY0IDE1LjAyNjdDNDkuNTI4NSAxNS4wNTUxIDQ4LjQxOSAxNS4wNTUxIDQ3LjIyNDEgMTUuMDU1MVpNNy43MDk0OSAxNy44NjcyQzcuODgwMTggMTcuODk1NiA3Ljk5Mzk3IDE3LjkyNDEgOC4xMzYyMiAxNy45MjQxQzE5LjgyODUgMTcuOTI0MSAzMS41MjA3IDE3LjkyNDEgNDMuMjEyOSAxNy45MjQxQzQzLjQ0MDUgMTcuOTI0MSA0My43MjUgMTguMDA5MyA0My44OTU3IDE3Ljg5NTZDNDQuMDk0OCAxNy43ODIgNDQuMzc5MyAxNy4zMjc1IDQ0LjMyMjQgMTcuMjcwN0M0My43NTM1IDE2LjYxNzQgNDQuMjY1NSAxNS44MjIgNDMuOTgxIDE1LjExMTlDMzEuODYyMSAxNS4xMTE5IDE5Ljc3MTYgMTUuMTExOSA3LjY4MTA0IDE1LjExMTlDNy43MDk0OSAxNi4wNDkzIDcuNzA5NDkgMTYuOTI5OCA3LjcwOTQ5IDE3Ljg2NzJaTTcuNzA5NDkgMjYuMDQ4MUM3Ljc5NDg0IDI2LjEwNDkgNy44MjMyOSAyNi4xMzMzIDcuODUxNzMgMjYuMTMzM0MxOS42ODYyIDI2LjEzMzMgMzEuNTIwNyAyNi4xMzMzIDQzLjM1NTIgMjYuMTMzM0M0My41ODI4IDI2LjEzMzMgNDMuODEwMyAyNi4xMDQ5IDQ0LjAwOTUgMjYuMDE5N0M0NC4xMjMzIDI1Ljk2MjkgNDQuMjk0IDI1LjcwNzIgNDQuMjY1NSAyNS42NTA0QzQzLjc1MzUgMjQuOTQwMyA0NC4yOTQgMjQuMDg4MSA0NC4wMDk1IDIzLjM0OTZDMzEuODkwNSAyMy4zNDk2IDE5LjggMjMuMzQ5NiA3LjcwOTQ5IDIzLjM0OTZDNy43MDk0OSAyNC4zMTU0IDcuNzA5NDkgMjUuMTk1OSA3LjcwOTQ5IDI2LjA0ODFaTTcuNzA5NDkgOS42Mjk1NkMxOS44Mjg1IDkuNjI5NTYgMzEuOTQ3NCA5LjYyOTU2IDQ0LjA2NjQgOS42Mjk1NkM0NC4wNjY0IDguNjkyMTcgNDQuMDY2NCA3LjgxMTU5IDQ0LjA2NjQgNi44NzQyQzMxLjk0NzQgNi44NzQyIDE5LjgyODUgNi44NzQyIDcuNzA5NDkgNi44NzQyQzcuNzA5NDkgNy44MTE1OSA3LjcwOTQ5IDguNjkyMTcgNy43MDk0OSA5LjYyOTU2Wk00NC4wNjY0IDMxLjU4NzJDMzEuODYyMSAzMS41ODcyIDE5LjggMzEuNTg3MiA3LjY4MTA0IDMxLjU4NzJDNy42ODEwNCAzMi4wOTg1IDcuNjgxMDQgMzIuNTUzIDcuNjgxMDQgMzMuMDM1OUM3LjY4MTA0IDMzLjQ5MDQgNy43MDk0OSAzMy45NDQ5IDcuNzM3OTQgMzQuMzQyNkMxOS45MTM4IDM0LjM0MjYgMzEuOTc1OSAzNC4zNDI2IDQ0LjA2NjQgMzQuMzQyNkM0NC4wNjY0IDMzLjQwNTIgNDQuMDY2NCAzMi41MjQ2IDQ0LjA2NjQgMzEuNTg3MloiIGZpbGw9IiNDQjE0MTkiPjwvcGF0aD4KICAgIDxwYXRoIGQ9Ik0xNDMuMDM4IDMxLjA1OThDMTQ0LjI2MiAzMS4wNTk4IDE0NS4zNDMgMzEuMDU5OCAxNDYuNDUyIDMxLjA1OThDMTQ2LjQ1MiAzNi41NzQxIDE0Ni40NTIgNDIuMDYgMTQ2LjQ1MiA0Ny42NTkxQzE0NS45MTIgNDcuNzcyMiAxNDUuMzE0IDQ3LjY4NzQgMTQ0LjcxNyA0Ny43MTU3QzE0NC4wOTEgNDcuNzQzOSAxNDMuNDkzIDQ3Ljc0MzkgMTQyLjg2OCA0Ny43MTU3QzE0Mi42NjggNDcuNzE1NyAxNDIuNDEyIDQ3LjU0NiAxNDIuMzI3IDQ3LjM3NjNDMTQxLjY0NCA0Ni4yNzM1IDE0MC45NjIgNDUuMTk4OSAxNDAuMjc5IDQ0LjEyNDNDMTM5LjUxMSA0Mi45MzY2IDEzOC43NzEgNDEuNzQ5IDEzOC4wMDMgNDAuNTMzQzEzNy4xNDkgMzkuMjAzOSAxMzYuNDY3IDM3Ljc5IDEzNS42NyAzNi40MDQ0QzEzNS41ODUgMzYuMjYzIDEzNS40NzEgMzYuMTIxNiAxMzUuMzU3IDM1Ljk4MDJDMTM1LjMyOSAzNS45NTE5IDEzNS4zIDM1LjkyMzcgMTM1LjI0MyAzNS44NjcxQzEzNS4xODcgMzUuOTIzNyAxMzUuMTAxIDM1Ljk4MDIgMTM1LjEwMSAzNi4wNjVDMTM1LjA3MyAzNi4yNjMgMTM1LjAxNiAzNi40ODkyIDEzNS4xMDEgMzYuNjMwNkMxMzUuNTI4IDM3LjQ1MDcgMTM1LjM1NyAzOC4zNTU2IDEzNS4zNTcgMzkuMjAzOUMxMzUuMzg2IDQyLjAzMTcgMTM1LjM1NyA0NC44MzEzIDEzNS4zNTcgNDcuNzE1N0MxMzQuMTM0IDQ3LjcxNTcgMTMyLjk2OCA0Ny43MTU3IDEzMS43MTYgNDcuNzE1N0MxMzEuNzE2IDQyLjIyOTcgMTMxLjcxNiAzNi42ODcyIDEzMS43MTYgMzEuMDg4MUMxMzIuNjgzIDMxLjA4ODEgMTMzLjY3OSAzMS4xNDQ2IDEzNC42NDYgMzEuMDU5OEMxMzUuMzg2IDMxLjAwMzIgMTM1Ljg0MSAzMS4yODYgMTM2LjIxMSAzMS45MDgxQzEzNy4xMjEgMzMuNDA2OSAxMzguMTE3IDM0Ljg3NzQgMTM5LjAyNyAzNi40MDQ0QzE0MC4zMzYgMzguNTUzNSAxNDEuNjQ0IDQwLjczMDkgMTQyLjk1MyA0Mi45MDg0QzE0My4wMzggMzguOTc3NyAxNDMuMDM4IDM1LjA3NTMgMTQzLjAzOCAzMS4wNTk4WiIgZmlsbD0iI0NCMTQxOSI+PC9wYXRoPgogICAgPHBhdGggZD0iTTgwLjI4MTIgMzEuMDMxMkM4MS40NDc2IDMxLjAzMTIgODIuNTg1NiAzMS4wMzEyIDgzLjcyMzUgMzEuMDMxMkM4NC4wMzY0IDMxLjAzMTIgODQuMTIxOCAzMS4xNDQ0IDg0LjI2NCAzMS40NTU0Qzg0LjY2MjMgMzIuMjc1NSA4NS4yMzEyIDMzLjAzOSA4NS43NDMzIDMzLjgwMjVDODYuOTk1IDM1LjY2ODggODguMjE4MyAzNy41OTE3IDg5LjM1NjIgMzkuNTE0NkM4OS45ODIxIDQwLjYxNzQgOTAuNjA4IDQxLjcyMDMgOTEuMjMzOCA0Mi44MjMxQzkxLjIwNTQgNDIuODIzMSA5MS4yMzM4IDQyLjgyMzEgOTEuMjYyMyA0Mi43OTQ4QzkxLjI5MDcgNDIuNzY2NiA5MS4zMTkyIDQyLjcxIDkxLjMxOTIgNDIuNjgxN0M5MS4zMTkyIDM4LjgzNTkgOTEuMzE5MiAzNC45NjE5IDkxLjMxOTIgMzEuMDU5NUM5Mi40NTcxIDMxLjA1OTUgOTMuNTY2NiAzMS4wNTk1IDk0LjcwNDUgMzEuMDU5NUM5NC43MDQ1IDM2LjU3MzcgOTQuNzA0NSA0Mi4wODc5IDk0LjcwNDUgNDcuNjg2OUM5My41NjY2IDQ3LjY4NjkgOTIuNDU3MSA0Ny42NTg2IDkxLjMxOTIgNDcuNjg2OUM5MC44OTI0IDQ3LjY4NjkgOTAuNjM2NCA0Ny41MTcyIDkwLjQzNzMgNDcuMTQ5NkM4OS4zNTYyIDQ1LjE3MDIgODguMTA0NSA0My4zMDM4IDg2Ljg4MTIgNDEuNDA5MkM4NS44MDAyIDM5LjcxMjYgODQuNzQ3NiAzNy45ODc2IDgzLjg5NDIgMzYuMTQ5NUM4My44NjU3IDM2LjA2NDcgODMuODk0MiAzNS45MjMzIDgzLjgzNzMgMzUuODk1QzgzLjcyMzUgMzUuODM4NSA4My41ODEyIDM1LjgxMDIgODMuNDY3NSAzNS44MTAyQzgzLjQxMDYgMzUuODEwMiA4My4zMjUyIDM1Ljk3OTkgODMuMzUzNyAzNi4wMzY0QzgzLjkyMjYgMzcuMjgwNyA4My41NTI4IDM4LjYwOTcgODMuNjA5NyAzOS44ODIyQzgzLjY2NjYgNDEuMTU0NyA4My42MzgxIDQyLjQyNzIgODMuNjM4MSA0My42OTk3QzgzLjYzODEgNDUuMDAwNSA4My42MzgxIDQ2LjMwMTMgODMuNjM4MSA0Ny42NTg2QzgyLjQ3MTggNDcuNjU4NiA4MS4zOTA3IDQ3LjY1ODYgODAuMjgxMiA0Ny42NTg2QzgwLjI4MTIgNDIuMTcyNyA4MC4yODEyIDM2LjY4NjggODAuMjgxMiAzMS4wMzEyWiIgZmlsbD0iI0NCMTQxOSI+PC9wYXRoPgogICAgPHBhdGggZD0iTTE2NC40MzEgMTUuODUwMkMxNjAuODE4IDE1Ljg1MDIgMTU3LjI5IDE1Ljg1MDIgMTUzLjY3OCAxNS44NTAyQzE1My41NjQgMTYuNzU1IDE1My42MjEgMTcuNjAzMiAxNTQuMjE4IDE4LjI4MThDMTU0Ljk4NiAxOS4xNTgzIDE1NS44NjggMTkuODM2OSAxNTcuMTIgMTkuODM2OUMxNTguNDU3IDE5LjgzNjkgMTU5LjYyMyAxOS41NTQyIDE2MC4yNzggMTguMTk3QzE2MC40NDggMTcuODI5NCAxNjAuOTAzIDE3LjY1OTggMTYxLjIxNiAxNy45MTQyQzE2MS42MTUgMTguMjUzNSAxNjIuMDEzIDE4LjA1NTYgMTYyLjQxMSAxOC4xMTIyQzE2My4wMzcgMTguMjI1MyAxNjMuNjYzIDE4LjMxMDEgMTY0LjI4OSAxOC4zOTQ5QzE2NC4zMTcgMTguMzk0OSAxNjQuMzQ2IDE4LjU2NDYgMTY0LjMxNyAxOC42NDk0QzE2My45NzYgMjAuMTE5NyAxNjIuOTUyIDIxLjAyNDUgMTYxLjcyOCAyMS43ODc5QzE2MC40NDggMjIuNjY0NSAxNTkuMDgzIDIzLjAzMiAxNTcuNjAzIDIyLjk0NzJDMTU2LjY5MyAyMi45MTg5IDE1NS43ODMgMjIuOTc1NSAxNTQuODcyIDIyLjY5MjdDMTUzLjM2NSAyMi4yMTIxIDE1MS4zNDUgMjEuMzA3MyAxNTAuNDYzIDE5LjUyNTlDMTQ5LjgwOSAxOC4xOTcgMTQ5LjI0IDE2Ljg2ODEgMTQ5LjMyNSAxNS4zNDEyQzE0OS4zNTMgMTQuNzE5MSAxNDkuMjk2IDE0LjEyNTQgMTQ5LjM1MyAxMy41MDMzQzE0OS40OTYgMTEuODM1MSAxNDkuOTc5IDEwLjI4IDE1MS4yMzEgOS4wNjQxMkMxNTEuNzE1IDguNTgzNDUgMTUyLjIyNyA4LjEwMjc3IDE1Mi43OTYgNy42Nzg2NEMxNTMuOTA1IDYuODU4NjcgMTU1LjE4NSA2LjQ5MTA5IDE1Ni41NzkgNi41NDc2NEMxNTcuMDkxIDYuNTc1OTEgMTU3LjYzMiA2LjUxOTM2IDE1OC4xNDQgNi41NzU5MUMxNTguNTcxIDYuNjA0MTkgMTU5LjAyNiA2LjcxNzI5IDE1OS40NTIgNi44NTg2N0MxNjAuMTM1IDcuMTEzMTQgMTYwLjkzMiA3LjMxMTA3IDE2MS40NzIgNy43NjM0N0MxNjMuMjA4IDkuMjA1NSAxNjQuNDMxIDEwLjkzMDMgMTY0LjQzMSAxMy4zMDU0QzE2NC40MzEgMTQuMTUzNiAxNjQuNDMxIDE0Ljk3MzYgMTY0LjQzMSAxNS44NTAyWk0xNTcuMDYzIDEzLjA1MDlDMTU4LjA1OSAxMy4wNTA5IDE1OS4wNTQgMTMuMDIyNiAxNjAuMDUgMTMuMDUwOUMxNjAuNDc3IDEzLjA3OTIgMTYwLjUzNCAxMi44NTMgMTYwLjQ3NyAxMi41NDJDMTYwLjE5MiAxMC45MzAzIDE1OC42ODQgOS42ODYxOCAxNTcuMDYzIDkuNjI5NjNDMTU2LjQzNyA5LjYwMTM1IDE1NS44NCA5Ljc3MSAxNTUuNDEzIDEwLjA4MkMxNTQuNzAyIDEwLjU5MSAxNTQuMDc2IDExLjIxMyAxNTMuNzA2IDEyLjAzM0MxNTMuNjIxIDEyLjE3NDQgMTUzLjY3NyAxMi40MDA2IDE1My42NDkgMTIuNTcwMkMxNTMuNjIxIDEyLjkwOTUgMTUzLjcwNiAxMy4wNzkyIDE1NC4xMDQgMTMuMDUwOUMxNTUuMDcxIDEzLjA1MDkgMTU2LjA2NyAxMy4wNTA5IDE1Ny4wNjMgMTMuMDUwOVoiIGZpbGw9ImJsYWNrIj48L3BhdGg+CiAgICA8cGF0aCBkPSJNOTUuMzAxOSAyMi42MzY5QzkzLjkzNjMgMjIuNjM2OSA5Mi42ODQ2IDIyLjYzNjkgOTEuMzE5MSAyMi42MzY5QzkxLjMxOTEgMjEuOTg2NSA5MS4zMTkxIDIxLjMzNjIgOTEuMzE5MSAyMC42ODU4QzkwLjYwNzkgMjEuMjUxNCA4OS44OTY3IDIxLjY3NTUgODkuMzU2MiAyMi4yNjkzQzg4Ljk4NjMgMjIuNjY1MiA4OC41MzEyIDIyLjY5MzQgODguMTYxMyAyMi44MDY2Qzg2LjE5ODQgMjMuNDg1MiA4My4yMTEzIDIyLjY2NTIgODEuOTAyNyAyMS4wNTM0QzgxLjUzMjkgMjAuNjAxIDgxLjI0ODQgMjAuMDA3MiA4MS4xMDYyIDE5LjQxMzRDODAuNDgwMyAxNi45ODE3IDgxLjg3NDMgMTQuNjM0OCA4NC4zNDkzIDEzLjg5OTZDODUuMzczNCAxMy41ODg2IDg2LjQyNiAxMy4yNzc1IDg3LjUwNyAxMy4zMzQxQzg4LjMzMiAxMy4zNjIzIDg5LjEwMDEgMTMuMDIzIDg5Ljk1MzYgMTIuOTk0OEM5MC45NDkzIDEyLjk2NjUgOTEuMzQ3NSAxMS44NjM3IDkwLjkyMDggMTAuNzg5MkM5MC42MDc5IDkuOTk3NDkgODkuODM5OCA5Ljc3MTI4IDg5LjE1NyA5LjY1ODE4Qzg3Ljk2MjIgOS40NjAyNSA4Ni44NTI3IDkuNjU4MTggODYuMDg0NiAxMC43NjA5Qzg1LjkxMzkgMTAuOTg3MiA4NS42ODYzIDExLjE1NjggODUuNDg3MiAxMS4zNTQ3Qzg1LjMxNjUgMTEuNTUyNyA4NS4yMDI3IDExLjQ5NjEgODQuOTc1MSAxMS40MTEzQzg0LjYwNTMgMTEuMjY5OSA4NC4xNTAxIDExLjM4MyA4My43MjM0IDExLjMyNjVDODMuMDk3NSAxMS4yNjk5IDgyLjQ3MTcgMTEuMTU2OCA4MS44NDU4IDExLjA0MzdDODEuNzYwNSAxMS4wMTU0IDgxLjYxODIgMTAuNzA0NCA4MS42NzUxIDEwLjU5MTNDODEuODc0MyAxMC4wNTQgODIuMDQ1IDkuNDg4NTIgODIuNDE0OCA5LjA2NDM4QzgzLjM4MiA3LjkwNTA2IDg0LjYwNTMgNy4xNjk4OCA4Ni4xNDE1IDYuNzc0MDJDODcuNTA3IDYuNDM0NzEgODguODE1NiA2LjU3NjA5IDkwLjE1MjcgNi41NzYwOUM5MC44NjM5IDYuNTc2MDkgOTEuNjYwNSA2Ljg1ODg1IDkyLjMxNDggNy4xNDE2MUM5My43MDg4IDcuNzkxOTYgOTQuNjQ3NSA4Ljk3OTU1IDk1LjMwMTkgMTAuMzY1MUM5NS40NDQxIDEwLjY0NzggOTUuMzMwMyAxMS4wMTU0IDk1LjMzMDMgMTEuMzU0N0M5NS4zMDE5IDE1LjA4NzIgOTUuMzAxOSAxOC43OTE0IDk1LjMwMTkgMjIuNjM2OVpNOTEuMzE5MSAxNi4zODc5QzkxLjMxOTEgMTYuMTMzNCA5MS4zMTkxIDE2LjAyMDMgOTEuMzE5MSAxNS45MzU1QzkxLjMxOTEgMTUuMzQxNyA5MS4wOTE1IDE1LjE0MzcgOTAuNTc5NCAxNS4zNjk5Qzg5Ljk1MzYgMTUuNjUyNyA4OS4yOTkzIDE1LjczNzUgODguNjczNCAxNS44Nzg5Qzg4LjAxOTEgMTYuMDIwMyA4Ny4zMDc5IDE2LjAyMDMgODYuNjgyIDE2LjE5Qzg2LjI4MzggMTYuMzAzMSA4NS44NTcgMTYuNjE0MSA4NS42MDEgMTYuOTUzNEM4NS4zNDUgMTcuMjY0NCA4NS4xMTc0IDE3Ljc3MzQgODUuMjAyNyAxOC4wODQ0Qzg1LjM0NSAxOC41OTM0IDg1LjM0NSAxOS4yMTU1IDg1Ljg4NTUgMTkuNTU0OEM4Ny4xMDg4IDIwLjM0NjUgODkuMTI4NiAyMC4xMjAzIDkwLjIwOTYgMTkuMTMwN0M5MS4wOTE1IDE4LjMzODkgOTEuMDkxNSAxNy4yMzYyIDkxLjMxOTEgMTYuMzg3OVoiIGZpbGw9ImJsYWNrIj48L3BhdGg+CiAgICA8cGF0aCBkPSJNMTA2Ljk2NiA2Ljg2MTMzQzEwOC40MTcgNi44NjEzMyAxMDkuNzgyIDYuODYxMzMgMTExLjIzMyA2Ljg2MTMzQzExMS4yMzMgOC44NDAzNSAxMTEuMjMzIDEwLjgxOTQgMTExLjIzMyAxMi43NzAxQzExMS4yMzMgMTQuMDQyNCAxMTEuMTQ4IDE1LjMxNDYgMTExLjI2MiAxNi41ODY4QzExMS4zNzUgMTcuNjYxMiAxMTEuNzQ1IDE4LjczNTUgMTEyLjk0IDE5LjEwM0MxMTQuNjc1IDE5LjY2ODUgMTE2Ljk1MSAxOS4wMTgyIDExNy4yNjQgMTYuODY5NUMxMTcuMzc4IDE2LjAyMTQgMTE3LjQ5MiAxNS4xNDUgMTE3LjQ5MiAxNC4yOTY4QzExNy40OTIgMTEuODM3MiAxMTcuNDkyIDkuNDA1NzkgMTE3LjQ5MiA2Ljg4OTZDMTE4Ljg4NiA2Ljg4OTYgMTIwLjI4IDYuODg5NiAxMjEuNzMgNi44ODk2QzEyMS43MyAxMi4wOTE2IDEyMS43MyAxNy4zMjE5IDEyMS43MyAyMi42MDg3QzEyMC4zNjUgMjIuNjA4NyAxMTguOTcxIDIyLjYwODcgMTE3LjUyIDIyLjYwODdDMTE3LjUyIDIxLjcwNCAxMTcuNTIgMjAuODI3NiAxMTcuNTIgMTkuNjExOUMxMTcuMjY0IDIwLjA2NDMgMTE3LjE3OSAyMC4yMzM5IDExNy4wNjUgMjAuNDMxOEMxMTYuOTIzIDIwLjY1OCAxMTYuODA5IDIwLjkxMjQgMTE2LjYzOCAyMS4xMTAzQzExNS43NTYgMjIuMjQxMiAxMTQuNTA1IDIyLjcyMTggMTEzLjE2OCAyMi45NDhDMTEyLjAwMSAyMy4xNDU5IDExMC45MiAyMi43NTAxIDEwOS44MzkgMjIuMzgyNkMxMDkuNzgyIDIyLjM4MjYgMTA5LjcyNSAyMi40MTA4IDEwOS42OTcgMjIuMzgyNkMxMDguNjczIDIxLjM2NDggMTA3LjUwNiAyMC40ODgzIDEwNy4yMjIgMTguODc2OEMxMDcuMDUxIDE4LjAwMDQgMTA2Ljk5NCAxNy4xNTIzIDEwNi45OTQgMTYuMzA0MUMxMDYuOTY2IDEzLjIyMjUgMTA2Ljk2NiAxMC4xMTI2IDEwNi45NjYgNi44NjEzM1oiIGZpbGw9ImJsYWNrIj48L3BhdGg+CiAgICA8cGF0aCBkPSJNMTE0LjkzMSA0Ny43MTU0QzExNC45MzEgNDIuMDg3OSAxMTQuOTMxIDM2LjYzMDIgMTE0LjkzMSAzMS4wODc2QzExNS4yNzIgMzEuMDU5MyAxMTUuNjQyIDMxLjAzMSAxMTYuMDEyIDMxLjAzMUMxMTguMDYgMzEuMDMxIDEyMC4wOCAzMC45NzQ0IDEyMi4xMjggMzEuMDMxQzEyNS40IDMxLjE0NDEgMTI3LjUwNSAzMi43ODQzIDEyOC44MTMgMzUuNjEyMUMxMjkuMDEzIDM2LjAwOCAxMjguNyAzNi41NDUzIDEyOS4xODMgMzYuODg0N0MxMjkuMDk4IDM3LjkwMjcgMTI5LjYzOCAzOC44MzU5IDEyOS40NjggMzkuODgyMkMxMjkuMDk4IDQyLjA4NzkgMTI4LjcyOCA0NC4yOTM3IDEyNi45MDcgNDUuODc3M0MxMjUuNTk5IDQ3LjAwODQgMTI0LjAzNCA0Ny42ODcxIDEyMi4yNyA0Ny43MTU0QzExOS44NTIgNDcuNzQzNyAxMTcuNDM0IDQ3LjcxNTQgMTE0LjkzMSA0Ny43MTU0Wk0xMTguNjI5IDQwLjA1MTlDMTE4LjYyOSA0MC4zMzQ3IDExOC42NTcgNDAuNjE3NCAxMTguNjI5IDQwLjkwMDJDMTE4LjUxNSA0MS45NDY1IDExOC44ODUgNDMuMDIxMSAxMTguMzczIDQ0LjAzOTJDMTE4LjIwMiA0NC4zNTAyIDExOC41NzIgNDQuODg3NSAxMTguODg1IDQ0Ljg1OTJDMTIwLjIyMiA0NC44MzEgMTIxLjU4OCA0NS4wNTcyIDEyMi44OTYgNDQuNjA0N0MxMjQuODMxIDQzLjk1NDMgMTI1Ljc2OSA0Mi4zMTQyIDEyNS43NjkgNDAuNjc0QzEyNS43NjkgNDAuMTkzMyAxMjYuMDU0IDM5LjY4NDMgMTI1LjkxMiAzOS4yODg0QzEyNS42ODQgMzguNzIyOCAxMjUuODI2IDM4LjE4NTUgMTI1Ljc2OSAzNy42NDgyQzEyNS42NTYgMzYuNDg4OCAxMjUuMTcyIDM1LjM4NTkgMTI0LjIzMyAzNC43OTIxQzEyMy4zOCAzNC4yMjY1IDEyMi4zNTYgMzMuNzQ1NyAxMjEuMjE4IDMzLjgzMDZDMTIwLjY0OSAzMy44ODcxIDEyMC4wOCAzMy44MzA2IDExOS41MTEgMzMuODMwNkMxMTguNDg3IDMzLjgzMDYgMTE4LjQ1OCAzMy45NDM3IDExOC4zNDQgMzQuOTA1MkMxMTguMjMxIDM1LjcyNTIgMTE4LjY1NyAzNi40MDM5IDExOC42NTcgMzcuMTk1N0MxMTguNjI5IDM4LjE4NTUgMTE4LjYyOSAzOS4xMTg3IDExOC42MjkgNDAuMDUxOVoiIGZpbGw9IiNDQjE0MTkiPjwvcGF0aD4KICAgIDxwYXRoIGQ9Ik0xNDkuMDY5IDM4Ljk3N0MxNDguODk4IDM3Ljc4OSAxNDkuMzI1IDM2LjQwMyAxNDkuOTIzIDM1LjAxN0MxNTAuODMzIDMyLjk4MDUgMTUyLjM0MSAzMS42MjI4IDE1NC41MDMgMzEuMDU3MUMxNTYuNzIyIDMwLjQ2MzEgMTU4Ljk5OCAzMC41NDggMTYxLjMwMiAzMS44Nzc0QzE2Mi41ODIgMzIuNjQxMSAxNjQuMTE4IDM0LjgxOSAxNjQuNDYgMzYuMjg5OUMxNjQuNjg3IDM3LjMzNjQgMTY1LjE0MiAzOC4zODMgMTY1IDM5LjQyOTVDMTY0LjgwMSA0MC45ODUyIDE2NC43NDQgNDIuNTk3NSAxNjMuODkxIDQ0LjA0QzE2Mi44NjcgNDUuNzkzNyAxNjEuNDQ0IDQ3LjAwOTkgMTU5LjUzOCA0Ny42ODg4QzE1OC4xNDQgNDguMTY5NiAxNTUuNjk4IDQ4LjExMzEgMTU0LjI3NSA0Ny42MzIyQzE1MS42MyA0Ni43NTU0IDE1MC4xNzkgNDQuODAzNyAxNDkuMzU0IDQyLjI4NjNDMTQ5LjAxMiA0MS4zMjQ2IDE0OS4wOTggNDAuMzM0NiAxNDkuMDY5IDM4Ljk3N1pNMTYxLjMwMiAzOS43NDA3QzE2MS4zMDIgMzguODYzOCAxNjEuMzMgMzguMTg1IDE2MS4zMDIgMzcuNTM0NEMxNjEuMzAyIDM3LjMzNjQgMTYxLjEwMyAzNy4xMzg0IDE2MS4wNDYgMzYuOTQwNEMxNjAuNTM0IDM1LjA3MzYgMTU5LjQyNCAzMy45MTM5IDE1Ny4zNDggMzMuODU3M0MxNTcuMDkyIDMzLjg1NzMgMTU2LjkyMSAzMy40MzMxIDE1Ni41NTEgMzMuODI5MUMxNTYuNDA5IDMzLjk5ODggMTU1Ljg5NyAzMy44MDA4IDE1NS41ODQgMzMuODg1NkMxNTUuMjk5IDMzLjk0MjIgMTU0Ljk4NiAzNC4wODM2IDE1NC43NTkgMzQuMjgxNkMxNTQuMTMzIDM0Ljg3NTYgMTUzLjM2NSAzNS40MTMgMTUzLjA4IDM2LjI2MTZDMTUyLjc2NyAzNy4wODE5IDE1Mi40NTUgMzcuOTMwNCAxNTIuNDgzIDM4LjgzNTVDMTUyLjUxMSAzOS42ODQxIDE1Mi41NCA0MC41MzI2IDE1Mi43MzkgNDEuNDA5NUMxNTMuMTA5IDQyLjkzNjkgMTUzLjk2MiA0My45NTUxIDE1NS4yNzEgNDQuNzE4OEMxNTUuMzU2IDQ0Ljc3NTQgMTU1LjQxMyA0NC44ODg1IDE1NS40NyA0NC44ODg1QzE1Ni4wMzkgNDUuMDAxNyAxNTYuNjY1IDQ0LjY2MjMgMTU3LjE3NyA0NS4xNzE0QzE1Ny40MzMgNDQuNzc1NCAxNTcuODg4IDQ1LjA1ODMgMTU4LjIwMSA0NC44ODg1QzE1OS4wODMgNDQuNDA3NyAxNjAuMDIyIDQ0LjA2ODMgMTYwLjU2MiA0My4wMjE3QzE2MS4xMzEgNDEuODYyIDE2MS4zODcgNDAuNzU4OSAxNjEuMzAyIDM5Ljc0MDdaIiBmaWxsPSIjQ0IxNDE5Ij48L3BhdGg+CiAgICA8cGF0aCBkPSJNNjUuMTQ2OCAyMi41Nzc4QzY1LjE0NjggMjIuMDQwNyA2NS4yMDM3IDIxLjQ3NTQgNjUuMTQ2OCAyMC45MzgzQzY1LjA2MTUgMjAuMDkwMiA2NS40MDI5IDE5LjUyNDggNjUuOTQzNCAxOC44NDY0QzY3LjMzNzMgMTcuMTIxOSA2OC44MTY3IDE1LjQ4MjQgNzAuMTI1MyAxMy42NzMxQzcwLjg2NDkgMTIuNjI3MiA3MS44NjA2IDExLjc1MDkgNzIuNzE0MSAxMC43NjE0QzcyLjc5OTQgMTAuNjc2NiA3Mi44NTYzIDEwLjUwNyA3Mi44Mjc5IDEwLjM5NEM3Mi43OTk0IDEwLjMwOTEgNzIuNjI4NyAxMC4yMjQzIDcyLjUxNDkgMTAuMjI0M0M3MC45NTAzIDEwLjIyNDMgNjkuMzg1NiAxMC4yMjQzIDY3LjgyMSAxMC4yMjQzQzY2Ljk2NzUgMTAuMjI0MyA2Ni4xMTQxIDEwLjIyNDMgNjUuMjMyMiAxMC4yMjQzQzY1LjIzMjIgOS4xMjE4NSA2NS4yMzIyIDguMDQ3NjMgNjUuMjMyMiA2Ljk3MzQyQzY2LjM0MTcgNi43NDcyNiA3My4yMjYxIDYuNjkwNzMgNzguMjYxNSA2Ljg2MDM0Qzc4LjI2MTUgNy4zNjkxOCA3OC4xNzYxIDcuODc4MDIgNzguMjg5OSA4LjM1ODU5Qzc4LjQ4OTEgOS4xNzgzOSA3OC4wMDU0IDkuNzQzNzcgNzcuNjY0MSAxMC4zNjU3Qzc3LjM3OTYgMTAuODc0NSA3Ni45NTI5IDExLjI5ODYgNzYuNTgzIDExLjc1MDlDNzUuMTYwNiAxMy40NDcgNzMuNzA5OCAxNS4xNDMxIDcyLjQ1OCAxNi45NTIzQzcxLjkxNzUgMTcuNzE1NiA3MS4yMzQ4IDE4LjM5NDEgNzAuNTIzNSAxOS4yNDIxQzcyLjgyNzkgMTkuMjQyMSA3NC45ODk5IDE5LjI0MjEgNzcuMTUyIDE5LjI0MjFDNzcuNDkzNCAxOS4yNDIxIDc3LjgzNDggMTkuMTg1NiA3OC4xNDc3IDE5LjI3MDRDNzguMjMzIDE5LjI5ODcgNzguNjAyOSAxOS40MTE3IDc4LjQzMjIgMTkuODY0Qzc4LjE3NjEgMjAuNDg1OSA3OC4wOTA4IDIxLjIyMDkgNzguNDg5MSAyMS44NDI5Qzc4Ljc3MzUgMjIuMjk1MiA3OC4zNDY4IDIyLjQzNjUgNzguMTQ3NyAyMi42MDYxQzc4LjAwNTQgMjIuNzE5MiA3Ny42OTI1IDIyLjYzNDQgNzcuNDY0OSAyMi42MzQ0QzczLjQ4MjIgMjIuNjM0NCA2OS40OTk0IDIyLjYzNDQgNjUuNTE2NyAyMi42MzQ0QzY1LjQwMjkgMjIuNjYyNyA2NS4zMTc1IDIyLjYzNDQgNjUuMTQ2OCAyMi41Nzc4WiIgZmlsbD0iYmxhY2siPjwvcGF0aD4KICAgIDxwYXRoIGQ9Ik0xMDQuODAzIDMxLjAzMUMxMDUuNDU3IDMxLjAzMSAxMDYuMTQgMzEuMDU5MyAxMDYuNzk0IDMxLjAzMUMxMDcuMTY0IDMxLjAwMjcgMTA3LjIyMSAzMS4xNzI0IDEwNy4zMDYgMzEuNDgzNUMxMDcuNzMzIDMyLjc1NiAxMDguMjQ1IDM0LjAyODUgMTA4LjY3MiAzNS4zMDExQzEwOC45ODUgMzYuMTc3NyAxMDkuMTU2IDM3LjA4MjYgMTA5LjU1NCAzNy45MDI3QzEwOS45NTIgMzguNzIyOCAxMTAuMTggMzkuNjI3NyAxMTAuNTUgNDAuNDQ3OEMxMTEuMTc1IDQxLjg5IDExMS42MzEgNDMuMzg4OCAxMTIuMDU3IDQ0Ljg4NzVDMTEyLjI4NSA0NS41OTQ1IDExMi42NTUgNDYuMjQ0OSAxMTIuODgyIDQ2Ljk1MTlDMTEzLjA1MyA0Ny41MTc0IDExMi45MzkgNDcuNzE1NCAxMTIuMzcgNDcuNzE1NEMxMTEuNDYgNDcuNzE1NCAxMTAuNTc4IDQ3LjcxNTQgMTA5LjY2OCA0Ny43MTU0QzEwOS4yNjkgNDcuNzE1NCAxMDkuMDcgNDcuNDg5MiAxMDguOTU2IDQ3LjE3ODFDMTA4LjY0NCA0Ni4zMDE1IDEwOC4zMzEgNDUuNDUzMSAxMDguMTMxIDQ0LjU3NjVDMTA4LjA0NiA0NC4xODA2IDEwOC4wMTggNDQuMDM5MiAxMDcuNTYzIDQ0LjAzOTJDMTA1Ljg1NiA0NC4wNjc1IDEwNC4xNDkgNDQuMDY3NSAxMDIuNDQyIDQ0LjAzOTJDMTAxLjkwMSA0NC4wMzkyIDEwMS42NDUgNDQuMjA4OCAxMDEuNTAzIDQ0LjgwMjdDMTAxLjMzMiA0NS43MDc2IDEwMC45NjMgNDYuNTU2IDEwMC42NzggNDcuNDA0M0MxMDAuNjIxIDQ3LjUxNzQgMTAwLjQ1IDQ3LjY4NzEgMTAwLjMzNyA0Ny42ODcxQzk5LjE5ODcgNDcuNzE1NCA5OC4wNjA4IDQ3LjcxNTQgOTYuOTIyOSA0Ny42ODcxQzk2LjU1MyA0Ny42ODcxIDk2LjgzNzUgNDcuNDg5MiA5Ni44OTQ0IDQ3LjM0NzhDOTguMDAzOSA0NC4wOTU3IDk5LjExMzQgNDAuODE1NCAxMDAuMjIzIDM3LjU2MzRDMTAwLjU2NCAzNi41NzM2IDEwMC44NDkgMzUuNTU1NiAxMDEuMzA0IDM0LjYyMjRDMTAxLjc1OSAzMy42NjA5IDEwMS45NTggMzIuNjQyOSAxMDIuMzg1IDMxLjY4MTRDMTAyLjU1NiAzMS4yNTcyIDEwMi41MjcgMzAuOTQ2MiAxMDMuMDk2IDMxLjAwMjdDMTAzLjY2NSAzMS4wODc2IDEwNC4yMzQgMzEuMDMxIDEwNC44MDMgMzEuMDMxWk0xMDQuNzc1IDMzLjg1ODlDMTA0LjQwNSAzNS4zMDExIDEwNC4xNDkgMzYuNDg4OCAxMDMuODM2IDM3LjY3NjVDMTAzLjUyMyAzOC44MzU5IDEwMy4xNTMgMzkuOTk1MyAxMDIuNzU1IDQxLjIxMTNDMTA0LjA2MyA0MS4yMTEzIDEwNS4zNzIgNDEuMjExMyAxMDYuNjgxIDQxLjIxMTNDMTA3LjA1IDQxLjIxMTMgMTA3LjAyMiA0MC45NTY4IDEwNi45MDggNDAuNzMwNkMxMDYuNTk1IDQwLjEzNjcgMTA2LjUxIDM5LjQ4NjMgMTA2LjMzOSAzOC44MzU5QzEwNi4xOTcgMzguMjEzOCAxMDUuOTEzIDM3LjYxOTkgMTA1Ljc3IDM2Ljk2OTVDMTA1LjYgMzYuMDA4MSAxMDUuMjAxIDM1LjA3NDkgMTA0Ljc3NSAzMy44NTg5WiIgZmlsbD0iI0NCMTQxOSI+PC9wYXRoPgogICAgPHBhdGggZD0iTTEzOC40ODYgMjIuNjYyOEMxMzQuMDQ4IDIyLjY2MjggMTI5LjY2NyAyMi42NjI4IDEyNS4xNzIgMjIuNjYyOEMxMjUuMTcyIDIyLjE4MjIgMTI1LjIyOSAyMS43MyAxMjUuMTcyIDIxLjI3NzdDMTI1LjAwMSAyMC4xNDcgMTI1LjU5OSAxOS4zODM4IDEyNi4yODIgMTguNjIwNkMxMjcuNDE5IDE3LjMyMDMgMTI4LjUyOSAxNi4wMiAxMjkuNDk2IDE0LjU3ODNDMTI5Ljk1MSAxMy45NTY1IDEzMC41NzcgMTMuNTA0MiAxMzEuMDMyIDEyLjgyNThDMTMxLjYwMSAxMS45NDk1IDEzMi4zOTggMTEuMjE0NSAxMzMuMTM4IDEwLjMzODNDMTMxLjg4NiAxMC4xNDA0IDEzMC42MzQgMTAuMjgxNyAxMjkuNDExIDEwLjI1MzVDMTI4LjEzMSAxMC4yMjUyIDEyNi44NSAxMC4yNTM1IDEyNS41MTMgMTAuMjUzNUMxMjUuNTEzIDkuMTUxMDMgMTI1LjUxMyA4LjA0ODYxIDEyNS41MTMgNi44ODk2NUMxMjkuNzI0IDYuODg5NjUgMTMzLjkwNiA2Ljg4OTY1IDEzOC4yNTggNi44ODk2NUMxMzguMjU4IDcuNzM3NjcgMTM4LjI1OCA4LjYxMzk1IDEzOC4yNTggOS40OTAyNEMxMzguMjU4IDkuNjU5ODQgMTM4LjIwMSA5Ljg1NzcxIDEzOC4wODggOS45NzA3OEMxMzcuMjM0IDExLjA0NDkgMTM2LjI5NSAxMi4wMzQzIDEzNS40OTkgMTMuMTM2N0MxMzQuMzYxIDE0LjcxOTcgMTMyLjkzOCAxNi4wNDgyIDEzMS45NDMgMTcuNzQ0M0MxMzEuNzE1IDE4LjE0IDEzMS4yODggMTguNDIyNyAxMzAuOTQ3IDE4Ljc2MTlDMTMwLjg2MiAxOC44NDY3IDEzMC44MDUgMTkuMDQ0NiAxMzAuODYyIDE5LjEyOTRDMTMwLjg5IDE5LjIxNDIgMTMxLjA4OSAxOS4yNzA3IDEzMS4yMDMgMTkuMjcwN0MxMzIuNzY4IDE5LjI3MDcgMTM0LjMzMiAxOS4yNzA3IDEzNS44OTcgMTkuMjcwN0MxMzYuNzUgMTkuMjcwNyAxMzcuNTc1IDE5LjI3MDcgMTM4LjQ4NiAxOS4yNzA3QzEzOC40ODYgMjAuNDAxNCAxMzguNDg2IDIxLjQ0NzMgMTM4LjQ4NiAyMi42NjI4WiIgZmlsbD0iYmxhY2siPjwvcGF0aD4KICAgIDxwYXRoIGQ9Ik03Ny4zNTA1IDM1LjgzNzZDNzYuNDY4NiAzNS44Mzc2IDc1LjY0MzYgMzUuODM3NiA3NC43OTAyIDM1LjgzNzZDNzQuMzYzNSAzNS44Mzc2IDc0LjA3OSAzNS42Njc5IDczLjk5MzYgMzUuMjQzNkM3My43NjYgMzQuMjI1MyA3Mi44ODQxIDMzLjgyOTMgNzIuMDg3NiAzMy42MDNDNzEuMTc3MyAzMy4zNDg0IDcwLjE1MzEgMzMuNTE4MiA2OS4yNzEyIDMzLjkxNDJDNjguNTMxNiAzNC4yNTM2IDY4LjEwNDggMzUuMjE1MyA2OC4zMzI0IDM2LjAzNTZDNjguNDc0NyAzNi42Mjk2IDY4Ljk1ODMgMzcuMDI1NiA2OS41NTU3IDM3LjIyMzZDNzAuODkyOCAzNy43MDQ1IDcyLjI1ODMgMzcuOTg3MyA3My42MjM4IDM4LjM1NUM3NC41MzQxIDM4LjU4MTMgNzUuMzg3NiAzOC45NzczIDc2LjEyNzMgMzkuNTE0OEM3Ny4wNjYgNDAuMjUwMiA3Ny42NjM1IDQxLjI5NjggNzcuNjkxOSA0Mi41MTMxQzc3LjcyMDQgNDMuMTkxOSA3Ny42OTE5IDQzLjg0MjUgNzcuNDkyOCA0NC41MjE0Qzc3LjI2NTIgNDUuMjg1MSA3Ni44OTU0IDQ1LjgyMjUgNzYuMzU0OCA0Ni4zMzE2Qzc1Ljk1NjYgNDYuNzI3NyA3NS4wNzQ3IDQ3LjEyMzcgNzQuNTM0MSA0Ny40MDY1QzcyLjgyNzMgNDguMzM5OSA3MC45NzgxIDQ3LjkxNTcgNjkuMjE0MyA0Ny45MTU3QzY4LjM4OTMgNDcuOTE1NyA2Ny40NzkgNDcuNTE5NyA2Ni43MzkzIDQ3LjA2NzFDNjYuMDI4MSA0Ni42NDI4IDY1LjQ4NzYgNDUuOTM1NiA2NC45MTg2IDQ1LjI4NTFDNjQuNDkxOSA0NC44MDQyIDY0LjYzNDIgNDQuMTI1MyA2NC4zNDk3IDQzLjUzMTNDNjQuMDA4MyA0Mi44ODA4IDY0LjMyMTIgNDIuNTY5NiA2NS4xMTc4IDQyLjU2OTZDNjUuODg1OSA0Mi41Njk2IDY2LjYyNTUgNDIuNTQxMyA2Ny4zOTM2IDQyLjU2OTZDNjcuNTA3NCA0Mi41Njk2IDY3LjY3ODEgNDIuNzM5MyA2Ny43MDY2IDQyLjg4MDhDNjguMzA0IDQ0LjU3NzkgNjkuMDE1MiA0NS4xMTU0IDcwLjgwNzQgNDUuMTE1NEM3MS40MzMzIDQ1LjExNTQgNzIuMDg3NiA0NS4yMDAyIDcyLjY4NSA0NC44MzI1QzcyLjgyNzMgNDQuNzQ3NiA3My4wNTQ4IDQ0Ljc0NzYgNzMuMTY4NiA0NC42MzQ1QzczLjc5NDUgNDQuMDY4OCA3NC41MDU3IDQzLjQ0NjUgNzQuMjQ5NyA0Mi41Njk2Qzc0LjAyMjEgNDEuODA1OSA3My4yODI0IDQxLjM4MTYgNzIuNTE0MyA0MS4xNTUzQzcwLjkyMTIgNDAuNzAyOCA2OS4yNzEyIDQwLjMzNSA2Ny43MDY2IDM5Ljc0MUM2Ni45NjY5IDM5LjQ1ODIgNjYuMzY5NSAzOS4wMzM5IDY1LjgwMDUgMzguNTI0OEM2NS4yNiAzOC4wMTU2IDY0LjgzMzMgMzcuNDQ5OSA2NC44OTAyIDM2LjYyOTZDNjQuMjA3NCAzNS43NTI3IDY0Ljg2MTcgMzQuOTA0MiA2NS4wNjA5IDM0LjA4MzlDNjUuMjg4NSAzMy4xNTA0IDY1Ljk0MjggMzIuMzU4NCA2Ni43Njc4IDMxLjg0OTNDNjcuOTA1NyAzMS4xNDIxIDY5LjEyOSAzMC41NzY0IDcwLjU1MTQgMzAuNjYxM0M3MC43NzkgMzAuNjg5NiA3MS4wMzUgMzAuNjg5NiA3MS4yNjI2IDMwLjY2MTNDNzIuODI3MyAzMC40NjMzIDc0LjE2NDMgMzEuMDg1NiA3NS40NzI5IDMxLjc5MjdDNzUuNzU3NCAzMS45MzQxIDc2LjA0MTkgMzIuMTYwNCA3Ni4yMTI2IDMyLjQxNUM3Ni44MSAzMy40ODk5IDc3LjU3ODEgMzQuNDUxNiA3Ny4zNTA1IDM1LjgzNzZaIiBmaWxsPSIjQ0IxNDE5Ij48L3BhdGg+CiAgICA8cGF0aCBkPSJNOTkuMDU2NiAxLjE2Njk5QzEwMC40NzkgMS4xNjY5OSAxMDEuODE2IDEuMTY2OTkgMTAzLjIzOSAxLjE2Njk5QzEwMy4yMzkgOC4zMDM4MiAxMDMuMjM5IDE1LjQxMjMgMTAzLjIzOSAyMi42MDU4QzEwMi42MTMgMjIuNzE5MSAxMDEuOTAxIDIyLjYzNDEgMTAxLjIxOSAyMi42NjI1QzEwMC41MDcgMjIuNjkwOCA5OS44MjQ3IDIyLjY2MjUgOTkuMDU2NiAyMi42NjI1Qzk5LjA1NjYgMTUuNTI1NiA5OS4wNTY2IDguNDE3MTEgOTkuMDU2NiAxLjE2Njk5WiIgZmlsbD0iYmxhY2siPjwvcGF0aD4KICAgIDxwYXRoIGQ9Ik0xNDIuMDE0IDYuODYxMzNDMTQzLjQwOCA2Ljg2MTMzIDE0NC43NzMgNi44NjEzMyAxNDYuMTY3IDYuODYxMzNDMTQ2LjE2NyAxMi4xMTkgMTQ2LjE2NyAxNy4zNzY3IDE0Ni4xNjcgMjIuNjM0NUMxNDQuNzczIDIyLjYzNDUgMTQzLjQwOCAyMi42MzQ1IDE0Mi4wMTQgMjIuNjM0NUMxNDIuMDE0IDE3LjM3NjcgMTQyLjAxNCAxMi4xNDczIDE0Mi4wMTQgNi44NjEzM1oiIGZpbGw9ImJsYWNrIj48L3BhdGg+CiAgICA8cGF0aCBkPSJNNjUuNzcyNSAwLjI4MzIwMUM2Ny4xMDk1IDAuMjgzMjAxIDY4LjE5MDYgMC4yNTU0ODMgNjkuMjcxNiAwLjI4MzIwMUM2OS40NzA3IDAuMjgzMjAxIDY5LjY5ODMgMC40MjE3OTIgNjkuODY5IDAuNTYwMzgyQzcwLjQzOCAxLjA4NzAyIDcxLjAwNjkgMS41ODU5NSA3MS40MDUyIDIuMjc4OUM3MS41NDc1IDIuNTI4MzYgNzEuNzc1IDIuNjM5MjQgNzIuMDMxMSAyLjI1MTE4QzcyLjQyOTQgMS42MTM2NyA3Mi45MTMgMS4wMzE1OSA3My4zOTY2IDAuNDQ5NTFDNzMuNDgxOSAwLjMzODYzNyA3My43MDk1IDAuMjU1NDgzIDczLjg4MDIgMC4yNTU0ODNDNzQuODc1OSAwLjIyNzc2NSA3NS44NzE2IDAuMjU1NDgzIDc2Ljg2NzMgMC4yNTU0ODNDNzYuOTUyNiAwLjI1NTQ4MyA3Ny4wOTQ5IDAuMzEwOTE5IDc3LjA5NDkgMC4zNjYzNTVDNzcuMTIzMyAwLjQ3NzIyOCA3Ny4xMjMzIDAuNjcxMjU0IDc3LjA2NjQgMC43MjY2OUM3NS44NzE2IDEuOTE4NTcgNzQuNjc2OCAzLjA4MjczIDczLjQ1MzUgNC4yNDY4OEM3My4zMTEzIDQuMzg1NDcgNzMuMDI2OCA0LjM4NTQ3IDcyLjgyNzYgNC4zODU0N0M3Mi4xMTY0IDQuNDEzMTkgNzEuNDA1MiA0LjM4NTQ3IDcwLjY5NCA0LjM4NTQ3QzcwLjEyNSA0LjM4NTQ3IDcwLjEyNSA0LjM1Nzc2IDY5LjU1NjEgMy43NzU2OEM2OC43MzExIDIuOTQ0MTMgNjcuODc3NiAyLjE0MDMxIDY3LjA1MjYgMS4zMzY0OUM2Ni42NTQ0IDEuMDU5MzEgNjYuMzEzIDAuNzgyMTI2IDY1Ljc3MjUgMC4yODMyMDFaIiBmaWxsPSJibGFjayI+PC9wYXRoPgogICAgPHBhdGggZD0iTTE0NC4wMzQgNC40NDUzM0MxNDMuMzggNC44MzM4OSAxNDIuODk2IDQuMTQwMDMgMTQyLjQ0MSAzLjgzNDc0QzE0Mi4wNDMgMy41ODQ5NSAxNDEuOTI5IDIuODYzMzQgMTQxLjcwMSAyLjM2Mzc2QzE0Mi4yNyAyLjAwMjk1IDE0MS43MyAxLjI1MzU5IDE0Mi4zNTUgMC44MDk1MThDMTQyLjgzOSAwLjQ3NjQ2NiAxNDMuMjY2IDAuMzM3Njk1IDE0My44MzUgMC4yMjY2NzhDMTQ0Ljk3MyAwLjAzMjM5NzQgMTQ2LjMxIDAuNjE1MjM4IDE0Ni41MDkgMS41ODY2NEMxNDYuODc5IDMuMzkwNjcgMTQ1LjcxMiA0LjQ3MzA5IDE0NC4wMzQgNC40NDUzM1oiIGZpbGw9ImJsYWNrIj48L3BhdGg+CiAgPC9nPgogIDxkZWZzPgogICAgPGNsaXBQYXRoIGlkPSJjbGlwMF8xODAwMV8xNTY5MSI+CiAgICAgIDxyZWN0IHdpZHRoPSIxNjUiIGhlaWdodD0iNDgiIGZpbGw9IndoaXRlIj48L3JlY3Q+CiAgICA8L2NsaXBQYXRoPgogIDwvZGVmcz4KPC9zdmc+Cg==" style="height:36px;width:auto;filter:brightness(0) invert(1);opacity:.92;flex-shrink:0">
  </div>

  <div class="info-box">
    <div class="info-row"><span class="info-label">Klient</span><span class="info-val">${klient}</span></div>
    <div class="info-row"><span class="info-label">Číslo zakázky</span><span class="info-val">${zakazka}</span></div>
    <div class="info-row"><span class="info-label">Adresa</span><span class="info-val">${adresa}</span></div>
    <div class="info-row"><span class="info-label">Technik</span><span class="info-val">${technik}</span></div>
  </div>

  <div class="section-title">Přehled položek</div>
  <table>
    <thead><tr>
      <th style="width:40px">#</th>
      <th>Místnost / pozice</th>
      <th>Typ produktu</th>
      <th style="width:70px">Počet</th>
    </tr></thead>
    <tbody>${tableRows}</tbody>
  </table>

  <div class="footer">
    <div class="footer-left">
      Tato nabídka je nezávazná a slouží jako přehled<br>
      záměřených produktů. Rozměry a ceny budou<br>
      upřesněny po potvrzení objednávky.
    </div>
    <div class="footer-right">
      <div class="footer-brand">ŽaluzieSnadno.cz</div>
      info@zaluziesnadno.cz<br>
      266 266 792<br>
      Po–Pá, 8:00–16:00
    </div>
  </div>

  <div class="sign-box">
    <div class="sign-line">Podpis technika</div>
  </div>
</div>

</body></html>`;

  const win=window.open('','_blank');
  if(win){win.document.write(html);win.document.close();}
});

// ===================== INIT =====================
document.querySelectorAll('#klient,#zakazka,#adresa,#datum,#technik').forEach(function(el){el.addEventListener('input',saveState);});
document.getElementById('datum').value=new Date().toISOString().split('T')[0];
loadState();
if(activTyp){var sel=document.getElementById('typ-select');if(sel)sel.value=activTyp;renderForm(activTyp);}
renderList();
var btnReset=document.getElementById('btn-reset');
if(btnReset)btnReset.addEventListener('click',resetState);
