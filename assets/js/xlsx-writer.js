"use strict";
/* =====================================================================
   Gerador mínimo de ficheiros Excel (.xlsx) no navegador, sem bibliotecas externas.
   Suporta: várias folhas, estilos fixos, larguras de coluna, células unidas,
   painéis fixos, filtros automáticos, alturas de linha e gráficos de linhas nativos do Excel.
   ===================================================================== */
const XLSX_W = (()=>{
  const enc = new TextEncoder();
  const xe = s => String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,"");
  // ---------- ZIP (sem compressão) ----------
  const CRC = (()=>{ const t=new Uint32Array(256); for(let n=0;n<256;n++){ let c=n; for(let k=0;k<8;k++) c = c&1 ? 0xEDB88320^(c>>>1) : c>>>1; t[n]=c>>>0; } return t; })();
  const crc32 = u8 => { let c=0xFFFFFFFF; for(let i=0;i<u8.length;i++) c=CRC[(c^u8[i])&0xFF]^(c>>>8); return (c^0xFFFFFFFF)>>>0; };
  async function deflate(u8){
    if(typeof CompressionStream==="undefined") return null;
    try{ const cs=new CompressionStream("deflate-raw"); const out=new Response(new Blob([u8]).stream().pipeThrough(cs)).arrayBuffer(); return new Uint8Array(await out); }catch(e){ return null; } }
  async function zip(files){ // files: [{name, data:string|Uint8Array}]
    const now=new Date(), dosT=(now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1), dosD=((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate();
    const parts=[], central=[]; let off=0;
    for(const f of files){
      const name=enc.encode(f.name), raw=typeof f.data==="string"?enc.encode(f.data):f.data, crc=crc32(raw);
      const z=raw.length>512?await deflate(raw):null, method=z?8:0, data=z||raw;
      const h=new DataView(new ArrayBuffer(30));
      h.setUint32(0,0x04034b50,true); h.setUint16(4,20,true); h.setUint16(6,0x0800,true); h.setUint16(8,method,true); h.setUint16(10,dosT,true); h.setUint16(12,dosD,true);
      h.setUint32(14,crc,true); h.setUint32(18,data.length,true); h.setUint32(22,raw.length,true); h.setUint16(26,name.length,true); h.setUint16(28,0,true);
      parts.push(new Uint8Array(h.buffer),name,data);
      const c=new DataView(new ArrayBuffer(46));
      c.setUint32(0,0x02014b50,true); c.setUint16(4,20,true); c.setUint16(6,20,true); c.setUint16(8,0x0800,true); c.setUint16(10,method,true); c.setUint16(12,dosT,true); c.setUint16(14,dosD,true);
      c.setUint32(16,crc,true); c.setUint32(20,data.length,true); c.setUint32(24,raw.length,true); c.setUint16(28,name.length,true); c.setUint32(42,off,true);
      central.push(new Uint8Array(c.buffer),name);
      off += 30+name.length+data.length;
    }
    const cdSize=central.reduce((s,p)=>s+p.length,0);
    const e=new DataView(new ArrayBuffer(22)); e.setUint32(0,0x06054b50,true); e.setUint16(8,files.length,true); e.setUint16(10,files.length,true); e.setUint32(12,cdSize,true); e.setUint32(16,off,true);
    return new Blob([...parts,...central,new Uint8Array(e.buffer)],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }
  // ---------- estilos ----------
  // índices de estilo usados pelas folhas (ver cellXfs abaixo)
  const ST = { text:0, head:1, num:2, int:3, pct:4, pctUp:5, pctDown:6, title:7, note:8, date:9, cell:10, group:11, bold:12, kpiL:13, kpiV:14, numLow:15, numHigh:16, band:17, kpiN:18, wrap:19, subtitle:20, dateTxt:21, boldWrap:22 };
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="4"><numFmt numFmtId="164" formatCode="#,##0.00"/><numFmt numFmtId="165" formatCode="+0.00%;-0.00%;0.00%"/><numFmt numFmtId="166" formatCode="dd/mm/yyyy"/><numFmt numFmtId="167" formatCode="#,##0"/></numFmts>
<fonts count="9">
<font><sz val="11"/><color rgb="FF13202B"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FF13202B"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="16"/><color rgb="FF13202B"/><name val="Calibri"/><family val="2"/></font>
<font><i/><sz val="10"/><color rgb="FF5B6976"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FF0B8A4E"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FFC42F3A"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="14"/><color rgb="FF13202B"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="12"/><color rgb="FFB86E00"/><name val="Calibri"/><family val="2"/></font>
</fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F2D3C"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEEF1F4"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF6E7CF"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFC9D1D8"/></left><right style="thin"><color rgb="FFC9D1D8"/></right><top style="thin"><color rgb="FFC9D1D8"/></top><bottom style="thin"><color rgb="FFC9D1D8"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="23">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="5" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="165" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="167" fontId="7" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>
<xf numFmtId="164" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="164" fontId="5" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="164" fontId="7" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="0" fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="166" fontId="7" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  // ---------- utilitários ----------
  const colL = n => { let s=""; n++; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; };
  const ref = (r,c) => colL(c)+(r+1);
  const sheetRef = name => "'"+String(name).replace(/'/g,"''")+"'";
  const safeName = (n,used) => { let s=String(n).replace(/[\[\]:*?\/\\]/g," ").slice(0,31).trim()||"Folha"; let b=s, i=2; while(used.has(s.toLowerCase())){ s=(b.slice(0,28)+" "+i++); } used.add(s.toLowerCase()); return s; };

  /* Folha: construída linha a linha.
     cell: valor simples (número/texto) ou {v, s, t:"d"(data serial)} ; null = vazio */
  class Sheet {
    constructor(name){ this.name=name; this.rows=[]; this.widths={}; this.merges=[]; this.heights={}; this.freeze=null; this.filter=null; this.charts=[]; }
    set(r,c,cell){ (this.rows[r] ??= [])[c]=cell; }
    row(r,cells,c0=0){ cells.forEach((v,i)=>this.set(r,c0+i,v)); }
    merge(r1,c1,r2,c2){ this.merges.push(ref(r1,c1)+":"+ref(r2,c2)); }
    width(c,w){ this.widths[c]=w; }
    height(r,h){ this.heights[r]=h; }
    // gráfico de linhas: {title, cat:{sheet,col,r1,r2}, series:[{name, sheet, col, r1, r2, color}], at:{c1,r1,c2,r2}, yFmt}
    lineChart(o){ this.charts.push(o); }
    autoWidths(min=8,max=60){ const w={}; this.rows.forEach(row=>row&&row.forEach((cell,c)=>{ if(cell==null) return; const v=typeof cell==="object"?cell.v:cell; if(v==null) return;
      const len=(typeof v==="number")? (cell&&cell.t==="d"?11:String(Math.round(v*100)/100).length+3) : String(v).split("\n").reduce((m,l)=>Math.max(m,l.length),0)+2;
      if(cell && typeof cell==="object" && (cell.s===ST.title||cell.s===ST.note||cell.s===ST.wrap||cell.s===ST.band||cell.s===ST.subtitle||cell.noW)) return;
      w[c]=Math.max(w[c]||min,Math.min(max,len)); })); for(const c in w) if(this.widths[c]==null) this.widths[c]=w[c]; }
    xml(drawingRid){
      const out=[]; let maxC=0;
      this.rows.forEach((row,r)=>{ if(!row) return; const cs=[];
        row.forEach((cell,c)=>{ if(cell==null) return; maxC=Math.max(maxC,c);
          let v=cell, s=0, t=null; if(typeof cell==="object"){ v=cell.v; s=cell.s||0; t=cell.t||null; }
          const R=ref(r,c);
          if(v==null||v===""){ cs.push(`<c r="${R}" s="${s}"/>`); return; }
          if(typeof v==="number" && isFinite(v)) cs.push(`<c r="${R}" s="${s}"><v>${v}</v></c>`);
          else cs.push(`<c r="${R}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`); });
        const h=this.heights[r]; out.push(`<row r="${r+1}"${h?` ht="${h}" customHeight="1"`:""}>${cs.join("")}</row>`); });
      const cols=Object.keys(this.widths).map(Number).sort((a,b)=>a-b).map(c=>`<col min="${c+1}" max="${c+1}" width="${this.widths[c]}" customWidth="1"/>`).join("");
      const pane=this.freeze?`<pane ${this.freeze.c?`xSplit="${this.freeze.c}" `:""}${this.freeze.r?`ySplit="${this.freeze.r}" `:""}topLeftCell="${ref(this.freeze.r||0,this.freeze.c||0)}" activePane="${this.freeze.r&&this.freeze.c?"bottomRight":this.freeze.r?"bottomLeft":"topRight"}" state="frozen"/>`:"";
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetViews><sheetView workbookViewId="0" showGridLines="${this.grid===false?0:1}">${pane}</sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>${cols?`<cols>${cols}</cols>`:""}
<sheetData>${out.join("")}</sheetData>${this.filter?`<autoFilter ref="${this.filter}"/>`:""}${this.merges.length?`<mergeCells count="${this.merges.length}">${this.merges.map(m=>`<mergeCell ref="${m}"/>`).join("")}</mergeCells>`:""}
<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>${drawingRid?`<drawing r:id="${drawingRid}"/>`:""}</worksheet>`;
    }
  }
  function chartXml(o){
    const rng=(sh,col,r1,r2)=>`${sheetRef(sh)}!$${colL(col)}$${r1+1}:$${colL(col)}$${r2+1}`;
    const PAL=["B86E00","1F6FB2","0B8A4E","C42F3A","6B4FA0","2A9D8F","8A6D3B","D35400"];
    const ser=o.series.map((s,i)=>`<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${xe(s.name)}</c:v></c:tx>
<c:spPr><a:ln w="${s.width||22225}" cap="rnd"><a:solidFill><a:srgbClr val="${s.color||PAL[i%PAL.length]}"/></a:solidFill>${s.dash?`<a:prstDash val="${s.dash}"/>`:""}<a:round/></a:ln></c:spPr>
<c:marker><c:symbol val="none"/></c:marker>
<c:cat><c:numRef><c:f>${rng(o.cat.sheet,o.cat.col,o.cat.r1,o.cat.r2)}</c:f></c:numRef></c:cat>
<c:val><c:numRef><c:f>${rng(s.sheet,s.col,s.r1,s.r2)}</c:f></c:numRef></c:val><c:smooth val="0"/></c:ser>`).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:roundedCorners val="0"/>
<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="pt-PT" sz="1200" b="1"/><a:t>${xe(o.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>
<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/><c:axId val="5001"/><c:axId val="5002"/></c:lineChart>
<c:dateAx><c:axId val="5001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="dd/mm/yy" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>
<c:txPr><a:bodyPr rot="-2700000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr><a:endParaRPr lang="pt-PT"/></a:p></c:txPr><c:crossAx val="5002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblOffset val="100"/><c:baseTimeUnit val="days"/></c:dateAx>
<c:valAx><c:axId val="5002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E1E6EA"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
<c:numFmt formatCode="${xe(o.yFmt||"#,##0.00")}" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr><a:endParaRPr lang="pt-PT"/></a:p></c:txPr><c:crossAx val="5001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>
</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="span"/></c:chart>
<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="D3DAE0"/></a:solidFill></a:ln></c:spPr></c:chartSpace>`;
  }
  function drawingXml(charts, firstChart){
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
${charts.map((o,i)=>`<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${o.at.c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${o.at.r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${o.at.c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${o.at.r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i+2}" name="${xe("Gráfico "+(i+1))}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId${i+1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`).join("")}
</xdr:wsDr>`;
  }
  class Workbook {
    constructor(){ this.sheets=[]; this.used=new Set(); this.title=""; this.creator=""; }
    sheet(name){ const s=new Sheet(safeName(name,this.used)); this.sheets.push(s); return s; }
    async blob(){
      const files=[], ct=[`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`,
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
        `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
        `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`];
      const wbRels=[]; let chartN=0, drawN=0;
      this.sheets.forEach((s,i)=>{
        const n=i+1; let drawingRid=null;
        if(s.charts.length){ drawN++; drawingRid="rId1";
          files.push({name:`xl/worksheets/_rels/sheet${n}.xml.rels`,data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawN}.xml"/></Relationships>`});
          const rels=[]; s.charts.forEach((o,k)=>{ chartN++; files.push({name:`xl/charts/chart${chartN}.xml`,data:chartXml(o)});
            ct.push(`<Override PartName="/xl/charts/chart${chartN}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
            rels.push(`<Relationship Id="rId${k+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartN}.xml"/>`); });
          files.push({name:`xl/drawings/drawing${drawN}.xml`,data:drawingXml(s.charts)});
          files.push({name:`xl/drawings/_rels/drawing${drawN}.xml.rels`,data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`});
          ct.push(`<Override PartName="/xl/drawings/drawing${drawN}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`); }
        files.push({name:`xl/worksheets/sheet${n}.xml`,data:s.xml(drawingRid)});
        ct.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
        wbRels.push(`<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`);
      });
      const ns=this.sheets.length;
      wbRels.push(`<Relationship Id="rId${ns+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
      const defs=this.sheets.map((s,i)=>s.filter?`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${xe(sheetRef(s.name))}!${s.filter.replace(/([A-Z]+)(\d+)/g,"$$$1$$$2")}</definedName>`:"").join("");
      files.unshift(
        {name:"[Content_Types].xml",data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${ct.join("")}</Types>`},
        {name:"_rels/.rels",data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`},
        {name:"docProps/core.xml",data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xe(this.title)}</dc:title><dc:creator>${xe(this.creator)}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0,19)}Z</dcterms:created></cp:coreProperties>`},
        {name:"xl/workbook.xml",data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${this.sheets.map((s,i)=>`<sheet name="${xe(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")}</sheets>${defs?`<definedNames>${defs}</definedNames>`:""}</workbook>`},
        {name:"xl/_rels/workbook.xml.rels",data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels.join("")}</Relationships>`},
        {name:"xl/styles.xml",data:STYLES});
      return await zip(files);
    }
  }
  return { Workbook, ST, colL, ref };
})();
