"use strict";
/* =====================================================================
   Exportação para Excel: (1) dados brutos  (2) relatório formatado com tabelas e gráficos
   Depende de: xlsx-writer.js, app.js, bulletin.js
   ===================================================================== */
const XL_T = {
  pt:{ btn:"Exportar Excel", raw:"Dados brutos", rawD:"Todos os registos de preço, um por linha, prontos para filtrar e analisar.", fmt:"Relatório formatado",
    fmtD:"As tabelas e os gráficos do sistema: cotações, séries, gráficos, preço por mercado, registos recentes e boletim.", working:"A preparar o ficheiro Excel…", done:"Ficheiro Excel pronto", fail:"Não foi possível gerar o Excel",
    sData:"Dados", sReadme:"Leia-me", sSummary:"Resumo", sCharts:"Gráficos", sSeries:"Séries", sMkt:"Preço por mercado", sRecent:"Registos recentes", sBull:"Boletim",
    cDate:"Data", cYear:"Ano", cMonth:"Mês", cWeek:"Semana (ISO)", cProv:"Província", cDist:"Distrito", cMkt:"Mercado", cSel:"Tipo de vendedor", cCat:"Categoria", cCode:"Código",
    cProd:"Produto", cPrice:"Preço", cUnit:"Unidade", cOrig:"Proveniência", cQty:"Quantidade (kg)", cF:"Vendedoras (mulheres)", cM:"Vendedores (homens)",
    cLast:"Último", cChg:"Var. %", cOpen:"Abertura", cLow:"Mínimo", cHigh:"Máximo", cAvg:"Média", cRec:"Registos", cMkts:"Mercados", cLastD:"Último registo", cIdx:"Índice SIMA", cVsAvg:"vs média",
    cPrev:"Semana anterior", cCur:"Semana recente", cMaxM:"Mercados com o máximo", cMinM:"Mercados com o mínimo", cZone:"Zona", cMonthPrev:"Mês anterior", cMonthCur:"Mês actual",
    title:"SIMA - Sistema de Monitoria de Preços de Produtos Agrícolas", repTitle:"Relatório de preços", exported:"Exportado em", source:"Fonte", filters:"Filtros", period:"Período",
    demo:"Dados de demonstração (fictícios)", all:"Todos", allF:"Todas", quotes:"Cotações", chartsNote:"Gráficos de linhas do Excel, ligados à folha «Séries»: pode editá-los, mudar cores ou copiá-los.",
    seriesNote:"Fecho de cada intervalo (média do último dia com registos), em MZN/kg. Células vazias = sem registos nesse intervalo.",
    mktNote:"Preço médio (MZN/kg) de cada produto em cada mercado no período filtrado. A vermelho o mínimo e a verde o máximo de cada produto.",
    rawNote:"Uma linha por produto registado em cada visita ao mercado. Preço = média dos preco_kg1-3 preenchidos (ou preco_medio; ou média de Preço-1/2/3 por unidade).",
    readme:[["Data","Data de monitoria (hora de Maputo)."],["Ano / Mês / Semana (ISO)","Período a que o registo pertence; a semana começa à segunda-feira."],
      ["Província / Distrito / Mercado","Local da monitoria (camada de registos do inquérito)."],["Tipo de vendedor","Nível de comercialização: Produtor, Grossista ou Retalhista."],
      ["Categoria / Código / Produto","Produto registado; o código é o usado nas cotações do sistema."],["Preço","Média dos preço por kg das amostras (preco_kg1-3). Se não houver, preco_medio; se também não houver, média de Preço-1/2/3 (por unidade)."],
      ["Unidade","MZN/kg ou MZN por unidade."],["Proveniência","Origem do produto indicada pelos vendedores."],["Quantidade (kg)","Quantidade disponível declarada nas amostras, convertida para kg quando possível."],
      ["Vendedoras / Vendedores","Número de amostras recolhidas a mulheres e a homens."]],
    bullT:"Boletim semanal", part:"Parte", dyn:"Dinâmica de preços", weekly:"Preços por mercado e variação", flows:"Fluxos (proveniência)", qty:"Quantidades no mercado grossista (kg)", live:"Pecuária: médias mensais (MZN/kg)",
    level:"Nível de mercado", allLevels:"Todos os níveis", week:"Semana", records:"registos" },
  en:{ btn:"Export Excel", raw:"Raw data", rawD:"Every price record, one per row, ready to filter and analyse.", fmt:"Formatted report",
    fmtD:"The system's tables and charts: quotes, series, charts, price by market, recent records and bulletin.", working:"Preparing the Excel file…", done:"Excel file ready", fail:"Could not build the Excel file",
    sData:"Data", sReadme:"Read me", sSummary:"Summary", sCharts:"Charts", sSeries:"Series", sMkt:"Price by market", sRecent:"Recent records", sBull:"Bulletin",
    cDate:"Date", cYear:"Year", cMonth:"Month", cWeek:"Week (ISO)", cProv:"Province", cDist:"District", cMkt:"Market", cSel:"Seller type", cCat:"Category", cCode:"Code",
    cProd:"Product", cPrice:"Price", cUnit:"Unit", cOrig:"Origin", cQty:"Quantity (kg)", cF:"Female sellers", cM:"Male sellers",
    cLast:"Last", cChg:"Chg. %", cOpen:"Open", cLow:"Low", cHigh:"High", cAvg:"Average", cRec:"Records", cMkts:"Markets", cLastD:"Latest record", cIdx:"SIMA index", cVsAvg:"vs avg",
    cPrev:"Previous week", cCur:"This week", cMaxM:"Markets at the high", cMinM:"Markets at the low", cZone:"Region", cMonthPrev:"Previous month", cMonthCur:"Current month",
    title:"SIMA - Agricultural Commodity Price Monitoring System", repTitle:"Price report", exported:"Exported", source:"Source", filters:"Filters", period:"Period",
    demo:"Demo data (fictitious)", all:"All", allF:"All", quotes:"Quotes", chartsNote:"Excel line charts linked to the «Series» sheet: you can edit them, change colours or copy them.",
    seriesNote:"Close of each interval (average of the last day with records), in MZN/kg. Empty cells = no records in that interval.",
    mktNote:"Average price (MZN/kg) of each product in each market over the filtered period. Red = lowest and green = highest for each product.",
    rawNote:"One row per product recorded at each market visit. Price = average of the filled preco_kg1-3 (or preco_medio; or average of Price-1/2/3 per unit).",
    readme:[["Date","Monitoring date (Maputo time)."],["Year / Month / Week (ISO)","Period of the record; weeks start on Monday."],
      ["Province / District / Market","Where the monitoring took place (survey records layer)."],["Seller type","Trading level: Producer, Wholesaler or Retailer."],
      ["Category / Code / Product","Product recorded; the code is the one used in the system's quotes."],["Price","Average per-kg price of the samples (preco_kg1-3). Otherwise preco_medio; otherwise the average of Price-1/2/3 (per unit)."],
      ["Unit","MZN/kg or MZN per unit."],["Origin","Product origin reported by the sellers."],["Quantity (kg)","Quantity available reported in the samples, converted to kg when possible."],
      ["Female / Male sellers","Number of samples collected from women and from men."]],
    bullT:"Weekly bulletin", part:"Part", dyn:"Price dynamics", weekly:"Prices by market and change", flows:"Trade flows (origin)", qty:"Wholesale quantities (kg)", live:"Livestock: monthly averages (MZN/kg)",
    level:"Market level", allLevels:"All levels", week:"Week", records:"records" }
};
const xt = k => (XL_T[LANG]||XL_T.pt)[k];
const XS = XLSX_W.ST;
const xSerial = d => d + 25569;                 // dia (desde 1970) -> data do Excel
const xNum = v => (v==null||!isFinite(v)) ? null : Math.round(v*100)/100;
const xPct = v => (v==null||!isFinite(v)) ? null : {v:Math.round(v*100)/10000, s: Math.abs(v)<0.005?XS.pct : v>0?XS.pctUp:XS.pctDown};
const xDate = d => d==null ? null : {v:xSerial(d), s:XS.date, t:"d"};

function filtersLine(){
  const L=[[xt("cProv"),S.prov||xt("allF")],[xt("cMkt"),S.mkt||xt("all")],[xt("cSel"),S.sel?sellerName(S.sel):xt("all")],[xt("cCat"),S.cat?catName(S.cat):xt("allF")],
    [xt("period"), S.period?`${S.period} ${LANG==="pt"?"dias":"days"}`:(LANG==="pt"?"Todo o histórico":"Full history")],[t("agg"), t("agg"+S.agg)]];
  return L.map(([k,v])=>`${k}: ${v}`).join(" · ");
}
async function saveFile(fname, blob){
  let dl=null; try{ if(window.claude?.use) dl=await window.claude.use("downloads"); }catch(e){}
  if(dl){ await dl.save({filename:fname,data:blob}); return; }
  const url=URL.createObjectURL(blob), a=document.createElement("a"); a.href=url; a.download=fname; document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); },2000);
}
function header(sh, r, labels, c0=0){ sh.row(r, labels.map(l=>({v:l,s:XS.head})), c0); sh.height(r,30); }
function band(sh, r, text, cols){ sh.set(r,0,{v:text,s:XS.band}); for(let c=1;c<cols;c++) sh.set(r,c,{v:"",s:XS.band}); sh.height(r,22); }

/* ---------------- (1) Dados brutos ---------------- */
function buildRaw(){
  const wb=new XLSX_W.Workbook(); wb.title=xt("title"); wb.creator="SIMA";
  const sh=wb.sheet(xt("sData"));
  const cols=[xt("cDate"),xt("cYear"),xt("cMonth"),xt("cWeek"),xt("cProv"),xt("cDist"),xt("cMkt"),xt("cSel"),xt("cCat"),xt("cCode"),xt("cProd"),xt("cPrice"),xt("cUnit"),xt("cOrig"),xt("cQty"),xt("cF"),xt("cM")];
  header(sh,0,cols);
  const obs=S.obs.slice().sort((a,b)=>a.d-b.d||a.prov.localeCompare(b.prov)||a.mkt.localeCompare(b.mkt)||a.prod.localeCompare(b.prod));
  obs.forEach((o,i)=>{ const dt=new Date(o.d*DAY), wk=isoWeek(o.d);
    sh.row(i+1,[xDate(o.d), dt.getUTCFullYear(), dt.getUTCMonth()+1, `${wk.y}-S${String(wk.w).padStart(2,"0")}`, o.prov, o.dist||"", o.mkt, sellerName(o.sel), catName(o.cat), pTk(o.prod), pName(o.prod),
      {v:xNum(o.price),s:XS.num}, o.unit==="un"?(LANG==="pt"?"MZN/unidade":"MZN/unit"):"MZN/kg", (o.orig||[]).join("; "), o.qty?{v:Math.round(o.qty*100)/100,s:XS.num}:null, o.sf||0, o.sm||0]); });
  sh.freeze={r:1,c:0}; sh.filter=`A1:${XLSX_W.colL(cols.length-1)}${obs.length+1}`;
  [11,7,6,11,16,14,30,14,16,8,24,10,11,30,13,10,10].forEach((w,c)=>sh.width(c,w));
  // leia-me
  const rm=wb.sheet(xt("sReadme")); rm.grid=false;
  rm.set(0,0,{v:xt("title"),s:XS.title}); rm.height(0,24);
  rm.set(1,0,{v:`${xt("exported")}: ${new Date().toLocaleString(LANG==="pt"?"pt-PT":"en-GB",{timeZone:"Africa/Maputo"})} (Maputo)`,s:XS.note});
  rm.set(2,0,{v:`${xt("source")}: ${S.mode==="demo"?xt("demo"):S.url}`,s:XS.note});
  rm.set(3,0,{v:`${fmtN(obs.length,0)} ${xt("records")}`,s:XS.note});
  rm.set(4,0,{v:xt("rawNote"),s:XS.wrap}); rm.merge(4,0,4,1); rm.height(4,34);
  header(rm,6,[LANG==="pt"?"Coluna":"Column",LANG==="pt"?"Descrição":"Description"]);
  xt("readme").forEach(([k,v],i)=>{ rm.row(7+i,[{v:k,s:XS.bold},{v,s:XS.cell}]); });
  rm.width(0,30); rm.width(1,100);
  return wb;
}

/* ---------------- (2) Relatório formatado ---------------- */
function buildFormatted(){
  const wb=new XLSX_W.Workbook(); wb.title=xt("title")+" - "+xt("repTitle"); wb.creator="SIMA";
  const list=filtered(false), sums=summarize(list).sort((a,b)=>(a.cat||"").localeCompare(b.cat||"")||a.tk.localeCompare(b.tk));
  const idx=indexSeries(sums);
  const dMin=list.reduce((m,o)=>Math.min(m,o.d),Infinity), dMax=list.reduce((m,o)=>Math.max(m,o.d),0);
  const unitOf=p=>PROD_UNIT[p]==="un"?(LANG==="pt"?"MZN/unidade":"MZN/unit"):"MZN/kg";

  // --- Resumo ---
  const R=wb.sheet(xt("sSummary")); R.grid=false;
  R.set(0,0,{v:xt("title"),s:XS.title}); R.height(0,26);
  R.set(1,0,{v:xt("repTitle"),s:XS.subtitle});
  R.set(2,0,{v:`${xt("exported")}: ${new Date().toLocaleString(LANG==="pt"?"pt-PT":"en-GB",{timeZone:"Africa/Maputo"})} (Maputo) · ${xt("source")}: ${S.mode==="demo"?xt("demo"):"ArcGIS Survey123"}`,s:XS.note});
  R.set(3,0,{v:`${xt("filters")}: ${filtersLine()}`,s:XS.note});
  const kp=[[xt("cRec"),{v:list.length,s:XS.kpiV}],[t("rptProducts"),{v:sums.length,s:XS.kpiV}],[xt("cMkts"),{v:new Set(list.map(o=>o.mkt)).size,s:XS.kpiV}],
    [xt("cProv"),{v:new Set(list.map(o=>o.prov)).size,s:XS.kpiV}],[LANG==="pt"?"Primeiro registo":"First record",isFinite(dMin)?{v:xSerial(dMin),s:XS.dateTxt}:"—"],
    [xt("cLastD"),dMax?{v:xSerial(dMax),s:XS.dateTxt}:"—"],[xt("cIdx"),idx.length?{v:xNum(idx[idx.length-1].c),s:XS.kpiN}:"—"]];
  kp.forEach(([k,v],i)=>{ R.set(5,i*2,{v:k,s:XS.kpiL}); R.set(5,i*2+1,{v:"",s:XS.kpiL}); R.merge(5,i*2,5,i*2+1); R.set(6,i*2,v); R.set(6,i*2+1,{v:"",s:XS.kpiL}); R.merge(6,i*2,6,i*2+1); });
  R.height(6,24);
  band(R,8,xt("quotes").toUpperCase(),13);
  const qh=[xt("cCode"),xt("cProd"),xt("cCat"),xt("cUnit"),xt("cLast"),xt("cChg"),xt("cOpen"),xt("cLow"),xt("cHigh"),xt("cAvg"),xt("cRec"),xt("cMkts"),xt("cLastD")];
  header(R,9,qh);
  sums.forEach((s,i)=>R.row(10+i,[{v:s.tk,s:XS.bold},{v:s.name,s:XS.cell},{v:catName(s.cat),s:XS.cell},{v:unitOf(s.prod),s:XS.cell},{v:xNum(s.last),s:XS.num},xPct(s.pct)||{v:"",s:XS.cell},
    {v:xNum(s.open),s:XS.num},{v:xNum(s.lo),s:XS.num},{v:xNum(s.hi),s:XS.num},{v:xNum(s.avg),s:XS.num},{v:s.n,s:XS.int},{v:s.mkts,s:XS.int},xDate(s.lastD)]));
  R.filter=`A10:M${10+sums.length}`;
  [9,24,18,12,11,10,11,11,11,11,10,10,13].forEach((w,c)=>R.width(c,w));

  // --- Séries (dados dos gráficos) ---
  const SE=wb.sheet(xt("sSeries"));
  const buckets=[...new Set([...idx.map(x=>x.t), ...sums.flatMap(s=>s.cs.map(c=>c.t))])].sort((a,b)=>a-b);
  const top=sums.slice().sort((a,b)=>b.n-a.n);
  header(SE,0,[xt("cDate"),xt("cIdx"),...top.map(s=>`${s.tk} - ${s.name}`)]);
  const iMap=new Map(idx.map(x=>[x.t,x.c]));
  const pMaps=top.map(s=>new Map(s.cs.map(c=>[c.t,c.c])));
  buckets.forEach((b,i)=>SE.row(i+1,[xDate(b),{v:xNum(iMap.get(b)),s:XS.num},...pMaps.map(m=>m.has(b)?{v:xNum(m.get(b)),s:XS.num}:null)]));
  SE.freeze={r:1,c:1}; SE.width(0,12); SE.width(1,12); top.forEach((s,i)=>SE.width(i+2,Math.min(28,Math.max(12,(s.tk+s.name).length+4))));
  SE.set(buckets.length+2,0,{v:xt("seriesNote"),s:XS.note});
  const nB=buckets.length;

  // --- Gráficos (nativos do Excel, ligados a «Séries») ---
  const G=wb.sheet(xt("sCharts")); G.grid=false;
  G.set(0,0,{v:xt("sCharts")+" - "+xt("title"),s:XS.title}); G.height(0,24);
  G.set(1,0,{v:xt("chartsNote"),s:XS.note});
  if(nB){
    const cat={sheet:SE.name,col:0,r1:1,r2:nB};
    G.lineChart({title:`${xt("cIdx")} (${LANG==="pt"?"base 100":"base 100"})`,cat,series:[{name:xt("cIdx"),sheet:SE.name,col:1,r1:1,r2:nB,color:"B86E00",width:28575}],at:{c1:0,r1:3,c2:16,r2:21},yFmt:"0.0"});
    const nCharts=Math.min(12,top.length);
    for(let k=0;k<nCharts;k++){ const s=top[k], row=23+Math.floor(k/2)*18, c1=(k%2)*8;
      G.lineChart({title:`${s.tk} - ${s.name} (${unitOf(s.prod)})`,cat,series:[{name:s.name,sheet:SE.name,col:k+2,r1:1,r2:nB}],at:{c1,r1:row,c2:c1+8,r2:row+17}}); }
  }
  for(let c=0;c<16;c++) G.width(c,11);

  // --- Preço por mercado ---
  const PM=wb.sheet(xt("sMkt"));
  PM.set(0,0,{v:xt("sMkt"),s:XS.title}); PM.height(0,24); PM.set(1,0,{v:xt("mktNote"),s:XS.note}); PM.set(2,0,{v:`${xt("filters")}: ${filtersLine()}`,s:XS.note});
  const mk=[...new Set(list.map(o=>o.mkt))].sort((a,b)=>{ const pa=list.find(o=>o.mkt===a)?.prov||"", pb=list.find(o=>o.mkt===b)?.prov||""; return pa.localeCompare(pb)||a.localeCompare(b); });
  const provOf=new Map(list.map(o=>[o.mkt,o.prov]));
  const acc=new Map(); for(const o of list){ const k=o.mkt+"|"+o.prod; const a=acc.get(k)||{s:0,n:0}; a.s+=o.price; a.n++; acc.set(k,a); }
  header(PM,4,[xt("cMkt"),xt("cProv"),...sums.map(s=>`${s.tk}\n${s.name}`)]); PM.height(4,44);
  const colMinMax=sums.map(s=>{ const v=mk.map(m=>acc.get(m+"|"+s.prod)).filter(Boolean).map(a=>a.s/a.n); return v.length>1?[Math.min(...v),Math.max(...v)]:[null,null]; });
  mk.forEach((m,i)=>PM.row(5+i,[{v:m,s:XS.bold},{v:provOf.get(m),s:XS.cell},...sums.map((s,j)=>{ const a=acc.get(m+"|"+s.prod); if(!a) return {v:"",s:XS.cell}; const v=a.s/a.n;
    return {v:xNum(v), s: v===colMinMax[j][0]?XS.numLow : v===colMinMax[j][1]?XS.numHigh : XS.num}; })]));
  PM.freeze={r:5,c:2}; PM.width(0,32); PM.width(1,18); sums.forEach((s,j)=>PM.width(j+2,12));

  // --- Registos recentes ---
  const RR=wb.sheet(xt("sRecent"));
  header(RR,0,[xt("cDate"),xt("cCode"),xt("cProd"),xt("cMkt"),xt("cProv"),xt("cSel"),xt("cPrice"),xt("cUnit"),xt("cVsAvg")]);
  const avg=new Map(sums.map(s=>[s.prod,s.avg]));
  const recent=list.slice().sort((a,b)=>b.d-a.d||a.prod.localeCompare(b.prod)).slice(0,1000);
  recent.forEach((o,i)=>RR.row(i+1,[xDate(o.d),{v:pTk(o.prod),s:XS.bold},{v:pName(o.prod),s:XS.cell},{v:o.mkt,s:XS.cell},{v:o.prov,s:XS.cell},{v:sellerName(o.sel),s:XS.cell},
    {v:xNum(o.price),s:XS.num},{v:unitOf(o.prod),s:XS.cell},xPct(avg.get(o.prod)?(o.price/avg.get(o.prod)-1)*100:null)||{v:"",s:XS.cell}]));
  RR.freeze={r:1,c:0}; RR.filter=`A1:I${recent.length+1}`; [12,9,24,32,18,14,11,12,10].forEach((w,c)=>RR.width(c,w));

  // --- Boletim (selecção actual do separador Boletim semanal) ---
  if(typeof bCompute==="function") buildBulletinSheet(wb);
  return wb;
}
function buildBulletinSheet(wb){
  const levels=bLevelsToShow(), multi=B.level==="ALL";
  const Rs=levels.map(L=>bCompute(L)).filter(r=>r&&r.mainStats.length);
  const R0=Rs.find(r=>r.level===BC.level)||Rs[0]; if(!R0) return;
  const sh=wb.sheet(xt("sBull")); sh.grid=false; const W=10; let r=0;
  sh.set(r,0,{v:`${LANG==="pt"?BC.title:BC.titleEn} - ${t("bNo")} ${R0.number}`,s:XS.title}); sh.height(r,26); r++;
  sh.set(r,0,{v:weekLabel(R0.w0),s:XS.subtitle}); r++;
  sh.set(r,0,{v:bScopeLabel(),s:XS.note}); r+=2;
  const para=(txt,s=XS.wrap,h)=>{ sh.set(r,0,{v:txt,s}); for(let c=1;c<W;c++) sh.set(r,c,{v:"",s}); sh.merge(r,0,r,W-1); sh.height(r,h||Math.max(18,Math.ceil(String(txt).length/125)*16)); r++; };
  para(bHeadline(R0),XS.boldWrap,22); para(bIntro(R0)); r++;
  Rs.forEach((R,i)=>{
    const L=R.level;
    if(multi){ band(sh,r,`${xt("part")} ${ROMAN[i]||i+1}: ${lvName(L)}`.toUpperCase(),W); r+=2; }
    // dinâmica de preços
    sh.set(r,0,{v:`${xt("dyn")} - ${lvName(L)}`,s:XS.subtitle}); r++;
    header(sh,r,[xt("cProd"),xt("cAvg"),xt("cPrev"),xt("cChg"),xt("cHigh"),xt("cMaxM"),xt("cLow"),xt("cMinM"),xt("cMkts"),xt("cUnit")]); r++;
    R.mainStats.forEach(s=>{ sh.row(r++,[{v:pName(s.prod),s:XS.bold},{v:xNum(s.avg),s:XS.num},{v:xNum(s.avgP),s:XS.num},xPct(s.chg)||{v:"",s:XS.cell},{v:xNum(s.max),s:XS.numHigh},{v:s.maxM.join(", "),s:XS.cell},
      {v:xNum(s.min),s:XS.numLow},{v:s.minM.join(", "),s:XS.cell},{v:s.nMk,s:XS.int},{v:unitLbl(s.prod),s:XS.cell}]); });
    r++;
    // preços por mercado e variação (todos os produtos deste nível)
    const T=R.levelTables[L];
    if(T){ sh.set(r,0,{v:`${xt("weekly")} - ${lvName(L)}`,s:XS.subtitle}); r++;
      header(sh,r,[xt("cProd"),xt("cMkt"),xt("cProv"),xt("cPrev"),xt("cCur"),xt("cChg")]); r++;
      T.prods.forEach(p=>T.mkts.forEach(m=>{ const x=T.cell[p+"|"+m]; if(!x) return;
        sh.row(r++,[{v:pName(p),s:XS.cell},{v:m,s:XS.cell},{v:T.provOf[m],s:XS.cell},{v:xNum(x.prev),s:XS.num},{v:xNum(x.cur),s:XS.num},xPct(x.pct)||{v:"",s:XS.cell}]); }));
      r++; }
    // fluxos
    const fl=Object.entries(R.flows);
    if(fl.length){ sh.set(r,0,{v:`${xt("flows")} - ${lvName(L)}`,s:XS.subtitle}); r++;
      header(sh,r,[xt("cZone"),xt("cProd"),xt("cMkt"),xt("cOrig"),xt("cAvg")]); r++;
      ["Sul","Centro","Norte","Outra"].forEach(z=>Object.entries(R.flows[z]||{}).forEach(([p,M])=>Object.values(M).forEach(e=>{
        sh.row(r++,[{v:t("z"+z),s:XS.cell},{v:pName(p),s:XS.cell},{v:e.mkt,s:XS.cell},{v:[...e.orig].join(", "),s:XS.cell},{v:xNum(mean(e.pr)),s:XS.num}]); })));
      r++; }
  });
  if(multi){ band(sh,r,t("bCommon").toUpperCase(),W); r+=2; }
  if(R0.qtyRows.length){ sh.set(r,0,{v:xt("qty"),s:XS.subtitle}); r++;
    header(sh,r,[xt("cMkt"),xt("cProd"),xt("cPrev"),xt("cCur"),xt("cChg")]); r++;
    R0.qtyRows.forEach(q=>sh.row(r++,[{v:q.mkt,s:XS.cell},{v:pName(q.prod),s:XS.cell},q.p?{v:Math.round(q.p),s:XS.int}:{v:"",s:XS.cell},q.c?{v:Math.round(q.c),s:XS.int}:{v:"",s:XS.cell},xPct(q.c&&q.p?pct(q.c,q.p):null)||{v:"",s:XS.cell}]));
    r++; }
  if(R0.liveP.length){ sh.set(r,0,{v:xt("live"),s:XS.subtitle}); r++;
    header(sh,r,[xt("cProv"),xt("cProd"),`${xt("cMonthPrev")} (${fmtD(R0.monthPrev,{month:"long"})})`,`${xt("cMonthCur")} (${fmtD(R0.monthCur,{month:"long"})})`,xt("cChg")]); r++;
    R0.liveProv.forEach(pv=>R0.liveP.forEach(p=>{ const x=R0.live[pv][p]; if(!x) return; const c=mean(x.c), pp=mean(x.p);
      sh.row(r++,[{v:pv,s:XS.cell},{v:pName(p),s:XS.cell},{v:xNum(pp),s:XS.num},{v:xNum(c),s:XS.num},xPct(pct(c,pp))||{v:"",s:XS.cell}]); })); }
  [26,30,18,14,14,28,12,28,11,13].forEach((w,c)=>sh.width(c,w));
}

/* ---------------- interface ---------------- */
async function exportExcel(kind){
  const menu=$("xlMenu"); if(menu) menu.hidden=true;
  if(!S.obs.length){ toast(t("noData")); return; }
  toast(xt("working"),15000); await new Promise(r=>setTimeout(r,40));
  try{
    const wb = kind==="raw" ? buildRaw() : buildFormatted();
    const day=new Date().toISOString().slice(0,10);
    const fname = kind==="raw" ? `SIMA_${LANG==="pt"?"dados_brutos":"raw_data"}_${day}.xlsx` : `SIMA_${LANG==="pt"?"relatorio":"report"}_${day}.xlsx`;
    await saveFile(fname, await wb.blob()); toast(xt("done"));
  }catch(e){ console.error(e); toast(e?.code==="declined"?t("pdfDeclined"):xt("fail")+": "+(e?.message||e)); }
}
function renderXlMenu(){
  $("btnXl").textContent=xt("btn");
  $("xlMenu").innerHTML=`<button type="button" data-xl="raw"><b>${esc(xt("raw"))}</b><span>${esc(xt("rawD"))}</span></button><button type="button" data-xl="fmt"><b>${esc(xt("fmt"))}</b><span>${esc(xt("fmtD"))}</span></button>`;
}
$("btnXl").addEventListener("click",e=>{ e.stopPropagation(); const m=$("xlMenu"); m.hidden=!m.hidden; $("btnXl").setAttribute("aria-expanded",String(!m.hidden)); });
$("xlMenu").addEventListener("click",e=>{ const b=e.target.closest("[data-xl]"); if(b) exportExcel(b.dataset.xl); });
document.addEventListener("click",e=>{ if(!e.target.closest(".xl-wrap")) { const m=$("xlMenu"); if(m) m.hidden=true; } });
document.addEventListener("keydown",e=>{ if(e.key==="Escape"){ const m=$("xlMenu"); if(m) m.hidden=true; } });
$("lang-pt").addEventListener("click",renderXlMenu); $("lang-en").addEventListener("click",renderXlMenu);
renderXlMenu();
