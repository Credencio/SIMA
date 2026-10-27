"use strict";
/* =====================================================================
   Mapa dos locais de registo de preços (aba Preços)
   Um marcador por mercado, na posição média dos registos com coordenadas.
   Tamanho = número de registos; cor = preço do produto seleccionado face à média dos mercados.
   Depende de: Leaflet (cdnjs), app.js
   ===================================================================== */
const MAP_T = {
  pt:{ title:"Locais de registo de preços", hint:"Um ponto por mercado e tipo de mercado", noCoords:"Os registos filtrados não têm coordenadas.",
    noLib:"Não foi possível carregar o mapa (sem ligação à internet).", colorBy:"Cor por", bySel:"Tipo de mercado", byLoc:"Localização",
    records:"Registos", lastRec:"Último registo", products:"Produtos registados (preço médio)", filterMkt:"Ver só este mercado", clearMkt:"Ver todos os mercados",
    loc:"Localização", type:"Tipo de mercado", points:"pontos no mapa", withoutXY:"mercados sem coordenadas", unknown:"Não indicado" },
  en:{ title:"Where prices were recorded", hint:"One point per market and market type", noCoords:"The filtered records have no coordinates.",
    noLib:"The map could not be loaded (no internet connection).", colorBy:"Colour by", bySel:"Market type", byLoc:"Location",
    records:"Records", lastRec:"Latest record", products:"Products recorded (average price)", filterMkt:"Show only this market", clearMkt:"Show all markets",
    loc:"Location", type:"Market type", points:"points on the map", withoutXY:"markets without coordinates", unknown:"Not stated" }
};
const mt = (k,o={}) => String((MAP_T[LANG]||MAP_T.pt)[k]).replace(/\{(\w+)\}/g,(m,x)=>o[x]??m);
// cores fixas por categoria (iguais em tema claro e escuro, com contorno branco)
const MAP_PAL = { sel:{ Produtor:"#2E9E5B", Grossista:"#2F6FB3", Retalhista:"#D9822B" },
                  loc:{ "Cidade":"#7B4FB0", "Municipio/Vila":"#1F9C9C", "Rural":"#9C6B30" } };
const MAP_EXTRA = ["#C2417A","#5C6BC0","#8D9A2E","#B5651D"], MAP_NONE="#8494A3";
let MAP=null, MAP_LAYER=null, MAP_TILES=null, MAP_FITKEY="", MAP_BY="sel";
try{ const v=localStorage.getItem("sima.mapBy"); if(v==="loc"||v==="sel") MAP_BY=v; }catch(e){}

function isDark(){ const th=document.documentElement.getAttribute("data-theme"); return th ? th==="dark" : window.matchMedia("(prefers-color-scheme: dark)").matches; }
// Mapa base Esri (cinzento, sem chave de acesso), com camada de nomes por cima
function setTiles(){
  if(!MAP) return; const dark=isDark(), key=dark?"dark":"light"; if(MAP_TILES && MAP_TILES._k===key) return;
  if(MAP_TILES) MAP_TILES.forEach(l=>MAP.removeLayer(l));
  const base=dark?"World_Dark_Gray":"World_Light_Gray", attr='Mapa base &copy; <a href="https://www.esri.com" target="_blank" rel="noopener">Esri</a>, HERE, Garmin, &copy; OpenStreetMap';
  MAP_TILES=[L.tileLayer(`https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/${base}_Base/MapServer/tile/{z}/{y}/{x}`,{maxZoom:16,attribution:attr}).addTo(MAP),
             L.tileLayer(`https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/${base}_Reference/MapServer/tile/{z}/{y}/{x}`,{maxZoom:16,zIndex:5}).addTo(MAP)];
  MAP_TILES._k=key;
}
function ensureMap(){
  if(MAP) return true;
  const el=$("map"); if(!el) return false;
  if(typeof L==="undefined"){ $("mapMsg").textContent=mt("noLib"); $("mapMsg").hidden=false; return false; }
  MAP=L.map(el,{zoomControl:true,scrollWheelZoom:false,attributionControl:true,maxZoom:16}).setView([-18.6,35.5],5);
  setTiles();
  MAP_LAYER=L.layerGroup().addTo(MAP);
  MAP.on("click",()=>MAP.scrollWheelZoom.enable()); MAP.on("mouseout",()=>MAP.scrollWheelZoom.disable());
  return true;
}
function catLabel(by,v){ if(!v||v==="—") return mt("unknown"); return by==="sel" ? sellerName(v) : v; }
function catColor(by,v,extra){ const P=MAP_PAL[by]; if(!v||v==="—") return MAP_NONE; if(P[v]) return P[v];
  if(!(v in extra)) extra[v]=MAP_EXTRA[Object.keys(extra).length%MAP_EXTRA.length]; return extra[v]; }
function renderMap(){
  const view=$("viewMarket"); if(!view || view.hidden) return;
  $("mapTitle").textContent=mt("title"); $("mapHint").textContent=mt("hint");
  $("mapByLbl").textContent=mt("colorBy"); const sel=$("mapBy");
  sel.innerHTML=`<option value="sel"${MAP_BY==="sel"?" selected":""}>${esc(mt("bySel"))}</option><option value="loc"${MAP_BY==="loc"?" selected":""}>${esc(mt("byLoc"))}</option>`;
  if(!ensureMap()) return;
  setTiles();
  const list=filtered(false), by=MAP_BY;
  // um ponto por mercado e categoria (tipo de mercado ou localização)
  const groups=new Map(), mktAll=new Set(), mktXY=new Set();
  for(const o of list){ mktAll.add(o.mkt); const cat=by==="sel"?o.sel:(o.loc||"—"); const k=o.mkt+"|"+cat;
    let g=groups.get(k); if(!g){ g={mkt:o.mkt,cat,prov:o.prov,dist:o.dist,loc:o.loc,sel:o.sel,lat:[],lon:[],n:0,last:0,prods:new Map()}; groups.set(k,g); }
    g.n++; g.last=Math.max(g.last,o.d); if(o.lat!=null&&o.lon!=null){ g.lat.push(o.lat); g.lon.push(o.lon); mktXY.add(o.mkt); }
    const p=g.prods.get(o.prod)||{s:0,n:0}; p.s+=o.price; p.n++; g.prods.set(o.prod,p); }
  const med=a=>{ const s=a.slice().sort((x,y)=>x-y), k=Math.floor(s.length/2); return s.length%2?s[k]:(s[k-1]+s[k])/2; };
  const pts=[...groups.values()].filter(g=>g.lat.length).map(g=>({...g,y:med(g.lat),x:med(g.lon)}));
  // pontos do mesmo mercado no mesmo sítio: afastar ligeiramente para todos ficarem visíveis
  const near=new Map(); pts.forEach(p=>{ const k=p.y.toFixed(3)+","+p.x.toFixed(3); (near.get(k)||near.set(k,[]).get(k)).push(p); });
  near.forEach(arr=>{ if(arr.length<2) return; arr.sort((a,b)=>String(a.cat).localeCompare(String(b.cat))).forEach((p,i)=>{ const ang=2*Math.PI*i/arr.length; p.y+=0.0022*Math.sin(ang); p.x+=0.0022*Math.cos(ang); }); });
  const extra={};
  MAP_LAYER.clearLayers();
  for(const g of pts){
    const col=catColor(by,g.cat,extra);
    const top=[...g.prods.entries()].sort((a,b)=>b[1].n-a[1].n).slice(0,6);
    const html=`<div class="mp"><div class="mp-h">${esc(g.mkt)}</div><div class="mp-s">${esc(g.prov)}${g.dist?" · "+esc(g.dist):""}</div>
      <div class="mp-r"><span>${esc(mt("type"))}</span><b class="mp-tag" style="--c:${col}">${esc(by==="sel"?catLabel("sel",g.cat):sellerName(g.sel))}</b></div>
      ${g.loc?`<div class="mp-r"><span>${esc(mt("loc"))}</span><b>${esc(g.loc)}</b></div>`:""}
      <div class="mp-r"><span>${esc(mt("records"))}</span><b>${fmtN(g.n,0)}</b></div><div class="mp-r"><span>${esc(mt("lastRec"))}</span><b>${fmtDate(g.last)}</b></div>
      <div class="mp-l">${esc(mt("products"))}</div>${top.map(([pp,a])=>`<div class="mp-r"><span>${esc(pName(pp))}</span><b>${fmtN(a.s/a.n)}</b></div>`).join("")}
      <button type="button" class="btn mp-btn" data-mapmkt="${esc(g.mkt)}" data-mapprov="${esc(g.prov)}">${esc(S.mkt===g.mkt?mt("clearMkt"):mt("filterMkt"))}</button></div>`;
    L.circleMarker([g.y,g.x],{radius:8,color:"#ffffff",weight:2,fillColor:col,fillOpacity:.95})
      .bindTooltip(`<b>${esc(g.mkt)}</b> · ${esc(catLabel(by,g.cat))}`,{direction:"top",offset:[0,-8]})
      .bindPopup(html,{maxWidth:290}).addTo(MAP_LAYER);
  }
  // legenda: só as categorias presentes
  const cats=[...new Set(pts.map(p=>p.cat))].sort((a,b)=>catLabel(by,a).localeCompare(catLabel(by,b)));
  const noXY=mktAll.size-mktXY.size;
  $("mapLegend").innerHTML = `<span class="ml-t">${esc(by==="sel"?mt("bySel"):mt("byLoc"))}</span>`+
    cats.map(c=>`<span class="ml-i"><i style="background:${catColor(by,c,extra)}"></i>${esc(catLabel(by,c))} <small>(${pts.filter(p=>p.cat===c).length})</small></span>`).join("")+
    `<span class="ml-n">${pts.length} ${esc(mt("points"))}${noXY?` · ${noXY} ${esc(mt("withoutXY"))}`:""}</span>`;
  $("mapMsg").hidden = pts.length>0; if(!pts.length) $("mapMsg").textContent=mt("noCoords");
  const key=pts.map(p=>p.mkt).sort().join("|");
  if(key!==MAP_FITKEY && pts.length){ MAP_FITKEY=key; MAP.fitBounds(L.latLngBounds(pts.map(p=>[p.y,p.x])).pad(0.25),{maxZoom:pts.length<3?12:9}); }
  setTimeout(()=>MAP.invalidateSize(),60);
}
$("mapBy").addEventListener("change",e=>{ MAP_BY=e.target.value; try{ localStorage.setItem("sima.mapBy",MAP_BY); }catch(_){} renderMap(); });
// botão do popup: filtrar pelo mercado (ou limpar)
document.addEventListener("click",e=>{ const b=e.target.closest("[data-mapmkt]"); if(!b) return;
  const m=b.dataset.mapmkt; if(S.mkt===m){ S.mkt=""; } else { S.mkt=m; S.prov=b.dataset.mapprov||S.prov; }
  MAP && MAP.closePopup(); renderFilters(); renderAll(); });
$("tabs").addEventListener("click",e=>{ const b=e.target.closest("button"); if(b && b.dataset.v==="market") setTimeout(renderMap,30); });
$("lang-pt").addEventListener("click",()=>setTimeout(renderMap,0)); $("lang-en").addEventListener("click",()=>setTimeout(renderMap,0));
window.matchMedia("(prefers-color-scheme: dark)").addEventListener?.("change",()=>renderMap());
setTimeout(renderMap,0);
