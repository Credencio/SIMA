"use strict";
/* =====================================================================
   Aba "Área Técnica": fluxo de validação, contagem por estado (campo "validacao"
   da tabela de produtos) e aplicação ArcGIS Experience de validação embebida.
   Depende de: app.js (S.valid é preenchido na leitura dos dados)
   ===================================================================== */
const TECH_T = {
  pt:{ flow:[["Submissão","Sem login: formulário Survey123","ok"],["Pendente","Todo o novo registo","pend"],["Validação","Técnico com login ArcGIS","ok"],["Preços e Boletim","Só dados validados","ok"]],
    status:"Estado dos registos na base de dados:", pend:"Pendente", val:"Validado", rej:"Rejeitado", total:"Total", updated:"Contagem lida às {h}", demo:"Contagem de demonstração",
    noField:"A tabela de produtos não tem o campo «{f}»: todos os registos contam como pendentes.", sub:"Validação de dados", subTitle:"Validação de dados (ArcGIS Experience)",
    subNote:"Acesso restrito: inicie sessão com utilizador e palavra-passe do ArcGIS Online. Validado → aparece nas abas Preços e Boletim · Rejeitado → fica de fora.",
    open:"Abrir em nova janela", refresh:"Actualizar contagem", frameNote:"Se a aplicação não abrir aqui (por exemplo, por causa do início de sessão), use «Abrir em nova janela».",
    iframeTitle:"Aplicação de validação de dados (ArcGIS Experience)", records:"registos de produto" },
  en:{ flow:[["Submission","No login: Survey123 form","ok"],["Pending","Every new record","pend"],["Validation","Technician with ArcGIS login","ok"],["Prices and Bulletin","Validated data only","ok"]],
    status:"Record status in the database:", pend:"Pending", val:"Validated", rej:"Rejected", total:"Total", updated:"Counted at {h}", demo:"Demo count",
    noField:"The products table has no «{f}» field: every record counts as pending.", sub:"Data validation", subTitle:"Data validation (ArcGIS Experience)",
    subNote:"Restricted access: sign in with your ArcGIS Online username and password. Validated → shown in the Prices and Bulletin tabs · Rejected → left out.",
    open:"Open in a new window", refresh:"Refresh count", frameNote:"If the app does not open here (for example because of the sign-in), use «Open in a new window».",
    iframeTitle:"Data validation app (ArcGIS Experience)", records:"product records" }
};
const tt = (k,o={}) => { const v=(TECH_T[LANG]||TECH_T.pt)[k]; return typeof v==="string" ? v.replace(/\{(\w+)\}/g,(m,x)=>o[x]??m) : v; };

function renderTech(){
  const root=$("viewTech"); if(!root || root.hidden) return;
  const cfg=SIMA_CONFIG.technical||{}, url=cfg.experienceUrl||"";
  // estrutura criada uma só vez: o iframe nunca é recriado nem movido (não perde a sessão iniciada)
  if(!root.dataset.built){
    root.innerHTML=`<div class="tk-flow" id="techFlow"></div><div class="tk-status" id="techStatus"></div><div id="techWarn"></div>
      <div class="tk-subtabs" role="tablist"><button role="tab" aria-pressed="true" data-sub="validation" id="techSubBtn"></button></div>
      <section class="panel tk-panel"><div class="tk-head"><div><h2 id="techSubTitle"></h2><p id="techSubNote"></p></div>
        <a class="btn tk-open" id="techOpen" target="_blank" rel="noopener"></a></div>
        <div class="tk-frame" id="techFrameBox"></div><p class="tk-note" id="techNote"></p></section>`;
    if(url){ const f=document.createElement("iframe"); f.id="techFrame"; f.src=url; f.title=tt("iframeTitle");
      f.setAttribute("allow","geolocation; clipboard-read; clipboard-write; fullscreen"); f.setAttribute("allowfullscreen",""); $("techFrameBox").appendChild(f); }
    root.dataset.built="1";
    root.addEventListener("click",e=>{ if(e.target.id==="techRefresh" && typeof load==="function") load(); });
  }
  const V=S.valid||{val:0,pend:0,rej:0,field:true}, tot=V.val+V.pend+V.rej;
  const share=n=>tot?` <small>${fmtN(n/tot*100,0)}%</small>`:"";
  const when = V.demo ? tt("demo") : (V.at ? tt("updated",{h:new Date(V.at).toLocaleTimeString(LANG==="pt"?"pt-PT":"en-GB",{hour:"2-digit",minute:"2-digit",timeZone:"Africa/Maputo"})}) : "");
  $("techFlow").innerHTML=tt("flow").map(([h,s,k],i)=>`<div class="tk-step tk-${k}"><b>${esc(h)}</b><span>${esc(s)}</span></div>${i<3?'<span class="tk-arrow" aria-hidden="true">→</span>':""}`).join("");
  $("techStatus").innerHTML=`<b>${esc(tt("status"))}</b>
    <span class="tk-pill tk-pend">${esc(tt("pend"))}: <b class="num">${fmtN(V.pend,0)}</b>${share(V.pend)}</span>
    <span class="tk-pill tk-val">${esc(tt("val"))}: <b class="num">${fmtN(V.val,0)}</b>${share(V.val)}</span>
    <span class="tk-pill tk-rej">${esc(tt("rej"))}: <b class="num">${fmtN(V.rej,0)}</b>${share(V.rej)}</span>
    <span class="tk-meta">${esc(tt("total"))}: ${fmtN(tot,0)} ${esc(tt("records"))}${when?" · "+esc(when):""}</span>
    <button class="btn tk-refresh" type="button" id="techRefresh">${esc(tt("refresh"))}</button>`;
  $("techWarn").innerHTML = V.field===false ? `<div class="banner">${esc(tt("noField",{f:cfg.validationField||"validacao"}))}</div>` : "";
  $("techSubBtn").textContent=tt("sub"); $("techSubTitle").textContent=tt("subTitle"); $("techSubNote").textContent=tt("subNote");
  const a=$("techOpen"); a.textContent=tt("open"); a.href=url; $("techNote").textContent=tt("frameNote");
  const f=$("techFrame"); if(f) f.title=tt("iframeTitle");
}
$("lang-pt").addEventListener("click",()=>setTimeout(renderTech,0));
$("lang-en").addEventListener("click",()=>setTimeout(renderTech,0));
