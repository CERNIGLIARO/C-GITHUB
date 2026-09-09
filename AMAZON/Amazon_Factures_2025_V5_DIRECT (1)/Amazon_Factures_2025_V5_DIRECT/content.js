
const sleep=ms=>new Promise(r=>setTimeout(r,ms));

async function getState(){
  const {amazonV5State}=await chrome.storage.local.get("amazonV5State");
  return amazonV5State||{};
}
async function setState(s){await chrome.storage.local.set({amazonV5State:s});}
const clean=e=>(e?.innerText||e?.textContent||"").replace(/\s+/g," ").trim();
const ordersUrl=y=>`https://www.amazon.fr/gp/your-account/order-history?orderFilter=year-${encodeURIComponent(y)}`;

function orderIdFrom(el){
  const m=clean(el).match(/\b\d{3}-\d{7}-\d{7}\b/);
  return m?m[0]:(el.getAttribute?.("data-order-id")||"commande_sans_numero");
}

function findCards(){
  const candidates=[
    ...document.querySelectorAll("div.order-card"),
    ...document.querySelectorAll("div.js-order-card"),
    ...document.querySelectorAll("[class*='order-card']"),
    ...document.querySelectorAll("[data-order-id]")
  ];
  const out=[],seen=new Set();
  for(const el of candidates){
    const id=orderIdFrom(el);
    if(id==="commande_sans_numero"||seen.has(id))continue;
    seen.add(id);out.push(el);
  }
  if(out.length)return out;

  for(const el of document.querySelectorAll("div.a-box,div.a-section")){
    const id=orderIdFrom(el);
    if(id!=="commande_sans_numero"&&!seen.has(id)){
      seen.add(id);out.push(el);
    }
  }
  return out;
}

function extractDate(card){
  const t=clean(card);
  const m=t.match(/(?:commande effectuée le|commandé le|ordered on)\s*[:\-]?\s*([^|]{5,35})/i);
  return m?m[1].trim():"";
}
function extractTotal(card){
  const t=clean(card);
  const m=t.match(/(?:total|montant)\s*[:\-]?\s*([0-9\s.,]+\s*(?:€|EUR))/i);
  return m?m[1].trim():"";
}

function isForbiddenLabel(t){
  return /r[eé]capitulatif|order summary|reçu|receipt|d[eé]tails?|bon de commande|demander|request|avoir|credit note|note de cr[eé]dit/i.test(t);
}
function isInvoiceLabel(t){
  t=String(t||"").replace(/\s+/g," ").trim();
  if(!t||isForbiddenLabel(t))return false;
  return /^(facture|factures|invoice|invoices)\b/i.test(t);
}

function normalizeLinks(anchors){
  const out=[],seen=new Set();
  for(const a of anchors){
    const label=clean(a);
    if(!isInvoiceLabel(label))continue;

    const href=a.getAttribute("href")||"";
    if(!href||/^javascript:|^#$/i.test(href))continue;

    let url;
    try{url=new URL(href,location.href).href;}catch{continue;}
    if(!url.startsWith("https://www.amazon.fr/"))continue;
    if(seen.has(url))continue;
    seen.add(url);
    out.push({url,label});
  }
  return out;
}

function visiblePopovers(){
  return [...document.querySelectorAll(".a-popover,[role='dialog']")].filter(el=>{
    const s=getComputedStyle(el);
    return s.display!=="none"&&s.visibility!=="hidden"&&el.getClientRects().length>0;
  });
}

async function invoiceLinksForCard(card){
  let all=normalizeLinks([...card.querySelectorAll("a[href]")]);

  const triggers=[...card.querySelectorAll("a,button,span[role='button']")].filter(el=>{
    const t=clean(el);
    if(!isInvoiceLabel(t))return false;
    return !isForbiddenLabel(t);
  });

  for(const tr of triggers.slice(0,4)){
    try{
      tr.click();
      await sleep(550);
      const popLinks=normalizeLinks(
        visiblePopovers().flatMap(p=>[...p.querySelectorAll("a[href]")])
      );
      all=[...popLinks,...all];
      if(popLinks.length)break;
    }catch{}
  }

  const out=[],seen=new Set();
  for(const x of all){
    if(seen.has(x.url))continue;
    seen.add(x.url);out.push(x);
  }
  return out;
}

function nextPage(){
  for(const s of[
    ".a-pagination .a-last a",
    "li.a-last a",
    "a[aria-label*='Suivant']",
    "a[aria-label*='Next']"
  ]){
    const a=document.querySelector(s);
    if(a?.href)return a.href;
  }
  return [...document.querySelectorAll("a[href]")].find(a=>/^(suivant|next)\b/i.test(clean(a)))?.href||null;
}

async function appendRow(row){
  const st=await getState();
  st.rows=Array.isArray(st.rows)?st.rows:[];
  if(!st.rows.some(r=>r.orderId===row.orderId))st.rows.push(row);
  await setState(st);
}

async function processPage(){
  let st=await getState();
  if(!st.running)return;

  if(!location.href.includes(`orderFilter=year-${st.year}`)){
    location.href=ordersUrl(st.year);return;
  }

  st.visited=Array.isArray(st.visited)?st.visited:[];
  if(st.visited.includes(location.href)){
    st.running=false;
    st.status="Arrêt sécurité : page déjà analysée.";
    await setState(st);return;
  }

  st.visited.push(location.href);
  st.pages=(st.pages||0)+1;
  st.status=`Analyse page ${st.pages}…`;
  await setState(st);
  await sleep(900);

  const cards=findCards();
  if(!cards.length){
    st=await getState();st.running=false;
    st.status="Aucune commande détectée ou reconnexion Amazon nécessaire.";
    await setState(st);return;
  }

  for(const card of cards){
    st=await getState();
    if(!st.running)return;

    const oid=orderIdFrom(card),date=extractDate(card),total=extractTotal(card);
    st.orders=(st.orders||0)+1;
    st.status=`Commande ${st.orders} : ${oid}`;
    await setState(st);

    try{
      const links=await invoiceLinksForCard(card);

      st=await getState();
      st.candidates=(st.candidates||0)+links.length;
      st.status=`${oid} : ${links.length} lien(s) Facture détecté(s).`;
      await setState(st);

      if(!links.length){
        st=await getState();
        st.missing=(st.missing||0)+1;
        await setState(st);
        await appendRow({
          date,orderId:oid,total,status:"Aucun lien Facture",
          candidates:0,downloaded:0,rejected:0,details:""
        });
        continue;
      }

      let downloaded=0,rejected=0;
      const details=[];

      for(let i=0;i<links.length;i++){
        st=await getState();
        if(!st.running)return;

        const l=links[i];
        st.status=`${oid} : tentative directe « ${l.label} »…`;
        await setState(st);

        const resp=await chrome.runtime.sendMessage({
          type:"V5_TRY_DOWNLOAD",
          url:l.url,
          orderId:oid,
          label:l.label,
          index:i+1
        });

        if(resp?.ok){
          downloaded++;
          details.push(`OK ${l.label} | ${resp.mime||""} | ${resp.bytes||0} octets`);
        }else if(resp?.rejected){
          rejected++;
          details.push(`REJET ${l.label} | ${resp.reason||"non-PDF"}`);
        }else{
          details.push(`ERREUR ${l.label} | ${resp?.error||"inconnue"}`);
        }
        await sleep(250);
      }

      if(downloaded===0){
        st=await getState();st.missing=(st.missing||0)+1;await setState(st);
      }

      await appendRow({
        date,orderId:oid,total,
        status:downloaded?"Vrai PDF téléchargé":"Aucun vrai PDF",
        candidates:links.length,downloaded,rejected,
        details:details.join(" | ")
      });

    }catch(e){
      st=await getState();
      st.errors=(st.errors||0)+1;
      st.status=`Erreur ${oid} : ${e.message}. Suite…`;
      await setState(st);
      await appendRow({
        date,orderId:oid,total,status:"Erreur",
        candidates:0,downloaded:0,rejected:0,details:e.message
      });
    }
  }

  st=await getState();
  if(!st.running)return;

  const next=nextPage();
  if(next){
    st.status=`Page ${st.pages} terminée. Page suivante…`;
    await setState(st);await sleep(500);location.href=next;
  }else{
    st.running=false;
    st.status=`Terminé : ${st.orders||0} commandes, ${st.candidates||0} liens Facture, ${st.downloads||0} vrais PDF gardés, ${st.rejected||0} faux documents rejetés.`;
    await setState(st);
  }
}

chrome.runtime.onMessage.addListener(m=>{
  if(m.type==="V5_START"){
    (async()=>{
      const st=await getState();
      st.running=true;st.year=Number(m.year||2025);st.status=`Démarrage ${st.year}…`;
      await setState(st);
      if(!location.href.includes(`orderFilter=year-${st.year}`))location.href=ordersUrl(st.year);
      else processPage();
    })();
  }
  if(m.type==="V5_STOP"){
    (async()=>{
      const st=await getState();st.running=false;st.status="Arrêté.";await setState(st);
    })();
  }
});

(async()=>{
  const {amazonV5AutoStart}=await chrome.storage.local.get("amazonV5AutoStart");
  const st=await getState();
  if(amazonV5AutoStart){
    await chrome.storage.local.remove("amazonV5AutoStart");
    st.running=true;await setState(st);processPage();
  }else if(st.running){
    processPage();
  }
})();
