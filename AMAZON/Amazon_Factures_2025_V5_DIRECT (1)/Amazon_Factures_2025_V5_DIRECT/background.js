
async function getState(){
  const {amazonV5State}=await chrome.storage.local.get("amazonV5State");
  return amazonV5State||{};
}
async function setState(s){await chrome.storage.local.set({amazonV5State:s});}
function safe(s){
  return String(s||"").replace(/[<>:"/\\|?*\x00-\x1F]/g,"_").replace(/\s+/g," ").trim().slice(0,100);
}

function waitDownload(id,timeout=45000){
  return new Promise((resolve,reject)=>{
    const timer=setTimeout(async()=>{
      chrome.downloads.onChanged.removeListener(listener);
      try{await chrome.downloads.cancel(id);}catch{}
      reject(new Error("Délai téléchargement dépassé"));
    },timeout);

    async function listener(d){
      if(d.id!==id)return;
      if(d.state?.current==="complete"){
        clearTimeout(timer);chrome.downloads.onChanged.removeListener(listener);
        const items=await chrome.downloads.search({id});
        resolve(items[0]||null);
      }else if(d.state?.current==="interrupted"){
        clearTimeout(timer);chrome.downloads.onChanged.removeListener(listener);
        const items=await chrome.downloads.search({id});
        reject(new Error(items[0]?.error||"Téléchargement interrompu"));
      }
    }
    chrome.downloads.onChanged.addListener(listener);
  });
}

function looksLikeRealPdf(item){
  const mime=String(item?.mime||"").toLowerCase();
  const finalUrl=String(item?.finalUrl||"").toLowerCase();
  const size=Number(item?.fileSize||item?.totalBytes||0);

  if(mime.includes("application/pdf")) return {ok:true,why:`MIME ${mime}`};

  // Certains serveurs renvoient les PDF comme binaire générique.
  // On ne les accepte que si l'URL finale ressemble clairement à un téléchargement de document.
  if(
    (mime.includes("application/octet-stream") || mime==="") &&
    size>10000 &&
    (
      finalUrl.includes(".pdf") ||
      finalUrl.includes("invoice") ||
      finalUrl.includes("document") ||
      finalUrl.includes("download")
    )
  ){
    return {ok:true,why:`binaire ${mime||"sans MIME"}, ${size} octets`};
  }

  return {ok:false,why:`MIME=${mime||"inconnu"}, taille=${size}, finalUrl=${finalUrl.slice(0,140)}`};
}

chrome.runtime.onMessage.addListener((m,sender,send)=>{
  if(m.type==="V5_TRY_DOWNLOAD"){
    (async()=>{
      let id=null;
      try{
        const st=await getState();
        const year=st.year||2025;
        const oid=safe(m.orderId||"commande");
        const idx=Number(m.index||1);
        const filename=`Amazon_Factures_${year}/${oid}_FACTURE_${String(idx).padStart(2,"0")}.pdf`;

        id=await chrome.downloads.download({
          url:m.url,
          filename,
          conflictAction:"uniquify",
          saveAs:false
        });

        const item=await waitDownload(id,45000);
        const check=looksLikeRealPdf(item);

        if(!check.ok){
          try{await chrome.downloads.removeFile(id);}catch{}
          try{await chrome.downloads.erase({id});}catch{}
          const f=await getState();
          f.rejected=(f.rejected||0)+1;
          f.status=`Rejeté ${oid} : Amazon a renvoyé HTML/non-PDF.`;
          await setState(f);
          send({ok:false,rejected:true,reason:check.why});
          return;
        }

        const f=await getState();
        f.downloads=(f.downloads||0)+1;
        f.status=`PDF OK : ${oid} — ${m.label||"Facture"}`;
        await setState(f);
        send({
          ok:true,
          mime:item?.mime||"",
          bytes:item?.fileSize||item?.totalBytes||0,
          finalUrl:item?.finalUrl||"",
          reason:check.why
        });

      }catch(e){
        if(id!==null){
          try{await chrome.downloads.removeFile(id);}catch{}
        }
        const f=await getState();
        f.errors=(f.errors||0)+1;
        f.status=`Erreur : ${e.message}. Passage à la suite…`;
        await setState(f);
        send({ok:false,error:e.message});
      }
    })();
    return true;
  }

  if(m.type==="V5_EXPORT"){
    (async()=>{
      const st=await getState(),rows=Array.isArray(st.rows)?st.rows:[];
      const esc=v=>'"'+String(v??"").replaceAll('"','""')+'"';
      const header=[
        "annee","date","numero_commande","montant","statut",
        "liens_facture","pdf_gardes","rejetes","details"
      ];
      const lines=[header.map(esc).join(";")];
      for(const r of rows){
        lines.push([
          st.year,r.date,r.orderId,r.total,r.status,
          r.candidates,r.downloaded,r.rejected,r.details
        ].map(esc).join(";"));
      }
      const csv="\ufeff"+lines.join("\r\n");
      const id=await chrome.downloads.download({
        url:"data:text/csv;charset=utf-8,"+encodeURIComponent(csv),
        filename:`Amazon_Factures_${st.year||2025}/rapport_V5_${st.year||2025}.csv`,
        conflictAction:"overwrite",
        saveAs:false
      });
      send({ok:true,id});
    })().catch(e=>send({ok:false,error:e.message}));
    return true;
  }
});
