
const $=id=>document.getElementById(id);

async function getState(){
  const {amazonV5State}=await chrome.storage.local.get("amazonV5State");
  return amazonV5State||{
    running:false,year:2025,pages:0,orders:0,candidates:0,downloads:0,
    rejected:0,missing:0,errors:0,status:"Prêt.",rows:[],visited:[]
  };
}
async function saveState(s){await chrome.storage.local.set({amazonV5State:s});}
async function activeTab(){return (await chrome.tabs.query({active:true,currentWindow:true}))[0];}
const ordersUrl=y=>`https://www.amazon.fr/gp/your-account/order-history?orderFilter=year-${encodeURIComponent(y)}`;

async function render(){
  const s=await getState();
  $("year").value=s.year||2025;
  $("orders").textContent=s.orders||0;
  $("candidates").textContent=s.candidates||0;
  $("downloads").textContent=s.downloads||0;
  $("rejected").textContent=s.rejected||0;
  $("status").textContent=s.status||"Prêt.";
}

$("open").onclick=async()=>{
  const y=Number($("year").value||2025),t=await activeTab(),s=await getState();
  s.year=y;s.status=`Ouverture des commandes ${y}…`;await saveState(s);
  await chrome.tabs.update(t.id,{url:ordersUrl(y)});window.close();
};

$("start").onclick=async()=>{
  const y=Number($("year").value||2025);
  await saveState({
    running:true,year:y,pages:0,orders:0,candidates:0,downloads:0,
    rejected:0,missing:0,errors:0,status:`Démarrage ${y}…`,rows:[],visited:[]
  });
  const t=await activeTab();
  if(!t?.url||!t.url.includes("amazon.fr")||!t.url.includes("order")){
    await chrome.storage.local.set({amazonV5AutoStart:true});
    await chrome.tabs.update(t.id,{url:ordersUrl(y)});
  }else{
    try{await chrome.tabs.sendMessage(t.id,{type:"V5_START",year:y});}
    catch{
      await chrome.storage.local.set({amazonV5AutoStart:true});
      await chrome.tabs.update(t.id,{url:ordersUrl(y)});
    }
  }
  window.close();
};

$("stop").onclick=async()=>{
  const s=await getState();s.running=false;s.status="Arrêt demandé.";await saveState(s);
  const t=await activeTab();if(t?.id)chrome.tabs.sendMessage(t.id,{type:"V5_STOP"}).catch(()=>{});
  render();
};

$("report").onclick=()=>chrome.runtime.sendMessage({type:"V5_EXPORT"});
$("reset").onclick=async()=>{
  const y=Number($("year").value||2025);
  await saveState({
    running:false,year:y,pages:0,orders:0,candidates:0,downloads:0,
    rejected:0,missing:0,errors:0,status:"Compteurs remis à zéro.",rows:[],visited:[]
  });
  render();
};

chrome.storage.onChanged.addListener(c=>{if(c.amazonV5State)render();});
render();
