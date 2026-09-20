/**
 * Rapprochement AS-Ventes <-> AS-CBC.
 * Ajoute Q:W dans AS-Ventes-20-09-2026 et cree AS-VENTES-CBC-DETAIL.
 * CONFIRME = reference FA/VFA trouvee dans la banque.
 * PROBABLE = nom du payeur reconnu ou montant exact unique.
 */
const VR_CFG = Object.freeze({
  SALES_SHEET: 'AS-Ventes-20-09-2026',
  BANK_SHEET: 'AS-CBC',
  DETAIL_SHEET: 'AS-VENTES-CBC-DETAIL',
  OUTPUT_COL: 17,
  TOL: 0.02,
  NAME_DAYS: 730,
  AMOUNT_DAYS: 365,
  MIN_AMOUNT_ONLY: 50,
  TZ: 'Europe/Brussels'
});

function rapprocherVentesAS() {
  return rapprocherVentesAS_({silent:false});
}

function rapprocherVentesAS_(opt) {
  opt = opt || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ventes = ss.getSheetByName(VR_CFG.SALES_SHEET);
  const banque = ss.getSheetByName(VR_CFG.BANK_SHEET);
  if (!ventes) throw new Error('Onglet introuvable : ' + VR_CFG.SALES_SHEET);
  if (!banque) throw new Error('Onglet introuvable : ' + VR_CFG.BANK_SHEET);

  const vLast = ventes.getLastRow();
  const bLast = banque.getLastRow();
  const vv = ventes.getRange(1,1,vLast,16).getValues();
  const bv = banque.getRange(1,1,bLast,14).getValues();

  const invoices = [];
  const byRef = {};

  for (let i=1;i<vv.length;i++) {
    const r = vv[i];
    const ref = vrRef_(r[0]);
    const date = convertirDate_(r[1]);
    const total = convertirNombre_(r[6]);
    if (!ref || !date || !(total>0)) continue;
    const due = convertirNombre_(r[7]);
    const inv = {
      row:i+1, ref:ref, rawRef:String(r[0]||''), date:date,
      total:arrondir2_(total), due:due===null?0:arrondir2_(due),
      client:String(r[2]||r[8]||('ROW-'+(i+1))).trim(),
      aliases:[r[2],r[3],r[8],r[10]].filter(Boolean),
      confirmed:0, probable:0, balance:arrondir2_(total), alloc:[]
    };
    invoices.push(inv);
    byRef[ref]=inv;
  }

  const clientIndex = vrClientIndex_(invoices);
  const payments = [];

  for (let i=2;i<bv.length;i++) {
    const r=bv[i];
    const date=convertirDate_(r[0]);
    const amount=convertirNombre_(r[8]);
    if (!date || !(amount>0)) continue;
    payments.push({
      row:i+1,date:date,extract:String(r[1]||''),payer:String(r[4]||''),
      desc:String(r[5]||''),comm:String(r[6]||''),
      amount:arrondir2_(amount),remaining:arrondir2_(amount),alloc:[]
    });
  }
  payments.sort(function(a,b){return a.date-b.date || a.row-b.row;});

  // 1. Reference explicite FA/VFA = confirme.
  payments.forEach(function(p){
    const refs=vrUnique_(vrRefs_(p.desc+' '+p.comm)).filter(function(x){return !!byRef[x];});
    refs.forEach(function(ref){
      if (p.remaining<=VR_CFG.TOL) return;
      const inv=byRef[ref];
      if (!inv || inv.balance<=VR_CFG.TOL) return;
      if (inv.date.getTime()>p.date.getTime()+86400000) return;
      vrAlloc_(inv,p,p.remaining,'CONFIRME','Reference '+inv.rawRef+' trouvee dans AS-CBC');
    });
  });

  // 2. Nom du payeur reconnu = probable.
  payments.forEach(function(p){
    if (p.remaining<=VR_CFG.TOL) return;
    const match=vrMatchClient_(p.payer,clientIndex);
    if (!match) return;
    const cut=new Date(p.date.getTime()-VR_CFG.NAME_DAYS*86400000);
    const cand=match.group.invoices.filter(function(inv){
      return inv.balance>VR_CFG.TOL && inv.date<=p.date && inv.date>=cut;
    }).sort(function(a,b){return a.date-b.date || a.row-b.row;});
    if (!cand.length) return;

    const exact=cand.filter(function(inv){return Math.abs(inv.balance-p.remaining)<=VR_CFG.TOL;});
    if (exact.length===1) {
      vrAlloc_(exact[0],p,p.remaining,'PROBABLE','Nom payeur + montant exact');
      return;
    }
    if (cand.length===1 && match.score>=0.82) {
      vrAlloc_(cand[0],p,p.remaining,'PROBABLE','Nom payeur + facture ouverte unique');
      return;
    }
    if (match.score>=0.88) {
      cand.forEach(function(inv){
        if (p.remaining>VR_CFG.TOL) vrAlloc_(inv,p,p.remaining,'PROBABLE','Nom payeur fort + FIFO');
      });
    }
  });

  // 3. Montant exact unique sur une facture recente = probable.
  payments.forEach(function(p){
    if (p.remaining<=VR_CFG.TOL || p.remaining<VR_CFG.MIN_AMOUNT_ONLY) return;
    const cut=new Date(p.date.getTime()-VR_CFG.AMOUNT_DAYS*86400000);
    const cand=invoices.filter(function(inv){
      return inv.balance>VR_CFG.TOL && inv.date<=p.date && inv.date>=cut &&
        Math.abs(inv.balance-p.remaining)<=VR_CFG.TOL;
    });
    if (cand.length===1) vrAlloc_(cand[0],p,p.remaining,'PROBABLE','Montant exact unique');
  });

  const bankStart=payments.length?payments[0].date:null;
  const out=[];
  let paid=0,partial=0,diffs=0;

  for (let i=1;i<vv.length;i++) {
    const ref=vrRef_(vv[i][0]);
    const inv=ref?byRef[ref]:null;
    if (!inv || inv.row!==i+1) {
      out.push(['','','','','','','']);
      continue;
    }

    const conf=arrondir2_(inv.confirmed);
    const prob=arrondir2_(inv.probable);
    const bal=Math.abs(inv.balance)<=VR_CFG.TOL?0:arrondir2_(inv.balance);
    const gap=arrondir2_(bal-inv.due);
    const totalPaid=arrondir2_(conf+prob);
    let status='';

    if (bankStart && inv.date<bankStart && totalPaid<=VR_CFG.TOL) status='⚪ Historique AS-CBC incomplet';
    else if (bal<=VR_CFG.TOL) {status='✅ Payee selon CBC';paid++;}
    else if (totalPaid>VR_CFG.TOL) {status='🔵 Partiellement payee';partial++;}
    else status='❌ Aucun paiement rapproche';

    if (Math.abs(gap)<=VR_CFG.TOL) status+=' — concorde avec Solde Du';
    else if (gap<0) {status+=' — paiement CBC possiblement non comptabilise';diffs++;}
    else {status+=' — paiement manquant/non rapproche a verifier';diffs++;}

    const details=inv.alloc.map(function(a){
      return Utilities.formatDate(a.date,VR_CFG.TZ,'dd/MM/yyyy')+' +'+
        arrondir2_(a.amount).toFixed(2).replace('.',',')+' EUR ['+a.method+'] '+a.reason;
    }).join(' | ');
    const lines=inv.alloc.map(function(a){return 'AS-CBC!'+a.bankRow;}).join(', ');

    out.push([conf,prob,bal,gap,status,details,lines]);
  }

  const headers=['Paiements CBC confirmes','Paiements CBC probables','Solde CBC',
    'Ecart vs Solde Du','Statut CBC','Detail paiements','Lignes AS-CBC'];

  if (ventes.getMaxColumns()<23) ventes.insertColumnsAfter(ventes.getMaxColumns(),23-ventes.getMaxColumns());
  try {ventes.getRange(1,16).copyFormatToRange(ventes,17,23,1,1);} catch(e) {}
  ventes.getRange(1,17,1,7).setValues([headers]).setFontWeight('bold').setWrap(true);
  if (out.length) {
    ventes.getRange(2,17,out.length,7).clearContent().setValues(out);
    ventes.getRange(2,17,out.length,4).setNumberFormat('#,##0.00');
    ventes.getRange(2,21,out.length,3).setWrap(true);
  }
  [125,125,110,115,285,440,190].forEach(function(w,i){ventes.setColumnWidth(17+i,w);});

  vrDetail_(ss,banque,payments);
  SpreadsheetApp.flush();

  const confirmedPayments=payments.filter(function(p){return p.alloc.some(function(a){return a.method==='CONFIRME';});}).length;
  const probablePayments=payments.filter(function(p){return p.alloc.some(function(a){return a.method==='PROBABLE';});}).length;
  const unmatched=payments.filter(function(p){return p.remaining>VR_CFG.TOL;}).length;
  const result={invoices:invoices.length,payments:payments.length,confirmedPayments:confirmedPayments,
    probablePayments:probablePayments,unmatchedPayments:unmatched,paidInvoices:paid,
    partialInvoices:partial,balanceDifferences:diffs};

  if (!opt.silent) ss.toast(
    invoices.length+' factures — '+confirmedPayments+' paiement(s) confirme(s), '+
    probablePayments+' probable(s), '+unmatched+' non rapproche(s).',
    'RAPPROCHEMENT VENTES / CBC',8
  );
  return result;
}

function vrAlloc_(inv,p,wanted,method,reason) {
  const amount=arrondir2_(Math.min(inv.balance,p.remaining,wanted));
  if (!(amount>VR_CFG.TOL)) return 0;
  inv.balance=arrondir2_(inv.balance-amount);
  p.remaining=arrondir2_(p.remaining-amount);
  if (method==='CONFIRME') inv.confirmed=arrondir2_(inv.confirmed+amount);
  else inv.probable=arrondir2_(inv.probable+amount);
  const a={ref:inv.rawRef,salesRow:inv.row,bankRow:p.row,date:p.date,amount:amount,method:method,reason:reason};
  inv.alloc.push(a); p.alloc.push(a);
  return amount;
}

function vrRef_(v) {
  const m=String(v||'').toUpperCase().match(/\b(VFA|FA)\s*[-\/]?\s*0*(\d+)\b/);
  return m?m[1]+String(parseInt(m[2],10)):'';
}

function vrRefs_(text) {
  const out=[],re=/\b(VFA|FA)\s*[-\/]?\s*0*(\d+)\b/g,s=String(text||'').toUpperCase();
  let m; while((m=re.exec(s))!==null) out.push(m[1]+String(parseInt(m[2],10)));
  return out;
}

function vrNorm_(v) {
  return normaliserRecherche_(v||'').replace(/&/g,' et ').replace(/[^a-z0-9]+/g,' ').replace(/\s+/g,' ').trim();
}

function vrTokens_(v) {
  const stop={srl:1,sprl:1,sa:1,nv:1,bv:1,asbl:1,sc:1,scrl:1,mr:1,mme:1,
    monsieur:1,madame:1,et:1,de:1,du:1,des:1,la:1,le:1,les:1,aux:1,
    societe:1,company:1,construct:1,construction:1,constructions:1,maison:1,
    maisons:1,group:1,groupe:1,belgium:1,belgique:1,industrie:1,industry:1};
  return vrNorm_(v).split(' ').filter(function(t){return t.length>=3&&!stop[t]&&!/^\d+$/.test(t);});
}

function vrClientIndex_(invoices) {
  const groups={};
  invoices.forEach(function(inv){
    const key=inv.client||('ROW-'+inv.row);
    if (!groups[key]) groups[key]={key:key,invoices:[],aliases:{},tokens:{}};
    const g=groups[key]; g.invoices.push(inv);
    inv.aliases.forEach(function(a){
      const n=vrNorm_(a); if(n) g.aliases[n]=true;
      vrTokens_(a).forEach(function(t){g.tokens[t]=true;});
    });
  });
  const tokenGroups={};
  Object.keys(groups).forEach(function(k){
    Object.keys(groups[k].tokens).forEach(function(t){
      if(!tokenGroups[t]) tokenGroups[t]=[];
      tokenGroups[t].push(groups[k]);
    });
  });
  return {groups:groups,tokenGroups:tokenGroups};
}

function vrMatchClient_(payer,index) {
  const pn=vrNorm_(payer),pt=vrUnique_(vrTokens_(payer));
  if(!pn||!pt.length) return null;
  const cand={};
  pt.forEach(function(t){(index.tokenGroups[t]||[]).forEach(function(g){cand[g.key]=g;});});
  let best=null,score=0,tied=false;
  Object.keys(cand).forEach(function(k){
    const g=cand[k],gt=Object.keys(g.tokens);
    let overlap=0,uniqueHits=0;
    pt.forEach(function(t){
      if(!g.tokens[t]) return;
      overlap++;
      if((index.tokenGroups[t]||[]).length===1&&t.length>=4) uniqueHits++;
    });
    let s=overlap/Math.max(1,Math.min(pt.length,gt.length));
    Object.keys(g.aliases).forEach(function(a){
      if(a.length>=5&&(pn.indexOf(a)!==-1||a.indexOf(pn)!==-1)) s=Math.max(s,1);
    });
    if(overlap>=2) s=Math.max(s,0.90);
    if(uniqueHits>=1) s=Math.max(s,0.82+Math.min(0.08,(uniqueHits-1)*0.04));
    if(s>score+0.000001){best=g;score=s;tied=false;}
    else if(s>0&&Math.abs(s-score)<=0.000001) tied=true;
  });
  return (!best||tied||score<0.82)?null:{group:best,score:score};
}

function vrUnique_(arr) {
  const seen={};
  return (arr||[]).filter(function(x){const k=String(x);if(seen[k])return false;seen[k]=true;return true;});
}

function vrDetail_(ss,bank,payments) {
  let sh=ss.getSheetByName(VR_CFG.DETAIL_SHEET);
  if(!sh) sh=ss.insertSheet(VR_CFG.DETAIL_SHEET);
  const f=sh.getFilter(); if(f) f.remove();
  sh.clear();
  if(sh.getMaxRows()<payments.length+5) sh.insertRowsAfter(sh.getMaxRows(),payments.length+5-sh.getMaxRows());
  if(sh.getMaxColumns()<13) sh.insertColumnsAfter(sh.getMaxColumns(),13-sh.getMaxColumns());

  const h=['Date','Ligne AS-CBC','N° extrait','Payeur','Communication','Montant recu',
    'Facture(s)','Confirme','Probable','Non affecte','Methode','Statut','Ouvrir AS-CBC'];
  sh.getRange(1,1,1,13).setValues([h]).setBackground('#5f6368').setFontColor('#fff').setFontWeight('bold');

  const rows=payments.map(function(p){
    const refs=vrUnique_(p.alloc.map(function(a){return a.ref;})).join(', ');
    const conf=arrondir2_(p.alloc.filter(function(a){return a.method==='CONFIRME';}).reduce(function(s,a){return s+a.amount;},0));
    const prob=arrondir2_(p.alloc.filter(function(a){return a.method==='PROBABLE';}).reduce(function(s,a){return s+a.amount;},0));
    const meth=vrUnique_(p.alloc.map(function(a){return a.method+' — '+a.reason;})).join(' | ');
    let st='❌ Non rapproche';
    if(p.remaining<=VR_CFG.TOL&&conf>0&&prob<=VR_CFG.TOL) st='✅ Rapproche confirme';
    else if(p.remaining<=VR_CFG.TOL&&prob>0) st='🟡 Rapproche avec probable';
    else if(p.alloc.length) st='🔵 Partiellement rapproche';
    return [p.date,p.row,p.extract,p.payer,p.comm,p.amount,refs,conf,prob,arrondir2_(p.remaining),meth,st,''];
  });
  if(rows.length){
    sh.getRange(2,1,rows.length,13).setValues(rows);
    sh.getRange(2,1,rows.length,1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(2,6,rows.length,1).setNumberFormat('#,##0.00');
    sh.getRange(2,8,rows.length,3).setNumberFormat('#,##0.00');
    sh.getRange(2,4,rows.length,2).setWrap(true);
    sh.getRange(2,11,rows.length,2).setWrap(true);
    const base=ss.getUrl().replace(/#.*$/,'');
    const rich=payments.map(function(p){return [SpreadsheetApp.newRichTextValue().setText('Ouvrir')
      .setLinkUrl(base+'#gid='+bank.getSheetId()+'&range=A'+p.row).build()];});
    sh.getRange(2,13,rich.length,1).setRichTextValues(rich).setHorizontalAlignment('center');
    sh.getRange(1,1,rows.length+1,13).createFilter();
  }
  [95,85,95,190,330,110,150,100,100,100,310,190,90].forEach(function(w,i){sh.setColumnWidth(i+1,w);});
  sh.setFrozenRows(1);
}
