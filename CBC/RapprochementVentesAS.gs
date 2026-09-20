/**
 * RAPPROCHEMENT VENTES AS <-> AS-CBC
 * ---------------------------------
 * Source ventes : AS-Ventes-20-09-2026
 * Source banque  : AS-CBC
 *
 * Objectif :
 * - rapprocher les encaissements positifs de AS-CBC avec les factures de vente ;
 * - NE PAS modifier la colonne H "Solde Dû" existante ;
 * - ajouter Q:W pour un contrôle indépendant :
 *     Q Paiements CBC
 *     R Solde CBC
 *     S Écart vs Solde Dû
 *     T Statut CBC
 *     U Méthode
 *     V Paiements liés
 *     W Dernière MAJ
 * - créer/mettre à jour RAPPROCHEMENT-VENTES-CBC pour les paiements non affectés
 *   ou ambigus.
 *
 * Le rapprochement automatique est volontairement prudent :
 * 1) référence VFA explicite dans la banque = priorité maximale ;
 * 2) montant exact égal au solde d'une seule facture = automatique ;
 * 3) nom client suffisamment proche + une seule facture compatible = automatique ;
 * 4) sinon le paiement est envoyé au contrôle manuel.
 */

const AS_VENTES_CBC_CFG = Object.freeze({
  SALES_SHEET: 'AS-Ventes-20-09-2026',
  BANK_SHEET: 'AS-CBC',
  CONTROL_SHEET: 'RAPPROCHEMENT-VENTES-CBC',

  SALES_HEADER_ROW: 1,
  SALES_FIRST_DATA_ROW: 2,
  SALES_BASE_COLS: 16, // A:P
  SALES_COL_DOC: 1,    // A N° pièce
  SALES_COL_DATE: 2,   // B Date Pièce
  SALES_COL_CLIENT: 3, // C Client
  SALES_COL_COMPANY: 4,// D Société
  SALES_COL_TTC: 7,    // G Total TTC
  SALES_COL_OLD_BAL: 8,// H Solde Dû
  SALES_COL_NAME: 11,  // K Nom

  BANK_HEADER_ROW: 1,
  BANK_FIRST_DATA_ROW: 3,
  BANK_COL_DATE: 1,    // A
  BANK_COL_COUNTERPARTY: 5, // E
  BANK_COL_DESC: 6,    // F
  BANK_COL_COMM: 7,    // G
  BANK_COL_AMOUNT: 9,  // I

  OUT_FIRST_COL: 17,   // Q
  OUT_WIDTH: 7,        // Q:W
  EPS: 0.02,
  NAME_MIN_SCORE: 0.56,
  MAX_DAYS_AFTER_INVOICE: 1825
});

function rapprocherVentesASCBC() {
  const cfg = AS_VENTES_CBC_CFG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSh = ss.getSheetByName(cfg.SALES_SHEET);
  const bankSh = ss.getSheetByName(cfg.BANK_SHEET);
  if (!salesSh) throw new Error('Onglet introuvable : ' + cfg.SALES_SHEET);
  if (!bankSh) throw new Error('Onglet introuvable : ' + cfg.BANK_SHEET);

  const salesLastRow = salesSh.getLastRow();
  const bankLastRow = bankSh.getLastRow();
  if (salesLastRow < cfg.SALES_FIRST_DATA_ROW) throw new Error('Aucune vente à rapprocher.');
  if (bankLastRow < cfg.BANK_FIRST_DATA_ROW) throw new Error('Aucun mouvement bancaire à rapprocher.');

  const salesValues = salesSh
    .getRange(cfg.SALES_FIRST_DATA_ROW, 1, salesLastRow - cfg.SALES_FIRST_DATA_ROW + 1, cfg.SALES_BASE_COLS)
    .getValues();

  const bankValues = bankSh
    .getRange(cfg.BANK_FIRST_DATA_ROW, 1, bankLastRow - cfg.BANK_FIRST_DATA_ROW + 1, 14)
    .getValues();

  const invoices = [];
  const invoiceByKey = new Map();

  salesValues.forEach((row, i) => {
    const sheetRow = cfg.SALES_FIRST_DATA_ROW + i;
    const rawDoc = row[cfg.SALES_COL_DOC - 1];
    const doc = String(rawDoc || '').trim();
    const amount = asNumberVentes_(row[cfg.SALES_COL_TTC - 1]);
    const date = asDateVentes_(row[cfg.SALES_COL_DATE - 1]);

    if (!doc || !isFinite(amount) || amount <= cfg.EPS || !date) return;
    if (/^totaux?$/i.test(doc)) return;

    const key = canonicalPieceVentes_(doc);
    const inv = {
      row: sheetRow,
      key,
      doc,
      date,
      client: String(row[cfg.SALES_COL_CLIENT - 1] || ''),
      company: String(row[cfg.SALES_COL_COMPANY - 1] || ''),
      name: String(row[cfg.SALES_COL_NAME - 1] || ''),
      total: round2Ventes_(amount),
      oldBalance: nullableNumberVentes_(row[cfg.SALES_COL_OLD_BAL - 1]),
      paid: 0,
      payments: [],
      methods: new Set(),
      confidence: 'Aucun'
    };

    invoices.push(inv);
    if (key) {
      if (!invoiceByKey.has(key)) invoiceByKey.set(key, []);
      invoiceByKey.get(key).push(inv);
    }
  });

  invoices.sort((a, b) => a.date - b.date || a.row - b.row);

  const incoming = [];
  bankValues.forEach((row, i) => {
    const amount = asNumberVentes_(row[cfg.BANK_COL_AMOUNT - 1]);
    const date = asDateVentes_(row[cfg.BANK_COL_DATE - 1]);
    if (!date || !isFinite(amount) || amount <= cfg.EPS) return;

    incoming.push({
      bankRow: cfg.BANK_FIRST_DATA_ROW + i,
      date,
      counterparty: String(row[cfg.BANK_COL_COUNTERPARTY - 1] || ''),
      desc: String(row[cfg.BANK_COL_DESC - 1] || ''),
      comm: String(row[cfg.BANK_COL_COMM - 1] || ''),
      amount: round2Ventes_(amount),
      allocated: false,
      reason: '',
      candidates: []
    });
  });

  incoming.sort((a, b) => a.date - b.date || a.bankRow - b.bankRow);

  // 1. Références VFA explicites.
  incoming.forEach(p => {
    if (p.allocated) return;
    const refs = extractVfaRefsVentes_(p.desc + ' ' + p.comm);
    if (!refs.length) return;

    const candidates = [];
    refs.forEach(ref => {
      const list = invoiceByKey.get(ref) || [];
      list.forEach(inv => {
        if (isInvoiceDateCompatibleVentes_(inv, p, cfg)) candidates.push(inv);
      });
    });

    const unique = uniqueInvoicesVentes_(candidates);
    if (unique.length === 1) {
      allocatePaymentVentes_(unique[0], p, p.amount, 'Référence VFA', 'Très haute');
      return;
    }

    // Plusieurs références : n'affecter automatiquement que si une seule facture
    // a exactement le montant restant du paiement.
    const exact = unique.filter(inv => Math.abs(outstandingVentes_(inv) - p.amount) <= cfg.EPS);
    if (exact.length === 1) {
      allocatePaymentVentes_(exact[0], p, p.amount, 'Référence VFA + montant', 'Très haute');
      return;
    }

    p.reason = 'Référence VFA ambiguë';
    p.candidates = unique.map(x => x.doc);
  });

  // 2. Montant exact unique sur une facture encore ouverte.
  incoming.forEach(p => {
    if (p.allocated) return;

    const exact = invoices.filter(inv =>
      isInvoiceDateCompatibleVentes_(inv, p, cfg) &&
      outstandingVentes_(inv) > cfg.EPS &&
      Math.abs(outstandingVentes_(inv) - p.amount) <= cfg.EPS
    );

    if (exact.length === 1) {
      allocatePaymentVentes_(exact[0], p, p.amount, 'Montant exact unique', 'Haute');
    } else if (exact.length > 1 && !p.reason) {
      p.reason = 'Montant exact présent sur plusieurs factures';
      p.candidates = exact.slice(0, 10).map(x => x.doc);
    }
  });

  // 3. Nom client + compatibilité montant.
  incoming.forEach(p => {
    if (p.allocated) return;

    const scored = invoices
      .filter(inv =>
        isInvoiceDateCompatibleVentes_(inv, p, cfg) &&
        outstandingVentes_(inv) > cfg.EPS &&
        p.amount <= outstandingVentes_(inv) + cfg.EPS
      )
      .map(inv => ({
        inv,
        score: bestNameScoreVentes_(p.counterparty, [inv.client, inv.company, inv.name])
      }))
      .filter(x => x.score >= cfg.NAME_MIN_SCORE)
      .sort((a, b) => b.score - a.score || a.inv.date - b.inv.date);

    if (!scored.length) return;

    // S'il n'existe qu'une facture crédible pour ce client, on accepte un acompte.
    const best = scored[0];
    const second = scored[1];

    if (!second || best.score >= second.score + 0.14) {
      allocatePaymentVentes_(
        best.inv,
        p,
        p.amount,
        Math.abs(outstandingVentes_(best.inv) - p.amount) <= cfg.EPS
          ? 'Client + montant'
          : 'Client + acompte',
        best.score >= 0.75 ? 'Haute' : 'Moyenne'
      );
    } else {
      p.reason = p.reason || 'Plusieurs factures possibles pour le même client';
      p.candidates = scored.slice(0, 8).map(x => x.inv.doc);
    }
  });

  // Prépare les sorties Q:W en respectant les lignes existantes.
  const out = Array.from({ length: salesLastRow - cfg.SALES_FIRST_DATA_ROW + 1 }, () =>
    ['', '', '', '', '', '', '']
  );

  const now = new Date();

  invoices.forEach(inv => {
    const idx = inv.row - cfg.SALES_FIRST_DATA_ROW;
    const paid = round2Ventes_(inv.paid);
    const balance = round2Ventes_(inv.total - paid);
    const oldBalance = inv.oldBalance;
    const diff = oldBalance === null ? '' : round2Ventes_(balance - oldBalance);

    let status;
    if (Math.abs(balance) <= cfg.EPS) status = 'PAYÉ';
    else if (balance < -cfg.EPS) status = 'SURPAYÉ';
    else if (paid > cfg.EPS) status = 'PARTIEL';
    else status = 'NON RAPPROCHÉ';

    const methods = Array.from(inv.methods).join(' + ');
    const detail = inv.payments
      .map(x =>
        Utilities.formatDate(x.date, cfgTimeZoneVentes_(), 'dd/MM/yyyy') +
        ' | ' + formatMoneyVentes_(x.amount) +
        ' | AS-CBC ligne ' + x.bankRow +
        (x.counterparty ? ' | ' + x.counterparty : '')
      )
      .join('\n');

    out[idx] = [
      paid,
      balance,
      diff,
      status,
      methods,
      detail,
      now
    ];
  });

  const header = [[
    'Paiements CBC',
    'Solde CBC',
    'Écart vs Solde Dû',
    'Statut CBC',
    'Méthode rapprochement',
    'Paiements liés AS-CBC',
    'Dernière MAJ'
  ]];

  salesSh.getRange(cfg.SALES_HEADER_ROW, cfg.OUT_FIRST_COL, 1, cfg.OUT_WIDTH)
    .setValues(header)
    .setFontWeight('bold');

  const outputRange = salesSh.getRange(
    cfg.SALES_FIRST_DATA_ROW,
    cfg.OUT_FIRST_COL,
    out.length,
    cfg.OUT_WIDTH
  );
  outputRange.clearContent();
  outputRange.setValues(out);

  // Formats : Q:R:S en montant, W en date/heure.
  salesSh.getRange(cfg.SALES_FIRST_DATA_ROW, 17, out.length, 3).setNumberFormat('#,##0.00');
  salesSh.getRange(cfg.SALES_FIRST_DATA_ROW, 23, out.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  salesSh.getRange(cfg.SALES_FIRST_DATA_ROW, 22, out.length, 1).setWrap(true);

  // Feuille de contrôle.
  writeControlSheetVentes_(ss, incoming, invoices, cfg);

  SpreadsheetApp.flush();

  const allocatedCount = incoming.filter(x => x.allocated).length;
  const unmatchedCount = incoming.length - allocatedCount;
  const fullyPaid = invoices.filter(x => Math.abs(outstandingVentes_(x)) <= cfg.EPS).length;
  const partial = invoices.filter(x => x.paid > cfg.EPS && outstandingVentes_(x) > cfg.EPS).length;

  SpreadsheetApp.getUi().alert(
    'Rapprochement ventes / AS-CBC terminé',
    [
      'Factures analysées : ' + invoices.length,
      'Encaissements analysés : ' + incoming.length,
      'Encaissements rapprochés : ' + allocatedCount,
      'Encaissements à vérifier : ' + unmatchedCount,
      'Factures payées : ' + fullyPaid,
      'Factures partielles : ' + partial,
      '',
      'La colonne H "Solde Dû" n\'a pas été modifiée.',
      'Le nouveau solde indépendant est en colonne R "Solde CBC".',
      'Les cas non sûrs sont dans "' + cfg.CONTROL_SHEET + '".'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function writeControlSheetVentes_(ss, incoming, invoices, cfg) {
  let sh = ss.getSheetByName(cfg.CONTROL_SHEET);
  if (!sh) sh = ss.insertSheet(cfg.CONTROL_SHEET);

  sh.clearContents();

  const headers = [[
    'Ligne AS-CBC',
    'Date',
    'Contrepartie',
    'Montant',
    'Communication / Description',
    'Motif',
    'Factures candidates'
  ]];

  const rows = incoming
    .filter(p => !p.allocated)
    .map(p => [
      p.bankRow,
      p.date,
      p.counterparty,
      p.amount,
      (p.comm || p.desc).slice(0, 1000),
      p.reason || 'Aucune correspondance suffisamment sûre',
      (p.candidates || []).join(', ')
    ]);

  sh.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');

  if (rows.length) {
    sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 4, rows.length, 1).setNumberFormat('#,##0.00');
    sh.getRange(2, 5, rows.length, 1).setWrap(true);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 4);
  sh.setColumnWidth(5, 500);
  sh.setColumnWidth(6, 260);
  sh.setColumnWidth(7, 280);
}

function allocatePaymentVentes_(inv, p, amount, method, confidence) {
  if (p.allocated) return;

  inv.paid = round2Ventes_(inv.paid + amount);
  inv.methods.add(method);
  inv.payments.push({
    bankRow: p.bankRow,
    date: p.date,
    amount,
    counterparty: p.counterparty,
    method
  });
  inv.confidence = confidence;
  p.allocated = true;
  p.reason = '';
  p.candidates = [inv.doc];
}

function outstandingVentes_(inv) {
  return round2Ventes_(inv.total - inv.paid);
}

function isInvoiceDateCompatibleVentes_(inv, payment, cfg) {
  const deltaDays = Math.floor((payment.date.getTime() - inv.date.getTime()) / 86400000);
  return deltaDays >= -2 && deltaDays <= cfg.MAX_DAYS_AFTER_INVOICE;
}

function extractVfaRefsVentes_(text) {
  const s = normalizeTextVentes_(text).toUpperCase();
  const refs = new Set();

  // VFA4690 / VFA 4690 / VFA-4690 / VFA 004690
  const re = /\bVFA\s*[-./]?\s*0*(\d{1,7})\b/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    refs.add('VFA' + String(parseInt(m[1], 10)));
  }
  return Array.from(refs);
}

function canonicalPieceVentes_(value) {
  const s = normalizeTextVentes_(value).toUpperCase().replace(/\s+/g, '');
  let m = s.match(/^VFA[-./]?0*(\d{1,7})$/);
  if (m) return 'VFA' + String(parseInt(m[1], 10));
  return s;
}

function bestNameScoreVentes_(bankName, invoiceNames) {
  let best = 0;
  invoiceNames.forEach(name => {
    best = Math.max(best, nameScoreVentes_(bankName, name));
  });
  return best;
}

function nameScoreVentes_(a, b) {
  const aa = tokenizeNameVentes_(a);
  const bb = tokenizeNameVentes_(b);
  if (!aa.length || !bb.length) return 0;

  const setA = new Set(aa);
  const setB = new Set(bb);
  let common = 0;
  setA.forEach(t => { if (setB.has(t)) common++; });

  const union = new Set([...setA, ...setB]).size;
  let score = union ? common / union : 0;

  const sa = aa.join('');
  const sb = bb.join('');
  if (sa && sb && (sa.includes(sb) || sb.includes(sa))) score = Math.max(score, 0.82);

  // Bonus si deux mots significatifs identiques.
  if (common >= 2) score = Math.max(score, 0.72);
  return score;
}

function tokenizeNameVentes_(value) {
  const stop = new Set([
    'SA','SRL','SPRL','NV','BV','SOCIETE','SOCIETE','CONSTRUCTION','CONSTRUCTIONS',
    'ENTREPRISE','ENTREPRISES','MONSIEUR','MADAME','MME','MR','M','DE','DU','DES',
    'LA','LE','LES','ET','THE','AS'
  ]);

  return normalizeTextVentes_(value)
    .toUpperCase()
    .replace(/[^A-Z0-9 ]+/g, ' ')
    .split(/\s+/)
    .filter(x => x.length >= 3 && !stop.has(x));
}

function normalizeTextVentes_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function uniqueInvoicesVentes_(arr) {
  const seen = new Set();
  return arr.filter(x => {
    if (seen.has(x.row)) return false;
    seen.add(x.row);
    return true;
  });
}

function nullableNumberVentes_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return null;
  const n = asNumberVentes_(value);
  return isFinite(n) ? round2Ventes_(n) : null;
}

function asNumberVentes_(value) {
  if (typeof value === 'number') return value;
  if (value === null || typeof value === 'undefined' || value === '') return NaN;
  let s = String(value)
    .replace(/\u00A0|\u202F/g, '')
    .replace(/\s/g, '')
    .replace(/€/g, '')
    .replace(/,/g, '.');
  // Si plusieurs points subsistent, les premiers sont probablement des séparateurs de milliers.
  const parts = s.split('.');
  if (parts.length > 2) {
    const dec = parts.pop();
    s = parts.join('') + '.' + dec;
  }
  return Number(s);
}

function asDateVentes_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'number' && isFinite(value)) {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }
  const s = String(value || '').trim();
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (!m) return null;
  let y = Number(m[3]);
  if (y < 100) y += y >= 70 ? 1900 : 2000;
  const d = new Date(y, Number(m[2]) - 1, Number(m[1]));
  return isNaN(d.getTime()) ? null : d;
}

function round2Ventes_(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

function formatMoneyVentes_(n) {
  return round2Ventes_(n).toFixed(2).replace('.', ',') + ' €';
}

function cfgTimeZoneVentes_() {
  return Session.getScriptTimeZone() || 'Europe/Brussels';
}

/**
 * À lancer une fois si tu veux un menu séparé "AS VENTES".
 * Le script crée aussi un déclencheur à l'ouverture du classeur.
 */
function installerRapprochementVentesAS() {
  ajouterMenuRapprochementVentesAS_();

  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'ajouterMenuRapprochementVentesAS_'
  );

  if (!exists) {
    ScriptApp.newTrigger('ajouterMenuRapprochementVentesAS_')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
  }

  SpreadsheetApp.getUi().alert(
    'Rapprochement ventes installé',
    'Le menu "AS VENTES" a été ajouté. Utilise "Rapprocher avec AS-CBC" pour recalculer les soldes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function ajouterMenuRapprochementVentesAS_() {
  SpreadsheetApp.getUi()
    .createMenu('AS VENTES')
    .addItem('Rapprocher avec AS-CBC', 'rapprocherVentesASCBC')
    .addToUi();
}
