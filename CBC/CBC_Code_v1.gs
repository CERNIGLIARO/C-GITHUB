/**
 * CBC BANK MANAGER - Google Sheets
 * Conversion de la logique principale de CBC-2025.xlsm vers Apps Script.
 *
 * IMPORTANT :
 * - Les CSV CBC doivent être déposés dans un dossier Google Drive.
 * - Le script prend le fichier export_BE...csv le plus récent pour chaque compte.
 * - Il ne supprime pas l'historique des onglets société.
 * - Il reconstruit DIVERS-BANQUE à partir des derniers CSV et ajoute seulement les
 *   opérations absentes dans les onglets société.
 */

const CBC = Object.freeze({
  VERSION: '1.0.0',
  PROP_FOLDER_ID: 'CBC_IMPORT_FOLDER_ID',
  PROP_LAST_BACKUP: 'CBC_LAST_BACKUP_DATE',
  PROP_SPREADSHEET_ID: 'CBC_SPREADSHEET_ID',
  DIVERS_SHEET: 'DIVERS-BANQUE',
  TOTAUX_SHEET: 'TOTAUX',
  JOURNAL_SHEET: 'CBC-JOURNAL',
  DATA_START_ROW: 3,
  DATA_WIDTH: 14, // A:N dans les onglets banque
  DIVERS_WIDTH: 18, // A:R
  ACCOUNTS: [
    { iban: 'BE07732021748966', sheet: 'AS-CBC',       totauxRow: 6,  totalCorrection: -13.98 },
    { iban: 'BE24732066860838', sheet: 'WINPRO-CBC',   totauxRow: 8,  totalCorrection: 0 },
    { iban: 'BE87732049241594', sheet: 'EMC-CBC',      totauxRow: 10, totalCorrection: 0 },
    { iban: 'BE05732042517575', sheet: 'IMMO-ACS-CBC', totauxRow: 12, totalCorrection: 0 },
    { iban: 'BE09732046046557', sheet: 'MC-ALFA-CBC',  totauxRow: 14, totalCorrection: 0 },
    { iban: 'BE29732652200264', sheet: 'ALFANO-CBC',   totauxRow: 20, totalCorrection: 0 },
    { iban: 'BE52732652078309', sheet: 'TONI-CBC',     totauxRow: 22, totalCorrection: 0 }
  ]
});

function onOpen() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) PropertiesService.getDocumentProperties().setProperty(CBC.PROP_SPREADSHEET_ID, active.getId());
  SpreadsheetApp.getUi()
    .createMenu('CBC BANQUE')
    .addItem('Mettre à jour maintenant', 'miseAJourCBC')
    .addSeparator()
    .addItem('Configurer le dossier CSV', 'configurerDossierImports')
    .addItem('Vérifier la configuration', 'verifierConfiguration')
    .addSeparator()
    .addItem('Recalculer les totaux', 'recalculerTotauxCBC')
    .addItem('Créer une sauvegarde', 'creerSauvegardeCBC')
    .addSeparator()
    .addItem('Activer mise à jour auto (1 h)', 'activerMiseAJourAuto')
    .addItem('Désactiver mise à jour auto', 'desactiverMiseAJourAuto')
    .addToUi();
}

/** À lancer une seule fois depuis Apps Script si le menu n'apparaît pas. */
function installerCBC() {
  onOpen();
  SpreadsheetApp.getUi().alert(
    'CBC BANK MANAGER installé',
    'Le menu "CBC BANQUE" est maintenant disponible dans le classeur.\n\n' +
    'Étape suivante : CBC BANQUE > Configurer le dossier CSV.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function configurerDossierImports() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Dossier des exports CBC',
    'Colle ici l’URL complète du dossier Google Drive "CBC - IMPORTS BANCAIRES" (ou simplement son ID).',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const folderId = extraireIdDrive_(response.getResponseText());
  if (!folderId) {
    ui.alert('Impossible de reconnaître l’ID du dossier. Copie l’URL du dossier depuis Google Drive.');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    PropertiesService.getDocumentProperties().setProperty(CBC.PROP_FOLDER_ID, folderId);
    ui.alert(
      'Configuration enregistrée',
      'Dossier : ' + folder.getName() + '\n\n' +
      'Tu peux maintenant déposer les fichiers export_BE....csv dans ce dossier puis cliquer sur "Mettre à jour maintenant".',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erreur', 'Je ne peux pas ouvrir ce dossier Drive :\n' + e.message, ui.ButtonSet.OK);
  }
}

function verifierConfiguration() {
  const ui = SpreadsheetApp.getUi();
  try {
    const folder = getImportFolder_();
    const latest = getLatestCsvFiles_(folder);
    const lines = CBC.ACCOUNTS.map(a => {
      const f = latest[a.iban];
      return (f ? '✓ ' : '✗ ') + a.sheet + ' — ' + a.iban + (f ? '\n   ' + f.getName() : '\n   aucun CSV trouvé');
    });

    ui.alert(
      'Configuration CBC',
      'Dossier : ' + folder.getName() + '\n\n' + lines.join('\n\n'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Configuration incomplète', e.message, ui.ButtonSet.OK);
  }
}

function miseAJourCBC() {
  const ui = SpreadsheetApp.getUi();
  try {
    getImportFolder_(); // vérifie la configuration avant la sauvegarde
    creerSauvegardeQuotidienne_();
    const result = executerMiseAJour_(false);

    ui.alert(
      'Mise à jour CBC terminée',
      'CSV lus : ' + result.filesRead + '\n' +
      'Opérations présentes dans les derniers CSV : ' + result.rowsParsed + '\n' +
      'Nouvelles opérations ajoutées : ' + result.rowsAdded + '\n\n' +
      (result.rowsAdded === 0
        ? 'Aucune nouvelle opération : le classeur était déjà à jour.'
        : 'Les onglets banque et TOTAUX ont été mis à jour.'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erreur CBC', e.message, ui.ButtonSet.OK);
    throw e;
  }
}

/** Fonction appelée par le déclencheur horaire. Ne montre aucune boîte de dialogue. */
function miseAJourCBCAuto() {
  try {
    executerMiseAJour_(true);
  } catch (e) {
    console.error('CBC auto : ' + e.stack);
  }
}

function executerMiseAJour_(silent) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('Une autre mise à jour CBC est déjà en cours. Réessaie dans quelques instants.');

  try {
    const ss = getSpreadsheet_();
    verifierOngletsRequis_(ss);
    const folder = getImportFolder_();
    const latest = getLatestCsvFiles_(folder);

    const allTransactions = [];
    const parsedBySheet = {};
    const journalRows = [];
    let filesRead = 0;
    let rowsParsed = 0;

    CBC.ACCOUNTS.forEach(def => {
      parsedBySheet[def.sheet] = [];
      const file = latest[def.iban];
      if (!file) {
        journalRows.push([new Date(), '', def.iban, def.sheet, 0, 0, 'Aucun CSV trouvé']);
        return;
      }

      const parsed = parseCbcCsv_(file, def);
      filesRead++;
      rowsParsed += parsed.length;
      parsedBySheet[def.sheet] = parsed;
      parsed.forEach(t => allTransactions.push(t));
      journalRows.push([new Date(), file.getName(), def.iban, def.sheet, parsed.length, '', 'CSV lu']);
    });

    if (filesRead === 0) {
      throw new Error(
        'Aucun fichier CSV CBC n’a été trouvé dans le dossier configuré.\n\n' +
        'Les noms attendus contiennent par exemple : export_BE07732021748966_...csv'
      );
    }

    // Comme la macro Excel : DIVERS-BANQUE reflète les derniers CSV lus.
    ecrireDiversBanque_(ss, allTransactions);

    let rowsAdded = 0;
    const summaries = [];
    const addedBySheet = {};

    CBC.ACCOUNTS.forEach(def => {
      const sheet = ss.getSheetByName(def.sheet);
      const added = ajouterNouvellesOperations_(sheet, parsedBySheet[def.sheet] || [], def);
      rowsAdded += added;
      addedBySheet[def.sheet] = added;
      summaries.push(calculerResumeOnglet_(sheet, def));
    });

    // On ne déplace Actuel -> Précédent que si au moins une nouvelle opération a réellement été ajoutée.
    mettreAJourTotaux_(ss, summaries, rowsAdded > 0);
    SpreadsheetApp.flush();

    // Complète le journal avec le nombre réellement ajouté par onglet.
    journalRows.forEach(r => {
      if (r[3] && addedBySheet[r[3]] !== undefined) {
        r[5] = addedBySheet[r[3]];
        if (r[6] === 'CSV lu') r[6] = addedBySheet[r[3]] > 0 ? 'Import OK' : 'Déjà à jour';
      }
    });
    ecrireJournal_(ss, journalRows);

    if (!silent) ss.toast('CBC : ' + rowsAdded + ' nouvelle(s) opération(s) ajoutée(s).', 'CBC BANQUE', 5);
    return { filesRead, rowsParsed, rowsAdded };
  } finally {
    lock.releaseLock();
  }
}

function getLatestCsvFiles_(folder) {
  const selected = {};
  const ranks = {};
  const known = {};
  CBC.ACCOUNTS.forEach(a => known[a.iban] = true);

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (!/\.csv$/i.test(name)) continue;

    const mIban = name.match(/export_(BE\d{14})/i);
    if (!mIban) continue;
    const iban = mIban[1].toUpperCase();
    if (!known[iban]) continue;

    const mDate = name.match(/_(\d{8})_(\d{4})/);
    let rank;
    if (mDate) {
      const ds = mDate[1];
      const ts = mDate[2];
      rank = new Date(
        Number(ds.slice(0, 4)), Number(ds.slice(4, 6)) - 1, Number(ds.slice(6, 8)),
        Number(ts.slice(0, 2)), Number(ts.slice(2, 4)), 0
      ).getTime();
    } else {
      rank = file.getLastUpdated().getTime();
    }

    if (!selected[iban] || rank > ranks[iban]) {
      selected[iban] = file;
      ranks[iban] = rank;
    }
  }
  return selected;
}

function parseCbcCsv_(file, def) {
  const blob = file.getBlob();
  let text = blob.getDataAsString('UTF-8');
  if (text.indexOf('\uFFFD') !== -1) text = blob.getDataAsString('windows-1252');
  text = text.replace(/^\uFEFF/, '');

  let rows = Utilities.parseCsv(text, ';');
  rows = rows.filter(row => row.some(v => String(v || '').trim() !== ''));
  if (!rows.length) return [];

  const headerIndex = trouverLigneEntete_(rows);
  const headers = headerIndex >= 0 ? rows[headerIndex] : [];
  const map = construireMapColonnes_(headers);
  const start = headerIndex >= 0 ? headerIndex + 1 : 0;
  const transactions = [];

  for (let i = start; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(v => String(v || '').trim() === '')) continue;

    const date = convertirDate_(cell_(row, map.date));
    const amount = convertirNombre_(cell_(row, map.amount));
    const balance = convertirNombre_(cell_(row, map.balance));

    // Ignore les lignes d'en-tête, total, solde initial, etc.
    if (!date || amount === null || balance === null) continue;

    const commParts = (map.communication || [])
      .map(idx => nettoyerTexte_(cell_(row, idx)))
      .filter(Boolean);

    const transaction = {
      date: date,
      extract: nettoyerTexte_(cell_(row, map.extract)),
      account: normaliserIban_(cell_(row, map.account)) || def.iban,
      counterpartyAccount: nettoyerTexte_(cell_(row, map.counterpartyAccount)),
      counterpartyName: nettoyerTexte_(cell_(row, map.counterpartyName)),
      description: nettoyerTexte_(cell_(row, map.description)),
      communication: valeursUniques_(commParts).join(' '),
      amount: arrondir2_(amount),
      balance: arrondir2_(balance),
      sourceDataLine: i - start + 1,
      sourceCsvLine: i + 1,
      positionInExtract: 0,
      targetSheet: def.sheet,
      sourceFile: file.getName()
    };

    transactions.push(transaction);
  }

  // Les exports CBC sont généralement du plus récent au plus ancien.
  // On remet ici les opérations dans l'ordre chronologique du classeur Excel.
  transactions.sort(comparerTransactions_);

  let currentExtract = null;
  let pos = 0;
  transactions.forEach(t => {
    const key = t.extract || cleDate_(t.date);
    if (key !== currentExtract) {
      currentExtract = key;
      pos = 1;
    } else {
      pos++;
    }
    t.positionInExtract = pos;
  });

  return transactions;
}

function trouverLigneEntete_(rows) {
  const max = Math.min(rows.length, 15);
  for (let i = 0; i < max; i++) {
    const h = rows[i].map(normaliserEntete_);
    let score = 0;
    if (h.some(x => x.includes('date'))) score++;
    if (h.some(x => x.includes('montant') || x === 'amount')) score++;
    if (h.some(x => x.includes('solde') || x === 'balance')) score++;
    if (h.some(x => x.includes('description') || x.includes('communication'))) score++;
    if (h.some(x => x.includes('extrait') || x.includes('statement'))) score++;
    if (score >= 3) return i;
  }
  return -1;
}

function construireMapColonnes_(headers) {
  const h = headers.map(normaliserEntete_);

  // Les index de secours correspondent à la disposition du fichier CBC utilisé
  // par le classeur Excel : A,B,C,D,E,F,G,(H),I,(J),K.
  const map = {
    date: trouverColonne_(h, ['date comptable', 'date operation', 'date valeur', 'date'], 0),
    extract: trouverColonne_(h, ['numero extrait', 'n extrait', 'extrait', 'statement'], 1),
    account: trouverColonne_(h, ['numero de compte', 'compte donneur ordre', 'donneur ordre', 'account'], 2),
    counterpartyAccount: trouverColonne_(h, ['compte contrepartie', 'numero compte contrepartie', 'compte beneficiaire', 'counterparty account'], 3),
    counterpartyName: trouverColonne_(h, ['nom contrepartie', 'nom beneficiaire', 'counterparty name', 'beneficiaire'], 4),
    description: trouverColonne_(h, ['description', 'details', 'detail'], 5),
    amount: trouverColonne_(h, ['montant', 'amount'], 8),
    balance: trouverColonne_(h, ['solde', 'balance'], 10),
    communication: []
  };

  if (h.length) {
    h.forEach((x, idx) => {
      if (x.includes('communication')) map.communication.push(idx);
    });
  }
  if (!map.communication.length) map.communication = [6];
  return map;
}

function trouverColonne_(headers, aliases, fallback) {
  if (!headers.length) return fallback;

  const normalizedAliases = aliases.map(normaliserEntete_);
  for (let a = 0; a < normalizedAliases.length; a++) {
    const idx = headers.indexOf(normalizedAliases[a]);
    if (idx >= 0) return idx;
  }
  for (let a = 0; a < normalizedAliases.length; a++) {
    const alias = normalizedAliases[a];
    const idx = headers.findIndex(x => x && (x.includes(alias) || alias.includes(x)));
    if (idx >= 0) return idx;
  }
  return fallback;
}

function ecrireDiversBanque_(ss, transactions) {
  const sheet = ss.getSheetByName(CBC.DIVERS_SHEET);
  if (!sheet) throw new Error('Onglet manquant : ' + CBC.DIVERS_SHEET);

  const oldLast = sheet.getLastRow();
  if (oldLast >= CBC.DATA_START_ROW) {
    sheet.getRange(CBC.DATA_START_ROW, 1, oldLast - CBC.DATA_START_ROW + 1, CBC.DIVERS_WIDTH).clearContent();
  }

  if (!transactions.length) return;
  const order = {};
  CBC.ACCOUNTS.forEach((a, i) => order[a.sheet] = i);
  transactions.sort((a, b) => {
    const d = comparerTransactions_(a, b);
    if (d !== 0) return d;
    return (order[a.targetSheet] || 0) - (order[b.targetSheet] || 0);
  });

  assurerNombreLignes_(sheet, CBC.DATA_START_ROW + transactions.length + 2);
  const values = transactions.map(t => [
    t.date, t.extract, t.account, t.counterpartyAccount, t.counterpartyName,
    t.description, t.communication, '', t.amount, '', t.balance, '', '', '',
    t.positionInExtract, t.sourceDataLine, t.sourceCsvLine, t.targetSheet
  ]);

  sheet.getRange(CBC.DATA_START_ROW, 1, values.length, CBC.DIVERS_WIDTH).setValues(values);
  sheet.getRange(CBC.DATA_START_ROW, 1, values.length, 1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(CBC.DATA_START_ROW, 9, values.length, 1).setNumberFormat('#,##0.00');
  sheet.getRange(CBC.DATA_START_ROW, 11, values.length, 1).setNumberFormat('#,##0.00');
}

function ajouterNouvellesOperations_(sheet, transactions, def) {
  if (!transactions.length) {
    ecrireTotal_(sheet, def);
    return 0;
  }

  const lastOp = derniereLigneOperation_(sheet);
  const existingKeys = new Set();

  if (lastOp >= CBC.DATA_START_ROW) {
    const values = sheet.getRange(CBC.DATA_START_ROW, 1, lastOp - CBC.DATA_START_ROW + 1, 11).getValues();
    values.forEach(r => {
      if (r[0] === '' || r[0] === null) return;
      existingKeys.add(construireCle_(r[0], r[1], r[5], r[8], r[10]));
    });
  }

  const toAdd = [];
  const pendingKeys = new Set();
  transactions.forEach(t => {
    const key = construireCle_(t.date, t.extract, t.description, t.amount, t.balance);
    if (!existingKeys.has(key) && !pendingKeys.has(key)) {
      pendingKeys.add(key);
      toAdd.push(t);
    }
  });

  toAdd.sort(comparerTransactions_);
  const oldTotalRow = trouverLigneTotal_(sheet);
  if (oldTotalRow) sheet.getRange(oldTotalRow, 8, 1, 7).clearContent(); // H:N

  if (!toAdd.length) {
    ecrireTotal_(sheet, def);
    return 0;
  }

  const startRow = Math.max(CBC.DATA_START_ROW, lastOp + 1);
  const endRow = startRow + toAdd.length - 1;
  assurerNombreLignes_(sheet, endRow + 3);

  // Reprend le format de la dernière ligne bancaire existante.
  if (lastOp >= CBC.DATA_START_ROW) {
    try {
      sheet.getRange(lastOp, 1, 1, CBC.DATA_WIDTH)
        .copyFormatToRange(sheet.getSheetId(), 1, CBC.DATA_WIDTH, startRow, endRow);
    } catch (e) {
      console.warn('Copie de format ignorée pour ' + sheet.getName() + ' : ' + e.message);
    }
  }

  let previousCumulative = 0;
  if (lastOp >= CBC.DATA_START_ROW) {
    const m = convertirNombre_(sheet.getRange(lastOp, 13).getValue());
    const k = convertirNombre_(sheet.getRange(lastOp, 11).getValue());
    previousCumulative = m !== null ? m : (k !== null ? k : 0);
  }

  const rows = [];
  toAdd.forEach(t => {
    previousCumulative = arrondir2_(previousCumulative + t.amount);
    const verification = arrondir2_(t.balance - previousCumulative);
    rows.push([
      t.date, t.extract, t.account, t.counterpartyAccount, t.counterpartyName,
      t.description, t.communication, '', t.amount, '', t.balance, '',
      previousCumulative, verification
    ]);
  });

  sheet.getRange(startRow, 1, rows.length, CBC.DATA_WIDTH).setValues(rows);
  sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy');
  [9, 11, 13, 14].forEach(col => sheet.getRange(startRow, col, rows.length, 1).setNumberFormat('#,##0.00'));

  ecrireTotal_(sheet, def);
  return toAdd.length;
}

function ecrireTotal_(sheet, def) {
  const oldTotal = trouverLigneTotal_(sheet);
  if (oldTotal) sheet.getRange(oldTotal, 8, 1, 7).clearContent();

  const lastOp = derniereLigneOperation_(sheet);
  if (lastOp < CBC.DATA_START_ROW) return;

  const totalRow = lastOp + 2;
  assurerNombreLignes_(sheet, totalRow);
  sheet.getRange(totalRow, 8, 1, 7).clearContent();
  sheet.getRange(totalRow, 8).setValue('TOTAL').setFontWeight('bold');

  let totalFormula = '=SUM(I' + CBC.DATA_START_ROW + ':I' + lastOp + ')';
  if (def.totalCorrection) {
    totalFormula += (def.totalCorrection < 0 ? '' : '+') + def.totalCorrection;
  }
  sheet.getRange(totalRow, 9).setFormula(totalFormula).setFontWeight('bold').setNumberFormat('#,##0.00');

  const verifStart = Math.min(4, lastOp);
  sheet.getRange(totalRow, 14)
    .setFormula('=SUM(N' + verifStart + ':N' + lastOp + ')')
    .setFontWeight('bold')
    .setNumberFormat('#,##0.00');
}

function recalculerTotauxCBC() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = getSpreadsheet_();
    verifierOngletsRequis_(ss);
    const summaries = CBC.ACCOUNTS.map(def => {
      const sheet = ss.getSheetByName(def.sheet);
      ecrireTotal_(sheet, def);
      return calculerResumeOnglet_(sheet, def);
    });
    // Recalcul manuel : ne modifie pas la colonne "Précédent".
    mettreAJourTotaux_(ss, summaries, false);
    SpreadsheetApp.flush();
    ui.alert('Totaux recalculés sans modifier les valeurs "Précédent".');
  } catch (e) {
    ui.alert('Erreur', e.message, ui.ButtonSet.OK);
  }
}

function calculerResumeOnglet_(sheet, def) {
  const lastOp = derniereLigneOperation_(sheet);
  if (lastOp < CBC.DATA_START_ROW) {
    return { def, lastOp: '', extract: '', date: '', total: 0, verification: 0 };
  }

  const lastValues = sheet.getRange(lastOp, 1, 1, 14).getValues()[0];
  const amounts = sheet.getRange(CBC.DATA_START_ROW, 9, lastOp - CBC.DATA_START_ROW + 1, 1).getValues();
  let total = amounts.reduce((s, r) => s + (convertirNombre_(r[0]) || 0), 0) + (def.totalCorrection || 0);
  total = arrondir2_(total);

  let verification = 0;
  if (lastOp >= 4) {
    const vals = sheet.getRange(4, 14, lastOp - 3, 1).getValues();
    verification = arrondir2_(vals.reduce((s, r) => s + (convertirNombre_(r[0]) || 0), 0));
  }

  return {
    def: def,
    lastOp: lastOp,
    extract: lastValues[1],
    date: lastValues[0],
    total: total,
    verification: verification
  };
}

function mettreAJourTotaux_(ss, summaries, shiftPrevious) {
  const sheet = ss.getSheetByName(CBC.TOTAUX_SHEET);
  if (!sheet) throw new Error('Onglet manquant : ' + CBC.TOTAUX_SHEET);

  summaries.forEach(s => {
    const r = s.def.totauxRow;

    if (shiftPrevious) {
      const current = sheet.getRange(r, 11, 1, 5).getValues()[0]; // K:O
      const hasCurrent = current.some(v => v !== '' && v !== null);
      if (hasCurrent) sheet.getRange(r, 5, 1, 5).setValues([current]); // E:I
    }

    sheet.getRange(r, 11, 1, 5).setValues([[
      s.lastOp, s.extract, s.date, s.total, s.verification
    ]]);
    sheet.getRange(r, 13).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(r, 14, 1, 2).setNumberFormat('#,##0.00');
    sheet.getRange(r, 17).setFormula('=N' + r + '-H' + r).setNumberFormat('#,##0.00');
  });

  // Conserve/répare les totaux généraux du modèle Excel.
  sheet.getRange('H25').setFormula('=SUM(H4:H24)').setNumberFormat('#,##0.00');
  sheet.getRange('I25').setFormula('=SUM(I4:I24)').setNumberFormat('#,##0.00');
  sheet.getRange('N25').setFormula('=SUM(N4:N24)').setNumberFormat('#,##0.00');
  sheet.getRange('O25').setFormula('=SUM(O4:O24)').setNumberFormat('#,##0.00');
  sheet.getRange('Q25').setFormula('=N25-H25').setNumberFormat('#,##0.00');
}

function derniereLigneOperation_(sheet) {
  const max = sheet.getLastRow();
  if (max < CBC.DATA_START_ROW) return CBC.DATA_START_ROW - 1;

  const values = sheet.getRange(CBC.DATA_START_ROW, 1, max - CBC.DATA_START_ROW + 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) return CBC.DATA_START_ROW + i;
  }
  return CBC.DATA_START_ROW - 1;
}

function trouverLigneTotal_(sheet) {
  const max = sheet.getLastRow();
  if (max < CBC.DATA_START_ROW) return 0;
  const vals = sheet.getRange(CBC.DATA_START_ROW, 8, max - CBC.DATA_START_ROW + 1, 1).getDisplayValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim().toUpperCase() === 'TOTAL') return CBC.DATA_START_ROW + i;
  }
  return 0;
}

function construireCle_(date, extract, description, amount, balance) {
  return [
    cleDate_(date),
    nettoyerPourCle_(extract),
    nettoyerPourCle_(description),
    cents_(amount),
    cents_(balance)
  ].join('|');
}

function comparerTransactions_(a, b) {
  const da = a.date instanceof Date ? a.date.getTime() : 0;
  const db = b.date instanceof Date ? b.date.getTime() : 0;
  if (da !== db) return da - db;

  const ea = Number(String(a.extract || '').replace(/\D/g, '')) || 0;
  const eb = Number(String(b.extract || '').replace(/\D/g, '')) || 0;
  if (ea !== eb) return ea - eb;

  // Dans le CSV CBC les lignes sont souvent en ordre inverse : la ligne la plus basse
  // du CSV est donc généralement la première chronologiquement dans le même extrait.
  return (b.sourceCsvLine || 0) - (a.sourceCsvLine || 0);
}


function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getDocumentProperties().setProperty(CBC.PROP_SPREADSHEET_ID, active.getId());
    return active;
  }
  const id = PropertiesService.getDocumentProperties().getProperty(CBC.PROP_SPREADSHEET_ID);
  if (!id) throw new Error('Impossible d’identifier le Google Sheet lié au script. Ouvre le classeur puis recharge la page.');
  return SpreadsheetApp.openById(id);
}

function getImportFolder_() {
  const id = PropertiesService.getDocumentProperties().getProperty(CBC.PROP_FOLDER_ID);
  if (!id) {
    throw new Error(
      'Le dossier des CSV n’est pas encore configuré.\n\n' +
      'Dans le classeur : CBC BANQUE > Configurer le dossier CSV.'
    );
  }
  try {
    return DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('Le dossier Drive configuré n’est plus accessible. Reconfigure-le depuis le menu CBC BANQUE.');
  }
}

function extraireIdDrive_(text) {
  const s = String(text || '').trim();
  if (/^[A-Za-z0-9_-]{15,}$/.test(s)) return s;
  const m = s.match(/\/folders\/([A-Za-z0-9_-]+)/) || s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  return m ? m[1] : '';
}

function verifierOngletsRequis_(ss) {
  const required = [CBC.DIVERS_SHEET, CBC.TOTAUX_SHEET].concat(CBC.ACCOUNTS.map(a => a.sheet));
  const missing = required.filter(name => !ss.getSheetByName(name));
  if (missing.length) throw new Error('Onglet(s) manquant(s) dans le Google Sheet : ' + missing.join(', '));
}

function ecrireJournal_(ss, rows) {
  if (!rows || !rows.length) return;
  let sheet = ss.getSheetByName(CBC.JOURNAL_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CBC.JOURNAL_SHEET);
    sheet.getRange(1, 1, 1, 7).setValues([[
      'Date/heure', 'Fichier CSV', 'Compte', 'Onglet', 'Lignes lues', 'Nouvelles lignes', 'Statut'
    ]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

function creerSauvegardeCBC() {
  const ui = SpreadsheetApp.getUi();
  try {
    const file = creerSauvegarde_(true);
    ui.alert('Sauvegarde créée :\n' + file.getName());
  } catch (e) {
    ui.alert('Erreur sauvegarde', e.message, ui.ButtonSet.OK);
  }
}

function creerSauvegardeQuotidienne_() {
  const props = PropertiesService.getDocumentProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  if (props.getProperty(CBC.PROP_LAST_BACKUP) === today) return;
  creerSauvegarde_(false);
  props.setProperty(CBC.PROP_LAST_BACKUP, today);
}

function creerSauvegarde_(manual) {
  const ss = getSpreadsheet_();
  const src = DriveApp.getFileById(ss.getId());
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const name = ss.getName() + '_BACKUP_' + stamp;
  const parents = src.getParents();
  const copy = parents.hasNext() ? src.makeCopy(name, parents.next()) : src.makeCopy(name);
  console.log((manual ? 'Sauvegarde manuelle : ' : 'Sauvegarde quotidienne : ') + copy.getName());
  return copy;
}

function activerMiseAJourAuto() {
  const ui = SpreadsheetApp.getUi();
  try {
    getImportFolder_();
    supprimerTriggersCBC_();
    ScriptApp.newTrigger('miseAJourCBCAuto').timeBased().everyHours(1).create();
    ui.alert('Mise à jour automatique activée : le dossier sera vérifié environ une fois par heure.');
  } catch (e) {
    ui.alert('Erreur', e.message, ui.ButtonSet.OK);
  }
}

function desactiverMiseAJourAuto() {
  supprimerTriggersCBC_();
  SpreadsheetApp.getUi().alert('Mise à jour automatique désactivée.');
}

function supprimerTriggersCBC_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'miseAJourCBCAuto') ScriptApp.deleteTrigger(t);
  });
}

function assurerNombreLignes_(sheet, wantedLastRow) {
  const max = sheet.getMaxRows();
  if (wantedLastRow > max) sheet.insertRowsAfter(max, wantedLastRow - max);
}

function cell_(row, idx) {
  return (idx >= 0 && idx < row.length) ? row[idx] : '';
}

function normaliserEntete_(value) {
  return String(value || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[°º]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function nettoyerTexte_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function nettoyerPourCle_(value) {
  return nettoyerTexte_(value).toUpperCase();
}

function normaliserIban_(value) {
  const s = String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return /^BE\d{14}$/.test(s) ? s : '';
}

function convertirNombre_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return isFinite(value) ? value : null;

  let s = String(value).trim();
  if (!s) return null;
  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1);
  }
  s = s.replace(/\u00A0/g, '').replace(/\s/g, '').replace(/€/g, '');

  if (s.includes(',') && s.includes('.')) {
    // Format européen 1.234,56
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (s.includes(',')) {
    s = s.replace(/\./g, '').replace(',', '.');
  }

  s = s.replace(/[^0-9+\-.]/g, '');
  const n = Number(s);
  if (!isFinite(n)) return null;
  return negative ? -n : n;
}

function convertirDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === 'number' && value > 20000 && value < 100000) {
    const utc = Date.UTC(1899, 11, 30) + Math.round(value) * 86400000;
    const d = new Date(utc);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), 12, 0, 0);
  }

  const s = String(value || '').trim();
  if (!s) return null;

  let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/);
  if (m) {
    let y = Number(m[3]);
    if (y < 100) y += (y >= 70 ? 1900 : 2000);
    return dateValide_(y, Number(m[2]), Number(m[1]));
  }

  m = s.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/);
  if (m) return dateValide_(Number(m[1]), Number(m[2]), Number(m[3]));

  if (/^\d+(?:[.,]\d+)?$/.test(s)) {
    const serial = convertirNombre_(s);
    if (serial !== null && serial > 20000 && serial < 100000) return convertirDate_(serial);
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function dateValide_(year, month, day) {
  const d = new Date(year, month - 1, day, 12, 0, 0);
  if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) return null;
  return d;
}

function cleDate_(value) {
  const d = convertirDate_(value);
  if (!d) return nettoyerPourCle_(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function cents_(value) {
  const n = convertirNombre_(value);
  return n === null ? '' : String(Math.round(n * 100));
}

function arrondir2_(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

function valeursUniques_(values) {
  const seen = {};
  const out = [];
  values.forEach(v => {
    const key = nettoyerPourCle_(v);
    if (key && !seen[key]) {
      seen[key] = true;
      out.push(v);
    }
  });
  return out;
}
