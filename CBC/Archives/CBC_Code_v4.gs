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
  VERSION: '2.0.0',
  TIMEZONE: 'Europe/Brussels',
  SEARCH_SHEET: 'RECHERCHE-CBC',
  PROP_SEARCH_CRITERIA: 'CBC_SEARCH_CRITERIA',
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

const GERICO_CFG = Object.freeze({
  SHEET_NAMES: ['GERICO', 'GERICO-'],
  SOURCE_SHEET: 'AS-CBC',
  COUNTERPARTY_IBANS: ['BE34068908139790'],
  NAME_CONTAINS: ['GERICO'],
  DATA_START_ROW: 5,
  DATA_WIDTH: 9
});

function onOpen() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) PropertiesService.getDocumentProperties().setProperty(CBC.PROP_SPREADSHEET_ID, active.getId());
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CBC BANQUE')
    .addItem('Mettre à jour maintenant', 'miseAJourCBC')
    .addItem('Mettre à jour GERICO', 'mettreAJourGericoManuel')
    .addSeparator()
    .addItem('🔎 Recherche avancée', 'ouvrirRechercheAvancee')
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
    'Le menu contient maintenant aussi GERICO et la Recherche avancée.\n\nÉtape suivante : CBC BANQUE > Configurer le dossier CSV, puis sélectionne simplement ton dossier Drive.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function configurerDossierImports() {
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; margin: 0; color: #202124; }
    .wrap { padding: 16px; }
    h2 { margin: 0 0 6px; font-size: 18px; }
    .hint { color: #5f6368; font-size: 12px; margin-bottom: 12px; }
    .toolbar { display: flex; gap: 8px; margin-bottom: 10px; }
    input { flex: 1; padding: 9px 10px; border: 1px solid #dadce0; border-radius: 6px; }
    button { border: 1px solid #dadce0; background: #fff; border-radius: 6px; padding: 8px 12px; cursor: pointer; }
    button.primary { background: #1a73e8; color: #fff; border-color: #1a73e8; }
    button:disabled { opacity: .5; cursor: default; }
    .path { padding: 8px 10px; background: #f8f9fa; border-radius: 6px; font-size: 12px; margin-bottom: 8px; }
    .list { height: 300px; overflow: auto; border: 1px solid #dadce0; border-radius: 6px; }
    .row { display: flex; align-items: center; gap: 8px; padding: 10px 12px; border-bottom: 1px solid #eee; cursor: pointer; }
    .row:last-child { border-bottom: 0; }
    .row:hover { background: #f1f3f4; }
    .icon { font-size: 18px; }
    .name { flex: 1; }
    .empty { padding: 22px; text-align: center; color: #777; }
    .footer { display: flex; justify-content: space-between; align-items: center; margin-top: 12px; gap: 8px; }
    .status { font-size: 12px; color: #5f6368; flex: 1; }
  </style>
</head>
<body>
  <div class="wrap">
    <h2>Choisir le dossier des exports CBC</h2>
    <div class="hint">Ouvre les dossiers puis clique sur <b>Sélectionner ce dossier</b>. Aucune URL à copier.</div>

    <div class="toolbar">
      <button id="back" onclick="back()">← Retour</button>
      <input id="search" type="text" placeholder="Filtrer les dossiers affichés…" oninput="filterRows()">
    </div>

    <div id="path" class="path">Chargement…</div>
    <div id="list" class="list"><div class="empty">Chargement des dossiers…</div></div>

    <div class="footer">
      <div id="status" class="status"></div>
      <button onclick="google.script.host.close()">Annuler</button>
      <button id="select" class="primary" onclick="selectCurrent()">Sélectionner ce dossier</button>
    </div>
  </div>

<script>
  let current = null;
  let history = [];

  function load(id, pushHistory) {
    document.getElementById('status').textContent = 'Chargement…';
    google.script.run
      .withSuccessHandler(data => {
        if (pushHistory && current) history.push(current.id);
        current = data;
        render(data);
      })
      .withFailureHandler(err => {
        document.getElementById('status').textContent = 'Erreur : ' + err.message;
      })
      .cbcListerDossiersDrive(id || 'ROOT');
  }

  function render(data) {
    document.getElementById('path').textContent = 'Dossier actuel : ' + data.name;
    document.getElementById('back').disabled = history.length === 0 && data.isRoot;
    document.getElementById('status').textContent = data.folders.length + ' dossier(s)';
    document.getElementById('search').value = '';

    const list = document.getElementById('list');
    list.innerHTML = '';
    if (!data.folders.length) {
      list.innerHTML = '<div class="empty">Aucun sous-dossier. Tu peux sélectionner ce dossier.</div>';
      return;
    }

    data.folders.forEach(f => {
      const row = document.createElement('div');
      row.className = 'row';
      row.dataset.name = f.name.toLowerCase();
      row.innerHTML = '<span class="icon">📁</span><span class="name"></span><span>›</span>';
      row.querySelector('.name').textContent = f.name;
      row.onclick = () => load(f.id, true);
      list.appendChild(row);
    });
  }

  function back() {
    if (history.length) {
      const id = history.pop();
      load(id, false);
    } else if (current && current.parentId) {
      load(current.parentId, false);
    }
  }

  function filterRows() {
    const q = document.getElementById('search').value.toLowerCase().trim();
    document.querySelectorAll('.row').forEach(r => {
      r.style.display = !q || r.dataset.name.includes(q) ? 'flex' : 'none';
    });
  }

  function selectCurrent() {
    if (!current) return;
    document.getElementById('select').disabled = true;
    document.getElementById('status').textContent = 'Enregistrement…';
    google.script.run
      .withSuccessHandler(result => {
        document.getElementById('status').textContent = 'Dossier enregistré : ' + result.name;
        setTimeout(() => google.script.host.close(), 700);
      })
      .withFailureHandler(err => {
        document.getElementById('select').disabled = false;
        document.getElementById('status').textContent = 'Erreur : ' + err.message;
      })
      .cbcEnregistrerDossierImport(current.id);
  }

  load('ROOT', false);
</script>
</body>
</html>`)
    .setWidth(640)
    .setHeight(470);

  SpreadsheetApp.getUi().showModalDialog(html, 'CBC — Sélection du dossier');
}

/** Appelée par la fenêtre HTML : liste les sous-dossiers du dossier choisi. */
function cbcListerDossiersDrive(folderId) {
  const root = DriveApp.getRootFolder();
  const folder = (!folderId || folderId === 'ROOT') ? root : DriveApp.getFolderById(folderId);
  const folders = [];
  const it = folder.getFolders();

  while (it.hasNext()) {
    const f = it.next();
    folders.push({ id: f.getId(), name: f.getName() });
  }
  folders.sort((a, b) => a.name.localeCompare(b.name, 'fr', { sensitivity: 'base' }));

  let parentId = '';
  try {
    const parents = folder.getParents();
    if (parents.hasNext()) parentId = parents.next().getId();
  } catch (e) {}

  return {
    id: folder.getId(),
    name: folder.getName() || 'Mon Drive',
    parentId: parentId,
    isRoot: folder.getId() === root.getId(),
    folders: folders
  };
}

/** Appelée par la fenêtre HTML : mémorise le dossier sélectionné. */
function cbcEnregistrerDossierImport(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  PropertiesService.getDocumentProperties().setProperty(CBC.PROP_FOLDER_ID, folderId);
  return { id: folderId, name: folder.getName() };
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
      'Nouvelles opérations banque : ' + result.rowsAdded + '\n' +
      'Nouvelles lignes GERICO : ' + result.gericoAdded + '\n\n' +
      (result.rowsAdded === 0 && result.gericoAdded === 0
        ? 'Aucune nouvelle opération : le classeur était déjà à jour.'
        : 'Les onglets banque, GERICO et TOTAUX ont été mis à jour.'),
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

    // Synchronise aussi l'onglet GERICO depuis AS-CBC.
    const gericoResult = mettreAJourGerico_(ss, { silent: true });
    const gericoAdded = gericoResult.added || 0;

    // On ne déplace Actuel -> Précédent que si au moins une nouvelle opération bancaire a réellement été ajoutée.
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

    if (!silent) ss.toast('CBC : ' + rowsAdded + ' opération(s) banque + ' + gericoAdded + ' ligne(s) GERICO.', 'CBC BANQUE', 5);
    return { filesRead, rowsParsed, rowsAdded, gericoAdded };
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
  const existingStableKeys = new Set();

  if (lastOp >= CBC.DATA_START_ROW) {
    const values = sheet.getRange(CBC.DATA_START_ROW, 1, lastOp - CBC.DATA_START_ROW + 1, 11).getValues();
    values.forEach(r => {
      if (r[0] === '' || r[0] === null) return;
      existingKeys.add(construireCle_(r[0], r[1], r[5], r[8], r[10]));
      existingStableKeys.add(construireCleStable_(r[0], r[1], r[2], r[8], r[10]));
    });
  }

  const toAdd = [];
  const pendingKeys = new Set();
  const pendingStableKeys = new Set();
  transactions.forEach(t => {
    const key = construireCle_(t.date, t.extract, t.description, t.amount, t.balance);
    const stableKey = construireCleStable_(t.date, t.extract, t.account, t.amount, t.balance);
    const exists = existingKeys.has(key) || existingStableKeys.has(stableKey) ||
                   pendingKeys.has(key) || pendingStableKeys.has(stableKey);
    if (!exists) {
      pendingKeys.add(key);
      pendingStableKeys.add(stableKey);
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
    normaliserExtraitPourCle_(extract),
    nettoyerPourCle_(description),
    cents_(amount),
    cents_(balance)
  ].join('|');
}

/**
 * CBC exporte souvent le n° d'extrait avec un zéro devant (ex. 02026203),
 * alors que Google Sheets/Excel l'historique l'affiche comme 2026203.
 * Sans cette normalisation, une opération déjà présente paraît "nouvelle".
 */
function normaliserExtraitPourCle_(value) {
  const s = nettoyerTexte_(value).replace(/\s+/g, '');
  if (!s) return '';
  if (/^\d+$/.test(s)) return String(Number(s)); // retire les zéros à gauche
  return s.toUpperCase();
}

/** Clé de secours très stable, indépendante du texte de description. */
function construireCleStable_(date, extract, account, amount, balance) {
  return [
    cleDate_(date),
    normaliserExtraitPourCle_(extract),
    normaliserIban_(account) || nettoyerPourCle_(account),
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
  const today = Utilities.formatDate(new Date(), CBC.TIMEZONE, 'yyyyMMdd');
  if (props.getProperty(CBC.PROP_LAST_BACKUP) === today) return;
  creerSauvegarde_(false);
  props.setProperty(CBC.PROP_LAST_BACKUP, today);
}

function creerSauvegarde_(manual) {
  const ss = getSpreadsheet_();
  const src = DriveApp.getFileById(ss.getId());
  const stamp = Utilities.formatDate(new Date(), CBC.TIMEZONE, 'yyyyMMdd_HHmmss');
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
  return Utilities.formatDate(d, CBC.TIMEZONE, 'yyyy-MM-dd');
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

/* ========================================================================== */
/* GERICO                                                                     */
/* ========================================================================== */

function mettreAJourGericoManuel() {
  const ui = SpreadsheetApp.getUi();
  try {
    creerSauvegardeQuotidienne_();
    const ss = getSpreadsheet_();
    const result = mettreAJourGerico_(ss, { silent: false });
    ui.alert(
      'GERICO mis à jour',
      'Nouvelles lignes ajoutées : ' + result.added + '\n' +
      'Dernière année : ' + (result.lastYear || '') + '\n' +
      'Mois pris en compte : ' + result.currentMonth + '/12\n\n' +
      (result.added ? 'Les nouvelles opérations ont été ajoutées sans doublon.' : 'Aucune nouvelle opération GERICO à ajouter.'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erreur GERICO', e.message, ui.ButtonSet.OK);
    throw e;
  }
}

function mettreAJourGerico_(ss, options) {
  options = options || {};
  const sheet = trouverOngletGerico_(ss);
  if (!sheet) {
    if (options.silent) return { added: 0, lastYear: '', currentMonth: moisCourantBelgique_(), skipped: true };
    throw new Error('Onglet GERICO introuvable. Le script cherche un onglet nommé GERICO ou GERICO-.');
  }

  const source = ss.getSheetByName(GERICO_CFG.SOURCE_SHEET);
  if (!source) throw new Error('Onglet source GERICO introuvable : ' + GERICO_CFG.SOURCE_SHEET);

  const currentYear = anneeCouranteBelgique_();
  const currentMonth = moisCourantBelgique_();
  const zoneAvant = analyserZoneGerico_(sheet);
  const sourceTransactions = lireTransactionsGericoDepuisAS_(source);

  const existing = construireIndexGericoExistant_(sheet, zoneAvant);
  const latestDate = existing.latestDate;
  const existingKeys = existing.keys;
  const toAdd = [];

  sourceTransactions.forEach(t => {
    // Le besoin est d'ajouter les nouvelles opérations à la suite de l'historique existant.
    // On ne tente pas de reconstruire rétroactivement les anciennes années qui utilisent
    // parfois un ancien format de colonnes.
    if (latestDate && t.date.getTime() < latestDate.getTime()) return;
    const key = construireCleGerico_(t.date, t.extract, t.counterpartyAccount, t.amount, t.description, t.communication);
    if (!existingKeys.has(key)) {
      existingKeys.add(key);
      toAdd.push(t);
    }
  });

  toAdd.sort(comparerTransactions_);
  const grouped = {};
  toAdd.forEach(t => {
    const y = t.date.getFullYear();
    (grouped[y] || (grouped[y] = [])).push(t);
  });

  let added = 0;
  Object.keys(grouped).map(Number).sort((a, b) => a - b).forEach(year => {
    added += insererTransactionsGericoAnnee_(sheet, year, grouped[year]);
  });

  // Même sans nouvelle opération, on garde le décompte annuel vivant :
  // B de l'année courante = n° du mois, et B du précompte = mois/12.
  mettreAJourDecompteAnnuelGerico_(sheet, currentYear, currentMonth);
  SpreadsheetApp.flush();

  if (!options.silent) {
    ss.toast('GERICO : ' + added + ' nouvelle(s) ligne(s).', 'CBC BANQUE', 5);
  }

  return {
    added: added,
    lastYear: Math.max(currentYear, dernierAnneeDonneesGerico_(sheet) || currentYear),
    currentMonth: currentMonth
  };
}

function trouverOngletGerico_(ss) {
  for (let i = 0; i < GERICO_CFG.SHEET_NAMES.length; i++) {
    const exact = ss.getSheetByName(GERICO_CFG.SHEET_NAMES[i]);
    if (exact) return exact;
  }
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = String(sheets[i].getName() || '').trim().toUpperCase();
    if (n === 'GERICO' || n === 'GERICO-') return sheets[i];
  }
  return null;
}

function lireTransactionsGericoDepuisAS_(source) {
  const lastOp = derniereLigneOperation_(source);
  if (lastOp < CBC.DATA_START_ROW) return [];

  const values = source.getRange(CBC.DATA_START_ROW, 1, lastOp - CBC.DATA_START_ROW + 1, 11).getValues();
  const out = [];
  const ibans = new Set(GERICO_CFG.COUNTERPARTY_IBANS.map(normaliserIban_));
  const nameNeedles = GERICO_CFG.NAME_CONTAINS.map(normaliserRecherche_);

  values.forEach((r, idx) => {
    const date = convertirDate_(r[0]);
    const amount = convertirNombre_(r[8]);
    if (!date || amount === null) return;

    const cpIban = normaliserIban_(r[3]);
    const nameText = normaliserRecherche_([r[4], r[5], r[6]].join(' '));
    const isGerico = (cpIban && ibans.has(cpIban)) || nameNeedles.some(n => n && nameText.indexOf(n) !== -1);
    if (!isGerico) return;

    out.push({
      date: date,
      extract: nettoyerTexte_(r[1]),
      account: normaliserIban_(r[2]) || nettoyerTexte_(r[2]),
      counterpartyAccount: nettoyerTexte_(r[3]),
      counterpartyName: nettoyerTexte_(r[4]),
      description: nettoyerTexte_(r[5]),
      communication: nettoyerTexte_(r[6]),
      amount: arrondir2_(amount),
      sourceCsvLine: CBC.DATA_START_ROW + idx,
      sourceRow: CBC.DATA_START_ROW + idx
    });
  });

  return out;
}

function analyserZoneGerico_(sheet) {
  const max = Math.max(sheet.getLastRow(), 1);
  const vals = sheet.getRange(1, 1, max, 11).getValues();
  const formulas = sheet.getRange(1, 1, max, 11).getFormulas();

  let decompteRow = 0;
  for (let i = 0; i < vals.length; i++) {
    if (normaliserRecherche_(vals[i][0]) === 'decompte') {
      decompteRow = i + 1;
      break;
    }
  }
  if (!decompteRow) throw new Error('Impossible de repérer la zone "Decompte" dans GERICO.');

  const dataRows = [];
  const years = {};
  let latestDate = null;
  for (let r = GERICO_CFG.DATA_START_ROW; r < decompteRow; r++) {
    const row = vals[r - 1];
    const date = convertirDate_(row[0]);
    const amount = convertirNombre_(row[8]);
    if (!date || amount === null) continue;
    dataRows.push(r);
    const y = date.getFullYear();
    (years[y] || (years[y] = [])).push(r);
    if (!latestDate || date > latestDate) latestDate = date;
  }

  let grandTotalRow = 0;
  for (let r = decompteRow - 1; r >= GERICO_CFG.DATA_START_ROW; r--) {
    const fi = String(formulas[r - 1][8] || '').toUpperCase(); // col I
    if (fi.indexOf('SUM(') !== -1) {
      grandTotalRow = r;
      break;
    }
  }
  if (!grandTotalRow) grandTotalRow = Math.max(GERICO_CFG.DATA_START_ROW, decompteRow - 3);

  return {
    decompteRow: decompteRow,
    grandTotalRow: grandTotalRow,
    dataRows: dataRows,
    years: years,
    latestDate: latestDate,
    values: vals,
    formulas: formulas
  };
}

function construireIndexGericoExistant_(sheet, zone) {
  const keys = new Set();
  let latestDate = null;
  const decompteRow = zone.decompteRow;
  if (decompteRow <= GERICO_CFG.DATA_START_ROW) return { keys: keys, latestDate: null };

  const vals = sheet.getRange(GERICO_CFG.DATA_START_ROW, 1, decompteRow - GERICO_CFG.DATA_START_ROW, 9).getValues();
  vals.forEach(r => {
    const date = convertirDate_(r[0]);
    const amount = convertirNombre_(r[8]);
    if (!date || amount === null) return;
    if (!latestDate || date > latestDate) latestDate = date;
    keys.add(construireCleGerico_(date, r[1], r[3], amount, r[5], r[6]));
  });
  return { keys: keys, latestDate: latestDate };
}

function construireCleGerico_(date, extract, cpAccount, amount, description, communication) {
  return [
    cleDate_(date),
    normaliserExtraitPourCle_(extract),
    normaliserCompteRecherche_(cpAccount),
    cents_(amount),
    normaliserRecherche_((description || '') + ' ' + (communication || ''))
  ].join('|');
}

function insererTransactionsGericoAnnee_(sheet, year, transactions) {
  if (!transactions || !transactions.length) return 0;

  let zone = analyserZoneGerico_(sheet);
  const yearRows = zone.years[year] || [];
  let startRow;

  if (yearRows.length) {
    const lastData = Math.max.apply(null, yearRows);
    startRow = lastData + 1;
    // Utilise d'abord les lignes vides déjà présentes. Si elles ne suffisent pas,
    // on insère juste avant le premier contenu suivant pour ne rien écraser.
    let free = 0;
    let r = startRow;
    while (r < zone.decompteRow && ligneGericoVide_(sheet, r)) {
      free++;
      r++;
      if (free >= transactions.length) break;
    }
    if (free < transactions.length) {
      const insertBefore = startRow + free;
      sheet.insertRowsBefore(insertBefore, transactions.length - free);
    }
  } else {
    // Nouvelle année : on laisse au minimum une ligne vide après la dernière année,
    // puis on crée les transactions et une ligne de sous-total annuel.
    zone = analyserZoneGerico_(sheet);
    const allDataRows = zone.dataRows;
    const lastData = allDataRows.length ? Math.max.apply(null, allDataRows) : GERICO_CFG.DATA_START_ROW - 1;

    // Cherche une éventuelle ligne de sous-total J juste après la dernière année.
    let endOfPreviousYear = lastData;
    for (let r = lastData + 1; r < zone.grandTotalRow; r++) {
      const f = String(sheet.getRange(r, 10).getFormula() || '').toUpperCase();
      if (f.indexOf('SUM(I') !== -1) {
        endOfPreviousYear = r;
        break;
      }
      if (!ligneGericoVide_(sheet, r)) break;
    }

    startRow = endOfPreviousYear + 2; // 1 ligne vide entre les années
    const neededEnd = startRow + transactions.length; // +1 pour sous-total
    if (neededEnd >= zone.grandTotalRow) {
      sheet.insertRowsBefore(zone.grandTotalRow, neededEnd - zone.grandTotalRow + 1);
    }
  }

  // Reprend le format de la dernière ligne bancaire précédente.
  const formatSourceRow = trouverDerniereLigneDonneeAvant_(sheet, startRow - 1);
  if (formatSourceRow >= GERICO_CFG.DATA_START_ROW) {
    try {
      sheet.getRange(formatSourceRow, 1, 1, GERICO_CFG.DATA_WIDTH)
        .copyFormatToRange(sheet.getSheetId(), 1, GERICO_CFG.DATA_WIDTH, startRow, startRow + transactions.length - 1);
    } catch (e) {
      console.warn('Format GERICO : ' + e.message);
    }
  }

  const rows = transactions.map(t => [
    t.date, t.extract, t.account, t.counterpartyAccount, t.counterpartyName,
    t.description, t.communication, '', t.amount
  ]);
  sheet.getRange(startRow, 1, rows.length, GERICO_CFG.DATA_WIDTH).setValues(rows);
  sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(startRow, 9, rows.length, 1).setNumberFormat('#,##0.00');

  // Sous-total de l'année dans la colonne J.
  zone = analyserZoneGerico_(sheet);
  const rowsThisYear = zone.years[year] || [];
  if (rowsThisYear.length) {
    const first = Math.min.apply(null, rowsThisYear);
    const last = Math.max.apply(null, rowsThisYear);
    let subtotalRow = last + 1;
    // Si la ligne juste après est déjà un sous-total, on la réutilise.
    const existingFormula = String(sheet.getRange(subtotalRow, 10).getFormula() || '').toUpperCase();
    if (!existingFormula && !ligneGericoVide_(sheet, subtotalRow)) {
      sheet.insertRowsBefore(subtotalRow, 1);
    }
    sheet.getRange(subtotalRow, 10).setFormula('=SUM(I' + first + ':I' + last + ')').setNumberFormat('#,##0.00');
  }

  return transactions.length;
}

function ligneGericoVide_(sheet, row) {
  if (row < 1 || row > sheet.getMaxRows()) return false;
  const vals = sheet.getRange(row, 1, 1, 11).getDisplayValues()[0];
  const formulas = sheet.getRange(row, 1, 1, 11).getFormulas()[0];
  return vals.every(v => String(v || '').trim() === '') && formulas.every(v => !v);
}

function trouverDerniereLigneDonneeAvant_(sheet, row) {
  for (let r = row; r >= GERICO_CFG.DATA_START_ROW; r--) {
    const a = sheet.getRange(r, 1).getValue();
    const i = sheet.getRange(r, 9).getValue();
    if (convertirDate_(a) && convertirNombre_(i) !== null) return r;
  }
  return 0;
}

function dernierAnneeDonneesGerico_(sheet) {
  const zone = analyserZoneGerico_(sheet);
  const years = Object.keys(zone.years).map(Number);
  return years.length ? Math.max.apply(null, years) : 0;
}

function mettreAJourDecompteAnnuelGerico_(sheet, currentYear, currentMonth) {
  // Assure les deux lignes annuelles, y compris lors d'un changement d'année.
  assurerLignesAnnuellesGerico_(sheet, currentYear, currentMonth);

  const occ = trouverOccurrencesAnneesGerico_(sheet);
  const first = occ.first[currentYear];
  const second = occ.second[currentYear];
  if (!first || !second) throw new Error('Impossible de repérer les deux lignes de décompte GERICO pour ' + currentYear + '.');

  // Première zone (loyer) : B = mois courant.
  sheet.getRange(first, 2).setValue(currentMonth);
  sheet.getRange(first, 4).setFormula('=B' + first + '*C' + first).setNumberFormat('#,##0.00');
  sheet.getRange(first, 5).setFormula('=D' + second).setNumberFormat('#,##0.00');
  sheet.getRange(first, 7).setFormula('=D' + first + '+E' + first).setNumberFormat('#,##0.00');
  sheet.getRange(first, 9)
    .setValue(totalGericoAnnee_(sheet, currentYear))
    .setNumberFormat('#,##0.00');
  sheet.getRange(first, 10).setFormula('=G' + first + '-I' + first).setNumberFormat('#,##0.00');

  // Deuxième zone (précompte) : B = mois / 12.
  sheet.getRange(second, 2).setFormula('=' + currentMonth + '/12');
  sheet.getRange(second, 4).setFormula('=B' + second + '*C' + second).setNumberFormat('#,##0.00');

  // Si une nouvelle année vient d'être créée, l'année précédente devient complète.
  const prev = currentYear - 1;
  if (occ.first[prev] && prev >= 2018) sheet.getRange(occ.first[prev], 2).setValue(12);
  if (occ.second[prev] && prev >= 2018) sheet.getRange(occ.second[prev], 2).setValue(1);

  // Répare les totaux des blocs après éventuelles insertions de lignes.
  reparerTotauxDecompteGerico_(sheet);
}

function totalGericoAnnee_(sheet, year) {
  const zone = analyserZoneGerico_(sheet);
  const rows = zone.years[year] || [];
  if (!rows.length) return 0;
  let total = 0;
  rows.forEach(r => {
    const n = convertirNombre_(sheet.getRange(r, 9).getValue());
    if (n !== null) total += n;
  });
  return arrondir2_(total);
}

function assurerLignesAnnuellesGerico_(sheet, currentYear, currentMonth) {
  let occ = trouverOccurrencesAnneesGerico_(sheet);
  let years = Object.keys(occ.first).map(Number);
  let lastYear = years.length ? Math.max.apply(null, years) : currentYear - 1;

  for (let year = lastYear + 1; year <= currentYear; year++) {
    // 1) ajoute la nouvelle ligne dans le premier bloc juste après la dernière année.
    occ = trouverOccurrencesAnneesGerico_(sheet);
    years = Object.keys(occ.first).map(Number);
    const prevYear = years.length ? Math.max.apply(null, years) : year - 1;
    const prevFirst = occ.first[prevYear];
    if (!prevFirst) throw new Error('Bloc annuel GERICO introuvable avant ' + year + '.');

    sheet.insertRowsAfter(prevFirst, 1);
    const newFirst = prevFirst + 1;
    sheet.getRange(prevFirst, 1, 1, 10).copyFormatToRange(sheet.getSheetId(), 1, 10, newFirst, newFirst);
    const prevRent = convertirNombre_(sheet.getRange(prevFirst, 3).getValue()) || 0;
    sheet.getRange(newFirst, 1).setValue('année ' + year);
    sheet.getRange(newFirst, 2).setValue(year === currentYear ? currentMonth : 12);
    sheet.getRange(newFirst, 3).setValue(prevRent);
    sheet.getRange(newFirst, 4).setFormula('=B' + newFirst + '*C' + newFirst);
    sheet.getRange(newFirst, 6).clearContent();
    sheet.getRange(newFirst, 7).setFormula('=D' + newFirst + '+E' + newFirst);
    sheet.getRange(newFirst, 8).clearContent();
    sheet.getRange(newFirst, 9).setValue(totalGericoAnnee_(sheet, year));
    sheet.getRange(newFirst, 10).setFormula('=G' + newFirst + '-I' + newFirst);
    sheet.getRange(newFirst, 4, 1, 7).setNumberFormat('#,##0.00');

    // 2) recalcule les positions puis ajoute la ligne dans le second bloc.
    occ = trouverOccurrencesAnneesGerico_(sheet);
    const prevSecond = occ.second[prevYear];
    if (!prevSecond) throw new Error('Deuxième bloc annuel GERICO introuvable avant ' + year + '.');

    sheet.insertRowsAfter(prevSecond, 1);
    const newSecond = prevSecond + 1;
    sheet.getRange(prevSecond, 1, 1, 10).copyTo(sheet.getRange(newSecond, 1, 1, 10));
    sheet.getRange(newSecond, 1).setValue('année-' + year);
    sheet.getRange(newSecond, 2).setFormula('=' + (year === currentYear ? currentMonth : 12) + '/12');

    // La copie relative conserve la logique du modèle : par exemple C2027 = C2026 * 1,03
    // si la ligne précédente utilisait déjà cette formule.
    if (!sheet.getRange(newSecond, 3).getFormula()) {
      const prevC = convertirNombre_(sheet.getRange(prevSecond, 3).getValue()) || 0;
      sheet.getRange(newSecond, 3).setValue(prevC);
    }
    sheet.getRange(newSecond, 4).setFormula('=B' + newSecond + '*C' + newSecond).setNumberFormat('#,##0.00');

    // 3) lie le précompte à la première ligne de la même année.
    occ = trouverOccurrencesAnneesGerico_(sheet);
    if (occ.first[year] && occ.second[year]) {
      sheet.getRange(occ.first[year], 5).setFormula('=D' + occ.second[year]);
    }
  }
}

function trouverOccurrencesAnneesGerico_(sheet) {
  const max = Math.min(sheet.getLastRow(), 800);
  const vals = sheet.getRange(1, 1, max, 1).getDisplayValues();
  const matches = {};
  vals.forEach((r, i) => {
    const m = normaliserRecherche_(r[0]).match(/^annee\s*-?\s*(20\d{2}|19\d{2})$/);
    if (!m) return;
    const y = Number(m[1]);
    (matches[y] || (matches[y] = [])).push(i + 1);
  });

  const first = {};
  const second = {};
  Object.keys(matches).forEach(y => {
    first[y] = matches[y][0];
    if (matches[y].length > 1) second[y] = matches[y][1];
  });
  return { first: first, second: second, all: matches };
}

function reparerTotauxDecompteGerico_(sheet) {
  const occ = trouverOccurrencesAnneesGerico_(sheet);
  const yearsFirst = Object.keys(occ.first).map(Number).sort((a, b) => a - b);
  const yearsSecond = Object.keys(occ.second).map(Number).sort((a, b) => a - b);
  if (!yearsFirst.length || !yearsSecond.length) return;

  const firstStart = occ.first[yearsFirst[0]];
  const firstEnd = occ.first[yearsFirst[yearsFirst.length - 1]];
  const secondStart = occ.second[yearsSecond[0]];
  const secondEnd = occ.second[yearsSecond[yearsSecond.length - 1]];

  // Ligne de total du premier bloc : première ligne après le bloc ayant une formule SUM en G/I/J.
  for (let r = firstEnd + 1; r < secondStart; r++) {
    const vals = sheet.getRange(r, 1, 1, 10).getDisplayValues()[0];
    const formulas = sheet.getRange(r, 1, 1, 10).getFormulas()[0];
    if (formulas[6] || formulas[8] || formulas[9] || vals.some(v => String(v || '').trim() !== '')) {
      if (formulas[6] || formulas[8] || formulas[9]) {
        sheet.getRange(r, 7).setFormula('=SUM(G' + firstStart + ':G' + firstEnd + ')').setNumberFormat('#,##0.00');
        sheet.getRange(r, 9).setFormula('=SUM(I' + firstStart + ':I' + firstEnd + ')').setNumberFormat('#,##0.00');
        sheet.getRange(r, 10).setFormula('=SUM(J' + firstStart + ':J' + firstEnd + ')').setNumberFormat('#,##0.00');
      }
      break;
    }
  }

  // "Total du" situé après le second bloc.
  const max = Math.min(sheet.getLastRow(), secondEnd + 15);
  const labels = sheet.getRange(secondEnd + 1, 1, Math.max(1, max - secondEnd), 1).getDisplayValues();
  for (let i = 0; i < labels.length; i++) {
    if (normaliserRecherche_(labels[i][0]).indexOf('total du') === 0) {
      const r = secondEnd + 1 + i;
      sheet.getRange(r, 4).setFormula('=SUM(D' + firstStart + ':D' + secondEnd + ')').setNumberFormat('#,##0.00');
      break;
    }
  }
}

function anneeCouranteBelgique_() {
  return Number(Utilities.formatDate(new Date(), CBC.TIMEZONE, 'yyyy'));
}

function moisCourantBelgique_() {
  return Number(Utilities.formatDate(new Date(), CBC.TIMEZONE, 'M'));
}

/* ========================================================================== */
/* RECHERCHE AVANCÉE                                                         */
/* ========================================================================== */

function ouvrirRechercheAvancee() {
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; margin: 0; color: #202124; background:#f8f9fa; }
    .wrap { padding: 16px; }
    h2 { margin: 0 0 4px; font-size: 20px; }
    .sub { color:#5f6368; font-size:12px; margin-bottom:12px; }
    .grid { display:grid; grid-template-columns: 1fr 1fr 1fr; gap:10px; }
    .card { background:white; border:1px solid #dadce0; border-radius:8px; padding:12px; }
    .card.full { grid-column: 1 / -1; }
    label { display:block; font-size:12px; color:#5f6368; margin:0 0 4px; }
    input, select { width:100%; box-sizing:border-box; padding:8px 9px; border:1px solid #dadce0; border-radius:6px; background:#fff; }
    .row2 { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
    .sheets { display:grid; grid-template-columns:repeat(4,1fr); gap:6px 10px; max-height:145px; overflow:auto; padding:4px; }
    .check { display:flex; align-items:center; gap:6px; font-size:12px; }
    .check input { width:auto; }
    .toolbar { display:flex; gap:8px; margin-bottom:8px; }
    button { border:1px solid #dadce0; background:white; border-radius:6px; padding:8px 12px; cursor:pointer; }
    button.primary { background:#1a73e8; border-color:#1a73e8; color:white; font-weight:600; }
    button.secondary { background:#e8f0fe; border-color:#d2e3fc; color:#174ea6; }
    .footer { display:flex; align-items:center; gap:8px; margin-top:12px; }
    .status { flex:1; font-size:12px; color:#5f6368; }
    .fields { display:flex; flex-wrap:wrap; gap:8px; margin-top:6px; }
    .fieldCheck { display:flex; align-items:center; gap:4px; font-size:12px; }
    .fieldCheck input { width:auto; }
    .examples { font-size:11px; color:#5f6368; line-height:1.45; margin-top:8px; }
    .loading { opacity:.65; pointer-events:none; }
  </style>
</head>
<body>
<div class="wrap" id="wrap">
  <h2>Recherche avancée CBC</h2>
  <div class="sub">Combine plusieurs critères : mot, compte, nom, année, dates, montant, entrées/sorties et onglets.</div>

  <div class="grid">
    <div class="card full">
      <label>Onglets à rechercher</label>
      <div class="toolbar">
        <button type="button" onclick="selectAll(true)">Tout sélectionner</button>
        <button type="button" onclick="selectAll(false)">Tout enlever</button>
        <button type="button" class="secondary" onclick="selectBankSheets()">Onglets bancaires</button>
      </div>
      <div id="sheets" class="sheets"></div>
    </div>

    <div class="card">
      <label>Texte / mot / phrase</label>
      <input id="text" placeholder="ex. GERICO, loyer, Construct Center…">
      <div class="row2" style="margin-top:8px">
        <select id="textMode">
          <option value="CONTAINS">Contient</option>
          <option value="EXACT">Exact</option>
          <option value="STARTS">Commence par</option>
        </select>
        <select id="caseMode">
          <option value="IGNORE">Ignorer majuscules/accents</option>
          <option value="CASE">Respecter majuscules</option>
        </select>
      </div>
      <label style="margin-top:8px">Chercher ce texte dans</label>
      <div class="fields" id="textFields">
        <label class="fieldCheck"><input type="checkbox" value="ALL" checked>Tous champs</label>
        <label class="fieldCheck"><input type="checkbox" value="NAME">Nom</label>
        <label class="fieldCheck"><input type="checkbox" value="ACCOUNT">Compte</label>
        <label class="fieldCheck"><input type="checkbox" value="DESCRIPTION">Description</label>
        <label class="fieldCheck"><input type="checkbox" value="COMMUNICATION">Communication</label>
        <label class="fieldCheck"><input type="checkbox" value="EXTRACT">N° extrait</label>
      </div>
    </div>

    <div class="card">
      <label>Numéro de compte (source ou contrepartie)</label>
      <input id="account" placeholder="BE34 0689… ou BE34">
      <select id="accountMode" style="margin-top:8px">
        <option value="CONTAINS">Contient</option>
        <option value="EXACT">Exact</option>
      </select>
      <label style="margin-top:10px">Nom / contrepartie contient</label>
      <input id="name" placeholder="ex. GERICO INDUSTRY">
    </div>

    <div class="card">
      <label>Année</label>
      <input id="year" type="number" min="2000" max="2100" placeholder="ex. 2026">
      <div class="row2" style="margin-top:8px">
        <div><label>Date du</label><input id="dateFrom" type="date"></div>
        <div><label>Date au</label><input id="dateTo" type="date"></div>
      </div>
    </div>

    <div class="card">
      <label>Type de mouvement</label>
      <select id="direction">
        <option value="ALL">Toutes les opérations</option>
        <option value="IN">Entrées uniquement</option>
        <option value="OUT">Sorties uniquement</option>
      </select>
      <label style="margin-top:10px">Mode montant</label>
      <select id="amountMode">
        <option value="ABS">Valeur absolue (recommandé)</option>
        <option value="SIGNED">Montant signé</option>
      </select>
    </div>

    <div class="card">
      <label>Montant minimum</label>
      <input id="minAmount" placeholder="ex. 500">
      <label style="margin-top:10px">Montant maximum</label>
      <input id="maxAmount" placeholder="ex. 2500">
    </div>

    <div class="card">
      <label>Tri des résultats</label>
      <select id="sort">
        <option value="DATE_DESC">Date : plus récent d'abord</option>
        <option value="DATE_ASC">Date : plus ancien d'abord</option>
        <option value="AMOUNT_DESC">Montant : plus grand d'abord</option>
        <option value="AMOUNT_ASC">Montant : plus petit d'abord</option>
      </select>
      <div class="examples">
        Exemples : <b>"loyer" + année 2026</b> ; <b>compte BE34… + entrées</b> ;
        <b>"Construct" + montant 500 à 2 500 + tous les onglets</b>.
      </div>
    </div>
  </div>

  <div class="footer">
    <div id="status" class="status">Chargement des onglets…</div>
    <button onclick="google.script.host.close()">Fermer</button>
    <button class="primary" id="searchBtn" onclick="runSearch()">Rechercher</button>
  </div>
</div>

<script>
let config = null;
function esc(s){ return String(s).replace(/[&<>\"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c])); }
function load(){
  google.script.run.withSuccessHandler(c => {
    config = c;
    const box = document.getElementById('sheets');
    box.innerHTML = c.sheets.map(s => '<label class="check"><input class="sheetChk" type="checkbox" value="'+esc(s.name)+'" '+(s.defaultSelected?'checked':'')+'><span>'+esc(s.name)+'</span></label>').join('');
    if (c.lastCriteria) restore(c.lastCriteria);
    document.getElementById('status').textContent = c.sheets.length + ' onglet(s) disponibles.';
  }).withFailureHandler(showError).cbcRechercheGetConfig();
}
function restore(c){
  ['text','account','name','year','dateFrom','dateTo','minAmount','maxAmount'].forEach(id => { if(c[id] !== undefined && c[id] !== null) document.getElementById(id).value = c[id]; });
  ['textMode','caseMode','accountMode','direction','amountMode','sort'].forEach(id => { if(c[id]) document.getElementById(id).value = c[id]; });
  if (Array.isArray(c.sheets) && c.sheets.length) document.querySelectorAll('.sheetChk').forEach(x => x.checked = c.sheets.includes(x.value));
  if (Array.isArray(c.textFields) && c.textFields.length) document.querySelectorAll('#textFields input').forEach(x => x.checked = c.textFields.includes(x.value));
}
function selectAll(v){ document.querySelectorAll('.sheetChk').forEach(x => x.checked=v); }
function selectBankSheets(){ document.querySelectorAll('.sheetChk').forEach(x => x.checked = !!(config.bankSheets||[]).includes(x.value)); }
function collect(){
  return {
    sheets:[...document.querySelectorAll('.sheetChk:checked')].map(x=>x.value),
    text:document.getElementById('text').value.trim(),
    textMode:document.getElementById('textMode').value,
    caseMode:document.getElementById('caseMode').value,
    textFields:[...document.querySelectorAll('#textFields input:checked')].map(x=>x.value),
    account:document.getElementById('account').value.trim(),
    accountMode:document.getElementById('accountMode').value,
    name:document.getElementById('name').value.trim(),
    year:document.getElementById('year').value.trim(),
    dateFrom:document.getElementById('dateFrom').value,
    dateTo:document.getElementById('dateTo').value,
    direction:document.getElementById('direction').value,
    amountMode:document.getElementById('amountMode').value,
    minAmount:document.getElementById('minAmount').value.trim(),
    maxAmount:document.getElementById('maxAmount').value.trim(),
    sort:document.getElementById('sort').value
  };
}
function runSearch(){
  const c=collect();
  if(!c.sheets.length){ document.getElementById('status').textContent='Sélectionne au moins un onglet.'; return; }
  document.getElementById('wrap').classList.add('loading');
  document.getElementById('status').textContent='Recherche en cours…';
  google.script.run.withSuccessHandler(r=>{
    document.getElementById('wrap').classList.remove('loading');
    document.getElementById('status').textContent=r.count+' résultat(s). Onglet '+r.sheetName+' affiché.';
  }).withFailureHandler(showError).executerRechercheAvancee(c);
}
function showError(e){ document.getElementById('wrap').classList.remove('loading'); document.getElementById('status').textContent='Erreur : '+(e.message||e); }
load();
</script>
</body>
</html>`).setWidth(980).setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'CBC — Recherche avancée');
}

function cbcRechercheGetConfig() {
  const ss = getSpreadsheet_();
  const excluded = new Set([CBC.SEARCH_SHEET, CBC.JOURNAL_SHEET]);
  const bankSet = new Set([CBC.DIVERS_SHEET, 'GERICO', 'GERICO-', 'ALFANO-CBC-Epargne']
    .concat(CBC.ACCOUNTS.map(a => a.sheet)));
  const sheets = ss.getSheets().filter(s => !excluded.has(s.getName())).map(s => ({
    name: s.getName(),
    defaultSelected: bankSet.has(s.getName())
  }));

  let lastCriteria = null;
  try {
    const raw = PropertiesService.getUserProperties().getProperty(CBC.PROP_SEARCH_CRITERIA);
    if (raw) lastCriteria = JSON.parse(raw);
  } catch (e) {}

  return {
    sheets: sheets,
    bankSheets: sheets.filter(s => bankSet.has(s.name)).map(s => s.name),
    lastCriteria: lastCriteria
  };
}

function executerRechercheAvancee(criteria) {
  criteria = criteria || {};
  const ss = getSpreadsheet_();
  const selected = Array.isArray(criteria.sheets) ? criteria.sheets : [];
  if (!selected.length) throw new Error('Sélectionne au moins un onglet.');

  PropertiesService.getUserProperties().setProperty(CBC.PROP_SEARCH_CRITERIA, JSON.stringify(criteria));

  const parsed = preparerCriteresRecherche_(criteria);
  let results = [];
  const maxResults = 10000;

  selected.forEach(sheetName => {
    if (results.length >= maxResults) return;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getName() === CBC.SEARCH_SHEET) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const width = Math.max(11, Math.min(sheet.getLastColumn(), 18));
    const vals = sheet.getRange(1, 1, lastRow, width).getValues();

    for (let i = 1; i < vals.length && results.length < maxResults; i++) {
      const r = vals[i];
      const date = convertirDate_(r[0]);
      const amount = convertirNombre_(r[8]);
      if (!date || amount === null) continue; // ligne bancaire réelle uniquement

      const rec = {
        sheet: sheet.getName(),
        sheetId: sheet.getSheetId(),
        row: i + 1,
        date: date,
        extract: valeurRecherche_(r[1]),
        account: valeurRecherche_(r[2]),
        counterpartyAccount: valeurRecherche_(r[3]),
        name: valeurRecherche_(r[4]),
        description: valeurRecherche_(r[5]),
        communication: valeurRecherche_(r[6]),
        amount: arrondir2_(amount),
        balance: convertirNombre_(r[10])
      };

      if (!correspondAuxCriteres_(rec, parsed)) continue;
      results.push(rec);
    }
  });

  trierResultatsRecherche_(results, parsed.sort);
  ecrireResultatsRecherche_(ss, results, criteria, maxResults);
  return { count: results.length, sheetName: CBC.SEARCH_SHEET, limited: results.length >= maxResults };
}

function preparerCriteresRecherche_(c) {
  const year = Number(c.year) || 0;
  let dateFrom = convertirDate_(c.dateFrom);
  let dateTo = convertirDate_(c.dateTo);
  if (dateTo) dateTo = new Date(dateTo.getFullYear(), dateTo.getMonth(), dateTo.getDate(), 23, 59, 59, 999);

  return {
    text: String(c.text || ''),
    textNorm: normaliserRecherche_(c.text || ''),
    textMode: c.textMode || 'CONTAINS',
    caseMode: c.caseMode || 'IGNORE',
    textFields: Array.isArray(c.textFields) && c.textFields.length ? c.textFields : ['ALL'],
    account: String(c.account || ''),
    accountNorm: normaliserCompteRecherche_(c.account || ''),
    accountMode: c.accountMode || 'CONTAINS',
    nameNorm: normaliserRecherche_(c.name || ''),
    year: year,
    dateFrom: dateFrom,
    dateTo: dateTo,
    direction: c.direction || 'ALL',
    amountMode: c.amountMode || 'ABS',
    minAmount: convertirNombre_(c.minAmount),
    maxAmount: convertirNombre_(c.maxAmount),
    sort: c.sort || 'DATE_DESC'
  };
}

function correspondAuxCriteres_(rec, c) {
  if (c.year && rec.date.getFullYear() !== c.year) return false;
  if (c.dateFrom && rec.date < c.dateFrom) return false;
  if (c.dateTo && rec.date > c.dateTo) return false;

  if (c.direction === 'IN' && !(rec.amount > 0)) return false;
  if (c.direction === 'OUT' && !(rec.amount < 0)) return false;

  const amountToTest = c.amountMode === 'SIGNED' ? rec.amount : Math.abs(rec.amount);
  if (c.minAmount !== null && c.minAmount !== undefined && amountToTest < c.minAmount) return false;
  if (c.maxAmount !== null && c.maxAmount !== undefined && amountToTest > c.maxAmount) return false;

  if (c.accountNorm) {
    const accounts = [rec.account, rec.counterpartyAccount].map(normaliserCompteRecherche_);
    const ok = c.accountMode === 'EXACT'
      ? accounts.some(x => x === c.accountNorm)
      : accounts.some(x => x.indexOf(c.accountNorm) !== -1);
    if (!ok) return false;
  }

  if (c.nameNorm && normaliserRecherche_(rec.name).indexOf(c.nameNorm) === -1) return false;

  if (c.text) {
    let fields = [];
    const tf = c.textFields || ['ALL'];
    if (tf.indexOf('ALL') !== -1) {
      fields = [rec.extract, rec.account, rec.counterpartyAccount, rec.name, rec.description, rec.communication];
    } else {
      if (tf.indexOf('NAME') !== -1) fields.push(rec.name);
      if (tf.indexOf('ACCOUNT') !== -1) fields.push(rec.account, rec.counterpartyAccount);
      if (tf.indexOf('DESCRIPTION') !== -1) fields.push(rec.description);
      if (tf.indexOf('COMMUNICATION') !== -1) fields.push(rec.communication);
      if (tf.indexOf('EXTRACT') !== -1) fields.push(rec.extract);
    }
    const okText = fields.some(v => texteCorrespond_(v, c));
    if (!okText) return false;
  }

  return true;
}

function texteCorrespond_(value, c) {
  const original = String(value === null || value === undefined ? '' : value);
  const hay = c.caseMode === 'CASE' ? original : normaliserRecherche_(original);
  const needle = c.caseMode === 'CASE' ? c.text : c.textNorm;
  if (!needle) return true;
  if (c.textMode === 'EXACT') return hay === needle;
  if (c.textMode === 'STARTS') return hay.indexOf(needle) === 0;
  return hay.indexOf(needle) !== -1;
}

function trierResultatsRecherche_(results, sort) {
  results.sort((a, b) => {
    if (sort === 'DATE_ASC') return a.date - b.date || a.row - b.row;
    if (sort === 'AMOUNT_DESC') return Math.abs(b.amount) - Math.abs(a.amount) || b.date - a.date;
    if (sort === 'AMOUNT_ASC') return Math.abs(a.amount) - Math.abs(b.amount) || b.date - a.date;
    return b.date - a.date || b.row - a.row;
  });
}

function ecrireResultatsRecherche_(ss, results, criteria, maxResults) {
  let sheet = ss.getSheetByName(CBC.SEARCH_SHEET);
  if (!sheet) sheet = ss.insertSheet(CBC.SEARCH_SHEET);

  const oldFilter = sheet.getFilter();
  if (oldFilter) oldFilter.remove();
  try { sheet.getBandings().forEach(b => b.remove()); } catch (e) {}
  try { sheet.getRange(1, 1, sheet.getMaxRows(), Math.min(15, sheet.getMaxColumns())).breakApart(); } catch (e) {}
  sheet.clear();
  assurerNombreLignes_(sheet, Math.max(100, results.length + 10));
  if (sheet.getMaxColumns() < 15) sheet.insertColumnsAfter(sheet.getMaxColumns(), 15 - sheet.getMaxColumns());

  sheet.getRange('A1:O1').merge();
  sheet.getRange('A1').setValue('RECHERCHE AVANCÉE CBC — RÉSULTATS')
    .setFontWeight('bold').setFontSize(14).setFontColor('#ffffff').setBackground('#1a73e8')
    .setHorizontalAlignment('center');

  const summary = resumeCriteresRecherche_(criteria) + (results.length >= maxResults ? ' — LIMITE : ' + maxResults + ' résultats' : '');
  sheet.getRange('A2:O2').merge();
  sheet.getRange('A2').setValue(summary).setWrap(true).setBackground('#e8f0fe').setFontColor('#174ea6');

  const headers = ['Onglet', 'Ligne', 'Date', 'N° extrait', 'Compte', 'Compte contrepartie', 'Nom contrepartie', 'Description', 'Communication', 'Entrée', 'Sortie', 'Montant signé', 'Solde', 'Année', 'Ouvrir'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center');

  if (results.length) {
    const values = results.map(r => [
      r.sheet, r.row, r.date, r.extract, r.account, r.counterpartyAccount, r.name,
      r.description, r.communication,
      r.amount > 0 ? r.amount : '',
      r.amount < 0 ? Math.abs(r.amount) : '',
      r.amount,
      r.balance === null ? '' : r.balance,
      r.date.getFullYear(),
      ''
    ]);
    sheet.getRange(5, 1, values.length, 15).setValues(values);
    sheet.getRange(5, 3, values.length, 1).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(5, 10, values.length, 4).setNumberFormat('#,##0.00');

    const formulas = results.map(r => ['=HYPERLINK("#gid=' + r.sheetId + '&range=A' + r.row + '","Ouvrir")']);
    sheet.getRange(5, 15, formulas.length, 1).setFormulas(formulas);

    sheet.getRange(4, 1, results.length + 1, 15).createFilter();
    try { sheet.getRange(5, 1, results.length, 15).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) {}
  } else {
    sheet.getRange('A5:O5').merge();
    sheet.getRange('A5').setValue('Aucun résultat avec ces critères.').setHorizontalAlignment('center').setFontStyle('italic');
  }

  sheet.setFrozenRows(4);
  const widths = [120, 60, 90, 100, 150, 155, 180, 320, 220, 95, 95, 105, 105, 70, 70];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(1, 1, Math.max(5, results.length + 4), 15).setVerticalAlignment('middle');
  if (results.length) sheet.getRange(5, 7, results.length, 3).setWrap(true);
  ss.setActiveSheet(sheet);
  sheet.activate();
}

function resumeCriteresRecherche_(c) {
  const parts = [];
  if (c.text) parts.push('Texte : "' + c.text + '"');
  if (c.account) parts.push('Compte : ' + c.account);
  if (c.name) parts.push('Nom : ' + c.name);
  if (c.year) parts.push('Année : ' + c.year);
  if (c.dateFrom || c.dateTo) parts.push('Dates : ' + (c.dateFrom || '…') + ' → ' + (c.dateTo || '…'));
  if (c.minAmount || c.maxAmount) parts.push('Montant : ' + (c.minAmount || '…') + ' → ' + (c.maxAmount || '…'));
  if (c.direction === 'IN') parts.push('Entrées');
  if (c.direction === 'OUT') parts.push('Sorties');
  parts.push('Onglets : ' + ((c.sheets || []).join(', ') || 'aucun'));
  return parts.join('  |  ');
}

function normaliserRecherche_(value) {
  return String(value === null || value === undefined ? '' : value)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function normaliserCompteRecherche_(value) {
  return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function valeurRecherche_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return Utilities.formatDate(value, CBC.TIMEZONE, 'dd/MM/yyyy');
  return String(value).trim();
}
