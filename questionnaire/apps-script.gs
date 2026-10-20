/**
 * Questionnaire Deezer Étudiant – enregistrement des réponses dans Google Sheets.
 *
 * Installation (ou mise à jour) :
 * 1. Dans la feuille Google Sheets : Extensions > Apps Script, remplacer tout le code par celui-ci, enregistrer.
 * 2. Déployer > Gérer les déploiements > crayon (modifier) > Version : « Nouvelle version » > Déployer.
 *    L'URL en /exec ne change pas : rien à modifier dans index.html.
 *    (Premier déploiement : Déployer > Nouveau déploiement > Application Web,
 *     exécuter en tant que « Moi », accès « Tout le monde », puis coller l'URL /exec dans index.html.)
 *
 * Les réponses arrivent dans l'onglet ci-dessous (créé automatiquement s'il n'existe pas).
 * La ligne d'en-tête se construit toute seule : une nouvelle question = une nouvelle colonne,
 * sans jamais décaler les colonnes existantes.
 */
const ONGLET = "reponses_etudiant";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    data.horodatage = new Date();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ONGLET) || ss.insertSheet(ONGLET);

    let headers = sh.getLastColumn() > 0
      ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String)
      : [];
    if (headers.length === 0) headers = ["horodatage"];
    const nouvelles = Object.keys(data).filter(k => headers.indexOf(k) === -1);
    if (nouvelles.length) {
      headers = headers.concat(nouvelles);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }

    const row = headers.map(h => (h in data && data[h] !== null) ? data[h] : "");
    sh.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
