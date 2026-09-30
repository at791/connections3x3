/**
 * Connections 3x3: receives finished games from the GitHub Pages site.
 * Each game is saved as a JSON file in a Drive folder, and a one-line
 * summary is added to the sheet this script is attached to.
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me    Who has access: Anyone
 * Copy the web app URL into ENDPOINT in index.html.
 *
 * Export: sim/pull_github.ps1 downloads the saved games. It sends the private
 * key made by makeExportKey (kept in Script Properties, never in this code).
 */
const FOLDER_NAME = 'Connections 3x3 games';
const MAX_BYTES = 1000000;                 // ignore anything over ~1 MB
const HEADER = ['received', 'started', 'player', 'puzzle', 'categories', 'outcome',
  'seconds', 'mistakes', 'session', 'slot', 'level', 'level kind', 'puzzle rating',
  'score', 'skill estimate', 'anchor first', 'file'];

function doPost(e) {
  let g;
  try {
    const raw = e && e.postData ? e.postData.contents : '';
    if (!raw || raw.length > MAX_BYTES) return reply_('ignored');
    g = JSON.parse(raw);
    if (g && g.export !== undefined) return exportGames_(g);
  } catch (err) {
    console.error(err);
    return reply_('error: ' + (err && err.message ? err.message : err));
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const raw = e.postData.contents;
    const id = String(g.id || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
    if (!id || !Array.isArray(g.cats) || !g.outcome) return reply_('ignored');

    const folder = folder_();
    const name = id + '.json';
    const existing = folder.getFilesByName(name);
    const file = existing.hasNext() ? existing.next() : null;
    if (file) file.setContent(raw); else folder.createFile(name, raw, MimeType.PLAIN_TEXT);
    const url = (file || folder.getFilesByName(name).next()).getUrl();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADER);
    if (!file) {
      sheet.appendRow([
        new Date(), g.startedAt || '', String(g.player || ''), g.puzzle || '',
        g.cats.join(' / '), g.outcome, Math.round((g.durationMs || 0) / 1000), g.mistakes,
        g.session ? g.session.id : '', g.session ? g.session.n : '',
        g.level ? g.level.n : '', g.level ? (g.level.kind || '') : '',
        g.adaptive ? g.adaptive.difficulty : '', g.adaptive ? g.adaptive.score : '',
        g.adaptive ? g.adaptive.skill : '',
        g.anchor ? g.anchor.anchorFirst : '', url,
      ]);
    }
    return reply_('ok');
  } catch (err) {
    // the reason goes back to the sender (a test POST shows it) and into Executions
    console.error(err);
    return reply_('error: ' + (err && err.message ? err.message : err));
  } finally {
    lock.releaseLock();
  }
}

/* Export for sim/pull_github.ps1. {export: key} lists the saved games (name,
   size, last change); {export: key, names: [...]} returns those files' contents.
   A wrong or missing key gets "denied" and nothing else. */
function exportGames_(g) {
  const key = PropertiesService.getScriptProperties().getProperty('EXPORT_KEY');
  if (!key || typeof g.export !== 'string' || g.export !== key) return reply_('denied');
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  if (!it.hasNext()) return json_({ files: [] });
  const folder = it.next();
  if (Array.isArray(g.names)) {
    const out = {};
    g.names.slice(0, 50).forEach((n) => {
      const f = folder.getFilesByName(String(n));
      if (f.hasNext()) out[n] = f.next().getBlob().getDataAsString('UTF-8');
    });
    return json_({ contents: out });
  }
  const files = [];
  const all = folder.getFiles();
  while (all.hasNext()) {
    const f = all.next();
    if (!/\.json$/.test(f.getName())) continue;
    files.push({ name: f.getName(), size: f.getSize(), updated: f.getLastUpdated().toISOString() });
  }
  return json_({ files });
}

// Setup check: pick testPost in the editor's function menu and click Run. It saves one
// test game the way the site does and prints "ok" or the reason it failed (delete the
// test row and file afterwards). Running it also asks for any permission still missing.
function testPost() {
  const game = { id: 'gTESTeditor01', player: 'TEST-delete-me', puzzle: 0,
    cats: ['TEST ROW', 'DELETE ME', 'FROM THE EDITOR'], outcome: 'test', durationMs: 1000, mistakes: 0 };
  console.log(doPost({ postData: { contents: JSON.stringify(game) } }).getContent());
}

// Run once from the editor: makes the private export key, stores it in Script
// Properties and prints it. The key goes in logs_github/export_key.txt on your computer,
// outside the github folder so it never gets uploaded. Running it again replaces the key.
function makeExportKey() {
  const key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('EXPORT_KEY', key);
  console.log('Export key: ' + key);
}

// Visiting the URL in a browser shows this, which confirms the deployment works.
function doGet() { return reply_('Connections 3x3 collector is running.'); }

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}
function reply_(text) { return ContentService.createTextOutput(text); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// Run this once from the editor if Google never asked for Drive access (or it was unticked):
// it shows the permission screen for everything the script needs.
function authorize() { ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL); }
