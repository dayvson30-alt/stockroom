/**
 * StockRoom — backend (Google Sheets + Apps Script)
 * Planilha "Produtos" = cadastro. Fotos numa pasta do Drive.
 * Acesso protegido por TEAM_CODE (Script Properties) — rode setup() uma vez.
 */
const SHEET_NAME = 'Produtos';
const HEAD = ['codigo', 'nome', 'marca', 'categoria', 'controle', 'unidade', 'quantidade', 'local',
  'fotoUrl', 'obs', 'criadoPor', 'criadoEm', 'atualizadoPor', 'atualizadoEm'];
const PHOTO_FOLDER = 'StockRoom - Fotos';
const EDITABLE = ['nome', 'marca', 'categoria', 'controle', 'unidade', 'quantidade', 'local', 'obs'];

/** Rode uma vez pelo editor: cria a aba, a pasta de fotos e o código da equipe. */
function setup() {
  sheet_();
  folder_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TEAM_CODE')) {
    props.setProperty('TEAM_CODE', String(Math.floor(100000 + Math.random() * 900000)));
  }
  Logger.log('TEAM_CODE = ' + props.getProperty('TEAM_CODE'));
}

function doGet(e) { return handle_((e && e.parameter) || {}); }

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return handle_(body);
}

function handle_(p) {
  let out;
  try {
    const code = PropertiesService.getScriptProperties().getProperty('TEAM_CODE');
    if (code && String(p.key || '') !== code) throw new Error('bad_key');
    const user = String(p.user || '').trim().slice(0, 60);
    switch (p.action) {
      case 'ping': out = {}; break;
      case 'list': out = { items: readAll_() }; break;
      case 'lookup': out = lookup_(String(p.code || '')); break;
      case 'save': out = { item: save_(p.item || {}, p.photo, user) }; break;
      default: throw new Error('bad_action');
    }
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, sh.getMaxRows(), HEAD.length).setNumberFormat('@'); // tudo texto: preserva zeros do código
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}

function rowToObj_(r) {
  const o = {};
  HEAD.forEach(function (h, i) { o[h] = r[i]; });
  return o;
}

function readAll_() {
  const sh = sheet_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEAD.length).getDisplayValues()
    .filter(function (r) { return r[0]; })
    .map(rowToObj_);
}

function save_(item, photo, user) {
  const codigo = String(item.codigo || '').trim().slice(0, 64);
  if (!codigo) throw new Error('no_code');
  if (!String(item.nome || '').trim()) throw new Error('no_name');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const n = sh.getLastRow();
    let row = -1;
    if (n >= 2) {
      const codes = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
      for (let i = 0; i < codes.length; i++) {
        if (codes[i][0] === codigo) { row = i + 2; break; }
      }
    }
    const now = new Date().toISOString();
    const o = row > 0
      ? rowToObj_(sh.getRange(row, 1, 1, HEAD.length).getDisplayValues()[0])
      : { codigo: codigo, criadoPor: user, criadoEm: now };

    EDITABLE.forEach(function (k) {
      if (item[k] !== undefined && item[k] !== null) o[k] = String(item[k]).trim().slice(0, 300);
    });
    if (photo && /^data:image\/(jpeg|png|webp);base64,/.test(photo)) {
      o.fotoUrl = savePhoto_(photo, codigo);
    } else if (!o.fotoUrl && /^https:\/\//.test(String(item.fotoUrl || ''))) {
      o.fotoUrl = String(item.fotoUrl).slice(0, 500); // imagem pública achada na busca
    }
    o.atualizadoPor = user;
    o.atualizadoEm = now;

    if (row < 0) row = Math.max(n, 1) + 1;
    const rg = sh.getRange(row, 1, 1, HEAD.length);
    rg.setNumberFormat('@');
    rg.setValues([HEAD.map(function (h) { return o[h] == null ? '' : String(o[h]); })]);
    return o;
  } finally {
    lock.releaseLock();
  }
}

function savePhoto_(dataUrl, codigo) {
  const m = dataUrl.match(/^data:(image\/\w+);base64,(.+)$/);
  const name = codigo.replace(/[^\w-]/g, '_') + '_' + Date.now() + '.jpg';
  const f = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w600';
}

/** Procura o código em bases públicas (Open Food/Products/Beauty Facts + UPCitemdb). */
function lookup_(code) {
  code = code.replace(/[^\w-]/g, '').slice(0, 64);
  if (!code) throw new Error('no_code');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('lk_' + code);
  if (hit) return { product: JSON.parse(hit) };

  const enc = encodeURIComponent(code);
  const hosts = ['world.openfoodfacts.org', 'world.openproductsfacts.org', 'world.openbeautyfacts.org'];
  const reqs = hosts.map(function (h) {
    return { url: 'https://' + h + '/api/v2/product/' + enc + '.json?fields=product_name,brands,image_front_small_url', muteHttpExceptions: true };
  });
  reqs.push({ url: 'https://api.upcitemdb.com/prod/trial/lookup?upc=' + enc, muteHttpExceptions: true });

  let res = [];
  try { res = UrlFetchApp.fetchAll(reqs); } catch (e) {}

  let product = null;
  for (let i = 0; i < hosts.length && !product; i++) {
    try {
      const j = JSON.parse(res[i].getContentText());
      if (j.status === 1 && j.product && j.product.product_name) {
        product = {
          name: j.product.product_name,
          brand: String(j.product.brands || '').split(',')[0].trim(),
          image: j.product.image_front_small_url || '',
          source: hosts[i].split('.')[1]
        };
      }
    } catch (e) {}
  }
  if (!product && res[hosts.length]) {
    try {
      const j = JSON.parse(res[hosts.length].getContentText());
      if (j.items && j.items.length) {
        const it = j.items[0];
        product = { name: it.title, brand: it.brand || '', image: (it.images || [])[0] || '', source: 'upcitemdb' };
      }
    } catch (e) {}
  }
  cache.put('lk_' + code, JSON.stringify(product), 21600);
  return { product: product };
}
