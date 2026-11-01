/**
 * StockRoom — backend (Google Sheets + Apps Script)
 * Planilha "Produtos" = cadastro. Fotos numa pasta do Drive.
 * Acesso protegido por TEAM_CODE (Script Properties) — rode setup() uma vez.
 */
const SHEET_NAME = 'Produtos';
const HEAD = ['codigo', 'nome', 'marca', 'categoria', 'controle', 'unidade', 'quantidade', 'local',
  'fotoUrl', 'obs', 'criadoPor', 'criadoEm', 'atualizadoPor', 'atualizadoEm', 'minimo'];
const PHOTO_FOLDER = 'StockRoom - Fotos';
const EDITABLE = ['nome', 'marca', 'categoria', 'controle', 'unidade', 'quantidade', 'local', 'obs', 'minimo'];
const OUT_SHEET = 'Saidas';
const OUT_HEAD = ['retirada', 'data', 'pessoa', 'casa', 'codigo', 'nome', 'quantidade', 'controle'];
const MAX_LINES = 100, MAX_QTY = 999;
const HOUSES_URL = 'https://gist.githubusercontent.com/dayvson30-alt/%ID%/raw/casas.json'; // mesma lista do Deck Report
const HOUSES_CACHE_S = 21600;
const MAX_OFFLINE_MS = 30 * 24 * 3600 * 1000, MAX_CLOCK_SKEW_MS = 3600 * 1000;
const ORDER_SHEET = 'Compras';
const ORDER_HEAD = ['pedido', 'data', 'solicitante', 'codigo', 'nome', 'quantidade', 'saldo', 'status'];
const MANAGER_ACTIONS = ['panel', 'order'];
const SHEET_ONLY_FIELDS = ['criadoPor', 'atualizadoPor']; // quem cadastrou fica só na planilha
const FALLBACK_MIN = 5;
const PANEL_URL = 'https://dayvson30-alt.github.io/stockroom/painel.html';

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
    const isManager = MANAGER_ACTIONS.indexOf(p.action) >= 0; // o painel tem código próprio
    const code = PropertiesService.getScriptProperties().getProperty(isManager ? 'MANAGER_CODE' : 'TEAM_CODE');
    if (isManager && !code) throw new Error('bad_key');
    if (code && String(p.key || '') !== code) throw new Error('bad_key');
    const user = String(p.user || '').trim().slice(0, 60);
    switch (p.action) {
      case 'ping': out = {}; break;
      case 'list': out = { items: readAll_().map(forApp_) }; break;
      case 'lookup': out = lookup_(String(p.code || '')); break;
      case 'save': out = { item: forApp_(save_(p.item || {}, p.photo, user)) }; break;
      case 'checkout': out = { items: checkout_(p, user) }; break;
      case 'houses': out = { houses: houses_() }; break;
      case 'panel': out = panel_(); break;
      case 'order': order_(p, user); out = {}; break;
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
  } else if (sh.getLastColumn() < HEAD.length) { // planilha criada antes de existir a última coluna
    sh.getRange(1, HEAD.length, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
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

function forApp_(item) {
  const o = {};
  Object.keys(item).forEach(function (k) { if (SHEET_ONLY_FIELDS.indexOf(k) < 0) o[k] = item[k]; });
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

function logSheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function outSheet_() { return logSheet_(OUT_SHEET, OUT_HEAD); }
function orderSheet_() { return logSheet_(ORDER_SHEET, ORDER_HEAD); }

function readLog_(sh, head) {
  const m = sh.getLastRow();
  if (m < 2) return [];
  return sh.getRange(2, 1, m - 1, head.length).getDisplayValues().map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

function defaultMin_() {
  const n = Number(PropertiesService.getScriptProperties().getProperty('DEFAULT_MIN'));
  return Number.isInteger(n) && n > 0 ? n : FALLBACK_MIN;
}

function panel_() {
  return {
    items: readAll_().map(forApp_),
    saidas: readLog_(outSheet_(), OUT_HEAD),
    pedidos: readLog_(orderSheet_(), ORDER_HEAD),
    minimo: defaultMin_(),
    aprovador: PropertiesService.getScriptProperties().getProperty('APPROVER_PHONE') || ''
  };
}

/** Registra o pedido enviado ao aprovador; o id torna o envio repetível. */
function order_(p, user) {
  const id = String(p.id || '');
  if (!/^[\w-]{6,40}$/.test(id)) throw new Error('bad_id');
  if (!user) throw new Error('no_user');
  const lines = checkLines_(p.lines);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const items = readAll_();
    const log = lines.map(function (l) {
      const it = items.filter(function (i) { return i.codigo === l.codigo; })[0];
      if (!it) throw new Error('unknown_code');
      return [id, new Date().toISOString(), user, l.codigo, it.nome, String(l.quantidade), it.quantidade, 'solicitado'];
    });
    const sh = orderSheet_();
    const m = sh.getLastRow();
    const seen = m >= 2 && sh.getRange(2, 1, m - 1, 1).getDisplayValues().some(function (r) { return r[0] === id; });
    if (seen) return;
    const rg = sh.getRange(m + 1, 1, log.length, ORDER_HEAD.length);
    rg.setNumberFormat('@');
    rg.setValues(log);
  } finally {
    lock.releaseLock();
  }
}

function lowStock_(items, fallbackMin) {
  return items.filter(function (i) {
    if (!/^-?\d+$/.test(String(i.quantidade).trim())) return false; // sem contagem não dá alerta
    const min = /^\d+$/.test(String(i.minimo).trim()) ? Number(i.minimo) : fallbackMin;
    return Number(i.quantidade) <= min;
  });
}

/** Rodado por gatilho diário: só envia e-mail se houver algo para comprar. */
function alertaDiario() {
  const to = PropertiesService.getScriptProperties().getProperty('ALERT_EMAIL');
  if (!to) return;
  const low = lowStock_(readAll_(), defaultMin_());
  if (!low.length) return;
  const body = low.map(function (i) { return '- ' + i.nome + ': restam ' + i.quantidade; }).join('\n');
  MailApp.sendEmail(to, 'FF Stock: ' + low.length + ' produto(s) para comprar',
    'Estoque baixo, recomendamos comprar:\n\n' + body + '\n\nMonte o pedido no painel: ' + PANEL_URL);
}

function createDailyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'alertaDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('alertaDiario').timeBased().everyDays(1).atHour(7).create();
}

// Retirada feita sem sinal chega depois: vale a hora do celular, se for plausível
function whenOf_(value, now) {
  const t = new Date(String(value || '')).getTime();
  if (isNaN(t) || t > now + MAX_CLOCK_SKEW_MS || t < now - MAX_OFFLINE_MS) return new Date(now).toISOString();
  return new Date(t).toISOString();
}

function checkLines_(lines) {
  if (!Array.isArray(lines) || !lines.length || lines.length > MAX_LINES) throw new Error('bad_lines');
  return lines.map(function (l) {
    const qtd = Number(l && l.quantidade);
    if (!Number.isInteger(qtd) || qtd < 1 || qtd > MAX_QTY) throw new Error('bad_qty');
    return { codigo: String(l.codigo || '').trim(), quantidade: qtd };
  });
}

/**
 * Registra uma retirada: uma linha em Saidas por produto e baixa na quantidade.
 * Só desconta de quem já tem contagem; o saldo pode ficar negativo (sinal de recontar).
 * O id torna o envio repetível: se a resposta se perdeu, reenviar não duplica.
 */
function checkout_(p, user) {
  const id = String(p.id || '');
  if (!/^[\w-]{6,40}$/.test(id)) throw new Error('bad_id');
  if (!user) throw new Error('no_user');
  const lines = checkLines_(p.lines);
  const casa = String(p.casa || '').trim().slice(0, 100);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const n = sh.getLastRow();
    const rows = n >= 2 ? sh.getRange(2, 1, n - 1, HEAD.length).getDisplayValues() : [];
    const qCol = HEAD.indexOf('quantidade');
    const rowOf = {};
    rows.forEach(function (r, i) { if (r[0]) rowOf[r[0]] = i; });
    lines.forEach(function (l) { if (rowOf[l.codigo] === undefined) throw new Error('unknown_code'); });
    const result = function () {
      return lines.map(function (l) { return { codigo: l.codigo, quantidade: rows[rowOf[l.codigo]][qCol] }; });
    };

    const out = outSheet_();
    const m = out.getLastRow();
    const seen = m >= 2 && out.getRange(2, 1, m - 1, 1).getDisplayValues().some(function (r) { return r[0] === id; });
    if (seen) return result();

    const now = whenOf_(p.quando, Date.now());
    const log = lines.map(function (l) {
      const r = rows[rowOf[l.codigo]];
      return [id, now, user, casa, l.codigo, r[HEAD.indexOf('nome')], String(l.quantidade), r[HEAD.indexOf('controle')]];
    });
    const rg = out.getRange(m + 1, 1, log.length, OUT_HEAD.length);
    rg.setNumberFormat('@');
    rg.setValues(log);

    lines.forEach(function (l) {
      const r = rows[rowOf[l.codigo]];
      if (!/^-?\d+$/.test(String(r[qCol]).trim())) return; // sem contagem: só registra
      r[qCol] = String(Number(r[qCol]) - l.quantidade);
      sh.getRange(rowOf[l.codigo] + 2, qCol + 1).setValue(r[qCol]);
    });
    return result();
  } finally {
    lock.releaseLock();
  }
}

/** Lista de casas (l = endereço, c = condomínio). O id do gist fica em Script Properties (HOUSES_GIST). */
function houses_() {
  const id = PropertiesService.getScriptProperties().getProperty('HOUSES_GIST');
  if (!id) return [];
  const cache = CacheService.getScriptCache();
  const hit = cache.get('houses');
  if (hit) return JSON.parse(hit);

  const res = UrlFetchApp.fetch(HOUSES_URL.replace('%ID%', encodeURIComponent(id)), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return [];
  const list = (JSON.parse(res.getContentText()).addresses || [])
    .map(function (a) { return { l: String(a.label || ''), c: String(a.comm || '') }; })
    .filter(function (a) { return a.l; });
  cache.put('houses', JSON.stringify(list), HOUSES_CACHE_S);
  return list;
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
