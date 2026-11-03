/**
 * Мувофиқасозии дархостҳо — серверная часть на Google Apps Script.
 * Таблица Google Sheets выступает базой; этот скрипт — единственный способ
 * читать и писать её из статического сайта на GitHub Pages.
 *
 * Установка описана в README.md репозитория isosaid/avesto-soglasovanie.
 *
 * Разработано Отделом автоматизации холдинга ООО «Авесто Групп» · ISJ
 */

var SHEETS = {
  contracts: 'Шартномаҳо',
  payments:  'Пардохтҳо',
  requests:  'Дархостҳо',
  items:     'Сатрҳо',
  approvals: 'Имзоҳо',
  signers:   'Имзокунандагон'
};

var COLS = {
  contracts: ['id','no','date','supplier','country','object','product','currency','rate','bank','invoice','gtd','file'],
  payments:  ['id','contractId','amount','date','rate','tjs','note','requestId'],
  requests:  ['id','number','date','company','contractId','currency','rate','purpose','purposeShort','submittedAt','paidAt','createdAt'],
  items:     ['id','requestId','name','unit','qty','price'],
  approvals: ['id','requestId','roleKey','decision','at','byName','comment'],
  signers:   ['key','posTj','posRu','name','pin','order']
};

/* ─────────────────────────── вход ─────────────────────────── */

function doGet(e) {
  try {
    return json({ok: true, data: bootstrap()});
  } catch (err) {
    return json({ok: false, error: String(err && err.message || err)});
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json({ok: false, error: 'Система банд аст, такрор кунед · система занята, повторите'});
  }
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var action = body.action;
    var p = body.payload || {};
    var out;
    switch (action) {
      case 'saveContract':  out = saveContract(p); break;
      case 'saveRequest':   out = saveRequest(p); break;
      case 'submitRequest': out = submitRequest(p); break;
      case 'decide':        out = decide(p); break;
      case 'markPaid':      out = markPaid(p); break;
      case 'deleteRequest': out = deleteRequest(p); break;
      case 'importRegistry': out = importRegistry(p); break;
      default: throw new Error('Амали номаълум · неизвестное действие: ' + action);
    }
    return json({ok: true, data: out, state: bootstrap()});
  } catch (err) {
    return json({ok: false, error: String(err && err.message || err)});
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ─────────────────────── доступ к листам ─────────────────────── */

function book() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet(key) {
  var sh = book().getSheetByName(SHEETS[key]);
  if (!sh) {
    sh = book().insertSheet(SHEETS[key]);
    sh.getRange(1, 1, 1, COLS[key].length).setValues([COLS[key]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll(key) {
  var sh = sheet(key);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var cols = COLS[key];
  var values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (String(row[0]).trim() === '') continue;
    var o = {};
    for (var c = 0; c < cols.length; c++) o[cols[c]] = norm(row[c]);
    out.push(o);
  }
  return out;
}

function norm(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}

function rowIndexById(key, id) {
  var sh = sheet(key);
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function toRow(key, obj) {
  return COLS[key].map(function (c) { return obj[c] === undefined || obj[c] === null ? '' : obj[c]; });
}

function upsert(key, obj) {
  var sh = sheet(key);
  var r = rowIndexById(key, obj.id);
  var row = toRow(key, obj);
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return obj;
}

function removeWhere(key, field, value) {
  var sh = sheet(key);
  var last = sh.getLastRow();
  if (last < 2) return;
  var cols = COLS[key];
  var idx = cols.indexOf(field);
  var values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][idx]) === String(value)) sh.deleteRow(i + 2);
  }
}

function uid(prefix) {
  return prefix + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
}

/* ─────────────────────────── чтение ─────────────────────────── */

function bootstrap() {
  var contracts = readAll('contracts');
  var payments  = readAll('payments');
  var requests  = readAll('requests');
  var items     = readAll('items');
  var approvals = readAll('approvals');
  var signers   = readAll('signers');

  contracts.forEach(function (c) {
    c.rate = Number(c.rate) || 0;
    c.invoice = Number(c.invoice) || 0;
    c.payments = payments.filter(function (p) { return String(p.contractId) === String(c.id); })
      .map(function (p) {
        return {id: p.id, amount: Number(p.amount) || 0, date: p.date, rate: Number(p.rate) || 0,
                tjs: Number(p.tjs) || 0, note: p.note, requestId: p.requestId};
      });
  });

  requests.forEach(function (r) {
    r.rate = Number(r.rate) || 0;
    r.items = items.filter(function (i) { return String(i.requestId) === String(r.id); })
      .map(function (i) {
        return {id: i.id, name: i.name, unit: i.unit, qty: Number(i.qty) || 0, price: Number(i.price) || 0};
      });
    r.approvals = {};
    approvals.filter(function (a) { return String(a.requestId) === String(r.id); })
      .forEach(function (a) {
        r.approvals[a.roleKey] = {decision: a.decision, at: a.at, byName: a.byName, comment: a.comment};
      });
  });

  // PIN наружу не отдаём — только признак, что он задан
  var publicSigners = signers
    .sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); })
    .map(function (s) {
      return {key: s.key, posTj: s.posTj, posRu: s.posRu, name: s.name,
              hasPin: String(s.pin || '').trim() !== '', order: Number(s.order) || 0};
    });

  return {contracts: contracts, requests: requests, signers: publicSigners, serverTime: new Date().toISOString()};
}

/* ─────────────────────── проверка подписанта ─────────────────────── */

function checkPin(roleKey, pin) {
  var signers = readAll('signers');
  var s = null;
  for (var i = 0; i < signers.length; i++) if (String(signers[i].key) === String(roleKey)) s = signers[i];
  if (!s) throw new Error('Имзокунанда ёфт нашуд · подписант не найден: ' + roleKey);
  var want = String(s.pin || '').trim();
  if (want === '') throw new Error('Барои ин вазифа PIN таъин нашудааст · для этой должности не задан PIN');
  if (String(pin || '').trim() !== want) throw new Error('PIN нодуруст · неверный PIN');
  return s;
}

/* ─────────────────────────── запись ─────────────────────────── */

function saveContract(p) {
  var c = p.contract || {};
  if (!c.no) throw new Error('Рақами шартнома холӣ · не указан номер контракта');
  if (!c.id) c.id = uid('c');
  upsert('contracts', {
    id: c.id, no: c.no, date: c.date || '', supplier: c.supplier || '', country: c.country || '',
    object: c.object || '', product: c.product || '', currency: c.currency || '', rate: Number(c.rate) || 0,
    bank: c.bank || '', invoice: Number(c.invoice) || 0, gtd: c.gtd || '', file: c.file || ''
  });
  return {id: c.id};
}

function saveRequest(p) {
  var r = p.request || {};
  if (!r.number) throw new Error('Рақами дархост холӣ · не указан номер заявки');
  var isNew = !r.id;
  if (isNew) r.id = uid('r');

  var existing = isNew ? null : findRequest(r.id);
  if (existing && existing.submittedAt && !p.force) {
    throw new Error('Дархост аллакай ба мувофиқа фиристода шуд · заявка уже на согласовании, править нельзя');
  }

  upsert('requests', {
    id: r.id, number: r.number, date: r.date || '', company: r.company || '',
    contractId: r.contractId || '', currency: r.currency || '', rate: Number(r.rate) || 0,
    purpose: r.purpose || '', purposeShort: r.purposeShort || '',
    submittedAt: existing ? existing.submittedAt : '', paidAt: existing ? existing.paidAt : '',
    createdAt: existing ? existing.createdAt : new Date().toISOString()
  });

  removeWhere('items', 'requestId', r.id);
  (r.items || []).forEach(function (it) {
    upsert('items', {id: uid('i'), requestId: r.id, name: it.name || '', unit: it.unit || '',
                     qty: Number(it.qty) || 0, price: Number(it.price) || 0});
  });
  return {id: r.id};
}

function findRequest(id) {
  var all = readAll('requests');
  for (var i = 0; i < all.length; i++) if (String(all[i].id) === String(id)) return all[i];
  return null;
}

function submitRequest(p) {
  var r = findRequest(p.requestId);
  if (!r) throw new Error('Дархост ёфт нашуд · заявка не найдена');
  if (r.submittedAt) throw new Error('Аллакай фиристода шуд · уже отправлена');
  var items = readAll('items').filter(function (i) { return String(i.requestId) === String(r.id); });
  if (!items.length) throw new Error('Маблағ холӣ · в заявке нет позиций');
  r.submittedAt = new Date().toISOString();
  upsert('requests', r);
  return {submittedAt: r.submittedAt};
}

function decide(p) {
  var s = checkPin(p.roleKey, p.pin);
  var r = findRequest(p.requestId);
  if (!r) throw new Error('Дархост ёфт нашуд · заявка не найдена');
  if (!r.submittedAt) throw new Error('Дархост ҳанӯз фиристода нашуд · заявка ещё не отправлена на согласование');
  if (r.paidAt) throw new Error('Дархост пардохт шуд · заявка уже оплачена');

  var signers = readAll('signers').sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
  var approvals = readAll('approvals').filter(function (a) { return String(a.requestId) === String(r.id); });
  var byRole = {};
  approvals.forEach(function (a) { byRole[a.roleKey] = a; });

  for (var i = 0; i < approvals.length; i++) {
    if (approvals[i].decision === 'rejected') throw new Error('Дархост рад шудааст · заявка отклонена, визы больше не ставятся');
  }

  // Порядок виз свободный; утверждение холдинга (последний по order) — только после остальных.
  if (!signers.length) throw new Error('Справочник подписантов пуст');
  var finalKey = signers[signers.length - 1].key;

  var mine = byRole[p.roleKey];
  if (mine && mine.decision === 'approved') {
    throw new Error('Шумо аллакай имзо гузоштаед · вы уже подписали эту заявку');
  }

  if (String(p.roleKey) === String(finalKey)) {
    var missing = [];
    for (var k = 0; k < signers.length; k++) {
      var sk = signers[k].key;
      if (String(sk) === String(finalKey)) continue;
      var vv = byRole[sk];
      if (!vv || vv.decision !== 'approved') missing.push(signers[k].name);
    }
    if (missing.length) {
      throw new Error('Аввал дигарон имзо гузоранд · сначала должны подписать: ' + missing.join(', '));
    }
  }
  if (p.decision === 'rejected' && !String(p.comment || '').trim()) {
    throw new Error('Сабаби радро нависед · укажите причину отказа');
  }

  var existing = byRole[p.roleKey];
  upsert('approvals', {
    id: existing ? existing.id : uid('a'),
    requestId: r.id, roleKey: p.roleKey,
    decision: p.decision === 'rejected' ? 'rejected' : 'approved',
    at: new Date().toISOString(), byName: s.name || '', comment: String(p.comment || '').trim()
  });
  return {roleKey: p.roleKey, decision: p.decision};
}

function markPaid(p) {
  checkPin(p.roleKey, p.pin);
  var r = findRequest(p.requestId);
  if (!r) throw new Error('Дархост ёфт нашуд · заявка не найдена');
  if (r.paidAt) throw new Error('Пардохт аллакай қайд шуд · оплата уже отмечена');

  var signers = readAll('signers');
  var approvals = readAll('approvals').filter(function (a) { return String(a.requestId) === String(r.id); });
  var okCount = approvals.filter(function (a) { return a.decision === 'approved'; }).length;
  if (okCount < signers.length) throw new Error('Масир нопурра · маршрут не пройден полностью');

  if (!r.contractId) throw new Error('Шартнома пайваст нашуд · контракт не привязан');
  var amount = Number(p.amount) || 0;
  if (amount <= 0) throw new Error('Маблағ нодуруст · неверная сумма');
  var rate = Number(p.rate) || 0;

  upsert('payments', {
    id: uid('p'), contractId: r.contractId, amount: amount, date: p.date || '',
    rate: rate, tjs: amount * rate, note: p.note || '', requestId: r.id
  });
  r.paidAt = new Date().toISOString();
  upsert('requests', r);
  return {paidAt: r.paidAt};
}

function deleteRequest(p) {
  checkPin(p.roleKey, p.pin);
  var r = findRequest(p.requestId);
  if (!r) throw new Error('Дархост ёфт нашуд');
  if (r.submittedAt) throw new Error('Фиристодашударо нест кардан мумкин нест · отправленную заявку удалить нельзя');
  removeWhere('items', 'requestId', r.id);
  removeWhere('approvals', 'requestId', r.id);
  var sh = sheet('requests');
  var row = rowIndexById('requests', r.id);
  if (row > 0) sh.deleteRow(row);
  return {deleted: r.id};
}

/**
 * Импорт реестра из Excel: сайт разбирает файл в браузере и присылает
 * готовые контракты с платежами. Листы «Шартномаҳо» и «Пардохтҳо»
 * перезаписываются целиком; заявки, визы и подписанты не трогаются.
 */
function importRegistry(p) {
  var list = p.contracts || [];
  if (!list.length) throw new Error('Рӯйхат холӣ · список контрактов пуст');

  var shC = sheet('contracts'), shP = sheet('payments');
  if (shC.getLastRow() > 1) shC.getRange(2, 1, shC.getLastRow() - 1, COLS.contracts.length).clearContent();
  if (shP.getLastRow() > 1) shP.getRange(2, 1, shP.getLastRow() - 1, COLS.payments.length).clearContent();

  var crows = [], prows = [], pn = 0;
  list.forEach(function (c) {
    var cid = c.id || uid('c');
    crows.push(toRow('contracts', {
      id: cid, no: c.no || '', date: c.date || '', supplier: c.supplier || '',
      country: c.country || '', object: c.object || '', product: c.product || '',
      currency: c.currency || '', rate: Number(c.rate) || 0, bank: c.bank || '',
      invoice: Number(c.invoice) || 0, gtd: c.gtd || '', file: c.file || ''
    }));
    (c.payments || []).forEach(function (q) {
      pn++;
      prows.push(toRow('payments', {
        id: 'p' + pn + '-' + cid, contractId: cid, amount: Number(q.amount) || 0,
        date: q.date || '', rate: Number(q.rate) || 0, tjs: Number(q.tjs) || 0,
        note: q.note || '', requestId: q.requestId || ''
      }));
    });
  });
  if (crows.length) shC.getRange(2, 1, crows.length, COLS.contracts.length).setValues(crows);
  if (prows.length) shP.getRange(2, 1, prows.length, COLS.payments.length).setValues(prows);
  return {contracts: crows.length, payments: prows.length};
}

/* ───────────────── первичная настройка таблицы ─────────────────
 * Запустите setup() один раз из редактора скриптов: он создаст листы,
 * заголовки и справочник подписантов с маршрутом согласования.
 * PIN-коды после этого впишите вручную в лист «Имзокунандагон».
 */
function setup() {
  Object.keys(SHEETS).forEach(function (k) {
    var sh = sheet(k);
    sh.getRange(1, 1, 1, COLS[k].length).setValues([COLS[k]]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS[k].length).setFontWeight('bold');
  });

  if (readAll('signers').length === 0) {
    // Порядок — как визируется бумажный бланк, снизу вверх
    var route = [
      ['branch','Муовини директори генералӣ оид ба самти паррандапарварӣ ҶДММ «Мармарӣ»','Зам. ген. директора по птицеводству','Ҳакимов З.','',1],
      ['finance','Муовини директори генералӣ оиди молияи ҶДММ «Мармарӣ»','Зам. ген. директора по финансам','Ибодов Д.М.','',2],
      ['chiefacc','Сармуҳосиби ҶДММ «Мармарӣ»','Главный бухгалтер','Зоиров С.А.','',3],
      ['initiator','Директори генералии ҶДММ «Мармарӣ»','Ген. директор предприятия-заявителя','Турсунов А.Г.','',4],
      ['holding','Директори генералӣ ҶДММ «Авесто Гуруҳ»','Ген. директор холдинга — утверждение','Муҳамад Мирзо','',5]
    ];
    sheet('signers').getRange(2, 1, route.length, COLS.signers.length).setValues(route);
  }
  SpreadsheetApp.getActive().toast('Листҳо тайёр · структура создана. Впишите PIN в лист «Имзокунандагон».');
}
