/**
 * Бюджети сармоягузорӣ — серверная часть на Google Apps Script.
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
  signers:   'Имзокунандагон',
  budget:    'Буҷет',
  log:       'Журнал',
  codes:     'Моддаҳо',
  cehs:      'Сехҳо',
  protocols: 'Протоколҳо',
  docs:      'Ҳуҷҷатҳо'
};

var COLS = {
  contracts: ['id','no','date','supplier','country','object','product','currency','rate','bank','invoice','gtd','file',
              'reqAmount','reqTjs','allocAmount','allocTjs','req2Rate','req2Amount','req2Tjs',
              'alloc2Rate','alloc2Amount','alloc2Tjs','allocTotal',
              'plan1Cond','plan1Sum','plan2Cond','plan2Sum','plan3Cond','plan3Sum','code','protocolId','ceh'],
  payments:  ['id','contractId','amount','date','rate','tjs','note','requestId'],
  requests:  ['id','number','date','company','contractId','currency','rate','purpose','purposeShort','submittedAt','paidAt','createdAt','code','tranche','ceh'],
  items:     ['id','requestId','name','unit','qty','price'],
  approvals: ['id','requestId','roleKey','decision','at','byName','comment'],
  signers:   ['key','posTj','posRu','name','login','password','order','role'],
  // план инвестиционного бюджета: строка на организацию, год и код расхода
  budget:    ['id','company','year','code','m1','m2','m3','m4','m5','m6','m7','m8','m9','m10','m11','m12','ceh'],
  // справочник цехов (подразделений)
  cehs:      ['name','order'],
  // справочник статей бюджета: код, название, родительская группа
  codes:     ['code','name','ru','parent','order'],
  // журнал действий: кто, когда и что сделал
  log:       ['id','at','byKey','byName','action','object','details'],
  // протоколы заседания Совета директоров
  protocols: ['id','number','date','company','subject','decisions','fileUrl','fileName','createdAt'],
  // подтверждающие документы заявки (файлы лежат на Диске)
  docs:      ['id','requestId','kind','name','url','size','addedAt','byName']
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
    var who = whoIs(p);                    // кто действует — до действия: после выхода токена уже нет
    var subj = describeAction(action, p);  // что за объект — до действия: после удаления его не прочитать
    switch (action) {
      case 'saveContract':  out = saveContract(p); break;
      case 'saveRequest':   out = saveRequest(p); break;
      case 'submitRequest': out = submitRequest(p); break;
      case 'decide':        out = decide(p); break;
      case 'markPaid':      out = markPaid(p); break;
      case 'deleteRequest': out = deleteRequest(p); break;
      case 'deleteContract': out = deleteContract(p); break;
      case 'importRegistry': out = importRegistry(p); break;
      case 'login':         out = login(p); break;
      case 'logout':        out = logout(p); break;
      case 'changePassword': out = changePassword(p); break;
      case 'saveBudget':    out = saveBudget(p); break;
      case 'saveCodes':     out = saveCodes(p); break;
      case 'saveCehs':      out = saveCehs(p); break;
      case 'saveProtocol':  out = saveProtocol(p); break;
      case 'deleteProtocol':out = deleteProtocol(p); break;
      case 'uploadDoc':     out = uploadDoc(p); break;
      case 'deleteDoc':     out = deleteDoc(p); break;
      case 'getLog':
        if (!who.key || !isAdmin(findSigner('key', who.key)))
          throw new Error('Журнал танҳо барои маъмур · журнал доступен только администратору');
        out = {};
        break;
      default: throw new Error('Амали номаълум · неизвестное действие: ' + action);
    }
    if (action === 'login') who = {key: out.role.key, name: out.role.name};
    if (action !== 'getLog') writeLog(who, action, subj);
    var admin = action !== 'logout' && !!who.key && isAdmin(findSigner('key', who.key));
    return json({ok: true, data: out, state: bootstrap(admin)});
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

/**
 * Даты отдаём строкой «гггг-мм-дд» в часовом поясе таблицы.
 * Через getScriptTimeZone дата уезжала на день: пояс скрипта и пояс
 * таблицы разные, и 10 июня превращалось в 9 июня 19:00 UTC.
 */
function norm(v) {
  if (v && typeof v.getTime === 'function' && !isNaN(v.getTime())) {
    var tz = book().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
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
  // Опознаём строку по первой колонке листа: у подписантов это «key», у остальных «id».
  var r = rowIndexById(key, obj[COLS[key][0]]);
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

function bootstrap(withLog) {
  var contracts = readAll('contracts');
  var payments  = readAll('payments');
  var requests  = readAll('requests');
  var items     = readAll('items');
  var approvals = readAll('approvals');
  var signers   = readAll('signers');
  var budget    = readAll('budget');
  var protocols = readAll('protocols');
  var docs      = readAll('docs');
  var cehs      = readAll('cehs').sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
  var codes     = readAll('codes').sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
  // журнал видит только администратор: в общий ответ он не попадает
  var log = null;
  if (withLog) {
    log = readAll('log');
    log = log.slice(Math.max(0, log.length - 300)).reverse();
  }

  contracts.forEach(function (c) {
    c.rate = Number(c.rate) || 0;
    c.invoice = Number(c.invoice) || 0;
    ['reqAmount','reqTjs','allocAmount','allocTjs','req2Rate','req2Amount','req2Tjs',
     'alloc2Rate','alloc2Amount','alloc2Tjs','allocTotal'].forEach(function (k) {
      c[k] = Number(c[k]) || 0;
    });
    c.plan = [1, 2, 3].map(function (i) {
      return {cond: c['plan' + i + 'Cond'] || '', amount: Number(c['plan' + i + 'Sum']) || 0};
    }).filter(function (x) { return x.cond || x.amount; });
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

  // Пароль наружу не отдаём — только признак, что он задан.
  // Повторы по key отбрасываем: один подписант — один шаг маршрута.
  var seen = {};
  var publicSigners = signers
    .filter(function (s) {
      var k = String(s.key || '').trim();
      if (isAdmin(s)) return false;          // администратор в маршрут не входит
      if (!k || seen[k]) return false;
      seen[k] = true;
      return true;
    })
    .sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); })
    .map(function (s) {
      return {key: s.key, posTj: s.posTj, posRu: s.posRu, name: s.name,
              login: s.login || '', hasPassword: String(s.password || '').trim() !== '',
              order: Number(s.order) || 0};
    });

  budget.forEach(function (b) {
    b.year = Number(b.year) || 0;
    for (var m = 1; m <= 12; m++) b['m' + m] = Number(b['m' + m]) || 0;
  });

  // подтверждающие документы кладём внутрь своей заявки
  var byReq = {};
  docs.forEach(function (d) {
    (byReq[d.requestId] = byReq[d.requestId] || []).push(d);
  });
  requests.forEach(function (r) { r.docs = byReq[r.id] || []; });

  return {contracts: contracts, requests: requests, signers: publicSigners,
          budget: budget, protocols: protocols, log: log || undefined, codes: codes, cehs: cehs,
          serverTime: new Date().toISOString()};
}

/**
 * План инвестиционного бюджета: по строке на организацию, год и код расхода.
 * Приходит целиком за год — переписываем только строки этой организации и года.
 */
function saveBudget(p) {
  auth(p);
  var company = String(p.company || '').trim();
  var year = Number(p.year) || 0;
  if (!company) throw new Error('Ташкилот интихоб нашуд · не выбрана организация');
  if (!year) throw new Error('Сол нишон дода нашуд · не указан год');

  // план хранится отдельно по цехам; пустой цех — общий план организации
  var ceh = String(p.ceh || '').trim();
  var rows = readAll('budget');
  var keep = rows.filter(function (b) {
    return !(String(b.company) === company && Number(b.year) === year && String(b.ceh || '') === ceh);
  });

  var fresh = (p.rows || []).map(function (b) {
    var o = {id: b.id || uid('b'), company: company, year: year, code: String(b.code || ''), ceh: ceh};
    for (var m = 1; m <= 12; m++) o['m' + m] = Number(b['m' + m]) || 0;
    return o;
  });

  var all = keep.concat(fresh);
  var sh = sheet('budget');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, COLS.budget.length).clearContent();
  if (all.length) {
    sh.getRange(2, 1, all.length, COLS.budget.length)
      .setValues(all.map(function (b) { return toRow('budget', b); }));
  }
  return {saved: fresh.length};
}

/* ───────────────────── журнал действий ─────────────────────
 * Каждое успешное действие — строка: когда, кто, что и над чем.
 * Пишется из doPost, поэтому мимо журнала не проходит ничего.
 */
function whoIs(p) {
  try {
    var s = auth(p);
    return {key: s.key, name: s.name};
  } catch (e) {
    return {key: '', name: ''};
  }
}

function describeAction(action, p) {
  try {
    var r, c, x;
    switch (action) {
      case 'saveContract':
        c = p.contract || {};
        return {object: 'Шартнома № ' + (c.no || ''), details: c.id ? 'таҳрир · правка' : 'нав · новый'};
      case 'deleteContract':
        c = readAll('contracts').filter(function (k) { return String(k.id) === String(p.contractId); })[0] || {};
        return {object: 'Шартнома № ' + (c.no || p.contractId), details: c.supplier || ''};
      case 'saveRequest':
        r = p.request || {};
        return {object: 'Дархост № ' + (r.number || ''),
                details: (r.id && findRequest(r.id) ? 'таҳрир · правка' : 'нав · новая') + (p.force ? ' · визы сняты' : '')};
      case 'submitRequest':
      case 'deleteRequest':
        r = findRequest(p.requestId) || {};
        return {object: 'Дархост № ' + (r.number || p.requestId), details: ''};
      case 'decide':
        r = findRequest(p.requestId) || {};
        return {object: 'Дархост № ' + (r.number || p.requestId),
                details: (p.decision === 'rejected' ? 'рад · отказ' : 'имзо · виза') +
                         (p.comment ? ': ' + String(p.comment).slice(0, 200) : '')};
      case 'markPaid':
        r = findRequest(p.requestId) || {};
        return {object: 'Дархост № ' + (r.number || p.requestId),
                details: (p.amount ? p.amount + ' ' : '') + (p.date || '')};
      case 'importRegistry':
        return {object: 'Реестр', details: 'шартномаҳо: ' + ((p.contracts || []).length)};
      case 'saveBudget':
        var sum = 0;
        (p.rows || []).forEach(function (b) { for (var m = 1; m <= 12; m++) sum += Number(b['m' + m]) || 0; });
        return {object: 'Буҷет ' + (p.year || '') + ' · ' + (p.company || '') + (p.ceh ? ' · ' + p.ceh : ''),
                details: 'нақша: ' + sum};
      case 'saveCehs':
        return {object: 'Сехҳо · цеха', details: 'сехҳо: ' + ((p.cehs || []).length)};
      case 'saveCodes':
        return {object: 'Моддаҳои буҷет · статьи бюджета', details: 'моддаҳо: ' + ((p.codes || []).length)};
      case 'saveProtocol':
        x = p.protocol || {};
        return {object: 'Протокол № ' + (x.number || ''), details: x.id ? 'таҳрир · правка' : 'нав · новый'};
      case 'deleteProtocol':
        x = readAll('protocols').filter(function (k) { return String(k.id) === String(p.protocolId); })[0] || {};
        return {object: 'Протокол № ' + (x.number || p.protocolId), details: ''};
      case 'uploadDoc':
        r = findRequest(p.requestId) || {};
        return {object: 'Дархост № ' + (r.number || p.requestId), details: (p.kind || '') + ': ' + (p.name || '')};
      case 'deleteDoc':
        x = readAll('docs').filter(function (k) { return String(k.id) === String(p.docId); })[0] || {};
        r = findRequest(x.requestId) || {};
        return {object: 'Дархост № ' + (r.number || ''), details: (x.kind || '') + ': ' + (x.name || '')};
      default:
        return {object: '', details: ''};
    }
  } catch (e) {
    return {object: '', details: ''};
  }
}

function writeLog(who, action, subj) {
  try {
    sheet('log').appendRow([
      uid('l'), new Date().toISOString(),
      (who && who.key) || '', (who && who.name) || '',
      action, (subj && subj.object) || '', (subj && subj.details) || ''
    ]);
  } catch (e) { /* журнал не должен ронять само действие */ }
}

/* ──────────────── справочник статей бюджета ────────────────
 * Приходит целиком. Код — только цифры и уникален, у статьи один уровень
 * вложенности. Статью, на которую уже ссылаются заявки, контракты или план,
 * удалить нельзя — иначе суммы молча выпали бы из бюджета.
 */
function saveCodes(p) {
  auth(p);
  var list = (p.codes || []).map(function (c, i) {
    return {
      code: String(c.code || '').replace(/\s+/g, ''),
      name: String(c.name || '').trim(),
      ru: String(c.ru || '').trim(),
      parent: String(c.parent || '').replace(/\s+/g, ''),
      order: i + 1
    };
  });
  if (!list.length) throw new Error('Рӯйхат холӣ · список статей пуст');

  var seen = {};
  list.forEach(function (c) {
    if (!/^\d{1,10}$/.test(c.code)) throw new Error('Рамз бояд рақам бошад · код из цифр: «' + c.code + '»');
    if (seen[c.code]) throw new Error('Рамз такрор шуд · код повторяется: ' + c.code);
    if (!c.name) throw new Error('Номгӯй холӣ · нет названия у кода ' + c.code);
    seen[c.code] = c;
  });
  list.forEach(function (c) {
    if (!c.parent) return;
    if (c.parent === c.code) throw new Error('Модда наметавонад гурӯҳи худ бошад · ' + c.code);
    if (!seen[c.parent]) throw new Error('Гурӯҳ ёфт нашуд · нет группы ' + c.parent + ' для ' + c.code);
    if (seen[c.parent].parent) throw new Error('Танҳо як сатҳ · группа ' + c.parent + ' сама входит в группу');
  });

  // удаляемые статьи не должны быть в работе
  var used = {};
  readAll('requests').forEach(function (r) { if (r.code) used[r.code] = 'дархост № ' + r.number; });
  readAll('contracts').forEach(function (c) { if (c.code) used[c.code] = used[c.code] || 'шартнома № ' + c.no; });
  readAll('budget').forEach(function (b) {
    var sum = 0; for (var m = 1; m <= 12; m++) sum += Number(b['m' + m]) || 0;
    if (sum) used[b.code] = used[b.code] || 'нақшаи ' + b.year;
  });
  var old = readAll('codes');
  old.forEach(function (c) {
    if (!seen[c.code] && used[c.code])
      throw new Error('Рамзи ' + c.code + ' истифода мешавад (' + used[c.code] + ') · статья в работе, удалить нельзя');
  });

  var sh = sheet('codes');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, COLS.codes.length).clearContent();
  sh.getRange(2, 1, list.length, COLS.codes.length)
    .setValues(list.map(function (c) { return toRow('codes', c); }));
  return {saved: list.length};
}

/* Справочник цехов. Цех, который уже стоит в заявке, контракте или плане,
   удалить нельзя. */
function saveCehs(p) {
  auth(p);
  var seen = {};
  var list = (p.cehs || []).map(function (n) { return String(n || '').trim(); })
    .filter(function (n) {
      if (!n) return false;
      var k = n.toLowerCase();
      if (seen[k]) throw new Error('Сех такрор шуд · цех повторяется: ' + n);
      seen[k] = true;
      return true;
    });
  var used = {};
  readAll('requests').forEach(function (r) { if (r.ceh) used[r.ceh] = 'дархост № ' + r.number; });
  readAll('contracts').forEach(function (c) { if (c.ceh) used[c.ceh] = used[c.ceh] || 'шартнома № ' + c.no; });
  readAll('budget').forEach(function (b) { if (b.ceh) used[b.ceh] = used[b.ceh] || 'нақшаи ' + b.year; });
  readAll('cehs').forEach(function (c) {
    if (!seen[String(c.name).toLowerCase()] && used[c.name])
      throw new Error('Сехи «' + c.name + '» истифода мешавад (' + used[c.name] + ') · цех в работе, удалить нельзя');
  });
  var sh = sheet('cehs');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, COLS.cehs.length).clearContent();
  if (list.length) sh.getRange(2, 1, list.length, 2).setValues(list.map(function (n, i) { return [n, i + 1]; }));
  return {saved: list.length};
}

/* ──────────────── протоколы Совета директоров ──────────────── */

function saveProtocol(p) {
  auth(p);
  var x = p.protocol || {};
  var o = {
    id: x.id || uid('pr'),
    number: String(x.number || '').trim(),
    date: x.date || '',
    company: x.company || '',
    subject: x.subject || '',
    decisions: x.decisions || '',
    fileUrl: x.fileUrl || '',
    fileName: x.fileName || '',
    createdAt: x.createdAt || new Date().toISOString()
  };
  if (!o.number) throw new Error('Рақами протокол лозим аст · нужен номер протокола');
  upsert('protocols', o);
  return {id: o.id};
}

function deleteProtocol(p) {
  auth(p);
  var id = String(p.protocolId || '');
  var row = rowIndexById('protocols', id);
  if (row < 0) throw new Error('Протокол ёфт нашуд');
  // контракты не удаляем, только снимаем ссылку
  readAll('contracts').forEach(function (c) {
    if (String(c.protocolId) === id) { c.protocolId = ''; upsert('contracts', c); }
  });
  sheet('protocols').deleteRow(row);
  return {deleted: id};
}

/* ──────────── подтверждающие документы заявки ──────────────
 * Файл приходит в base64, ложится в папку рядом с таблицей и
 * открывается по ссылке: подписанты входят по логину, а не через Google.
 */
function docsFolder() {
  var name = 'Ҳуҷҷатҳои дархостҳо';
  var parents = DriveApp.getFileById(book().getId()).getParents();
  var root = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function uploadDoc(p) {
  var s = auth(p);
  var reqId = String(p.requestId || '');
  if (!reqId) throw new Error('Дархост нишон дода нашуд');
  if (!p.data) throw new Error('Файл холӣ аст · файл пустой');

  var blob = Utilities.newBlob(
    Utilities.base64Decode(p.data),
    p.mime || 'application/octet-stream',
    p.name || 'hujjat'
  );
  var file = docsFolder().createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* домен может запрещать внешний доступ — ссылка останется внутренней */ }

  var o = {
    id: uid('d'), requestId: reqId,
    kind: p.kind || 'Дигар',
    name: file.getName(), url: file.getUrl(),
    size: file.getSize(), addedAt: new Date().toISOString(),
    byName: s.name || ''
  };
  upsert('docs', o);
  return o;
}

function deleteDoc(p) {
  auth(p);
  var id = String(p.docId || '');
  var row = rowIndexById('docs', id);
  if (row < 0) throw new Error('Ҳуҷҷат ёфт нашуд');
  var rec = readAll('docs').filter(function (d) { return d.id === id; })[0];
  if (rec && rec.url) {
    try {
      var m = /[-\w]{25,}/.exec(rec.url);
      if (m) DriveApp.getFileById(m[0]).setTrashed(true);
    } catch (e) { /* файла может уже не быть — строку всё равно убираем */ }
  }
  sheet('docs').deleteRow(row);
  return {deleted: id};
}

/* ─────────────────────── вход и сессии ───────────────────────
 * Пароль в таблице можно написать открытым текстом: при первом удачном
 * входе он заменяется на хеш, и подсмотреть его больше нельзя.
 * Токен сессии живёт 12 часов и хранится в свойствах скрипта.
 */
var SESSION_HOURS = 12;

function sha256(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}
function isHash(s) { return /^[0-9a-f]{64}$/.test(String(s || '')); }

/** Маршрут согласования: все, кроме администраторов. */
function isAdmin(s) { return String((s && s.role) || '').trim().toLowerCase() === 'admin'; }
function routeSigners() {
  return readAll('signers').filter(function (s) { return !isAdmin(s); });
}

function findSigner(field, value) {
  var signers = readAll('signers');
  for (var i = 0; i < signers.length; i++) {
    if (String(signers[i][field] || '').trim().toLowerCase() === String(value || '').trim().toLowerCase()) {
      return signers[i];
    }
  }
  return null;
}

function login(p) {
  var s = findSigner('login', p.login);
  var pass = String(p.password || '');
  if (!s || !pass) throw new Error('Логин ё рамз нодуруст · неверный логин или пароль');

  var stored = String(s.password || '').trim();
  if (!stored) throw new Error('Барои ин корбар рамз таъин нашудааст · для этого пользователя не задан пароль');

  var ok = isHash(stored) ? (sha256(pass) === stored) : (pass === stored);
  if (!ok) throw new Error('Логин ё рамз нодуруст · неверный логин или пароль');

  if (!isHash(stored)) {            // первый вход — прячем пароль
    s.password = sha256(pass);
    upsert('signers', s);
  }

  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var until = Date.now() + SESSION_HOURS * 3600 * 1000;
  PropertiesService.getScriptProperties()
    .setProperty('s_' + token, JSON.stringify({key: s.key, until: until}));
  return {
    token: token, expiresAt: new Date(until).toISOString(),
    role: {key: s.key, name: s.name, posTj: s.posTj, posRu: s.posRu, order: Number(s.order) || 0,
           admin: isAdmin(s)}
  };
}

function logout(p) {
  if (p && p.token) PropertiesService.getScriptProperties().deleteProperty('s_' + p.token);
  return {ok: true};
}

/**
 * Смена пароля самим подписантом: старый пароль + новый.
 * Новый сразу кладётся отпечатком, прочитать его в таблице нельзя.
 * Остальные сессии этого человека закрываются — на случай чужого входа.
 */
function changePassword(p) {
  var s = auth(p);
  var cur = String((p && p.oldPassword) || '');
  var nw  = String((p && p.newPassword) || '');

  var stored = String(s.password || '').trim();
  var ok = isHash(stored) ? (sha256(cur) === stored) : (cur === stored);
  if (!ok) throw new Error('Рамзи кӯҳна нодуруст · старый пароль неверен');

  if (nw.length < 6) throw new Error('Рамзи нав бояд на кам аз 6 аломат бошад · новый пароль не короче 6 символов');
  if (nw === cur) throw new Error('Рамзи нав бояд фарқ кунад · новый пароль должен отличаться от старого');

  s.password = sha256(nw);
  upsert('signers', s);

  // Закрываем прочие сессии этого подписанта, текущую оставляем.
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('s_') !== 0 || k === 's_' + p.token) return;
    try {
      if (JSON.parse(all[k]).key === s.key) props.deleteProperty(k);
    } catch (e) { /* мусорное свойство — не трогаем */ }
  });

  return {ok: true};
}

/** Возвращает подписанта по токену; бросает ошибку, если сессия негодна. */
function auth(p) {
  var raw = p && p.token
    ? PropertiesService.getScriptProperties().getProperty('s_' + p.token) : null;
  if (!raw) throw new Error('Сессия тамом шуд · сессия истекла, войдите заново');
  var sess = JSON.parse(raw);
  if (!sess.until || Date.now() > sess.until) {
    PropertiesService.getScriptProperties().deleteProperty('s_' + p.token);
    throw new Error('Сессия тамом шуд · сессия истекла, войдите заново');
  }
  var s = findSigner('key', sess.key);
  if (!s) throw new Error('Имзокунанда ёфт нашуд · подписант не найден');
  return s;
}

/* ─────────────────────────── запись ─────────────────────────── */

/** Раскладывает контракт сайта по колонкам листа, включая график траншей. */
function contractRow(c) {
  var n = function (v) { return Number(v) || 0; };
  var row = {
    id: c.id, no: c.no || '', date: c.date || '', supplier: c.supplier || '',
    country: c.country || '', object: c.object || '', product: c.product || '',
    currency: c.currency || '', rate: n(c.rate), bank: c.bank || '',
    invoice: n(c.invoice), gtd: c.gtd || '', file: c.file || '',
    reqAmount: n(c.reqAmount), reqTjs: n(c.reqTjs),
    allocAmount: n(c.allocAmount), allocTjs: n(c.allocTjs),
    req2Rate: n(c.req2Rate), req2Amount: n(c.req2Amount), req2Tjs: n(c.req2Tjs),
    alloc2Rate: n(c.alloc2Rate), alloc2Amount: n(c.alloc2Amount), alloc2Tjs: n(c.alloc2Tjs),
    allocTotal: n(c.allocTotal), code: c.code || '', protocolId: c.protocolId || '', ceh: c.ceh || ''
  };
  var plan = c.plan || [];
  for (var i = 0; i < 3; i++) {
    row['plan' + (i + 1) + 'Cond'] = plan[i] ? (plan[i].cond || '') : '';
    row['plan' + (i + 1) + 'Sum'] = plan[i] ? n(plan[i].amount) : '';
  }
  return row;
}

function saveContract(p) {
  auth(p);
  var c = p.contract || {};
  if (!c.no) throw new Error('Рақами шартнома холӣ · не указан номер контракта');
  if (!c.id) c.id = uid('c');
  upsert('contracts', contractRow(c));
  return {id: c.id};
}

function saveRequest(p) {
  auth(p);
  var r = p.request || {};
  if (!r.number) throw new Error('Рақами дархост холӣ · не указан номер заявки');
  var isNew = !r.id;
  if (isNew) r.id = uid('r');

  var existing = isNew ? null : findRequest(r.id);
  if (existing && existing.submittedAt && !p.force) {
    throw new Error('Дархост аллакай ба мувофиқа фиристода шуд · заявка уже на согласовании, править нельзя');
  }
  // Правка отправленной заявки отменяет уже поставленные визы:
  // подписывали другой документ.
  if (existing && existing.submittedAt && p.force) removeWhere('approvals', 'requestId', r.id);

  upsert('requests', {
    id: r.id, number: r.number, date: r.date || '', company: r.company || '',
    contractId: r.contractId || '', currency: r.currency || '', rate: Number(r.rate) || 0,
    purpose: r.purpose || '', purposeShort: r.purposeShort || '',
    submittedAt: existing ? existing.submittedAt : '', paidAt: existing ? existing.paidAt : '',
    createdAt: existing ? existing.createdAt : new Date().toISOString(),
    code: r.code || '', tranche: r.tranche || '', ceh: r.ceh || ''
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
  auth(p);
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
  var s = auth(p);
  p.roleKey = s.key;   // подписать можно только за себя
  var r = findRequest(p.requestId);
  if (!r) throw new Error('Дархост ёфт нашуд · заявка не найдена');
  if (!r.submittedAt) throw new Error('Дархост ҳанӯз фиристода нашуд · заявка ещё не отправлена на согласование');
  if (r.paidAt) throw new Error('Дархост пардохт шуд · заявка уже оплачена');

  if (isAdmin(s)) throw new Error('Маъмур имзо намегузорад · администратор визы не ставит');
  var signers = routeSigners().sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
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
  auth(p);
  var r = findRequest(p.requestId);
  if (!r) throw new Error('Дархост ёфт нашуд · заявка не найдена');
  if (r.paidAt) throw new Error('Пардохт аллакай қайд шуд · оплата уже отмечена');

  var signers = routeSigners();
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

/**
 * Удаление контракта. Заявки по нему остаются, но теряют привязку —
 * иначе из бюджета пропали бы уже оплаченные суммы.
 */
function deleteContract(p) {
  auth(p);
  var id = String(p.contractId || '');
  var sh = sheet('contracts');
  var row = rowIndexById('contracts', id);
  if (row < 0) throw new Error('Шартнома ёфт нашуд · контракт не найден');

  var reqs = readAll('requests').filter(function (r) { return String(r.contractId) === id; });
  reqs.forEach(function (r) { r.contractId = ''; upsert('requests', r); });

  removeWhere('payments', 'contractId', id);
  sh.deleteRow(row);
  return {deleted: id, unlinked: reqs.length};
}

function deleteRequest(p) {
  auth(p);
  var r = findRequest(p.requestId);
  if (!r) throw new Error('Дархост ёфт нашуд');
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
  auth(p);
  var list = p.contracts || [];
  if (!list.length) throw new Error('Рӯйхат холӣ · список контрактов пуст');

  var shC = sheet('contracts'), shP = sheet('payments');
  if (shC.getLastRow() > 1) shC.getRange(2, 1, shC.getLastRow() - 1, COLS.contracts.length).clearContent();
  if (shP.getLastRow() > 1) shP.getRange(2, 1, shP.getLastRow() - 1, COLS.payments.length).clearContent();

  var crows = [], prows = [], pn = 0;
  list.forEach(function (c) {
    var cid = c.id || uid('c');
    c.id = cid;
    crows.push(toRow('contracts', contractRow(c)));
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
 * Пароли после этого впишите вручную в лист «Имзокунандагон», колонка password:
 * при первом входе скрипт заменит их отпечатком SHA-256.
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
      ['branch','Муовини директори генералӣ оид ба самти паррандапарварӣ ҶДММ «Мармарӣ»','Зам. ген. директора по птицеводству','Ҳакимов З.','hakimov','',1],
      ['finance','Муовини директори генералӣ оиди молияи ҶДММ «Мармарӣ»','Зам. ген. директора по финансам','Ибодов Д.М.','ibodov','',2],
      ['chiefacc','Сармуҳосиби ҶДММ «Мармарӣ»','Главный бухгалтер','Зоиров С.А.','zoirov','',3],
      ['initiator','Директори генералии ҶДММ «Мармарӣ»','Ген. директор предприятия-заявителя','Турсунов А.Г.','tursunov','',4],
      ['holding','Директори генералӣ ҶДММ «Авесто Гуруҳ»','Ген. директор холдинга — утверждение','Муҳамад Мирзо','mirzo','',5]
    ];
    sheet('signers').getRange(2, 1, route.length, COLS.signers.length).setValues(route);
  }
  SpreadsheetApp.getActive().toast('Листҳо тайёр · структура создана. Впишите пароли в лист «Имзокунандагон».');
}
