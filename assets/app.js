/* Мувофиқасозии дархостҳо — клиентская часть.
   Данные: Google Sheets через веб-приложение Apps Script (см. config.js).
   Разработано Отделом автоматизации холдинга ООО «Авесто Групп» · ISJ */
(function () {
"use strict";

var CFG = window.AVESTO_CONFIG || {};
var API = String(CFG.API_URL || "").trim();
var SHEET = String(CFG.SHEET_ID || "").trim();
/* три режима: полный (есть API), только чтение таблицы (есть SHEET_ID),
   автономный (нет ни того, ни другого — данные в этом браузере) */
var READONLY = !API && !!SHEET;
var OFFLINE = !API && !SHEET;

/* ─── маршрут по умолчанию (пока справочник не пришёл с сервера) ─── */
var DEFAULT_SIGNERS = [
  {key:"branch",    posTj:"Муовини директори генералӣ оид ба самти паррандапарварӣ ҶДММ «Мармарӣ»", posRu:"Зам. ген. директора по птицеводству", name:"Ҳакимов З.", order:1},
  {key:"finance",   posTj:"Муовини директори генералӣ оиди молияи ҶДММ «Мармарӣ»", posRu:"Зам. ген. директора по финансам", name:"Ибодов Д.М.", order:2},
  {key:"chiefacc",  posTj:"Сармуҳосиби ҶДММ «Мармарӣ»", posRu:"Главный бухгалтер", name:"Зоиров С.А.", order:3},
  {key:"initiator", posTj:"Директори генералии ҶДММ «Мармарӣ»", posRu:"Ген. директор предприятия-заявителя", name:"Турсунов А.Г.", order:4},
  {key:"holding",   posTj:"Директори генералӣ ҶДММ «Авесто Гуруҳ»", posRu:"Ген. директор холдинга — утверждение", name:"Муҳамад Мирзо", order:5}
];
/* Порядок подписей на бумажном бланке — как в исходном Excel */
var SHEET_ORDER = ["holding","initiator","chiefacc","finance","branch"];

var STATUS = {
  draft:    {ru:"Лоиҳа · черновик", short:"Лоиҳа", cls:"draft"},
  review:   {ru:"Дар мувофиқа · на согласовании", short:"Дар мувофиқа", cls:"review"},
  approved: {ru:"Тасдиқшуда · утверждена", short:"Тасдиқшуда", cls:"approved"},
  rejected: {ru:"Радшуда · отклонена", short:"Радшуда", cls:"rejected"},
  paid:     {ru:"Пардохтшуда · оплачена", short:"Пардохтшуда", cls:"paid"}
};

/* ═══ таджикские числительные и даты ═══ */
var U=["","як","ду","се","чор","панҷ","шаш","ҳафт","ҳашт","нӯҳ"];
var TEEN=["даҳ","ёздаҳ","дувоздаҳ","сенздаҳ","чордаҳ","понздаҳ","шонздаҳ","ҳабдаҳ","ҳаждаҳ","нуздаҳ"];
var TENS=["","","бист","сӣ","чил","панҷоҳ","шаст","ҳафтод","ҳаштод","навад"];
var HUND=["","сад","дусад","сесад","чорсад","панҷсад","шашсад","ҳафтсад","ҳаштсад","нӯҳсад"];
var SCALE=["","ҳазор","миллион","миллиард"];
function joinU(parts){
  var a=parts.filter(Boolean);
  return a.map(function(p,i){return i<a.length-1?p+"у":p;}).join(" ");
}
function under1000(n){
  var p=[];
  if(n>=100){p.push(HUND[Math.floor(n/100)]);n%=100;}
  if(n>=20){p.push(TENS[Math.floor(n/10)]);n%=10;}
  else if(n>=10){p.push(TEEN[n-10]);n=0;}
  if(n>0)p.push(U[n]);
  return joinU(p);
}
function tjWords(n){
  n=Math.floor(Math.abs(Number(n)||0));
  if(n===0)return "сифр";
  var g=[],s=0;
  while(n>0&&s<4){g.unshift({v:n%1000,s:s});n=Math.floor(n/1000);s++;}
  return joinU(g.filter(function(x){return x.v>0;}).map(function(x){
    return under1000(x.v)+(SCALE[x.s]?" "+SCALE[x.s]:"");
  }));
}
var MON_TJ=["Январи","Феврали","Марти","Апрели","Майи","Июни","Июли","Августи","Сентябри","Октябри","Ноябри","Декабри"];
function dTj(iso){
  if(!iso)return "—";
  var d=new Date(iso); if(isNaN(d))return "—";
  return "«"+String(d.getDate()).padStart(2,"0")+"» "+MON_TJ[d.getMonth()]+" "+d.getFullYear()+"с";
}
function dRu(iso){
  if(!iso)return "—";
  var d=new Date(iso); if(isNaN(d))return "—";
  return String(d.getDate()).padStart(2,"0")+"."+String(d.getMonth()+1).padStart(2,"0")+"."+d.getFullYear();
}
function dStamp(iso){
  if(!iso)return "";
  var d=new Date(iso); if(isNaN(d))return "";
  return dRu(iso)+" "+String(d.getHours()).padStart(2,"0")+":"+String(d.getMinutes()).padStart(2,"0");
}
function nf(v,dec){
  var n=Number(v)||0;
  return n.toLocaleString("ru-RU",{minimumFractionDigits:dec||0,maximumFractionDigits:dec===undefined?2:dec});
}
function money(v,cur){return nf(v,2)+(cur?" "+cur:"");}
function esc(s){return String(s==null?"":s).replace(/[&<>"']/g,function(c){
  return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];});}

/* ═══ расчёты ═══ */
function reqTotalCur(r){return (r.items||[]).reduce(function(s,i){return s+(Number(i.qty)||0)*(Number(i.price)||0);},0);}
function reqTotalTjs(r){return reqTotalCur(r)*(Number(r.rate)||0);}
function paidSum(c){return (c.payments||[]).reduce(function(s,p){return s+(Number(p.amount)||0);},0);}
/* Эквивалент в сомони берём по курсу самой оплаты, а не по курсу заявки:
   курс выделения и курс платежа различаются, иначе итог расходится с реестром. */
function paidTjs(c){
  return (c.payments||[]).reduce(function(s,p){
    var t=Number(p.tjs)||0;
    return s+(t||(Number(p.amount)||0)*(Number(p.rate)||Number(c.rate)||0));
  },0);
}
function balance(c){return (Number(c.invoice)||0)-paidSum(c);}
/* Порядок виз свободный: первые четверо подписывают когда угодно,
   ген. директор холдинга утверждает последним — только после остальных. */
function finalKey(){return S.signers.length?S.signers[S.signers.length-1].key:null;}
function signed(r,key){var v=(r.approvals||{})[key];return !!(v&&v.decision==="approved");}
function allSigned(r){
  if(!S.signers.length)return false;
  for(var i=0;i<S.signers.length;i++) if(!signed(r,S.signers[i].key))return false;
  return true;
}
function othersSigned(r){
  var fk=finalKey();
  for(var i=0;i<S.signers.length;i++){
    var k=S.signers[i].key;
    if(k!==fk&&!signed(r,k))return false;
  }
  return true;
}
function pendingNames(r,skipFinal){
  var fk=finalKey(),out=[];
  for(var i=0;i<S.signers.length;i++){
    var p=S.signers[i];
    if(skipFinal&&p.key===fk)continue;
    if(!signed(r,p.key))out.push(p.name);
  }
  return out;
}
function canSign(r,key){
  if(!key||key==="observer")return false;
  if(derivedStatus(r)!=="review")return false;
  if(signed(r,key))return false;
  if(!signerBy(key))return false;
  if(key===finalKey())return othersSigned(r);
  return true;
}
function derivedStatus(r){
  var a=r.approvals||{},keys=Object.keys(a);
  for(var i=0;i<keys.length;i++) if(a[keys[i]]&&a[keys[i]].decision==="rejected")return "rejected";
  if(r.paidAt)return "paid";
  if(allSigned(r))return "approved";
  return r.submittedAt?"review":"draft";
}
function signerBy(key){
  for(var i=0;i<S.signers.length;i++) if(S.signers[i].key===key)return S.signers[i];
  return null;
}

/* ═══ состояние ═══ */
var S={
  tab:"requests", filter:"all",
  contracts:[], requests:[], signers:DEFAULT_SIGNERS.slice(),
  selReq:null, selCon:null, sheetFor:null,
  q:"", cur:"all", rest:"all", conMode:"cards",
  me:null, busy:false, loaded:false, error:null
};
function lsGet(k){try{return localStorage.getItem(k);}catch(e){return null;}}
function lsSet(k,v){try{localStorage.setItem(k,v);}catch(e){}}
function ssGet(k){try{return sessionStorage.getItem(k);}catch(e){return null;}}
function ssSet(k,v){try{sessionStorage.setItem(k,v);}catch(e){}}

/* ═══ транспорт ═══
   POST отправляется с Content-Type: text/plain намеренно: это «простой»
   запрос, браузер не шлёт предварительный OPTIONS, который Apps Script
   не умеет обрабатывать. Тело при этом остаётся JSON. */
function apiGet(){
  return fetch(API+"?t="+Date.now(),{method:"GET"})
    .then(function(r){return r.json();})
    .then(function(j){ if(!j.ok)throw new Error(j.error||"Хатогии сервер"); return j.data; });
}
function apiPost(action,payload){
  return fetch(API,{
    method:"POST",
    headers:{"Content-Type":"text/plain;charset=utf-8"},
    body:JSON.stringify({action:action,payload:payload})
  }).then(function(r){return r.json();})
    .then(function(j){ if(!j.ok)throw new Error(j.error||"Хатогии сервер"); return j; });
}

/* автономный режим — данные в этом браузере */
function localLoad(){
  try{
    var raw=lsGet("avesto.soglasovanie.data");
    if(raw){var d=JSON.parse(raw);S.contracts=d.contracts||[];S.requests=d.requests||[];}
  }catch(e){}
}
function localSave(){
  try{lsSet("avesto.soglasovanie.data",JSON.stringify({contracts:S.contracts,requests:S.requests}));}catch(e){}
}

function applyState(d){
  S.contracts=d.contracts||[];
  S.requests=d.requests||[];
  if(d.signers&&d.signers.length)S.signers=d.signers;
  S.loaded=true; S.error=null;
}
function reload(){
  if(OFFLINE){localLoad();S.loaded=true;render();return Promise.resolve();}
  var p = API ? apiGet() : loadFromSheet(SHEET);
  return p.then(function(d){applyState(d);render();})
    .catch(function(e){S.loaded=true;S.error=e.message||String(e);render();});
}
function send(action,payload,okMsg){
  if(READONLY){toast("Режими танҳо дидан · режим только просмотра: запись в таблицу не подключена",true);return Promise.reject();}
  if(OFFLINE){toast("Базаи маълумот пайваст нашуд · база не подключена",true);return Promise.reject();}
  S.busy=true;render();
  return apiPost(action,payload).then(function(j){
    if(j.state)applyState(j.state);
    S.busy=false;render();
    if(okMsg)toast(okMsg);
    return j.data;
  }).catch(function(e){
    S.busy=false;render();
    toast(e.message||"Хатогӣ",true);
    throw e;
  });
}

/* ═══ PIN ═══ */
function cachedPin(role){return ssGet("avesto.pin."+role)||"";}
function cachePin(role,pin){ssSet("avesto.pin."+role,pin);}

/* ═══ рендер ═══ */
var view=document.getElementById("view");
function render(){
  view.innerHTML=(OFFLINE?offlineNotice():(READONLY?readonlyNotice():""))+(S.error?errorNotice():"")+
    (S.tab==="requests"?renderRequests():renderContracts());
  wire();
  var btn=document.getElementById("meBtn");
  var s=S.me&&signerBy(S.me);
  btn.textContent=s?s.name:(S.me==="observer"?"Нозир":"Кто вы?");
  btn.title=s?s.posRu:"Выбрать роль";
}
function offlineNotice(){
  return '<div class="notice"><span>⚠</span><span><b>База не подключена.</b> '+
    'Данные видны только в этом браузере. Загрузите реестр кнопкой <b>«Импорт Excel»</b> '+
    'или пропишите таблицу в <code>assets/config.js</code> — инструкция в README.</span></div>';
}
function readonlyNotice(){
  return '<div class="notice"><span>👁</span><span><b>Режим просмотра.</b> '+
    'Данные читаются прямо из таблицы Google Sheets. Визы и платежи записывать нельзя — '+
    'для этого нужно развернуть веб-приложение Apps Script и указать его адрес в <code>config.js</code>.</span></div>';
}
function errorNotice(){
  return '<div class="notice bad"><span>✕</span><span><b>Сервер недоступен.</b> '+esc(S.error)+
    ' <button class="btn sm" data-act="retry" style="margin-left:6px">Такрор</button></span></div>';
}

function renderRequests(){
  var list=S.requests.slice().sort(function(a,b){return String(b.date||"").localeCompare(String(a.date||""));});
  var shown=list.filter(function(r){return S.filter==="all"||derivedStatus(r)===S.filter;});
  var sel=null;
  for(var i=0;i<shown.length;i++) if(shown[i].id===S.selReq)sel=shown[i];
  if(!sel)sel=shown[0]||null;
  if(sel)S.selReq=sel.id;
  if(S.sheetFor&&sel&&S.sheetFor===sel.id)return sheetView(sel);

  var fdefs=[["all","Ҳама"],["draft","Лоиҳа"],["review","Дар мувофиқа"],["approved","Тасдиқшуда"],["paid","Пардохтшуда"],["rejected","Радшуда"]];
  return '<div class="cols"><div class="panel">'+
    '<div class="head"><h2>Дархостҳо</h2><span class="count">'+shown.length+" / "+list.length+'</span></div>'+
    '<div class="filters">'+fdefs.map(function(f){
      return '<button data-filter="'+f[0]+'" aria-pressed="'+(S.filter===f[0])+'">'+f[1]+"</button>";
    }).join("")+"</div>"+
    (shown.length?'<ul class="list">'+shown.map(function(r){return reqRow(r,sel);}).join("")+"</ul>":
      '<div class="empty"><h3>Дархост нест</h3><p>'+
      (S.loaded?"По этому фильтру заявок нет.":'<span class="spin"></span> Маълумот бор мешавад…')+"</p></div>")+
    '<div class="foot"><button class="btn sm" id="newReq">+ Дархости нав</button></div></div>'+
    (sel?reqDoc(sel):'<div class="panel"><div class="empty"><h3>Дархост интихоб нашуд</h3>'+
      "<p>Выберите заявку слева — откроются позиции, маршрут согласования и печатный бланк на таджикском.</p></div></div>")+
    "</div>";
}
function reqRow(r,sel){
  var st=derivedStatus(r);
  var dots=S.signers.map(function(p){
    var v=(r.approvals||{})[p.key];
    if(v&&v.decision==="approved")return '<i class="on"></i>';
    if(v&&v.decision==="rejected")return '<i class="no"></i>';
    if(st==="review"&&canSign(r,p.key))return '<i class="now"></i>';
    return "<i></i>";
  }).join("");
  var c=null;
  for(var i=0;i<S.contracts.length;i++) if(S.contracts[i].id===r.contractId)c=S.contracts[i];
  return '<li><button class="row" data-req="'+esc(r.id)+'" aria-current="'+(sel&&sel.id===r.id)+'">'+
    '<span class="l1"><b>№ '+esc(r.number||"—")+'</b><span class="badge '+STATUS[st].cls+'">'+STATUS[st].short+"</span>"+
    "<time>"+dRu(r.date)+"</time></span>"+
    '<span class="l2">'+esc(r.purpose||(c&&c.product)||"—")+"</span>"+
    '<span class="l3"><span class="dots">'+dots+'</span><span class="sum">'+money(reqTotalCur(r),r.currency)+"</span></span>"+
    "</button></li>";
}
function reqDoc(r){
  var st=derivedStatus(r);
  var c=null;
  for(var i=0;i<S.contracts.length;i++) if(S.contracts[i].id===r.contractId)c=S.contracts[i];
  var cur=reqTotalCur(r),tjs=reqTotalTjs(r);
  var myTurn=canSign(r,S.me);
  var canSubmit=st==="draft";
  var canPay=st==="approved"&&c&&S.me&&S.me!=="observer";

  return '<div class="panel doc">'+
    '<div class="dhead"><div class="t"><h2>Дархости № '+esc(r.number||"—")+"</h2>"+
    '<div class="sub">'+esc(r.company||"—")+" · "+dTj(r.date)+"</div></div>"+
    '<span class="badge '+STATUS[st].cls+'">'+STATUS[st].ru+"</span></div>"+

    "<section><h3>Асос · основание</h3>"+
    '<p class="purpose">'+esc(r.purpose||"—")+"</p>"+
    (c?'<dl class="kv" style="margin-top:14px">'+
      "<div><dt>Шартнома</dt><dd>№ "+esc(c.no)+" · "+dRu(c.date)+"</dd></div>"+
      "<div><dt>Интиқолдиҳанда</dt><dd>"+esc(c.supplier)+(c.country?" · "+esc(c.country):"")+"</dd></div>"+
      '<div><dt>Бақия барои пардохт</dt><dd class="num">'+money(balance(c),c.currency)+"</dd></div></dl>":
      '<div class="hint" style="margin-top:12px">Шартнома пайваст нашуд · контракт не привязан</div>')+
    "</section>"+

    "<section><h3>Маблағ · позиции</h3>"+
    '<div class="tw"><table class="items"><thead><tr><th style="width:34px">№</th><th>Номгӯй</th><th>Воҳ. ченак</th>'+
    '<th class="r">Миқдор</th><th class="r">Нарх, '+esc(r.currency||"")+'</th><th class="r">Ҷамъ, '+esc(r.currency||"")+'</th>'+
    '<th class="r">Ҷамъ, сомонӣ</th></tr></thead><tbody>'+
    ((r.items||[]).map(function(it,i){
      var s=(Number(it.qty)||0)*(Number(it.price)||0);
      return '<tr><td class="num">'+(i+1)+"</td><td>"+esc(it.name)+"</td><td>"+esc(it.unit||"")+"</td>"+
        '<td class="r num">'+nf(it.qty,0)+'</td><td class="r num">'+nf(it.price,2)+"</td>"+
        '<td class="r num">'+nf(s,2)+'</td><td class="r num">'+nf(s*(Number(r.rate)||0),2)+"</td></tr>";
    }).join("")||'<tr><td colspan="7" class="hint">Позиции не заданы</td></tr>')+
    '</tbody><tfoot><tr><td colspan="5">Ҳамагӣ</td><td class="r num">'+nf(cur,2)+'</td><td class="r num">'+nf(tjs,2)+
    "</td></tr></tfoot></table></div>"+
    '<div class="hint" style="margin-top:9px">Курси ҳисобӣ: <b class="num">'+nf(r.rate,4)+"</b> сомонӣ / "+
    esc(r.currency||"")+' · эквивалент <b class="num">'+nf(tjs,2)+"</b> сомонӣ</div></section>"+

    "<section><h3>Масири мувофиқа · маршрут</h3>"+
    '<ol class="route">'+S.signers.map(function(p,i){return routeStep(p,i,r,st);}).join("")+"</ol></section>"+

    '<section><div class="acts">'+
    (canSubmit?'<button class="btn pri" data-act="submit" data-id="'+esc(r.id)+'">Ба мувофиқа фиристодан</button>':"")+
    (myTurn?'<button class="btn pri" data-act="approve" data-id="'+esc(r.id)+'">Имзо гузоштан · завизировать</button>'+
      '<button class="btn dang" data-act="reject" data-id="'+esc(r.id)+'">Рад кардан</button>':"")+
    (canPay?'<button class="btn pri" data-act="pay" data-id="'+esc(r.id)+'">Пардохтро қайд кардан</button>':"")+
    '<button class="btn" data-act="sheet" data-id="'+esc(r.id)+'">Бланк дархост</button>'+
    (st==="draft"?'<button class="btn sm" data-act="editreq" data-id="'+esc(r.id)+'">Таҳрир</button>':"")+
    "</div>"+
    (st==="review"&&!myTurn?'<p class="hint" style="margin-top:10px">'+waitHint(r)+"</p>":"")+
    (st==="paid"?'<p class="hint" style="margin-top:10px">Пардохт қайд шуд '+dStamp(r.paidAt)+".</p>":"")+
    (st==="approved"&&!canPay?'<p class="hint" style="margin-top:10px">Масир пурра шуд. Отметить оплату может любой подписант.</p>':"")+
    "</section></div>";
}
function waitHint(r){
  var fk=finalKey(),tail="";
  if(!S.me||S.me==="observer")tail=" Укажите свою должность в шапке, чтобы визировать.";
  if(S.me&&S.me!=="observer"&&signed(r,S.me)){
    var left=pendingNames(r,false);
    return "Шумо имзо гузоштаед · вы уже подписали. Ожидают подписи: <b>"+esc(left.join(", "))+"</b>.";
  }
  if(S.me===fk&&!othersSigned(r)){
    return "Утверждение — последнее. Сначала должны подписать: <b>"+esc(pendingNames(r,true).join(", "))+"</b>.";
  }
  return "Ожидают подписи: <b>"+esc(pendingNames(r,false).join(", "))+"</b>."+tail;
}
function routeStep(p,i,r,st){
  var v=(r.approvals||{})[p.key];
  var cls="",mark=String(i+1),note="";
  if(v&&v.decision==="approved"){cls="done";mark="✓";}
  else if(v&&v.decision==="rejected"){cls="no";mark="✕";}
  else if(st==="review"&&canSign(r,p.key)){cls="now";}
  else if(st==="review"&&p.key===finalKey()){note="Утверждает после остальных";}
  return '<li class="'+cls+'"><span class="mark">'+mark+'</span><span class="step">'+
    '<span class="pos">'+esc(p.posTj)+'</span><span class="posru">'+esc(p.posRu)+'</span>'+
    '<span class="who">'+esc(p.name)+(note?' · <span style="color:var(--ink-faint)">'+note+"</span>":"")+"</span>"+
    (v?'<span class="vis '+(v.decision==="approved"?"ok":"no")+'">'+
      (v.decision==="approved"?"Имзо гузошт":"Рад кард")+' · <span class="meta">'+dStamp(v.at)+"</span>"+
      (v.byName?" · "+esc(v.byName):"")+
      (v.comment?'<span class="c">«'+esc(v.comment)+"»</span>":"")+"</span>":"")+
    "</span></li>";
}

/* ── печатный бланк ── */
function sheetView(r){
  var cur=reqTotalCur(r),tjs=reqTotalTjs(r);
  var c=null;
  for(var i=0;i<S.contracts.length;i++) if(S.contracts[i].id===r.contractId)c=S.contracts[i];
  var curName=r.currency==="EUR"?"евро":(r.currency==="USD"?"доллар":(r.currency||""));
  var order=S.signers.slice().sort(function(a,b){
    var ia=SHEET_ORDER.indexOf(a.key),ib=SHEET_ORDER.indexOf(b.key);
    return (ia<0?99:ia)-(ib<0?99:ib);
  });
  return '<div class="sheetbar"><button class="btn sm" data-act="closesheet">← Бозгашт</button>'+
    '<button class="btn sm" data-act="print">Чоп кардан · печать</button>'+
    '<span class="hint">Печать браузера: оттуда же сохранение в PDF.</span></div>'+
    '<div class="sheet" id="sheetBody">'+
    '<div class="org">'+esc(r.company||"")+"</div>"+
    '<div class="nomer">Дархости № '+esc(r.number||"")+" аз "+dTj(r.date)+"</div>"+
    '<p class="body">Аз Шумо эҳтиромона хоҳиш менамоем, ки '+esc(r.purposeShort||r.purpose||"")+
    " маблағи зерин дар ҳаҷми "+nf(cur,0)+" ("+tjWords(cur)+") "+esc(curName)+" ҷудо намоед.</p>"+
    '<table><thead><tr><th style="width:32px">№</th><th>Номгӯй</th><th style="width:78px">Воҳ. ченак</th>'+
    '<th class="r" style="width:66px">Миқдор</th><th class="r" style="width:104px">Нарх бо '+esc(curName)+"</th>"+
    '<th class="r" style="width:112px">Ҷамъ бо '+esc(curName)+'</th><th class="r" style="width:118px">Ҷамъ бо сомонӣ</th></tr></thead><tbody>'+
    (r.items||[]).map(function(it,i){
      var s=(Number(it.qty)||0)*(Number(it.price)||0);
      return "<tr><td>"+(i+1)+".0</td><td>"+esc(it.name)+"</td><td>"+esc(it.unit||"")+'</td><td class="r">'+
        nf(it.qty,0)+'</td><td class="r">'+nf(it.price,0)+'</td><td class="r">'+nf(s,0)+'</td><td class="r">'+
        nf(s*(Number(r.rate)||0),0)+"</td></tr>";
    }).join("")+
    '<tr><td></td><td><b>Ҳамагӣ</b></td><td></td><td></td><td></td><td class="r"><b>'+nf(cur,0)+
    '</b></td><td class="r"><b>'+nf(tjs,0)+"</b></td></tr></tbody></table>"+
    '<p class="after">Маблағи мазкур аз хазина дар ҳаҷми '+nf(cur,0)+" ("+tjWords(cur)+") "+esc(curName)+" ҷудо карда шавад.</p>"+
    (c?'<p class="after" style="font-size:13px">Асос: шартномаи № '+esc(c.no)+" аз "+dTj(c.date)+" бо "+
      esc(c.supplier)+" ("+esc(c.country||"")+"). Курси ҳисобӣ: "+nf(r.rate,4)+" сомонӣ.</p>":"")+
    '<div class="sigs">'+order.map(function(p){
      var v=(r.approvals||{})[p.key];
      return '<div class="sig"><span class="pos">'+esc(p.posTj)+'</span><span><span class="ln">'+
        (v?'<span class="vz '+(v.decision==="rejected"?"no":"")+'">'+
          (v.decision==="approved"?"имзо":"рад")+" · "+dRu(v.at)+"</span>":"")+
        '</span><span class="nm"> '+esc(p.name)+"</span></span></div>";
    }).join("")+"</div></div>";
}

/* ── контракты ── */
function conMatches(c){
  var q=(S.q||"").trim().toLowerCase();
  if(q){
    var hay=[c.no,c.supplier,c.product,c.country,c.object].join(" ").toLowerCase();
    if(hay.indexOf(q)<0)return false;
  }
  if(S.cur&&S.cur!=="all"&&(c.currency||"—")!==S.cur)return false;
  if(S.rest==="open"&&balance(c)<=0.0001)return false;
  if(S.rest==="done"&&!((Number(c.invoice)||0)>0&&balance(c)<=0.0001))return false;
  if(S.rest==="over"&&balance(c)>=-0.0001)return false;
  return true;
}
function renderContracts(){
  var all=S.contracts.slice().sort(function(a,b){return String(b.date||"").localeCompare(String(a.date||""));});
  var list=all.filter(conMatches);
  var curs=[];
  all.forEach(function(c){var k=c.currency||"—";if(curs.indexOf(k)<0)curs.push(k);});
  var totalTjs=list.reduce(function(s,c){return s+paidTjs(c);},0);

  var bar='<div class="toolbar">'+
    '<input id="conQ" type="search" placeholder="Поиск: поставщик, номер контракта, продукция" value="'+esc(S.q||"")+'">'+
    '<div class="chips">'+
      '<button data-cur="all" aria-pressed="'+(!S.cur||S.cur==="all")+'">Ҳама асъор</button>'+
      curs.map(function(k){
        return '<button data-cur="'+esc(k)+'" aria-pressed="'+(S.cur===k)+'">'+esc(k)+"</button>";}).join("")+
    '</div><div class="chips">'+
      [["all","Ҳама"],["open","Бо бақия"],["done","Пӯшида"],["over","Зиёдпардохт"]].map(function(f){
        return '<button data-rest="'+f[0]+'" aria-pressed="'+((S.rest||"all")===f[0])+'">'+f[1]+"</button>";}).join("")+
    '</div><div class="chips modes">'+
      [["cards","Корточкаҳо"],["table","Ҷадвал"]].map(function(m){
        return '<button data-mode="'+m[0]+'" aria-pressed="'+((S.conMode||"cards")===m[0])+'">'+m[1]+"</button>";}).join("")+
    "</div></div>";

  if((S.conMode||"cards")==="table")
    return bar+conTable(list)+
      '<div class="panel" style="margin-top:14px"><div class="foot"><button class="btn sm" id="newCon">+ Шартномаи нав</button></div></div>';

  var sel=null;
  for(var i=0;i<list.length;i++) if(list[i].id===S.selCon)sel=list[i];
  if(!sel)sel=list[0]||null;
  if(sel)S.selCon=sel.id;
  return bar+'<div class="cols"><div class="panel">'+
    '<div class="head"><h2>Шартномаҳо</h2><span class="count">'+list.length+" / "+all.length+"</span></div>"+
    (list.length?'<ul class="list">'+list.map(function(c){return conRow(c,sel);}).join("")+"</ul>":
      '<div class="empty"><h3>Ёфт нашуд</h3><p>'+
      (S.loaded?"По этому запросу контрактов нет.":'<span class="spin"></span> Маълумот бор мешавад…')+"</p></div>")+
    (list.length?'<div class="foot" style="font-size:13px;color:var(--ink-soft);display:flex;gap:8px">'+
      '<span>Ҳамагӣ пардохт</span><b class="num" style="margin-left:auto">'+nf(totalTjs,2)+" сом.</b></div>":"")+
    '<div class="foot"><button class="btn sm" id="newCon">+ Шартномаи нав</button></div></div>'+
    (sel?conDoc(sel):'<div class="panel"><div class="empty"><h3>Шартнома интихоб нашуд</h3>'+
      "<p>Выберите контракт — откроется график платежей и остаток к оплате.</p></div></div>")+"</div>";
}
function conRow(c,sel){
  var inv=Number(c.invoice)||0,pd=paidSum(c),bal=balance(c);
  var pct=inv?Math.max(0,Math.min(100,pd/inv*100)):0;
  var state=bal<-0.0001?"over":(inv>0&&bal<=0.0001?"done":(pd>0?"part":"none"));
  var tag={over:'<span class="badge rejected">Зиёд</span>',
           done:'<span class="badge approved">Пӯшида</span>',
           part:"",none:""}[state];
  return '<li><button class="row conrow" data-con="'+esc(c.id)+'" aria-current="'+(sel&&sel.id===c.id)+'">'+
    '<span class="l1"><b>'+esc(c.no||"—")+"</b>"+tag+"<time>"+dRu(c.date)+"</time></span>"+
    '<span class="l2">'+esc(c.supplier||"—")+(c.product?" · "+esc(c.product):"")+"</span>"+
    '<span class="money"><span class="num">'+nf(pd,0)+"</span>"+
    '<span class="sep">/</span><span class="num tot">'+nf(inv,0)+"</span>"+
    '<span class="cur">'+esc(c.currency||"")+"</span>"+
    '<span class="rest '+state+'">'+(bal<-0.0001?"зиёд "+nf(-bal,0):"бақия "+nf(bal,0))+"</span></span>"+
    '<span class="bar"><i class="'+state+'" style="width:'+pct.toFixed(1)+'%"></i></span>'+
    "</button></li>";
}
function conTable(list){
  var tot={};
  list.forEach(function(c){
    var k=c.currency||"—";
    tot[k]=tot[k]||{inv:0,paid:0,rest:0,n:0};
    tot[k].inv+=Number(c.invoice)||0; tot[k].paid+=paidSum(c);
    tot[k].rest+=balance(c); tot[k].n++;
  });
  var body=list.map(function(c,i){
    var bal=balance(c);
    return '<tr data-con="'+esc(c.id)+'" tabindex="0">'+
      '<td class="num">'+(i+1)+"</td>"+
      "<td>"+esc(c.no||"—")+'<span class="sub">'+dRu(c.date)+"</span></td>"+
      "<td>"+esc(c.supplier||"—")+'<span class="sub">'+esc(c.product||"")+"</span></td>"+
      "<td>"+esc(c.currency||"")+'</td><td class="r num">'+nf(c.invoice,0)+"</td>"+
      '<td class="r num">'+nf(paidSum(c),0)+'</td><td class="r num'+(bal<-0.0001?" neg":"")+'">'+
      nf(bal,0)+'</td><td class="r num">'+nf(paidTjs(c),0)+"</td></tr>";
  }).join("");
  var foot=Object.keys(tot).map(function(k){
    var g=tot[k];
    return '<tr class="sum"><td></td><td colspan="2">Итого '+esc(k)+" · "+g.n+'</td><td>'+esc(k)+
      '</td><td class="r num">'+nf(g.inv,0)+'</td><td class="r num">'+nf(g.paid,0)+
      '</td><td class="r num">'+nf(g.rest,0)+"</td><td></td></tr>";
  }).join("");
  return '<div class="panel"><div class="head"><h2>Реестр контрактов</h2>'+
    '<span class="count">'+list.length+"</span></div>"+
    '<div class="tw"><table class="grid"><thead><tr><th>№</th><th>Шартнома</th><th>Интиқолдиҳанда</th>'+
    '<th>Асъор</th><th class="r">Инвойс</th><th class="r">Пардохтшуда</th><th class="r">Бақия</th>'+
    '<th class="r">Сомонӣ</th></tr></thead><tbody>'+
    (body||'<tr><td colspan="8" class="hint">Ничего не найдено</td></tr>')+
    "</tbody><tfoot>"+foot+"</tfoot></table></div></div>";
}
function conDoc(c){
  var inv=Number(c.invoice)||0,pd=paidSum(c),bal=balance(c);
  var linked=S.requests.filter(function(r){return r.contractId===c.id;});
  return '<div class="panel doc"><div class="dhead"><div class="t"><h2>Шартномаи № '+esc(c.no||"—")+"</h2>"+
    '<div class="sub">'+esc(c.supplier||"")+" · "+esc(c.country||"")+" · "+dTj(c.date)+"</div></div>"+
    '<span class="badge '+(bal<-0.0001?"rejected":(bal<=0.0001&&inv>0?"approved":"review"))+'">'+
    (bal<-0.0001?"Зиёдпардохт "+nf(-bal,0)+" "+esc(c.currency||"")
      :(bal<=0.0001&&inv>0?"Пурра пардохт":"Бақия "+nf(bal,0)+" "+esc(c.currency||"")))+"</span></div>"+
    "<section><h3>Маълумот · реквизиты</h3>"+
    '<dl class="kv"><div><dt>Объекти дахлдор</dt><dd>'+esc(c.object||"—")+"</dd></div>"+
    "<div><dt>Маҳсулот</dt><dd>"+esc(c.product||"—")+"</dd></div>"+
    '<div><dt>Асъор / курс</dt><dd>'+esc(c.currency||"—")+' · <span class="num">'+nf(c.rate,4)+"</span></dd></div>"+
    "<div><dt>Бонки пардохткунанда</dt><dd>"+esc(c.bank||"—")+"</dd></div>"+
    '<div><dt>Маблағи ИНВОЙС</dt><dd class="num">'+money(inv,c.currency)+"</dd></div>"+
    "<div><dt>ГТД</dt><dd>"+esc(c.gtd||"—")+"</dd></div>"+
    (c.file?"<div><dt>Файли шартнома</dt><dd>"+esc(c.file)+"</dd></div>":"")+"</dl></section>"+
    (function(){
      var n=function(x){return Number(x)||0;};
      var hasReq=n(c.reqAmount)||n(c.allocAmount)||n(c.req2Amount)||n(c.alloc2Amount)||n(c.allocTotal);
      if(!hasReq)return "";
      return "<section><h3>Заявка ва тақсим · заявка и выделение</h3>"+
        '<dl class="kv">'+
        (n(c.reqAmount)?"<div><dt>Заявка аз рӯи шартнома</dt><dd class=\"num\">"+nf(c.reqAmount,0)+" "+esc(c.currency||"")+"</dd></div>":"")+
        (n(c.allocAmount)?"<div><dt>Ҷудошуда</dt><dd class=\"num\">"+nf(c.allocAmount,0)+" "+esc(c.currency||"")+"</dd></div>":"")+
        (n(c.req2Amount)?"<div><dt>Заявка-2</dt><dd class=\"num\">"+nf(c.req2Amount,0)+
          (n(c.req2Rate)?' <span style="color:var(--ink-faint);font-size:12px">курс '+nf(c.req2Rate,4)+"</span>":"")+"</dd></div>":"")+
        (n(c.alloc2Amount)?"<div><dt>Ҷудошуда-2</dt><dd class=\"num\">"+nf(c.alloc2Amount,0)+
          (n(c.alloc2Rate)?' <span style="color:var(--ink-faint);font-size:12px">курс '+nf(c.alloc2Rate,4)+"</span>":"")+"</dd></div>":"")+
        (n(c.allocTotal)?"<div><dt>Ҳамагӣ ҷудошуда</dt><dd class=\"num\">"+nf(c.allocTotal,0)+"</dd></div>":"")+
        "</dl></section>";
    })()+
    (function(){
      var plan=(c.plan||[]).filter(function(x){return Number(x.amount)||0;});
      if(!plan.length)return "";
      var ps=plan.reduce(function(s,x){return s+(Number(x.amount)||0);},0);
      return "<section><h3>Графики пардохт · график траншей</h3>"+
        '<div class="paylist">'+plan.map(function(x,i){
          return '<div class="pay"><span class="n">'+(i+1)+"</span><span>"+esc(x.cond||"—")+
            '</span><span class="a">'+nf(x.amount,0)+" "+esc(c.currency||"")+"</span></div>";
        }).join("")+"</div>"+
        '<p class="hint" style="margin-top:9px">По графику: <b class="num">'+nf(ps,0)+"</b> "+esc(c.currency||"")+
        " · по инвойсу: <b class=\"num\">"+nf(c.invoice,0)+"</b>"+
        (Math.abs(ps-(Number(c.invoice)||0))>0.5?' · <span style="color:var(--wait)">расхождение '+
          nf((Number(c.invoice)||0)-ps,0)+"</span>":"")+"</p></section>";
    })()+
    "<section><h3>Пардохтҳо · платежи</h3>"+
    '<div class="paylist">'+((c.payments||[]).map(function(p,i){
      return '<div class="pay"><span class="n">'+(i+1)+"</span><span>"+esc(p.note||"Пардохт")+
        (p.date?' · <span class="num" style="font-size:12px;color:var(--ink-faint)">'+dRu(p.date)+"</span>":"")+
        (p.tjs?'<br><span class="num" style="font-size:12px;color:var(--ink-faint)">'+nf(p.tjs,2)+" сомонӣ</span>":"")+
        '</span><span class="a">'+nf(p.amount,2)+" "+esc(c.currency||"")+"</span></div>";
    }).join("")||'<p class="hint">Пардохт сабт нашудааст · платежей нет</p>')+"</div>"+
    '<dl class="kv" style="margin-top:16px"><div><dt>Пардохтшуда</dt><dd class="num">'+money(pd,c.currency)+"</dd></div>"+
    '<div><dt>Бақия барои пардохт</dt><dd class="num" style="color:'+(bal>0.0001?"var(--wait)":"var(--ok)")+'">'+
    money(bal,c.currency)+"</dd></div>"+
    '<div><dt>Эквиваленти сомонӣ</dt><dd class="num">'+nf(paidTjs(c),2)+"</dd></div></dl>"+
    '<div class="bar" style="margin-top:12px;height:9px"><i style="width:'+(inv?Math.min(100,pd/inv*100).toFixed(1):0)+'%"></i></div></section>'+
    "<section><h3>Дархостҳои вобаста · связанные заявки</h3>"+
    (linked.length?'<div class="paylist">'+linked.map(function(r){
      var st=derivedStatus(r);
      return '<div class="pay"><span class="n">№'+esc(r.number)+'</span><span><span class="badge '+
        STATUS[st].cls+'">'+STATUS[st].short+"</span> · "+dRu(r.date)+'</span><span class="a">'+
        nf(reqTotalCur(r),2)+" "+esc(r.currency||"")+"</span></div>";
    }).join("")+"</div>":'<p class="hint">Дархост нест</p>')+
    '<div class="acts" style="margin-top:14px"><button class="btn sm" data-act="editcon" data-id="'+esc(c.id)+'">Таҳрир</button></div>'+
    "</section></div>";
}

/* ═══ события ═══ */
function wire(){
  var i,els;
  els=view.querySelectorAll("[data-filter]");
  for(i=0;i<els.length;i++)(function(b){b.onclick=function(){S.filter=b.dataset.filter;render();};})(els[i]);
  els=view.querySelectorAll("[data-req]");
  for(i=0;i<els.length;i++)(function(b){b.onclick=function(){S.selReq=b.dataset.req;S.sheetFor=null;render();};})(els[i]);
  els=view.querySelectorAll("[data-con]");
  for(i=0;i<els.length;i++)(function(b){
    b.onclick=function(){
      S.selCon=b.dataset.con;
      if((S.conMode||"cards")==="table")S.conMode="cards";
      render();
    };
    if(b.tagName==="TR")b.onkeydown=function(e){if(e.key==="Enter")b.onclick();};
  })(els[i]);
  var q=view.querySelector("#conQ");
  if(q){
    q.oninput=function(){
      S.q=q.value;
      clearTimeout(S._qT);
      S._qT=setTimeout(function(){
        var pos=q.selectionStart; render();
        var nq=view.querySelector("#conQ");
        if(nq){nq.focus();try{nq.setSelectionRange(pos,pos);}catch(e){}}
      },220);
    };
  }
  els=view.querySelectorAll("[data-cur]");
  for(i=0;i<els.length;i++)(function(b){b.onclick=function(){S.cur=b.dataset.cur;render();};})(els[i]);
  els=view.querySelectorAll("[data-rest]");
  for(i=0;i<els.length;i++)(function(b){b.onclick=function(){S.rest=b.dataset.rest;render();};})(els[i]);
  els=view.querySelectorAll("[data-mode]");
  for(i=0;i<els.length;i++)(function(b){b.onclick=function(){S.conMode=b.dataset.mode;render();};})(els[i]);
  els=view.querySelectorAll("[data-act]");
  for(i=0;i<els.length;i++)(function(b){b.onclick=function(){act(b.dataset.act,b.dataset.id);};})(els[i]);
  var nr=document.getElementById("newReq"); if(nr)nr.onclick=function(){editRequest(null);};
  var nc=document.getElementById("newCon"); if(nc)nc.onclick=function(){editContract(null);};
}
document.getElementById("tabs").onclick=function(e){
  var b=e.target.closest("[data-tab]"); if(!b)return;
  S.tab=b.dataset.tab; S.sheetFor=null;
  var all=document.querySelectorAll("#tabs [data-tab]");
  for(var i=0;i<all.length;i++)all[i].setAttribute("aria-selected",String(all[i]===b));
  render();
};
document.getElementById("meBtn").onclick=pickRole;
document.getElementById("reloadBtn").onclick=function(){reload();};
document.getElementById("importBtn").onclick=openImport;
document.getElementById("exportBtn").onclick=exportRegistry;
document.getElementById("xlsxInput").onchange=function(e){
  handleFile(e.target.files&&e.target.files[0]);
  e.target.value="";
};

function findReq(id){for(var i=0;i<S.requests.length;i++)if(S.requests[i].id===id)return S.requests[i];return null;}
function findCon(id){for(var i=0;i<S.contracts.length;i++)if(S.contracts[i].id===id)return S.contracts[i];return null;}

function act(a,id){
  if(a==="retry"){reload();return;}
  if(a==="sheet"){S.sheetFor=id;render();return;}
  if(a==="closesheet"){S.sheetFor=null;render();return;}
  if(a==="print"){window.print();return;}
  var r=findReq(id);
  if(a==="submit"){doSubmit(r);return;}
  if(a==="approve"||a==="reject"){decide(r,a==="approve"?"approved":"rejected");return;}
  if(a==="pay"){markPaid(r);return;}
  if(a==="editreq"){editRequest(r);return;}
  if(a==="editcon"){editContract(findCon(id));return;}
}

function doSubmit(r){
  if(!(r.items||[]).length){toast("Маблағ холӣ · добавьте хотя бы одну позицию",true);return;}
  if(OFFLINE){r.submittedAt=new Date().toISOString();localSave();render();toast("Фиристода шуд");return;}
  send("submitRequest",{requestId:r.id},"Ба мувофиқа фиристода шуд");
}

function decide(r,decision){
  var s=signerBy(S.me);
  if(!s){toast("Вазифаи худро интихоб кунед · выберите свою должность",true);return;}
  modal({
    title:decision==="approved"?"Имзо гузоштан":"Рад кардан",
    body:'<div><label class="f">PIN</label><input id="mPin" type="password" inputmode="numeric" '+
      'autocomplete="off" value="'+esc(cachedPin(S.me))+'" placeholder="PIN аз лист «Имзокунандагон»"></div>'+
      '<div><label class="f">'+(decision==="approved"?"Шарҳ (ихтиёрӣ) · комментарий":"Сабаби рад · причина отказа")+
      '</label><textarea id="mComment"></textarea></div>'+
      '<p class="hint">Визу поставит: <b>'+esc(s.name)+"</b> — "+esc(s.posRu)+"</p>",
    okText:decision==="approved"?"Имзо гузоштан":"Рад кардан",
    danger:decision==="rejected",
    onOk:function(root){
      var pin=root.querySelector("#mPin").value.trim();
      var cm=root.querySelector("#mComment").value.trim();
      if(!pin){toast("PIN-ро ворид кунед",true);return false;}
      if(decision==="rejected"&&!cm){toast("Сабабро нависед · укажите причину",true);return false;}
      cachePin(S.me,pin);
      send("decide",{requestId:r.id,roleKey:S.me,pin:pin,decision:decision,comment:cm},
        decision==="approved"?"Имзо сабт шуд":"Дархост рад шуд");
    }
  });
}

function markPaid(r){
  var c=findCon(r.contractId);
  if(!c){toast("Шартнома пайваст нашуд",true);return;}
  modal({
    title:"Пардохтро қайд кардан",
    body:'<div class="fgrid">'+
      '<div><label class="f">Маблағ, '+esc(r.currency||"")+'</label><input id="pAmt" class="num" value="'+reqTotalCur(r)+'"></div>'+
      '<div><label class="f">Сана</label><input id="pDate" type="date" value="'+new Date().toISOString().slice(0,10)+'"></div>'+
      '<div><label class="f">Курс</label><input id="pRate" class="num" value="'+esc(r.rate||c.rate||"")+'"></div></div>'+
      '<div><label class="f">Шарҳ</label><input id="pNote" value="Пардохт аз рӯи дархости № '+esc(r.number||"")+'"></div>'+
      '<div><label class="f">PIN</label><input id="pPin" type="password" inputmode="numeric" autocomplete="off" value="'+
      esc(cachedPin(S.me))+'"></div>'+
      '<p class="hint">Бақия то пардохт: <b class="num">'+nf(balance(c),2)+" "+esc(c.currency||"")+"</b></p>",
    okText:"Қайд кардан",
    onOk:function(root){
      var amount=Number(root.querySelector("#pAmt").value)||0;
      var pin=root.querySelector("#pPin").value.trim();
      if(amount<=0){toast("Маблағ нодуруст",true);return false;}
      if(!pin){toast("PIN-ро ворид кунед",true);return false;}
      cachePin(S.me,pin);
      send("markPaid",{requestId:r.id,roleKey:S.me,pin:pin,amount:amount,
        date:root.querySelector("#pDate").value,rate:Number(root.querySelector("#pRate").value)||0,
        note:root.querySelector("#pNote").value.trim()},"Пардохт сабт шуд");
    }
  });
}

/* ── редакторы ── */
function editRequest(r){
  var isNew=!r;
  var d=r?JSON.parse(JSON.stringify(r)):{
    number:"",date:new Date().toISOString().slice(0,10),company:'ҶДММ «Мармарӣ»',
    contractId:S.contracts[0]?S.contracts[0].id:"",purpose:"",purposeShort:"",
    currency:"EUR",rate:10.81,items:[{name:"",unit:"евро",qty:1,price:0}],approvals:{}
  };
  var opts=S.contracts.map(function(c){
    return '<option value="'+esc(c.id)+'"'+(c.id===d.contractId?" selected":"")+">№ "+esc(c.no)+" — "+esc(c.supplier)+"</option>";
  }).join("");
  modal({
    title:isNew?"Дархости нав":"Таҳрири дархост",
    body:'<div class="fgrid">'+
      '<div><label class="f">Рақам · номер</label><input id="fNum" value="'+esc(d.number)+'" placeholder="01/01"></div>'+
      '<div><label class="f">Сана</label><input id="fDate" type="date" value="'+esc(d.date)+'"></div>'+
      '<div><label class="f">Ташкилот</label><input id="fComp" value="'+esc(d.company)+'"></div></div>'+
      '<div><label class="f">Шартнома</label><select id="fCon"><option value="">— бе шартнома —</option>'+opts+"</select></div>"+
      '<div class="fgrid"><div><label class="f">Асъор</label><input id="fCur" value="'+esc(d.currency)+'"></div>'+
      '<div><label class="f">Курс, сомонӣ</label><input id="fRate" class="num" value="'+esc(d.rate)+'"></div></div>'+
      '<div><label class="f">Асос · обоснование</label><textarea id="fPurp">'+esc(d.purpose)+"</textarea></div>"+
      '<div><label class="f">Матни бланк (пас аз «хоҳиш менамоем, ки …»)</label>'+
      '<input id="fShort" value="'+esc(d.purposeShort||"")+'" placeholder="барои хариди таҷҳизот …"></div>'+
      '<div><label class="f">Маблағ · позиции</label><div id="itemsBox"></div>'+
      '<button class="btn sm" id="addItem" style="margin-top:8px">+ Сатр</button></div>',
    okText:"Нигоҳ доштан",
    onOpen:function(root){
      var box=root.querySelector("#itemsBox");
      function draw(){
        box.innerHTML=d.items.map(function(it,i){
          return '<div class="fgrid" style="margin-bottom:8px;grid-template-columns:minmax(110px,2fr) 80px 70px 100px 34px">'+
            '<input data-i="'+i+'" data-k="name" value="'+esc(it.name)+'" placeholder="Номгӯй">'+
            '<input data-i="'+i+'" data-k="unit" value="'+esc(it.unit||"")+'" placeholder="Воҳ.">'+
            '<input data-i="'+i+'" data-k="qty" class="num" value="'+esc(it.qty)+'" placeholder="Миқдор">'+
            '<input data-i="'+i+'" data-k="price" class="num" value="'+esc(it.price)+'" placeholder="Нарх">'+
            '<button class="btn sm" data-del="'+i+'">✕</button></div>';
        }).join("");
        var ins=box.querySelectorAll("input");
        for(var j=0;j<ins.length;j++)(function(inp){
          inp.oninput=function(){
            var it=d.items[+inp.dataset.i],k=inp.dataset.k;
            it[k]=(k==="qty"||k==="price")?(Number(inp.value)||0):inp.value;
          };
        })(ins[j]);
        var dels=box.querySelectorAll("[data-del]");
        for(var m=0;m<dels.length;m++)(function(b){
          b.onclick=function(){d.items.splice(+b.dataset.del,1);draw();};
        })(dels[m]);
      }
      draw();
      root.querySelector("#addItem").onclick=function(){d.items.push({name:"",unit:"",qty:1,price:0});draw();};
    },
    onOk:function(root){
      d.number=root.querySelector("#fNum").value.trim();
      d.date=root.querySelector("#fDate").value;
      d.company=root.querySelector("#fComp").value.trim();
      d.contractId=root.querySelector("#fCon").value;
      d.currency=root.querySelector("#fCur").value.trim();
      d.rate=Number(root.querySelector("#fRate").value)||0;
      d.purpose=root.querySelector("#fPurp").value.trim();
      d.purposeShort=root.querySelector("#fShort").value.trim();
      if(!d.number){toast("Рақами дархостро нависед",true);return false;}
      if(OFFLINE){
        if(!d.id)d.id="r"+Date.now().toString(36);
        var i=S.requests.findIndex(function(x){return x.id===d.id;});
        if(i>=0)S.requests[i]=d; else S.requests.push(d);
        S.selReq=d.id;localSave();render();toast("Сабт шуд");return;
      }
      send("saveRequest",{request:d},isNew?"Дархост сохта шуд":"Сабт шуд").then(function(res){
        if(res&&res.id)S.selReq=res.id; render();
      });
    }
  });
}
function editContract(c){
  var isNew=!c;
  var d=c?JSON.parse(JSON.stringify(c)):{
    no:"",date:new Date().toISOString().slice(0,10),supplier:"",country:"",object:"Мармарӣ",
    product:"",currency:"EUR",rate:10.81,bank:"",invoice:0,gtd:"",file:"",payments:[]
  };
  var pl=(d.plan||[]).concat([{},{},{}]).slice(0,3);
  var num=function(id,label,val,hint){
    return '<div><label class="f">'+label+'</label><input id="'+id+'" class="num" inputmode="decimal" value="'+
      esc(val===0||val?val:"")+'">'+(hint?'<span class="fh">'+hint+"</span>":"")+"</div>";
  };
  modal({
    title:isNew?"Шартномаи нав":"Таҳрири шартнома",
    wide:true,
    body:'<fieldset class="grp"><legend>Шартнома</legend>'+
      '<div class="fgrid"><div><label class="f">№ шартнома *</label><input id="cNo" value="'+esc(d.no)+'"></div>'+
      '<div><label class="f">Сана</label><input id="cDate" type="date" value="'+esc(d.date)+'"></div>'+
      '<div><label class="f">Объекти дахлдор</label><input id="cObj" value="'+esc(d.object)+'"></div></div>'+
      '<div class="fgrid"><div><label class="f">Интиқолдиҳанда *</label><input id="cSup" value="'+esc(d.supplier)+'"></div>'+
      '<div><label class="f">Давлат</label><input id="cCty" value="'+esc(d.country)+'"></div></div>'+
      '<div><label class="f">Маҳсулот</label><input id="cProd" value="'+esc(d.product)+'"></div></fieldset>'+

      '<fieldset class="grp"><legend>Маблағ ва курс</legend><div class="fgrid">'+
      '<div><label class="f">Асъор</label><input id="cCur" value="'+esc(d.currency)+'" placeholder="EUR / USD / СОМ"></div>'+
      num("cRate","Курси ҳисобӣ",d.rate,"курс заявки")+
      num("cInv","Маблағи ИНВОЙС *",d.invoice)+
      "</div></fieldset>"+

      '<fieldset class="grp"><legend>Заявка ва тақсим</legend><div class="fgrid">'+
      num("cReqA","Заявка (валюта)",d.reqAmount)+num("cAllA","Ҷудошуда (валюта)",d.allocAmount)+
      "</div><div class=\"fgrid\">"+
      num("cR2R","Заявка-2: курс",d.req2Rate)+num("cR2A","Заявка-2 (валюта)",d.req2Amount)+
      num("cA2R","Ҷудошуда-2: курс",d.alloc2Rate,"по нему считается сомони")+
      num("cA2A","Ҷудошуда-2 (валюта)",d.alloc2Amount)+
      "</div>"+num("cAllT","Ҳамагӣ ҷудошуда",d.allocTotal)+"</fieldset>"+

      '<fieldset class="grp"><legend>Графики пардохт</legend>'+
      pl.map(function(x,i){
        return '<div class="fgrid" style="grid-template-columns:minmax(140px,2fr) minmax(90px,1fr)">'+
          '<div><label class="f">Транш '+(i+1)+' — шарт/сана</label><input id="cP'+i+'c" value="'+
          esc(x.cond||"")+'" placeholder="до 01.10.2026 / по готовности"></div>'+
          num("cP"+i+"s","Сумма",x.amount)+"</div>";
      }).join("")+
      '<p class="hint" id="planHint"></p></fieldset>'+

      '<fieldset class="grp"><legend>Иловагӣ</legend><div class="fgrid">'+
      '<div><label class="f">Бонки пардохткунанда</label><input id="cBank" value="'+esc(d.bank)+'"></div>'+
      '<div><label class="f">ГТД</label><input id="cGtd" value="'+esc(d.gtd)+'"></div>'+
      '<div><label class="f">Файли шартнома</label><input id="cFile" value="'+esc(d.file||"")+
      '" placeholder="контракт.pdf"></div></div></fieldset>',
    okText:"Нигоҳ доштан",
    onOpen:function(root){
      function recalc(){
        var inv=Number(root.querySelector("#cInv").value)||0;
        var s=0;
        for(var i=0;i<3;i++)s+=Number(root.querySelector("#cP"+i+"s").value)||0;
        var h=root.querySelector("#planHint");
        if(!s){h.textContent="График не задан — остаток считается от фактических оплат.";h.style.color="";return;}
        var diff=inv-s;
        h.innerHTML="По графику <b>"+nf(s,0)+"</b> из <b>"+nf(inv,0)+"</b>"+
          (Math.abs(diff)>0.5?" · расхождение <b>"+nf(diff,0)+"</b>":" · сходится");
        h.style.color=Math.abs(diff)>0.5?"var(--wait)":"var(--ok)";
      }
      ["cInv","cP0s","cP1s","cP2s"].forEach(function(id){
        root.querySelector("#"+id).oninput=recalc;
      });
      recalc();
    },
    onOk:function(root){
      d.no=root.querySelector("#cNo").value.trim();
      d.date=root.querySelector("#cDate").value;
      d.supplier=root.querySelector("#cSup").value.trim();
      d.country=root.querySelector("#cCty").value.trim();
      d.object=root.querySelector("#cObj").value.trim();
      d.product=root.querySelector("#cProd").value.trim();
      d.currency=root.querySelector("#cCur").value.trim();
      d.rate=Number(root.querySelector("#cRate").value)||0;
      d.invoice=Number(root.querySelector("#cInv").value)||0;
      d.bank=root.querySelector("#cBank").value.trim();
      d.gtd=root.querySelector("#cGtd").value.trim();
      d.file=root.querySelector("#cFile").value.trim();
      var g=function(id){return Number(root.querySelector("#"+id).value)||0;};
      d.reqAmount=g("cReqA"); d.allocAmount=g("cAllA");
      d.req2Rate=g("cR2R");  d.req2Amount=g("cR2A");
      d.alloc2Rate=g("cA2R"); d.alloc2Amount=g("cA2A");
      d.allocTotal=g("cAllT");
      d.plan=[0,1,2].map(function(i){
        return {cond:root.querySelector("#cP"+i+"c").value.trim(), amount:g("cP"+i+"s")};
      }).filter(function(x){return x.cond||x.amount;});
      if(!d.no){toast("Рақами шартномаро нависед",true);return false;}
      if(!d.supplier){toast("Интиқолдиҳандаро нависед · укажите поставщика",true);return false;}
      if(!(Number(d.invoice)||0)){toast("Маблағи ИНВОЙС холӣ · укажите сумму инвойса",true);return false;}
      if(OFFLINE){
        if(!d.id)d.id="c"+Date.now().toString(36);
        var i=S.contracts.findIndex(function(x){return x.id===d.id;});
        if(i>=0)S.contracts[i]=d; else S.contracts.push(d);
        S.selCon=d.id;localSave();render();toast("Сабт шуд");return;
      }
      send("saveContract",{contract:d},isNew?"Шартнома сохта шуд":"Сабт шуд").then(function(res){
        if(res&&res.id)S.selCon=res.id; render();
      });
    }
  });
}

/* ── выбор роли ── */
function pickRole(){
  var chosen=S.me;
  modal({
    title:"Шумо кӣ ҳастед? · Ваша должность",
    body:'<p class="hint">Виза ставится от имени выбранной должности и подтверждается PIN-кодом. '+
      "Выбор сохраняется в этом браузере.</p>"+
      '<div class="pick" id="roleBox">'+S.signers.map(function(p){
        return '<button type="button" data-role="'+esc(p.key)+'" aria-pressed="'+(S.me===p.key)+'">'+
          '<span class="p1">'+esc(p.name)+'</span><span class="p2">'+esc(p.posTj)+"</span></button>";
      }).join("")+
      '<button type="button" data-role="observer" aria-pressed="'+(S.me==="observer")+'">'+
      '<span class="p1">Нозир</span><span class="p2">Наблюдатель — только просмотр</span></button></div>',
    okText:"Интихоб",
    onOpen:function(root){
      var bs=root.querySelectorAll("[data-role]");
      for(var i=0;i<bs.length;i++)(function(b){
        b.onclick=function(){
          chosen=b.dataset.role;
          for(var j=0;j<bs.length;j++)bs[j].setAttribute("aria-pressed",String(bs[j]===b));
        };
      })(bs[i]);
    },
    onOk:function(){
      if(!chosen){toast("Вазифаро интихоб кунед",true);return false;}
      S.me=chosen; lsSet("avesto.role",chosen); render();
    }
  });
}

/* ── диалог и уведомление ── */
function modal(opt){
  var host=document.getElementById("modalHost");
  host.innerHTML='<div class="scrim"><div class="modal'+(opt.wide?" wide":"")+'" role="dialog" aria-modal="true">'+
    '<div class="mh"><h3>'+esc(opt.title)+'</h3></div><div class="mb">'+opt.body+"</div>"+
    '<div class="mf"><button class="btn" data-x="c">Бекор</button>'+
    '<button class="btn '+(opt.danger?"dang":"pri")+'" data-x="o">'+esc(opt.okText||"ОК")+"</button></div></div></div>";
  var root=host.querySelector(".modal");
  function close(){host.innerHTML="";document.removeEventListener("keydown",onKey);}
  function onKey(e){if(e.key==="Escape")close();}
  host.querySelector('[data-x="c"]').onclick=close;
  host.querySelector('[data-x="o"]').onclick=function(){
    if(opt.onOk&&opt.onOk(root)===false)return;
    close();
  };
  host.querySelector(".scrim").onclick=function(e){if(e.target===e.currentTarget)close();};
  document.addEventListener("keydown",onKey);
  if(opt.onOpen)opt.onOpen(root);
  var f=root.querySelector("input,textarea,select,[data-role]"); if(f)f.focus();
}
var toastT=null;
function toast(msg,bad){
  var h=document.getElementById("toastHost");
  h.innerHTML='<div class="toast'+(bad?" bad":"")+'">'+esc(msg)+"</div>";
  clearTimeout(toastT);
  toastT=setTimeout(function(){h.innerHTML="";},bad?6000:3600);
}

/* ═══ импорт реестра из Excel ═══
   Структура файла «Мархилаи-2»: строка заголовков с «Номи Ташкилоти
   интиколдиханда», под ней два ряда подзаголовков, затем данные.
   Колонки ищем по названиям, а не по буквам — в разных версиях файла
   они сдвинуты. Первый «Курс» — курс заявки, последний — курс оплаты,
   по которому реально прошёл платёж. */
function cellText(v){ return v==null?"":String(v).replace(/\s+/g," ").trim(); }
function cellNum(v){
  if(v==null||v==="")return 0;
  if(typeof v==="number")return v;
  var s=String(v).replace(/\s| /g,"").replace(",",".").replace(/[^\d.\-]/g,"");
  var n=parseFloat(s);
  return isNaN(n)?0:n;
}
function excelDate(v){
  if(v==null||v==="")return "";
  if(v instanceof Date)return v.toISOString().slice(0,10);
  if(typeof v==="number"&&v>20000&&v<80000){
    var d=new Date(Date.UTC(1899,11,30+Math.floor(v)));
    return d.toISOString().slice(0,10);
  }
  var s=cellText(v);
  var m=s.match(/(\d{2})[.\/](\d{2})[.\/](\d{4})/);
  if(m)return m[3]+"-"+m[2]+"-"+m[1];
  if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);
  return "";
}
function findCol(hdr,needle){
  for(var i=0;i<hdr.length;i++){
    var t=cellText(hdr[i]).toLowerCase();
    if(t&&t.indexOf(needle.toLowerCase())>-1)return i;
  }
  return -1;
}
function allCols(hdr,needle){
  var out=[];
  for(var i=0;i<hdr.length;i++){
    var t=cellText(hdr[i]).toLowerCase();
    if(t&&t.indexOf(needle.toLowerCase())>-1)out.push(i);
  }
  return out;
}
function parseRegistry(wb){
  var best=null;
  wb.SheetNames.forEach(function(name){
    var rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:true,defval:null});
    var hdr=-1;
    for(var i=0;i<Math.min(rows.length,15);i++){
      if(findCol(rows[i]||[],"номи ташкилоти")>-1){hdr=i;break;}
    }
    if(hdr<0)return;
    var r=readSheet(rows,hdr,name);
    if(r&&(!best||r.contracts.length>best.contracts.length))best=r;
  });
  return best;
}
/* Шапка реестра трёхэтажная: строка групп («ПО ДОГОВОРАМ», «ЗАЯВКА-2»,
   «Фактическая Оплата»), под ней подгруппы («Заявка», «Выделено»,
   «Оплата-1»), под ними — что лежит в колонке («Курс», «Сумма (Валюта)»,
   «Сумма (СОМ)», «Дата»). Ищем по этим надписям, а не по буквам. */
function groupRange(H,name){
  var from=-1;
  for(var i=0;i<H.length;i++){
    var t=cellText(H[i]).toLowerCase();
    if(t&&t.indexOf(name.toLowerCase())>-1){from=i;break;}
  }
  if(from<0)return null;
  var to=H.length;
  for(var j=from+1;j<H.length;j++){ if(cellText(H[j])){to=j;break;} }
  return [from,to];
}
function subRange(sub,range,name){
  if(!range)return null;
  var from=-1;
  for(var i=range[0];i<range[1];i++){
    var t=cellText(sub[i]).toLowerCase();
    if(t&&t.indexOf(name.toLowerCase())>-1){from=i;break;}
  }
  if(from<0)return null;
  var to=range[1];
  for(var j=from+1;j<range[1];j++){ if(cellText(sub[j])){to=j;break;} }
  return [from,to];
}
function colIn(sub,range,name){
  if(!range)return -1;
  for(var i=range[0];i<range[1];i++){
    var t=cellText(sub[i]).toLowerCase();
    if(t&&t.indexOf(name.toLowerCase())>-1)return i;
  }
  return -1;
}
function pairsIn(sub,range){
  var out=[],pending=-1;
  if(!range)return out;
  for(var k=range[0];k<range[1];k++){
    var s=cellText(sub[k]).toLowerCase();
    if(s.indexOf("дата")>-1)pending=k;
    else if(s.indexOf("сумма")>-1&&pending>-1){out.push([pending,k]);pending=-1;}
  }
  return out;
}

function readSheet(rows,hdr,sheetName){
  var H=rows[hdr]||[], sub=rows[hdr+2]||[], mid=rows[hdr+1]||[];
  var ci={
    date:findCol(H,"сана"), supplier:findCol(H,"номи ташкилоти"), country:findCol(H,"давлати"),
    no:findCol(H,"шартнома"), object:findCol(H,"объекти"), product:findCol(H,"номи махсулоти"),
    cur:findCol(H,"асъор"), invoice:findCol(H,"инвойс")
  };
  if(ci.supplier<0||ci.invoice<0)return null;
  /* «Курс» встречается и в шапке, и в ряду подзаголовков: первый по счёту —
     курс заявки, последний — курс, по которому деньги реально выделены. */
  var rates=allCols(H,"курс").concat(allCols(sub,"курс")).sort(function(a,b){return a-b;});
  var rateCol=rates.length?rates[0]:-1;
  var payRateCol=rates.length>1?rates[rates.length-1]:rateCol;

  var gPay=groupRange(H,"фактическая оплата")||groupRange(H,"факт оплата");
  var pairs=pairsIn(sub,gPay);

  /* остальные группы реестра — нужны, чтобы выгрузка вернула такой же файл */
  var gDog=groupRange(H,"по договорам")||groupRange(H,"заявка (по договорам)");
  var gZ2=groupRange(H,"заявка-2");
  var gPlan=groupRange(H,"к оплате");
  var dogReq=subRange(mid,gDog,"заявка"), dogAlloc=subRange(mid,gDog,"выделено");
  var z2Req=subRange(mid,gZ2,"заявка"), z2Alloc=subRange(mid,gZ2,"выделено");
  var planPairs=pairsIn(sub,gPlan);
  var colAllocTotal=findCol(H,"всего выделено");
  var cDogReqV=colIn(sub,dogReq,"валюта"), cDogReqS=colIn(sub,dogReq,"сом");
  var cDogAllV=colIn(sub,dogAlloc,"валюта"), cDogAllS=colIn(sub,dogAlloc,"сом");
  var cZ2ReqR=colIn(sub,z2Req,"курс"), cZ2ReqV=colIn(sub,z2Req,"валюта"), cZ2ReqS=colIn(sub,z2Req,"сом");
  var cZ2AllR=colIn(sub,z2Alloc,"курс"), cZ2AllV=colIn(sub,z2Alloc,"валюта"), cZ2AllS=colIn(sub,z2Alloc,"сом");

  /* Начало данных ищем, а не отсчитываем: в одних версиях файла под шапкой
     два ряда подзаголовков, в других строки идут сразу. */
  var start=-1;
  for(var si=hdr+1;si<Math.min(rows.length,hdr+7);si++){
    var rr=rows[si]||[];
    if(cellText(rr[ci.supplier])&&cellNum(rr[ci.invoice])>0){start=si;break;}
  }
  if(start<0)start=hdr+3;

  var contracts=[],payments=[],skipped=0;
  for(var rI=start;rI<rows.length;rI++){
    var row=rows[rI]||[];
    var sup=cellText(row[ci.supplier]);
    var first=cellText(row[0]).toLowerCase();
    if(first.indexOf("итог")>-1)break;
    var inv=cellNum(row[ci.invoice]);
    var line=row.slice(0,10).map(cellText).join(" ").toLowerCase();
    if(line.indexOf("итого")>-1)continue;           // строки итогов из нашей же выгрузки
    if(!sup){ if(inv)skipped++; continue; }
    if(/^(свод|заявка|выделено|остаток|основные|строител)/i.test(sup))continue;
    var num=cellText(row[0]).replace(/\D/g,"");
    var cid="c-"+(num?("00"+num).slice(-2):String(contracts.length+1));
    var rate=rateCol>-1?cellNum(row[rateCol]):1; if(!rate)rate=1;
    var rp=payRateCol>-1?cellNum(row[payRateCol]):rate; if(!rp)rp=rate;
    var at=function(col){return col>-1?cellNum(row[col]):0;};
    contracts.push({
      id:cid, no:cellText(row[ci.no])||"—", date:excelDate(row[ci.date]),
      supplier:sup, country:cellText(row[ci.country]), object:cellText(row[ci.object]),
      product:cellText(row[ci.product]), currency:cellText(row[ci.cur])||"—",
      rate:rate, bank:"", invoice:inv, gtd:"", file:"", payments:[],
      reqAmount:at(cDogReqV), reqTjs:at(cDogReqS),
      allocAmount:at(cDogAllV), allocTjs:at(cDogAllS),
      req2Rate:at(cZ2ReqR), req2Amount:at(cZ2ReqV), req2Tjs:at(cZ2ReqS),
      alloc2Rate:at(cZ2AllR)||rp, alloc2Amount:at(cZ2AllV), alloc2Tjs:at(cZ2AllS),
      allocTotal:at(colAllocTotal),
      plan:planPairs.slice(0,3).map(function(pp){
        return {cond:cellText(row[pp[0]])||excelDate(row[pp[0]]), amount:cellNum(row[pp[1]])};
      })
    });
    var c=contracts[contracts.length-1], n=0;
    pairs.forEach(function(p){
      var am=cellNum(row[p[1]]);
      if(!am)return;
      n++;
      c.payments.push({amount:am, date:excelDate(row[p[0]]), rate:rp,
                       tjs:Math.round(am*rp*100)/100, note:"Пардохти "+n});
    });
    payments.push.apply(payments,c.payments);
  }
  return {sheet:sheetName, contracts:contracts, payments:payments, skipped:skipped};
}

function openImport(){
  document.getElementById("xlsxInput").click();
}
function handleFile(file){
  if(!file)return;
  if(typeof XLSX==="undefined"){toast("Модули хониши Excel бор нашуд · не загрузился модуль чтения Excel",true);return;}
  var fr=new FileReader();
  fr.onerror=function(){toast("Файлро хондан нашуд · не удалось прочитать файл",true);};
  fr.onload=function(e){
    var parsed;
    try{
      /* без cellDates: иначе библиотека отдаёт даты объектом Date в местном
         поясе и число месяца уезжает на день. Серийный номер переводим сами. */
      var wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
      parsed=parseRegistry(wb);
    }catch(err){ toast("Файл хонда нашуд: "+(err.message||err),true); return; }
    if(!parsed||!parsed.contracts.length){
      toast("Дар файл реестр ёфт нашуд · в файле не найден реестр контрактов",true);
      return;
    }
    previewImport(parsed,file.name);
  };
  fr.readAsArrayBuffer(file);
}
function previewImport(p,fname){
  var byCur={};
  p.contracts.forEach(function(c){
    var k=c.currency||"—";
    byCur[k]=byCur[k]||{inv:0,paid:0,n:0};
    byCur[k].inv+=Number(c.invoice)||0;
    byCur[k].paid+=paidSum(c);
    byCur[k].n++;
  });
  var rowsHtml=Object.keys(byCur).map(function(k){
    var g=byCur[k];
    return "<tr><td>"+esc(k)+'</td><td class="r num">'+g.n+'</td><td class="r num">'+nf(g.inv,2)+
      '</td><td class="r num">'+nf(g.paid,2)+'</td><td class="r num">'+nf(g.inv-g.paid,2)+"</td></tr>";
  }).join("");
  var over=p.contracts.filter(function(c){return paidSum(c)>(Number(c.invoice)||0)+0.5;});
  modal({
    title:"Импорт реестра",
    body:'<p class="hint">Файл: <b>'+esc(fname)+"</b> · лист «"+esc(p.sheet)+"»</p>"+
      '<div class="tw"><table class="items"><thead><tr><th>Асъор</th><th class="r">Шартнома</th>'+
      '<th class="r">Инвойс</th><th class="r">Пардохтшуда</th><th class="r">Бақия</th></tr></thead>'+
      "<tbody>"+rowsHtml+"</tbody></table></div>"+
      '<p class="hint">Распознано: <b>'+p.contracts.length+"</b> контрактов и <b>"+
      p.payments.length+"</b> фактических платежей."+
      (p.skipped?" Пропущено строк без поставщика: "+p.skipped+".":"")+"</p>"+
      (over.length?'<div class="notice" style="margin:0"><span>⚠</span><span>Переплата по '+over.length+
        " контракту(ам): "+esc(over.map(function(c){return c.supplier;}).slice(0,3).join(", "))+
        (over.length>3?" и др.":"")+"</span></div>":"")+
      '<div><label class="f">Что сделать с текущими данными</label>'+
      '<select id="impMode"><option value="replace">Заменить реестр целиком</option>'+
      '<option value="merge">Добавить и обновить по номеру контракта</option></select></div>'+
      (OFFLINE?'<p class="hint">База не подключена — данные сохранятся в этом браузере.</p>'
              :'<p class="hint">Данные будут записаны в таблицу Google Sheets.</p>'),
    okText:"Импорт",
    onOk:function(root){
      var mode=root.querySelector("#impMode").value;
      applyImport(p,mode);
    }
  });
}
function applyImport(p,mode){
  var next;
  if(mode==="replace")next=p.contracts.slice();
  else{
    next=S.contracts.slice();
    p.contracts.forEach(function(c){
      var i=-1;
      for(var j=0;j<next.length;j++)
        if(next[j].no&&c.no&&next[j].no===c.no&&next[j].supplier===c.supplier){i=j;break;}
      if(i>=0)next[i]=c; else next.push(c);
    });
  }
  if(!API){
    S.contracts=next; localSave(); render();
    toast("Импорт шуд: "+p.contracts.length+" шартнома"+(READONLY?" (танҳо дар ин браузер)":""));
    return;
  }
  send("importRegistry",{contracts:next},"Импорт шуд: "+p.contracts.length+" шартнома");
}

/* ═══ выгрузка реестра в Excel — та же шапка, что в файле финотдела ═══ */
var XCOL = 41; // A..AO
function exportRegistry(){
  if(typeof XLSX==="undefined"){toast("Модули Excel бор нашуд",true);return;}
  var list=S.contracts.slice().sort(function(a,b){
    return (parseInt(String(a.id).replace(/\D/g,""),10)||0)-(parseInt(String(b.id).replace(/\D/g,""),10)||0);
  });
  var A=function(){var r=[];for(var i=0;i<XCOL;i++)r.push(null);return r;};
  var rows=[A(),A(),A(),A(),A(),A()];
  rows[1][0]="Мархилаи -2 - ЗАЯВКИ на выделение ДС на "+dRu(new Date().toISOString());
  var h4=rows[3],h5=rows[4],h6=rows[5];
  ["№","Сана","Номи Ташкилоти интиколдиханда","Давлати интиколдиханда","№ шартнома",
   "Объекти дахлдор","Номи махсулоти воридшаванда","Асъор"].forEach(function(t,i){h4[i]=t;});
  h4[9]="Маблаги ИНВОЙС"; h4[10]="Курс";
  h4[11]="ПО ДОГОВОРАМ"; h4[17]="ЗАЯВКА-2"; h4[25]="ВСЕГО ВЫДЕЛЕНО";
  h4[26]="К оплате - согласно инвойса"; h4[33]="Фактическая Оплата";
  h4[39]="ВСЕГО ОПЛАТА"; h4[40]="Фактический Остаток к выплате от суммы контракта";
  h5[11]="Заявка"; h5[13]="Выделено"; h5[15]="Остаток";
  h5[17]="Заявка"; h5[20]="Выделено"; h5[23]="Остаток";
  h5[26]="К оплате-1транш"; h5[28]="К оплате-2транш"; h5[30]="К оплате-3транш"; h5[32]="ИТОГО";
  h5[33]="Оплата-1"; h5[35]="Оплата-2"; h5[37]="Оплата-3";
  var V="Сумма (Валюта)", C="Сумма (СОМ)";
  [11,13,15].forEach(function(i){h6[i]=V;h6[i+1]=C;});
  h6[17]="Курс"; h6[18]=V; h6[19]=C;
  h6[20]="Курс"; h6[21]=V; h6[22]=C;
  h6[23]=V; h6[24]=C;
  [26,28,30].forEach(function(i){h6[i]="Дата";h6[i+1]="Сумма";});
  [33,35,37].forEach(function(i){h6[i]="Дата";h6[i+1]="Сумма";});

  var tot={};
  list.forEach(function(c,idx){
    var r=A(), n=function(x){return Number(x)||0;};
    var inv=n(c.invoice), paid=paidSum(c), plan=(c.plan||[]);
    var planSum=plan.reduce(function(s,x){return s+n(x.amount);},0);
    /* даты пишем как «дд.мм.гггг»: ISO-строку Excel распознаёт как дату
       и сдвигает её на день по часовому поясу */
    var dOut=function(s){ var t=dRu(s); return t==="—"?(s||""):t; };
    r[0]=idx+1; r[1]=dOut(c.date); r[2]=c.supplier||""; r[3]=c.country||"";
    r[4]=c.no||""; r[5]=c.object||""; r[6]=c.product||""; r[7]=c.currency||"";
    r[9]=inv; r[10]=n(c.rate);
    r[11]=n(c.reqAmount); r[12]=n(c.reqTjs);
    r[13]=n(c.allocAmount); r[14]=n(c.allocTjs);
    r[15]=n(c.reqAmount)-n(c.allocAmount); r[16]=n(c.reqTjs)-n(c.allocTjs);
    r[17]=n(c.req2Rate); r[18]=n(c.req2Amount); r[19]=n(c.req2Tjs);
    r[20]=n(c.alloc2Rate); r[21]=n(c.alloc2Amount); r[22]=n(c.alloc2Tjs);
    r[23]=n(c.req2Amount)-n(c.alloc2Amount); r[24]=n(c.req2Tjs)-n(c.alloc2Tjs);
    r[25]=n(c.allocTotal);
    [0,1,2].forEach(function(k){
      if(plan[k]){ r[26+k*2]=plan[k].cond||""; r[27+k*2]=n(plan[k].amount); }
    });
    r[32]=inv-planSum;
    (c.payments||[]).slice(0,3).forEach(function(p,k){
      r[33+k*2]=dOut(p.date); r[34+k*2]=n(p.amount);
    });
    r[39]=paid; r[40]=inv-paid;
    rows.push(r);
    var cur=c.currency||"—";
    tot[cur]=tot[cur]||{inv:0,paid:0,rest:0};
    tot[cur].inv+=inv; tot[cur].paid+=paid; tot[cur].rest+=inv-paid;
  });

  rows.push(A());
  Object.keys(tot).forEach(function(cur){
    var r=A();
    r[6]="Итого:"; r[7]=cur; r[9]=tot[cur].inv; r[39]=tot[cur].paid; r[40]=tot[cur].rest;
    rows.push(r);
  });
  var rt=A();
  rt[6]="Итого в сомонӣ:";
  rt[39]=list.reduce(function(s,c){return s+paidTjs(c);},0);
  rows.push(rt);

  var ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"]=[
    {s:{r:1,c:0},e:{r:1,c:25}},
    {s:{r:3,c:11},e:{r:3,c:16}}, {s:{r:3,c:17},e:{r:3,c:24}},
    {s:{r:3,c:26},e:{r:3,c:32}}, {s:{r:3,c:33},e:{r:3,c:38}},
    {s:{r:4,c:11},e:{r:4,c:12}}, {s:{r:4,c:13},e:{r:4,c:14}}, {s:{r:4,c:15},e:{r:4,c:16}},
    {s:{r:4,c:17},e:{r:4,c:19}}, {s:{r:4,c:20},e:{r:4,c:22}}, {s:{r:4,c:23},e:{r:4,c:24}},
    {s:{r:4,c:26},e:{r:4,c:27}}, {s:{r:4,c:28},e:{r:4,c:29}}, {s:{r:4,c:30},e:{r:4,c:31}},
    {s:{r:4,c:33},e:{r:4,c:34}}, {s:{r:4,c:35},e:{r:4,c:36}}, {s:{r:4,c:37},e:{r:4,c:38}}
  ];
  [0,1,2,3,4,5,6,7,9,10,25,32,39,40].forEach(function(c){
    ws["!merges"].push({s:{r:3,c:c},e:{r:5,c:c}});
  });
  ws["!cols"]=[{wch:5},{wch:13},{wch:38},{wch:16},{wch:20},{wch:12},{wch:34},{wch:7},{wch:5},
               {wch:14},{wch:9}].concat(new Array(30).fill({wch:13}));
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"общ");
  var out=XLSX.write(wb,{bookType:"xlsx",type:"array"});
  var blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  var url=URL.createObjectURL(blob), a=document.createElement("a");
  a.href=url; a.download="Мархилаи-2 — реестр "+new Date().toISOString().slice(0,10)+".xlsx";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(url);},1500);
  toast("Реестр выгружен: "+list.length+" шартнома");
}

/* ═══ чтение опубликованной таблицы (без Apps Script) ═══ */
function csvUrl(id,sheet){
  return "https://docs.google.com/spreadsheets/d/"+encodeURIComponent(id)+
         "/gviz/tq?tqx=out:csv&sheet="+encodeURIComponent(sheet);
}
function parseCsv(text){
  var rows=[],row=[],cur="",q=false;
  for(var i=0;i<text.length;i++){
    var ch=text[i];
    if(q){
      if(ch==='"'){ if(text[i+1]==='"'){cur+='"';i++;} else q=false; }
      else cur+=ch;
    }else if(ch==='"')q=true;
    else if(ch===","){row.push(cur);cur="";}
    else if(ch==="\n"){row.push(cur);rows.push(row);row=[];cur="";}
    else if(ch!=="\r")cur+=ch;
  }
  if(cur!==""||row.length){row.push(cur);rows.push(row);}
  return rows;
}
function objects(rows){
  if(!rows.length)return [];
  var h=rows[0].map(function(x){return String(x).trim();});
  return rows.slice(1).filter(function(r){return r.join("").trim()!=="";}).map(function(r){
    var o={};
    h.forEach(function(k,i){o[k]=r[i]!==undefined?r[i]:"";});
    return o;
  });
}
function loadFromSheet(id){
  var names=["Шартномаҳо","Пардохтҳо","Дархостҳо","Сатрҳо","Имзоҳо","Имзокунандагон"];
  return Promise.all(names.map(function(n){
    return fetch(csvUrl(id,n)).then(function(r){
      if(!r.ok)throw new Error("лист «"+n+"» недоступен ("+r.status+")");
      return r.text();
    }).then(function(t){return objects(parseCsv(t));});
  })).then(function(res){
    var contracts=res[0],pays=res[1],reqs=res[2],items=res[3],apprs=res[4],signers=res[5];
    contracts.forEach(function(c){
      c.rate=Number(c.rate)||0; c.invoice=Number(c.invoice)||0;
      c.payments=pays.filter(function(p){return p.contractId===c.id;}).map(function(p){
        return {amount:Number(p.amount)||0,date:p.date,rate:Number(p.rate)||0,
                tjs:Number(p.tjs)||0,note:p.note,requestId:p.requestId};
      });
    });
    reqs.forEach(function(r){
      r.rate=Number(r.rate)||0;
      r.items=items.filter(function(i){return i.requestId===r.id;}).map(function(i){
        return {name:i.name,unit:i.unit,qty:Number(i.qty)||0,price:Number(i.price)||0};
      });
      r.approvals={};
      apprs.filter(function(a){return a.requestId===r.id;}).forEach(function(a){
        r.approvals[a.roleKey]={decision:a.decision,at:a.at,byName:a.byName,comment:a.comment};
      });
    });
    return {contracts:contracts,requests:reqs,
            signers:signers.filter(function(s){return s.key;}).sort(function(a,b){
              return (Number(a.order)||0)-(Number(b.order)||0);})};
  });
}

/* ═══ запуск ═══ */
S.me=lsGet("avesto.role")||null;
var foot=document.getElementById("footText");
if(foot&&CFG.FOOTER)foot.textContent=CFG.FOOTER;
render();
reload();
})();
