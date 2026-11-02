/* Мувофиқасозии дархостҳо — клиентская часть.
   Данные: Google Sheets через веб-приложение Apps Script (см. config.js).
   Разработано Отделом автоматизации холдинга ООО «Авесто Групп» · ISJ */
(function () {
"use strict";

var CFG = window.AVESTO_CONFIG || {};
var API = String(CFG.API_URL || "").trim();
var OFFLINE = !API;

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
  return apiGet().then(function(d){applyState(d);render();})
    .catch(function(e){S.loaded=true;S.error=e.message||String(e);render();});
}
function send(action,payload,okMsg){
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
  view.innerHTML=(OFFLINE?offlineNotice():"")+(S.error?errorNotice():"")+
    (S.tab==="requests"?renderRequests():renderContracts());
  wire();
  var btn=document.getElementById("meBtn");
  var s=S.me&&signerBy(S.me);
  btn.textContent=s?s.name:(S.me==="observer"?"Нозир":"Кто вы?");
  btn.title=s?s.posRu:"Выбрать роль";
}
function offlineNotice(){
  return '<div class="notice"><span>⚠</span><span><b>База не подключена.</b> '+
    'Записи сохраняются только в этом браузере и другим подписантам не видны. '+
    'Укажите адрес веб-приложения в <code>assets/config.js</code> — инструкция в README репозитория.</span></div>';
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
function renderContracts(){
  var list=S.contracts.slice().sort(function(a,b){return String(b.date||"").localeCompare(String(a.date||""));});
  var sel=null;
  for(var i=0;i<list.length;i++) if(list[i].id===S.selCon)sel=list[i];
  if(!sel)sel=list[0]||null;
  if(sel)S.selCon=sel.id;
  var totalTjs=list.reduce(function(s,c){return s+paidTjs(c);},0);
  return '<div class="cols"><div class="panel">'+
    '<div class="head"><h2>Шартномаҳо</h2><span class="count">'+list.length+"</span></div>"+
    (list.length?'<ul class="list">'+list.map(function(c){return conRow(c,sel);}).join("")+"</ul>":
      '<div class="empty"><h3>Шартнома нест</h3><p>'+
      (S.loaded?"Реестр контрактов пуст.":'<span class="spin"></span> Маълумот бор мешавад…')+"</p></div>")+
    (list.length?'<div class="foot" style="font-size:13px;color:var(--ink-soft);display:flex;gap:8px">'+
      '<span>Ҳамагӣ пардохт</span><b class="num" style="margin-left:auto">'+nf(totalTjs,2)+" сом.</b></div>":"")+
    '<div class="foot"><button class="btn sm" id="newCon">+ Шартномаи нав</button></div></div>'+
    (sel?conDoc(sel):'<div class="panel"><div class="empty"><h3>Шартнома интихоб нашуд</h3>'+
      "<p>Выберите контракт — откроется график платежей и остаток к оплате.</p></div></div>")+"</div>";
}
function conRow(c,sel){
  var inv=Number(c.invoice)||0,pd=paidSum(c),pct=inv?Math.min(100,pd/inv*100):0;
  return '<li><button class="row" data-con="'+esc(c.id)+'" aria-current="'+(sel&&sel.id===c.id)+'">'+
    '<span class="l1"><b>'+esc(c.no||"—")+"</b>"+
    (balance(c)<=0.0001&&inv>0?'<span class="badge approved">Пӯшида</span>':"")+"<time>"+dRu(c.date)+"</time></span>"+
    '<span class="l2">'+esc(c.supplier||"—")+" · "+esc(c.product||"")+"</span>"+
    '<span class="l3" style="display:block"><span class="bar"><i style="width:'+pct.toFixed(1)+'%"></i></span>'+
    '<span style="display:flex;gap:8px;margin-top:5px;font-size:12.5px;color:var(--ink-soft)">'+
    '<span class="num">'+nf(pd,0)+'</span><span>/</span><span class="num">'+nf(inv,0)+" "+esc(c.currency||"")+"</span>"+
    '<span class="num" style="margin-left:auto">бақия '+nf(balance(c),0)+"</span></span></span></button></li>";
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
  for(i=0;i<els.length;i++)(function(b){b.onclick=function(){S.selCon=b.dataset.con;render();};})(els[i]);
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
  modal({
    title:isNew?"Шартномаи нав":"Таҳрири шартнома",
    body:'<div class="fgrid"><div><label class="f">№ шартнома</label><input id="cNo" value="'+esc(d.no)+'"></div>'+
      '<div><label class="f">Сана</label><input id="cDate" type="date" value="'+esc(d.date)+'"></div></div>'+
      '<div class="fgrid"><div><label class="f">Интиқолдиҳанда</label><input id="cSup" value="'+esc(d.supplier)+'"></div>'+
      '<div><label class="f">Давлат</label><input id="cCty" value="'+esc(d.country)+'"></div></div>'+
      '<div class="fgrid"><div><label class="f">Объекти дахлдор</label><input id="cObj" value="'+esc(d.object)+'"></div>'+
      '<div><label class="f">Маҳсулот</label><input id="cProd" value="'+esc(d.product)+'"></div></div>'+
      '<div class="fgrid"><div><label class="f">Асъор</label><input id="cCur" value="'+esc(d.currency)+'"></div>'+
      '<div><label class="f">Курс</label><input id="cRate" class="num" value="'+esc(d.rate)+'"></div>'+
      '<div><label class="f">Маблағи ИНВОЙС</label><input id="cInv" class="num" value="'+esc(d.invoice)+'"></div></div>'+
      '<div class="fgrid"><div><label class="f">Бонки пардохткунанда</label><input id="cBank" value="'+esc(d.bank)+'"></div>'+
      '<div><label class="f">ГТД</label><input id="cGtd" value="'+esc(d.gtd)+'"></div>'+
      '<div><label class="f">Файли шартнома</label><input id="cFile" value="'+esc(d.file||"")+'" placeholder="контракт.pdf"></div></div>',
    okText:"Нигоҳ доштан",
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
      if(!d.no){toast("Рақами шартномаро нависед",true);return false;}
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
  host.innerHTML='<div class="scrim"><div class="modal" role="dialog" aria-modal="true">'+
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

/* ═══ запуск ═══ */
S.me=lsGet("avesto.role")||null;
var foot=document.getElementById("footText");
if(foot&&CFG.FOOTER)foot.textContent=CFG.FOOTER;
render();
reload();
})();
