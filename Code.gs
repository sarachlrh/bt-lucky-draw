// 블루타이거 300 번호 뽑기 - Google Apps Script (스프레드시트에 연결해 웹 앱으로 배포)
// 번호판 상태(경품 위치·공개된 번호)는 스크립트 속성에, 뽑기 기록은 '기록' 탭에 저장.
// 링크를 가진 진행자 누구나 뽑기·리셋·경품 설정·당첨 번호 보기를 할 수 있음.
var N = 300;
var DEFAULT_PRIZES = [
  {tier:'1등', name:'블루타이거 수업 100% 무료 수강권', qty:1},
  {tier:'2등', name:'네이버페이 1만원권', qty:3},
  {tier:'3등', name:'스타벅스 아메리카노', qty:5},
  {tier:'장려상', name:'GS25 편의점 상품권 3천원권', qty:10}
];
var FILLER = {tier:'참가상', name:'블루타이거 시크릿 자료집'};

function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function fresh_(prizes, prev){
  var pool=[];
  prizes.forEach(function(p,i){for(var k=0;k<p.qty;k++)pool.push(i)});
  while(pool.length<N)pool.push(-1);
  for(var i=pool.length-1;i>0;i--){var j=Math.floor(Math.random()*(i+1)),t=pool[i];pool[i]=pool[j];pool[j]=t}
  return {sid:prev?prev.sid:Utilities.getUuid(), ver:prev?prev.ver+1:1, round:prev?prev.round+1:1, prizes:prizes, slots:pool, opened:[]};
}
function load_(){
  var c=CacheService.getScriptCache().get('s');if(c)return JSON.parse(c);
  var p=PropertiesService.getScriptProperties().getProperty('s');if(!p)return null;
  CacheService.getScriptCache().put('s',p,21600);return JSON.parse(p);
}
function save_(s){var j=JSON.stringify(s);PropertiesService.getScriptProperties().setProperty('s',j);CacheService.getScriptCache().put('s',j,21600)}
// 참가자에게는 공개된 번호의 결과와 남은 수량만 보냄(안 열린 번호의 경품 위치는 숨김)
function public_(s){
  var left=s.prizes.map(function(){return 0});left.push(0);
  s.slots.forEach(function(v,k){if(s.opened.indexOf(k+1)<0)left[v<0?s.prizes.length:v]++});
  return {sid:s.sid, ver:s.ver, prizes:s.prizes, left:left, opened:s.opened.map(function(n){return [n,s.slots[n-1]]})};
}
function sheet_(name,head){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);sh.appendRow(head)}
  return sh;
}
function log_(s,n){
  var idx=s.slots[n-1],p=idx<0?FILLER:s.prizes[idx];
  sheet_('기록',['회차','시각','번호','등급','상품']).appendRow([s.round,new Date(),n,p.tier,p.name]);
}
function clean_(raw){
  var a=JSON.parse(raw||'[]');if(!Array.isArray(a))throw 'bad';
  a=a.map(function(p){return {tier:String(p.tier||'').trim().slice(0,20)||'경품', name:String(p.name||'').trim().slice(0,60)||'경품', qty:Math.max(0,parseInt(p.qty,10)||0)}}).filter(function(p){return p.qty>0});
  if(a.reduce(function(t,p){return t+p.qty},0)>N)throw 'bad';
  return a;
}

function doGet(e){
  var q=e.parameter,a=q.a||'state';
  if(a==='state'){var s0=load_();if(s0)return json_(public_(s0))}
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  var s,n,logIt=false,out;
  try{
    s=load_();
    if(!s){s=fresh_(DEFAULT_PRIZES,null);save_(s)}
    if(a==='draw'){
      n=parseInt(q.n,10);
      if(!(n>=1&&n<=N))return json_({ok:false,err:'range'});
      if(s.opened.indexOf(n)>=0){out=public_(s);out.ok=false;out.err='taken';return json_(out)}
      s.opened.push(n);s.ver++;save_(s);logIt=true;
      out=public_(s);out.ok=true;out.n=n;out.idx=s.slots[n-1];
    } else if(a==='reset'||a==='apply'||a==='peek'){
      if(a==='reset'){s=fresh_(s.prizes,s);save_(s)}
      if(a==='apply'){try{s=fresh_(clean_(q.p),s)}catch(err){return json_({ok:false,err:'prizes'})}save_(s)}
      out=public_(s);out.ok=true;
      if(a==='peek')out.slots=s.slots;
    } else {
      out=public_(s);
    }
  } finally {lock.releaseLock()}
  if(logIt)log_(s,n);
  return json_(out);
}
