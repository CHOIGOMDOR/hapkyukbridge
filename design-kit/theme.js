/* 디자인에 관련된 JS만 모았어요. 앱 데이터(시트, 퀴즈 등) 코드는 reference.html 에 있어요 */
var THEMES = [ // 월~일 : pri(메인) pri2(그라데이션 짝) soft(연한 배경) tint(중간 배경) deep(진한 강조)
  {n:"연보라", pri:"#b9a6f5", pri2:"#8fb4f7", soft:"#f4f0ff", tint:"#e4dcff", deep:"#4b2fb5"},
  {n:"연하늘", pri:"#8fd0f4", pri2:"#7fe3dc", soft:"#ebf8fe", tint:"#d0eefb", deep:"#17618f"},
  {n:"연노랑", pri:"#f5d86e", pri2:"#ffb877", soft:"#fffae3", tint:"#ffefb0", deep:"#7a5a00"},
  {n:"연분홍", pri:"#f4a9c8", pri2:"#ffb9a0", soft:"#fff1f7", tint:"#ffdbe9", deep:"#a02a62"},
  {n:"연민트", pri:"#8fdcbf", pri2:"#b5e88a", soft:"#e9faf3", tint:"#cdf2e2", deep:"#17704f"},
  {n:"연살구", pri:"#f7bd98", pri2:"#f7a1b4", soft:"#fff3ea", tint:"#ffe0cc", deep:"#9a4b14"},
  {n:"연연두", pri:"#c3e38a", pri2:"#8fe0b0", soft:"#f5fbe6", tint:"#e3f3bf", deep:"#3f6a0a"}
];
var DAYN = ["월","화","수","목","금","토","일"];

/* ============ 날짜 (KST) ============ */
var fmt = new Intl.DateTimeFormat("ko-KR",{timeZone:"Asia/Seoul",year:"numeric",month:"long",day:"numeric",weekday:"long",hour:"2-digit",minute:"2-digit",second:"2-digit",hour12:false});
function parts(){var o={};fmt.formatToParts(new Date()).forEach(function(p){o[p.type]=p.value});return o}
function todayStr(){return new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Seoul"}).format(new Date())}
function utc(s){return Date.UTC(+s.slice(0,4),+s.slice(5,7)-1,+s.slice(8,10))}
function addDays(s,n){return new Date(utc(s)+n*864e5).toISOString().slice(0,10)}
function wIdx(s){return (new Date(utc(s)).getUTCDay()+6)%7}   // 월0 ~ 일6
var TODAY = (location.search.match(/date=(\d{4}-\d{2}-\d{2})/)||[])[1] || todayStr(); // ?date=YYYY-MM-DD 로 다른 날 미리보기

/* ============ 유틸 ============ */
function $(id){return document.getElementById(id)}
function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]})}
function pick(a){return a[Math.floor(Math.random()*a.length)]}
function josa(w,a,b){var c=w.charCodeAt(w.length-1);if(c<0xAC00||c>0xD7A3)return a;return (c-0xAC00)%28?a:b}
var toastT;function toast(m,ms){var t=$("toast");t.textContent=m;t.classList.add("on");clearTimeout(toastT);toastT=setTimeout(function(){t.classList.remove("on")},ms||1800)}
function lget(k,d){try{var v=localStorage.getItem(k);return v?JSON.parse(v):d}catch(e){return d}}
function lset(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){}}
var CHEERS=["오늘도 화이팅! 💪","한 걸음씩, 천천히 가자","잘하고 있어, 치즈볼!","모르는 건 사진 찍어 두기 📷","오늘 할 일 다 하면 ⭐","수학은 하다 보면 쉬워져","차근차근 하나씩 ✨","작은 습관이 큰 실력!","쉬엄쉬엄 해도 괜찮아","오늘의 나, 칭찬해 🎉"];
function dayNum(){return Math.floor(utc(TODAY)/864e5)}
function kindOf(u){return /pompom/.test(u)?"jump":/lying/.test(u)?"crawl":/friends/.test(u)?"sway":"float"}
function setImg(el,u){el.onerror=function(){el.onerror=null;var f=IMAGES[0];el.src=IMG_DATA[f]||f};el.src=IMG_DATA[u]||u;el.dataset.kind=kindOf(u)}
function pickImg(used){
  var pool=IMAGES.filter(function(u,i){return IMAGES.indexOf(u)===i});
  var recent=lget("img-recent",[]),keep=Math.min(Math.max(pool.length-1,0),12);
  var fresh=pool.filter(function(u){return used.indexOf(u)<0&&recent.slice(-keep).indexOf(u)<0});
  if(!fresh.length)fresh=pool.filter(function(u){return used.indexOf(u)<0});
  if(!fresh.length)fresh=pool;
  var u=pick(fresh);recent.push(u);lset("img-recent",recent.slice(-30));return u;
}
/* 그림은 열 때마다, 탭을 바꿀 때마다, 눌렀을 때마다 새 그림으로 바뀌어요 (같은 화면에 같은 그림이 겹치지 않게) */
function assignImages(){
  var used=[];
  [].forEach.call(document.querySelectorAll("img.slot"),function(el){var u=pickImg(used);used.push(u);setImg(el,u)});
  $("bubble").textContent=pick(CHEERS);
}
document.addEventListener("click",function(e){
  var el=e.target.closest&&e.target.closest("img.slot");if(!el)return;
  var used=[].map.call(document.querySelectorAll("img.slot"),function(x){return x.getAttribute("src")});
  var u=pickImg([]);setImg(el,u);
});

/* ============ 데이터 ============ */
var DB = {plan:[],log:[],vocab:[],repeats:[],rchecks:[],minutes:[],qlog:[],questions:[],qa:[],concepts:[],mood:{}};
var SUN=["일","월","화","수","목","금","토"];                       // 한 주는 일요일에 시작해요
var KINDS=["한국사","학원","수학","학원","수학","학원","쉬는 날"]; // 일~토 순서
function sIdx(d){return (wIdx(d)+1)%7}
var POS_LABEL={"n.":"명사","v.":"동사","adj.":"형용사","adv.":"부사","prep.":"전치사","phr.":"구(숙어)","phr. v.":"구동사","pron.":"대명사"};

function samplePlan(){
  var mon=addDays(weekStart(TODAY),1), rows=[], id=0;
  function add(d,subj,unit,task,range){rows.push({id:"s"+(id++),date:addDays(mon,d),subject:subj,unit:unit,task:task,range:range||"",link:"",done:false})}
  var AC=["5:10~8:00","","5:20~9:00","","5:10~8:00"];
  [0,2,4].forEach(function(d){add(d,"영어 학원","Pine English School","영어 학원 ("+AC[d]+")");add(d,"수학 · 연산","소인수분해","연산 학습지","30분")});
  add(0,"영어 · 단어","유닛 7·8","단어 시험 (학원)","시험");
  [1,3].forEach(function(d){
    add(d,"수학 · 개념","소인수분해",d===1?"숨마쿰 EBS 강의 + 개념 읽기":"개념원리",d===1?"p.25~34":"p.20~29");
    add(d,"수학 · 문제","소인수분해",d===1?"숨마쿰 유형 EXERCISES":"쎈",d===1?"해당 쪽 문제":"p.15~21");
    add(d,"영어 · 문법","Grammar Inside STARTER · Ch01 명사와 관사","Unit 0"+(d===1?2:3)+" "+(d===1?"셀 수 없는 명사":"관사"),d===1?"교재 p.12~13 + 워크북 p.6~7":"교재 p.14~15 + 워크북 p.8~9");
  });
  add(3,"과학","초등 과학 6-2","1 계절의 변화: 남중 고도와 낮의 길이");
  [1,2,3,4,5,6].forEach(function(d){add(d,"영어 · 단어","유닛 9·10 (다음 월 시험)","단어 퀴즈 3문제","3문제")});
  add(6,"역사","한국사 1권 선사~고대","01 선사 시대, 돌을 사용하던 사람들","p.16");
  add(6,"역사","한국사 1권 선사~고대","02 단군 할아버지가 세운 국가, 고조선","p.24");
  return rows;
}
function sampleVocab(){
  return [[9,"Asian New Year Feasts","to welcome","v.","맞이하다,환영하다","O",""],[9,"Asian New Year Feasts","surplus","n.","과잉,잉여","O",""],[9,"Asian New Year Feasts","ancient","adj.","옛날의,고대의","O",""],[9,"Asian New Year Feasts","an ingredient","n.","재료","O",""],[9,"Asian New Year Feasts","abundance","n.","풍부,충만","O",""],[9,"Asian New Year Feasts","culture","n.","문화","O",""],[9,"Asian New Year Feasts","sticky","adj.","끈적끈적한","O",""],[9,"Asian New Year Feasts","truly","adv.","정말로","O",""],[9,"Asian New Year Feasts","to symbolize","v.","상징하다","X",""],[10,"Food Preservation","innovative","adj.","혁신적인","O",""],[10,"Food Preservation","a benefit","n.","혜택,이득","O",""],[10,"Food Preservation","a method","n.","방법","O",""],[10,"Food Preservation","vinegar","n.","식초","O",""],[10,"Food Preservation","liquid","n.","액체","O",""],[10,"Food Preservation","harmful","adj.","해로운","O","빨간★"],[10,"Food Preservation","to prove","v.","입증하다,증명하다","O",""],[10,"Food Preservation","to store","v.","저장하다,보관하다","X",""]];
}
function sampleLog(){
  return [{date:addDays(TODAY,-1),mood:"hard",c1:"소수는 약수가 2개인 자연수예요.",c2:"소인수분해는 소수들의 곱으로 나타내는 거예요.",c3:"작은 소수부터 차례로 나누어 보세요.",next:"개념원리 p.20~29, 쎈 p.15~21",summary:"(예시) 12를 2×2×3으로 나누는 과정에서 헷갈렸다고 말했어요.",photos:[]}];
}
function toVocab(a){return {u:+a[0],title:a[1],en:String(a[2]||""),pos:a[3]||"",ko:a[4]||"",spell:/^O/.test(a[5]||""),mark:a[6]||""}}
function applyData(j){
  DB.plan=j.plan||[]; DB.log=j.log||[]; DB.mood={};
  DB.log.forEach(function(l){if(l.mood)DB.mood[l.date]=l.mood});
  DB.vocab=(j.vocab||[]).map(toVocab);
  DB.qa=j.qa||[]; DB.concepts=j.concepts||[]; DB.repeats=j.repeats||[]; DB.rchecks=j.rchecks||[]; DB.minutes=j.minutes||[]; DB.qlog=j.qlog||[]; DB.questions=j.questions||[];
  if(j.images&&j.images.length)j.images.forEach(function(u){if(IMAGES.indexOf(u)<0)IMAGES.push(u)});
  if(j.goal){for(var k in j.goal){if(j.goal[k])GOAL[k]=j.goal[k]}}
}
function loadData(){
  if(HAS_GS){
    return new Promise(function(ok,no){google.script.run.withSuccessHandler(function(j){applyData(j);ok()}).withFailureHandler(no).apiGet()});
  }
  if(!LIVE){
    DB.plan=samplePlan(); DB.log=sampleLog(); DB.vocab=sampleVocab().map(toVocab);
    var saved=lget("local-done",{});
    DB.plan.forEach(function(r){if(saved[r.id])r.done=true});
    DB.mood=lget("local-mood",{});
    DB.repeats=[{name:"큐티",day:"자유",times:2,subject:"큐티",memo:"일요일 모임 준비"},{name:"주제 글쓰기",day:"금",times:1,subject:"학교 숙제",memo:"학교 글쓰기 숙제"}];
    DB.rchecks=lget("local-rchecks",[]); DB.minutes=lget("local-min",[]); DB.qlog=lget("local-qlog",[]);
    DB.questions=[{subject:"역사",key:"01 선사 시대, 돌을 사용하던 사람들",q:"돌을 깨뜨려 만든 도구를 쓰던 시대는?",opts:["구석기 시대","청동기 시대","철기 시대","삼국 시대"],ans:1,exp:"뗀석기를 쓰던 구석기 시대예요."}];
    return Promise.resolve();
  }
  return fetch(API_URL+"?action=all").then(function(r){return r.json()}).then(applyData);
}
function send(body){
  if(HAS_GS){
    return new Promise(function(ok){google.script.run.withSuccessHandler(function(r){if(r&&r.ok===false){toast("저장이 안 됐어요: "+(r.error||"알 수 없는 오류"),6000);ok()}else ok(r)}).withFailureHandler(function(er){toast("저장이 안 됐어요: "+((er&&er.message)||"연결을 확인해 주세요"),6000);ok()}).apiPost(JSON.stringify(body))});
  }
  if(!LIVE)return Promise.resolve();
  return fetch(API_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(body)}).then(function(r){return r.json()}).catch(function(){toast("저장이 안 됐어요. 다시 시도해 줘요")});
}
function persistLocal(){lset("local-rchecks",DB.rchecks);lset("local-min",DB.minutes);lset("local-qlog",DB.qlog)}

/* ============ 시계 / D-day ============ */
function tick(){
  var o=parts(),h=o.hour==="24"?"00":o.hour;
  $("date").textContent=o.year+"년 "+o.month+" "+o.day+" "+o.weekday;
  $("time").textContent=h+":"+o.minute+":"+o.second;
  var d=Math.round((utc(GOAL.end)-utc(todayStr()))/864e5);
  var dd=(d>0?"D-"+d:d===0?"D-Day":"D+"+(-d));if($("dday").dataset.v!==dd){$("dday").dataset.v=dd;$("dday").innerHTML=dd+"<small>1월 중후반 목표</small>"}
}

/* ============ 오늘 할 일 계산 ============ */
/* ---- 오늘의 개념 세 문장: 오늘 공부할 단원에서 뽑아요 (월·수·금=연산, 화·목·일=개념, 토=이번 주 복습) ---- */
function mathUnit(d){
  var m=DB.plan.filter(function(r){return /^수학/.test(r.subject)&&r.unit});if(!m.length)return null;
  var u=null,x=m.filter(function(r){return r.date===d})[0];
  if(x)u=x.unit;
  else{var ws=weekStart(d),we=addDays(ws,6),wk=m.filter(function(r){return r.date>=ws&&r.date<=we});
    if(!wk.length)return null;
    var after=wk.filter(function(r){return r.date>=d})[0];u=(after||wk[wk.length-1]).unit}
  var first=m.filter(function(r){return r.unit===u}).map(function(r){return r.date}).sort()[0];
  return {unit:u,first:first};
}
function conceptSet(d){
  var mu=mathUnit(d);if(!mu||!DB.concepts.length)return null;
  var w=wIdx(d),isSat=w===5,kind=(w===0||w===2||w===4)?"연산":"개념";
  var all=DB.concepts.filter(function(c){return c[0]===mu.unit});
  var list=all.filter(function(c){return c[1]===(isSat?"개념":kind)}).map(function(c){return c[2]});
  if(!list.length)return null;
  var out=[];
  if(isSat)out=list.slice(0,3);
  else{var k=Math.max(0,Math.round((utc(d)-utc(mu.first))/864e5)),st=(k*3)%list.length;
    for(var i=0;i<Math.min(3,list.length);i++)out.push(list[(st+i)%list.length])}
  return {items:out,chip:mu.unit+" · "+(isSat?"이번 주 복습":kind==="연산"?"연산 요령":"핵심 개념")};
}
/* ---- 줄바꿈 정리: "A + B" 는 줄을 나눠서, 분량은 알약으로 ---- */
function fmtTitle(t){
  var p=String(t||"").split(/\s+\+\s+/);
  return p.map(function(x,i){return '<span class="tl">'+(p.length>1?'<i>'+(i+1)+'</i>':'')+esc(x)+'</span>'}).join("");
}
function fmtRange(r){
  if(!r)return "";
  return '<span class="rgs">'+String(r).split(/\s+\+\s+/).map(function(x){return '<span class="rg">'+esc(x)+'</span>'}).join("")+'</span>';
}
function tasksOf(date){return DB.plan.filter(function(r){return r.date===date})}
function weekStart(d){return addDays(d,-sIdx(d))}
function repeatItems(date){
  var out=[], wi=wIdx(date), mon=weekStart(date), sun=addDays(mon,6);
  DB.repeats.forEach(function(r){
    var times=+r.times||1;
    var doneOn=DB.rchecks.some(function(c){return c.name===r.name&&c.date===date});
    if(r.day==="자유"){
      if(wi===6)return;                                  // 일요일은 쉬어요
      var cnt=DB.rchecks.filter(function(c){return c.name===r.name&&c.date>=mon&&c.date<=sun}).length;
      if(cnt>=times&&!doneOn)return;                     // 이번 주 목표를 채우면 사라져요
      out.push({id:"rep:"+r.name,rep:true,flex:true,subject:r.subject||"반복",task:r.name,range:"이번 주 "+cnt+"/"+times,unit:r.memo||"",done:doneOn});
    }else if(r.day===DAYN[wi]){
      out.push({id:"rep:"+r.name,rep:true,subject:r.subject||"반복",task:r.name,range:"",unit:r.memo||"",done:doneOn});
    }
  });
  return out;
}
function dayItems(date){
  var t=tasksOf(date).slice();
  repeatItems(date).forEach(function(x){if(!x.flex||date===TODAY)t.push(x)});
  return t;
}
function allDone(date){var t=dayItems(date);return t.length>0&&t.every(function(x){return x.done})}
function findPlan(id){return DB.plan.filter(function(x){return String(x.id)===String(id)})[0]}
function setDone(r,on){
  r.done=on;
  if(!LIVE){var s=lget("local-done",{});s[r.id]=on;lset("local-done",s)}
  else send({action:"check",row:r.id,done:on});
}
function setRepeat(name,date,on){
  DB.rchecks=DB.rchecks.filter(function(c){return !(c.name===name&&c.date===date)});
  if(on)DB.rchecks.push({date:date,name:name});
  if(!LIVE)persistLocal(); else send({action:"rcheck",date:date,name:name,on:on});
}

/* ============ 순공시간 ============ */
function minsOn(d){return DB.minutes.filter(function(m){return m.date===d}).reduce(function(a,m){return a+(+m.min||0)},0)}
function minsBetween(a,b){return DB.minutes.filter(function(m){return m.date>=a&&m.date<=b}).reduce(function(s,m){return s+(+m.min||0)},0)}
function fmtMin(n){var h=Math.floor(n/60),m=n%60;return h?(h+"시간"+(m?" "+m+"분":"")):(m+"분")}
function addMinutes(min){
  min=Math.round(+min);if(!(min>0)){toast("시간을 입력해 줘요");return}
  DB.minutes.push({date:TODAY,min:min});
  if(!LIVE)persistLocal(); else send({action:"minutes",date:TODAY,min:min});
  renderToday();toast("순공 "+fmtMin(min)+" 기록했어요 ⏱");
}
function undoMinutes(){
  for(var i=DB.minutes.length-1;i>=0;i--){if(DB.minutes[i].date===TODAY){var m=DB.minutes.splice(i,1)[0];
    if(!LIVE)persistLocal(); else send({action:"minutesUndo",date:TODAY});
    renderToday();toast(fmtMin(m.min)+" 취소했어요");return}}
  toast("오늘 취소할 기록이 없어요");
}

/* ============ 렌더 ============ */
function applyTheme(){
  var t=THEMES[wIdx(TODAY)],s=document.documentElement.style;
  s.setProperty("--pri",t.pri);s.setProperty("--pri2",t.pri2);s.setProperty("--soft",t.soft);s.setProperty("--tint",t.tint);s.setProperty("--deep",t.deep);
}

function kindOf(u){return /pompom/.test(u)?"jump":/lying/.test(u)?"crawl":/friends/.test(u)?"sway":"float"}
function setImg(el,u){el.onerror=function(){el.onerror=null;var f=IMAGES[0];el.src=IMG_DATA[f]||f};el.src=IMG_DATA[u]||u;el.dataset.kind=kindOf(u)}
function pickImg(used){
  var pool=IMAGES.filter(function(u,i){return IMAGES.indexOf(u)===i});
  var recent=lget("img-recent",[]),keep=Math.min(Math.max(pool.length-1,0),12);
  var fresh=pool.filter(function(u){return used.indexOf(u)<0&&recent.slice(-keep).indexOf(u)<0});
  if(!fresh.length)fresh=pool.filter(function(u){return used.indexOf(u)<0});
  if(!fresh.length)fresh=pool;
  var u=pick(fresh);recent.push(u);lset("img-recent",recent.slice(-30));return u;
}
/* 그림은 열 때마다, 탭을 바꿀 때마다, 눌렀을 때마다 새 그림으로 바뀌어요 (같은 화면에 같은 그림이 겹치지 않게) */
function assignImages(){
  var used=[];
  [].forEach.call(document.querySelectorAll("img.slot"),function(el){var u=pickImg(used);used.push(u);setImg(el,u)});
  $("bubble").textContent=pick(CHEERS);
}
document.addEventListener("click",function(e){
  var el=e.target.closest&&e.target.closest("img.slot");if(!el)return;
  var used=[].map.call(document.querySelectorAll("img.slot"),function(x){return x.getAttribute("src")});
  var u=pickImg([]);setImg(el,u);
});


function showTab(b){
  if(b.classList.contains("on"))return;
  [].forEach.call(document.querySelectorAll("nav button"),function(x){x.classList.toggle("on",x===b)});
  ["today","plan","qa","log"].forEach(function(id){$(id).classList.toggle("hide",id!==b.dataset.tab)});
  window.scrollTo(0,0);assignImages();
}
function tick(){
  var o=parts(),h=o.hour==="24"?"00":o.hour;
  $("date").textContent=o.year+"년 "+o.month+" "+o.day+" "+o.weekday;
  $("time").textContent=h+":"+o.minute+":"+o.second;
  var d=Math.round((utc(GOAL.end)-utc(todayStr()))/864e5);
  var dd=(d>0?"D-"+d:d===0?"D-Day":"D+"+(-d));if($("dday").dataset.v!==dd){$("dday").dataset.v=dd;$("dday").innerHTML=dd+"<small>1월 중후반 목표</small>"}
}

var toastT;function toast(m,ms){var t=$("toast");t.textContent=m;t.classList.add("on");clearTimeout(toastT);toastT=setTimeout(function(){t.classList.remove("on")},ms||1800)}