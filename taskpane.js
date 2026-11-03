/* שיוך למשימה – איילון | Outlook add-in taskpane (mail read). Talks to the Apps Script web app with text/plain POSTs. */
(function(){
'use strict';
var APP=window.EYALON_APP_URL, MAXF=25*1024*1024, TK='eyalon_tok', EM='eyalon_email';
var S={tok:'',me:null,mail:null,atts:[],sel:null,mode:'pick',q:'',proj:'',busy:false,email:''};
function $(id){return document.getElementById(id)}
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})}
function show(id,on){$(id).hidden=!on}
function msg(t,k){var m=$('msg');m.hidden=!t;m.className='msg '+(k||'info');m.textContent=t||''}
function kb(n){return n>1048576?(n/1048576).toFixed(1)+'MB':Math.max(1,Math.round(n/1024))+'KB'}

// ---- API: simple request (text/plain) => no CORS preflight; fetch follows Apps Script 302 to googleusercontent ----
function api(fn){var args=[].slice.call(arguments,1);
  return fetch(APP,{method:'POST',redirect:'follow',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({addin:1,fn:fn,args:args})})
   .then(function(r){if(!r.ok)throw new Error('שגיאת רשת '+r.status);return r.json()})
   .then(function(j){if(!j.ok){if(j.error==='AUTH'){logout(true);throw new Error('פג תוקף הכניסה – יש להתחבר מחדש')}throw new Error(j.error||'שגיאה')}return j.result})}
function tapi(fn){var a=[].slice.call(arguments,1);return api.apply(null,[fn,S.tok].concat(a))}

// ---- token memory: Outlook roamingSettings (follows the user across devices) + localStorage fallback ----
function rs(){try{return Office.context.roamingSettings}catch(e){return null}}
function loadTok(){var r=rs(),t=r&&r.get(TK);try{t=t||localStorage.getItem(TK)}catch(e){}var e=(r&&r.get(EM))||'';try{e=e||localStorage.getItem(EM)||''}catch(x){}S.email=e;return t||''}
function saveTok(t,e){var r=rs();if(r){if(t){r.set(TK,t);r.set(EM,e||'')}else r.remove(TK);r.saveAsync(function(){})}try{t?(localStorage.setItem(TK,t),localStorage.setItem(EM,e||'')):localStorage.removeItem(TK)}catch(x){}}
function logout(expired){S.tok='';S.me=null;saveTok('');show('vMain',false);show('btnOut',false);showLogin();if(!expired)msg('')}

// ---- current item ----
function readItem(){var it=Office.context.mailbox.item;if(!it)return null;
  var f=it.from||it.sender||{},d=it.dateTimeCreated||it.dateTimeModified,iid=it.itemId||'',url='';
  try{var mb=Office.context.mailbox,rid=mb.convertToRestId?mb.convertToRestId(iid,Office.MailboxEnums.RestVersion.v2_0):iid;
    url='https://outlook.office.com/mail/deeplink/read/'+encodeURIComponent(rid)+'?ItemID='+encodeURIComponent(iid)+'&exvsurl=1';}catch(e){url='https://outlook.office.com/owa/?ItemID='+encodeURIComponent(iid)+'&exvsurl=1&viewmodel=ReadMessageItem'}
  S.atts=(it.attachments||[]).filter(function(a){return !a.isInline}).map(function(a){return {id:a.id,name:a.name,size:a.size||0,type:a.attachmentType,ct:a.contentType||'',on:false}});
  return {item_id:iid,internet_message_id:it.internetMessageId||'',conversation_id:it.conversationId||'',subject:it.subject||'',from:(f.displayName?f.displayName+' <'+f.emailAddress+'>':f.emailAddress||''),date:d?new Date(d).toISOString():'',url:url}}
function drawMail(){var m=S.mail;$('mail').innerHTML=m?'<div class="s">✉ '+esc(m.subject||'(ללא נושא)')+'</div><div class="m">'+esc(m.from)+(m.date?' · '+esc(new Date(m.date).toLocaleString('he-IL',{dateStyle:'short',timeStyle:'short'})):'')+(S.atts.length?' · 📎 '+S.atts.length:'')+'</div>':'<div class="m">לא נבחר מייל</div>'}
function snippet(){return new Promise(function(res){try{Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text,function(r){res(r.status==='succeeded'?String(r.value||'').replace(/\s+/g,' ').trim().slice(0,1500):'')})}catch(e){res('')}})}
function attContent(a){return new Promise(function(res,rej){var it=Office.context.mailbox.item;if(!it.getAttachmentContentAsync)return rej(new Error('גרסת Outlook ישנה – אין תמיכה בהעלאת קבצים (נדרש Mailbox 1.8)'));
  it.getAttachmentContentAsync(a.id,function(r){if(r.status!=='succeeded')return rej(new Error(r.error&&r.error.message||'קריאת הקובץ נכשלה'));var v=r.value,F=Office.MailboxEnums.AttachmentContentFormat;
    if(v.format===F.Base64)return res({b64:v.content,name:a.name,mime:a.ct||'application/octet-stream'});
    if(v.format===F.Eml||v.format===F.ICalendar){var b=btoa(unescape(encodeURIComponent(v.content)));return res({b64:b,name:/\.(eml|ics)$/i.test(a.name)?a.name:a.name+(v.format===F.Eml?'.eml':'.ics'),mime:v.format===F.Eml?'message/rfc822':'text/calendar'})}
    rej(new Error('קובץ ענן (קישור) – לא ניתן להעלות: '+a.name))})})}

// ---- login ----
function showLogin(){show('vLogin',true);show('lStep1',true);show('lStep2',false);$('lEmail').value=S.email||(Office.context.mailbox.userProfile||{}).emailAddress||''}
function doCode(){var e=$('lEmail').value.trim();if(!e)return;busy('btnCode',1);api('requestCode',e).then(function(){S.email=e;show('lStep1',false);show('lStep2',true);$('lCode').focus();msg('הקוד נשלח למייל '+e,'info')}).catch(err).then(function(){busy('btnCode',0)})}
function doVerify(){busy('btnVerify',1);api('verifyCode',S.email,$('lCode').value.trim()).then(function(r){S.tok=r.token;saveTok(r.token,r.email);msg('');return start()}).catch(err).then(function(){busy('btnVerify',0)})}
function busy(id,on){$(id).disabled=!!on}
function err(e){msg(String(e&&e.message||e),'err')}

// ---- main ----
function start(){return tapi('addinMe').then(function(me){S.me=me;if(me.external)throw new Error('משתמש חיצוני אינו רשאי לשייך מיילים');show('vLogin',false);show('vMain',true);show('btnOut',true);
  var o=me.projects.map(function(p){return '<option>'+esc(p)+'</option>'}).join('');$('fProj').innerHTML='<option value="">כל הפרויקטים</option>'+o;$('nProj').innerHTML='<option value="">ללא פרויקט</option>'+o;
  $('nTitle').value=S.mail?S.mail.subject.replace(/^((RE|FW|FWD|השב|העבר)\s*:\s*)+/i,'').slice(0,300):'';drawAtts();lookup();return loadTasks()})}
function lookup(){if(!S.mail)return;tapi('addinLookup',[S.mail.item_id,S.mail.internet_message_id]).then(function(L){var a=$('already');a.hidden=!L.length;a.innerHTML=L.length?'🔗 המייל כבר משויך ל: '+L.map(function(t){return '<b>'+esc(t.title)+'</b> ('+esc(t.project||'—')+')'}).join(', '):''}).catch(function(){})}
var qT=0;function loadTasks(){clearTimeout(qT);return new Promise(function(res){qT=setTimeout(function(){$('list').innerHTML='<div class="it m">טוען…</div>';tapi('addinTasks',S.q,S.proj).then(drawList).catch(err).then(res)},S.q?300:0)})}
function row(t){return '<div class="it'+(S.sel&&S.sel.id===t.id?' on':'')+'" data-id="'+esc(t.id)+'"><b>'+esc(t.title)+'</b><span class="m">'+esc((t.project||'ללא פרויקט')+' · '+t.owner+' · '+t.status+(t.deadline?' · יעד '+t.deadline:'')+(t.emails?' · ✉'+t.emails:''))+'</span></div>'}
var LAST={recent:[],tasks:[]};
function drawList(r){LAST=r;var ids={},h='';if(r.recent.length){h+='<h4>אחרונות ששייכת אליהן</h4>'+r.recent.map(function(t){ids[t.id]=1;return row(t)}).join('')}
  var rest=r.tasks.filter(function(t){return !ids[t.id]});h+='<h4>'+(S.q?'תוצאות חיפוש ('+rest.length+')':'משימות פתוחות')+'</h4>'+(rest.length?rest.map(row).join(''):'<div class="it m">לא נמצאו משימות</div>');$('list').innerHTML=h;upd()}
function pick(id){var all=LAST.recent.concat(LAST.tasks);S.sel=all.filter(function(t){return t.id===id})[0]||null;drawList(LAST)}
function drawAtts(){$('atts').innerHTML=S.atts.length?'<div class="m">קבצים מצורפים להעלאה לתיקיית המשימה:</div>'+S.atts.map(function(a,i){var big=a.size>MAXF||a.type==='cloud';return '<label class="att'+(big?' big':'')+'"><input type="checkbox" data-i="'+i+'"'+(big?' disabled':'')+(a.on?' checked':'')+'> 📄 '+esc(a.name)+' <span class="m">'+kb(a.size)+(a.size>MAXF?' · מעל 25MB':a.type==='cloud'?' · קובץ ענן':'')+'</span></label>'}).join(''):''}
function upd(){var ok=S.mode==='new'?!!$('nTitle').value.trim():!!S.sel;$('btnGo').disabled=!ok||S.busy||!S.mail;var s=$('sel');s.hidden=S.mode!=='pick'||!S.sel;if(S.sel)s.innerHTML='נבחרה: <b>'+esc(S.sel.title)+'</b> · '+esc(S.sel.project||'ללא פרויקט');$('btnGo').textContent=S.mode==='new'?'צור משימה ושייך':'שייך למשימה'}
function setMode(m){S.mode=m;$('tPick').classList.toggle('on',m==='pick');$('tNew').classList.toggle('on',m==='new');show('pPick',m==='pick');show('pNew',m==='new');upd()}
function go(){if(S.busy)return;S.busy=true;upd();msg('');var p=$('prog'),task,mail=Object.assign({},S.mail),files=S.atts.filter(function(a){return a.on}),up=0,fail=[];
  (S.mode==='new'?(p.textContent='יוצר משימה…',tapi('createTask',{title:$('nTitle').value.trim(),project:$('nProj').value,deadline:$('nDl').value,priority:$('nPr').value,description:'נוצרה מתוך מייל: '+mail.subject+' · '+mail.from})):Promise.resolve(S.sel))
  .then(function(t){task=t;return $('oSnip').checked?snippet():''})
  .then(function(sn){if(sn)mail.snippet=sn;p.textContent='משייך את המייל…';return tapi('addinLinkEmail',task.id,mail)})
  .then(function(r){var chain=Promise.resolve();files.forEach(function(a,i){chain=chain.then(function(){p.textContent='מעלה קובץ '+(i+1)+'/'+files.length+': '+a.name;return attContent(a).then(function(c){return tapi('addinUpload',task.id,c.name,c.mime,c.b64,mail.internet_message_id||mail.item_id)}).then(function(){up++}).catch(function(e){fail.push(a.name+' – '+(e.message||e))})})});
    return chain.then(function(){p.textContent='';msg((r.duplicate?'המייל כבר היה משויך ל':'✔ המייל שויך ל')+'"'+(task.title||task.subject)+'"'+(files.length?' · הועלו '+up+'/'+files.length+' קבצים':'')+(fail.length?'\nנכשלו: '+fail.join('; '):''),fail.length?'err':'ok');lookup();S.sel=null;setMode('pick');return loadTasks()})})
  .catch(function(e){p.textContent='';err(e)}).then(function(){S.busy=false;upd()})}

function bind(){$('btnCode').onclick=doCode;$('btnVerify').onclick=doVerify;$('btnBack').onclick=function(){show('lStep1',true);show('lStep2',false)};$('btnOut').onclick=function(){logout()};
  $('lCode').onkeydown=function(e){if(e.key==='Enter')doVerify()};$('lEmail').onkeydown=function(e){if(e.key==='Enter')doCode()};
  $('fQ').oninput=function(){S.q=this.value.trim();loadTasks()};$('fProj').onchange=function(){S.proj=this.value;loadTasks()};
  $('list').onclick=function(e){var it=e.target.closest('.it[data-id]');if(it)pick(it.getAttribute('data-id'))};
  $('atts').onchange=function(e){var i=e.target.getAttribute('data-i');if(i!=null)S.atts[+i].on=e.target.checked};
  $('tPick').onclick=function(){setMode('pick')};$('tNew').onclick=function(){setMode('new')};$('nTitle').oninput=upd;$('btnGo').onclick=go}
function onItem(){S.mail=readItem();S.sel=null;drawMail();if(S.me){drawAtts();lookup();$('nTitle').value=S.mail?S.mail.subject.slice(0,300):'';upd()}}
Office.onReady(function(){bind();if(!APP||/\{\{/.test(APP)){msg('התוסף לא הוגדר: חסרה כתובת האפליקציה (config.js)','err');return}
  S.mail=readItem();drawMail();try{Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged,onItem)}catch(e){}
  S.tok=loadTok();if(!S.tok){showLogin();return}start().catch(function(e){if(S.tok)err(e)})});
window.__eyalon={S:S,api:api};
})();
