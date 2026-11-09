/* ===== خادم تطبيق الناخبين على Google Apps Script =====
   التخزين: Google Sheets (ملف VotersDB) + الملفات في Google Drive (مجلد VotersFiles)
   النشر: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone */
var _ss=null;
function prop_(k){return PropertiesService.getScriptProperties().getProperty(k)}
function setup(){ /* شغّلها مرة واحدة ثم افتح Executions/Logs لرؤية مفتاح الإدارة */
  var p=PropertiesService.getScriptProperties();
  if(!p.getProperty('ADMIN_KEY'))p.setProperty('ADMIN_KEY',Utilities.getUuid().replace(/-/g,'').slice(0,16));
  db_();folder_();
  Logger.log('ADMIN_KEY = '+p.getProperty('ADMIN_KEY'));
  Logger.log('Sheet = '+db_().getUrl());
}
function db_(){if(_ss)return _ss;var p=PropertiesService.getScriptProperties(),id=p.getProperty('SSID');
  if(id){try{_ss=SpreadsheetApp.openById(id);return _ss}catch(e){}}
  _ss=SpreadsheetApp.create('VotersDB');p.setProperty('SSID',_ss.getId());return _ss}
function folder_(){var p=PropertiesService.getScriptProperties(),id=p.getProperty('FOLDER');
  if(id){try{return DriveApp.getFolderById(id)}catch(e){}}
  var f=DriveApp.createFolder('VotersFiles');p.setProperty('FOLDER',f.getId());return f}
function sheet_(c){if(!/^[A-Za-z0-9_]{1,30}$/.test(c))throw new Error('اسم مجموعة غير صالح');
  var s=db_(),sh=s.getSheetByName(c);if(!sh){sh=s.insertSheet(c);sh.getRange('A:B').setNumberFormat('@')}return sh}
function all_(c){var sh=sheet_(c),n=sh.getLastRow();if(!n)return[];
  return sh.getRange(1,1,n,2).getValues().map(function(r){return{id:String(r[0]),data:JSON.parse(r[1]||'null')}})}
function rowOf_(sh,id){var n=sh.getLastRow();if(!n)return 0;var v=sh.getRange(1,1,n,1).getValues();for(var i=0;i<n;i++)if(String(v[i][0])===String(id))return i+1;return 0}
function js_(d){var t=JSON.stringify(d);if(t.length>49000)throw new Error('حجم السجل كبير جداً');return t}
function bump_(){PropertiesService.getScriptProperties().setProperty('ver',String(Date.now()))}
function pub_(r){
  switch(r.op){case'ver':case'login':case'stats':case'upload':case'me':case'mevote':return true;
   case'get':return r.c==='registry'||(r.c==='meta'&&r.id==='config');
   case'set':return r.c==='registry'||r.c==='submissions'}
  return false}
function handle_(r){
  var k=prop_('ADMIN_KEY'),admin=!!k&&r.key===k;
  if(!admin&&!pub_(r))throw new Error('غير مصرح (مفتاح الإدارة مطلوب)');
  var sh,row;
  switch(r.op){
   case'ver':return prop_('ver')||'0';
   case'login':return login_(r);
   case'geo':return geo_(r);
   case'get':sh=sheet_(r.c);row=rowOf_(sh,r.id);
     if(!row)return{exists:false,data:null};
     return{exists:true,data:(admin||(r.c==='meta'&&r.id==='config'))?JSON.parse(sh.getRange(row,2).getValue()):{at:1}};
   case'list':return r.lim?page_(r.c,+r.off||0,+r.lim):all_(r.c);
   case'listMany':var o={};r.cs.forEach(function(c){o[c]=all_(c)});return o;
   case'set':return lock_(function(){sh=sheet_(r.c);row=rowOf_(sh,r.id);
     if(!admin){if(row&&r.c==='submissions')throw new Error('مسجّل مسبقاً');if(r.c==='submissions'){r.data.status='new'}}
     if(!admin&&r.c==='submissions'){var d_=r.data||{},ph_=String(d_.ph||'').replace(/[٠-٩]/g,function(c){return '٠١٢٣٤٥٦٧٨٩'.indexOf(c)}).replace(/\D/g,'');
       if(ph_.length!==8)throw new Error('رقم الهاتف يجب أن يكون 8 أرقام بالضبط');d_.ph=ph_;
       ['fn','fa','mo','ln','dob','sx','sect','nid','reg','rg','rq','ra','lg','lq','la'].forEach(function(k){if(!String(d_[k]||'').trim())throw new Error('حقل إلزامي ناقص: '+k)})}
     if(row)sh.getRange(row,2).setValue(js_(r.data));else sh.appendRow([r.id,js_(r.data)]);
     bump_();return 1});
   case'update':return lock_(function(){sh=sheet_(r.c);row=rowOf_(sh,r.id);var cur={};
     if(row)cur=JSON.parse(sh.getRange(row,2).getValue()||'{}');for(var x in r.data)cur[x]=r.data[x];
     if(row)sh.getRange(row,2).setValue(js_(cur));else sh.appendRow([r.id,js_(cur)]);bump_();return 1});
   case'del':return lock_(function(){sh=sheet_(r.c);row=rowOf_(sh,r.id);if(row)sh.deleteRow(row);bump_();return 1});
   case'setMany':return lock_(function(){var by={};r.items.forEach(function(i){(by[i.c]=by[i.c]||[]).push(i)});
     for(var c in by){sh=sheet_(c);var n=sh.getLastRow(),ids=n?sh.getRange(1,1,n,1).getValues():[],m={};
      ids.forEach(function(v,i){m[String(v[0])]=i+1});var add=[];
      by[c].forEach(function(i){var t=js_(i.data);if(m[i.id])sh.getRange(m[i.id],2).setValue(t);else{add.push([i.id,t]);m[i.id]=-1}});
      if(add.length)sh.getRange(sh.getLastRow()+1,1,add.length,2).setValues(add)}
     bump_();return 1});
   case'delMany':return lock_(function(){var by={};r.items.forEach(function(i){(by[i.c]=by[i.c]||[]).push(String(i.id))});
     for(var c in by){sh=sheet_(c);var n=sh.getLastRow();if(!n)continue;var ids=sh.getRange(1,1,n,1).getValues(),set={};by[c].forEach(function(x){set[x]=1});
      for(var i=n-1;i>=0;i--)if(set[String(ids[i][0])])sh.deleteRow(i+1)}
     bump_();return 1});
   case'upload':
     if(!r.b64||r.b64.length>14000000)throw new Error('الملف كبير (الحد حوالي 10 ميغابايت)');
     var mime=r.mime||'application/octet-stream',nm=String(r.name||'file').replace(/[^\w\u0600-\u06FF.\- ]/g,'_');
     var f=folder_().createFile(Utilities.newBlob(Utilities.base64Decode(r.b64),mime,nm));
     f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
     var id=f.getId();return{id:id,url:/^image\//.test(mime)?'https://drive.google.com/thumbnail?id='+id+'&sz=w1600':'https://drive.google.com/file/d/'+id+'/view'};
   case'me':return me_(r.tk,false);
   case'mevote':return lock_(function(){return me_(r.tk,true)});
   case'stats':var vr=prop_('ver')||'0',cc=CacheService.getScriptCache(),ch=cc.get('st'+vr);if(ch)return JSON.parse(ch);var st=stats_();try{cc.put('st'+vr,JSON.stringify(st),600)}catch(e){}return st;
  }
  throw new Error('عملية غير معروفة')}
function page_(c,off,lim){var sh=sheet_(c),n=sh.getLastRow();if(off>=n)return[];var m=Math.min(lim,n-off);return sh.getRange(off+1,1,m,2).getValues().map(function(r){return{id:String(r[0]),data:JSON.parse(r[1]||'null')}})}
function find_(c,tk){var sh=sheet_(c),n=sh.getLastRow();if(!n)return null;var f=sh.getRange(1,2,n,1).createTextFinder('"tk":"'+tk+'"').matchCase(true).findNext();return f?{sh:sh,row:f.getRow()}:null}
function me_(tk,vote){tk=String(tk||'');if(!/^[A-Za-z0-9]{12,40}$/.test(tk))throw new Error('رابط غير صالح');
  var h=find_('voters',tk);
  if(!h){h=find_('submissions',tk);if(!h)throw new Error('لم يُعثر على هذا التسجيل');var s=JSON.parse(h.sh.getRange(h.row,2).getValue());return{found:true,approved:false,status:s.status||'new',name:[s.fn,s.fa,s.ln].join(' ')}}
  var d=JSON.parse(h.sh.getRange(h.row,2).getValue());if(d.del)throw new Error('هذا التسجيل محذوف');
  if(vote&&!d.vt){d.vt=true;d.vat=Date.now();d.vby='self';h.sh.getRange(h.row,2).setValue(js_(d));bump_()}
  return{found:true,approved:true,name:[d.fn,d.fa,d.ln].join(' '),ps:d.ps||'',yr:d.yr||'',rg:d.rg||'',rq:d.rq||'',ra:d.ra||'',vt:!!d.vt,vat:d.vat||null,by:d.vby||''}}
function lock_(fn){var l=LockService.getScriptLock();l.waitLock(30000);try{return fn()}finally{l.releaseLock()}}
function stats_(){
  var cfg={},ann={};all_('meta').forEach(function(m){if(m.id==='config')cfg=m.data;else if(m.id.indexOf('ann')===0)ann[m.id.slice(3)]={winners:m.data.winners,at:m.data.at}});
  var yrs=cfg.years||[new Date().getFullYear()],v={};
  all_('voters').forEach(function(x){var d=x.data;if(!d||d.del)return;var y=d.yr||yrs[0];v[y]=v[y]||{t:0,v:0};v[y].t++;if(d.vt)v[y].v++});
  return{years:yrs,active:cfg.active||yrs[0],voters:v,
    cands:all_('cands').map(function(x){x.data.id=x.id;return x.data}),
    counts:all_('counts').map(function(x){var d=x.data;d.id=x.id;delete d.by;return d}),ann:ann}}
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
/* يستدعيها المتصفح عبر google.script.run (أضمن من fetch داخل صفحة Apps Script) */
function rpc(s){try{return JSON.stringify({r:handle_(JSON.parse(s))})}catch(err){return JSON.stringify({error:String(err.message||err)})}}
function doPost(e){try{return out_({r:handle_(JSON.parse(e.postData.contents))})}catch(err){return out_({error:String(err.message||err)})}}
function doGet(e){
  var p=(e&&e.parameter)||{},url=ScriptApp.getService().getUrl();
  if(p.intake||p.me){var h2=HtmlService.createHtmlOutputFromFile('intake').getContent();h2=h2.replace('<script>',function(){return '<script>window.__GAS='+JSON.stringify(url)+';window.__ME='+JSON.stringify(String(p.me||'').replace(/[^A-Za-z0-9]/g,''))+';</script><script>'});return HtmlService.createHtmlOutput(h2).setTitle('تسجيل بيانات الناخب').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
  var html=HtmlService.createHtmlOutputFromFile('index').getContent();
  var F={gas:url,exec:url,intake:!!p.intake,watch:!!p.watch,k:p.k||''};
  html=html.replace('<script>',function(){return '<script>window.__FORCE='+JSON.stringify(F).replace(/</g,'\\u003c')+';</script><script>'});
  return HtmlService.createHtmlOutput(html).setTitle('الانتخابات النيابية - لبنان')
    .addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}

function hex_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(function(b){return('0'+(b&255).toString(16)).slice(-2)}).join('')}
/* دخول بدون مفتاح: يتحقق من اسم المستخدم وكلمة المرور ثم يعيد مفتاح الإدارة للجهاز */
function login_(r){var k=prop_('ADMIN_KEY');if(!k)throw new Error('شغّل الدالة setup في Apps Script أولاً');
  var u=String(r.u||'').trim().toLowerCase(),h=String(r.h||''),list=all_('accounts');
  if(!list.length){if(u==='admin'&&h===hex_('hb|admin'))return{k:k,id:'admin0'}}
  else for(var i=0;i<list.length;i++){var d=list[i].data;if(d&&String(d.u||'').toLowerCase()===u&&d.h===h)return{k:k,id:list[i].id}}
  Utilities.sleep(800);throw new Error('اسم المستخدم أو كلمة المرور غير صحيحة')}

/* تحديد مكان وعنوان من خرائط Google (خدمة Maps المدمجة في Apps Script، بدون مفتاح) */
function geo_(r){var q=String(r.q||'').trim();if(!q)throw new Error('اكتب اسم المكان');
  var vs=[q+' لبنان',q,q+', Lebanon'],gc=Maps.newGeocoder().setLanguage('ar').setRegion('lb');
  for(var i=0;i<vs.length;i++){var res=gc.geocode(vs[i]);
    if(res.status==='OK'&&res.results&&res.results.length){var x=res.results[0];return{addr:x.formatted_address,lat:x.geometry.location.lat,lng:x.geometry.location.lng}}}
  throw new Error('لم يُعثر على المكان في خرائط Google')}
