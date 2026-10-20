(function(){
'use strict';
var C=Object.assign({endpoint:'',downloadUrl:'downloads/OKV-CBT-Setup.zip',price:120000,whatsapp:'2348104141138'},window.OKV_CONFIG||{});
var $=function(s,r){return (r||document).querySelector(s)},$$=function(s,r){return Array.prototype.slice.call((r||document).querySelectorAll(s))};
var naira=function(n){return '\u20A6'+Number(n).toLocaleString('en-NG')};
$$('#yr').forEach(function(e){e.textContent=new Date().getFullYear()});

/* cookie choice */
var bar=$('#cookieBar');
if(bar&&!localStorage.getItem('okv_cookie')){bar.classList.remove('d-none')}
function ck(v){localStorage.setItem('okv_cookie',v);bar.classList.add('d-none')}
if($('#ckYes'))$('#ckYes').onclick=function(){ck('all')};if($('#ckNo'))$('#ckNo').onclick=function(){ck('essential')};

document.addEventListener('click',function(e){var b=e.target.closest&&e.target.closest('[data-os]');if(b){var s=$('#osSel');if(s)s.value=b.getAttribute('data-os')}});

/* live config from the OKV Apps Script (price, bank details, announcement, switches) */
function applyConfig(cfg){
  if(!cfg)return;
  C.cfg=cfg;
  if(cfg.price){C.price=cfg.price;$$('[data-price]').forEach(function(e){e.textContent=naira(cfg.price)})}
  $$('[data-cfg]').forEach(function(e){var k=e.getAttribute('data-cfg'),v=cfg[k];
    if(k==='old_price'){if(v){e.textContent=naira(v);e.classList.remove('d-none')}else e.classList.add('d-none')}
    else if(k==='price'){if(v)e.textContent=naira(v)}
    else if(v)e.textContent=v});
  if(cfg.manual_enabled===false){$$('#payManual').forEach(function(b){b.classList.add('d-none')})}
  if(cfg.downloads)C.downloads=cfg.downloads;
  if(cfg.announcement){var a=$('#announce');if(a){a.textContent=cfg.announcement;a.classList.remove('d-none')}}
  if(cfg.download_url)C.downloadUrl=cfg.download_url;
  if(cfg.trial_enabled===false){$$('[data-bs-target="#trialModal"]').forEach(function(b){b.disabled=true;b.title='Trial downloads are paused'})}
  if(cfg.paystack_enabled===false){var p=$('#payPaystack');if(p)p.classList.add('d-none')}
  if(cfg.bank&&$('#bankDetails')){var b=cfg.bank;$('#bankDetails').innerHTML='<div><b>Bank:</b> '+esc(b.bank_name||'')+'</div><div><b>Account name:</b> '+esc(b.account_name||'')+'</div><div><b>Account number:</b> <span class="fs-5">'+esc(b.account_number||'')+'</span></div>'+(b.instructions?'<div class="small mt-1">'+esc(b.instructions)+'</div>':'')}
}
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})}
if(C.endpoint){fetch(C.endpoint+'?action=config').then(function(r){return r.json()}).then(function(j){if(j&&j.ok)applyConfig(j.config)}).catch(function(){})}

/* posting to Google Apps Script */
function post(payload){
  if(!C.endpoint)return Promise.resolve({ok:false,error:'This site is not connected to the OKV Apps Script yet. Set "endpoint" in assets/js/config.js. Meanwhile, WhatsApp +234 810 414 1138.'});
  return fetch(C.endpoint,{method:'POST',body:JSON.stringify(payload)}).then(function(r){return r.json()}).catch(function(){return {ok:false,error:'Network problem. Check your internet connection and try again.'}});
}
function readFile(f){return new Promise(function(ok,no){var r=new FileReader();r.onload=function(){ok({name:f.name,mime:f.type||'application/octet-stream',data:String(r.result).split(',')[1]})};r.onerror=no;r.readAsDataURL(f)})}
function showMsg(form,type,text){var m=$('.okv-msg',form);if(m)m.innerHTML='<div class="alert alert-'+type+'">'+text+'</div>'}

$$('form[data-okv]').forEach(function(form){
  form.addEventListener('submit',function(ev){
    ev.preventDefault();
    var kind=form.getAttribute('data-okv'),btn=$('button[type=submit]',form),old=btn.innerHTML;
    var data={};new FormData(form).forEach(function(v,k){if(!(v instanceof File))data[k]=v});
    data.type=kind==='paystack'?'paystack_init':kind;if(kind==='trial'&&$('#osSel'))data.os=$('#osSel').value;data.source='website';data.page=location.pathname.split('/').pop()||'index.html';
    data.consent=form.elements.consent&&form.elements.consent.checked?'yes':'';
    data.install_help=form.elements.install_help&&form.elements.install_help.checked?'yes':'';
    btn.disabled=true;btn.innerHTML='<span class="spinner-border spinner-border-sm"></span> Please wait...';showMsg(form,'info','Sending...');
    var file=form.elements.receipt&&form.elements.receipt.files[0],p=Promise.resolve();
    if(file){if(file.size>4*1024*1024){btn.disabled=false;btn.innerHTML=old;return showMsg(form,'danger','The receipt file is too large (max 4 MB).')}p=readFile(file).then(function(f){data.receipt=f})}
    p.then(function(){return post(data)}).then(function(res){
      btn.disabled=false;btn.innerHTML=old;
      if(!res||!res.ok)return showMsg(form,'danger',esc((res&&res.error)||'Something went wrong. Please try again or WhatsApp us.'));
      if(kind==='paystack'){if(res.authorization_url){showMsg(form,'success','Redirecting you to Paystack...');location.href=res.authorization_url}else showMsg(form,'danger','Could not start the payment. Please use bank transfer or WhatsApp us.');return}
      var done=$('.okv-done',form.parentNode);
      if(kind==='trial'){var url=res.download_url||(C.downloads&&C.downloads[data.os])||C.downloadUrl;$('#dlBtn').href=url;form.classList.add('d-none');done.classList.remove('d-none');setTimeout(function(){window.location.href=url},800);return}
      if(done){form.classList.add('d-none');done.classList.remove('d-none')}else showMsg(form,'success','Thank you! We have received your message.');
    });
  });
});

/* arriving from the software (Licence tab) and from Paystack */
var q=new URLSearchParams(location.search),ref=q.get('reference')||q.get('trxref'),ps=$('#payStatus');
var osEl=$('#osSel');if(osEl){var ua=navigator.userAgent;osEl.value=(/Mac/.test(ua)&&!/iPhone|iPad/.test(ua))?'mac_arm':'windows'}
if(q.get('install_id'))$$('input[name=install_id]').forEach(function(i){i.value=q.get('install_id')});
if(q.get('school'))$$('input[name=school]').forEach(function(i){if(!i.value)i.value=q.get('school')});
var fa=$('#fromApp');if(fa&&q.get('from')==='app')fa.innerHTML='<div class="alert alert-primary mb-0"><b>Welcome from OKV CBT.</b> Pay below with Paystack to get your activation code instantly on this page (and by email), or pay by bank transfer and upload your receipt. Then return to the software, open <b>Licence</b> and enter the code.'+(q.get('install_id')?' Your Install ID <code>'+esc(q.get('install_id'))+'</code> is already filled in.':'')+'</div>';
function codeBox(code,email){
  return '<div class="alert alert-success"><h4 class="mb-1"><i class="bi bi-check-circle-fill"></i> Payment confirmed. Thank you!</h4><p class="mb-2">Your activation code is ready. <b>Copy it now</b> - we have also emailed it to <b>'+esc(email||'you')+'</b> from the OKV CBT system (check spam if you do not see it).</p>'
   +'<div class="codebox"><span id="actCode">'+esc(code)+'</span><button type="button" class="btn btn-success" id="copyCode"><i class="bi bi-clipboard"></i> Copy code</button></div>'
   +'<ol class="mt-3 mb-0 small"><li>Open <b>OKV CBT</b> and sign in as admin.</li><li>Click <b>Licence</b> in the left menu.</li><li>Paste the code in the box and press <b>Activate</b>.</li></ol></div>';
}
function showCode(code,email){ps.innerHTML=codeBox(code,email);var b=$('#copyCode');if(b)b.onclick=function(){var t=$('#actCode').textContent;var done=function(){b.innerHTML='<i class="bi bi-check2"></i> Copied!'};if(navigator.clipboard)navigator.clipboard.writeText(t).then(done,function(){sel()});else sel();function sel(){var r=document.createRange();r.selectNode($('#actCode'));var s=getSelection();s.removeAllRanges();s.addRange(r);try{document.execCommand('copy');done()}catch(e){}}};ps.scrollIntoView({behavior:'smooth',block:'center'})}
if(ref&&ps){
  ps.innerHTML='<div class="alert alert-info"><span class="spinner-border spinner-border-sm"></span> Confirming your payment with Paystack...</div>';
  var tryVerify=function(n){
    if(!C.endpoint){ps.innerHTML='<div class="alert alert-warning">Payment received by Paystack (ref '+esc(ref)+'). This site is not connected to verify it - WhatsApp +234 810 414 1138.</div>';return}
    fetch(C.endpoint+'?action=verify&reference='+encodeURIComponent(ref)).then(function(r){return r.json()}).then(function(j){
      if(j&&j.ok&&j.paid&&j.code)return showCode(j.code,j.email);
      if(n<6)return setTimeout(function(){tryVerify(n+1)},3000);
      ps.innerHTML='<div class="alert alert-warning">We could not confirm this payment yet (reference '+esc(ref)+'). If you were charged, your code will be emailed automatically within minutes. Still nothing? WhatsApp +234 810 414 1138 with this reference.</div>';
    }).catch(function(){if(n<6)setTimeout(function(){tryVerify(n+1)},3000);else ps.innerHTML='<div class="alert alert-warning">Could not reach the server to confirm. Please WhatsApp +234 810 414 1138 with reference '+esc(ref)+'.</div>'});
  };
  tryVerify(0);
}

/* demo page */
$$('.dtab').forEach(function(b){b.addEventListener('click',function(){
  $$('.dtab').forEach(function(x){x.classList.remove('active')});b.classList.add('active');
  $$('.dpane').forEach(function(p){p.classList.remove('show')});$('#pane-'+b.getAttribute('data-tab')).classList.add('show');window.scrollTo({top:0,behavior:'smooth'})})});
$$('[data-demo]').forEach(function(b){b.addEventListener('click',function(){var t=$('#demoToast');if(t&&window.bootstrap)bootstrap.Toast.getOrCreateInstance(t,{delay:3500}).show()})});
})();
