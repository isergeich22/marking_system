const headerComponent = `<!DOCTYPE html>
                            <html lang="en">
                            <head>
                                <meta charset="UTF-8">
                                <meta http-equiv="X-UA-Compatible" content="IE=edge">
                                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                                <link rel="stylesheet" href="/css/styles.css" type="text/css">
                                <link rel="shortcut icon" type="image/png" href="/favicon.png">
                                <link
                                    rel="stylesheet"
                                    href="https://cdn.jsdelivr.net/npm/bulma@1.0.4/css/bulma.min.css"
                                >`

const navComponent = `<header class="header">
                        <nav>
                            <img src="/img/chestnyj_znak.png" alt="честный знак">
                            <p class="nav-item" id="home"><a href="http://localhost:3030/home">Главная</a></p>
                            <p class="nav-item" id="import">Создание импорт-файлов</p>
                            <p class="nav-item" id="cis_actions">Действия с КИЗ</p>
                            <p class="nav-item" id="other_actions">Другие действия</p>
                        </nav>                    
                    </header>`

const footerComponent = `   <button id="top" class="button-top">
                            <svg width="24" height="24" fill="none" xmlns="http://www.w3.org/2000/svg">
                                <g clip-path="url(#ArrowLongUp_large_svg__clip0_35331_5070)">
                                    <path d="M12 2v20m0-20l7 6.364M12 2L5 8.364" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"></path>
                                </g><defs><clipPath id="ArrowLongUp_large_svg__clip0_35331_5070"><path fill="#fff" transform="rotate(90 12 12)" d="M0 0h24v24H0z">
                                </path></clipPath></defs></svg>
                            </button>    
                            <script src="/script.js"></script>
                            </body>
                        </html>`

const cryptoProPage = `<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Честный знак — Авторизация через ЭЦП</title>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--bg:#0a0e17;--surface:#111827;--surface-2:#1a2332;--border:#1e293b;--border-active:#3b82f6;--text:#e2e8f0;--text-dim:#64748b;--text-muted:#475569;--accent:#3b82f6;--accent-glow:rgba(59,130,246,.15);--success:#10b981;--success-glow:rgba(16,185,129,.15);--error:#ef4444;--error-glow:rgba(239,68,68,.15);--warning:#f59e0b;--mono:'JetBrains Mono',monospace;--sans:'Manrope',sans-serif}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh;display:flex;flex-direction:column;align-items:center}
.container{position:relative;z-index:1;width:100%;max-width:720px;padding:40px 24px}
header{text-align:center;margin-bottom:48px}
.logo-mark{display:inline-flex;align-items:center;gap:10px;margin-bottom:16px}
.logo-icon{width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,var(--accent),#6366f1);display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:800;color:#fff}
.logo-text{font-size:15px;font-weight:700;color:var(--text)}
h1{font-size:28px;font-weight:800;letter-spacing:-.8px;margin-bottom:8px;background:linear-gradient(to right,var(--text),var(--text-dim));-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.subtitle{color:var(--text-dim);font-size:14px;line-height:1.6}
.step-flow{display:flex;flex-direction:column;gap:16px}
.step{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:24px;transition:all .3s;position:relative;overflow:hidden}
.step::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--border);transition:background .3s}
.step.active{border-color:var(--border-active)}.step.active::before{background:var(--accent)}
.step.done{border-color:rgba(16,185,129,.3)}.step.done::before{background:var(--success)}
.step.error{border-color:rgba(239,68,68,.3)}.step.error::before{background:var(--error)}
.step.disabled{opacity:.4;pointer-events:none}
.step-header{display:flex;align-items:center;gap:14px;margin-bottom:12px}
.step-num{width:32px;height:32px;border-radius:50%;background:var(--surface-2);border:1.5px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;font-family:var(--mono);flex-shrink:0;transition:all .3s}
.step.active .step-num{background:var(--accent-glow);border-color:var(--accent);color:var(--accent)}
.step.done .step-num{background:var(--success-glow);border-color:var(--success);color:var(--success)}
.step.error .step-num{background:var(--error-glow);border-color:var(--error);color:var(--error)}
.step-title{font-size:16px;font-weight:700}
.step-desc{color:var(--text-dim);font-size:13px;line-height:1.6;margin-bottom:16px}
.btn{display:inline-flex;align-items:center;gap:8px;padding:10px 20px;border-radius:10px;font-family:var(--sans);font-size:13px;font-weight:600;border:none;cursor:pointer;transition:all .2s}
.btn-primary{background:var(--accent);color:#fff}.btn-primary:hover{background:#2563eb;transform:translateY(-1px)}.btn-primary:disabled{opacity:.5;cursor:not-allowed;transform:none}
.btn-ghost{background:var(--surface-2);color:var(--text-dim);border:1px solid var(--border)}.btn-ghost:hover{border-color:var(--text-muted);color:var(--text)}
.data-block{background:var(--bg);border:1px solid var(--border);border-radius:10px;padding:16px;margin-top:12px;font-family:var(--mono);font-size:12px;line-height:1.7;word-break:break-all;color:var(--text-dim);max-height:200px;overflow-y:auto}
.cert-list{display:flex;flex-direction:column;gap:8px;margin-top:12px}
.cert-item{padding:14px 16px;background:var(--bg);border:1.5px solid var(--border);border-radius:10px;cursor:pointer;transition:all .2s}
.cert-item:hover{border-color:var(--text-muted)}.cert-item.selected{border-color:var(--accent);background:var(--accent-glow)}
.cert-name{font-size:13px;font-weight:600;margin-bottom:4px}
.cert-meta{font-size:11px;color:var(--text-dim);font-family:var(--mono);display:flex;flex-wrap:wrap;gap:4px 12px;align-items:center}
.status{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:6px;font-size:11px;font-weight:600;font-family:var(--mono)}
.status-ok{background:var(--success-glow);color:var(--success)}.status-err{background:var(--error-glow);color:var(--error)}.status-warn{background:rgba(245,158,11,.15);color:var(--warning)}
.status-dot{width:6px;height:6px;border-radius:50%;background:currentColor}
.spinner{display:inline-block;width:16px;height:16px;border:2px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.token-display{background:var(--bg);border:1px solid rgba(16,185,129,.3);border-radius:10px;padding:16px;margin-top:12px}
.token-label{font-size:11px;color:var(--success);font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
.token-value{font-family:var(--mono);font-size:12px;color:var(--text);word-break:break-all;line-height:1.7}
.actions-row{display:flex;gap:8px;margin-top:12px}
.log-area{margin-top:32px;background:var(--surface);border:1px solid var(--border);border-radius:16px;overflow:hidden}
.log-header{padding:14px 20px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--border)}
.log-title{font-size:13px;font-weight:700;color:var(--text-dim)}
.log-content{padding:16px 20px;font-family:var(--mono);font-size:11px;line-height:1.8;color:var(--text-dim);max-height:300px;overflow-y:auto}
.log-line{margin-bottom:2px}.log-time{color:var(--text-muted)}.log-info{color:var(--accent)}.log-ok{color:var(--success)}.log-err{color:var(--error)}
</style>
</head>
<body>
<div class="container">
<header>
<div class="logo-mark"><div class="logo-icon">ЧЗ</div><div class="logo-text">Честный знак</div></div>
<h1>Авторизация через ЭЦП</h1>
<p class="subtitle">Получение токена через True API с подписанием CAdES-BES</p>
</header>
<div class="step-flow">
<div class="step active" id="step-plugin"><div class="step-header"><div class="step-num">0</div><div class="step-title">Проверка КриптоПро Browser plug-in</div></div><div class="step-desc">Проверяем наличие и работоспособность плагина.</div><div id="plugin-status"></div><button class="btn btn-primary" id="btn-check-plugin" onclick="checkPlugin()">Проверить плагин</button></div>
<div class="step disabled" id="step-auth-key"><div class="step-header"><div class="step-num">1</div><div class="step-title">Получение ключа авторизации</div></div><div class="step-desc">GET /api/v3/true-api/auth/key</div><button class="btn btn-primary" id="btn-get-key" onclick="getAuthKey()">Получить ключ</button><div id="auth-key-result"></div></div>
<div class="step disabled" id="step-cert"><div class="step-header"><div class="step-num">2</div><div class="step-title">Выбор сертификата</div></div><div class="step-desc">Выберите сертификат УКЭП.</div><div id="cert-list-container"></div></div>
<div class="step disabled" id="step-sign"><div class="step-header"><div class="step-num">3</div><div class="step-title">Подписание данных</div></div><div class="step-desc">CAdES-BES attached.</div><button class="btn btn-primary" id="btn-sign" onclick="signData()">Подписать</button><div id="sign-result"></div></div>
<div class="step disabled" id="step-token"><div class="step-header"><div class="step-num">4</div><div class="step-title">Получение токена</div></div><div class="step-desc">POST /api/v3/true-api/auth/simpleSignIn</div><button class="btn btn-primary" id="btn-get-token" onclick="getToken()">Получить токен</button><div id="token-result"></div></div>
</div>
<div class="log-area"><div class="log-header"><span class="log-title">Лог</span><button class="btn btn-ghost" style="padding:4px 10px;font-size:11px;" onclick="clearLog()">Очистить</button></div><div class="log-content" id="log"></div></div>
</div>
<script src="cadesplugin_api.js"></script>
<script>
var STORE_CU=2,STORE_MY="My",STORE_MAX=2,FIND_SHA1=0,CHAIN=1,CADES_BES=1,B64BIN=1;
var CFG={base:"",keyPath:"/api/v3/true-api/auth/key",signPath:"/api/v3/true-api/auth/simpleSignIn"};
var app={};
function log(m,t){var el=document.getElementById("log");var ts=new Date().toLocaleTimeString("ru-RU");var c=(t==="ok")?"log-ok":(t==="err")?"log-err":"log-info";var line=document.createElement("div");line.className="log-line";var a=document.createElement("span");a.className="log-time";a.textContent="["+ts+"] ";var b=document.createElement("span");b.className=c;b.textContent=m;line.appendChild(a);line.appendChild(b);el.appendChild(line);el.scrollTop=el.scrollHeight;}
function clearLog(){document.getElementById("log").innerHTML="";}
function setStep(id,s){document.getElementById(id).className="step"+(s?" "+s:"");}
function b64(s){return btoa(unescape(encodeURIComponent(s)));}
function cleanSig(str){var out="";var i=0;var len=str.length;while(i<len){var code=str.charCodeAt(i);if(code!==10&&code!==13&&code!==32){out+=str.charAt(i);}i++;}return out;}
function makeSpinner(){var s=document.createElement("span");s.className="spinner";return s;}
function makeStatus(text,type){var sp=document.createElement("span");sp.className="status status-"+type;var dot=document.createElement("span");dot.className="status-dot";sp.appendChild(dot);sp.appendChild(document.createTextNode(text));return sp;}
function setBtnLoading(btn,text){btn.disabled=true;btn.innerHTML="";btn.appendChild(makeSpinner());btn.appendChild(document.createTextNode(" "+text));}
async function checkPlugin(){var btn=document.getElementById("btn-check-plugin");var st=document.getElementById("plugin-status");setBtnLoading(btn,"Проверка...");log("Проверяем КриптоПро...");try{var oStore=await cadesplugin.CreateObjectAsync("CAdESCOM.Store");await oStore.Open(STORE_CU,STORE_MY,STORE_MAX);await oStore.Close();log("Плагин работает!","ok");st.innerHTML="";st.appendChild(makeStatus("Плагин работает","ok"));setStep("step-plugin","done");setStep("step-auth-key","active");btn.textContent="\u2713 Плагин найден";}catch(e){var m=(e&&e.message)?e.message:String(e);log("Ошибка: "+m,"err");st.innerHTML="";st.appendChild(makeStatus(m,"err"));setStep("step-plugin","error");btn.textContent="Повторить";btn.disabled=false;}}
async function getAuthKey(){var btn=document.getElementById("btn-get-key");var res=document.getElementById("auth-key-result");setBtnLoading(btn,"Запрос...");var url=CFG.base+CFG.keyPath;log("GET "+url);try{var r=await fetch(url,{headers:{"Accept":"application/json"}});if(!r.ok)throw new Error("HTTP "+r.status);var j=await r.json();app.uuid=j.uuid;app.data=j.data;log("uuid: "+app.uuid,"ok");log("data: "+app.data.substring(0,40)+"...","ok");var block=document.createElement("div");block.className="data-block";var d1=document.createElement("div");d1.textContent="uuid: "+app.uuid;var d2=document.createElement("div");d2.textContent="data: "+app.data;block.appendChild(d1);block.appendChild(d2);res.innerHTML="";res.appendChild(block);setStep("step-auth-key","done");btn.textContent="\u2713 Получен";loadCerts();}catch(e){log("Ошибка: "+e.message,"err");res.innerHTML="";res.appendChild(makeStatus(e.message,"err"));setStep("step-auth-key","error");btn.textContent="Повторить";btn.disabled=false;}}
async function loadCerts(){setStep("step-cert","active");var container=document.getElementById("cert-list-container");container.innerHTML="";container.appendChild(makeSpinner());container.appendChild(document.createTextNode(" Загрузка..."));log("Загружаем сертификаты...");try{var oStore=await cadesplugin.CreateObjectAsync("CAdESCOM.Store");await oStore.Open(STORE_CU,STORE_MY,STORE_MAX);var oCerts=await oStore.Certificates;var count=await oCerts.Count;log("Найдено: "+count,"ok");if(count===0){container.innerHTML="";container.appendChild(makeStatus("Нет сертификатов","warn"));setStep("step-cert","error");return;}var listDiv=document.createElement("div");listDiv.className="cert-list";var now=new Date();for(var i=1;i<=count;i++){var cert=await oCerts.Item(i);var sn="",tp="",vt="";try{sn=String(await cert.SubjectName);}catch(x){}try{tp=String(await cert.Thumbprint);}catch(x){}try{vt=String(await cert.ValidToDate);}catch(x){}var validTo=vt?new Date(vt):new Date(0);var cnArr=sn.match(/CN=([^,]+)/);var cn=cnArr?cnArr[1]:(sn.substring(0,50)||"Cert #"+i);var innArr=sn.match(/INN=(\d+)/i);if(!innArr)innArr=sn.match(/1\.2\.643\.3\.131\.1\.1=(\d+)/);var inn=innArr?innArr[1]:null;var expired=validTo<now;var item=document.createElement("div");item.className="cert-item";var nameDiv=document.createElement("div");nameDiv.className="cert-name";nameDiv.textContent=cn;item.appendChild(nameDiv);var metaDiv=document.createElement("div");metaDiv.className="cert-meta";var s1=document.createElement("span");s1.textContent=tp.substring(0,16)+"\u2026";metaDiv.appendChild(s1);if(inn){var s2=document.createElement("span");s2.textContent="\u0418\u041D\u041D: "+inn;metaDiv.appendChild(s2);}var s3=document.createElement("span");s3.textContent="\u0434\u043E "+validTo.toLocaleDateString("ru-RU");metaDiv.appendChild(s3);metaDiv.appendChild(makeStatus(expired?"\u0418\u0441\u0442\u0451\u043A":"OK",expired?"err":"ok"));item.appendChild(metaDiv);(function(element,thumbprint){element.addEventListener("click",function(){document.querySelectorAll(".cert-item.selected").forEach(function(x){x.classList.remove("selected");});element.classList.add("selected");app.tp=thumbprint;log("\u0412\u044B\u0431\u0440\u0430\u043D: "+thumbprint,"ok");setStep("step-cert","done");setStep("step-sign","active");});})(item,tp);listDiv.appendChild(item);}container.innerHTML="";container.appendChild(listDiv);await oStore.Close();}catch(e){var m=(e&&e.message)?e.message:String(e);log("Ошибка: "+m,"err");container.innerHTML="";container.appendChild(makeStatus(m,"err"));setStep("step-cert","error");}}
async function signData(){var btn=document.getElementById("btn-sign");var res=document.getElementById("sign-result");setBtnLoading(btn,"Подписание...");log("Подписываем...");try{var oStore=await cadesplugin.CreateObjectAsync("CAdESCOM.Store");await oStore.Open(STORE_CU,STORE_MY,STORE_MAX);var allCerts=await oStore.Certificates;var found=await allCerts.Find(FIND_SHA1,app.tp);var n=await found.Count;if(n===0)throw new Error("Сертификат не найден");var cert=await found.Item(1);log("Сертификат найден","ok");var signer=await cadesplugin.CreateObjectAsync("CAdESCOM.CPSigner");await signer.propset_Certificate(cert);await signer.propset_Options(CHAIN);var sd=await cadesplugin.CreateObjectAsync("CAdESCOM.CadesSignedData");await sd.propset_ContentEncoding(B64BIN);await sd.propset_Content(b64(app.data));var sig=await sd.SignCades(signer,CADES_BES,false);app.sig=cleanSig(String(sig));log("Подпись OK! Длина: "+app.sig.length,"ok");var block=document.createElement("div");block.className="data-block";block.textContent="signature: "+app.sig.substring(0,120)+"...";res.innerHTML="";res.appendChild(block);await oStore.Close();setStep("step-sign","done");setStep("step-token","active");btn.textContent="\u2713 Подписано";}catch(e){var m=(e&&e.message)?e.message:String(e);log("Ошибка: "+m,"err");res.innerHTML="";res.appendChild(makeStatus(m,"err"));setStep("step-sign","error");btn.textContent="Повторить";btn.disabled=false;}}
async function getToken(){var btn=document.getElementById("btn-get-token");var res=document.getElementById("token-result");setBtnLoading(btn,"Запрос...");var url=CFG.base+CFG.signPath;log("POST "+url);try{var r=await fetch(url,{method:"POST",headers:{"Content-Type":"application/json","Accept":"application/json"},body:JSON.stringify({uuid:app.uuid,data:app.sig})});var j=await r.json();if(!r.ok)throw new Error(j.error_message||j.message||"HTTP "+r.status);app.token=j.token;log("Токен получен!","ok");var wrap=document.createElement("div");wrap.className="token-display";var label=document.createElement("div");label.className="token-label";label.textContent="\u2713 Токен авторизации";var val=document.createElement("div");val.className="token-value";val.textContent=app.token;wrap.appendChild(label);wrap.appendChild(val);var actions=document.createElement("div");actions.className="actions-row";var copyBtn=document.createElement("button");copyBtn.className="btn btn-ghost";copyBtn.textContent="Копировать";copyBtn.addEventListener("click",function(){navigator.clipboard.writeText(app.token).then(function(){log("Скопировано","ok");});});actions.appendChild(copyBtn);res.innerHTML="";res.appendChild(wrap);res.appendChild(actions);setStep("step-token","done");btn.textContent="\u2713 Получен";}catch(e){log("Ошибка: "+e.message,"err");res.innerHTML="";res.appendChild(makeStatus(e.message,"err"));setStep("step-token","error");btn.textContent="Повторить";btn.disabled=false;}}
</script>
</body>
</html>`

module.exports = {

    headerComponent,
    navComponent,
    footerComponent,
    cryptoProPage

}