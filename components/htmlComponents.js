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

const cryptoProPage = `
                        <!DOCTYPE html>
                        <html lang="ru">
                        <head>
                            <meta charset="UTF-8">
                            <meta name="viewport" content="width=device-width, initial-scale=1.0">
                            <title>Честный знак — Авторизация через ЭЦП</title>
                            <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
                            <style>
                                :root {
                                    --bg: #0a0e17; --surface: #111827; --surface-2: #1a2332;
                                    --border: #1e293b; --border-active: #3b82f6;
                                    --text: #e2e8f0; --text-dim: #64748b; --text-muted: #475569;
                                    --accent: #3b82f6; --accent-glow: rgba(59, 130, 246, 0.15);
                                    --success: #10b981; --success-glow: rgba(16, 185, 129, 0.15);
                                    --error: #ef4444; --error-glow: rgba(239, 68, 68, 0.15);
                                    --warning: #f59e0b;
                                    --mono: 'JetBrains Mono', monospace; --sans: 'Manrope', sans-serif;
                                }
                                * { margin: 0; padding: 0; box-sizing: border-box; }
                                body { font-family: var(--sans); background: var(--bg); color: var(--text); min-height: 100vh; display: flex; flex-direction: column; align-items: center; }
                                .container { position: relative; z-index: 1; width: 100%; max-width: 720px; padding: 40px 24px; }
                                header { text-align: center; margin-bottom: 48px; }
                                .logo-mark { display: inline-flex; align-items: center; gap: 10px; margin-bottom: 16px; }
                                .logo-icon { width: 36px; height: 36px; border-radius: 10px; background: linear-gradient(135deg, var(--accent), #6366f1); display: flex; align-items: center; justify-content: center; font-size: 18px; font-weight: 800; color: white; }
                                .logo-text { font-size: 15px; font-weight: 700; color: var(--text); }
                                h1 { font-size: 28px; font-weight: 800; letter-spacing: -0.8px; margin-bottom: 8px; background: linear-gradient(to right, var(--text), var(--text-dim)); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
                                .subtitle { color: var(--text-dim); font-size: 14px; line-height: 1.6; }
                                .step-flow { display: flex; flex-direction: column; gap: 16px; }
                                .step { background: var(--surface); border: 1px solid var(--border); border-radius: 16px; padding: 24px; transition: all 0.3s; position: relative; overflow: hidden; }
                                .step::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px; background: var(--border); transition: background 0.3s; }
                                .step.active { border-color: var(--border-active); } .step.active::before { background: var(--accent); }
                                .step.done { border-color: rgba(16, 185, 129, 0.3); } .step.done::before { background: var(--success); }
                                .step.error { border-color: rgba(239, 68, 68, 0.3); } .step.error::before { background: var(--error); }
                                .step.disabled { opacity: 0.4; pointer-events: none; }
                                .step-header { display: flex; align-items: center; gap: 14px; margin-bottom: 12px; }
                                .step-num { width: 32px; height: 32px; border-radius: 50%; background: var(--surface-2); border: 1.5px solid var(--border); display: flex; align-items: center; justify-content: center; font-size: 13px; font-weight: 700; font-family: var(--mono); flex-shrink: 0; transition: all 0.3s; }
                                .step.active .step-num { background: var(--accent-glow); border-color: var(--accent); color: var(--accent); }
                                .step.done .step-num { background: var(--success-glow); border-color: var(--success); color: var(--success); }
                                .step.error .step-num { background: var(--error-glow); border-color: var(--error); color: var(--error); }
                                .step-title { font-size: 16px; font-weight: 700; }
                                .step-desc { color: var(--text-dim); font-size: 13px; line-height: 1.6; margin-bottom: 16px; }
                                .btn { display: inline-flex; align-items: center; gap: 8px; padding: 10px 20px; border-radius: 10px; font-family: var(--sans); font-size: 13px; font-weight: 600; border: none; cursor: pointer; transition: all 0.2s; }
                                .btn-primary { background: var(--accent); color: white; }
                                .btn-primary:hover { background: #2563eb; transform: translateY(-1px); }
                                .btn-primary:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
                                .btn-ghost { background: var(--surface-2); color: var(--text-dim); border: 1px solid var(--border); }
                                .btn-ghost:hover { border-color: var(--text-muted); color: var(--text); }
                                .data-block { background: var(--bg); border: 1px solid var(--border); border-radius: 10px; padding: 16px; margin-top: 12px; font-family: var(--mono); font-size: 12px; line-height: 1.7; word-break: break-all; color: var(--text-dim); max-height: 200px; overflow-y: auto; }
                                .data-block .key { color: var(--accent); } .data-block .val { color: var(--text); }
                                .cert-list { display: flex; flex-direction: column; gap: 8px; margin-top: 12px; }
                                .cert-item { padding: 14px 16px; background: var(--bg); border: 1.5px solid var(--border); border-radius: 10px; cursor: pointer; transition: all 0.2s; }
                                .cert-item:hover { border-color: var(--text-muted); }
                                .cert-item.selected { border-color: var(--accent); background: var(--accent-glow); }
                                .cert-name { font-size: 13px; font-weight: 600; margin-bottom: 4px; }
                                .cert-meta { font-size: 11px; color: var(--text-dim); font-family: var(--mono); }
                                .cert-meta span + span::before { content: '·'; margin: 0 6px; color: var(--text-muted); }
                                .status { display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px; border-radius: 6px; font-size: 11px; font-weight: 600; font-family: var(--mono); }
                                .status-ok { background: var(--success-glow); color: var(--success); }
                                .status-err { background: var(--error-glow); color: var(--error); }
                                .status-warn { background: rgba(245, 158, 11, 0.15); color: var(--warning); }
                                .status-dot { width: 6px; height: 6px; border-radius: 50%; background: currentColor; }
                                .spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid var(--border); border-top-color: var(--accent); border-radius: 50%; animation: spin 0.7s linear infinite; }
                                @keyframes spin { to { transform: rotate(360deg); } }
                                .token-display { background: var(--bg); border: 1px solid rgba(16, 185, 129, 0.3); border-radius: 10px; padding: 16px; margin-top: 12px; }
                                .token-label { font-size: 11px; color: var(--success); font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
                                .token-value { font-family: var(--mono); font-size: 12px; color: var(--text); word-break: break-all; line-height: 1.7; }
                                .actions-row { display: flex; gap: 8px; margin-top: 12px; }
                                .log-area { margin-top: 32px; background: var(--surface); border: 1px solid var(--border); border-radius: 16px; overflow: hidden; }
                                .log-header { padding: 14px 20px; display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid var(--border); }
                                .log-title { font-size: 13px; font-weight: 700; color: var(--text-dim); }
                                .log-content { padding: 16px 20px; font-family: var(--mono); font-size: 11px; line-height: 1.8; color: var(--text-dim); max-height: 300px; overflow-y: auto; }
                                .log-line { margin-bottom: 2px; } .log-time { color: var(--text-muted); }
                                .log-info { color: var(--accent); } .log-ok { color: var(--success); } .log-err { color: var(--error); }
                            </style>
                        </head>
                        <body>
                            <div class="container">
                                <header>
                                    <div class="logo-mark">
                                        <div class="logo-icon">ЧЗ</div>
                                        <div class="logo-text">Честный знак</div>
                                    </div>
                                    <h1>Авторизация через ЭЦП</h1>
                                    <p class="subtitle">Получение токена через True API с подписанием CAdES-BES</p>
                                </header>
                                <div class="step-flow">
                                    <div class="step active" id="step-plugin">
                                        <div class="step-header"><div class="step-num">0</div><div class="step-title">Проверка КриптоПро Browser plug-in</div></div>
                                        <div class="step-desc">Проверяем наличие и работоспособность плагина в браузере.</div>
                                        <div id="plugin-status"></div>
                                        <button class="btn btn-primary" id="btn-check-plugin" onclick="checkPlugin()">Проверить плагин</button>
                                    </div>
                                    <div class="step disabled" id="step-auth-key">
                                        <div class="step-header"><div class="step-num">1</div><div class="step-title">Получение ключа авторизации</div></div>
                                        <div class="step-desc">GET-запрос к <code style="color:var(--accent);">/api/v3/true-api/auth/key</code> — получаем <code>uuid</code> и <code>data</code>.</div>
                                        <button class="btn btn-primary" id="btn-get-key" onclick="getAuthKey()">Получить ключ</button>
                                        <div id="auth-key-result"></div>
                                    </div>
                                    <div class="step disabled" id="step-cert">
                                        <div class="step-header"><div class="step-num">2</div><div class="step-title">Выбор сертификата</div></div>
                                        <div class="step-desc">Выберите сертификат УКЭП, привязанный к организации в Честном знаке.</div>
                                        <div id="cert-list-container"></div>
                                    </div>
                                    <div class="step disabled" id="step-sign">
                                        <div class="step-header"><div class="step-num">3</div><div class="step-title">Подписание данных</div></div>
                                        <div class="step-desc">Подписываем строку <code>data</code> в формате CAdES-BES (attached CMS/PKCS#7).</div>
                                        <button class="btn btn-primary" id="btn-sign" onclick="signData()">Подписать</button>
                                        <div id="sign-result"></div>
                                    </div>
                                    <div class="step disabled" id="step-token">
                                        <div class="step-header"><div class="step-num">4</div><div class="step-title">Получение токена</div></div>
                                        <div class="step-desc">POST-запрос к <code style="color:var(--accent);">/api/v3/true-api/auth/simpleSignIn</code> — отправляем uuid и подпись.</div>
                                        <button class="btn btn-primary" id="btn-get-token" onclick="getToken()">Получить токен</button>
                                        <div id="token-result"></div>
                                    </div>
                                </div>
                                <div class="log-area">
                                    <div class="log-header">
                                        <span class="log-title">Лог операций</span>
                                        <button class="btn btn-ghost" style="padding:4px 10px;font-size:11px;" onclick="clearLog()">Очистить</button>
                                    </div>
                                    <div class="log-content" id="log"></div>
                                </div>
                            </div>

                            <script src="cadesplugin_api.js"></script>

                            <script>
                            // ================================================================
                            // Константы CAdES
                            // ================================================================
                            var STORE_CU = 2, STORE_MY = "My", STORE_MAX = 2;
                            var FIND_SHA1 = 0, CHAIN = 1, CADES_BES = 1, B64BIN = 1;

                            var CFG = {
                                base: 'https://markirovka.crpt.ru',
                                keyPath: '/api/v3/true-api/auth/key',
                                signPath: '/api/v3/true-api/auth/simpleSignIn',
                            };
                            var app = {};

                            // ================================================================
                            // Утилиты
                            // ================================================================
                            function log(m, t) {
                                t = t || 'info';
                                var el = document.getElementById('log');
                                var ts = new Date().toLocaleTimeString('ru-RU');
                                var c = t === 'ok' ? 'log-ok' : t === 'err' ? 'log-err' : 'log-info';
                                el.innerHTML += '<div class="log-line"><span class="log-time">[' + ts + ']</span> <span class="' + c + '">' + m + '</span></div>';
                                el.scrollTop = el.scrollHeight;
                            }
                            function clearLog() { document.getElementById('log').innerHTML = ''; }
                            function setStep(id, s) {
                                var el = document.getElementById(id);
                                el.className = 'step' + (s ? ' ' + s : '');
                            }
                            function b64(s) { return btoa(unescape(encodeURIComponent(s))); }

                            // ================================================================
                            // Шаг 0: Проверка
                            // НЕ используем cadesplugin.then() — он ломается.
                            // Просто пробуем CreateObjectAsync напрямую.
                            // ================================================================
                            async function checkPlugin() {
                                var btn = document.getElementById('btn-check-plugin');
                                var st = document.getElementById('plugin-status');
                                btn.disabled = true;
                                btn.innerHTML = '<span class="spinner"></span> Проверка...';
                                log('Проверяем КриптоПро Browser plug-in...');

                                try {
                                    // Напрямую пробуем создать объект Store
                                    var oStore = await cadesplugin.CreateObjectAsync("CAdESCOM.Store");
                                    await oStore.Open(STORE_CU, STORE_MY, STORE_MAX);
                                    await oStore.Close();

                                    log('Плагин работает!', 'ok');
                                    st.innerHTML = '<span class="status status-ok"><span class="status-dot"></span>Плагин работает</span>';
                                    setStep('step-plugin', 'done');
                                    setStep('step-auth-key', 'active');
                                    btn.innerHTML = '✓ Плагин найден';
                                } catch(e) {
                                    var m = (e && e.message) ? e.message : String(e);
                                    log('Ошибка: ' + m, 'err');
                                    st.innerHTML =
                                        '<span class="status status-err"><span class="status-dot"></span>' + m + '</span>' +
                                        '<div class="step-desc" style="margin-top:8px;">Установите <a href="https://www.cryptopro.ru/products/cades/plugin" target="_blank" style="color:var(--accent);">КриптоПро ЭЦП Browser plug-in</a></div>';
                                    setStep('step-plugin', 'error');
                                    btn.innerHTML = 'Повторить';
                                    btn.disabled = false;
                                }
                            }

                            // ================================================================
                            // Шаг 1: Ключ
                            // ================================================================
                            async function getAuthKey() {
                                var btn = document.getElementById('btn-get-key');
                                var res = document.getElementById('auth-key-result');
                                btn.disabled = true;
                                btn.innerHTML = '<span class="spinner"></span> Запрос...';
                                var url = CFG.base + CFG.keyPath;
                                log('GET ' + url);
                                try {
                                    var r = await fetch(url, { headers: { 'Accept': 'application/json' } });
                                    if (!r.ok) throw new Error('HTTP ' + r.status);
                                    var j = await r.json();
                                    app.uuid = j.uuid; app.data = j.data;
                                    log('uuid: ' + app.uuid, 'ok');
                                    log('data: ' + app.data.substring(0, 40) + '...', 'ok');
                                    res.innerHTML = '<div class="data-block"><div><span class="key">uuid:</span> <span class="val">' + app.uuid + '</span></div><div><span class="key">data:</span> <span class="val">' + app.data + '</span></div></div>';
                                    setStep('step-auth-key', 'done');
                                    btn.innerHTML = '✓ Получен';
                                    loadCerts();
                                } catch(e) {
                                    log('Ошибка: ' + e.message, 'err');
                                    res.innerHTML = '<span class="status status-err"><span class="status-dot"></span>' + e.message + '</span>';
                                    setStep('step-auth-key', 'error');
                                    btn.innerHTML = 'Повторить'; btn.disabled = false;
                                }
                            }

                            // ================================================================
                            // Шаг 2: Сертификаты
                            // ================================================================
                            async function loadCerts() {
                                setStep('step-cert', 'active');
                                var c = document.getElementById('cert-list-container');
                                c.innerHTML = '<div style="padding:8px 0;"><span class="spinner"></span> <span style="color:var(--text-dim);font-size:13px;margin-left:8px;">Загрузка...</span></div>';
                                log('Загружаем сертификаты...');
                                try {
                                    var oStore = await cadesplugin.CreateObjectAsync("CAdESCOM.Store");
                                    await oStore.Open(STORE_CU, STORE_MY, STORE_MAX);
                                    var oCerts = await oStore.Certificates;
                                    var count = await oCerts.Count;
                                    log('Найдено: ' + count, 'ok');
                                    if (count === 0) { c.innerHTML = '<span class="status status-warn"><span class="status-dot"></span>Нет сертификатов</span>'; setStep('step-cert', 'error'); return; }

                                    var h = '<div class="cert-list">';
                                    var now = new Date();
                                    for (var i = 1; i <= count; i++) {
                                        var cert = await oCerts.Item(i);
                                        var sn = '', tp = '', vt = '';
                                        try { sn = String(await cert.SubjectName); } catch(x) {}
                                        try { tp = String(await cert.Thumbprint); } catch(x) {}
                                        try { vt = String(await cert.ValidToDate); } catch(x) {}
                                        var d = vt ? new Date(vt) : new Date(0);
                                        var cn = (sn.match(/CN=([^,]+)/) || [])[1] || sn.substring(0, 50) || 'Серт #' + i;
                                        var inn = (sn.match(/(?:INN|ИНН|1\.2\.643\.3\.131\.1\.1)=(\d+)/i) || [])[1] || null;
                                        var exp = d < now;
                                        h += '<div class="cert-item" onclick="selectCert(this,\'' + tp + '\')">' +
                                            '<div class="cert-name">' + cn + '</div><div class="cert-meta">' +
                                            '<span>' + tp.substring(0, 16) + '…</span>' +
                                            (inn ? '<span>ИНН: ' + inn + '</span>' : '') +
                                            '<span>до ' + d.toLocaleDateString('ru-RU') + '</span>' +
                                            '<span class="status ' + (exp ? 'status-err' : 'status-ok') + '" style="display:inline-flex;padding:2px 6px;font-size:10px;"><span class="status-dot"></span>' + (exp ? 'Истёк' : 'OK') + '</span>' +
                                            '</div></div>';
                                    }
                                    h += '</div>';
                                    c.innerHTML = h;
                                    await oStore.Close();
                                } catch(e) {
                                    var m = (e && e.message) ? e.message : String(e);
                                    log('Ошибка: ' + m, 'err');
                                    c.innerHTML = '<span class="status status-err"><span class="status-dot"></span>' + m + '</span>';
                                    setStep('step-cert', 'error');
                                }
                            }

                            function selectCert(el, tp) {
                                document.querySelectorAll('.cert-item.selected').forEach(function(x) { x.classList.remove('selected'); });
                                el.classList.add('selected');
                                app.tp = tp;
                                log('Выбран: ' + tp, 'ok');
                                setStep('step-cert', 'done');
                                setStep('step-sign', 'active');
                            }

                            // ================================================================
                            // Шаг 3: Подписание
                            // ================================================================
                            async function signData() {
                                var btn = document.getElementById('btn-sign');
                                var res = document.getElementById('sign-result');
                                btn.disabled = true;
                                btn.innerHTML = '<span class="spinner"></span> Подписание...';
                                log('Подписываем...');
                                try {
                                    var oStore = await cadesplugin.CreateObjectAsync("CAdESCOM.Store");
                                    await oStore.Open(STORE_CU, STORE_MY, STORE_MAX);
                                    var all = await oStore.Certificates;
                                    var found = await all.Find(FIND_SHA1, app.tp);
                                    var n = await found.Count;
                                    if (n === 0) throw new Error('Сертификат не найден');
                                    var cert = await found.Item(1);
                                    log('Сертификат найден', 'ok');

                                    var signer = await cadesplugin.CreateObjectAsync("CAdESCOM.CPSigner");
                                    await signer.propset_Certificate(cert);
                                    await signer.propset_Options(CHAIN);

                                    var sd = await cadesplugin.CreateObjectAsync("CAdESCOM.CadesSignedData");
                                    await sd.propset_ContentEncoding(B64BIN);
                                    await sd.propset_Content(b64(app.data));

                                    var sig = await sd.SignCades(signer, CADES_BES, false);
                                    app.sig = String(sig).replace(/[\r\n\s]/g, '');

                                    log('Подпись OK! Длина: ' + app.sig.length, 'ok');
                                    res.innerHTML = '<div class="data-block"><span class="key">signature:</span><br><span class="val">' + app.sig.substring(0, 120) + '...</span></div>';
                                    await oStore.Close();
                                    setStep('step-sign', 'done');
                                    setStep('step-token', 'active');
                                    btn.innerHTML = '✓ Подписано';
                                } catch(e) {
                                    var m = (e && e.message) ? e.message : String(e);
                                    log('Ошибка: ' + m, 'err');
                                    res.innerHTML = '<span class="status status-err"><span class="status-dot"></span>' + m + '</span>';
                                    setStep('step-sign', 'error');
                                    btn.innerHTML = 'Повторить'; btn.disabled = false;
                                }
                            }

                            // ================================================================
                            // Шаг 4: Токен
                            // ================================================================
                            async function getToken() {
                                var btn = document.getElementById('btn-get-token');
                                var res = document.getElementById('token-result');
                                btn.disabled = true;
                                btn.innerHTML = '<span class="spinner"></span> Запрос...';
                                var url = CFG.base + CFG.signPath;
                                log('POST ' + url);
                                try {
                                    var r = await fetch(url, {
                                        method: 'POST',
                                        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
                                        body: JSON.stringify({ uuid: app.uuid, data: app.sig })
                                    });
                                    var j = await r.json();
                                    if (!r.ok) throw new Error(j.error_message || j.message || 'HTTP ' + r.status);
                                    app.token = j.token;
                                    log('Токен получен!', 'ok');
                                    res.innerHTML =
                                        '<div class="token-display"><div class="token-label">✓ Токен авторизации</div><div class="token-value">' + app.token + '</div></div>' +
                                        '<div class="actions-row"><button class="btn btn-ghost" onclick="copyToken()">Копировать</button></div>';
                                    setStep('step-token', 'done');
                                    btn.innerHTML = '✓ Получен';
                                } catch(e) {
                                    log('Ошибка: ' + e.message, 'err');
                                    res.innerHTML = '<span class="status status-err"><span class="status-dot"></span>' + e.message + '</span><div class="step-desc" style="margin-top:8px;">uuid живёт ~10 мин. Начните с шага 1.</div>';
                                    setStep('step-token', 'error');
                                    btn.innerHTML = 'Повторить'; btn.disabled = false;
                                }
                            }

                            function copyToken() {
                                if (app.token) navigator.clipboard.writeText(app.token).then(function() { log('Скопировано', 'ok'); });
                            }
                            </script>
                        </body>
                        </html>`

module.exports = {

    headerComponent,
    navComponent,
    footerComponent,
    cryptoProPage

}