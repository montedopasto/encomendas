/*
 * Encomendas — camada de plataforma.
 *
 * A app Encomendas foi feita para correr dentro do Claude (claude.ai), que lhe dá:
 *   - uma base de dados partilhada  -> claude.use('db')
 *   - o utilizador atual           -> claude.use('user')
 *   - descarregar ficheiros        -> claude.use('downloads')
 *   - conectores (Microsoft 365)   -> claude.use('mcp')
 *
 * Este ficheiro implementa a mesma interface (window.claude) fora do Claude, só com o Microsoft 365:
 *   - entrada com a conta Microsoft da empresa (MSAL.js, Microsoft Entra ID);
 *   - base de dados = ficheiros JSON numa pasta do SharePoint (pasta "dados" dentro de SalesVision).
 *     Quem tem acesso à pasta SalesVision usa a app; quem não tem, não entra;
 *   - email, Outlook e SharePoint pelo Microsoft Graph, com as permissões de quem entra.
 *
 * Assim o código da app (index.html) é o mesmo da versão do Claude.
 */
(function () {
  'use strict';
  const CFG = window.ENCOMENDAS_CONFIG || {};
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const SCOPES = ['User.Read', 'Mail.ReadWrite', 'Mail.Send', 'Files.ReadWrite.All'];
  const POLL_MS = CFG.pollMs || 15000;

  // ------------------------------------------------------------------ ecrã de entrada / erros
  function gate(html) {
    let el = document.getElementById('sv-gate');
    if (!el) {
      el = document.createElement('div'); el.id = 'sv-gate';
      el.style.cssText = 'position:fixed;inset:0;z-index:50;display:grid;place-items:center;padding:16px;background:var(--bg,#F4F2ED);font-family:Montserrat,Segoe UI,Arial,sans-serif';
      (document.body || document.documentElement).appendChild(el);
    }
    el.innerHTML = '<div style="max-width:440px;width:100%;background:var(--surface,#fff);border:1px solid var(--line,#ddd);border-radius:12px;padding:28px;box-shadow:0 12px 40px rgba(0,0,0,.12)">' + html + '</div>';
    el.hidden = false;
    return el;
  }
  function ungate() { const el = document.getElementById('sv-gate'); if (el) el.remove(); }
  const esc = (v) => String(v == null ? '' : v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const LOGO = '<div style="display:flex;align-items:center;gap:12px;margin-bottom:18px"><img src="icon.png" alt="" style="width:44px;height:44px;border-radius:8px"><div><b style="font-size:19px">Encomendas</b><div style="font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#85601C">Monte do Pasto · Frio</div></div></div>';

  // ------------------------------------------------------------------ autenticação (MSAL)
  if (!window.msal) { gate(LOGO + '<p>Falta a biblioteca de entrada Microsoft (vendor/msal-browser.min.js).</p>'); return; }
  if (!CFG.clientId || !CFG.tenantId || String(CFG.clientId).startsWith('COLOCAR')) {
    gate(LOGO + '<p><b>Instalação por concluir.</b></p><p>Falta preencher <code>config.js</code> com o registo da app no Microsoft Entra (clientId, tenantId) e a pasta SharePoint dos dados. Ver <code>docs/INSTALAR.md</code>.</p>');
    return;
  }
  const pca = new msal.PublicClientApplication({
    auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + CFG.tenantId, redirectUri: CFG.redirectUri || (location.origin + location.pathname) },
    cache: { cacheLocation: 'localStorage' }
  });
  let account = null;

  async function token() {
    try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|login_required|consent_required/.test(e.errorCode || '')) {
        await pca.acquireTokenRedirect({ scopes: SCOPES, account });
        return new Promise(() => {}); // a página vai ser redirecionada
      }
      throw e;
    }
  }

  // ------------------------------------------------------------------ Microsoft Graph
  class GraphError extends Error { constructor(status, body) { super((body && body.error && body.error.message) || ('HTTP ' + status)); this.status = status; this.body = body; this.code = status === 401 ? 'needs_reauth' : status === 403 ? 'tool_error' : status === 404 ? 'not_found' : status === 409 || status === 412 ? 'conflict' : status >= 500 ? 'server_unavailable' : 'tool_error'; } }
  async function graph(path, opts) {
    opts = opts || {};
    const t = await token();
    const headers = Object.assign({ Authorization: 'Bearer ' + t }, opts.headers || {});
    let body = opts.body;
    if (body !== undefined && !(body instanceof Blob) && !(body instanceof ArrayBuffer) && typeof body !== 'string') { body = JSON.stringify(body); headers['Content-Type'] = 'application/json'; }
    for (let attempt = 0; attempt < 4; attempt++) {
      const r = await fetch(path.startsWith('http') ? path : GRAPH + path, { method: opts.method || 'GET', headers, body });
      if (r.status === 429 || r.status === 503) { await sleep(((+r.headers.get('Retry-After')) || (attempt + 1)) * 1000); continue; }
      if (opts.raw) { if (!r.ok) throw new GraphError(r.status, await r.json().catch(() => null)); return r; }
      const txt = await r.text();
      const json = txt ? (() => { try { return JSON.parse(txt); } catch (e) { return txt; } })() : null;
      if (!r.ok) throw new GraphError(r.status, json);
      json && typeof json === 'object' && Object.defineProperty(json, '__etag', { value: r.headers.get('ETag'), enumerable: false });
      return json;
    }
    throw new GraphError(429, null);
  }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const enc = (p) => p.split('/').map(encodeURIComponent).join('/');

  // ------------------------------------------------------------------ base de dados em ficheiros JSON no SharePoint
  // Documento "a/b"      -> ficheiro dados/a/b.json
  // Coleção  "a/b/c"     -> pasta    dados/a/b/c/  (cada documento é um .json lá dentro)
  const DRIVE = CFG.dataDriveId, ROOT = CFG.dataFolderId;
  const itemPath = (rel) => '/drives/' + encodeURIComponent(DRIVE) + '/items/' + encodeURIComponent(ROOT) + ':/' + enc(rel) + ':';
  const cache = {}; // caminho -> { etag, data }

  async function readDoc(path) {
    let meta;
    try { meta = await graph(itemPath(path + '.json') + '?$select=id,eTag,@microsoft.graph.downloadUrl'); }
    catch (e) { if (e.status === 404) { delete cache[path]; return { exists: false, data: null, etag: null }; } throw e; }
    const c = cache[path];
    if (c && c.etag === meta.eTag) return { exists: true, data: c.data, etag: meta.eTag };
    const url = meta['@microsoft.graph.downloadUrl'];
    const r = url ? await fetch(url) : await graph(itemPath(path + '.json') + '/content', { raw: true });
    const data = await r.json();
    cache[path] = { etag: meta.eTag, data };
    return { exists: true, data, etag: meta.eTag };
  }
  async function writeDoc(path, data, opts) {
    opts = opts || {};
    const headers = {};
    if (opts.ifMatch) headers['If-Match'] = opts.ifMatch;
    const q = opts.createOnly ? '?@microsoft.graph.conflictBehavior=fail' : '';
    const body = new Blob([JSON.stringify(data)], { type: 'application/json' });
    const res = await graph(itemPath(path + '.json') + '/content' + q, { method: 'PUT', body, headers });
    cache[path] = { etag: res.eTag, data: JSON.parse(JSON.stringify(data)) };
    poke(path);
    return res;
  }
  async function listCol(path) {
    let res;
    try { res = await graph(itemPath(path) + '/children?$select=name,eTag,file,@microsoft.graph.downloadUrl&$top=999'); }
    catch (e) { if (e.status === 404) return []; throw e; }
    let items = res.value || []; let next = res['@odata.nextLink'];
    while (next) { const r2 = await graph(next); items = items.concat(r2.value || []); next = r2['@odata.nextLink']; }
    const files = items.filter((i) => i.file && /\.json$/i.test(i.name));
    const out = await Promise.all(files.map(async (f) => {
      const id = f.name.replace(/\.json$/i, ''); const p = path + '/' + id;
      const c = cache[p];
      if (c && c.etag === f.eTag) return { id, data: c.data };
      const r = await fetch(f['@microsoft.graph.downloadUrl']);
      const data = await r.json(); cache[p] = { etag: f.eTag, data };
      return { id, data };
    }));
    return out;
  }
  const docSnap = (id, exists, data) => ({ id, exists, data: () => (data == null ? data : JSON.parse(JSON.stringify(data))) });
  const colSnap = (list) => ({ docs: list.map((d) => docSnap(d.id, true, d.data)), size: list.length, empty: !list.length });

  // Atualizações "em tempo real": leitura periódica + leitura imediata depois de cada gravação nesta página.
  const watchers = new Set();
  function poke(path) { watchers.forEach((w) => { if (path === w.path || path.startsWith(w.path + '/')) w.run(); }); }
  function watch(path, fn, onErr) {
    let last = undefined, busy = false, again = false;
    const w = { path, run: async () => {
      if (busy) { again = true; return; } busy = true;
      try { const v = await fn(); const sig = JSON.stringify(v.sig); if (sig !== last) { last = sig; v.emit(); } }
      catch (e) { onErr && onErr(e); }
      finally { busy = false; if (again) { again = false; w.run(); } }
    } };
    watchers.add(w); w.run();
    return () => watchers.delete(w);
  }
  setInterval(() => { if (document.visibilityState !== 'hidden') watchers.forEach((w) => w.run()); }, POLL_MS);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') watchers.forEach((w) => w.run()); });

  function docRef(path) {
    const id = path.split('/').pop();
    return {
      id, path,
      async get() { const r = await readDoc(path); return docSnap(id, r.exists, r.data); },
      async set(data) { await writeDoc(path, data); },
      async update(patch) { const r = await readDoc(path); await writeDoc(path, Object.assign({}, r.data || {}, patch)); },
      async delete() { try { await graph(itemPath(path + '.json'), { method: 'DELETE' }); } catch (e) { if (e.status !== 404) throw e; } delete cache[path]; poke(path); },
      onSnapshot(f, onErr) { return watch(path, async () => { const r = await readDoc(path); return { sig: r.etag, emit: () => f(docSnap(id, r.exists, r.data)) }; }, onErr); },
      // Bloqueio com prazo (lease): gravação condicional pelo eTag do ficheiro.
      async acquire(o) {
        const holder = o.holder, ttl = o.ttlMs || 10000, now = Date.now();
        const lockPath = path;
        for (let i = 0; i < 3; i++) {
          let cur;
          try { cur = await readDoc(lockPath); } catch (e) { throw e; }
          const d = cur.data;
          if (cur.exists && d && d.holder !== holder && d.until > now) return { acquired: false, expiresAt: d.until };
          try {
            await writeDoc(lockPath, { holder, until: Date.now() + ttl }, cur.exists ? { ifMatch: cur.etag } : { createOnly: true });
            return { acquired: true, expiresAt: Date.now() + ttl };
          } catch (e) { if (e.code !== 'conflict') throw e; delete cache[lockPath]; }
        }
        return { acquired: false };
      }
    };
  }
  function colRef(path, order) {
    const sortList = (list) => {
      if (!order) return list.sort((a, b) => (a.id < b.id ? -1 : 1));
      const k = order.field, dir = order.dir === 'desc' ? -1 : 1;
      list.sort((a, b) => ((a.data || {})[k] < (b.data || {})[k] ? -dir : (a.data || {})[k] > (b.data || {})[k] ? dir : 0));
      return order.limit ? list.slice(0, order.limit) : list;
    };
    return {
      path,
      doc(id) { return docRef(path + '/' + id); },
      orderBy(field, dir) { return colRef(path, Object.assign({}, order, { field, dir })); },
      limit(n) { return colRef(path, Object.assign({}, order, { limit: n })); },
      async get() { return colSnap(sortList(await listCol(path))); },
      onSnapshot(f, onErr) { return watch(path, async () => { const l = sortList(await listCol(path)); return { sig: l.map((d) => [d.id, cache[path + '/' + d.id] && cache[path + '/' + d.id].etag]), emit: () => f(colSnap(l)) }; }, onErr); }
    };
  }
  const db = { doc: (p) => docRef(p), collection: (p) => colRef(p) };

  // ------------------------------------------------------------------ utilizador
  const me = { id: null, name: '', email: '' };
  const admins = (CFG.admins || []).map((x) => String(x).toLowerCase());
  const user = {
    async id() { return me.id; },
    async me() { return { id: me.id, name: me.name, email: me.email, avatarUrl: '', color: '#247305', isOwner: admins.includes(me.email) }; },
    async name() { return me.name; },
    async canEdit() { return !admins.length || admins.includes(me.email); },
    async can() { return true; },
    async profiles(ids) {
      const out = {};
      await Promise.all(ids.map(async (id) => { try { const r = await readDoc('users/' + id.replace(/[^A-Za-z0-9_\-.~:@+]/g, '_')); if (r.exists) out[id] = { name: r.data.name }; } catch (e) {} }));
      return out;
    }
  };

  // ------------------------------------------------------------------ descarregar ficheiros
  const downloads = {
    async save(o) {
      const blob = o.data instanceof Blob ? o.data : new Blob([o.data]);
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = o.filename || 'ficheiro';
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000);
    }
  };

  // ------------------------------------------------------------------ "conectores": Microsoft 365 pelo Graph
  const result = (payload) => ({ payload, content: [{ type: 'text', text: typeof payload === 'string' ? payload : JSON.stringify(payload) }] });
  const msgUri = (id) => 'mail:///messages/' + encodeURIComponent(id);
  const mapMsg = (m) => ({ id: m.id, uri: msgUri(m.id), subject: m.subject, sender: m.from ? { name: m.from.emailAddress.name, address: m.from.emailAddress.address } : m.sender ? { name: m.sender.emailAddress.name, address: m.sender.emailAddress.address } : null, receivedDateTime: m.receivedDateTime, createdDateTime: m.createdDateTime, lastModifiedDateTime: m.lastModifiedDateTime, hasAttachments: m.hasAttachments, webLink: m.webLink, isDraft: m.isDraft });
  const WELL_KNOWN = { drafts: 'drafts', rascunhos: 'drafts', inbox: 'inbox', 'caixa de entrada': 'inbox', 'sent items': 'sentitems', 'itens enviados': 'sentitems', 'deleted items': 'deleteditems', archive: 'archive' };
  const folderIds = {};
  async function folderId(name) {
    const k = String(name || '').trim().toLowerCase();
    if (WELL_KNOWN[k]) return WELL_KNOWN[k];
    if (folderIds[k]) return folderIds[k];
    const q = encodeURIComponent("displayName eq '" + String(name).replace(/'/g, "''") + "'");
    let r = await graph('/me/mailFolders?$filter=' + q + '&$select=id,displayName&$top=1');
    if (!(r.value && r.value.length)) r = await graph('/me/mailFolders/inbox/childFolders?$filter=' + q + '&$select=id,displayName&$top=1');
    if (!(r.value && r.value.length)) { const e = new Error('A pasta "' + name + '" não existe na caixa de correio.'); e.code = 'tool_error'; throw e; }
    return (folderIds[k] = r.value[0].id);
  }
  function sheetText(wb) {
    const out = [];
    wb.worksheets.forEach((ws) => {
      out.push('=== Sheet: ' + ws.name + ' ===');
      const nc = ws.columnCount;
      for (let r = 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r); const cells = [];
        for (let c = 1; c <= nc; c++) {
          let v = row.getCell(c).value;
          if (v && typeof v === 'object' && !(v instanceof Date)) v = v.richText ? v.richText.map((x) => x.text).join('') : ('result' in v ? v.result : v.text != null ? v.text : '');
          if (v instanceof Date) v = v.toISOString().slice(0, 10);
          cells.push(v == null ? '' : String(v).replace(/[\t\r\n]+/g, ' '));
        }
        while (cells.length && cells[cells.length - 1] === '') cells.pop();
        out.push(cells.join('\t'));
      }
      out.push('');
    });
    return out.join('\n');
  }
  const b64ToBytes = (b) => Uint8Array.from(atob(b), (c) => c.charCodeAt(0));
  const TOOLS = {
    async get_me() { const r = await graph('/me?$select=id,displayName,mail,userPrincipalName'); return result({ id: r.id, displayName: r.displayName, mail: r.mail, userPrincipalName: r.userPrincipalName }); },
    async outlook_email_search(a) {
      if (a.query && !a.folderName) {
        // Pesquisa em todas as pastas (ex.: "Relatório de preparação")
        // Aceita 'subject:"…"' (KQL) e afterDateTime (filtrado aqui: $search não combina com $filter).
        const q = String(a.query).replace(/"/g, '').trim();
        const top = Math.min(Math.max(+a.limit || 10, a.afterDateTime ? 50 : 1), 50);
        const r = await graph('/me/messages?$search=' + encodeURIComponent('"' + q + '"') + '&$top=' + top + '&$select=id,subject,from,sender,receivedDateTime,createdDateTime,lastModifiedDateTime,hasAttachments,webLink,isDraft');
        const after = a.afterDateTime ? String(a.afterDateTime).slice(0, 10) : '';
        return result((r.value || []).filter((x) => !after || String(x.receivedDateTime || '').slice(0, 10) >= after).map(mapMsg));
      }
      const fid = await folderId(a.folderName || 'inbox');
      const top = Math.min(+a.limit || 10, 50);
      const r = await graph('/me/mailFolders/' + encodeURIComponent(fid) + '/messages?$top=' + top + '&$orderby=receivedDateTime desc&$select=id,subject,from,sender,receivedDateTime,createdDateTime,lastModifiedDateTime,hasAttachments,webLink,isDraft');
      return result((r.value || []).map(mapMsg));
    },
    async read_resource(a) {
      const m = String(a.uri || '').match(/^mail:\/\/\/messages\/([^/?]+)(?:\/attachments\/([^/?]+))?/);
      if (!m) { const e = new Error('Recurso não suportado: ' + a.uri); e.code = 'tool_error'; throw e; }
      const mid = decodeURIComponent(m[1]);
      if (!m[2]) {
        const msg = await graph('/me/messages/' + encodeURIComponent(mid) + '?$expand=attachments($select=id,name,contentType,size,isInline)');
        const o = mapMsg(msg);
        o.body = msg.body;
        o.attachments = (msg.attachments || []).map((x) => ({ uri: msgUri(mid) + '/attachments/' + encodeURIComponent(x.id), name: x.name, contentType: x.contentType, size: x.size, isInline: x.isInline }));
        return result(o);
      }
      const att = await graph('/me/messages/' + encodeURIComponent(mid) + '/attachments/' + encodeURIComponent(decodeURIComponent(m[2])));
      if (/\.xlsx$/i.test(att.name || '') && att.contentBytes) {
        const wb = new window.ExcelJS.Workbook(); await wb.xlsx.load(b64ToBytes(att.contentBytes).buffer);
        return result(sheetText(wb));
      }
      return result({ name: att.name, contentType: att.contentType, contentBytes: att.contentBytes });
    },
    async sharepoint_upload_file(a) {
      // O pedido de rascunho (.outlook.json) já é a mensagem do Graph: aqui cria-se o rascunho diretamente, sem fluxo Power Automate.
      if (/\.outlook\.json$/i.test(a.filename || '') && a.content) {
        const d = await graph('/me/messages', { method: 'POST', body: JSON.parse(a.content) });
        return result({ id: d.id, name: a.filename, webUrl: d.webLink, draft: true });
      }
      const bytes = a.contentBase64 ? b64ToBytes(a.contentBase64) : new TextEncoder().encode(a.content || '');
      const q = '?@microsoft.graph.conflictBehavior=' + (a.conflictBehavior || 'fail');
      const path = '/drives/' + encodeURIComponent(a.driveId) + '/items/' + encodeURIComponent(a.parentItemId || 'root') + ':/' + encodeURIComponent(a.filename) + ':/content' + q;
      const r = await graph(path, { method: 'PUT', body: new Blob([bytes]) });
      return result({ id: r.id, name: r.name, webUrl: r.webUrl, item: { webUrl: r.webUrl } });
    }
  };
  TOOLS.outlook_send_mail = async (a) => {
    const rcp = (l) => (l || []).map((x) => ({ emailAddress: { address: x } }));
    await graph('/me/sendMail', { method: 'POST', body: { message: { subject: a.subject, body: { contentType: a.bodyType === 'html' ? 'HTML' : 'Text', content: a.body }, toRecipients: rcp(a.to), ccRecipients: rcp(a.cc), bccRecipients: rcp(a.bcc), attachments: (a.attachments || []).map((x) => ({ '@odata.type': '#microsoft.graph.fileAttachment', name: x.name, contentType: x.contentType || 'application/octet-stream', contentBytes: x.contentBytes })) }, saveToSentItems: true } });
    return result({ sent: true });
  };
  const mcp = {
    async callTool(server, tool, args) {
      if (server !== 'Microsoft 365') { const e = new Error(server + ' não está disponível nesta versão (só Microsoft 365).'); e.code = 'server_not_connected'; throw e; }
      const f = TOOLS[tool];
      if (!f) { const e = new Error('Ação não disponível: ' + tool); e.code = 'not_in_manifest'; throw e; }
      return f(args || {});
    }
  };

  // ------------------------------------------------------------------ arranque: entrar e verificar acesso à pasta
  let resolveReady, rejectReady;
  const ready = new Promise((res, rej) => { resolveReady = res; rejectReady = rej; });
  ready.catch(() => {});
  window.__SV_POLL_MS = 4000; // o rascunho é criado na hora; a app verifica-o depressa

  async function start() {
    await pca.initialize();
    const red = await pca.handleRedirectPromise().catch((e) => { gate(LOGO + '<p>Não foi possível entrar: ' + esc(e.message) + '</p><p><button id="sv-login">Tentar de novo</button></p>'); hookLogin(); return null; });
    account = (red && red.account) || pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    if (!account) {
      gate(LOGO + '<p>Entre com a sua conta Microsoft <b>@' + esc(CFG.domain || 'montedopasto.pt') + '</b>.</p><p><button id="sv-login" style="font:600 15px Montserrat,Segoe UI,Arial,sans-serif;background:#247305;color:#fff;border:0;border-radius:8px;padding:12px 18px;cursor:pointer;width:100%">Entrar com a conta Microsoft</button></p><p style="font-size:13px;color:#5F5A51">Só entra quem tem conta da empresa e acesso à pasta SalesVision no SharePoint.</p>');
      hookLogin(); return;
    }
    pca.setActiveAccount(account);
    const email = String(account.username || '').toLowerCase();
    if (CFG.domain && !email.endsWith('@' + CFG.domain.toLowerCase())) {
      gate(LOGO + '<p>A conta <b>' + esc(email) + '</b> não é da empresa. Entre com a conta @' + esc(CFG.domain) + '.</p><p><button id="sv-logout">Sair</button></p>'); hookLogout(); return;
    }
    me.email = email; me.name = account.name || email; me.id = email;
    try { await graph('/drives/' + encodeURIComponent(DRIVE) + '/items/' + encodeURIComponent(ROOT) + '?$select=id,name'); }
    catch (e) {
      gate(LOGO + '<p><b>Sem acesso à pasta de dados.</b></p><p>A conta ' + esc(email) + ' não tem acesso à pasta SalesVision no SharePoint (' + esc(e.message) + '). Peça acesso ao administrador.</p><p><button id="sv-logout">Sair</button></p>'); hookLogout(); return;
    }
    // Perfil (nome) para o histórico das encomendas.
    writeDoc('users/' + me.id.replace(/[^A-Za-z0-9_\-.~:@+]/g, '_'), { name: me.name, email: me.email, lastSeen: new Date().toISOString() }).catch(() => {});
    ungate();
    addUserChip();
    resolveReady();
  }
  function hookLogin() { const b = document.getElementById('sv-login'); if (b) b.onclick = () => pca.loginRedirect({ scopes: SCOPES, prompt: 'select_account' }); }
  function hookLogout() { const b = document.getElementById('sv-logout'); if (b) b.onclick = () => pca.logoutRedirect({ account }); }
  function addUserChip() {
    const tryAdd = () => {
      const row = document.querySelector('.top-row'); const th = document.getElementById('themebtn');
      if (!row || !th) return setTimeout(tryAdd, 300);
      if (document.getElementById('sv-user')) return;
      const b = document.createElement('button'); b.id = 'sv-user'; b.type = 'button'; b.className = 'btn themebtn';
      b.title = me.name + ' (' + me.email + ') · Sair'; b.setAttribute('aria-label', 'Sair (' + me.email + ')');
      b.style.cssText = 'width:auto;padding:0 12px;font-size:13px;font-weight:700';
      b.textContent = (me.name || me.email).split(/\s+/).map((x) => x[0]).slice(0, 2).join('').toUpperCase();
      b.onclick = () => { if (confirm('Sair da conta ' + me.email + '?')) pca.logoutRedirect({ account }); };
      th.after(b);
    };
    tryAdd();
  }

  window.claude = {
    async use(name) {
      await ready;
      if (name === 'db') return db;
      if (name === 'user') return user;
      if (name === 'downloads') return downloads;
      if (name === 'mcp') return mcp;
      return null;
    }
  };
  // Para testes e manutenção.
  window.Encomendas = { db, graph, me, sheetText };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => start().catch((e) => gate(LOGO + '<p>Erro ao arrancar: ' + esc(e.message) + '</p>')));
  else start().catch((e) => gate(LOGO + '<p>Erro ao arrancar: ' + esc(e.message) + '</p>'));
})();
