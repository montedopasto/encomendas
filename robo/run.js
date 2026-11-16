'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const vm = require('node:vm');

const publicDir = path.resolve(__dirname, '../public');
const site = 'https://encomendas.montedopasto.pt';
const mailbox = 'trueborn@montedopasto.pt';
const graphOrigin = 'https://graph.microsoft.com';
const mime = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.png': 'image/png',
  '.webmanifest': 'application/manifest+json',
};

async function run() {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Credenciais do robô por configurar.');
  }

  const { chromium } = require('playwright');
  let cachedToken;
  async function token() {
    if (cachedToken && cachedToken.expiresAt > Date.now() + 120000) return cachedToken.value;
    const response = await fetch(
      'https://login.microsoftonline.com/' + encodeURIComponent(TENANT_ID) + '/oauth2/v2.0/token',
      {
        method: 'POST',
        body: new URLSearchParams({
          client_id: CLIENT_ID,
          client_secret: CLIENT_SECRET,
          scope: graphOrigin + '/.default',
          grant_type: 'client_credentials',
        }),
        signal: AbortSignal.timeout(15000),
      },
    );
    if (!response.ok) throw new Error('Falha na autenticação do robô.');
    const result = await response.json();
    cachedToken = { value: result.access_token, expiresAt: Date.now() + result.expires_in * 1000 };
    return cachedToken.value;
  }

  let app;
  try {
    const config = { window: {} };
    vm.runInNewContext(await fs.readFile(path.join(publicDir, 'config.js'), 'utf8'), config, { timeout: 1000 });
    const { dataDriveId, dataFolderId } = config.window.ENCOMENDAS_CONFIG;
    const drive = '/v1.0/drives/' + encodeURIComponent(dataDriveId) + '/items/';
    async function readGraph(resource, stage) {
      const response = await fetch(graphOrigin + resource, {
        headers: { Authorization: 'Bearer ' + await token() },
        signal: AbortSignal.timeout(15000),
      }).catch(() => null);
      if (!response?.ok) throw new Error(stage + ' · ' + (response ? 'HTTP ' + response.status : 'sem resposta'));
      return response;
    }
    const data = await (await readGraph(drive + encodeURIComponent(dataFolderId) + '?$select=parentReference', 'localizar dados')).json();
    if (!data.parentReference?.id) throw new Error('pasta pai não encontrada');
    const parent = drive + encodeURIComponent(data.parentReference.id);
    const folder = await (await readGraph(parent + ':/app:/children?$select=id,name,lastModifiedDateTime&$top=100', 'listar app')).json();
    const index = folder.value?.find(item => item.name.toLowerCase() === 'index.html');
    const platform = folder.value?.find(item => item.name.toLowerCase() === 'platform.js');
    if (!index || !platform) throw new Error('index.html ou platform.js em falta');
    const [html, js] = await Promise.all([
      readGraph(drive + encodeURIComponent(index.id) + '/content', 'ler index.html').then(r => r.text()),
      readGraph(drive + encodeURIComponent(platform.id) + '/content', 'ler platform.js').then(r => r.text()),
    ]);
    if (!html.includes('__svEstado') || !js.trim()) throw new Error('ficheiros da app incompletos');
    app = { html, js };
    const version = new Date(index.lastModifiedDateTime);
    console.log('App: SharePoint (' + (Number.isNaN(version.valueOf()) ? 'data indisponível' : version.toISOString()) + ')');
  } catch (error) {
    const safe = /^(localizar dados|listar app|ler index\.html|ler platform\.js) · (HTTP \d{3}|sem resposta)$|^(pasta pai não encontrada|index\.html ou platform\.js em falta|ficheiros da app incompletos)$/;
    console.error('Pasta app do SharePoint indisponível (' + (safe.test(error.message) ? error.message : 'erro de leitura ou configuração') + ').');
    throw new Error('Ciclo interrompido para evitar usar uma versão antiga da app.');
  }

  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    let pageErrors = 0;
    page.on('pageerror', () => { pageErrors++; });

    await page.route(site + '/**', async (route) => {
      const url = new URL(route.request().url());
      if (['/', '/index.html', '/app.html'].includes(url.pathname)) {
        return route.fulfill({ body: app.html, contentType: mime['.html'] });
      }
      if (url.pathname === '/platform.js') {
        return route.fulfill({ body: app.js, contentType: mime['.js'] });
      }
      const requestedPath = url.pathname;
      const file = path.resolve(publicDir, '.' + requestedPath);
      if (!file.startsWith(publicDir + path.sep)) return route.abort();
      let body;
      try {
        body = await fs.readFile(file);
      } catch (error) {
        if (error.code === 'ENOENT' || error.code === 'EISDIR') return route.fulfill({ status: 404 });
        throw error;
      }
      return route.fulfill({ body, contentType: mime[path.extname(file)] || 'application/octet-stream' });
    });

    await page.exposeFunction('__svRobotToken', token);
    if (process.env.GEMINI_API_KEY) {
      const { criarLeitor } = require('./ler-guia');
      await page.exposeFunction('__svLerGuia', criarLeitor(process.env.GEMINI_API_KEY));
    }
    await page.addInitScript((address) => {
      window.__svRobot = { mailbox: address, nome: 'Robô Encomendas' };
    }, mailbox);

    await page.goto(site + '/', { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(15000);
    const deadline = Date.now() + 210000;
    let idle = 0;
    while (Date.now() < deadline && idle < 3) {
      if (pageErrors) throw new Error('A aplicação reportou um erro.');
      const status = await page.evaluate(() => ({
        blocked: Boolean(document.getElementById('sv-gate')),
        state: window.__svEstado?.(),
      }));
      if (status.blocked) throw new Error('A aplicação não conseguiu iniciar.');
      const state = status.state;
      idle = state && !state.prep && !state.mail && !state.busy ? idle + 1 : 0;
      if (idle < 3) await page.waitForTimeout(2000);
    }
    if (idle < 3) throw new Error('A aplicação não terminou dentro do limite.');
    console.log('Ciclo do robô terminado.');
  } finally {
    await browser.close();
  }
}

run().catch((error) => {
  const safeMessages = new Set([
    'Ciclo interrompido para evitar usar uma versão antiga da app.',
    'Credenciais do robô por configurar.',
    'Falha na autenticação do robô.',
    'A aplicação reportou um erro.',
    'A aplicação não conseguiu iniciar.',
    'A aplicação não terminou dentro do limite.',
  ]);
  console.error(safeMessages.has(error.message) ? error.message : 'O robô encontrou um erro inesperado.');
  process.exitCode = 1;
});
