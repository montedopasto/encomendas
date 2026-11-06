'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');

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

  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    let pageErrors = 0;
    page.on('pageerror', () => { pageErrors++; });

    await page.route(site + '/**', async (route) => {
      const url = new URL(route.request().url());
      const requestedPath = url.pathname === '/' ? '/app.html' : url.pathname;
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
    'Credenciais do robô por configurar.',
    'Falha na autenticação do robô.',
    'A aplicação reportou um erro.',
    'A aplicação não conseguiu iniciar.',
    'A aplicação não terminou dentro do limite.',
  ]);
  console.error(safeMessages.has(error.message) ? error.message : 'O robô encontrou um erro inesperado.');
  process.exitCode = 1;
});
