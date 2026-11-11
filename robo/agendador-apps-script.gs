// Projeto Apps Script independente. Configurar nas propriedades do script:
// GITHUB_ACTIONS_TOKEN: token fine-grained de montedopasto/encomendas, Actions: Read and write.
// Executar testarLigacao() antes de ativar. No GitHub, confirmar ENCOMENDAS_ROBO_ATIVO=true
// e definir ENCOMENDAS_ROBO_EXTERNO=true (desliga a execução pelo cron nativo).
// Executar instalarAgendamento() e confirmar o primeiro ciclo em Actions.

const API = 'https://api.github.com/repos/montedopasto/encomendas/actions/workflows/encomendas-robo.yml';
const BRANCH = 'main';

function github_(suffix, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_ACTIONS_TOKEN');
  if (!token) throw new Error('Configura GITHUB_ACTIONS_TOKEN nas propriedades do script.');
  const options = {
    method: payload ? 'post' : 'get',
    headers: {Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json'},
    muteHttpExceptions: true,
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  return UrlFetchApp.fetch(API + suffix, options);
}

function testarLigacao() {
  const response = github_('');
  if (response.getResponseCode() !== 200) {
    throw new Error('GitHub recusou a consulta: HTTP ' + response.getResponseCode());
  }
  const state = JSON.parse(response.getContentText()).state;
  if (state !== 'active') throw new Error('Workflow GitHub: ' + state);
  console.log('Workflow ativo e acessível.');
}

function instalarAgendamento() {
  testarLigacao();
  const existing = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'verificarHorario');
  if (!existing) ScriptApp.newTrigger('verificarHorario').timeBased().everyMinutes(1).create();
}

function verificarHorario() {
  const now = new Date();
  const hour = now.getUTCHours();
  if (hour < 6 || hour > 21 || now.getUTCMinutes() < 7) return;
  const slot = now.toISOString().slice(0, 13);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('LAST_DISPATCHED_SLOT') === slot) return;
    const list = github_('/runs?event=workflow_dispatch&branch=main&per_page=30');
    if (list.getResponseCode() !== 200) {
      throw new Error('Não foi possível verificar o histórico: HTTP ' + list.getResponseCode());
    }
    const runs = JSON.parse(list.getContentText()).workflow_runs;
    if (runs.some(run => run.display_title === 'Encomendas Robô UTC ' + slot)) {
      props.setProperty('LAST_DISPATCHED_SLOT', slot);
      return;
    }
    // Se a resposta se perder, esperar antes de consultar o histórico e tentar novamente.
    const lastAttempt = Number(props.getProperty('LAST_DISPATCH_ATTEMPT_MS') || 0);
    if (Date.now() - lastAttempt < 180000) return;
    props.setProperty('LAST_DISPATCH_ATTEMPT_MS', String(Date.now()));
    const result = github_('/dispatches', {ref: BRANCH, inputs: {slot}});
    const code = result.getResponseCode();
    if (code !== 204 && code !== 200 && code !== 201) {
      throw new Error('GitHub recusou o acionamento: HTTP ' + code);
    }
    props.setProperty('LAST_DISPATCHED_SLOT', slot);
    console.log('Pedido aceite para ' + slot + ':07 UTC.');
  } finally {
    lock.releaseLock();
  }
}
