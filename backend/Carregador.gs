/**
 * ICY Academy · carregador do script.
 *
 * Este é o ÚNICO código que fica colado no Apps Script. O código verdadeiro vive no
 * GitHub (icyacademy/icy-academy, ficheiro backend/codigo.js) e é publicado com a app.
 * Assim, as atualizações chegam sozinhas: não é preciso colar nada nem criar nova versão.
 *
 * Instalar (uma só vez):
 * 1. Extensões > Apps Script. Apagar tudo e colar este ficheiro. Guardar.
 * 2. Correr "experimentar" e autorizar quando pedir.
 * 3. Implementar > Gerir implementações > lápis > Versão: Nova versão > Implementar.
 * 4. Correr "ativarRelatorioMensal" uma vez (liga o relatório do dia 1).
 */

const CODIGO_URL = 'https://icy-academy.netlify.app/backend/codigo.js';
const CACHE_MIN = 5; // as alterações chegam em até 5 minutos

function doGet(e) { return icy_().doGet(e); }
function doPost(e) { return icy_().doPost(e); }
function configurar() { return icy_().configurar(); }
function relatorioMensal() { return icy_().relatorioMensal(); }
function ativarRelatorioMensal() { return icy_().ativarRelatorioMensal(); }
function relatorioTeste() { return icy_().relatorioTeste(); }

/** Corre isto depois de colar: confirma que o código é descarregado e mostra a versão. */
function experimentar() {
  CacheService.getScriptCache().remove('icy_codigo');
  const api = icy_();
  const msg = 'Código carregado (' + (api.VERSAO || 'sem versão') + ').';
  Logger.log(msg);
  try { SpreadsheetApp.getActive().toast(msg); } catch (e) {}
}

let ICY_API_ = null;
function icy_() {
  if (ICY_API_) return ICY_API_;
  const cache = CacheService.getScriptCache();
  let src = cache.get('icy_codigo');
  if (!src) {
    try {
      const r = UrlFetchApp.fetch(CODIGO_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
      if (r.getResponseCode() === 200 && r.getContentText('UTF-8').indexOf('ICY_CODIGO_OK') >= 0) {
        src = r.getContentText('UTF-8');
        guardarCopia_(src);
        cache.put('icy_codigo', src, CACHE_MIN * 60);
      }
    } catch (e) {}
    if (!src) src = lerCopia_(); // sem rede ou site em baixo: usa a última versão boa
    if (!src) throw new Error('Não foi possível carregar o código da ICY Academy.');
  }
  ICY_API_ = new Function(src + '\n;return { VERSAO: typeof VERSAO_CODIGO !== "undefined" ? VERSAO_CODIGO : "", doGet, doPost, configurar, relatorioMensal, ativarRelatorioMensal, relatorioTeste };')();
  return ICY_API_;
}

// Última versão boa, guardada em pedaços nas propriedades do script (limite de 9 KB por valor).
function guardarCopia_(src) {
  const p = PropertiesService.getScriptProperties();
  const partes = src.match(/[\s\S]{1,8000}/g) || [];
  const atual = Number(p.getProperty('icy_partes') || 0);
  const obj = { icy_partes: String(partes.length) };
  partes.forEach((x, i) => { obj['icy_p' + i] = x; });
  p.setProperties(obj);
  for (let i = partes.length; i < atual; i++) p.deleteProperty('icy_p' + i);
}
function lerCopia_() {
  const p = PropertiesService.getScriptProperties();
  const n = Number(p.getProperty('icy_partes') || 0);
  let s = '';
  for (let i = 0; i < n; i++) s += p.getProperty('icy_p' + i) || '';
  return s;
}

/*
 * Não apagar: estas referências dizem ao Google que autorizações o código precisa
 * (o código descarregado não é analisado na altura da autorização).
 */
function permissoes_() {
  SpreadsheetApp.getActive(); DriveApp.getRootFolder(); MailApp.getRemainingDailyQuota();
  UrlFetchApp.fetch(''); ScriptApp.getOAuthToken(); LockService.getScriptLock();
  ContentService.createTextOutput(''); Session.getEffectiveUser(); PropertiesService.getScriptProperties();
}
