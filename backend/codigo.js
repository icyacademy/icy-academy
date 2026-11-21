/**
 * ICY Academy · ligação entre a app e a folha "Base de dados – App ICY Academy".
 *
 * Este ficheiro NÃO se cola no Apps Script: é descarregado pelo "Carregador" (backend/Carregador.gs)
 * a partir de https://icy-academy.netlify.app/backend/codigo.js. Para mudar o script basta
 * alterar este ficheiro no GitHub; a alteração chega em até 5 minutos.
 * Segredos (SEGREDO, códigos de acesso) ficam na folha e nas propriedades do script, nunca aqui.
 */

const VERSAO_CODIGO = 'v12 · relatório mensal'; // ICY_CODIGO_OK
const TZ = 'Europe/Lisbon';
const TOKEN_DIAS = 180;

/* ---------- Configuração inicial ---------- */

function configurar() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SEGREDO')) {
    props.setProperty('SEGREDO', Utilities.getUuid() + Utilities.getUuid());
  }
  // Garante a coluna "Código de acesso" e cria códigos para quem não tem.
  const sh = folha_('Professores');
  let cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let col = cab.indexOf('Código de acesso') + 1;
  if (!col) {
    col = cab.length + 1;
    sh.getRange(1, col).setValue('Código de acesso')
      .setBackground('#4A47C4').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setColumnWidth(col, 160);
  }
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const emails = sh.getRange(2, 2, n, 1).getValues();
    const rng = sh.getRange(2, col, n, 1);
    const cods = rng.getValues();
    for (let i = 0; i < n; i++) {
      if (emails[i][0] && !cods[i][0]) cods[i][0] = novoCodigo_();
    }
    rng.setNumberFormat('@').setValues(cods);
  }
  SpreadsheetApp.getActive().toast('Configuração feita. Os códigos de acesso estão no separador Professores.');
}

function novoCodigo_() {
  const letras = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += letras[Math.floor(Math.random() * letras.length)];
  return s.slice(0, 4) + '-' + s.slice(4);
}

/* ---------- Entrada HTTP ---------- */

function doGet() {
  return json_({ ok: true, app: 'ICY Academy', hora: new Date().toISOString() });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const acao = req.acao;
    if (acao === 'login') return json_(login_(req));
    const user = verificarToken_(req.token);
    const lock = LockService.getScriptLock();
    const escreve = acao !== 'dados' && acao !== 'fotos';
    if (escreve) lock.waitLock(20000);
    try {
      switch (acao) {
        case 'dados': return json_(dados_(user));
        case 'fotos': return json_(fotos_(user, req));
        case 'guardarPresencas': return json_(guardarPresencas_(user, req));
        case 'guardarAluno': return json_(guardarAluno_(user, req));
        case 'guardarPagamento': return json_(guardarPagamento_(user, req));
        case 'guardarCalendario': return json_(guardarCalendario_(user, req));
        case 'guardarCover': return json_(guardarCover_(user, req));
        case 'guardarPosicoes': return json_(guardarPosicoes_(user, req));
        case 'guardarEvento': return json_(guardarEvento_(user, req));
        case 'apagarPagamento': return json_(apagarPagamento_(user, req));
        case 'apagarEvento': return json_(apagarPorId_('Eventos', req.id));
        case 'apagarCover': return json_(apagarCover_(user, req));
        default: throw new Error('Ação desconhecida');
      }
    } finally {
      if (escreve) lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, erro: String(err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Autenticação ---------- */

function login_(req) {
  const email = String(req.email || '').trim().toLowerCase();
  const codigo = String(req.codigo || '').trim().toUpperCase();
  const cache = CacheService.getScriptCache();
  const chave = 'falhas_' + email;
  const falhas = Number(cache.get(chave) || 0);
  if (falhas >= 8) throw new Error('Demasiadas tentativas. Tente daqui a 15 minutos.');
  const p = professores_().find(x => x.email === email);
  if (!p || !p.ativo || !p.codigo || p.codigo !== codigo) {
    cache.put(chave, String(falhas + 1), 900);
    throw new Error('Email ou código errado.');
  }
  cache.remove(chave);
  const exp = Date.now() + TOKEN_DIAS * 86400000;
  const corpo = email + '|' + exp;
  return { ok: true, token: Utilities.base64EncodeWebSafe(corpo) + '.' + assinar_(corpo), nome: p.nome, perfil: p.perfil, email };
}

function assinar_(txt) {
  const seg = PropertiesService.getScriptProperties().getProperty('SEGREDO');
  if (!seg) throw new Error('Falta correr a função "configurar" no Apps Script.');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(txt, seg));
}

function verificarToken_(token) {
  const partes = String(token || '').split('.');
  if (partes.length !== 2) throw new Error('SESSAO');
  const corpo = Utilities.newBlob(Utilities.base64DecodeWebSafe(partes[0])).getDataAsString();
  if (assinar_(corpo) !== partes[1]) throw new Error('SESSAO');
  const [email, exp] = corpo.split('|');
  if (Date.now() > Number(exp)) throw new Error('SESSAO');
  const p = professores_().find(x => x.email === email);
  if (!p || !p.ativo) throw new Error('SESSAO');
  return p;
}

function professores_() {
  return ler_('Professores').map(r => ({
    nome: r['Nome'],
    email: String(r['Email Google'] || '').trim().toLowerCase(),
    perfil: r['Perfil'] || 'Professor',
    ativo: r['Ativo'] !== 'Não',
    codigo: String(r['Código de acesso'] || '').trim().toUpperCase()
  }));
}

function gestao_(user) {
  if (user.perfil !== 'Gestão') throw new Error('Só o perfil Gestão pode fazer isto.');
}

/* ---------- Leitura ---------- */

function dados_(user) {
  const turmas = ler_('Turmas').filter(t => t['ID']);
  const minhas = user.perfil === 'Gestão' ? turmas : turmas.filter(t => String(t['Email do professor']).toLowerCase() === user.email);
  const idsMinhas = new Set(minhas.map(t => t['ID']));
  const inscr = ler_('Inscrições').filter(i => i['ID do aluno']);
  const alunosVisiveis = user.perfil === 'Gestão' ? null : new Set(inscr.filter(i => idsMinhas.has(i['ID da turma'])).map(i => i['ID do aluno']));
  let alunos = ler_('Alunos').filter(a => a['ID'] && (!alunosVisiveis || alunosVisiveis.has(a['ID'])));
  const limite = new Date(); limite.setDate(limite.getDate() - 400);
  const limiteTxt = fmtData_(limite);
  const out = {
    ok: true,
    eu: { nome: user.nome, perfil: user.perfil, email: user.email },
    turmas: turmas.map(t => ({
      id: t['ID'], nome: t['Nome da turma'], estilo: t['Estilo'], local: t['Local'], dia: t['Dia da semana'],
      inicio: t['Hora de início'], fim: t['Hora de fim'], professor: t['Email do professor'],
      mensalidade: num_(t['Mensalidade (€)']), precoAula: num_(t['Preço por aula (€)']), ativa: t['Ativa'] !== 'Não', idades: t['Idades']
    })),
    alunos: alunos.map(a => ({
      id: a['ID'], nome: a['Nome (primeiro e último)'] || a['Nome'] || '', stage: a['Stage name'] || '', nascimento: nasc_(a['Data de nascimento'], a['Aniversário (dia/mês)']), contacto: a['Contacto'], email: a['Email'] || '',
      encarregado: a['Encarregado de educação'], contactoEnc: a['Contacto do encarregado'], emailEnc: a['Email do encarregado'],
      imagem: a['Autorização de imagem'], notas: a['Notas (alergias, lesões)'], inscricao: a['Data de inscrição'],
      ativo: a['Ativo'] !== 'Não', foto: a['Foto (link)'], localidade: a['Localidade'] || ''
    })),
    inscricoes: inscr.map(i => ({ aluno: i['ID do aluno'], turma: i['ID da turma'], inicio: i['Data de início'], fim: i['Data de fim'], ativa: i['Ativa'] !== 'Não', plano: i['Plano'] || 'Mensal' })),
    presencas: ler_('Presenças').filter(p => p['Data'] && p['Data'] >= limiteTxt).map(p => ({ data: p['Data'], turma: p['ID da turma'], aluno: p['ID do aluno'], estado: p['Estado'], tipo: p['Tipo de aula'] || 'Normal' })),
    calendario: ler_('Calendário').filter(c => c['Data']).map(c => ({ data: c['Data'], turma: c['ID da turma'], inicio: c['Hora de início'], fim: c['Hora de fim'], local: c['Local'], tipo: c['Tipo'] || 'Normal', notas: c['Notas'] })),
    covers: ler_('Covers').filter(c => c['ID']).map(c => ({
      id: c['ID'], musica: c['Música'], grupo: c['Grupo'], membros: num_(c['Nº de membros']), turma: c['ID da turma'], estado: c['Estado'],
      inicio: c['Data de início'], apresentacao: c['Data de apresentação'], practice: c['Dance practice (YouTube)'],
      espelhado: c['Versão espelhada (YouTube)'], outros: c['Outros links'], notas: c['Notas'], final: c['Vídeo final']
    })),
    posicoes: ler_('Posições').filter(p => p['ID do cover']).map(p => ({ cover: p['ID do cover'], aluno: p['ID do aluno'], parte: p['Membro / parte'], notas: p['Notas'] })),
    eventos: lerSeExistir_('Eventos').filter(e => e['ID']).map(e => ({
      id: e['ID'], nome: e['Nome'], tipo: e['Tipo'], estado: e['Estado'] || 'A considerar', data: e['Data'], dataFim: e['Data de fim'],
      hora: e['Hora'], local: e['Local'], prazo: e['Prazo de inscrição'], turmas: e['Turmas'], covers: e['Covers'],
      requisitos: e['Requisitos'], custo: num_(e['Custo (€)']), link: e['Link'], notas: e['Notas']
    })),
    professores: professores_().filter(p => p.ativo).map(p => ({ nome: p.nome, email: p.email, perfil: p.perfil }))
  };
  if (user.perfil === 'Gestão') {
    out.pagamentos = ler_('Pagamentos').filter(p => p['ID do aluno']).map(p => ({ aluno: p['ID do aluno'], mes: p['Mês (AAAA-MM)'], valor: num_(p['Valor (€)']), metodo: p['Método'], data: p['Data de pagamento'], notas: p['Notas'] }));
  }
  return out;
}

/**
 * Fotos dos alunos guardadas no Drive: devolve miniaturas (sem ser preciso partilhar as fotos).
 * Recebe uma lista de IDs de alunos; responde { fotos: { A001: 'data:image/...' } }.
 */
function fotos_(user, req) {
  const visiveis = new Set(dados_(user).alunos.map(a => a.id));
  const alunos = ler_('Alunos').filter(a => visiveis.has(a['ID']) && (req.ids || []).indexOf(a['ID']) >= 0);
  const cache = CacheService.getScriptCache();
  const out = {};
  alunos.forEach(a => {
    const m = String(a['Foto (link)'] || '').match(/[-\w]{25,}/);
    if (!m) return;
    const chave = 'foto_' + m[0];
    let d = cache.get(chave);
    if (!d) {
      try {
        let blob;
        try {
          const r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?sz=w240&id=' + m[0],
            { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
          if (r.getResponseCode() === 200 && /^image\//.test(r.getBlob().getContentType() || '')) blob = r.getBlob();
        } catch (e) {}
        if (!blob) blob = DriveApp.getFileById(m[0]).getThumbnail();
        if (!blob || blob.getBytes().length > 120000) return;
        d = 'data:' + (blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes());
        if (d.length < 95000) cache.put(chave, d, 21600);
      } catch (e) { return; }
    }
    out[a['ID']] = d;
  });
  return { ok: true, fotos: out };
}

/* ---------- Escrita ---------- */

function guardarPresencas_(user, req) {
  const data = req.data, turma = req.turma, tipo = req.tipo || 'Normal';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) throw new Error('Data inválida');
  const sh = folha_('Presenças');
  const t = tabela_(sh);
  const agora = new Date();
  const existentes = {};
  t.linhas.forEach((r, i) => { if (r['Data'] === data && r['ID da turma'] === turma) existentes[r['ID do aluno']] = i + 2; });
  const novas = [];
  (req.marcas || []).forEach(m => {
    const linha = existentes[m.aluno];
    if (!m.estado) {
      if (linha) sh.getRange(linha, t.col['Estado']).setValue('');
      return;
    }
    if (linha) {
      sh.getRange(linha, t.col['Estado']).setValue(m.estado);
      sh.getRange(linha, t.col['Marcado por']).setValue(user.nome);
      sh.getRange(linha, t.col['Registado em']).setValue(agora);
      if (t.col['Tipo de aula']) sh.getRange(linha, t.col['Tipo de aula']).setValue(tipo);
    } else {
      novas.push({ 'Data': paraData_(data), 'ID da turma': turma, 'ID do aluno': m.aluno, 'Estado': m.estado, 'Marcado por': user.nome, 'Registado em': agora, 'Tipo de aula': tipo });
    }
  });
  acrescentar_(sh, t, novas);
  return { ok: true };
}

function guardarAluno_(user, req) {
  const a = req.aluno || {};
  if (!a.nome && !a.stage) throw new Error('Falta o nome');
  const sh = folha_('Alunos');
  const t = tabela_(sh);
  let id = a.id;
  const obj = {
    'Nome': a.nome || '', 'Nome (primeiro e último)': a.nome || '', 'Stage name': a.stage || '', 'Data de nascimento': nascParaFolha_(a.nascimento), 'Contacto': a.contacto || '', 'Email': a.email || '',
    'Encarregado de educação': a.encarregado || '', 'Contacto do encarregado': a.contactoEnc || '', 'Email do encarregado': a.emailEnc || '',
    'Autorização de imagem': a.imagem || '', 'Notas (alergias, lesões)': a.notas || '', 'Ativo': a.ativo === false ? 'Não' : 'Sim', 'Foto (link)': a.foto || '',
    'Localidade': a.localidade || ''
  };
  if (id) {
    const i = t.linhas.findIndex(r => r['ID'] === id);
    if (i < 0) throw new Error('Aluno não encontrado');
    escreverLinha_(sh, t, i + 2, obj);
  } else {
    const nums = t.linhas.map(r => parseInt(String(r['ID']).replace(/\D/g, ''), 10) || 0);
    id = 'A' + String(Math.max(0, ...nums) + 1).padStart(3, '0');
    obj['ID'] = id;
    obj['Data de inscrição'] = paraData_(a.inscricao || fmtData_(new Date()));
    acrescentar_(sh, t, [obj]);
  }
  // Inscrições: a lista enviada substitui as ativas deste aluno
  if (Array.isArray(req.turmas)) {
    const si = folha_('Inscrições');
    const ti = tabela_(si);
    const hoje = paraData_(fmtData_(new Date()));
    const atuais = {};
    ti.linhas.forEach((r, i) => { if (r['ID do aluno'] === id && r['Ativa'] !== 'Não') atuais[r['ID da turma']] = i + 2; });
    const novas = [];
    req.turmas.forEach(x => {
      if (atuais[x.turma]) {
        si.getRange(atuais[x.turma], ti.col['Plano']).setValue(x.plano || 'Mensal');
        delete atuais[x.turma];
      } else {
        novas.push({ 'ID do aluno': id, 'ID da turma': x.turma, 'Data de início': hoje, 'Ativa': 'Sim', 'Plano': x.plano || 'Mensal' });
      }
    });
    Object.keys(atuais).forEach(tu => {
      si.getRange(atuais[tu], ti.col['Ativa']).setValue('Não');
      si.getRange(atuais[tu], ti.col['Data de fim']).setValue(hoje);
    });
    acrescentar_(si, ti, novas);
  }
  return { ok: true, id };
}

function guardarPagamento_(user, req) {
  gestao_(user);
  const p = req.pagamento || {};
  if (!p.aluno || !/^\d{4}-\d{2}$/.test(p.mes) || !(Number(p.valor) > 0)) throw new Error('Pagamento incompleto');
  const sh = folha_('Pagamentos');
  acrescentar_(sh, tabela_(sh), [{
    'ID do aluno': p.aluno, 'Mês (AAAA-MM)': p.mes, 'Valor (€)': Number(p.valor), 'Método': p.metodo || '',
    'Data de pagamento': paraData_(p.data || fmtData_(new Date())), 'Registado por': user.nome, 'Notas': p.notas || ''
  }]);
  return { ok: true };
}

/** Substitui as linhas do calendário de um mês (gestão) ou acrescenta uma aula (qualquer perfil). */
function guardarCalendario_(user, req) {
  const sh = folha_('Calendário');
  const t = tabela_(sh);
  const linha = c => ({ 'Data': paraData_(c.data), 'ID da turma': c.turma, 'Hora de início': c.inicio || '', 'Hora de fim': c.fim || '', 'Local': c.local || '', 'Tipo': c.tipo || 'Normal', 'Notas': c.notas || '' });
  if (req.mes) {
    gestao_(user);
    const apagar = [];
    t.linhas.forEach((r, i) => { if (String(r['Data']).slice(0, 7) === req.mes) apagar.push(i + 2); });
    apagar.reverse().forEach(r => sh.deleteRow(r));
    acrescentar_(sh, tabela_(sh), (req.aulas || []).map(linha));
  } else if (req.aula) {
    acrescentar_(sh, t, [linha(req.aula)]);
  }
  return { ok: true };
}

function guardarCover_(user, req) {
  const c = req.cover || {};
  const sh = folha_('Covers');
  const t = tabela_(sh);
  const obj = {
    'Música': c.musica || '', 'Grupo': c.grupo || '', 'Nº de membros': c.membros || '', 'ID da turma': c.turma || '', 'Estado': c.estado || 'Próximo',
    'Data de início': c.inicio ? paraData_(c.inicio) : '', 'Data de apresentação': c.apresentacao ? paraData_(c.apresentacao) : '',
    'Dance practice (YouTube)': c.practice || '', 'Versão espelhada (YouTube)': c.espelhado || '', 'Outros links': c.outros || '', 'Notas': c.notas || '', 'Vídeo final': c.final || ''
  };
  let id = c.id;
  if (id) {
    const i = t.linhas.findIndex(r => r['ID'] === id);
    if (i < 0) throw new Error('Cover não encontrado');
    escreverLinha_(sh, t, i + 2, obj);
  } else {
    const nums = t.linhas.map(r => parseInt(String(r['ID']).replace(/\D/g, ''), 10) || 0);
    id = 'C' + String(Math.max(0, ...nums) + 1).padStart(3, '0');
    obj['ID'] = id;
    acrescentar_(sh, t, [obj]);
  }
  return { ok: true, id };
}

function guardarPosicoes_(user, req) {
  const sh = folha_('Posições');
  const t = tabela_(sh);
  const apagar = [];
  t.linhas.forEach((r, i) => { if (r['ID do cover'] === req.cover) apagar.push(i + 2); });
  apagar.reverse().forEach(r => sh.deleteRow(r));
  acrescentar_(sh, tabela_(sh), (req.posicoes || []).filter(p => p.aluno && p.parte).map(p => ({ 'ID do cover': req.cover, 'ID do aluno': p.aluno, 'Membro / parte': p.parte, 'Notas': p.notas || '' })));
  return { ok: true };
}

function guardarEvento_(user, req) {
  const e = req.evento || {};
  if (!e.nome) throw new Error('Falta o nome do evento');
  const sh = folha_('Eventos');
  const t = tabela_(sh);
  const obj = {
    'Nome': e.nome, 'Tipo': e.tipo || 'Outro', 'Estado': e.estado || 'A considerar',
    'Data': e.data ? paraData_(e.data) : '', 'Data de fim': e.dataFim ? paraData_(e.dataFim) : '', 'Hora': e.hora || '',
    'Local': e.local || '', 'Prazo de inscrição': e.prazo ? paraData_(e.prazo) : '', 'Turmas': e.turmas || '', 'Covers': e.covers || '',
    'Requisitos': e.requisitos || '', 'Custo (€)': e.custo === '' || e.custo == null ? '' : Number(e.custo), 'Link': e.link || '', 'Notas': e.notas || ''
  };
  let id = e.id;
  if (id) {
    const i = t.linhas.findIndex(r => r['ID'] === id);
    if (i < 0) throw new Error('Evento não encontrado');
    escreverLinha_(sh, t, i + 2, obj);
  } else {
    const nums = t.linhas.map(r => parseInt(String(r['ID']).replace(/\D/g, ''), 10) || 0);
    id = 'E' + String(Math.max(0, ...nums) + 1).padStart(3, '0');
    obj['ID'] = id;
    acrescentar_(sh, t, [obj]);
  }
  return { ok: true, id };
}

function apagarPagamento_(user, req) {
  gestao_(user);
  const p = req.pagamento || {};
  const sh = folha_('Pagamentos');
  const t = tabela_(sh);
  const i = t.linhas.findIndex(r => r['ID do aluno'] === p.aluno && String(r['Mês (AAAA-MM)']) === p.mes &&
    Math.abs(Number(r['Valor (€)']) - Number(p.valor)) < 0.005 && (!p.data || r['Data de pagamento'] === p.data));
  if (i < 0) throw new Error('Pagamento não encontrado na folha');
  sh.deleteRow(i + 2);
  return { ok: true };
}

function apagarPorId_(nome, id) {
  if (!id) throw new Error('Falta o ID');
  const sh = folha_(nome);
  const i = tabela_(sh).linhas.findIndex(r => r['ID'] === id);
  if (i < 0) throw new Error('Não encontrado');
  sh.deleteRow(i + 2);
  return { ok: true };
}

function apagarCover_(user, req) {
  apagarPorId_('Covers', req.id);
  const sp = folha_('Posições');
  const tp = tabela_(sp);
  const linhas = [];
  tp.linhas.forEach((r, i) => { if (r['ID do cover'] === req.id) linhas.push(i + 2); });
  linhas.reverse().forEach(l => sp.deleteRow(l));
  return { ok: true };
}

/* ---------- Ajudantes de folha ---------- */

function folha_(nome) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nome);
  if (!sh) throw new Error('Falta o separador "' + nome + '"');
  return sh;
}

function tabela_(sh) {
  const ult = Math.max(sh.getLastRow(), 1), cols = sh.getLastColumn();
  const rng = sh.getRange(1, 1, ult, cols);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();
  const cab = vals[0].map(String);
  const col = {};
  cab.forEach((h, i) => { if (h) col[h] = i + 1; });
  const linhas = vals.slice(1).map((r, k) => {
    const o = {};
    cab.forEach((h, i) => {
      if (!h) return;
      // Horas: usa o texto mostrado na folha (evita desvios de fuso horário antigos)
      o[h] = (r[i] instanceof Date && r[i].getFullYear() < 1901) ? disp[k + 1][i] : normal_(r[i]);
    });
    return o;
  });
  return { cab, col, linhas };
}

function ler_(nome) { return tabela_(folha_(nome)).linhas; }

function lerSeExistir_(nome) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nome);
  return sh ? tabela_(sh).linhas : [];
}

function normal_(v) {
  if (v instanceof Date) {
    if (v.getFullYear() < 1901) return Utilities.formatDate(v, TZ, 'HH:mm');
    return fmtData_(v);
  }
  return v === null || v === undefined ? '' : (typeof v === 'string' ? v.trim() : v);
}

// Data de nascimento: completa (yyyy-mm-dd) ou só dia/mês (--mm-dd) quando falta o ano
function nasc_(v, extra) {
  if (v === '' || v === null || v === undefined) v = extra || '';
  if (v instanceof Date) return fmtData_(v);
  const t = String(v).replace(/^'/, '').trim();
  let m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return t.slice(0, 10);
  m = t.match(/^(\d{1,2})\D(\d{1,2})(?:\D(\d{4}))?$/) || t.match(/^(\d{2})(\d{2})$/);
  if (!m) return '';
  const p2 = x => ('0' + x).slice(-2);
  return (m[3] || '-') + '-' + p2(m[2]) + '-' + p2(m[1]);
}

function nascParaFolha_(n) {
  if (!n) return '';
  const m = String(n).match(/^--(\d{2})-(\d{2})/);
  if (m) return "'" + m[2] + '/' + m[1];
  return paraData_(n);
}

function fmtData_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }

function paraData_(txt) {
  const m = String(txt).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return txt;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
}

function num_(v) { const n = Number(v); return isNaN(n) || v === '' ? null : n; }

function acrescentar_(sh, t, objs) {
  if (!objs.length) return;
  const largura = t.cab.length;
  const linhas = objs.map(o => t.cab.map(h => (h in o ? o[h] : '')));
  const inicio = sh.getLastRow() + 1;
  sh.getRange(inicio, 1, linhas.length, largura).setValues(linhas);
}

function escreverLinha_(sh, t, linha, obj) {
  Object.keys(obj).forEach(h => { if (t.col[h]) sh.getRange(linha, t.col[h]).setValue(obj[h]); });
}

/* =====================================================================
 * RELATÓRIO MENSAL
 * Enviado no dia 1 de cada mês (sobre o mês anterior), a partir de dezembro de 2026.
 * Destinatários: coluna "Relatório mensal" = Sim no separador Professores.
 * Para ligar: correr "ativarRelatorioMensal" uma vez (autorizar quando pedir).
 * Para experimentar já: correr "relatorioTeste" (envia o mês atual só para quem o corre).
 * ===================================================================== */

const RELATORIO_DESDE = '2026-12';
const ALERTA_ASSIDUIDADE = 70;
const ESPACO_EUR_HORA = 15; // aluguer da sala, calculado automaticamente pelas aulas dadas

function ativarRelatorioMensal() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'relatorioMensal').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('relatorioMensal').timeBased().onMonthDay(1).atHour(9).inTimezone(TZ).create();
  garantirDespesas_();
  try { SpreadsheetApp.getActive().toast('Relatório mensal ativado: chega no dia 1 de cada mês, por volta das 9h.'); } catch (e) {}
}

function relatorioMensal() {
  const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1);
  const mes = Utilities.formatDate(d, TZ, 'yyyy-MM');
  if (mes < RELATORIO_DESDE) return;
  const para = ler_('Professores').filter(r => r['Ativo'] !== 'Não' && r['Relatório mensal'] === 'Sim').map(r => String(r['Email Google'] || '').trim()).filter(Boolean);
  if (para.length) gerarRelatorio_(mes, para);
}

function relatorioTeste() {
  const mes = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const eu = Session.getEffectiveUser().getEmail();
  gerarRelatorio_(mes, [eu], true);
  try { SpreadsheetApp.getActive().toast('Relatório de teste enviado para ' + eu); } catch (e) {}
}

function garantirDespesas_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName('Despesas')) return;
  const sh = ss.insertSheet('Despesas');
  const cab = ['Data', 'Descrição', 'Categoria', 'Valor (€)', 'Notas'];
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setBackground('#4A47C4').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
  sh.getRange('D2:D').setNumberFormat('#,##0.00 €');
  const cats = ['Espaço', 'Professora', 'Guarda-roupa', 'Material', 'Eventos e concursos', 'Transporte', 'Divulgação', 'Outros'];
  sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(cats, true).setAllowInvalid(true).build());
  [110, 260, 170, 100, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function nomeMes_(mes) {
  const M = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  const p = mes.split('-');
  return M[Number(p[1]) - 1] + ' de ' + p[0];
}
function mesMais_(mes, n) { const p = mes.split('-').map(Number); const d = new Date(p[0], p[1] - 1 + n, 1, 12); return Utilities.formatDate(d, TZ, 'yyyy-MM'); }
function eur_(v) { return (Math.round((v || 0) * 100) / 100).toFixed(2).replace('.', ',') + ' €'; }
function pct_(a, b) { return b ? Math.round(a * 100 / b) + '%' : '–'; }
function dm_(iso) { const p = String(iso || '').split('-'); return p.length >= 3 ? p[2] + '/' + p[1] : ''; }

function despesasDoMes_(mes) {
  return lerSeExistir_('Despesas').filter(x => x['Data'] && String(x['Data']).slice(0, 7) === mes && num_(x['Valor (€)']) != null)
    .map(x => ({ data: x['Data'], desc: x['Descrição'] || '', cat: x['Categoria'] || 'Outros', valor: num_(x['Valor (€)']) || 0, notas: x['Notas'] || '' }));
}

/** Calcula tudo o que o relatório precisa para um mês (AAAA-MM). */
function contasDoMes_(D, mes) {
  const turmas = D.turmas.filter(t => t.ativa);
  const nomeT = id => (D.turmas.find(t => t.id === id) || {}).nome || id;
  const ativos = D.alunos.filter(a => a.ativo);
  const nomeA = id => { const a = D.alunos.find(x => x.id === id); return a ? (a.stage || a.nome) : id; };
  const insc = D.inscricoes.filter(i => i.ativa);
  const pres = D.presencas.filter(p => String(p.data).slice(0, 7) === mes && p.estado);

  // Presenças por turma
  const porTurma = turmas.map(t => {
    const ps = pres.filter(p => p.turma === t.id);
    const aulas = new Set(ps.map(p => p.data + '|' + p.tipo));
    const rep = new Set(ps.filter(p => p.tipo === 'Reposição').map(p => p.data));
    const n = e => ps.filter(p => p.estado === e).length;
    return { nome: t.nome, aulas: aulas.size, reposicoes: rep.size, presente: n('Presente'), falta: n('Falta'), just: n('Justificada'), total: ps.length };
  });

  // Presenças por aluna
  const porAluna = ativos.map(a => {
    const ps = pres.filter(p => p.aluno === a.id);
    const n = e => ps.filter(p => p.estado === e).length;
    const ts = insc.filter(i => i.aluno === a.id).map(i => nomeT(i.turma)).join(', ');
    return { id: a.id, nome: a.stage || a.nome, turmas: ts, presente: n('Presente'), falta: n('Falta'), just: n('Justificada'), total: ps.length };
  }).filter(x => x.total || x.turmas).sort((a, b) => a.nome.localeCompare(b.nome, 'pt'));

  // Pagamentos: esperado vs recebido
  const pagMes = (D.pagamentos || []).filter(p => p.mes === mes);
  const pagamentos = ativos.map(a => {
    const ins = insc.filter(i => i.aluno === a.id);
    if (!ins.length) return null;
    let esperado = 0;
    ins.forEach(i => {
      const t = D.turmas.find(x => x.id === i.turma) || {};
      if (i.plano === 'Por aula') esperado += (t.precoAula || 0) * pres.filter(p => p.aluno === a.id && p.turma === i.turma && p.estado === 'Presente').length;
      else esperado += t.mensalidade || 0;
    });
    const ps = pagMes.filter(p => p.aluno === a.id);
    const pago = ps.reduce((s, p) => s + (p.valor || 0), 0);
    const metodo = Array.from(new Set(ps.map(p => p.metodo).filter(Boolean))).join(', ');
    const estado = esperado === 0 ? (pago ? 'Pago' : '–') : pago >= esperado ? 'Pago' : pago > 0 ? 'Parcial' : 'Em falta';
    return { nome: a.stage || a.nome, turmas: ins.map(i => nomeT(i.turma)).join(', '), plano: ins.map(i => i.plano).join(', '), esperado: esperado, pago: pago, metodo: metodo, estado: estado };
  }).filter(Boolean).sort((a, b) => a.nome.localeCompare(b.nome, 'pt'));
  // Pagamentos de alunas que já não estão ativas, mas pagaram este mês
  const idsAtivos = new Set(ativos.map(a => a.id));
  const outros = pagMes.filter(p => !idsAtivos.has(p.aluno)).reduce((s, p) => s + (p.valor || 0), 0);

  const receitas = pagMes.reduce((s, p) => s + (p.valor || 0), 0);
  const porAula = new Set(insc.filter(i => i.plano === 'Por aula').map(i => i.aluno));
  const recAula = pagMes.filter(p => idsAtivos.has(p.aluno) && porAula.has(p.aluno)).reduce((s, p) => s + (p.valor || 0), 0);
  const recMensal = receitas - recAula - outros;
  const fimMes = mes + '-31';
  const movPag = pagMes.map(p => ({ data: p.data && String(p.data) <= fimMes && String(p.data) >= mes + '-01' ? p.data : (p.data || mes + '-01'), mes: p.mes, nome: nomeA(p.aluno), metodo: p.metodo, valor: p.valor || 0 }));
  const esperado = pagamentos.reduce((s, p) => s + p.esperado, 0);
  // Aluguer do espaço: horas de aula dadas no mês × preço por hora (agrupado por dia)
  const horas = (i, f) => { const a = String(i || '').split(':'), b = String(f || '').split(':'); const h = (Number(b[0]) * 60 + Number(b[1] || 0) - Number(a[0]) * 60 - Number(a[1] || 0)) / 60; return h > 0 ? h : 0; };
  const sessoes = {};
  pres.forEach(p => { sessoes[p.data + '|' + p.turma + '|' + p.tipo] = { data: p.data, turma: p.turma, tipo: p.tipo }; });
  const hojeTxt = fmtData_(new Date());
  (D.calendario || []).filter(c => String(c.data).slice(0, 7) === mes && c.tipo !== 'Cancelada' && String(c.data) <= hojeTxt)
    .forEach(c => { const k = c.data + '|' + c.turma + '|' + (c.tipo || 'Normal'); sessoes[k] = Object.assign(sessoes[k] || {}, { data: c.data, turma: c.turma, tipo: c.tipo || 'Normal', inicio: c.inicio, fim: c.fim }); });
  const porDia = {};
  Object.keys(sessoes).forEach(k => {
    const x = sessoes[k]; const t = D.turmas.find(y => y.id === x.turma) || {};
    porDia[x.data] = (porDia[x.data] || 0) + horas(x.inicio || t.inicio, x.fim || t.fim);
  });
  const espaco = Object.keys(porDia).filter(d => porDia[d] > 0).sort().map(d => ({
    data: d, desc: 'Aluguer do espaço (' + String(porDia[d]).replace('.', ',') + ' h × ' + ESPACO_EUR_HORA + ' €)', cat: 'Espaço', valor: porDia[d] * ESPACO_EUR_HORA, notas: '', auto: true }));
  const despesas = espaco.concat(despesasDoMes_(mes));
  const totalDesp = despesas.reduce((s, x) => s + x.valor, 0);
  const porCat = {};
  despesas.forEach(x => { porCat[x.cat] = (porCat[x.cat] || 0) + x.valor; });

  const novas = D.alunos.filter(a => String(a.inscricao || '').slice(0, 7) === mes).map(a => a.stage || a.nome);
  const presTot = pres.filter(p => p.estado === 'Presente').length;

  return {
    mes: mes, ativos: ativos.length, novas: novas, porTurma: porTurma, porAluna: porAluna,
    aulas: porTurma.reduce((s, t) => s + t.aulas, 0), assiduidade: pct_(presTot, pres.length),
    pagamentos: pagamentos, receitas: receitas, outros: outros, esperado: esperado, recAula: recAula, recMensal: recMensal, movPag: movPag,
    emFalta: pagamentos.filter(p => p.estado === 'Em falta' || p.estado === 'Parcial'),
    despesas: despesas, totalDesp: totalDesp, porCat: porCat, saldo: receitas - totalDesp,
    nomeA: nomeA, nomeT: nomeT
  };
}

function gerarRelatorio_(mes, para, teste) {
  const D = dados_({ perfil: 'Gestão', nome: 'Relatório', email: '' });
  let logo = '';
  try {
    const r = UrlFetchApp.fetch('https://icy-academy.netlify.app/logo-branco.png', { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) logo = 'data:image/png;base64,' + Utilities.base64Encode(r.getBlob().getBytes());
  } catch (e) {}
  const R = htmlRelatorio_(D, mes, logo, teste);
  const pdf = Utilities.newBlob(R.html, 'text/html', 'relatorio.html').getAs('application/pdf').setName(R.titulo + '.pdf');

  // Guarda o PDF na pasta "Relatórios", ao lado da folha
  let url = '';
  const pai = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
  if (pai.hasNext()) {
    const base = pai.next();
    const it = base.getFoldersByName('Relatórios');
    const pasta = it.hasNext() ? it.next() : base.createFolder('Relatórios');
    url = pasta.createFile(pdf).getUrl();
  }
  MailApp.sendEmail({ to: para.join(','), subject: R.titulo, htmlBody: R.email.replace('{{URL}}', url), attachments: [pdf], name: 'ICY Academy' });
  return url;
}

/** Monta o relatório (PDF em HTML) e o corpo do email. */
function htmlRelatorio_(D, mes, logo, teste) {
  const C = contasDoMes_(D, mes);
  const A = contasDoMes_(D, mesMais_(mes, -1));
  const prox = mesMais_(mes, 1);
  const p2 = n => ('0' + n).slice(-2);
  const ultimo = new Date(Number(mes.slice(0, 4)), Number(mes.slice(5, 7)), 0).getDate();
  const periodo = '01/' + mes.slice(5, 7) + '/' + mes.slice(0, 4) + ' a ' + ultimo + '/' + mes.slice(5, 7) + '/' + mes.slice(0, 4);
  const emitido = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
  const nm = m => nomeMes_(m).split(' ')[0];
  const cap = t => t.charAt(0).toUpperCase() + t.slice(1);
  const h = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const n = v => (v < 0 ? '−' : '') + eur_(Math.abs(v || 0));
  const varPct = (a, b) => !b ? '–' : ((a - b) / Math.abs(b) * 100 >= 0 ? '+' : '−') + Math.abs(Math.round((a - b) / Math.abs(b) * 100)) + '%';

  // Acumulado do ano
  const ano = { rec: 0, desp: 0, porCat: {}, mens: 0, aula: 0, outros: 0 };
  for (let m = mes.slice(0, 4) + '-01'; m <= mes; m = mesMais_(m, 1)) {
    const X = m === mes ? C : m === A.mes ? A : contasDoMes_(D, m);
    ano.rec += X.receitas; ano.desp += X.totalDesp; ano.mens += X.recMensal; ano.aula += X.recAula; ano.outros += X.outros;
    Object.keys(X.porCat).forEach(k => { ano.porCat[k] = (ano.porCat[k] || 0) + X.porCat[k]; });
  }
  const cats = Array.from(new Set(Object.keys(C.porCat).concat(Object.keys(A.porCat)).concat(Object.keys(ano.porCat)))).sort();

  const ROXO = '#4A47C4', NAVY = '#05051F', CINZA = '#6B6B80', LINHA = '#E3E3EE';
  const css = `
    body{font-family:Arial,Helvetica,sans-serif;font-size:9.5pt;color:#1E1E2E;margin:0}
    table{border-collapse:collapse;width:100%}
    .sec{font-size:10.5pt;font-weight:bold;color:${ROXO};text-transform:uppercase;letter-spacing:.6px;border-bottom:2px solid ${ROXO};padding:0 0 4px;margin:22px 0 8px}
    .t td,.t th{padding:5px 6px;border-bottom:1px solid ${LINHA};vertical-align:top}
    .t th{font-size:8pt;color:${CINZA};text-transform:uppercase;letter-spacing:.4px;text-align:left;border-bottom:1.5px solid #B9B9CC;font-weight:bold}
    .r{text-align:right;white-space:nowrap}
    .sub td{font-weight:bold;background:#F4F4FA;border-top:1px solid #B9B9CC}
    .tot td{font-weight:bold;font-size:10.5pt;background:${NAVY};color:#FFFFFF;border:none}
    .grp td{font-weight:bold;color:${ROXO};padding-top:9px;border-bottom:none}
    .ind td{padding-left:18px}
    .mut{color:${CINZA}}
    .ok{color:#1C7C3A;font-weight:bold}.bad{color:#C0264F;font-weight:bold}.warn{color:#B7791F;font-weight:bold}
    .sec{page-break-after:avoid} tr{page-break-inside:avoid}
    .nota{font-size:8.5pt;color:${CINZA};margin:6px 0 0}
  `;
  let o = `<html><head><meta charset="utf-8"><style>${css}</style></head><body>`;

  // Cabeçalho
  o += `<table style="background:${NAVY};color:#fff"><tr>
    <td style="padding:16px 20px;width:90px">${logo ? `<img src="${logo}" style="width:72px;height:70px">` : `<b style="font-size:20pt">ICY</b>`}</td>
    <td style="padding:16px 0"><div style="font-size:17pt;font-weight:bold;letter-spacing:.5px">ICY ACADEMY</div><div style="font-size:9pt;color:#B9B8F5;margin-top:3px">Escola de dança K-pop · São Brás de Alportel</div></td>
    <td style="padding:16px 20px;text-align:right"><div style="font-size:8pt;color:#B9B8F5;text-transform:uppercase;letter-spacing:1px">Relatório mensal${teste ? ' · teste' : ''}</div>
      <div style="font-size:15pt;font-weight:bold;margin-top:2px">${cap(nomeMes_(mes))}</div>
      <div style="font-size:8pt;color:#B9B8F5;margin-top:4px">Período ${periodo}<br>Emitido em ${emitido}</div></td></tr></table>`;

  // Indicadores
  const kpi = (rot, val, sub, cor) => `<td style="width:25%;padding:10px 12px;border:1px solid ${LINHA};border-top:3px solid ${cor || ROXO}">
    <div style="font-size:7.5pt;color:${CINZA};text-transform:uppercase;letter-spacing:.5px">${rot}</div>
    <div style="font-size:14pt;font-weight:bold;margin-top:3px">${val}</div><div style="font-size:8pt;color:${CINZA};margin-top:2px">${sub}</div></td>`;
  o += `<table style="margin-top:14px"><tr>
    ${kpi('Receitas', n(C.receitas), varPct(C.receitas, A.receitas) + ' vs ' + nm(A.mes))}
    ${kpi('Despesas', n(C.totalDesp), varPct(C.totalDesp, A.totalDesp) + ' vs ' + nm(A.mes))}
    ${kpi('Resultado do mês', n(C.saldo), 'Margem ' + (C.receitas ? Math.round(C.saldo / C.receitas * 100) + '%' : '–'), C.saldo >= 0 ? '#1C7C3A' : '#C0264F')}
    ${kpi('Cobrança', pct_(C.receitas - C.outros, C.esperado), n(C.receitas - C.outros) + ' de ' + n(C.esperado), C.emFalta.length ? '#B7791F' : '#1C7C3A')}
  </tr></table>`;

  // 1. Demonstração de resultados
  o += `<div class="sec">1. Demonstração de resultados</div><table class="t">
    <tr><th></th><th class="r">${cap(nm(mes))}</th><th class="r">${cap(nm(A.mes))}</th><th class="r">Variação</th><th class="r">Acumulado ${mes.slice(0, 4)}</th></tr>
    <tr class="grp"><td colspan="5">Receitas</td></tr>
    <tr class="ind"><td>Mensalidades</td><td class="r">${n(C.recMensal)}</td><td class="r">${n(A.recMensal)}</td><td class="r">${varPct(C.recMensal, A.recMensal)}</td><td class="r">${n(ano.mens)}</td></tr>
    <tr class="ind"><td>Aulas avulsas</td><td class="r">${n(C.recAula)}</td><td class="r">${n(A.recAula)}</td><td class="r">${varPct(C.recAula, A.recAula)}</td><td class="r">${n(ano.aula)}</td></tr>
    ${(C.outros || A.outros || ano.outros) ? `<tr class="ind"><td>Outras receitas</td><td class="r">${n(C.outros)}</td><td class="r">${n(A.outros)}</td><td class="r">${varPct(C.outros, A.outros)}</td><td class="r">${n(ano.outros)}</td></tr>` : ''}
    <tr class="sub"><td>Total de receitas</td><td class="r">${n(C.receitas)}</td><td class="r">${n(A.receitas)}</td><td class="r">${varPct(C.receitas, A.receitas)}</td><td class="r">${n(ano.rec)}</td></tr>
    <tr class="grp"><td colspan="5">Despesas</td></tr>
    ${cats.length ? cats.map(k => `<tr class="ind"><td>${h(k)}</td><td class="r">${n(C.porCat[k] || 0)}</td><td class="r">${n(A.porCat[k] || 0)}</td><td class="r">${varPct(C.porCat[k] || 0, A.porCat[k] || 0)}</td><td class="r">${n(ano.porCat[k] || 0)}</td></tr>`).join('')
      : `<tr class="ind"><td class="mut" colspan="5">Sem despesas registadas</td></tr>`}
    <tr class="sub"><td>Total de despesas</td><td class="r">${n(C.totalDesp)}</td><td class="r">${n(A.totalDesp)}</td><td class="r">${varPct(C.totalDesp, A.totalDesp)}</td><td class="r">${n(ano.desp)}</td></tr>
    <tr class="tot"><td>Resultado líquido</td><td class="r">${n(C.saldo)}</td><td class="r">${n(A.saldo)}</td><td class="r">${varPct(C.saldo, A.saldo)}</td><td class="r">${n(ano.rec - ano.desp)}</td></tr>
  </table>`;

  // 2. Extrato de movimentos
  const mov = C.movPag.map(x => ({ data: x.data, desc: 'Mensalidade ' + nm(x.mes) + ' · ' + x.nome, cat: x.metodo || 'Recebimento', ent: x.valor, sai: 0 }))
    .concat(C.despesas.map(x => ({ data: x.data, desc: x.desc, cat: x.cat, ent: 0, sai: x.valor })))
    .sort((a, b) => String(a.data).localeCompare(String(b.data)));
  let saldo = 0;
  o += `<div class="sec">2. Extrato de movimentos</div><table class="t">
    <tr><th style="width:52px">Data</th><th>Descrição</th><th>Categoria / meio</th><th class="r">Entrada</th><th class="r">Saída</th><th class="r">Saldo</th></tr>
    ${mov.length ? mov.map(x => { saldo += x.ent - x.sai; return `<tr><td>${dm_(x.data) || '–'}</td><td>${h(x.desc)}</td><td class="mut">${h(x.cat)}</td><td class="r">${x.ent ? n(x.ent) : ''}</td><td class="r">${x.sai ? n(x.sai) : ''}</td><td class="r">${n(saldo)}</td></tr>`; }).join('')
      : `<tr><td colspan="6" class="mut">Sem movimentos neste mês</td></tr>`}
    <tr class="sub"><td></td><td>Totais do período</td><td></td><td class="r">${n(C.receitas)}</td><td class="r">${n(C.totalDesp)}</td><td class="r">${n(C.saldo)}</td></tr>
  </table>`;

  // 3. Conta corrente das alunas
  const est = e => e === 'Pago' ? '<span class="ok">Pago</span>' : e === 'Em falta' ? '<span class="bad">Em falta</span>' : e === 'Parcial' ? '<span class="warn">Parcial</span>' : e;
  o += `<div class="sec">3. Mensalidades · conta corrente</div><table class="t">
    <tr><th>Aluna</th><th>Turma</th><th>Plano</th><th class="r">Devido</th><th class="r">Pago</th><th class="r">Em dívida</th><th class="r">Estado</th></tr>
    ${C.pagamentos.map(x => `<tr><td>${h(x.nome)}</td><td class="mut">${h(x.turmas)}</td><td class="mut">${h(x.plano)}</td><td class="r">${n(x.esperado)}</td><td class="r">${n(x.pago)}</td><td class="r">${x.esperado > x.pago ? n(x.esperado - x.pago) : '–'}</td><td class="r">${est(x.estado)}</td></tr>`).join('')}
    <tr class="sub"><td colspan="3">Total (${C.pagamentos.length} alunas)</td><td class="r">${n(C.esperado)}</td><td class="r">${n(C.receitas - C.outros)}</td><td class="r">${n(C.emFalta.reduce((s, x) => s + x.esperado - x.pago, 0))}</td><td class="r">${pct_(C.receitas - C.outros, C.esperado)}</td></tr>
  </table>`;

  // 4. Assiduidade
  o += `<div class="sec" style="page-break-before:always">4. Assiduidade</div><table class="t">
    <tr><th>Turma</th><th class="r">Aulas</th><th class="r">Reposições</th><th class="r">Presenças</th><th class="r">Faltas</th><th class="r">Justificadas</th><th class="r">Assiduidade</th></tr>
    ${C.porTurma.map(t => `<tr><td>${h(t.nome)}</td><td class="r">${t.aulas}</td><td class="r">${t.reposicoes}</td><td class="r">${t.presente}</td><td class="r">${t.falta}</td><td class="r">${t.just}</td><td class="r"><b>${pct_(t.presente, t.total)}</b></td></tr>`).join('')}
    <tr class="sub"><td>Total</td><td class="r">${C.aulas}</td><td class="r">${C.porTurma.reduce((s, t) => s + t.reposicoes, 0)}</td><td class="r">${C.porTurma.reduce((s, t) => s + t.presente, 0)}</td><td class="r">${C.porTurma.reduce((s, t) => s + t.falta, 0)}</td><td class="r">${C.porTurma.reduce((s, t) => s + t.just, 0)}</td><td class="r">${C.assiduidade}</td></tr>
  </table>`;
  // por aluna em duas colunas
  const linhasA = C.porAluna.map(a => {
    const pc = a.total ? Math.round(a.presente * 100 / a.total) : null;
    const cls = pc == null ? 'mut' : pc < ALERTA_ASSIDUIDADE ? 'bad' : '';
    return `<tr><td>${h(a.nome)}</td><td class="r mut">${a.presente}/${a.falta}/${a.just}</td><td class="r ${cls}">${pc == null ? '–' : pc + '%'}</td></tr>`;
  });
  const meio = Math.ceil(linhasA.length / 2);
  const col = ls => `<table class="t"><tr><th>Aluna</th><th class="r">P / F / J</th><th class="r">%</th></tr>${ls.join('')}</table>`;
  o += `<table style="margin-top:10px"><tr><td style="width:49%;vertical-align:top">${col(linhasA.slice(0, meio))}</td><td style="width:2%"></td><td style="width:49%;vertical-align:top">${col(linhasA.slice(meio))}</td></tr></table>`;
  o += `<div class="nota">P = presenças · F = faltas · J = faltas justificadas. A vermelho, assiduidade abaixo de ${ALERTA_ASSIDUIDADE}%.</div>`;

  // 5. Previsão
  const aulas = D.calendario.filter(c => String(c.data).slice(0, 7) === prox).sort((a, c) => String(a.data).localeCompare(String(c.data)));
  const normais = aulas.filter(a => a.tipo !== 'Cancelada');
  const recPrev = C.pagamentos.filter(x => x.plano.indexOf('Mensal') >= 0).reduce((s, x) => s + x.esperado, 0);
  const md = v => { const m = String(v || '').match(/-(\d{2})-(\d{2})$/); return m ? { m: m[1], d: m[2] } : null; };
  const anivs = D.alunos.filter(a => a.ativo && md(a.nascimento) && md(a.nascimento).m === prox.slice(5, 7)).sort((a, c) => md(a.nascimento).d.localeCompare(md(c.nascimento).d));
  const evs = D.eventos.filter(e => e.estado !== 'Não vamos' && e.estado !== 'Concluído' && ((e.data && String(e.data).slice(0, 7) === prox) || (e.prazo && String(e.prazo).slice(0, 7) === prox)));
  const cvs = D.covers.filter(c => c.estado === 'Em ensaio' || (c.apresentacao && String(c.apresentacao).slice(0, 7) === prox));
  const item = (k, v) => `<tr><td style="width:150px" class="mut">${k}</td><td>${v}</td></tr>`;
  o += `<div class="sec">5. Perspetiva para ${nm(prox)}</div><table class="t">
    ${item('Receita prevista', '<b>' + n(recPrev) + '</b> em mensalidades (alunas ativas, plano mensal)')}
    ${item('Aulas marcadas', normais.length + (aulas.length - normais.length ? ' · ' + (aulas.length - normais.length) + ' cancelada(s): ' + aulas.filter(a => a.tipo === 'Cancelada').map(a => dm_(a.data) + ' ' + C.nomeT(a.turma)).join(', ') : ''))}
    ${aulas.filter(a => a.tipo === 'Reposição' || (a.tipo && a.tipo !== 'Normal' && a.tipo !== 'Cancelada')).length ? item('Aulas especiais', aulas.filter(a => a.tipo && a.tipo !== 'Normal' && a.tipo !== 'Cancelada').map(a => dm_(a.data) + ' ' + C.nomeT(a.turma) + ' (' + h(a.tipo) + ')').join(', ')) : ''}
    ${item('Eventos e prazos', evs.length ? evs.map(e => '<b>' + h(e.nome) + '</b>' + (e.prazo && String(e.prazo).slice(0, 7) === prox ? ' · prazo de inscrição ' + dm_(e.prazo) : '') + (e.data && String(e.data).slice(0, 7) === prox ? ' · ' + dm_(e.data) : '') + ' <span class="mut">(' + h(e.estado) + ')</span>').join('<br>') : '<span class="mut">Nada marcado</span>')}
    ${item('Covers', cvs.length ? cvs.map(c => h(c.grupo) + ' – ' + h(c.musica) + (c.apresentacao ? ' · apresentação ' + dm_(c.apresentacao) : '') + ' <span class="mut">(' + h(c.estado) + ')</span>').join('<br>') : '<span class="mut">Nenhum em curso</span>')}
    ${item('Aniversários', anivs.length ? anivs.map(a => h(a.stage || a.nome) + ' (' + md(a.nascimento).d + '/' + md(a.nascimento).m + ')').join(', ') : '<span class="mut">Nenhum</span>')}
  </table>`;

  o += `<div style="margin-top:26px;padding-top:8px;border-top:1px solid ${LINHA};font-size:7.5pt;color:${CINZA}">
    ICY Academy · Relatório gerado automaticamente a partir da app de gestão em ${emitido}. Valores em euros. Receitas por mês de referência da mensalidade; despesas pela data registada no separador "Despesas".</div>`;
  o += '</body></html>';

  const titulo = 'ICY Academy · Relatório de ' + nomeMes_(mes) + (teste ? ' (teste)' : '');
  const linha = (k, v) => `<tr><td style="padding:6px 18px 6px 0;color:#6B6B80;border-bottom:1px solid #EEE">${k}</td><td style="padding:6px 0;border-bottom:1px solid #EEE;text-align:right">${v}</td></tr>`;
  const email = `<div style="font-family:Arial,sans-serif;font-size:14px;color:#1E1E2E;max-width:520px">
    <div style="background:${NAVY};color:#fff;padding:14px 18px;border-radius:8px 8px 0 0"><div style="font-size:11px;color:#B9B8F5;text-transform:uppercase;letter-spacing:1px">ICY Academy · Relatório mensal</div><div style="font-size:20px;font-weight:bold;margin-top:2px">${cap(nomeMes_(mes))}</div></div>
    <div style="border:1px solid #E3E3EE;border-top:none;padding:14px 18px;border-radius:0 0 8px 8px">
    <table style="border-collapse:collapse;width:100%">
    ${linha('Receitas', n(C.receitas))}${linha('Despesas', n(C.totalDesp))}${linha('<b>Resultado do mês</b>', '<b>' + n(C.saldo) + '</b>')}
    ${linha('Cobrança', pct_(C.receitas - C.outros, C.esperado) + (C.emFalta.length ? ' · em falta: ' + C.emFalta.map(x => h(x.nome)).join(', ') : ''))}
    ${linha('Assiduidade', C.assiduidade + ' em ' + C.aulas + ' aulas')}
    ${linha('Alunas ativas', C.ativos + (C.novas.length ? ' (+' + C.novas.length + (C.novas.length === 1 ? ' nova)' : ' novas)') : ''))}
    </table>
    <p style="margin:14px 0 0;font-size:13px">O relatório completo segue em PDF, em anexo. <a href="{{URL}}" style="color:${ROXO}">Ver no Google Drive</a></p></div></div>`;
  return { html: o, email: email, titulo: titulo };
}
