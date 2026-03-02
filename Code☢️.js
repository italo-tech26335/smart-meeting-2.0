const ID_PLANILHA = '1rjcUe9iGoguQVNcmwRHoxICmWB44q3tv5ARdp58pqxo';

/** =====================================================================
 *                          ABAS DA PLANILHAss
=========================================================================*/

const NOME_ABA_SETORES = 'Setores'
const NOME_ABA_PROJETOS = 'Projetos';
const NOME_ABA_ETAPAS = 'Atividades';
const NOME_ABA_RESPONSAVEIS = 'Responsaveis';
const NOME_ABA_DEPENDENCIAS = 'Dependencias';
const NOME_ABA_PRIORIDADES = 'PrioridadesResponsavel';
// NOME_ABA_PERMISSOES removido — sistema substituído por Auth.js
const NOME_ABA_LIXEIRA = 'LixeiraProjetos';

/** Reuniões */

const NOME_ABA_REUNIOES = 'Reuniões';

/** =====================================================================
 *                          COLUNAS DAS ABAS
=========================================================================*/

// COLUNAS_PERMISSOES removido — sistema substituído por Auth.js (COLUNAS_USUARIOS)

const COLUNAS_PROJETOS = {
  ID: 0,
  NOME: 1,
  DESCRICAO: 2,
  TIPO: 3,
  PARA_QUEM: 4,
  STATUS: 5,
  PRIORIDADE: 6,
  LINK: 7,
  GRAVIDADE: 8,
  URGENCIA: 9,
  ESFORCO: 10,
  SETOR: 11,
  PILAR: 12,
  RESPONSAVEIS_IDS: 13,
  VALOR_PRIORIDADE: 14,
  DATA_INICIO: 15,
  DATA_FIM: 16,
  DATA_CRIACAO: 17,
  DATA_ULTIMA_MODIFICACAO: 18
};

const COLUNAS_RESPONSAVEIS = {
  ID: 0,
  NOME: 1,
  EMAIL: 2,
  CARGO: 3,
};

const COLUNAS_ETAPAS = {
  ID: 0,
  PROJETO_ID: 1,
  RESPONSAVEIS_IDS: 2,
  NOME: 3,
  O_QUE_FAZER: 4,
  STATUS: 5
};

const COLUNAS_DEPENDENCIAS = {
  ID: 0,
  ETAPA_ORIGEM_ID: 1,
  ORIGEM_ANCHOR: 2,
  ETAPA_DESTINO_ID: 3,
  DESTINO_ANCHOR: 4
};

const COLUNAS_SETORES = {
  ID: 0,
  NOME: 1,
  DESCRICAO: 2,
  RESPONSAVEIS_IDS: 3
};

const COLUNAS_PRIORIDADES = {
  ID: 0,
  RESPONSAVEL_ID: 1,
  TIPO_ITEM: 2,     
  ITEM_ID: 3,             
  ORDEM_PRIORIDADE: 4,   
  PROJETO_REFERENCIA: 5   
};

/** Reuniões */

const COLUNAS_REUNIOES = {
  ID: 0,
  TITULO: 1,
  DATA_INICIO: 2,
  DATA_FIM: 3,
  DURACAO: 4,
  STATUS: 5,
  PARTICIPANTES: 6,
  TRANSCRICAO: 7,
  ATA: 8,
  SUGESTOES_IA: 9,
  LINK_AUDIO: 10,
  LINK_ATA: 11,
  EMAILS_ENVIADOS: 12,
  PROJETOS_IMPACTADOS: 13,
  ETAPAS_IMPACTADAS: 14
};

/** =====================================================================
 *                          OUTROS
=========================================================================*/

const STATUS_ETAPAS = {
  A_FAZER: 'A Fazer',
  EM_ANDAMENTO: 'Em Andamento',
  BLOQUEADA: 'Bloqueada',
  CONCLUIDA: 'Conclu?da'
};

const TIPOS_PROJETO = {
  MELHORIA: 'Melhoria',
  CORRECAO: 'Corre?o',
  NOVA_IMPLEMENTACAO: 'Nova Implementa?o'
};

const URGENCIA_PENDENCIA = {
  ALTA: 'alta',
  MEDIA: 'media',
  BAIXA: 'baixa'
};

const CORES_URGENCIA = {
  alta: { cor: '#dc2626', bgCor: 'rgba(220, 38, 38, 0.12)', icone: 'fa-exclamation-circle' },
  media: { cor: '#ca8a04', bgCor: 'rgba(202, 138, 4, 0.12)', icone: 'fa-clock' },
  baixa: { cor: '#65a30d', bgCor: 'rgba(101, 163, 13, 0.12)', icone: 'fa-check-circle' }
};

const NIVEIS_ACESSO = {
  ADMIN: 'admin',
  GESTOR: 'gestor',
  COLABORADOR: 'colaborador'
};

const MODELO_GEMINI = 'gemini-2.5-flash';
const MODELO_GEMINI_FLASH = 'gemini-2.5-flash';
const PREFIXO_CHAVE_GEMINI = 'GEMINI_API_KEY_';
const CHAVE_PROPRIEDADE_GEMINI = 'GEMINI_API_KEY_PROJETO_EDITOR';
const QUANTIDADE_CHAVES_GEMINI = 5;
const ID_PASTA_DRIVE_REUNIOES = '1h10xPrk1VVFUdgkc6n8-kxYZ9z2X1TGh';
const PARTICIPANTES_CADASTRADOS = [];
const EMAILS_DESTINATARIOS_PADRAO = [
  'napa02@christus.com.br',
  'napa05@christus.com.br',
  'napa11@christus.com.br',
  'napa12@christus.com.br',
  'napa13@christus.com.br',
  'napa14@christus.com.br',
  'napa16@christus.com.br'
];

function doGet(e) {
  try {
    const pagina = (e?.parameter?.pagina || '').toString().trim().toLowerCase();
    const token  = (e?.parameter?.token  || '').toString().trim();
    const rt     = (e?.parameter?.rt     || '').toString().trim(); // token de reset

    // ── Página de login (sempre acessível) ──────────────────────────────
    if (pagina === 'login' || pagina === '') {
      return _servirLogin('');
    }

    // ── Reset de senha (acessível sem sessão) ────────────────────────────
    if (pagina === 'reset') {
      return _servirReset(rt);
    }

    // ── Validar sessão ───────────────────────────────────────────────────
    const sessao = _obterSessao(token);
    if (!sessao) {
      return _servirLogin('Sessão expirada. Faça login novamente.');
    }

    // ── Painel admin (apenas perfil admin) ───────────────────────────────
    if (pagina === 'admin') {
      if (sessao.perfil !== 'admin') return _servirLogin('Acesso negado.');
      const tmpl = HtmlService.createTemplateFromFile('PaginaAdmin');
      tmpl.sessaoToken   = token;
      tmpl.usuarioNome   = sessao.nome;
      tmpl.usuarioPerfil = sessao.perfil;
      return tmpl.evaluate()
        .setTitle('Smart Meeting - Painel Admin')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    // ── Renovar sessão (sliding session) ─────────────────────────────────
    _renovarSessao(token, sessao);

    // ── Roteamento de páginas ─────────────────────────────────────────────
    let nomeArquivoHtml, titulo;

    if (pagina === 'projeto') {
      nomeArquivoHtml = 'PaginaProjetoDetalhe🟡';
      titulo = 'Smart Meeting - Projetos';
    } else if (pagina === 'relatorios') {
      nomeArquivoHtml = 'PaginaRelatorios';
      titulo = 'Smart Meeting - Relatórios';
    } else {
      nomeArquivoHtml = 'PaginaReunioes▶️';
      titulo = 'Smart Meeting - Reuniões';
    }

    const tmpl = HtmlService.createTemplateFromFile(nomeArquivoHtml);
    tmpl.sessaoToken   = token;
    tmpl.usuarioNome   = sessao.nome;
    tmpl.usuarioPerfil = sessao.perfil;

    return tmpl.evaluate()
      .setTitle(titulo)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (erro) {
    Logger.log('ERRO no doGet: ' + erro.toString());
    return HtmlService.createHtmlOutput('Erro ao carregar página: ' + erro.message);
  }
}

function _servirLogin(mensagemErro) {
  const tmpl = HtmlService.createTemplateFromFile('Login');
  tmpl.mensagemErroUrl = mensagemErro || '';
  tmpl.tokenReset = '';
  return tmpl.evaluate()
    .setTitle('Smart Meeting - Login')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _servirReset(rt) {
  const tmpl = HtmlService.createTemplateFromFile('Login');
  tmpl.mensagemErroUrl = '';
  tmpl.tokenReset = rt || '';
  return tmpl.evaluate()
    .setTitle('Smart Meeting - Redefinir Senha')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function abrirPaginaRelatorios() {
  const url = ScriptApp.getService().getUrl() + '?pagina=relatorios';
  const html = `<script>window.open('${url}', '_blank');google.script.host.close();</script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), 'Abrindo...');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Grava timestamps de criação e/ou modificação em uma linha da planilha.
 * @param {Sheet} aba - Objeto Sheet do Google Sheets
 * @param {number} linhaIndex - Índice da linha (base 1)
 * @param {number} colCriacao - Índice da coluna DATA_CRIACAO (base 0); -1 para ignorar
 * @param {number} colModificacao - Índice da coluna DATA_ULTIMA_MODIFICACAO (base 0); -1 para ignorar
 * @param {boolean} ehNovo - Se true, grava também a data de criação
 */
function gravarTimestamp(aba, linhaIndex, colCriacao, colModificacao, ehNovo) {
  const agora = new Date();
  if (ehNovo && colCriacao >= 0) {
    aba.getRange(linhaIndex, colCriacao + 1).setValue(agora);
  }
  if (colModificacao >= 0) {
    aba.getRange(linhaIndex, colModificacao + 1).setValue(agora);
  }
}

function obterUrlWebApp() {
  return ScriptApp.getService().getUrl();
}

// Alias usado por PaginaReunioes para configurar links de navegação
function obterUrlBaseWebApp() {
  return ScriptApp.getService().getUrl();
}

function abrirPaginaReunioes() {
  const url = ScriptApp.getService().getUrl();
  const html = `<script>window.open('${url}', '_blank');google.script.host.close();</script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), 'Abrindo...');
}

function abrirPaginaProjetos() {
  const url = ScriptApp.getService().getUrl() + '?pagina=projetos';
  const html = `<script>window.open('${url}', '_blank');google.script.host.close();</script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), 'Abrindo...');
}

function obterPlanilhaAtiva() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(ID_PLANILHA);
  } catch (e) {
    Logger.log('ERRO ao obter planilha: ' + e.toString());
    return SpreadsheetApp.openById(ID_PLANILHA);
  }
}

function obterPlanilha() {
  return obterPlanilhaAtiva();
}

const CACHE_TTL_PADRAO_SEGUNDOS = 120;

function _chaveCacheAba(nomeAba) {
  return 'aba_cache::' + nomeAba;
}

function _cacheGetJson(chave) {
  try {
    const valor = CacheService.getScriptCache().get(chave);
    return valor ? JSON.parse(valor) : null;
  } catch (e) {
    return null;
  }
}

function _cachePutJson(chave, valor, ttlSegundos) {
  try {
    CacheService.getScriptCache().put(chave, JSON.stringify(valor), ttlSegundos || CACHE_TTL_PADRAO_SEGUNDOS);
  } catch (e) {
    //     // Falha de cache n?o deve interromper fluxo.
  }
}

function limparCacheAba(nomeAba) {
  try {
    CacheService.getScriptCache().remove(_chaveCacheAba(nomeAba));
  } catch (e) {
    // Ignora erro de cache.
  }
}

function obterDadosAbaComCache(nomeAba, ttlSegundos) {
  const chave = _chaveCacheAba(nomeAba);
  const cache = _cacheGetJson(chave);
  if (cache && Array.isArray(cache)) return cache;

  const aba = obterAba(nomeAba);
  const dados = aba ? aba.getDataRange().getValues() : [];
  _cachePutJson(chave, dados, ttlSegundos || CACHE_TTL_PADRAO_SEGUNDOS);
  return dados;
}

function obterAba(nomeAba) {
  const planilha = obterPlanilha();
  let aba = planilha.getSheetByName(nomeAba);
  if (!aba) {
    aba = planilha.insertSheet(nomeAba);
    inicializarCabecalhoAba(aba, nomeAba);
  }
  return aba;
}

function inicializarCabecalhoAba(aba, nomeAba) {
  const cabecalhos = {
    [NOME_ABA_PROJETOS]:      ['ID', 'Nome', 'Descricao', 'Tipo', 'ParaQuem', 'Status', 'Prioridade', 'Link', 'Gravidade', 'Urgencia', 'Esforco', 'Setor', 'Pilar', 'ResponsaveisIds', 'ValorPrioridade', 'DataInicio', 'DataFim'],
    [NOME_ABA_RESPONSAVEIS]:  ['ID', 'Nome', 'Email', 'Cargo'],
    [NOME_ABA_ETAPAS]:        ['ID', 'ProjetoId', 'ResponsaveisIds', 'Nome', 'OQueFazer', 'Status'],
    [NOME_ABA_DEPENDENCIAS]:  ['ID', 'EtapaOrigemId', 'OrigemAnchor', 'EtapaDestinoId', 'DestinoAnchor'],
    [NOME_ABA_REUNIOES]:      ['ID', 'Titulo', 'DataInicio', 'DataFim', 'DuracaoMin', 'Status', 'Participantes', 'Transcricao', 'Ata', 'SugestoesIA', 'LinkAudio', 'LinkAta', 'EmailsEnviados', 'ProjetosImpactados', 'EtapasCriadasOuAlteradas'],
    [NOME_ABA_SETORES]:       ['ID', 'Nome', 'Descricao', 'Cor'],
    [NOME_ABA_PRIORIDADES]:   ['ID', 'ResponsavelId', 'TipoItem', 'ItemId', 'OrdemPrioridade', 'ProjetoReferencia']
  };

  if (cabecalhos[nomeAba]) {
    aba.getRange(1, 1, 1, cabecalhos[nomeAba].length).setValues([cabecalhos[nomeAba]]);
    aba.getRange(1, 1, 1, cabecalhos[nomeAba].length).setFontWeight('bold');
  }
}


function gerarId() { return Date.now().toString(36) + Math.random().toString(36).substring(2, 9); }

function obterEmailUsuario() { 
  try { return Session.getActiveUser().getEmail(); } 
  catch (e) { 
    Logger.log('ERRO ao obter email: ' + e.toString());
    return ''; 
  } 
}

// verificarPermissaoEdicao e exigirPermissaoEdicao removidos — auth via token

function salvarChavesGemini(chaves) {
  try {
    const propriedades = PropertiesService.getScriptProperties();
    let chavesSalvas = 0;
    for (let i = 1; i <= QUANTIDADE_CHAVES_GEMINI; i++) {
      const chave = chaves['chave' + i];
      if (chave && chave.trim() !== '') {
        propriedades.setProperty(PREFIXO_CHAVE_GEMINI + i, chave.trim());
        chavesSalvas++;
      } else {
        propriedades.deleteProperty(PREFIXO_CHAVE_GEMINI + i);
      }
    }
    return { sucesso: true, mensagem: `${chavesSalvas} chave(s) salva(s)!` };
  } catch (erro) {
    Logger.log('ERRO salvar chaves: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function temChaveApiConfigurada() {
  const propriedades = PropertiesService.getScriptProperties();
  
  // Verificar chaves numeradas (1 a 5)
  for (let i = 1; i <= QUANTIDADE_CHAVES_GEMINI; i++) {
    const chave = propriedades.getProperty(PREFIXO_CHAVE_GEMINI + i);
    if (chave && chave.trim() !== '') return true;
  }
  
  //   // Verificar chave espec?fica do projeto editor
  const chaveProjeto = propriedades.getProperty(CHAVE_PROPRIEDADE_GEMINI);
  if (chaveProjeto && chaveProjeto.trim() !== '') return true;
  
  return false;
}

function abrirConfiguracoesChaves() {
  const html = HtmlService.createHtmlOutput(`
    <style>body{font-family:Arial;padding:20px}input{width:100%;padding:8px;margin:5px 0 15px;box-sizing:border-box}button{background:#4a86e8;color:white;padding:10px 20px;border:none;cursor:pointer}</style>
    <p>Configure suas chaves API do Gemini:</p>
    <input type=\"password\" id=\"c1\" placeholder=\"Chave 1\">
    <input type=\"password\" id=\"c2\" placeholder=\"Chave 2\">
    <input type=\"password\" id=\"c3\" placeholder=\"Chave 3\">
    <button onclick=\"salvar()\">Salvar</button>
    <script>function salvar(){google.script.run.withSuccessHandler(r=>{alert(r.mensagem);google.script.host.close();}).salvarChavesGemini({chave1:document.getElementById('c1').value,chave2:document.getElementById('c2').value,chave3:document.getElementById('c3').value});}</script>
  `).setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '🔑 Configurar Chaves API');
}

function obterChavesGeminiMascaradas() {
  const props = PropertiesService.getScriptProperties();
  const res = {};
  for (let i = 1; i <= QUANTIDADE_CHAVES_GEMINI; i++) {
    const k = props.getProperty(PREFIXO_CHAVE_GEMINI + i);
    if (k) res['chave' + i] = k.substring(0, 4) + '...' + k.substring(k.length - 4);
  }
  return res;
}

function obterEmailsDestinatariosPadrao() {
  try {
    return { sucesso: true, emails: EMAILS_DESTINATARIOS_PADRAO || [] };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message, emails: [] };
  }
}

function obterChaveGeminiProjeto() {
  return PropertiesService.getScriptProperties().getProperty(CHAVE_PROPRIEDADE_GEMINI) || 'AIzaSyD237PQ4GwuChutMZ4HDQ9i15m-4Y6Id4p';
}

function salvarChaveGeminiProjeto(chave) {
  try {
    if (!chave || chave.trim() === '') {
      return { sucesso: false, mensagem: 'Chave inv?lida' };
    }
    PropertiesService.getScriptProperties().setProperty(CHAVE_PROPRIEDADE_GEMINI, chave.trim());
    return { sucesso: true, mensagem: 'Chave salva com sucesso!' };
  } catch (e) {
    Logger.log('Erro ao salvar chave projeto: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

// ── Sistema de permissões granular removido — substituído por Auth.js ──────────
// As funções abaixo são helpers de conveniência baseados no novo sistema de perfis.

function ehAdministrador(token) {
  try {
    var sessao = _obterSessao(token || '');
    return sessao && sessao.perfil === 'admin';
  } catch (e) { return false; }
}

function podeEditar(token) {
  try {
    var sessao = _obterSessao(token || '');
    return sessao && (sessao.perfil === 'admin' || sessao.perfil === 'usuario');
  } catch (e) { return false; }
}

function parseArrayString(valor) {
  if (!valor) return [];
  if (Array.isArray(valor)) return valor;
  return valor.toString().split(',').map(function(v) { return v.trim(); }).filter(function(v) { return v !== ''; });
}

function obterProjetoPorIdSimples(projetoId) {
  try {
    const dados = obterDadosAbaComCache(NOME_ABA_PROJETOS);
    if (!dados || dados.length <= 1) return null;
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID] === projetoId) {
        const rawIds = dados[i][COLUNAS_PROJETOS.RESPONSAVEIS_IDS] || '';
        return {
          id: dados[i][COLUNAS_PROJETOS.ID],
          nome: dados[i][COLUNAS_PROJETOS.NOME],
          setor: dados[i][COLUNAS_PROJETOS.SETOR],
          responsaveisIds: rawIds.toString().split(',').map(id => id.trim()).filter(id => id !== '')
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('ERRO obterProjetoPorIdSimples: ' + e.toString());
    return null;
  }
}

function obterEtapaPorId(etapaId) {
  try {
    const dados = obterDadosAbaComCache(NOME_ABA_ETAPAS);
    if (!dados || dados.length <= 1) return null;
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_ETAPAS.ID] === etapaId) {
        const rawIds = dados[i][COLUNAS_ETAPAS.RESPONSAVEIS_IDS] || '';
        return {
          id:              dados[i][COLUNAS_ETAPAS.ID],
          projetoId:       dados[i][COLUNAS_ETAPAS.PROJETO_ID],
          responsaveisIds: rawIds.toString().split(',').map(id => id.trim()).filter(id => id !== '')
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('ERRO obterEtapaPorId: ' + e.toString());
    return null;
  }
}

function obterResponsavelPorEmail(email) {
  try {
    if (!email) return null;
    const emailLower = email.toLowerCase();
    const dados = obterDadosAbaComCache(NOME_ABA_RESPONSAVEIS);
    if (!dados || dados.length <= 1) return null;
    for (let i = 1; i < dados.length; i++) {
      const emailResp = (dados[i][COLUNAS_RESPONSAVEIS.EMAIL] || '').toString().toLowerCase();
      if (emailResp === emailLower) {
        return {
          id: dados[i][COLUNAS_RESPONSAVEIS.ID],
          nome: dados[i][COLUNAS_RESPONSAVEIS.NOME],
          email: dados[i][COLUNAS_RESPONSAVEIS.EMAIL]
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('ERRO obterResponsavelPorEmail: ' + e.toString());
    return null;
  }
}

// listarTodasPermissoes, salvarPermissaoUsuario, removerPermissaoUsuario
// removidos — gerenciamento de usuários feito via PaginaAdmin + Auth.js
