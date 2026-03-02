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
const NOME_ABA_PERMISSOES = 'Permissoes';
const NOME_ABA_LIXEIRA = 'LixeiraProjetos';

/** Reuniões */

const NOME_ABA_REUNIOES = 'Reuniões';

/** =====================================================================
 *                          COLUNAS DAS ABAS
=========================================================================*/

const COLUNAS_PERMISSOES = {
  ID: 0,
  EMAIL_USUARIO: 1,
  NIVEL_ACESSO: 2,             // 'admin', 'gestor', 'colaborador'
  SETORES_PERMITIDOS: 3,       // IDs separados por v?rgula (vazio = todos para admin)
  PROJETOS_PERMITIDOS: 4,      // IDs separados por v?rgula (vazio = herda do setor)
  PODE_CRIAR_PROJETO: 5,       // true/false
  PODE_CRIAR_ETAPA: 6,         // true/false
  ATIVO: 7,                    // true/false
  FILTRAR_POR_RESPONSAVEL: 8   // true/false — exibe apenas projetos onde o usuário é responsável
};

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
    const pagina = (e?.parameter?.pagina || 'reunioes').toString().trim().toLowerCase();

    // Modo admin: ativado via ?modo=admin, mas só válido para usuários admin na planilha
    const solicitouAdmin = (e?.parameter?.modo === 'admin');
    const modoAdmin = solicitouAdmin && ehAdministrador();

    let nomeArquivoHtml, titulo;

    // ==================== PÁGINA DE DETALHE DO PROJETO ====================
    if (pagina === 'projeto') {
      nomeArquivoHtml = 'PaginaProjetoDetalhe🟡';
      titulo = 'Smart Meeting - Detalhes do Projeto';

      const tmpl = HtmlService.createTemplateFromFile(nomeArquivoHtml);
      tmpl.modoAdmin = modoAdmin;
      return tmpl.evaluate()
        .setTitle(titulo)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    // ==================== PÁGINA DE RELATÓRIOS DE IDENTIFICAÇÃO ====================
    else if (pagina === 'relatorios') {
      nomeArquivoHtml = 'PaginaRelatorios';
      titulo = 'Smart Meeting - Relatórios de Identificação';
    }
    // ==================== PÁGINA DE REUNIÕES (PADRÃO) ====================
    else {
      nomeArquivoHtml = 'PaginaReunioes▶️';
      titulo = 'Smart Meeting - Reuniões';
    }

    const template = HtmlService.createTemplateFromFile(nomeArquivoHtml);
    template.modoAdmin = modoAdmin;
    return template.evaluate()
      .setTitle(titulo)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (erro) {
    Logger.log('ERRO no doGet: ' + erro.toString());
    return HtmlService.createHtmlOutput('Erro ao carregar página: ' + erro.message);
  }
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

function verificarPermissaoEdicao() {
  const emailUsuario = obterEmailUsuario();
  if (!emailUsuario) return false;
  const dados = obterDadosAbaComCache(NOME_ABA_RESPONSAVEIS);
  if (dados.length <= 1) return true;
  const emailLower = emailUsuario.toLowerCase();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_RESPONSAVEIS.EMAIL] && dados[i][COLUNAS_RESPONSAVEIS.EMAIL].toString().toLowerCase() === emailLower) return true;
  }
  return true;
}

function exigirPermissaoEdicao() {
    // Sem restri?o ? qualquer usu?rio autenticado pode editar
  return true;
}

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

/**  ==========================================
//        //        SISTEMA DE PERMISS?ES GRANULAR
//   ========================================== */

function obterPermissoesUsuarioAtual() {
  try {
    const emailUsuario = obterEmailUsuario();
    if (!emailUsuario) {
      return criarPermissaoPadrao('visitante');
    }
    
    const aba = obterAba(NOME_ABA_PERMISSOES);
    if (!aba || aba.getLastRow() <= 1) {
      //       // Se n?o h? permiss?es configuradas, verifica se ? o primeiro acesso
      //       // Primeiro usu?rio vira admin automaticamente
      const totalLinhas = aba ? aba.getLastRow() : 0;
      if (totalLinhas <= 1) {
        const novaPermissao = criarPermissaoAdmin(emailUsuario);
        salvarPermissao(novaPermissao);
        return novaPermissao;
      }
      return criarPermissaoPadrao('colaborador');
    }
    
    const dados = obterDadosAbaComCache(NOME_ABA_PERMISSOES);
    const emailLower = emailUsuario.toLowerCase();
    
    for (let i = 1; i < dados.length; i++) {
      const emailRegistro = (dados[i][COLUNAS_PERMISSOES.EMAIL_USUARIO] || '').toString().toLowerCase();
      if (emailRegistro === emailLower) {
        const ativo = dados[i][COLUNAS_PERMISSOES.ATIVO];
        if (ativo === false || ativo === 'false') {
          return criarPermissaoPadrao('inativo');
        }
        
        return {
          id: dados[i][COLUNAS_PERMISSOES.ID],
          email: emailUsuario,
          nivelAcesso: dados[i][COLUNAS_PERMISSOES.NIVEL_ACESSO] || NIVEIS_ACESSO.COLABORADOR,
          setoresPermitidos: parseArrayString(dados[i][COLUNAS_PERMISSOES.SETORES_PERMITIDOS]),
          projetosPermitidos: parseArrayString(dados[i][COLUNAS_PERMISSOES.PROJETOS_PERMITIDOS]),
          podeCriarProjeto: dados[i][COLUNAS_PERMISSOES.PODE_CRIAR_PROJETO] === true || dados[i][COLUNAS_PERMISSOES.PODE_CRIAR_PROJETO] === 'true',
          podeCriarEtapa: dados[i][COLUNAS_PERMISSOES.PODE_CRIAR_ETAPA] === true || dados[i][COLUNAS_PERMISSOES.PODE_CRIAR_ETAPA] === 'true',
          filtrarPorResponsavel: dados[i][COLUNAS_PERMISSOES.FILTRAR_POR_RESPONSAVEL] === true || dados[i][COLUNAS_PERMISSOES.FILTRAR_POR_RESPONSAVEL] === 'true',
          ativo: true
        };
      }
    }
    
    //     // Usu?rio n?o encontrado - verifica se est? na aba de respons?veis
    const dadosResp = obterDadosAbaComCache(NOME_ABA_RESPONSAVEIS);
    
    for (let i = 1; i < dadosResp.length; i++) {
      const emailResp = (dadosResp[i][COLUNAS_RESPONSAVEIS.EMAIL] || '').toString().toLowerCase();
      if (emailResp === emailLower) {
        //         // Est? cadastrado como respons?vel mas sem permiss?o expl?cita
        //         // Cria permiss?o de colaborador automaticamente
        const novaPermissao = criarPermissaoPadrao('colaborador');
        novaPermissao.email = emailUsuario;
        salvarPermissao(novaPermissao);
        return novaPermissao;
      }
    }
    
    return criarPermissaoPadrao('visitante');
  } catch (e) {
    Logger.log('ERRO obterPermissoesUsuarioAtual: ' + e.toString());
    return criarPermissaoPadrao('erro');
  }
}

function criarPermissaoPadrao(nivel) {
  const base = {
    id: null,
    email: '',
    nivelAcesso: nivel,
    setoresPermitidos: [],
    projetosPermitidos: [],
    podeCriarProjeto: false,
    podeCriarEtapa: false,
    ativo: nivel !== 'inativo' && nivel !== 'visitante'
  };
  
  if (nivel === 'colaborador') {
    base.podeCriarEtapa = true;
  }
  
  return base;
}

function criarPermissaoAdmin(email) {
  return {
    id: gerarId(),
    email: email,
    nivelAcesso: NIVEIS_ACESSO.ADMIN,
    setoresPermitidos: [], // Vazio = todos
    projetosPermitidos: [], // Vazio = todos
    podeCriarProjeto: true,
    podeCriarEtapa: true,
    ativo: true
  };
}

function salvarPermissao(permissao) {
  try {
    const aba = obterAba(NOME_ABA_PERMISSOES);
    const id = permissao.id || gerarId();
    
    const linha = [
      id,
      permissao.email || '',
      permissao.nivelAcesso || NIVEIS_ACESSO.COLABORADOR,
      Array.isArray(permissao.setoresPermitidos) ? permissao.setoresPermitidos.join(',') : '',
      Array.isArray(permissao.projetosPermitidos) ? permissao.projetosPermitidos.join(',') : '',
      permissao.podeCriarProjeto === true,
      permissao.podeCriarEtapa === true,
      permissao.ativo !== false
    ];
    
    aba.appendRow(linha);
    limparCacheAba(NOME_ABA_PERMISSOES);
    return { sucesso: true, id: id };
  } catch (e) {
    Logger.log('ERRO salvarPermissao: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function parseArrayString(valor) {
  if (!valor) return [];
  if (Array.isArray(valor)) return valor;
  return valor.toString().split(',').map(v => v.trim()).filter(v => v !== '');
}

function podeVerProjeto(projetoId, permissoes = null) {
  try {
    if (!permissoes) permissoes = obterPermissoesUsuarioAtual();
    
        // Admin v? tudo
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.ADMIN) return true;
    
    //     // Inativo ou visitante n?o v? nada
    if (!permissoes.ativo || permissoes.nivelAcesso === 'visitante' || permissoes.nivelAcesso === 'inativo') {
      return false;
    }
    
    //     // Verifica se projeto est? na lista de permitidos
    if (permissoes.projetosPermitidos.length > 0) {
      if (permissoes.projetosPermitidos.includes(projetoId)) return true;
    }
    
    //     // Verifica se o setor do projeto est? nos setores permitidos
    const projeto = obterProjetoPorIdSimples(projetoId);
    if (projeto && projeto.setor && permissoes.setoresPermitidos.length > 0) {
      if (permissoes.setoresPermitidos.includes(projeto.setor)) return true;
    }
    
    // Gestor pode ver projetos do seu setor
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.GESTOR) {
      if (projeto && projeto.setor && permissoes.setoresPermitidos.includes(projeto.setor)) {
        return true;
      }
    }
    
    //     // Colaborador pode ver projetos onde est? atribu?do
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.COLABORADOR) {
      const emailUsuario = obterEmailUsuario().toLowerCase();
      const responsavel = obterResponsavelPorEmail(emailUsuario);
      
      if (responsavel) {
        //         // Verifica se est? nos respons?veis do projeto
        if (projeto && projeto.responsaveisIds && projeto.responsaveisIds.includes(responsavel.id)) {
          return true;
        }
        
        //         // Verifica se est? em alguma etapa do projeto
        const etapas = obterTodasEtapas().filter(e => e.projetoId === projetoId);
        for (const etapa of etapas) {
          if (etapa.responsaveisIds && etapa.responsaveisIds.includes(responsavel.id)) {
            return true;
          }
        }
      }
    }
    
    //     // Se n?o tem restri?es de setor/projeto configuradas, permite ver
    if (permissoes.setoresPermitidos.length === 0 && permissoes.projetosPermitidos.length === 0) {
      return true;
    }
    
    return false;
  } catch (e) {
    Logger.log('ERRO podeVerProjeto: ' + e.toString());
    return false;
  }
}

function podeEditarProjeto(projetoId, permissoes = null) {
  try {
    if (!permissoes) permissoes = obterPermissoesUsuarioAtual();
    
    // Admin edita tudo
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.ADMIN) return true;
    
    //     // Inativo ou visitante n?o edita nada
    if (!permissoes.ativo || permissoes.nivelAcesso === 'visitante' || permissoes.nivelAcesso === 'inativo') {
      return false;
    }
    
    // Primeiro precisa poder ver
    if (!podeVerProjeto(projetoId, permissoes)) return false;
    
    const projeto = obterProjetoPorIdSimples(projetoId);
    if (!projeto) return false;
    
    // Gestor pode editar projetos do seu setor
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.GESTOR) {
      if (projeto.setor && permissoes.setoresPermitidos.includes(projeto.setor)) {
        return true;
      }
    }
    
    //     // Colaborador N?O pode editar projeto, apenas etapas
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.COLABORADOR) {
      return false;
    }
    
    return false;
  } catch (e) {
    Logger.log('ERRO podeEditarProjeto: ' + e.toString());
    return false;
  }
}

function podeEditarEtapa(etapaId, permissoes = null) {
  try {
    if (!permissoes) permissoes = obterPermissoesUsuarioAtual();
    
    // Admin edita tudo
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.ADMIN) return true;
    
    //     // Inativo ou visitante n?o edita nada
    if (!permissoes.ativo) return false;
    
    const etapa = dadosDiagrama?.etapas?.find(e => e.id === etapaId) || obterEtapaPorId(etapaId);
    if (!etapa) return false;
    
    // Verifica se pode ver o projeto da etapa
    if (!podeVerProjeto(etapa.projetoId, permissoes)) return false;
    
    // Gestor pode editar todas as etapas do seu setor
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.GESTOR) {
      const projeto = obterProjetoPorIdSimples(etapa.projetoId);
      if (projeto && projeto.setor && permissoes.setoresPermitidos.includes(projeto.setor)) {
        return true;
      }
    }
    
    //     // Colaborador pode editar etapas onde est? atribu?do
    if (permissoes.nivelAcesso === NIVEIS_ACESSO.COLABORADOR) {
      const emailUsuario = obterEmailUsuario().toLowerCase();
      const responsavel = obterResponsavelPorEmail(emailUsuario);
      
      if (responsavel && etapa.responsaveisIds && etapa.responsaveisIds.includes(responsavel.id)) {
        return true;
      }
    }
    
    return false;
  } catch (e) {
    Logger.log('ERRO podeEditarEtapa: ' + e.toString());
    return false;
  }
}

function podeCriarProjeto(permissoes = null) {
  if (!permissoes) permissoes = obterPermissoesUsuarioAtual();
  return permissoes.nivelAcesso === NIVEIS_ACESSO.ADMIN || permissoes.podeCriarProjeto === true;
}

function podeCriarEtapa(projetoId, permissoes = null) {
  if (!permissoes) permissoes = obterPermissoesUsuarioAtual();
  
  if (permissoes.nivelAcesso === NIVEIS_ACESSO.ADMIN) return true;
  if (!permissoes.podeCriarEtapa) return false;
  
  // Precisa poder ver o projeto para criar etapa nele
  return podeVerProjeto(projetoId, permissoes);
}

function ehAdministrador(permissoes = null) {
  if (!permissoes) permissoes = obterPermissoesUsuarioAtual();
  return permissoes.nivelAcesso === NIVEIS_ACESSO.ADMIN;
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

function listarTodasPermissoes() {
  try {
    const permissoesUsuario = obterPermissoesUsuarioAtual();
    if (!ehAdministrador(permissoesUsuario)) {
      return { sucesso: false, mensagem: 'Acesso negado. Apenas administradores.' };
    }
    
    const dados = obterDadosAbaComCache(NOME_ABA_PERMISSOES);
    if (!dados || dados.length <= 1) {
      return { sucesso: true, permissoes: [] };
    }

    const permissoes = [];
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PERMISSOES.ID]) {
        permissoes.push({
          id: dados[i][COLUNAS_PERMISSOES.ID],
          email: dados[i][COLUNAS_PERMISSOES.EMAIL_USUARIO],
          nivelAcesso: dados[i][COLUNAS_PERMISSOES.NIVEL_ACESSO],
          setoresPermitidos: parseArrayString(dados[i][COLUNAS_PERMISSOES.SETORES_PERMITIDOS]),
          projetosPermitidos: parseArrayString(dados[i][COLUNAS_PERMISSOES.PROJETOS_PERMITIDOS]),
          podeCriarProjeto: dados[i][COLUNAS_PERMISSOES.PODE_CRIAR_PROJETO] === true || dados[i][COLUNAS_PERMISSOES.PODE_CRIAR_PROJETO] === 'true',
          podeCriarEtapa: dados[i][COLUNAS_PERMISSOES.PODE_CRIAR_ETAPA] === true || dados[i][COLUNAS_PERMISSOES.PODE_CRIAR_ETAPA] === 'true',
          ativo: dados[i][COLUNAS_PERMISSOES.ATIVO] !== false && dados[i][COLUNAS_PERMISSOES.ATIVO] !== 'false',
          filtrarPorResponsavel: dados[i][COLUNAS_PERMISSOES.FILTRAR_POR_RESPONSAVEL] === true || dados[i][COLUNAS_PERMISSOES.FILTRAR_POR_RESPONSAVEL] === 'true'
        });
      }
    }
    
    return { sucesso: true, permissoes: permissoes };
  } catch (e) {
    Logger.log('ERRO listarTodasPermissoes: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function salvarPermissaoUsuario(dadosPermissao) {
  try {
    const permissoesUsuario = obterPermissoesUsuarioAtual();
    if (!ehAdministrador(permissoesUsuario)) {
      return { sucesso: false, mensagem: 'Acesso negado. Apenas administradores.' };
    }
    
    if (!dadosPermissao.email) {
      return { sucesso: false, mensagem: 'Email ? obrigat?rio' };
    }
    
    const aba = obterAba(NOME_ABA_PERMISSOES);
    const dados = aba.getDataRange().getValues();
    const emailLower = dadosPermissao.email.toLowerCase();
    
        // Verifica se j? existe
    for (let i = 1; i < dados.length; i++) {
      const emailRegistro = (dados[i][COLUNAS_PERMISSOES.EMAIL_USUARIO] || '').toString().toLowerCase();
      if (emailRegistro === emailLower) {
        const linha = i + 1;
        const linhaAtualizada = dados[i].slice();
        linhaAtualizada[COLUNAS_PERMISSOES.NIVEL_ACESSO] = dadosPermissao.nivelAcesso || NIVEIS_ACESSO.COLABORADOR;
        linhaAtualizada[COLUNAS_PERMISSOES.SETORES_PERMITIDOS] =
          Array.isArray(dadosPermissao.setoresPermitidos) ? dadosPermissao.setoresPermitidos.join(',') : '';
        linhaAtualizada[COLUNAS_PERMISSOES.PROJETOS_PERMITIDOS] =
          Array.isArray(dadosPermissao.projetosPermitidos) ? dadosPermissao.projetosPermitidos.join(',') : '';
        linhaAtualizada[COLUNAS_PERMISSOES.PODE_CRIAR_PROJETO] = dadosPermissao.podeCriarProjeto === true;
        linhaAtualizada[COLUNAS_PERMISSOES.PODE_CRIAR_ETAPA] = dadosPermissao.podeCriarEtapa === true;
        linhaAtualizada[COLUNAS_PERMISSOES.ATIVO] = dadosPermissao.ativo !== false;
        linhaAtualizada[COLUNAS_PERMISSOES.FILTRAR_POR_RESPONSAVEL] = dadosPermissao.filtrarPorResponsavel === true;

        aba.getRange(linha, 1, 1, linhaAtualizada.length).setValues([linhaAtualizada]);
        limparCacheAba(NOME_ABA_PERMISSOES);

        return { sucesso: true, mensagem: 'Permiss?o atualizada!', id: dados[i][COLUNAS_PERMISSOES.ID] };
      }
    }
    
    // Cria nova
    const id = gerarId();
    const novaLinha = [
      id,
      dadosPermissao.email,
      dadosPermissao.nivelAcesso || NIVEIS_ACESSO.COLABORADOR,
      Array.isArray(dadosPermissao.setoresPermitidos) ? dadosPermissao.setoresPermitidos.join(',') : '',
      Array.isArray(dadosPermissao.projetosPermitidos) ? dadosPermissao.projetosPermitidos.join(',') : '',
      dadosPermissao.podeCriarProjeto === true,
      dadosPermissao.podeCriarEtapa === true,
      dadosPermissao.ativo !== false,
      dadosPermissao.filtrarPorResponsavel === true
    ];

    aba.appendRow(novaLinha);
    limparCacheAba(NOME_ABA_PERMISSOES);
    return { sucesso: true, mensagem: 'Permiss?o criada!', id: id };
  } catch (e) {
    Logger.log('ERRO salvarPermissaoUsuario: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function removerPermissaoUsuario(permissaoId) {
  try {
    const permissoesUsuario = obterPermissoesUsuarioAtual();
    if (!ehAdministrador(permissoesUsuario)) {
      return { sucesso: false, mensagem: 'Acesso negado. Apenas administradores.' };
    }
    
    const aba = obterAba(NOME_ABA_PERMISSOES);
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PERMISSOES.ID] === permissaoId) {
                // Não permite remover o próprio admin
        const emailPermissao = dados[i][COLUNAS_PERMISSOES.EMAIL_USUARIO];
        const emailAtual = obterEmailUsuario();
        if (emailPermissao.toLowerCase() === emailAtual.toLowerCase()) {
          return { sucesso: false, mensagem: 'Voc? n?o pode remover sua pr?pria permiss?o!' };
        }
        
        aba.deleteRow(i + 1);
        limparCacheAba(NOME_ABA_PERMISSOES);
        return { sucesso: true, mensagem: 'Permiss?o removida!' };
      }
    }
    
    return { sucesso: false, mensagem: 'Permiss?o n?o encontrada' };
  } catch (e) {
    Logger.log('ERRO removerPermissaoUsuario: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}
