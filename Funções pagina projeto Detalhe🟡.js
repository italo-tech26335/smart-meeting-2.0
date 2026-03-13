/** =====================================================================
 *                    CONFIGURAÇÃO DE URGÊNCIA (PARA QUEM)
=========================================================================*/

const PESSOAS_PARA_QUEM = {
  'Diretoria': { peso: 5, cor: '#dc2626', bgCor: 'rgba(220, 38, 38, 0.15)', label: '🔴 Crítico' },
  'Demais áreas': { peso: 4, cor: '#ea580c', bgCor: 'rgba(234, 88, 12, 0.15)', label: '🟠 Urgente' },
};

const CONFIG_RESUMO_SEMANAL = {
  EMAIL_GESTOR: 'murilobcr@gmail.com',       // ← trocar pelo email real
  NOME_GESTOR: 'Murilo',                     // ← nome para saudação
  DIA_SEMANA: ScriptApp.WeekDay.FRIDAY,      // sexta-feira
  HORA_DISPARO: 18,                          // 18:00
  ASSUNTO: '📊 Resumo Semanal — Projetos em Andamento',
  // Status que entram no resumo (apenas "em andamento" e variações ativas)
  STATUS_INCLUIDOS: ['Em Andamento', 'A Fazer', 'Aguardando Setor']
};

// Função para obter configuração de pessoas (será chamada pelo frontend)
function obterPessoasParaQuem() {
  return PESSOAS_PARA_QUEM;
}

/**
 * Retorna as reuniões vinculadas a um projeto específico,
 * incluindo o conteúdo completo de ata e transcrição.
 * @param {string} projetoId
 * @returns {{ sucesso: boolean, reunioes: Array, mensagem?: string }}
 */
function obterReunioesDoProjeto(projetoId) {
  try {
    if (!projetoId) return { sucesso: false, reunioes: [], mensagem: 'projetoId não informado' };

    const nomeAba = typeof NOME_ABA_REUNIOES !== 'undefined' ? NOME_ABA_REUNIOES : 'Reuniões';
    const colunas = typeof COLUNAS_REUNIOES !== 'undefined' ? COLUNAS_REUNIOES : {
      ID: 0, TITULO: 1, DATA_INICIO: 2, DATA_FIM: 3, DURACAO: 4, STATUS: 5,
      PARTICIPANTES: 6, TRANSCRICAO: 7, ATA: 8, SUGESTOES_IA: 9, LINK_AUDIO: 10,
      LINK_ATA: 11, EMAILS_ENVIADOS: 12, PROJETOS_IMPACTADOS: 13
    };

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName(nomeAba);
    if (!aba || aba.getLastRow() <= 1) return { sucesso: true, reunioes: [] };

    const dados = aba.getDataRange().getValues();
    const reunioes = [];
    const idBuscado = String(projetoId).trim();

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const idCelula = linha[colunas.ID] ? String(linha[colunas.ID]).trim() : '';
      if (!idCelula) continue;

      const projVinculado = linha[colunas.PROJETOS_IMPACTADOS]
        ? String(linha[colunas.PROJETOS_IMPACTADOS]).trim()
        : '';
      if (projVinculado !== idBuscado) continue;

      const ataTexto = linha[colunas.ATA] ? String(linha[colunas.ATA]).trim() : '';
      const transcricaoTexto = linha[colunas.TRANSCRICAO] ? String(linha[colunas.TRANSCRICAO]).trim() : '';

      reunioes.push({
        id: idCelula,
        titulo: linha[colunas.TITULO] ? String(linha[colunas.TITULO]) : 'Reunião sem título',
        data: linha[colunas.DATA_INICIO] ? String(linha[colunas.DATA_INICIO]) : '',
        duracao: linha[colunas.DURACAO] ? String(linha[colunas.DURACAO]) : '',
        participantes: linha[colunas.PARTICIPANTES] ? String(linha[colunas.PARTICIPANTES]) : '',
        linkAudio: linha[colunas.LINK_AUDIO] ? String(linha[colunas.LINK_AUDIO]) : '',
        ata: ataTexto,
        transcricao: transcricaoTexto,
        temAta: ataTexto.length > 10,
        temTranscricao: transcricaoTexto.length > 10
      });
    }

    // Ordenar da mais recente para a mais antiga
    reunioes.sort(function(a, b) { return b.data > a.data ? 1 : -1; });

    return { sucesso: true, reunioes: reunioes };
  } catch (e) {
    Logger.log('ERRO obterReunioesDoProjeto: ' + e.toString());
    return { sucesso: false, reunioes: [], mensagem: e.message };
  }
}

/**
 * Carrega todas as etapas e prioridades de um projeto específico
 * @param {string} projetoId - ID do projeto
 * @returns {Object} { sucesso, etapas, prioridades, mensagem }
 */
function carregarEtapasDoProjeto(projetoId) {
  
  // ==================== CONFIGURAÇÃO ====================
  const ID_PLANILHA = SpreadsheetApp.getActiveSpreadsheet().getId(); // ou cole o ID fixo
  const planilha = SpreadsheetApp.openById(ID_PLANILHA);
  
  // Abas
  const ABA_ETAPAS = 'Etapas';
  const ABA_PRIORIDADES = 'Prioridades'; // Opcional - pode não existir
  
  // Colunas da aba Etapas (baseado no índice, começando em 0)
  const COLUNAS_ETAPAS = {
    id: 0,
    projetoId: 1,
    etapaPaiId: 2,
    nome: 3,
    descricao: 4,
    oQueFazer: 5,
    pendencias: 6,
    status: 7,
    responsaveisIds: 8,  // Separados por vírgula ou JSON
    dataCriacao: 9,
    dataAtualizacao: 10
  };
  
  // Colunas da aba Prioridades (se existir)
  const COLUNAS_PRIORIDADES = {
    id: 0,
    tipoItem: 1,        // 'etapa' ou 'projeto'
    itemId: 2,
    projetoId: 3,
    ordemPrioridade: 4,
    responsavelId: 5
  };
  
  // ==================== VALIDAÇÃO ====================
  if (!projetoId) {
    return {
      sucesso: false,
      mensagem: 'ID do projeto não informado',
      etapas: [],
      prioridades: []
    };
  }
  
  try {
    // ==================== BUSCAR ETAPAS ====================
    const abaEtapas = planilha.getSheetByName(ABA_ETAPAS);
    
    if (!abaEtapas) {
      return {
        sucesso: false,
        mensagem: `Aba "${ABA_ETAPAS}" não encontrada`,
        etapas: [],
        prioridades: []
      };
    }
    
    const dadosEtapas = abaEtapas.getDataRange().getValues();
    const etapas = [];
    
    // Pular cabeçalho (linha 0)
    for (let i = 1; i < dadosEtapas.length; i++) {
      const linha = dadosEtapas[i];
      const etapaProjetoId = String(linha[COLUNAS_ETAPAS.projetoId] || '').trim();
      
      // Filtrar apenas etapas do projeto solicitado
      if (etapaProjetoId === projetoId) {
        
        // Processar responsáveis (pode ser JSON ou separado por vírgula)
        let responsaveisIds = [];
        const valorResponsaveis = linha[COLUNAS_ETAPAS.responsaveisIds];
        
        if (valorResponsaveis) {
          const valorStr = String(valorResponsaveis).trim();
          
          if (valorStr.startsWith('[')) {
            // É um JSON array
            try {
              responsaveisIds = JSON.parse(valorStr);
            } catch (e) {
              responsaveisIds = [];
            }
          } else if (valorStr.includes(',')) {
            // Separado por vírgula
            responsaveisIds = valorStr.split(',').map(id => id.trim()).filter(id => id);
          } else if (valorStr) {
            // Único valor
            responsaveisIds = [valorStr];
          }
        }
        
        etapas.push({
          id: String(linha[COLUNAS_ETAPAS.id] || ''),
          projetoId: etapaProjetoId,
          etapaPaiId: String(linha[COLUNAS_ETAPAS.etapaPaiId] || ''),
          nome: String(linha[COLUNAS_ETAPAS.nome] || ''),
          descricao: String(linha[COLUNAS_ETAPAS.descricao] || ''),
          oQueFazer: String(linha[COLUNAS_ETAPAS.oQueFazer] || ''),
          pendencias: String(linha[COLUNAS_ETAPAS.pendencias] || ''),
          status: String(linha[COLUNAS_ETAPAS.status] || 'A Fazer'),
          responsaveisIds: responsaveisIds
        });
      }
    }
    
    // ==================== BUSCAR PRIORIDADES (OPCIONAL) ====================
    let prioridades = [];
    const abaPrioridades = planilha.getSheetByName(ABA_PRIORIDADES);
    
    if (abaPrioridades) {
      const dadosPrioridades = abaPrioridades.getDataRange().getValues();
      
      // Pular cabeçalho (linha 0)
      for (let i = 1; i < dadosPrioridades.length; i++) {
        const linha = dadosPrioridades[i];
        const prioridadeProjetoId = String(linha[COLUNAS_PRIORIDADES.projetoId] || '').trim();
        const tipoItem = String(linha[COLUNAS_PRIORIDADES.tipoItem] || '').trim();
        
        // Filtrar prioridades de etapas do projeto solicitado
        if (prioridadeProjetoId === projetoId && tipoItem === 'etapa') {
          prioridades.push({
            id: String(linha[COLUNAS_PRIORIDADES.id] || ''),
            tipoItem: tipoItem,
            itemId: String(linha[COLUNAS_PRIORIDADES.itemId] || ''),
            projetoId: prioridadeProjetoId,
            ordemPrioridade: Number(linha[COLUNAS_PRIORIDADES.ordemPrioridade]) || 0,
            responsavelId: String(linha[COLUNAS_PRIORIDADES.responsavelId] || '')
          });
        }
      }
      
      // Ordenar por ordem de prioridade
      prioridades.sort((a, b) => a.ordemPrioridade - b.ordemPrioridade);
    }
    
    // ==================== RETORNO ====================
    return {
      sucesso: true,
      mensagem: `${etapas.length} etapa(s) carregada(s)`,
      etapas: etapas,
      prioridades: prioridades
    };
    
  } catch (erro) {
    console.error('Erro ao carregar etapas do projeto:', erro);
    return {
      sucesso: false,
      mensagem: 'Erro ao carregar etapas: ' + erro.message,
      etapas: [],
      prioridades: []
    };
  }
}

const CONFIG_PROJETO_DETALHE = {
  MODELO_IA: 'gemini-2.5-flash',
  MAX_TOKENS_RESPOSTA: 8192,
  TEMPERATURA: 0.7,
  URL_API: 'https://generativelanguage.googleapis.com/v1beta/models'
};

function carregarDadosEditorProjeto(token) {
  Logger.log('Iniciando carregarDadosEditorProjeto (sem projeto específico)');
  try {
    const sessao = token ? _obterSessao(token) : null;
    if (!sessao) return { sucesso: false, mensagem: 'Sessão inválida ou expirada.' };

    const setores = listarSetores();
    const responsaveis = listarResponsaveisCompletos();
    
    const todosProjetosRaw = listarProjetosCompletos(token);
    const listaProjetosCompleta = todosProjetosRaw.map(p => ({
      id: p.id,
      nome: p.nome,
      descricao: p.descricao || '',
      tipo: p.tipo || '',
      paraQuem: p.paraQuem || '',
      status: p.status || 'A Fazer',
      prioridade: p.prioridade || 'Média',
      link: p.link || '',
      gravidade: p.gravidade || '',
      urgencia: p.urgencia || '',
      esforco: p.esforco || '',
      setor: p.setor || '',
      pilar: p.pilar || '',
      responsaveisIds: p.responsaveisIds || [],
      valorPrioridade: p.valorPrioridade || 0,
      dataInicio: p.dataInicio || '',
      dataFim: p.dataFim || '',
      departamentoId: p.departamentoId || ''
    }));

        // ── Calcular duração de cada projeto ──
    listaProjetosCompleta.forEach(function(p) {
      if (p.dataInicio && p.dataFim) {
        var resDuracao = calcularDuracaoProjetoHoras(p.dataInicio, p.dataFim);
        p.duracaoFormatada = resDuracao.sucesso ? resDuracao.textoFormatado : '';
        p.horasTotais      = resDuracao.sucesso ? resDuracao.horasTotais    : 0;
      } else {
        p.duracaoFormatada = '';
        p.horasTotais      = 0;
      }
    });

    // ── Obter tema e permissões do usuário ──
    var temaUsuario = obterPreferenciaTema();
    const permissoes = {
      perfil: sessao.perfil || 'visualizador',
      podeEditar: sessao.perfil === 'admin' || sessao.perfil === 'usuario',
      modoAdmin: sessao.perfil === 'admin',
      departamentosIds: sessao.departamentosIds || []
    };
    
    let prioridadesProjetos = [];
    try {
      const resPrioProjetos = obterPrioridadesGeraisDeProjetos();
      if (resPrioProjetos.sucesso) {
        prioridadesProjetos = resPrioProjetos.prioridades;
      }
    } catch (e) {
      Logger.log('Erro ao carregar prioridades de projetos: ' + e.toString());
    }
    
    const pessoasParaQuem = PESSOAS_PARA_QUEM || {};
    const configCalculoPrioridade = obterConfigCalculoPrioridade();
    const resDepartamentos = listarDepartamentos(null);
    const departamentos = resDepartamentos.sucesso ? (resDepartamentos.departamentos || []) : [];

    return {
      sucesso: true,
      setores: setores,
      responsaveis: responsaveis,
      listaProjetos: listaProjetosCompleta,
      projeto: null,
      etapas: [],
      prioridadesProjetos: prioridadesProjetos,
      pessoasParaQuem: pessoasParaQuem,
      configCalculoPrioridade: configCalculoPrioridade,
      tema: temaUsuario,
      permissoes: permissoes,
      departamentos: departamentos
    };
    
  } catch (e) {
    Logger.log('ERRO carregarDadosEditorProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function obterContextoParaIA() {
  try {
    const setores = listarSetores().map(s => s.nome).join(', ');
    const projetos = listarProjetosCompletos().map(p => p.nome).join(', ');
    const responsaveis = listarResponsaveisCompletos().map(r => `${r.nome} (${r.cargo})`).join(', ');
    
    return `
    CONTEXTO ATUAL DO SISTEMA:
    - Setores Disponíveis: ${setores}
    - Projetos Existentes: ${projetos}
    - Equipe/Responsáveis: ${responsaveis}
    `;
  } catch (e) {
    Logger.log('Erro ao obter contexto IA: ' + e.toString());
    return '';
  }
}

function enviarMensagemParaGemini(mensagemUsuario, historico, projetoAtualJson) {
  Logger.log('Enviando mensagem para Gemini...');
  try {
    const chave = obterChaveGeminiProjeto();
    if (!chave) throw new Error('Chave API não configurada.');

    const contexto = obterContextoParaIA();
    const dadosProjeto = projetoAtualJson ? `DADOS DO PROJETO ATUAL NO EDITOR: ${JSON.stringify(projetoAtualJson)}` : 'Nenhum projeto aberto no momento.';

    const systemPrompt = `
      Você é um Gerente de Projetos Sênior e Assistente de IA especialista.
      Seu objetivo é ajudar o usuário a estruturar projetos, definir etapas, identificar riscos e sugerir melhorias.
      
      ${contexto}
      ${dadosProjeto}
      
      DIRETRIZES:
      1. SEMPRE use formatação Markdown nas respostas:
         - Use ## para títulos de seções
         - Use **negrito** para destaques importantes
         - Use listas com - para itens
         - Use \`código\` para termos técnicos
      2. Seja direto e prático nas respostas.
      3. Se o usuário pedir para criar etapas, sugira uma lista estruturada com:
         ## Etapas Sugeridas
         - **Etapa 1**: Descrição
         - **Etapa 2**: Descrição
      4. Analise o contexto dos setores e responsáveis para fazer sugestões realistas.
      5. Mantenha um tom profissional mas colaborativo.
      6. Organize respostas longas em seções com subtítulos.
    `;

    const contents = [];
    
    if (historico && historico.length > 0) {
      historico.forEach(msg => {
        contents.push({ role: msg.role, parts: [{ text: msg.texto }] });
      });
    }

    contents.push({ role: 'user', parts: [{ text: mensagemUsuario }] });

    const payload = {
      contents: contents,
      systemInstruction: { parts: [{ text: systemPrompt }] },
      generationConfig: {
        temperature: CONFIG_PROJETO_DETALHE.TEMPERATURA,
        maxOutputTokens: CONFIG_PROJETO_DETALHE.MAX_TOKENS_RESPOSTA
      }
    };

    const url = `${CONFIG_PROJETO_DETALHE.URL_API}/${CONFIG_PROJETO_DETALHE.MODELO_IA}:generateContent?key=${chave}`;
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.error) {
      throw new Error(json.error.message);
    }

    const respostaTexto = json.candidates[0].content.parts[0].text;
    return { sucesso: true, resposta: respostaTexto };

  } catch (e) {
    Logger.log('ERRO Gemini Texto: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro na IA: ' + e.message };
  }
}

function processarAudioParaGemini(audioBase64, mimeType, projetoAtualJson) {
  Logger.log('Processando áudio para Gemini. Mime: ' + mimeType);
  try {
    const chave = obterChaveGeminiProjeto();
    if (!chave) throw new Error('Chave API não configurada.');

    const contexto = obterContextoParaIA();
    const dadosProjeto = projetoAtualJson ? `DADOS PROJETO: ${JSON.stringify(projetoAtualJson)}` : '';
    
    const cleanBase64 = audioBase64.split(',')[1] || audioBase64;

    const payload = {
      contents: [{
        role: 'user',
        parts: [
          { text: `Analise este áudio. Ele contém instruções sobre um projeto. Responda sugerindo ações ou estruturando o projeto baseada no que foi falado. ${contexto} ${dadosProjeto}` },
          { inline_data: { mime_type: mimeType, data: cleanBase64 } }
        ]
      }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 8192
      }
    };

    const url = `${CONFIG_PROJETO_DETALHE.URL_API}/${CONFIG_PROJETO_DETALHE.MODELO_IA}:generateContent?key=${chave}`;
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.error) {
      Logger.log('Erro API Gemini: ' + JSON.stringify(json.error));
      throw new Error(json.error.message);
    }

    const respostaTexto = json.candidates[0].content.parts[0].text;
    return { sucesso: true, resposta: respostaTexto };

  } catch (e) {
    Logger.log('ERRO Gemini Audio: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro ao processar áudio: ' + e.message };
  }
}

function listarSetores() {
  try {
    const aba = obterAba(NOME_ABA_SETORES);
    const dados = aba.getDataRange().getValues();
    const setores = [];
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_SETORES.ID]) {
        setores.push({
          id: dados[i][COLUNAS_SETORES.ID],
          nome: dados[i][COLUNAS_SETORES.NOME],
          descricao: dados[i][COLUNAS_SETORES.DESCRICAO],
          cor: dados[i][COLUNAS_SETORES.COR]
        });
      }
    }
    return setores;
  } catch (e) {
    Logger.log('ERRO listarSetores: ' + e.toString());
    return [];
  }
}

function salvarSetor(dados) {
  try {
    const aba = obterAba(NOME_ABA_SETORES);
    const valores = aba.getDataRange().getValues();
    let linha = -1;
    
    if (dados.id) {
      for (let i = 1; i < valores.length; i++) {
        if (valores[i][COLUNAS_SETORES.ID] == dados.id) {
          linha = i + 1;
          break;
        }
      }
    }
    
    const novoId = dados.id || gerarId();
    const linhaDados = [];
    linhaDados[COLUNAS_SETORES.ID] = novoId;
    linhaDados[COLUNAS_SETORES.NOME] = dados.nome;
    linhaDados[COLUNAS_SETORES.DESCRICAO] = dados.descricao;
    linhaDados[COLUNAS_SETORES.COR] = dados.cor;
    
    if (linha > 0) {
      aba.getRange(linha, 1, 1, linhaDados.length).setValues([linhaDados]);
    } else {
      aba.appendRow(linhaDados);
    }
    
    return { sucesso: true, id: novoId };
  } catch (e) {
    Logger.log('ERRO salvarSetor: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function excluirSetor(id) {
  try {
    const aba = obterAba(NOME_ABA_SETORES);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_SETORES.ID] == id) {
        aba.deleteRow(i + 1);
        return { sucesso: true };
      }
    }
    return { sucesso: false, mensagem: 'Setor não encontrado' };
  } catch (e) {
    Logger.log('ERRO excluirSetor: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function _normalizarDepartamentoValor(valor) {
  return (valor || '').toString().trim().toLowerCase();
}

/**
 * Obtém departamentos atualizados do usuário diretamente da planilha,
 * garantindo que sessões antigas não causem problemas de validação.
 */
function _obterDepsAtualizadosUsuario(sessao) {
  try {
    const usuario = _buscarUsuarioPorEmail(sessao.email);
    const deps = usuario ? (usuario.departamentosIds || []) : (sessao.departamentosIds || []);
    return _resolverIdsDepartamento(deps);
  } catch (e) {
    return _resolverIdsDepartamento(sessao.departamentosIds || []);
  }
}

function _resolverIdsDepartamento(valores) {
  const lista = Array.isArray(valores) ? valores : [valores];
  const mapa = {};
  const dados = obterDadosAbaComCache(NOME_ABA_DEPARTAMENTOS) || [];

  for (let i = 1; i < dados.length; i++) {
    const id = (dados[i][COLUNAS_DEPARTAMENTOS.ID] || '').toString().trim();
    const nome = (dados[i][COLUNAS_DEPARTAMENTOS.NOME] || '').toString().trim();
    if (!id && !nome) continue;
    if (id) mapa[_normalizarDepartamentoValor(id)] = id;
    if (nome) mapa[_normalizarDepartamentoValor(nome)] = id || nome;
  }

  const resolvidos = [];
  lista.forEach(function(v) {
    const bruto = (v || '').toString().trim();
    if (!bruto) return;
    const normalizado = _normalizarDepartamentoValor(bruto);
    resolvidos.push(mapa[normalizado] || bruto);
  });
  return resolvidos;
}

function _obterSessaoEdicaoProjetos(token) {
  const sessao = token ? _obterSessao(token) : null;
  if (!sessao) return { ok: false, mensagem: 'Sessão inválida ou expirada.' };
  if (sessao.perfil === 'visualizador') return { ok: false, mensagem: 'Sem permissão para editar.' };
  return { ok: true, sessao: sessao };
}

function _usuarioPodeAcessarProjetoPorDepartamento(sessao, projetoId) {
  if (!sessao || !projetoId) return false;
  if (sessao.perfil === 'admin') return true;

  const depsUsuario = _obterDepsAtualizadosUsuario(sessao);
  // Se usuário não tem departamentos, permite acesso (compatibilidade retroativa)
  if (depsUsuario.length === 0) return true;

  const aba = obterAba(NOME_ABA_PROJETOS);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_PROJETOS.ID] === projetoId) {
      const depProjetoRaw = (dados[i][COLUNAS_PROJETOS.DEPARTAMENTO_ID] || '').toString().trim();
      // Projetos sem departamento são acessíveis por todos (retrocompat)
      if (!depProjetoRaw) return true;
      const depProjeto = _resolverIdsDepartamento([depProjetoRaw])[0] || depProjetoRaw;
      return depsUsuario.includes(depProjeto);
    }
  }
  return false;
}

function listarProjetosCompletos(token) {
  const sessao = token ? _obterSessao(token) : null;
  const isAdmin = sessao && sessao.perfil === 'admin';
  // Busca departamentos atualizados da planilha (não usa somente a sessão cacheada)
  const depsFiltro = (sessao && !isAdmin) ? _obterDepsAtualizadosUsuario(sessao) : null;

  const aba = obterAba(NOME_ABA_PROJETOS);
  const dados = aba.getDataRange().getValues();
  const projetos = [];

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_PROJETOS.ID]) {
      // Filtro de departamento: não-admin só vê projetos do(s) seu(s) departamento(s)
      // Projetos SEM departamento são visíveis para todos (compatibilidade retroativa)
      if (depsFiltro !== null && depsFiltro.length > 0) {
        const depProjetoRaw = (dados[i][COLUNAS_PROJETOS.DEPARTAMENTO_ID] || '').toString().trim();
        if (depProjetoRaw) {
          const depProjeto = _resolverIdsDepartamento([depProjetoRaw])[0] || depProjetoRaw;
          if (!depsFiltro.includes(depProjeto)) continue;
        }
        // Se projeto sem departamento: mostra para todos
      }

      const respRaw = dados[i][COLUNAS_PROJETOS.RESPONSAVEIS_IDS];
      let respIds = [];
      if (respRaw) {
        try { respIds = JSON.parse(respRaw); } catch(e) { respIds = [respRaw.toString()]; }
      }

      const converterData = (val) => {
        if (!val) return '';
        if (val instanceof Date) {
          return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return val.toString();
      };

      projetos.push({
        id:              dados[i][COLUNAS_PROJETOS.ID],
        nome:            dados[i][COLUNAS_PROJETOS.NOME],
        descricao:       dados[i][COLUNAS_PROJETOS.DESCRICAO],
        tipo:            dados[i][COLUNAS_PROJETOS.TIPO],
        paraQuem:        dados[i][COLUNAS_PROJETOS.PARA_QUEM],
        status:          dados[i][COLUNAS_PROJETOS.STATUS],
        prioridade:      dados[i][COLUNAS_PROJETOS.PRIORIDADE],
        link:            dados[i][COLUNAS_PROJETOS.LINK],
        gravidade:       dados[i][COLUNAS_PROJETOS.GRAVIDADE]  || '',
        urgencia:        dados[i][COLUNAS_PROJETOS.URGENCIA]   || '',
        esforco:         dados[i][COLUNAS_PROJETOS.ESFORCO]    || '',
        setor:           dados[i][COLUNAS_PROJETOS.SETOR],
        pilar:           dados[i][COLUNAS_PROJETOS.PILAR],
        responsaveisIds: respIds,
        valorPrioridade: dados[i][COLUNAS_PROJETOS.VALOR_PRIORIDADE] || 0,
        dataInicio:      converterData(dados[i][COLUNAS_PROJETOS.DATA_INICIO]),
        dataFim:         converterData(dados[i][COLUNAS_PROJETOS.DATA_FIM]),
        departamentoId:  (dados[i][COLUNAS_PROJETOS.DEPARTAMENTO_ID] || '').toString()
      });
    }
  }
  return projetos;
}

function salvarProjetoCompleto(dados, token) {
  Logger.log('Salvando projeto: ' + JSON.stringify(dados));
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    const sessao = auth.sessao;

    const aba = obterAba(NOME_ABA_PROJETOS);
    const valores = aba.getDataRange().getValues();
    let linha = -1;
    
    if (dados.id) {
      for (let i = 1; i < valores.length; i++) {
        if (valores[i][COLUNAS_PROJETOS.ID] == dados.id) {
          linha = i + 1;
          break;
        }
      }
    }
    
    const novoId = dados.id || gerarId();
    const linhaDados = [];
    for(let k in COLUNAS_PROJETOS) linhaDados[COLUNAS_PROJETOS[k]] = '';

    if (linha > 0) {
      const dadosAntigos = valores[linha-1];
      for(let k = 0; k < dadosAntigos.length; k++) linhaDados[k] = dadosAntigos[k];
    } else {
      // Novo projeto: criar pasta no Drive automaticamente
      try {
        const pastaPrincipal = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
        const novaSubPasta = pastaPrincipal.createFolder(dados.nome || 'Projeto ' + novoId);
        dados.link = 'https://drive.google.com/drive/folders/' + novaSubPasta.getId();
      } catch(e) {
        Logger.log('AVISO: não foi possível criar pasta no Drive: ' + e.message);
        dados.link = dados.link || '';
      }
    }

    linhaDados[COLUNAS_PROJETOS.ID]              = novoId;
    linhaDados[COLUNAS_PROJETOS.NOME]            = dados.nome;
    linhaDados[COLUNAS_PROJETOS.DESCRICAO]       = dados.descricao;
    linhaDados[COLUNAS_PROJETOS.TIPO]            = dados.tipo;
    linhaDados[COLUNAS_PROJETOS.PARA_QUEM]       = dados.paraQuem;
    linhaDados[COLUNAS_PROJETOS.STATUS]          = dados.status;
    linhaDados[COLUNAS_PROJETOS.PRIORIDADE]      = dados.prioridade;
    linhaDados[COLUNAS_PROJETOS.LINK]            = dados.link;
    linhaDados[COLUNAS_PROJETOS.GRAVIDADE]       = dados.gravidade   || '';
    linhaDados[COLUNAS_PROJETOS.URGENCIA]        = dados.urgencia    || '';
    linhaDados[COLUNAS_PROJETOS.ESFORCO]         = dados.esforco     || '';  // ← novo
    linhaDados[COLUNAS_PROJETOS.SETOR]           = dados.setor;
    linhaDados[COLUNAS_PROJETOS.PILAR]           = dados.pilar;
    linhaDados[COLUNAS_PROJETOS.RESPONSAVEIS_IDS]= JSON.stringify(dados.responsaveisIds || []);
    linhaDados[COLUNAS_PROJETOS.VALOR_PRIORIDADE]= dados.valorPrioridade || 0;
    linhaDados[COLUNAS_PROJETOS.DATA_INICIO]      = dados.dataInicio    || '';
    linhaDados[COLUNAS_PROJETOS.DATA_FIM]         = dados.dataFim       || '';

    // Busca departamentos ATUALIZADOS da planilha (ignora sessão cacheada)
    const depsUsuario = (sessao.perfil !== 'admin') ? _obterDepsAtualizadosUsuario(sessao) : [];
    let departamentoIdFinal = _resolverIdsDepartamento([dados.departamentoId || ''])[0] || (dados.departamentoId || '').toString().trim();
    const depExistenteRaw = (linha > 0 ? (linhaDados[COLUNAS_PROJETOS.DEPARTAMENTO_ID] || '').toString().trim() : '');
    const depExistente = _resolverIdsDepartamento([depExistenteRaw])[0] || depExistenteRaw;

    if (sessao.perfil === 'admin') {
      // Admin: preserva departamento existente se não enviou um novo
      if (!departamentoIdFinal && depExistente) departamentoIdFinal = depExistente;
    } else {
      // Não-admin: valida departamento
      if (!departamentoIdFinal) departamentoIdFinal = depExistente;
      if (!departamentoIdFinal && depsUsuario.length > 0) departamentoIdFinal = depsUsuario[0];

      // Se usuário tem departamentos, valida que o departamento escolhido é permitido
      if (depsUsuario.length > 0 && departamentoIdFinal && !depsUsuario.includes(departamentoIdFinal)) {
        return { sucesso: false, mensagem: 'Departamento inválido para este usuário.' };
      }
      // Se usuário sem departamentos: permite salvar (compatibilidade retroativa)
    }

    linhaDados[COLUNAS_PROJETOS.DEPARTAMENTO_ID]  = departamentoIdFinal;

    if (linha > 0) {
      aba.getRange(linha, 1, 1, linhaDados.length).setValues([linhaDados]);
    } else {
      aba.appendRow(linhaDados);
    }
    
    return { sucesso: true, id: novoId, link: linhaDados[COLUNAS_PROJETOS.LINK] || '' };
  } catch (e) {
    Logger.log('ERRO salvarProjetoCompleto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function listarEtapasPorProjeto(projetoId) {
  const aba = obterAba(NOME_ABA_ETAPAS);
  const dados = aba.getDataRange().getValues();
  const etapas = [];

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_ETAPAS.PROJETO_ID] == projetoId) {
      const respRaw = dados[i][COLUNAS_ETAPAS.RESPONSAVEIS_IDS];
      let respIds = [];
      if (respRaw) {
        try {
          respIds = JSON.parse(respRaw);
        } catch(e) {
          respIds = respRaw.toString() ? [respRaw.toString()] : [];
        }
      }

      etapas.push({
        id:              dados[i][COLUNAS_ETAPAS.ID],
        projetoId:       dados[i][COLUNAS_ETAPAS.PROJETO_ID],
        nome:            dados[i][COLUNAS_ETAPAS.NOME],
        oQueFazer:       dados[i][COLUNAS_ETAPAS.O_QUE_FAZER] || '',
        status:          dados[i][COLUNAS_ETAPAS.STATUS] || 'A Fazer',
        responsaveisIds: respIds
      });
    }
  }
  return etapas;
}

function salvarEtapaCompleta(dados, token) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    const sessao = auth.sessao;
    if (!_usuarioPodeAcessarProjetoPorDepartamento(sessao, dados && dados.projetoId)) {
      return { sucesso: false, mensagem: 'Sem acesso ao projeto informado.' };
    }

    const aba = obterAba(NOME_ABA_ETAPAS);
    const valores = aba.getDataRange().getValues();
    let linha = -1;

    if (dados.id) {
      for (let i = 1; i < valores.length; i++) {
        if (valores[i][COLUNAS_ETAPAS.ID] == dados.id) {
          linha = i + 1;
          break;
        }
      }
    }

    const novoId = dados.id || gerarId();

    // Monta linha com exatamente 6 colunas
    const linhaDados = ['', '', '', '', '', ''];
    linhaDados[COLUNAS_ETAPAS.ID]              = novoId;
    linhaDados[COLUNAS_ETAPAS.PROJETO_ID]      = dados.projetoId;
    linhaDados[COLUNAS_ETAPAS.RESPONSAVEIS_IDS]= JSON.stringify(dados.responsaveisIds || []);
    linhaDados[COLUNAS_ETAPAS.NOME]            = dados.nome || '';
    linhaDados[COLUNAS_ETAPAS.O_QUE_FAZER]     = dados.oQueFazer || '';
    linhaDados[COLUNAS_ETAPAS.STATUS]          = dados.status || 'A Fazer';

    if (linha > 0) {
      aba.getRange(linha, 1, 1, linhaDados.length).setValues([linhaDados]);
    } else {
      aba.appendRow(linhaDados);
    }

    return { sucesso: true, id: novoId };
  } catch (e) {
    Logger.log('ERRO salvarEtapaCompleta: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function excluirEtapaCompleta(id, token) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    const sessao = auth.sessao;

    const aba = obterAba(NOME_ABA_ETAPAS);
    const dados = aba.getDataRange().getValues();

    for (let i = dados.length - 1; i >= 1; i--) {
      if (dados[i][COLUNAS_ETAPAS.ID] === id) {
        const projetoId = (dados[i][COLUNAS_ETAPAS.PROJETO_ID] || '').toString();
        if (!_usuarioPodeAcessarProjetoPorDepartamento(sessao, projetoId)) {
          return { sucesso: false, mensagem: 'Sem acesso ao projeto desta atividade.' };
        }
        aba.deleteRow(i + 1);
        return { sucesso: true, mensagem: 'Atividade excluída.' };
      }
    }

    return { sucesso: false, mensagem: 'Atividade não encontrada.' };
  } catch (e) {
    Logger.log('ERRO excluirEtapaCompleta: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function listarResponsaveisCompletos() {
  const aba = obterAba(NOME_ABA_RESPONSAVEIS);
  const dados = aba.getDataRange().getValues();
  const res = [];
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_RESPONSAVEIS.ID]) {
      res.push({
        id: dados[i][COLUNAS_RESPONSAVEIS.ID],
        nome: dados[i][COLUNAS_RESPONSAVEIS.NOME],
        email: dados[i][COLUNAS_RESPONSAVEIS.EMAIL],
        cargo: dados[i][COLUNAS_RESPONSAVEIS.CARGO]
      });
    }
  }
  return res;
}

function salvarResponsavelCompleto(dados) {
  try {
    const aba = obterAba(NOME_ABA_RESPONSAVEIS);
    const valores = aba.getDataRange().getValues();
    let linha = -1;
    
    if (dados.id) {
      for (let i = 1; i < valores.length; i++) {
        if (valores[i][COLUNAS_RESPONSAVEIS.ID] == dados.id) {
          linha = i + 1;
          break;
        }
      }
    }
    
    const novoId = dados.id || gerarId();
    const linhaDados = [];
    for(let k in COLUNAS_RESPONSAVEIS) linhaDados[COLUNAS_RESPONSAVEIS[k]] = '';
    
    if (linha > 0) {
      const dadosAntigos = valores[linha-1];
      for(let k = 0; k < dadosAntigos.length; k++) linhaDados[k] = dadosAntigos[k];
    }
    
    linhaDados[COLUNAS_RESPONSAVEIS.ID] = novoId;
    linhaDados[COLUNAS_RESPONSAVEIS.NOME] = dados.nome;
    linhaDados[COLUNAS_RESPONSAVEIS.EMAIL] = dados.email;
    linhaDados[COLUNAS_RESPONSAVEIS.CARGO] = dados.cargo;
    
    if (linha > 0) {
      aba.getRange(linha, 1, 1, linhaDados.length).setValues([linhaDados]);
    } else {
      aba.appendRow(linhaDados);
    }
    
    return { sucesso: true, id: novoId };
  } catch (e) {
    Logger.log('ERRO salvarResponsavelCompleto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

// ===================== ENVIO DE EMAIL PARA RESPONSÁVEIS =====================

/**
 * Envia email para um ou mais responsáveis com resumo de seus projetos e etapas
 * @param {Array} idsResponsaveis - Lista de IDs dos responsáveis
 * @param {string} projetoId - ID do projeto atual (opcional, se vazio envia todos)
 * @param {string} assuntoPersonalizado - Assunto do email (opcional)
 * @param {string} mensagemAdicional - Mensagem adicional no corpo (opcional)
 * @returns {Object} Resultado da operação
 */
function enviarEmailParaResponsaveis(idsResponsaveis, projetoId, assuntoPersonalizado, mensagemAdicional) {
  Logger.log('Iniciando envio de emails. Responsáveis: ' + JSON.stringify(idsResponsaveis));
  
  try {
    if (!idsResponsaveis || idsResponsaveis.length === 0) {
      return { sucesso: false, mensagem: 'Nenhum responsável selecionado.' };
    }
    
    const todosResponsaveis = listarResponsaveisCompletos();
    const todosProjetos = listarProjetosCompletos();
    const todasEtapas = obterTodasEtapas();
    
    let emailsEnviados = 0;
    let erros = [];
    
    for (const idResp of idsResponsaveis) {
      const responsavel = todosResponsaveis.find(r => r.id === idResp);
      
      if (!responsavel || !responsavel.email) {
        erros.push(`Responsável ${idResp} sem email cadastrado.`);
        continue;
      }
      
      // Filtra projetos e etapas do responsável
      let projetosDoResponsavel = todosProjetos.filter(p => 
        p.responsaveisIds && p.responsaveisIds.includes(idResp)
      );
      
      let etapasDoResponsavel = todasEtapas.filter(e => 
        e.responsaveisIds && e.responsaveisIds.includes(idResp)
      );
      
      // Se projetoId foi passado, filtra apenas esse projeto
      if (projetoId) {
        projetosDoResponsavel = projetosDoResponsavel.filter(p => p.id === projetoId);
        etapasDoResponsavel = etapasDoResponsavel.filter(e => e.projetoId === projetoId);
      }
      
      // Monta o corpo do email
      const corpoEmail = montarCorpoEmailResponsavel(
        responsavel, 
        projetosDoResponsavel, 
        etapasDoResponsavel,
        todosProjetos,
        mensagemAdicional
      );
      
      const assunto = assuntoPersonalizado || 
        `📋 Resumo de Projetos e Etapas - ${responsavel.nome}`;
      
      try {
        MailApp.sendEmail({
          to: responsavel.email,
          subject: assunto,
          htmlBody: corpoEmail,
          name: 'Smart Meeting',
          replyTo: 'setorbiunichristus@gmail.com'
        });
        emailsEnviados++;
        Logger.log('Email enviado para: ' + responsavel.email);
      } catch (emailErro) {
        erros.push(`Erro ao enviar para ${responsavel.email}: ${emailErro.message}`);
        Logger.log('ERRO envio email: ' + emailErro.toString());
      }
    }
    
    if (emailsEnviados > 0) {
      return { 
        sucesso: true, 
        mensagem: `${emailsEnviados} email(s) enviado(s) com sucesso!`,
        erros: erros.length > 0 ? erros : null
      };
    } else {
      return { 
        sucesso: false, 
        mensagem: 'Nenhum email foi enviado.',
        erros: erros
      };
    }
    
  } catch (e) {
    Logger.log('ERRO enviarEmailParaResponsaveis: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

function montarCorpoEmailResponsavel(responsavel, projetos, etapas, todosProjetos, mensagemAdicional) {
  // Delegar para a versão com prioridades (sem prioridades = array vazio)
  return montarCorpoEmailComPrioridades(
    responsavel, projetos, etapas, todosProjetos, mensagemAdicional, []
  );
}

function montarCorpoEmailComPrioridades(responsavel, projetos, etapas, todosProjetos, mensagemAdicional, prioridades) {
  var urlBase = ScriptApp.getService().getUrl();
  
  // ── Mapa de prioridades para lookup rápido ──
  var mapaPrioridades = {};
  (prioridades || []).forEach(function(p) {
    mapaPrioridades[p.itemId] = p.ordemPrioridade;
  });

  // ── Mapa de responsáveis para lookup rápido (id → nome) ──
  var todosResponsaveis = listarResponsaveisCompletos();
  var mapaResponsaveis = {};
  todosResponsaveis.forEach(function(r) { mapaResponsaveis[r.id] = r.nome; });

  // ── Ordenar atividades por prioridade (se existir) ──
  var etapasOrdenadas = (etapas || []).slice().sort(function(a, b) {
    var prioA = mapaPrioridades[a.id] || 9999;
    var prioB = mapaPrioridades[b.id] || 9999;
    return prioA - prioB;
  });

  // ── Helpers inline ──
  function badgeStatus(status) {
    var mapa = {
      'A Fazer':          { bg: '#fef3c7', cor: '#92400e', label: '📋 A Fazer'          },
      'Em Andamento':     { bg: '#dbeafe', cor: '#1e40af', label: '🔄 Em Andamento'     },
      'Aguardando Setor': { bg: '#fde8d8', cor: '#9a3412', label: '⏳ Aguardando Setor' },
      'Concluída':        { bg: '#d1fae5', cor: '#065f46', label: '✅ Concluída'        },
      'Suspenso':         { bg: '#ede9f7', cor: '#5a3e9e', label: '⏸ Suspenso'         }
    };
    var cfg = mapa[status] || mapa['A Fazer'];
    return '<span style="display:inline-block;padding:3px 10px;border-radius:20px;font-size:0.72rem;font-weight:700;background:' +
      cfg.bg + ';color:' + cfg.cor + ';">' + cfg.label + '</span>';
  }

  function badgePrioridade(valor) {
    if (!valor || valor === 0) return '<span style="font-size:0.72rem;color:#9ca3af;font-style:italic;">Sem score</span>';
    var cor, label;
    if (valor >= 2102)      { cor = '#dc2626'; label = '🔴 ALTA — ' + valor;  }
    else if (valor >= 1078) { cor = '#ca8a04'; label = '🟡 MÉDIA — ' + valor; }
    else                    { cor = '#059669'; label = '🟢 BAIXA — ' + valor; }
    return '<span style="display:inline-block;padding:3px 10px;border-radius:20px;font-size:0.72rem;font-weight:700;background:rgba(0,0,0,0.06);color:' + cor + ';">' + label + '</span>';
  }

  function nomesResponsaveis(ids) {
    if (!ids || ids.length === 0) return '<span style="color:#9ca3af;font-style:italic;">Sem responsável</span>';
    return ids.map(function(id) {
      return '<span style="display:inline-block;background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:10px;font-size:0.7rem;font-weight:600;margin:1px 2px;">' +
        (mapaResponsaveis[id] || id) + '</span>';
    }).join('');
  }

  function linhaInfo(icone, label, valor) {
    if (!valor || String(valor).trim() === '') return '';
    return '<tr>' +
      '<td style="padding:4px 8px 4px 0;font-size:0.74rem;color:#9a8a78;font-weight:600;white-space:nowrap;vertical-align:top;">' + icone + ' ' + label + '</td>' +
      '<td style="padding:4px 0;font-size:0.78rem;color:#3d2b1f;vertical-align:top;">' + valor + '</td>' +
      '</tr>';
  }

  // ── Cabeçalho HTML ──
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:\'Segoe UI\',Arial,sans-serif;line-height:1.6;color:#333;max-width:720px;margin:0 auto;background:#f5f0e8;">';

  // Banner principal
  html += '<div style="background:linear-gradient(135deg,#78350f 0%,#92400e 100%);color:white;padding:28px 24px;border-radius:12px 12px 0 0;">' +
    '<h1 style="margin:0;font-size:1.5rem;font-weight:700;">📋 Resumo de Atividades</h1>' +
    '<p style="margin:8px 0 0;opacity:0.9;font-size:0.95rem;">Olá, <strong>' + responsavel.nome + '</strong>! Aqui está seu resumo atualizado.</p>' +
    '</div>';

  html += '<div style="background:#fffbf2;padding:24px;border:1px solid #fcd34d;border-top:none;border-radius:0 0 12px 12px;">';

  // ── Mensagem adicional ──
  if (mensagemAdicional && mensagemAdicional.trim()) {
    html += '<div style="background:#f0f9ff;border-left:4px solid #3b82f6;padding:14px 16px;margin-bottom:22px;border-radius:0 8px 8px 0;">' +
      '<strong style="color:#1e40af;">📝 Mensagem:</strong><br>' +
      '<span style="font-size:0.88rem;">' + mensagemAdicional.replace(/\n/g, '<br>') + '</span>' +
      '</div>';
  }

  // ══════════════════════════════════════════
  //  SEÇÃO: PROJETOS
  // ══════════════════════════════════════════
  if (projetos && projetos.length > 0) {
    html += '<div style="margin-bottom:28px;">' +
      '<div style="color:#78350f;font-size:1.05rem;font-weight:700;margin-bottom:14px;border-bottom:2px solid #fcd34d;padding-bottom:8px;">' +
        '📁 Projetos sob sua responsabilidade (' + projetos.length + ')' +
      '</div>';

    projetos.forEach(function(proj) {
      var valorPrio = proj.valorPrioridade || 0;

      html += '<div style="background:white;border:1px solid #e7e5e4;border-radius:10px;padding:18px;margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,0.06);">';

      // Nome + badges
      html += '<div style="margin-bottom:12px;">' +
        '<div style="font-weight:700;color:#451a03;font-size:1rem;margin-bottom:8px;">' + proj.nome + '</div>' +
        '<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;">' +
          badgeStatus(proj.status) +
          badgePrioridade(valorPrio) +
          (proj.tipo ? '<span style="display:inline-block;padding:3px 10px;border-radius:20px;font-size:0.72rem;font-weight:600;background:#e0e7ff;color:#3730a3;">' + proj.tipo + '</span>' : '') +
        '</div>' +
      '</div>';

      // Descrição
      if (proj.descricao && proj.descricao.trim()) {
        html += '<p style="font-size:0.84rem;color:#57534e;margin:0 0 12px;padding:10px 12px;background:#fafaf9;border-radius:6px;border-left:3px solid #fcd34d;">' +
          proj.descricao + '</p>';
      }

      // Tabela de informações detalhadas
      html += '<table style="width:100%;border-collapse:collapse;margin-bottom:12px;">';
      html += linhaInfo('📅', 'Início:', proj.dataInicio ? _formatarDataEmail(proj.dataInicio) : '');
      html += linhaInfo('🏁', 'Fim:', proj.dataFim ? _formatarDataEmail(proj.dataFim) : '');
      html += linhaInfo('⏱️', 'Duração:', proj.duracaoFormatada || '');
      html += linhaInfo('🏢', 'Setor:', proj.setor || '');
      html += linhaInfo('🏷️', 'Tipo:', proj.tipo || '');
      html += linhaInfo('👥', 'Para Quem:', proj.paraQuem || '');
      html += linhaInfo('⚠️', 'Gravidade:', proj.gravidade || '');
      html += linhaInfo('🕐', 'Urgência:', proj.urgencia || '');
      html += linhaInfo('⚡', 'Esforço:', proj.esforco || '');
      html += linhaInfo('🏛️', 'Pilar:', proj.pilar || '');
      html += '</table>';

      // Responsáveis
      html += '<div style="margin-bottom:10px;">' +
        '<span style="font-size:0.72rem;color:#9a8a78;font-weight:600;">👤 Responsáveis: </span>' +
        nomesResponsaveis(proj.responsaveisIds || []) +
      '</div>';

      // Botão abrir projeto
      html += '<a href="' + urlBase + '?pagina=projeto&id=' + proj.id + '" ' +
        'style="display:inline-block;background:#b45309;color:white;padding:9px 18px;border-radius:8px;text-decoration:none;font-weight:600;font-size:0.82rem;">Abrir Projeto →</a>';

      // Link do Drive (se existir)
      if (proj.link && proj.link.trim()) {
        html += ' <a href="' + proj.link + '" target="_blank" ' +
          'style="display:inline-block;background:#059669;color:white;padding:9px 18px;border-radius:8px;text-decoration:none;font-weight:600;font-size:0.82rem;margin-left:6px;">📁 Documentos</a>';
      }

      html += '</div>'; // /card projeto
    });

    html += '</div>'; // /seção projetos
  }

  // ══════════════════════════════════════════
  //  SEÇÃO: ATIVIDADES
  // ══════════════════════════════════════════
  if (etapasOrdenadas && etapasOrdenadas.length > 0) {
    html += '<div style="margin-bottom:28px;">' +
      '<div style="color:#78350f;font-size:1.05rem;font-weight:700;margin-bottom:14px;border-bottom:2px solid #fcd34d;padding-bottom:8px;">' +
        '✅ Atividades atribuídas a você (' + etapasOrdenadas.length + ')' +
      '</div>';

    etapasOrdenadas.forEach(function(etapa) {
      var projeto = todosProjetos.find(function(p) { return p.id === etapa.projetoId; });
      var nomeProjeto = projeto ? projeto.nome : 'Projeto não encontrado';
      var numeroPrioridade = mapaPrioridades[etapa.id];

      html += '<div style="background:white;border:1px solid #e7e5e4;border-radius:10px;padding:16px;margin-bottom:10px;box-shadow:0 1px 4px rgba(0,0,0,0.06);' +
        (numeroPrioridade ? 'border-left:4px solid #b45309;' : '') + '">';

      // Número de prioridade (se existir)
      if (numeroPrioridade) {
        html += '<div style="display:inline-block;background:linear-gradient(135deg,#b45309,#92400e);color:white;width:26px;height:26px;border-radius:50%;text-align:center;line-height:26px;font-size:0.78rem;font-weight:700;margin-bottom:8px;">' +
          numeroPrioridade + '</div> ';
      }

      // Nome da atividade
      html += '<span style="font-weight:700;color:#451a03;font-size:0.95rem;">' + etapa.nome + '</span>';

      // Status + projeto
      html += '<div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:6px;align-items:center;">' +
        badgeStatus(etapa.status || 'A Fazer') +
        '<span style="font-size:0.75rem;color:#78350f;background:#fef3c7;padding:3px 8px;border-radius:4px;">📁 ' + nomeProjeto + '</span>' +
      '</div>';

      // Responsáveis da atividade
      if (etapa.responsaveisIds && etapa.responsaveisIds.length > 0) {
        html += '<div style="margin-top:8px;">' +
          '<span style="font-size:0.72rem;color:#9a8a78;font-weight:600;">👤 </span>' +
          nomesResponsaveis(etapa.responsaveisIds) +
        '</div>';
      }

      // O que fazer
      if (etapa.oQueFazer && etapa.oQueFazer.trim()) {
        html += '<div style="margin-top:10px;padding:10px 12px;background:#fafaf9;border-radius:6px;border-left:3px solid #fcd34d;">' +
          '<div style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;color:#9a8a78;margin-bottom:4px;">📋 O que fazer:</div>' +
          '<div style="font-size:0.82rem;color:#3d2b1f;line-height:1.6;">' + etapa.oQueFazer.replace(/\n/g, '<br>') + '</div>' +
        '</div>';
      }

      html += '</div>'; // /card atividade
    });

    html += '</div>'; // /seção atividades
  }

  // ── Estado vazio ──
  if ((!projetos || projetos.length === 0) && (!etapasOrdenadas || etapasOrdenadas.length === 0)) {
    html += '<div style="text-align:center;padding:32px;color:#78350f;">' +
      '<div style="font-size:2.5rem;margin-bottom:12px;">📭</div>' +
      '<p style="font-size:0.88rem;">Você não possui projetos ou atividades atribuídos no momento.</p>' +
    '</div>';
  }

  // ── Rodapé ──
  html += '<div style="text-align:center;padding:20px;border-top:1px solid #fcd34d;margin-top:8px;">' +
    '<p style="font-size:0.78rem;color:#78350f;margin:0 0 12px;">Email enviado automaticamente pelo sistema <strong>Smart Meeting</strong></p>' +
    '<a href="' + urlBase + '?pagina=projetos" style="display:inline-block;background:#b45309;color:white;padding:10px 22px;border-radius:8px;text-decoration:none;font-weight:600;font-size:0.85rem;">🚀 Acessar Sistema</a>' +
  '</div>';

  html += '</div></body></html>';
  return html;
}

// Helper privado: formata data ISO (yyyy-MM-dd) para dd/mm/yyyy
function _formatarDataEmail(dataIso) {
  if (!dataIso) return '';
  var p = String(dataIso).split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : dataIso;
}

/**
 * Retorna a classe CSS do badge de status
 */
function obterClasseBadgeStatus(status) {
  const mapa = {
    'A Fazer': 'badge-afazer',
    'Em Andamento': 'badge-andamento',
    'Bloqueada': 'badge-bloqueada',
    'Concluída': 'badge-concluida',
    'Suspenso': 'badge-suspenso'
  };
  return mapa[status] || 'badge-afazer';
}

/**
 * Obtém todas as etapas do sistema
 */
function obterTodasEtapas() {
  try {
    const aba = obterAba(NOME_ABA_ETAPAS);
    const dados = aba.getDataRange().getValues();
    const etapas = [];

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_ETAPAS.ID]) {
        const respRaw = dados[i][COLUNAS_ETAPAS.RESPONSAVEIS_IDS];
        let respIds = [];
        if (respRaw) {
          try { respIds = JSON.parse(respRaw); }
          catch(e) { respIds = respRaw.toString() ? [respRaw.toString()] : []; }
        }

        etapas.push({
          id:              dados[i][COLUNAS_ETAPAS.ID],
          projetoId:       dados[i][COLUNAS_ETAPAS.PROJETO_ID],
          nome:            dados[i][COLUNAS_ETAPAS.NOME],
          oQueFazer:       dados[i][COLUNAS_ETAPAS.O_QUE_FAZER] || '',
          status:          dados[i][COLUNAS_ETAPAS.STATUS] || 'A Fazer',
          responsaveisIds: respIds
        });
      }
    }
    return etapas;
  } catch (e) {
    Logger.log('ERRO obterTodasEtapas: ' + e.toString());
    return [];
  }
}

function obterResponsaveisComEtapasDoProjeto(projetoId) {
  try {
    const projeto = listarProjetosCompletos().find(p => p.id === projetoId);
    const etapas = listarEtapasPorProjeto(projetoId);
    const todosResponsaveis = listarResponsaveisCompletos();
    
    // Coleta todos os IDs de responsáveis (do projeto e das etapas)
    const idsResponsaveis = new Set();
    
    if (projeto && projeto.responsaveisIds) {
      projeto.responsaveisIds.forEach(id => idsResponsaveis.add(id));
    }
    
    etapas.forEach(etapa => {
      if (etapa.responsaveisIds) {
        etapa.responsaveisIds.forEach(id => idsResponsaveis.add(id));
      }
    });
    
    // Monta a lista com detalhes
    const resultado = [];
    
    for (const idResp of idsResponsaveis) {
      const resp = todosResponsaveis.find(r => r.id === idResp);
      if (!resp) continue;
      
      const etapasDoResp = etapas.filter(e => 
        e.responsaveisIds && e.responsaveisIds.includes(idResp)
      );
      
      const ehResponsavelProjeto = projeto && 
        projeto.responsaveisIds && 
        projeto.responsaveisIds.includes(idResp);
      
      resultado.push({
        id: resp.id,
        nome: resp.nome,
        email: resp.email,
        cargo: resp.cargo,
        ehResponsavelProjeto: ehResponsavelProjeto,
        quantidadeEtapas: etapasDoResp.length,
        etapas: etapasDoResp.map(e => ({
          id: e.id,
          nome: e.nome,
          status: e.status
        }))
      });
    }
    
    return { sucesso: true, responsaveis: resultado };
    
  } catch (e) {
    Logger.log('ERRO obterResponsaveisComEtapasDoProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message, responsaveis: [] };
  }
}

// ===================== SISTEMA DE PRIORIDADES =====================

function obterPrioridadesDoProjeto(projetoId) {
  try {
    if (!projetoId) {
      return { sucesso: false, mensagem: 'ID do projeto é obrigatório', prioridades: [] };
    }
    
    const aba = obterAba(NOME_ABA_PRIORIDADES);
    if (!aba || aba.getLastRow() <= 1) {
      return { sucesso: true, prioridades: [] };
    }
    
    const dados = aba.getDataRange().getValues();
    const prioridades = [];
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA] === projetoId) {
        prioridades.push({
          id: dados[i][COLUNAS_PRIORIDADES.ID],
          responsavelId: dados[i][COLUNAS_PRIORIDADES.RESPONSAVEL_ID],
          tipoItem: dados[i][COLUNAS_PRIORIDADES.TIPO_ITEM],
          itemId: dados[i][COLUNAS_PRIORIDADES.ITEM_ID],
          ordemPrioridade: dados[i][COLUNAS_PRIORIDADES.ORDEM_PRIORIDADE],
          projetoReferencia: dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA]
        });
      }
    }
    
    // Ordenar por ordem de prioridade
    prioridades.sort((a, b) => a.ordemPrioridade - b.ordemPrioridade);
    
    return { sucesso: true, prioridades: prioridades };
  } catch (e) {
    Logger.log('ERRO obterPrioridadesDoProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message, prioridades: [] };
  }
}

function salvarPrioridadesDoProjeto(projetoId, listaPrioridades) {
  try {
    if (!projetoId) {
      return { sucesso: false, mensagem: 'ID do projeto é obrigatório' };
    }
    
    const aba = obterAba(NOME_ABA_PRIORIDADES);
    const dados = aba.getDataRange().getValues();
    
    // Remover prioridades antigas deste projeto (de trás para frente)
    for (let i = dados.length - 1; i >= 1; i--) {
      if (dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA] === projetoId) {
        aba.deleteRow(i + 1);
      }
    }
    
    // Inserir novas prioridades
    if (listaPrioridades && listaPrioridades.length > 0) {
      const novasLinhas = listaPrioridades.map((item, index) => {
        return [
          gerarId(),                        // ID
          item.responsavelId || '',         // RESPONSAVEL_ID
          item.tipoItem || 'etapa',         // TIPO_ITEM
          item.itemId,                      // ITEM_ID
          index + 1,                        // ORDEM_PRIORIDADE (1-based)
          projetoId                         // PROJETO_REFERENCIA
        ];
      });
      
      // Adicionar todas as linhas de uma vez
      const ultimaLinha = aba.getLastRow();
      aba.getRange(ultimaLinha + 1, 1, novasLinhas.length, novasLinhas[0].length)
         .setValues(novasLinhas);
    }
    
    return { sucesso: true, mensagem: `${listaPrioridades.length} prioridade(s) salva(s)!` };
  } catch (e) {
    Logger.log('ERRO salvarPrioridadesDoProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function obterPrioridadeEtapa(etapaId, projetoId) {
  try {
    const resultado = obterPrioridadesDoProjeto(projetoId);
    if (!resultado.sucesso) return null;
    
    const prioridade = resultado.prioridades.find(p => p.itemId === etapaId && p.tipoItem === 'etapa');
    return prioridade ? prioridade.ordemPrioridade : null;
  } catch (e) {
    Logger.log('ERRO obterPrioridadeEtapa: ' + e.toString());
    return null;
  }
}

/**
 * Gera e envia um relatório consolidado de projetos selecionados por email.
 * opcoes: { projetoIds, apenasAtivos, incluirDescricao, incluirAtividades,
 *           incluirPrioridade, destinatariosIds, emailsCC, assunto, mensagem }
 */
function enviarRelatorioProjetosEmail(token, opcoes) {
  try {
    const sessao = _obterSessao(token);
    if (!sessao) return { sucesso: false, mensagem: 'Sessão inválida.' };

    opcoes = opcoes || {};
    const projetoIds       = opcoes.projetoIds       || [];
    const apenasAtivos     = opcoes.apenasAtivos     !== false;
    const incluirDescricao = opcoes.incluirDescricao !== false;
    const incluirAtividades= opcoes.incluirAtividades !== false;
    const incluirPrioridade= opcoes.incluirPrioridade !== false;
    const destinatariosIds = opcoes.destinatariosIds || [];
    const emailsCC         = (opcoes.emailsCC || []).filter(e => e && e.includes('@'));
    const assunto          = opcoes.assunto || '📊 Relatório de Projetos';
    const mensagem         = opcoes.mensagem || '';

    if (projetoIds.length === 0) return { sucesso: false, mensagem: 'Nenhum projeto selecionado.' };

    const todosResponsaveis = listarResponsaveisCompletos();
    const todosProjetos     = listarProjetosCompletos(token);
    const todasEtapas       = obterTodasEtapas();

    // Filtra e ordena projetos por prioridade
    const projetos = todosProjetos
      .filter(p => projetoIds.includes(p.id))
      .sort((a, b) => (b.valorPrioridade || 0) - (a.valorPrioridade || 0));

    if (projetos.length === 0) return { sucesso: false, mensagem: 'Projetos não encontrados.' };

    // Resolve destinatários
    const emailsTo = destinatariosIds
      .map(id => todosResponsaveis.find(r => r.id === id))
      .filter(r => r && r.email)
      .map(r => r.email);

    if (emailsTo.length === 0 && emailsCC.length === 0) {
      return { sucesso: false, mensagem: 'Nenhum destinatário com email válido.' };
    }

    // Monta HTML do relatório — usa corpoHtml pré-gerado (personalizado) se fornecido
    const corpo = opcoes.corpoHtml || _montarCorpoRelatorioHTML(projetos, todasEtapas, todosResponsaveis, {
      apenasAtivos, incluirDescricao, incluirAtividades, incluirPrioridade, mensagem
    });

    // Envia
    const emailOpts = { subject: assunto, htmlBody: corpo, name: 'Smart Meeting', replyTo: 'setorbiunichristus@gmail.com' };
    if (emailsCC.length > 0) emailOpts.cc = emailsCC.join(',');

    if (emailsTo.length > 0) {
      emailOpts.to = emailsTo.join(',');
      MailApp.sendEmail(emailOpts);
    } else {
      // Apenas CC: envia para o primeiro CC, restante em CC
      MailApp.sendEmail(Object.assign({}, emailOpts, {
        to: emailsCC[0],
        cc: emailsCC.slice(1).join(',') || undefined
      }));
    }

    Logger.log('[enviarRelatorioProjetosEmail] Enviado para: ' + (emailsTo.concat(emailsCC)).join(', '));
    return { sucesso: true, mensagem: 'Relatório enviado para ' + (emailsTo.length + emailsCC.length) + ' destinatário(s).' };

  } catch (e) {
    Logger.log('ERRO enviarRelatorioProjetosEmail: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function _montarCorpoRelatorioHTML(projetos, todasEtapas, todosResponsaveis, opc) {
  const data = new Date().toLocaleDateString('pt-BR', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

  function corStatus(status) {
    const m = {
      'A Fazer':          { bg: '#f0e6d3', cor: '#8b6b58' },
      'Em Andamento':     { bg: '#e0e8f0', cor: '#4a6fa5' },
      'Aguardando Setor': { bg: '#f5ede3', cor: '#c2855a' },
      'Concluída':        { bg: '#e8f0e5', cor: '#5a7247' },
      'Suspenso':         { bg: '#ede9f7', cor: '#7c5cbf' }
    };
    return m[status] || { bg: '#f0ece8', cor: '#7a6555' };
  }

  function corPrioridade(valor) {
    if (!valor) return null;
    if (valor >= 2102) return { bg: '#fce8e8', cor: '#a63d40', label: '🔴 ALTA' };
    if (valor >= 1078) return { bg: '#fef9e7', cor: '#b8860b', label: '🟡 MÉDIA' };
    return { bg: '#e8f5e9', cor: '#5a7247', label: '🟢 BAIXA' };
  }

  function nomeResps(ids) {
    if (!ids || ids.length === 0) return '—';
    return ids.map(id => { const r = todosResponsaveis.find(x => x.id === id); return r ? r.nome : null; })
      .filter(Boolean).join(', ') || '—';
  }

  // Contadores globais
  let totalAtvGlobal = 0;
  projetos.forEach(p => {
    let atv = todasEtapas.filter(e => e.projetoId === p.id);
    if (opc.apenasAtivos) atv = atv.filter(e => e.status !== 'Concluída' && e.status !== 'Suspenso');
    totalAtvGlobal += atv.length;
  });

  // ── Monta HTML ──
  let corpo = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>body{margin:0;padding:0;background:#f5f0e8;font-family:'Segoe UI',Arial,sans-serif;}
table{border-collapse:collapse;}td,th{padding:0;}
</style></head><body>
<div style="max-width:700px;margin:28px auto;background:#fffdf8;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.12);">`;

  // Banner
  corpo += `<div style="background:linear-gradient(135deg,#3d2b1f 0%,#6b4c3b 55%,#8b6b58 100%);padding:28px 28px 22px;">
<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;">
<div style="width:52px;height:52px;background:rgba(255,255,255,0.14);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:1.6rem;">📊</div>
<div><h1 style="margin:0;color:#faf6ee;font-size:1.35rem;font-family:Georgia,serif;">Relatório de Projetos</h1>
<p style="margin:3px 0 0;color:rgba(250,246,238,0.7);font-size:0.78rem;">${data}</p></div></div>
<div style="display:flex;gap:12px;flex-wrap:wrap;">
<div style="background:rgba(255,255,255,0.12);border-radius:8px;padding:8px 16px;text-align:center;">
<div style="color:rgba(250,246,238,0.65);font-size:0.62rem;text-transform:uppercase;letter-spacing:0.08em;">Projetos</div>
<div style="color:#faf6ee;font-size:1.2rem;font-weight:700;">${projetos.length}</div></div>
<div style="background:rgba(255,255,255,0.12);border-radius:8px;padding:8px 16px;text-align:center;">
<div style="color:rgba(250,246,238,0.65);font-size:0.62rem;text-transform:uppercase;letter-spacing:0.08em;">${opc.apenasAtivos ? 'Pendentes' : 'Atividades'}</div>
<div style="color:#faf6ee;font-size:1.2rem;font-weight:700;">${totalAtvGlobal}</div></div>
</div></div>`;

  // Mensagem de abertura
  if (opc.mensagem) {
    corpo += `<div style="padding:14px 28px;background:#faf6ee;border-bottom:1px solid #d4c5b0;font-size:0.85rem;color:#5a4333;line-height:1.6;">${opc.mensagem.replace(/\n/g, '<br>')}</div>`;
  }

  // Projetos
  corpo += `<div style="padding:22px 28px;">`;

  projetos.forEach((p, idx) => {
    let etapas = todasEtapas.filter(e => e.projetoId === p.id);
    if (opc.apenasAtivos) etapas = etapas.filter(e => e.status !== 'Concluída' && e.status !== 'Suspenso');

    const totalEtapas  = todasEtapas.filter(e => e.projetoId === p.id).length;
    const conclEtapas  = todasEtapas.filter(e => e.projetoId === p.id && e.status === 'Concluída').length;
    const pct          = totalEtapas > 0 ? Math.round((conclEtapas / totalEtapas) * 100) : 0;
    const corPct       = pct === 100 ? '#5a7247' : pct >= 50 ? '#b8860b' : '#6b4c3b';
    const stCor        = corStatus(p.status);
    const prioCor      = opc.incluirPrioridade ? corPrioridade(p.valorPrioridade) : null;
    const respsStr     = nomeResps(p.responsaveisIds);

    corpo += `<div style="margin-bottom:20px;background:#fffdf8;border:1px solid #d4c5b0;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(74,50,40,0.07);">`;

    // Header do projeto
    corpo += `<div style="padding:14px 18px;background:linear-gradient(135deg,#faf6ee,#f0e8d8);border-bottom:1px solid #d4c5b0;">
<div style="display:flex;align-items:flex-start;gap:12px;">
<div style="flex:1;">
<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:7px;">
<span style="font-size:0.72rem;color:#9a8878;font-weight:700;">#${String(idx + 1).padStart(2, '0')}</span>
<strong style="font-size:1rem;color:#3d2b1f;font-family:Georgia,serif;">${p.nome || ''}</strong>
<span style="display:inline-block;padding:2px 10px;border-radius:100px;font-size:0.68rem;font-weight:600;background:${stCor.bg};color:${stCor.cor};">${p.status || 'A Fazer'}</span>
${prioCor ? `<span style="display:inline-block;padding:2px 10px;border-radius:100px;font-size:0.68rem;font-weight:600;background:${prioCor.bg};color:${prioCor.cor};">${prioCor.label}</span>` : ''}
${p.tipo ? `<span style="display:inline-block;padding:2px 10px;border-radius:100px;font-size:0.68rem;background:#f0ece8;color:#7a6555;">${p.tipo}</span>` : ''}
</div>
<div style="display:flex;flex-wrap:wrap;gap:14px;font-size:0.72rem;color:#7a6555;">
${p.setor ? `<span>🏢 ${p.setor}</span>` : ''}
${respsStr !== '—' ? `<span>👤 ${respsStr}</span>` : ''}
${p.dataInicio ? `<span>📅 ${p.dataInicio}</span>` : ''}
${p.dataFim ? `<span>🏁 ${p.dataFim}</span>` : ''}
${opc.incluirPrioridade && p.valorPrioridade ? `<span>⚡ Score: ${p.valorPrioridade}</span>` : ''}
</div>
</div>
<div style="text-align:center;min-width:56px;">
<div style="font-size:1.35rem;font-weight:700;color:${corPct};">${pct}%</div>
<div style="font-size:0.62rem;color:#9a8878;">${conclEtapas}/${totalEtapas} concl.</div>
</div>
</div></div>`;

    // Barra de progresso
    corpo += `<div style="height:4px;background:#e8dfd2;"><div style="height:100%;width:${pct}%;background:${corPct};transition:width 0.4s;"></div></div>`;

    // Descrição
    if (opc.incluirDescricao && p.descricao) {
      corpo += `<div style="padding:10px 18px;border-bottom:1px solid #e8dfd2;font-size:0.8rem;color:#5a4333;line-height:1.55;">${p.descricao}</div>`;
    }

    // Atividades
    if (opc.incluirAtividades) {
      corpo += `<div style="padding:12px 18px;">`;
      if (etapas.length === 0) {
        corpo += `<p style="margin:0;font-size:0.78rem;color:#9a8878;font-style:italic;">Nenhuma atividade${opc.apenasAtivos ? ' pendente' : ''}.</p>`;
      } else {
        corpo += `<table style="width:100%;border-collapse:collapse;font-size:0.78rem;">
<thead><tr style="background:#f5f0e8;">
<th style="padding:6px 10px;text-align:left;color:#7a6555;font-weight:600;border-bottom:2px solid #d4c5b0;font-size:0.68rem;text-transform:uppercase;">Atividade</th>
<th style="padding:6px 10px;text-align:center;color:#7a6555;font-weight:600;border-bottom:2px solid #d4c5b0;font-size:0.68rem;text-transform:uppercase;white-space:nowrap;">Status</th>
<th style="padding:6px 10px;text-align:left;color:#7a6555;font-weight:600;border-bottom:2px solid #d4c5b0;font-size:0.68rem;text-transform:uppercase;">Responsável</th>
</tr></thead><tbody>`;
        etapas.forEach((e, ei) => {
          const ec  = corStatus(e.status);
          const bgR = ei % 2 === 0 ? '#fffdf8' : '#faf6ee';
          corpo += `<tr style="background:${bgR};">
<td style="padding:7px 10px;border-bottom:1px solid #e8dfd2;color:#3d2b1f;">${e.nome || ''}</td>
<td style="padding:7px 10px;border-bottom:1px solid #e8dfd2;text-align:center;">
<span style="display:inline-block;padding:2px 9px;border-radius:100px;font-size:0.66rem;font-weight:600;background:${ec.bg};color:${ec.cor};">${e.status || ''}</span></td>
<td style="padding:7px 10px;border-bottom:1px solid #e8dfd2;color:#7a6555;font-size:0.75rem;">${nomeResps(e.responsaveisIds)}</td>
</tr>`;
        });
        corpo += `</tbody></table>`;
      }
      corpo += `</div>`;
    }

    corpo += `</div>`; // /projeto card
  });

  corpo += `</div>`; // /projetos area

  // Footer
  corpo += `<div style="padding:14px 28px;background:#f5f0e8;border-top:1px solid #d4c5b0;text-align:center;">
<p style="margin:0;font-size:0.7rem;color:#9a8878;">Gerado automaticamente pelo <strong>Smart Meeting</strong></p>
</div>`;

  corpo += `</div></body></html>`;
  return corpo;
}

function enviarEmailParaResponsaveisComCopia(idsResponsaveis, projetoId, assuntoPersonalizado, mensagemAdicional, emailsEmCopia) {
  Logger.log('Iniciando envio de emails com CC. Responsáveis: ' + JSON.stringify(idsResponsaveis));
  Logger.log('Emails em cópia: ' + JSON.stringify(emailsEmCopia));
  
  try {
    if (!idsResponsaveis || idsResponsaveis.length === 0) {
      return { sucesso: false, mensagem: 'Nenhum responsável selecionado.' };
    }
    
    const todosResponsaveis = listarResponsaveisCompletos();
    const todosProjetos = listarProjetosCompletos();
    const todasEtapas = obterTodasEtapas();
    
    // Obter prioridades do projeto
    let prioridadesProjeto = [];
    if (projetoId) {
      const resPrioridades = obterPrioridadesDoProjeto(projetoId);
      if (resPrioridades.sucesso) {
        prioridadesProjeto = resPrioridades.prioridades;
      }
    }
    
    // Validar emails em cópia
    const emailsCcValidos = [];
    if (emailsEmCopia && Array.isArray(emailsEmCopia)) {
      emailsEmCopia.forEach(email => {
        const emailLimpo = email.trim();
        if (emailLimpo && emailLimpo.includes('@')) {
          emailsCcValidos.push(emailLimpo);
        }
      });
    }
    
    let emailsEnviados = 0;
    let erros = [];
    
    for (const idResp of idsResponsaveis) {
      const responsavel = todosResponsaveis.find(r => r.id === idResp);
      
      if (!responsavel || !responsavel.email) {
        erros.push(`Responsável ${idResp} sem email cadastrado.`);
        continue;
      }
      
      // Filtra projetos e etapas do responsável
      let projetosDoResponsavel = todosProjetos.filter(p => 
        p.responsaveisIds && p.responsaveisIds.includes(idResp)
      );
      
      let etapasDoResponsavel = todasEtapas.filter(e => 
        e.responsaveisIds && e.responsaveisIds.includes(idResp)
      );
      
      // Se projetoId foi passado, filtra apenas esse projeto
      if (projetoId) {
        projetosDoResponsavel = projetosDoResponsavel.filter(p => p.id === projetoId);
        etapasDoResponsavel = etapasDoResponsavel.filter(e => e.projetoId === projetoId);
      }
      
      // Monta o corpo do email com prioridades
      const corpoEmail = montarCorpoEmailComPrioridades(
        responsavel, 
        projetosDoResponsavel, 
        etapasDoResponsavel,
        todosProjetos,
        mensagemAdicional,
        prioridadesProjeto
      );
      
      const assunto = assuntoPersonalizado || 
        `📋 Resumo de Projetos e Etapas - ${responsavel.nome}`;
      
      try {
        const opcoesEmail = {
          to: responsavel.email,
          subject: assunto,
          htmlBody: corpoEmail,
          name: 'Smart Meeting',
          replyTo: 'setorbiunichristus@gmail.com'
        };
        
        // Adicionar CC se houver
        if (emailsCcValidos.length > 0) {
          opcoesEmail.cc = emailsCcValidos.join(',');
        }
        
        MailApp.sendEmail(opcoesEmail);
        emailsEnviados++;
        Logger.log('Email enviado para: ' + responsavel.email + (emailsCcValidos.length > 0 ? ' CC: ' + emailsCcValidos.join(',') : ''));
      } catch (emailErro) {
        erros.push(`Erro ao enviar para ${responsavel.email}: ${emailErro.message}`);
        Logger.log('ERRO envio email: ' + emailErro.toString());
      }
    }
    
    if (emailsEnviados > 0) {
      return { 
        sucesso: true, 
        mensagem: `${emailsEnviados} email(s) enviado(s) com sucesso!` + 
                  (emailsCcValidos.length > 0 ? ` (CC: ${emailsCcValidos.length})` : ''),
        erros: erros.length > 0 ? erros : null
      };
    } else {
      return { 
        sucesso: false, 
        mensagem: 'Nenhum email foi enviado.',
        erros: erros
      };
    }
    
  } catch (e) {
    Logger.log('ERRO enviarEmailParaResponsaveisComCopia: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}


// ===================== SISTEMA DE PRIORIDADES DE PROJETOS =====================

/**
 * Obtém as prioridades gerais de todos os projetos
 * @returns {Object} Lista de prioridades de projetos
 */
function obterPrioridadesGeraisDeProjetos() {
  try {
    const aba = obterAba(NOME_ABA_PRIORIDADES);
    if (!aba || aba.getLastRow() <= 1) {
      return { sucesso: true, prioridades: [] };
    }
    
    const dados = aba.getDataRange().getValues();
    const prioridades = [];
    
    for (let i = 1; i < dados.length; i++) {
      // Filtrar apenas prioridades de projetos (tipoItem = 'projeto' e sem projetoReferencia ou com valor especial)
      if (dados[i][COLUNAS_PRIORIDADES.TIPO_ITEM] === 'projeto' && 
          (!dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA] || 
           dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA] === 'GERAL')) {
        prioridades.push({
          id: dados[i][COLUNAS_PRIORIDADES.ID],
          responsavelId: dados[i][COLUNAS_PRIORIDADES.RESPONSAVEL_ID],
          tipoItem: dados[i][COLUNAS_PRIORIDADES.TIPO_ITEM],
          itemId: dados[i][COLUNAS_PRIORIDADES.ITEM_ID],
          ordemPrioridade: dados[i][COLUNAS_PRIORIDADES.ORDEM_PRIORIDADE],
          projetoReferencia: dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA]
        });
      }
    }
    
    // Ordenar por ordem de prioridade
    prioridades.sort((a, b) => a.ordemPrioridade - b.ordemPrioridade);
    
    return { sucesso: true, prioridades: prioridades };
  } catch (e) {
    Logger.log('ERRO obterPrioridadesGeraisDeProjetos: ' + e.toString());
    return { sucesso: false, mensagem: e.message, prioridades: [] };
  }
}

/**
 * Salva as prioridades gerais de projetos
 * @param {Array} listaPrioridades - Lista de projetos com suas prioridades
 * @returns {Object} Resultado da operação
 */
function salvarPrioridadesGeraisDeProjetos(listaPrioridades) {
  try {
    const aba = obterAba(NOME_ABA_PRIORIDADES);
    const dados = aba.getDataRange().getValues();
    
    // Remover prioridades antigas de projetos gerais (de trás para frente)
    for (let i = dados.length - 1; i >= 1; i--) {
      if (dados[i][COLUNAS_PRIORIDADES.TIPO_ITEM] === 'projeto' && 
          (!dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA] || 
           dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA] === 'GERAL')) {
        aba.deleteRow(i + 1);
      }
    }
    
    // Inserir novas prioridades
    if (listaPrioridades && listaPrioridades.length > 0) {
      const novasLinhas = listaPrioridades.map((item, index) => {
        return [
          gerarId(),                        // ID
          item.responsavelId || '',         // RESPONSAVEL_ID
          'projeto',                        // TIPO_ITEM
          item.itemId,                      // ITEM_ID (ID do projeto)
          index + 1,                        // ORDEM_PRIORIDADE (1-based)
          'GERAL'                           // PROJETO_REFERENCIA (marcador especial)
        ];
      });
      
      // Adicionar todas as linhas de uma vez
      const ultimaLinha = aba.getLastRow();
      aba.getRange(ultimaLinha + 1, 1, novasLinhas.length, novasLinhas[0].length)
         .setValues(novasLinhas);
    }
    
    return { sucesso: true, mensagem: `${listaPrioridades.length} prioridade(s) de projeto(s) salva(s)!` };
  } catch (e) {
    Logger.log('ERRO salvarPrioridadesGeraisDeProjetos: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Obtém o número de prioridade de um projeto específico
 * @param {string} projetoId - ID do projeto
 * @returns {number|null} Número da prioridade ou null
 */
function obterNumeroPrioridadeProjeto(projetoId) {
  try {
    const resultado = obterPrioridadesGeraisDeProjetos();
    if (!resultado.sucesso) return null;
    
    const prioridade = resultado.prioridades.find(p => p.itemId === projetoId);
    return prioridade ? prioridade.ordemPrioridade : null;
  } catch (e) {
    Logger.log('ERRO obterNumeroPrioridadeProjeto: ' + e.toString());
    return null;
  }
}

/**
 * Obtém as prioridades de projetos para um responsável específico
 * @param {string} responsavelId - ID do responsável (ou 'TODOS' para geral)
 * @returns {Object} Lista de prioridades de projetos
 */
function obterPrioridadesProjetosPorResponsavel(responsavelId) {
  try {
    const aba = obterAba(NOME_ABA_PRIORIDADES);
    if (!aba || aba.getLastRow() <= 1) {
      return { sucesso: true, prioridades: [] };
    }
    
    const dados = aba.getDataRange().getValues();
    const prioridades = [];
    
    // Definir o filtro de responsável
    const filtroResp = responsavelId || 'TODOS';
    
    for (let i = 1; i < dados.length; i++) {
      const tipoItem = dados[i][COLUNAS_PRIORIDADES.TIPO_ITEM];
      const respId = dados[i][COLUNAS_PRIORIDADES.RESPONSAVEL_ID] || 'TODOS';
      
      // Filtrar por tipo 'projeto' e pelo responsável específico
      if (tipoItem === 'projeto' && respId === filtroResp) {
        prioridades.push({
          id: dados[i][COLUNAS_PRIORIDADES.ID],
          responsavelId: respId,
          tipoItem: tipoItem,
          itemId: dados[i][COLUNAS_PRIORIDADES.ITEM_ID],
          ordemPrioridade: dados[i][COLUNAS_PRIORIDADES.ORDEM_PRIORIDADE],
          projetoReferencia: dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA]
        });
      }
    }
    
    // Ordenar por ordem de prioridade
    prioridades.sort((a, b) => a.ordemPrioridade - b.ordemPrioridade);
    
    return { sucesso: true, prioridades: prioridades };
  } catch (e) {
    Logger.log('ERRO obterPrioridadesProjetosPorResponsavel: ' + e.toString());
    return { sucesso: false, mensagem: e.message, prioridades: [] };
  }
}

/**
 * Salva as prioridades de projetos para um responsável específico
 * @param {string} responsavelId - ID do responsável (ou 'TODOS' para geral)
 * @param {Array} listaPrioridades - Lista de projetos com suas prioridades
 * @returns {Object} Resultado da operação
 */
function salvarPrioridadesProjetosPorResponsavel(responsavelId, listaPrioridades) {

  try {
    const aba = obterAba(NOME_ABA_PRIORIDADES);
    const dados = aba.getDataRange().getValues();
    
    // Definir o filtro de responsável
    const filtroResp = responsavelId || 'TODOS';
    
    // Remover prioridades antigas deste responsável (de trás para frente)
    for (let i = dados.length - 1; i >= 1; i--) {
      const tipoItem = dados[i][COLUNAS_PRIORIDADES.TIPO_ITEM];
      const respId = dados[i][COLUNAS_PRIORIDADES.RESPONSAVEL_ID] || 'TODOS';
      
      if (tipoItem === 'projeto' && respId === filtroResp) {
        aba.deleteRow(i + 1);
      }
    }
    
    // Inserir novas prioridades
    if (listaPrioridades && listaPrioridades.length > 0) {
      const novasLinhas = listaPrioridades.map((item, index) => {
        return [
          gerarId(),                        // ID
          filtroResp,                       // RESPONSAVEL_ID
          'projeto',                        // TIPO_ITEM
          item.itemId,                      // ITEM_ID (ID do projeto)
          index + 1,                        // ORDEM_PRIORIDADE (1-based)
          'PRIORIZACAO'                     // PROJETO_REFERENCIA (marcador)
        ];
      });
      
      // Adicionar todas as linhas de uma vez
      const ultimaLinha = aba.getLastRow();
      aba.getRange(ultimaLinha + 1, 1, novasLinhas.length, novasLinhas[0].length)
         .setValues(novasLinhas);
    }
    
    return { 
      sucesso: true, 
      mensagem: `${listaPrioridades.length} prioridade(s) salva(s) para ${filtroResp === 'TODOS' ? 'Todos' : 'o responsável'}!` 
    };
  } catch (e) {
    Logger.log('ERRO salvarPrioridadesProjetosPorResponsavel: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function salvarValorPrioridadeProjeto(projetoId, valorPrioridade, dadosCalculo, token) {

  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    if (!_usuarioPodeAcessarProjetoPorDepartamento(auth.sessao, projetoId)) {
      return { sucesso: false, mensagem: 'Sem acesso ao projeto informado.' };
    }

    if (!projetoId) {
      return { sucesso: false, mensagem: 'ID do projeto é obrigatório' };
    }
    
    const aba = obterAba(NOME_ABA_PROJETOS);
    const dados = aba.getDataRange().getValues();
    
    dadosCalculo = dadosCalculo || {};
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID] === projetoId) {
        const linha = i + 1;
        const linhaAtualizada = dados[i].slice();
        linhaAtualizada[COLUNAS_PROJETOS.VALOR_PRIORIDADE] = valorPrioridade;
        if (dadosCalculo.gravidadeLabel) linhaAtualizada[COLUNAS_PROJETOS.GRAVIDADE] = dadosCalculo.gravidadeLabel;
        if (dadosCalculo.urgenciaLabel) linhaAtualizada[COLUNAS_PROJETOS.URGENCIA] = dadosCalculo.urgenciaLabel;
        if (dadosCalculo.paraQuemLabel) linhaAtualizada[COLUNAS_PROJETOS.PARA_QUEM] = dadosCalculo.paraQuemLabel;
        if (dadosCalculo.tipoProjetoLabel) linhaAtualizada[COLUNAS_PROJETOS.TIPO] = dadosCalculo.tipoProjetoLabel;
        if (dadosCalculo.esforcoLabel) linhaAtualizada[COLUNAS_PROJETOS.ESFORCO] = dadosCalculo.esforcoLabel;

        let categoriaPrioridade = 'Baixa';
        if (valorPrioridade >= 2102) {
          categoriaPrioridade = 'Alta';
        } else if (valorPrioridade >= 1078) {
          categoriaPrioridade = 'Média';
        }
        linhaAtualizada[COLUNAS_PROJETOS.PRIORIDADE] = categoriaPrioridade;
        aba.getRange(linha, 1, 1, linhaAtualizada.length).setValues([linhaAtualizada]);
        
        return { 
          sucesso: true, 
          mensagem: `Prioridade calculada: ${valorPrioridade} (${categoriaPrioridade})`,
          valorPrioridade: valorPrioridade,
          categoriaPrioridade: categoriaPrioridade
        };
      }
    }
    
    return { sucesso: false, mensagem: 'Projeto não encontrado' };
    
  } catch (e) {
    Logger.log('ERRO salvarValorPrioridadeProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function montarOpcoes(lista, pesoMaximo) {
  if (!lista || lista.length === 0) return [];
  const pesoMax = Math.max(pesoMaximo, lista.length);
  return lista.map((label, index) => ({
    valor: pesoMax - index,
    label: label.toString()
  })).filter(o => o.valor > 0);
}

// ===================== CONFIG PRIORIDADE =====================

const CONFIG_PRIORIDADE_DEFAULTS = {
  // Gravidade
  grav_critico:        5,
  grav_alto:           4,
  grav_medio:          3,
  // Urgência
  urg_imediata:        5,
  urg_muito_urgente:   4,
  urg_urgente:         3,
  urg_pouco_urgente:   2,
  urg_pode_esperar:    1,
  // Tipo de Projeto
  tipo_correcao:       5,
  tipo_nova_impl:      4,
  tipo_melhoria:       3,
  // Para Quem
  para_diretoria:      5,
  para_outros:         3,
  // Esforço
  esf_turno:           5,
  esf_dia:             4,
  esf_semana:          3,
  esf_mais_semana:     2,
  // Escala de prioridade (thresholds)
  escala_alta_min:     2102,
  escala_media_min:    1078
};

function _lerConfigPrioridadeTab_() {
  try {
    const ss  = SpreadsheetApp.openById(ID_PLANILHA);
    let aba   = ss.getSheetByName(NOME_ABA_CONFIG_PRIORIDADE);
    if (!aba) {
      aba = ss.insertSheet(NOME_ABA_CONFIG_PRIORIDADE);
      inicializarCabecalhoAba(aba, NOME_ABA_CONFIG_PRIORIDADE);
      return Object.assign({}, CONFIG_PRIORIDADE_DEFAULTS);
    }
    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return Object.assign({}, CONFIG_PRIORIDADE_DEFAULTS);
    const cfg = Object.assign({}, CONFIG_PRIORIDADE_DEFAULTS);
    for (let i = 1; i < dados.length; i++) {
      const chave = dados[i][0] ? dados[i][0].toString().trim() : '';
      const val   = dados[i][1];
      if (chave && chave in cfg && val !== '' && val !== null) {
        cfg[chave] = Number(val);
      }
    }
    return cfg;
  } catch (e) {
    Logger.log('ERRO _lerConfigPrioridadeTab_: ' + e.toString());
    return Object.assign({}, CONFIG_PRIORIDADE_DEFAULTS);
  }
}

function obterConfigPrioridade(token) {
  try {
    const sess = verificarSessao(token);
    if (!sess || !sess.valida) return { sucesso: false, mensagem: 'Sessão inválida.' };
    const cfg = _lerConfigPrioridadeTab_();
    return { sucesso: true, config: cfg };
  } catch (e) {
    Logger.log('ERRO obterConfigPrioridade: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function salvarConfigPrioridade(token, config) {
  try {
    const sess = verificarSessao(token);
    if (!sess || !sess.valida) return { sucesso: false, mensagem: 'Sessão inválida.' };
    if (sess.usuario.perfil !== 'admin') return { sucesso: false, mensagem: 'Apenas administradores podem alterar a configuração de prioridade.' };

    const ss  = SpreadsheetApp.openById(ID_PLANILHA);
    let aba   = ss.getSheetByName(NOME_ABA_CONFIG_PRIORIDADE);
    if (!aba) {
      aba = ss.insertSheet(NOME_ABA_CONFIG_PRIORIDADE);
      inicializarCabecalhoAba(aba, NOME_ABA_CONFIG_PRIORIDADE);
    }

    // Reconstrói a aba inteiramente para evitar linhas órfãs
    const linhas = [['Chave', 'Valor']];
    const chavesValidas = Object.keys(CONFIG_PRIORIDADE_DEFAULTS);
    chavesValidas.forEach(function(chave) {
      const val = (config && config[chave] !== undefined) ? Number(config[chave]) : CONFIG_PRIORIDADE_DEFAULTS[chave];
      linhas.push([chave, val]);
    });

    aba.clearContents();
    aba.getRange(1, 1, linhas.length, 2).setValues(linhas);
    aba.getRange(1, 1, 1, 2).setFontWeight('bold');

    return { sucesso: true, mensagem: 'Configuração salva com sucesso!' };
  } catch (e) {
    Logger.log('ERRO salvarConfigPrioridade: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function obterConfigCalculoPrioridade() {
  try {
    const cfg = _lerConfigPrioridadeTab_();

    return {
      gravidade: {
        titulo: 'Gravidade',
        descricao: 'O que acontece se eu não fizer?',
        opcoes: [
          { valor: cfg.grav_critico, label: 'Crítico - Não é possível cumprir as atividades' },
          { valor: cfg.grav_alto,    label: 'Alto - É possível cumprir parcialmente' },
          { valor: cfg.grav_medio,   label: 'Médio - É possível mas demora muito' }
        ]
      },
      urgencia: {
        titulo: 'Urgência',
        descricao: 'Para quando?',
        opcoes: [
          { valor: cfg.urg_imediata,       label: 'Imediata - Executar imediatamente' },
          { valor: cfg.urg_muito_urgente,  label: 'Muito urgente - Prazo curto (5 dias)' },
          { valor: cfg.urg_urgente,        label: 'Urgente - Curto prazo (10 dias)' },
          { valor: cfg.urg_pouco_urgente,  label: 'Pouco urgente - Mais de 10 dias' },
          { valor: cfg.urg_pode_esperar,   label: 'Pode esperar' }
        ]
      },
      tipoProjeto: {
        titulo: 'Tipo de Projeto',
        descricao: 'Natureza do projeto',
        opcoes: [
          { valor: cfg.tipo_correcao,  label: 'Correção' },
          { valor: cfg.tipo_nova_impl, label: 'Nova Implementação' },
          { valor: cfg.tipo_melhoria,  label: 'Melhoria' }
        ]
      },
      quemSolicita: {
        titulo: 'Para Quem',
        descricao: 'Quem solicitou',
        opcoes: [
          { valor: cfg.para_diretoria, label: 'Diretoria / Crítico' },
          { valor: cfg.para_outros,    label: 'Demais áreas' }
        ]
      },
      esforco: {
        titulo: 'Esforço',
        descricao: 'Tempo de desenvolvimento necessário',
        opcoes: [
          { valor: cfg.esf_turno,      label: '1 turno ou menos (4 horas)' },
          { valor: cfg.esf_dia,        label: '1 dia ou menos (8 horas)' },
          { valor: cfg.esf_semana,     label: 'uma semana (40h)' },
          { valor: cfg.esf_mais_semana,label: 'mais de uma semana (40h)' }
        ]
      },
      escala: {
        alta:  { minimo: cfg.escala_alta_min,  cor: '#dc2626', bgCor: 'rgba(220, 38, 38, 0.15)' },
        media: { minimo: cfg.escala_media_min, maximo: cfg.escala_alta_min - 1, cor: '#ca8a04', bgCor: 'rgba(202, 138, 4, 0.15)' },
        baixa: { maximo: cfg.escala_media_min - 1, cor: '#059669', bgCor: 'rgba(5, 150, 105, 0.15)' }
      }
    };
  } catch (e) {
    Logger.log('ERRO obterConfigCalculoPrioridade: ' + e.toString());
    return {
      gravidade:   { titulo: 'Gravidade',      descricao: 'O que acontece?',     opcoes: [{ valor: 5, label: 'Crítico' }, { valor: 4, label: 'Alto' }, { valor: 3, label: 'Médio' }] },
      urgencia:    { titulo: 'Urgência',        descricao: 'Para quando?',        opcoes: [{ valor: 5, label: 'Imediata' }, { valor: 4, label: 'Muito urgente' }, { valor: 3, label: 'Urgente' }, { valor: 2, label: 'Pouco urgente' }, { valor: 1, label: 'Pode esperar' }] },
      tipoProjeto: { titulo: 'Tipo de Projeto', descricao: 'Natureza',            opcoes: [{ valor: 5, label: 'Correção' }, { valor: 4, label: 'Nova Implementação' }, { valor: 3, label: 'Melhoria' }] },
      quemSolicita:{ titulo: 'Para Quem',       descricao: 'Solicitante',         opcoes: [{ valor: 5, label: 'Diretoria' }, { valor: 3, label: 'Demais áreas' }] },
      esforco:     { titulo: 'Esforço',         descricao: 'Tempo de desenvolvimento', opcoes: [{ valor: 5, label: '1 turno ou menos (4 horas)' }, { valor: 4, label: '1 dia ou menos (8 horas)' }, { valor: 3, label: 'uma semana (40h)' }, { valor: 2, label: 'mais de uma semana (40h)' }] },
      escala: {
        alta:  { minimo: 2102, cor: '#dc2626', bgCor: 'rgba(220, 38, 38, 0.15)' },
        media: { minimo: 1078, maximo: 2101,   cor: '#ca8a04', bgCor: 'rgba(202, 138, 4, 0.15)' },
        baixa: { maximo: 1077, cor: '#059669', bgCor: 'rgba(5, 150, 105, 0.15)' }
      }
    };
  }
}

// ===================== LIXEIRA DE PROJETOS =====================

/**
 * Move um projeto para a aba de lixeira (soft delete)
 * @param {string} projetoId - ID do projeto a excluir
 * @returns {Object} Resultado da operação
 */
function excluirProjeto(projetoId, token) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    if (!_usuarioPodeAcessarProjetoPorDepartamento(auth.sessao, projetoId)) {
      return { sucesso: false, mensagem: 'Sem acesso ao projeto informado.' };
    }

    const abaProjetos = obterAba(NOME_ABA_PROJETOS);
    const dados = abaProjetos.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID] === projetoId) {
        const linhaCompleta = dados[i];

        // Cria aba lixeira se não existir
        const abaLixeira = obterOuCriarAbaLixeira();

        // Adiciona data de exclusão no final da linha
        const linhaComData = [...linhaCompleta, new Date().toISOString()];
        abaLixeira.appendRow(linhaComData);

        // Remove da aba de projetos
        abaProjetos.deleteRow(i + 1);

        // Remove etapas associadas ao projeto
        const abaEtapas = obterAba(NOME_ABA_ETAPAS);
        const dadosEtapas = abaEtapas.getDataRange().getValues();
        const linhasParaExcluir = [];
        for (let j = 1; j < dadosEtapas.length; j++) {
          if (dadosEtapas[j][COLUNAS_ETAPAS.PROJETO_ID] === projetoId) {
            linhasParaExcluir.push(j + 1);
          }
        }
        for (let k = linhasParaExcluir.length - 1; k >= 0; k--) {
          abaEtapas.deleteRow(linhasParaExcluir[k]);
        }

        Logger.log('Projeto movido para lixeira: ' + projetoId);
        return { sucesso: true, mensagem: 'Projeto movido para a lixeira!' };
      }
    }

    return { sucesso: false, mensagem: 'Projeto não encontrado' };
  } catch (e) {
    Logger.log('ERRO excluirProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Lista todos os projetos na lixeira
 * @returns {Object} Lista de projetos excluídos
 */
function listarLixeira(token) {
  try {
    const sessao = token ? _obterSessao(token) : null;
    if (!sessao) return { sucesso: false, mensagem: 'Sessão inválida ou expirada.', projetos: [] };

    const depsFiltro = (sessao.perfil === 'admin') ? null : _obterDepsAtualizadosUsuario(sessao);
    const abaLixeira = obterOuCriarAbaLixeira();
    if (abaLixeira.getLastRow() <= 1) return { sucesso: true, projetos: [] };

    const dados = abaLixeira.getDataRange().getValues();
    const projetos = [];

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID]) {
        if (depsFiltro !== null && depsFiltro.length > 0) {
          const depProjetoRaw = (dados[i][COLUNAS_PROJETOS.DEPARTAMENTO_ID] || '').toString().trim();
          if (depProjetoRaw) {
            const depProjeto = _resolverIdsDepartamento([depProjetoRaw])[0] || depProjetoRaw;
            if (!depsFiltro.includes(depProjeto)) continue;
          }
          // Projetos sem departamento: visíveis por todos (retrocompat)
        }
        const respRaw = dados[i][COLUNAS_PROJETOS.RESPONSAVEIS_IDS];
        let respIds = [];
        if (respRaw) {
          try { respIds = JSON.parse(respRaw); } catch(e) { respIds = [respRaw.toString()]; }
        }
        projetos.push({
          id: dados[i][COLUNAS_PROJETOS.ID],
          nome: dados[i][COLUNAS_PROJETOS.NOME],
          descricao: dados[i][COLUNAS_PROJETOS.DESCRICAO],
          status: dados[i][COLUNAS_PROJETOS.STATUS],
          setor: dados[i][COLUNAS_PROJETOS.SETOR],
          prioridade: dados[i][COLUNAS_PROJETOS.PRIORIDADE],
          valorPrioridade: dados[i][COLUNAS_PROJETOS.VALOR_PRIORIDADE] || 0,
          responsaveisIds: respIds,
          // Data de exclusão está na última coluna (índice 17)
          dataExclusao: dados[i][17] || ''
        });
      }
    }

    return { sucesso: true, projetos: projetos };
  } catch (e) {
    Logger.log('ERRO listarLixeira: ' + e.toString());
    return { sucesso: false, mensagem: e.message, projetos: [] };
  }
}

/**
 * Restaura um projeto da lixeira de volta para projetos ativos
 * @param {string} projetoId - ID do projeto a restaurar
 * @returns {Object} Resultado da operação
 */
function restaurarDaLixeira(projetoId, token) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const abaLixeira = obterOuCriarAbaLixeira();
    const dados = abaLixeira.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID] === projetoId) {
        const depProjetoRaw = (dados[i][COLUNAS_PROJETOS.DEPARTAMENTO_ID] || '').toString().trim();
        const depProjeto = _resolverIdsDepartamento([depProjetoRaw])[0] || depProjetoRaw;
        const depsUsuario = _resolverIdsDepartamento(auth.sessao.departamentosIds || []);
        if (auth.sessao.perfil !== 'admin' && !depsUsuario.includes(depProjeto)) {
          return { sucesso: false, mensagem: 'Sem acesso ao departamento deste projeto.' };
        }

        // Pega apenas as colunas do projeto (sem a coluna de data de exclusão)
        const linhaOriginal = dados[i].slice(0, 20);

        // Restaura para aba de projetos
        const abaProjetos = obterAba(NOME_ABA_PROJETOS);
        abaProjetos.appendRow(linhaOriginal);

        // Remove da lixeira
        abaLixeira.deleteRow(i + 1);

        Logger.log('Projeto restaurado: ' + projetoId);
        return { sucesso: true, mensagem: 'Projeto restaurado com sucesso!' };
      }
    }

    return { sucesso: false, mensagem: 'Projeto não encontrado na lixeira' };
  } catch (e) {
    Logger.log('ERRO restaurarDaLixeira: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Exclui definitivamente um projeto da lixeira
 * @param {string} projetoId - ID do projeto
 * @returns {Object} Resultado da operação
 */
function excluirDefinitivamente(projetoId, token) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const abaLixeira = obterOuCriarAbaLixeira();
    const dados = abaLixeira.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID] === projetoId) {
        const depProjetoRaw = (dados[i][COLUNAS_PROJETOS.DEPARTAMENTO_ID] || '').toString().trim();
        const depProjeto = _resolverIdsDepartamento([depProjetoRaw])[0] || depProjetoRaw;
        const depsUsuario = _resolverIdsDepartamento(auth.sessao.departamentosIds || []);
        if (auth.sessao.perfil !== 'admin' && !depsUsuario.includes(depProjeto)) {
          return { sucesso: false, mensagem: 'Sem acesso ao departamento deste projeto.' };
        }
        abaLixeira.deleteRow(i + 1);
        Logger.log('Projeto excluído definitivamente: ' + projetoId);
        return { sucesso: true, mensagem: 'Projeto excluído permanentemente!' };
      }
    }

    return { sucesso: false, mensagem: 'Projeto não encontrado na lixeira' };
  } catch (e) {
    Logger.log('ERRO excluirDefinitivamente: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Obtém ou cria a aba de lixeira com cabeçalhos
 * @returns {Sheet} Aba da lixeira
 */
function obterOuCriarAbaLixeira() {
  const planilha = obterPlanilha();
  let aba = planilha.getSheetByName(NOME_ABA_LIXEIRA);
  if (!aba) {
    aba = planilha.insertSheet(NOME_ABA_LIXEIRA);
    // Cabeçalhos iguais aos de Projetos + coluna DataExclusao
    const cabecalho = ['ID','Nome','Descricao','Tipo','ParaQuem','Status','Prioridade',
                       'Link','Gravidade','Urgencia','Bloqueado','Setor','Pilar',
                       'ResponsaveisIds','ValorPrioridade','DataInicio','DataFim','DataExclusao'];
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
    aba.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold');
    aba.setTabColor('#dc2626'); // aba vermelha para identificação visual
  }
  return aba;
}

function listarEtapasPorProjetoOrdenadas(projetoId) {
  const etapas = listarEtapasPorProjeto(projetoId);

  try {
    const res = obterPrioridadesDoProjeto(projetoId);
    if (!res.sucesso || !res.prioridades || res.prioridades.length === 0) return etapas;

    const mapaOrdem = {};
    res.prioridades.forEach(p => {
      if (p.tipoItem === 'etapa') mapaOrdem[p.itemId] = p.ordemPrioridade;
    });

    return etapas.sort((a, b) => {
      const oA = mapaOrdem[a.id] !== undefined ? mapaOrdem[a.id] : 99999;
      const oB = mapaOrdem[b.id] !== undefined ? mapaOrdem[b.id] : 99999;
      return oA - oB;
    });
  } catch (e) {
    Logger.log('ERRO listarEtapasPorProjetoOrdenadas: ' + e.toString());
    return etapas;
  }
}

/**
 * Salva a ordem das atividades raiz de um projeto na aba Prioridades.
 * Reutiliza salvarPrioridadesDoProjeto() com tipoItem = 'etapa'.
 * @param {string} projetoId
 * @param {Array<string>} idsOrdenados - IDs das atividades raiz na nova ordem
 * @returns {Object} { sucesso, mensagem }
 */
function salvarOrdemAtividades(projetoId, idsOrdenados, token) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    if (!_usuarioPodeAcessarProjetoPorDepartamento(auth.sessao, projetoId)) {
      return { sucesso: false, mensagem: 'Sem acesso ao projeto informado.' };
    }

    if (!projetoId || !Array.isArray(idsOrdenados) || idsOrdenados.length === 0) {
      return { sucesso: false, mensagem: 'Dados inválidos para salvar ordem' };
    }

    const listaPrioridades = idsOrdenados.map((itemId, index) => ({
      responsavelId: '',
      tipoItem: 'etapa',
      itemId: itemId,
      ordemPrioridade: index + 1
    }));

    return salvarPrioridadesDoProjeto(projetoId, listaPrioridades);
  } catch (e) {
    Logger.log('ERRO salvarOrdemAtividades: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Calcula a duração de um projeto em dias úteis e horas,
 * descontando feriados da aba FERIADOS26.
 * Fórmula: (diasUteisCompletos × 9h) - (sabadosUteis × 1h)
 * Replica a lógica: DIATRABALHOTOTAL(...) * 9 - DIATRABALHOTOTAL.INTL(..., "1111011", ...) * 1
 *
 * @param {string} dataInicioStr - Data início (yyyy-MM-dd)
 * @param {string} dataFimStr    - Data fim (yyyy-MM-dd)
 * @returns {Object} { sucesso, diasUteis, horasTotais, diasExibicao, horasExibicao, textoFormatado }
 */
function calcularDuracaoProjetoHoras(dataInicioStr, dataFimStr) {
  try {
    if (!dataInicioStr || !dataFimStr) {
      return { sucesso: false, mensagem: 'Datas não informadas', textoFormatado: '' };
    }

    var inicio = new Date(dataInicioStr + 'T00:00:00');
    var fim    = new Date(dataFimStr + 'T00:00:00');

    if (isNaN(inicio.getTime()) || isNaN(fim.getTime())) {
      return { sucesso: false, mensagem: 'Datas inválidas', textoFormatado: '' };
    }

    if (fim < inicio) {
      return { sucesso: false, mensagem: 'Data fim anterior à data início', textoFormatado: '' };
    }

    // ── Carregar feriados da aba FERIADOS26 ──
    var ABA_FERIADOS = 'FERIADOS26';
    var COLUNA_FERIADOS = 0; // coluna A
    var feriados = [];

    try {
      var planilha = SpreadsheetApp.getActiveSpreadsheet();
      var abaFeriados = planilha.getSheetByName(ABA_FERIADOS);
      if (abaFeriados && abaFeriados.getLastRow() > 1) {
        var dadosFeriados = abaFeriados.getRange(2, 1, abaFeriados.getLastRow() - 1, 1).getValues();
        dadosFeriados.forEach(function(linha) {
          if (linha[COLUNA_FERIADOS] instanceof Date) {
            // Normalizar para meia-noite
            var d = new Date(linha[COLUNA_FERIADOS]);
            d.setHours(0, 0, 0, 0);
            feriados.push(d.getTime());
          }
        });
      }
    } catch (e) {
      Logger.log('Aviso: Aba FERIADOS26 não encontrada ou erro ao ler: ' + e.toString());
      // Continua sem feriados
    }

    // ── Função auxiliar: verifica se uma data é feriado ──
    function ehFeriado(data) {
      var ts = new Date(data.getFullYear(), data.getMonth(), data.getDate()).getTime();
      return feriados.indexOf(ts) !== -1;
    }

    // ── Contar dias úteis (seg-sex, excluindo feriados) ──
    // Equivale a DIATRABALHOTOTAL(inicio, fim, feriados)
    var diasUteisCompletos = 0;

    // ── Contar sábados úteis (não feriados) ──
    // Equivale a DIATRABALHOTOTAL.INTL(inicio, fim, "1111011", feriados)
    // "1111011" = seg/ter/qua/qui são não-trabalho(1), sex=não-trabalho(0→trabalho? NÃO)
    // Na verdade "1111011": posições = seg(1)ter(1)qua(1)qui(1)sex(0)sab(1)dom(1)
    // Isso significa: APENAS sexta é dia de trabalho... mas na fórmula original
    // está subtraindo esse valor × 1h.
    // Revisando a fórmula: DIATRABALHOTOTAL * 9 - DIATRABALHOTOTAL.INTL(..., "1111011") * 1
    // "1111011" marca como NÃO útil: seg,ter,qua,qui,sab,dom → só SEXTA é útil
    // Isso conta apenas as sextas-feiras úteis (não feriadas) no período
    // E subtrai 1h por cada sexta → sextas valem 8h em vez de 9h
    var sextasUteis = 0;

    var dataAtual = new Date(inicio);
    dataAtual.setHours(0, 0, 0, 0);
    var dataFimNorm = new Date(fim);
    dataFimNorm.setHours(0, 0, 0, 0);

    while (dataAtual <= dataFimNorm) {
      var diaSemana = dataAtual.getDay(); // 0=dom, 1=seg, ..., 5=sex, 6=sab

      if (diaSemana >= 1 && diaSemana <= 5 && !ehFeriado(dataAtual)) {
        // É dia útil (seg a sex, não feriado)
        diasUteisCompletos++;

        if (diaSemana === 5) {
          // É sexta-feira útil
          sextasUteis++;
        }
      }

      dataAtual.setDate(dataAtual.getDate() + 1);
    }

    // ── Cálculo final: dias úteis × 9h - sextas × 1h ──
    var horasTotais = (diasUteisCompletos * 9) - (sextasUteis * 1);

    // Converter para dias e horas
    var diasExibicao  = Math.floor(horasTotais / 9); // dias de 9h
    var horasExibicao = horasTotais % 9;

    // Texto formatado
    var partes = [];
    if (diasExibicao > 0) {
      partes.push(diasExibicao + ' dia' + (diasExibicao !== 1 ? 's' : ''));
    }
    if (horasExibicao > 0) {
      partes.push(horasExibicao + ' hora' + (horasExibicao !== 1 ? 's' : ''));
    }
    var textoFormatado = partes.length > 0 ? partes.join(' e ') : '0 horas';

    return {
      sucesso: true,
      diasUteis: diasUteisCompletos,
      sextasUteis: sextasUteis,
      horasTotais: horasTotais,
      diasExibicao: diasExibicao,
      horasExibicao: horasExibicao,
      textoFormatado: textoFormatado
    };

  } catch (e) {
    Logger.log('ERRO calcularDuracaoProjetoHoras: ' + e.toString());
    return { sucesso: false, mensagem: e.message, textoFormatado: '' };
  }
}

/**
 * Salva a preferência de tema do usuário no servidor
 * @param {string} tema - 'cafe', 'azul' ou 'preto'
 */
function salvarPreferenciaTema(tema) {
  try {
    var temasValidos = ['cafe', 'azul', 'preto'];
    var temaFinal = temasValidos.indexOf(tema) >= 0 ? tema : 'cafe';
    var props = PropertiesService.getUserProperties();
    props.setProperty('smartmeeting_tema', temaFinal);
    return { sucesso: true };
  } catch (e) {
    Logger.log('ERRO salvarPreferenciaTema: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Obtém a preferência de tema do usuário
 * @returns {string} 'cafe', 'azul' ou 'preto'
 */
function obterPreferenciaTema() {
  try {
    var props = PropertiesService.getUserProperties();
    var tema = props.getProperty('smartmeeting_tema') || 'cafe';
    var temasValidos = ['cafe', 'azul', 'preto'];
    return temasValidos.indexOf(tema) >= 0 ? tema : 'cafe';
  } catch (e) {
    Logger.log('ERRO obterPreferenciaTema: ' + e.toString());
    return 'cafe';
  }
}

/**
 * Instala o trigger semanal para envio do resumo ao gestor.
 * Execute esta função UMA VEZ pelo editor do Apps Script.
 * Ela remove triggers antigos do mesmo tipo antes de criar um novo.
 */
function instalarTriggerResumoSemanal() {
  // Remove triggers anteriores desta função para evitar duplicatas
  removerTriggerResumoSemanal();

  ScriptApp.newTrigger('enviarResumoSemanalParaGestor')
    .timeBased()
    .onWeekDay(CONFIG_RESUMO_SEMANAL.DIA_SEMANA)
    .atHour(CONFIG_RESUMO_SEMANAL.HORA_DISPARO)
    .create();

  Logger.log('✅ Trigger do resumo semanal instalado: sexta-feira às '
    + CONFIG_RESUMO_SEMANAL.HORA_DISPARO + ':00');
}

/**
 * Remove todos os triggers associados à função enviarResumoSemanalParaGestor.
 * Use para desativar o envio automático.
 */
function removerTriggerResumoSemanal() {
  var todosOsTriggers = ScriptApp.getProjectTriggers();
  todosOsTriggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'enviarResumoSemanalParaGestor') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Envia o resumo semanal consolidado para o gestor.
 * Chamada automaticamente pelo trigger semanal.
 * Pode ser executada manualmente para testes.
 *
 * Fluxo:
 *  1. Filtra projetos pelos status definidos em CONFIG_RESUMO_SEMANAL
 *  2. Agrupa projetos por responsável
 *  3. Monta HTML do email
 *  4. Envia via MailApp
 *
 * @return {Object} { sucesso: boolean, mensagem: string }
 */
function enviarResumoSemanalParaGestor() {
  try {
    // 1. Busca projetos e filtra pelos status ativos
    var todosProjetos = listarProjetosCompletos();
    var projetosAtivos = todosProjetos.filter(function(projeto) {
      return CONFIG_RESUMO_SEMANAL.STATUS_INCLUIDOS.indexOf(projeto.status) !== -1;
    });

    // Se não houver projetos ativos, não envia email
    if (projetosAtivos.length === 0) {
      Logger.log('ℹ️ Nenhum projeto em andamento. Email do resumo semanal não enviado.');
      return { sucesso: true, mensagem: 'Sem projetos ativos — email não enviado.' };
    }

    // 2. Busca responsáveis e todas as etapas (atividades)
    var responsaveis = listarResponsaveisCompletos();
    var todasEtapas = obterTodasEtapas();

    // 3. Agrupa projetos por responsável
    var projetosPorResponsavel = {};
    responsaveis.forEach(function(resp) {
      projetosPorResponsavel[resp.id] = [];
    });

    projetosAtivos.forEach(function(projeto) {
      var idsResponsaveis = projeto.responsaveisIds || [];
      idsResponsaveis.forEach(function(idResp) {
        if (projetosPorResponsavel[idResp]) {
          projetosPorResponsavel[idResp].push(projeto);
        }
      });
    });

    // 4. Monta o HTML e envia
    var corpoHtmlEmail = montarEmailResumoSemanal(
      projetosPorResponsavel,
      responsaveis,
      todasEtapas,
      projetosAtivos
    );

var dataEnvio = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
MailApp.sendEmail({
  to: CONFIG_RESUMO_SEMANAL.EMAIL_GESTOR,
  cc:"napa13@christus.com.br,napa02@christus.com.br,michelle.furuya@gmail.com",
  subject: CONFIG_RESUMO_SEMANAL.ASSUNTO + ' · ' + dataEnvio,
  htmlBody: corpoHtmlEmail,
  name: 'Smart Meeting',
  replyTo: 'setorbiunichristus@gmail.com'
});

    Logger.log('✅ Resumo semanal enviado para ' + CONFIG_RESUMO_SEMANAL.EMAIL_GESTOR);
    return { sucesso: true, mensagem: 'Email enviado com sucesso.' };

  } catch (erro) {
    Logger.log('❌ ERRO em enviarResumoSemanalParaGestor: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}


function montarEmailResumoSemanal(projetosPorResponsavel, responsaveis, todasEtapas, projetosAtivos) {

  var urlSistema = '';
  try {
    // getUrl() pode retornar a URL /dev (ambiente de teste).
    // Substituímos por /exec para garantir que o gestor acesse a versão publicada.
    urlSistema = ScriptApp.getService().getUrl()
      .replace('/dev', '/exec') + '?pagina=projeto';
  } catch (e) { }

  var dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // ── Apenas projetos Em Andamento ──
  var emAndamento = projetosAtivos.filter(function(p) { return p.status === 'Em Andamento'; });

  // ── Agrupar projetos por responsável ──
  var mapaRespProjetos = {};
  responsaveis.forEach(function(r) {
    mapaRespProjetos[r.id] = { nome: r.nome, projetos: [] };
  });

emAndamento.forEach(function(proj) {
  // Formatar data início
  var dataInicio = '';
  if (proj.dataInicio) {
    var p = String(proj.dataInicio).split('-');
    dataInicio = p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : proj.dataInicio;
  }

  // Categoria de prioridade
  var valorPrio = proj.valorPrioridade || 0;
  var labelPrio = valorPrio >= 2102 ? '🔴 Alta — ' + valorPrio
                : valorPrio >= 1078 ? '🟡 Média — ' + valorPrio
                : valorPrio > 0     ? '🟢 Baixa — ' + valorPrio
                : '';

  // ── Calcular progresso das atividades do projeto ──
  var etapasDoProjeto = todasEtapas.filter(function(e) {
    return String(e.projetoId) === String(proj.id);
  });
  var totalEtapas      = etapasDoProjeto.length;
  var concluidasEtapas = etapasDoProjeto.filter(function(e) {
    return e.status === 'Concluída';
  }).length;
  var percentual = totalEtapas > 0 ? Math.round((concluidasEtapas / totalEtapas) * 100) : 0;

  // Cor da barra conforme percentual
  var corBarra = percentual === 100 ? '#5a7247'   // verde (concluído)
               : percentual >= 50  ? '#b45309'   // âmbar (andamento)
               :                     '#9a3412';  // laranja-escuro (início)

  (proj.responsaveisIds || []).forEach(function(id) {
    if (mapaRespProjetos[id]) {
      mapaRespProjetos[id].projetos.push({
        nome:             proj.nome        || '',
        descricao:        proj.descricao   || '',
        dataInicio:       dataInicio,
        labelPrio:        labelPrio,
        // ── progresso ──
        totalEtapas:      totalEtapas,
        concluidasEtapas: concluidasEtapas,
        percentual:       percentual,
        corBarra:         corBarra
      });
    }
  });
});

  // ── Montar linhas por responsável ──
  var linhasResponsaveis = '';
  responsaveis.forEach(function(r) {
    var entry = mapaRespProjetos[r.id];
    if (!entry || entry.projetos.length === 0) return;

    // Nome do responsável
    linhasResponsaveis +=
      '<p style="margin:0 0 8px 0;font-size:0.9rem;color:#3d2b1f;">' +
        '<strong>' + entry.nome + '</strong>' +
      '</p>';

    // Cards dos projetos
entry.projetos.forEach(function(proj) {
  linhasResponsaveis +=
    '<div style="background:#fef9f0;border-left:3px solid #b45309;padding:10px 14px;margin-bottom:7px;border-radius:4px;">' +

      // Linha: nome + data início (lado a lado)
      '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
        '<td style="vertical-align:top;">' +
          '<div style="font-size:0.85rem;font-weight:600;color:#451a03;">' + proj.nome + '</div>' +
          (proj.descricao
            ? '<div style="font-size:0.78rem;color:#7a6555;margin-top:3px;">' + proj.descricao + '</div>'
            : '') +
        '</td>' +
        '<td style="text-align:right;vertical-align:top;white-space:nowrap;padding-left:12px;">' +
          (proj.dataInicio
            ? '<div style="font-size:0.75rem;color:#9a8a78;">📅 ' + proj.dataInicio + '</div>'
            : '') +
          (proj.labelPrio
            ? '<div style="font-size:0.72rem;color:#7a6555;margin-top:3px;">' + proj.labelPrio + '</div>'
            : '') +
        '</td>' +
      '</tr></table>' +

      // ── NOVO: Barra de progresso das atividades ──
      (proj.totalEtapas > 0
        ? '<div style="margin-top:10px;">' +
            // Cabeçalho: rótulo + contagem
            '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">' +
              '<span style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;color:#9a8a78;">📋 Progresso das atividades</span>' +
              '<span style="font-size:0.72rem;color:#7a6555;font-weight:600;">' +
                proj.concluidasEtapas + ' / ' + proj.totalEtapas +
                ' &nbsp;(' + proj.percentual + '%)' +
              '</span>' +
            '</div>' +
            // Trilha da barra
            '<div style="width:100%;height:7px;background:#e7e5e4;border-radius:4px;overflow:hidden;">' +
              '<div style="width:' + proj.percentual + '%;height:100%;background:' + proj.corBarra + ';border-radius:4px;transition:width 0.3s;"></div>' +
            '</div>' +
          '</div>'
        : '<div style="margin-top:8px;font-size:0.7rem;color:#c0b0a0;font-style:italic;">Sem atividades cadastradas</div>') +

    '</div>';
});

    linhasResponsaveis += '<div style="height:10px;"></div>';
  });

  if (!linhasResponsaveis) {
    linhasResponsaveis =
      '<p style="color:#9a8a78;font-style:italic;font-size:0.85rem;">Nenhum projeto em andamento no momento.</p>';
  }

  var h = '';
  h += '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>';
  h += '<body style="margin:0;padding:24px;font-family:\'Segoe UI\',Arial,sans-serif;background:#ffffff;">';
  h += '<div style="max-width:520px;">';

  // Link no topo
  if (urlSistema) {
    h += '<p style="margin:0 0 20px 0;">';
    h += '<a href="' + urlSistema + '" style="color:#b45309;font-size:0.82rem;font-weight:600;text-decoration:none;">→ Abrir Smart Meeting</a>';
    h += '</p>';
  }

  // Saudação
  h += '<p style="margin:0 0 6px 0;font-size:0.92rem;color:#3d2b1f;line-height:1.7;">';
  h += 'Olá, <strong>' + CONFIG_RESUMO_SEMANAL.NOME_GESTOR + '</strong>!<br>';
  h += 'Projetos <strong>Em Andamento</strong> · ' + dataHoje;
  h += '</p>';

  h += '<hr style="border:none;border-top:1px solid #e8dfd2;margin:16px 0 20px 0;">';

  // Conteúdo
  h += linhasResponsaveis;

  h += '</div></body></html>';
  return h;
}

// ========================= MÓDULO: DOCUMENTOS DO DRIVE =========================

function _extrairIdPastaDrive(link) {
  if (!link) return null;
  var m = link.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

function _obterLinkProjeto(projetoId) {
  var aba = obterAba(NOME_ABA_PROJETOS);
  var valores = aba.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (valores[i][COLUNAS_PROJETOS.ID] == projetoId) {
      return { link: (valores[i][COLUNAS_PROJETOS.LINK] || '').toString().trim(), linha: i + 1 };
    }
  }
  return null;
}

function criarPastaProjetoDrive(token, projetoId) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const aba = obterAba(NOME_ABA_PROJETOS);
    const valores = aba.getDataRange().getValues();
    let linha = -1, nomeProjeto = '', linkExistente = '';
    for (let i = 1; i < valores.length; i++) {
      if (valores[i][COLUNAS_PROJETOS.ID] == projetoId) {
        linha = i + 1;
        nomeProjeto = valores[i][COLUNAS_PROJETOS.NOME] || 'Projeto ' + projetoId;
        linkExistente = (valores[i][COLUNAS_PROJETOS.LINK] || '').toString().trim();
        break;
      }
    }
    if (linha === -1) return { sucesso: false, mensagem: 'Projeto não encontrado.' };
    if (linkExistente) return { sucesso: true, link: linkExistente };

    const pastaPrincipal = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const novaSubPasta = pastaPrincipal.createFolder(nomeProjeto);
    const linkNovo = 'https://drive.google.com/drive/folders/' + novaSubPasta.getId();
    aba.getRange(linha, COLUNAS_PROJETOS.LINK + 1).setValue(linkNovo);
    return { sucesso: true, link: linkNovo };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function listarDocumentosDrive(token, projetoId) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const info = _obterLinkProjeto(projetoId);
    if (!info || !info.link) return { sucesso: true, arquivos: [], temPasta: false };

    const folderId = _extrairIdPastaDrive(info.link);
    if (!folderId) return { sucesso: true, arquivos: [], temPasta: false };

    let pasta;
    try { pasta = DriveApp.getFolderById(folderId); }
    catch(e) { return { sucesso: true, arquivos: [], temPasta: false }; }

    const arquivos = [];
    const it = pasta.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      arquivos.push({
        id: f.getId(),
        nome: f.getName(),
        mimeType: f.getMimeType(),
        tamanho: f.getSize(),
        dataModificacao: Utilities.formatDate(f.getLastUpdated(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
        urlVisualizacao: f.getUrl()
      });
    }
    arquivos.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
    return { sucesso: true, arquivos: arquivos, temPasta: true, link: info.link };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function uploadDocumentoDrive(token, projetoId, nomeArquivo, conteudoBase64, mimeType) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const info = _obterLinkProjeto(projetoId);
    if (!info || !info.link) return { sucesso: false, mensagem: 'Este projeto não tem pasta no Drive.' };

    const folderId = _extrairIdPastaDrive(info.link);
    if (!folderId) return { sucesso: false, mensagem: 'Link de pasta inválido.' };

    const pasta = DriveApp.getFolderById(folderId);
    const dados = conteudoBase64.split(',')[1] || conteudoBase64;
    const blob = Utilities.newBlob(Utilities.base64Decode(dados), mimeType, nomeArquivo);
    const arquivo = pasta.createFile(blob);
    return {
      sucesso: true,
      arquivo: {
        id: arquivo.getId(),
        nome: arquivo.getName(),
        mimeType: arquivo.getMimeType(),
        tamanho: arquivo.getSize(),
        dataModificacao: Utilities.formatDate(arquivo.getLastUpdated(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
        urlVisualizacao: arquivo.getUrl()
      }
    };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

// ── Upload em chunks (arquivos grandes) ───────────────────────────────────────

function iniciarUploadChunked(token, projetoId) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const info = _obterLinkProjeto(projetoId);
    if (!info || !info.link) return { sucesso: false, mensagem: 'Este projeto não tem pasta no Drive.' };

    const folderId = _extrairIdPastaDrive(info.link);
    if (!folderId) return { sucesso: false, mensagem: 'Link de pasta inválido.' };

    const pasta = DriveApp.getFolderById(folderId);
    const tempNome = '__upload_temp_' + Utilities.getUuid();
    const tempPasta = pasta.createFolder(tempNome);

    return { sucesso: true, uploadId: tempPasta.getId(), folderId: folderId };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function enviarChunkUpload(token, uploadId, chunkIndex, dadosBase64) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const tempPasta = DriveApp.getFolderById(uploadId);
    const bytes = Utilities.base64Decode(dadosBase64);
    const nome = 'chunk_' + ('000000' + chunkIndex).slice(-6);
    tempPasta.createFile(Utilities.newBlob(bytes, 'application/octet-stream', nome));

    return { sucesso: true };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function finalizarUploadChunked(token, uploadId, totalChunks, folderId, nomeArquivo, mimeType) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const tempPasta = DriveApp.getFolderById(uploadId);

    // Coleta chunks em ordem
    var fileMap = {};
    var iter = tempPasta.getFiles();
    while (iter.hasNext()) {
      var f = iter.next();
      fileMap[f.getName()] = f.getBlob().getBytes();
    }

    // Concatena bytes
    var totalLen = 0;
    var arrays = [];
    for (var i = 0; i < totalChunks; i++) {
      var key = 'chunk_' + ('000000' + i).slice(-6);
      if (!fileMap[key]) return { sucesso: false, mensagem: 'Parte ' + i + ' não encontrada.' };
      var arr = fileMap[key];
      arrays.push(arr);
      totalLen += arr.length;
    }
    var combined = new Array(totalLen);
    var offset = 0;
    for (var a = 0; a < arrays.length; a++) {
      for (var b = 0; b < arrays[a].length; b++) {
        combined[offset++] = arrays[a][b];
      }
    }

    // Salva arquivo final e apaga pasta temp
    var pasta = DriveApp.getFolderById(folderId);
    var blob = Utilities.newBlob(combined, mimeType, nomeArquivo);
    var arquivo = pasta.createFile(blob);
    tempPasta.setTrashed(true);

    return {
      sucesso: true,
      arquivo: {
        id: arquivo.getId(),
        nome: arquivo.getName(),
        mimeType: arquivo.getMimeType(),
        tamanho: arquivo.getSize(),
        dataModificacao: Utilities.formatDate(arquivo.getLastUpdated(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
        urlVisualizacao: arquivo.getUrl()
      }
    };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function renomearDocumentoDrive(token, fileId, novoNome) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    DriveApp.getFileById(fileId).setName(novoNome);
    return { sucesso: true };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function removerDocumentoDrive(token, fileId) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    DriveApp.getFileById(fileId).setTrashed(true);
    return { sucesso: true };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function obterConteudoDocumentoDrive(token, fileId) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const arquivo = DriveApp.getFileById(fileId);
    const mimeType = arquivo.getMimeType();

    if (mimeType.startsWith('text/') || mimeType === 'application/json') {
      return { sucesso: true, tipo: 'texto', conteudo: arquivo.getBlob().getDataAsString('utf-8') };
    } else if (mimeType.startsWith('image/')) {
      const base64 = Utilities.base64Encode(arquivo.getBlob().getBytes());
      return { sucesso: true, tipo: 'imagem', conteudo: 'data:' + mimeType + ';base64,' + base64 };
    } else {
      return { sucesso: true, tipo: 'iframe', url: 'https://drive.google.com/file/d/' + fileId + '/preview' };
    }
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function salvarTextoNaDrive(token, projetoId, nome, conteudo, mimeTypeParam) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };

    const info = _obterLinkProjeto(projetoId);
    if (!info || !info.link) return { sucesso: false, mensagem: 'Este projeto não tem pasta no Drive.' };

    const folderId = _extrairIdPastaDrive(info.link);
    if (!folderId) return { sucesso: false, mensagem: 'Link de pasta inválido.' };

    const mime = mimeTypeParam || 'text/html';
    const ext  = mime === 'text/html' ? '.html' : '.txt';
    const nomeArquivo = nome.match(/\.(html?|txt)$/i) ? nome : nome + ext;
    const pasta = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(conteudo, mime, nomeArquivo);
    const arquivo = pasta.createFile(blob);
    return {
      sucesso: true,
      arquivo: {
        id: arquivo.getId(),
        nome: arquivo.getName(),
        mimeType: arquivo.getMimeType(),
        tamanho: arquivo.getSize(),
        dataModificacao: Utilities.formatDate(arquivo.getLastUpdated(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
        urlVisualizacao: arquivo.getUrl()
      }
    };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function atualizarTextoNaDrive(token, fileId, novoConteudo) {
  try {
    const auth = _obterSessaoEdicaoProjetos(token);
    if (!auth.ok) return { sucesso: false, mensagem: auth.mensagem };
    DriveApp.getFileById(fileId).setContent(novoConteudo);
    return { sucesso: true };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}
