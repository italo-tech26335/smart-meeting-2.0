const PREFIXO_ARQUIVO_RELATORIO = 'Relatorio_Identificacoes_';
const EXTENSAO_RELATORIO = '.md';
const MARCADOR_DEP_RELATORIO = '_DEP-'; // separador de departamento no nome do arquivo

/**
 * Monta os slugs de departamento do usuário para uso no filtro de relatórios.
 * Retorna null para admin (vê tudo) ou sessão inexistente.
 */
function _obterSlugsDepsUsuarioParaRelatorio(token) {
  if (!token) return null;
  try {
    var sessao = _obterSessao(token);
    if (!sessao) return null;
    if (sessao.perfil === 'admin') return null; // admin vê tudo

    var depsIds = _obterDepsAtualizadosUsuario(sessao);
    if (!depsIds || depsIds.length === 0) return null; // sem restrição (retrocompat)

    // Busca nomes na aba de departamentos
    var resDeps = listarDepartamentos(null);
    var todosDeps = (resDeps && resDeps.sucesso) ? (resDeps.departamentos || []) : [];

    // Para cada ID do usuário, encontra o nome e gera o slug
    var slugs = [];
    depsIds.forEach(function(id) {
      var dep = todosDeps.find(function(d) { return d.id === id || d.nome === id; });
      var nome = dep ? dep.nome : id;
      var slug = _slugificarNomeDep(nome);
      if (slug) slugs.push(slug);
    });
    return slugs.length > 0 ? slugs : null;
  } catch (e) {
    Logger.log('AVISO _obterSlugsDepsUsuarioParaRelatorio: ' + e.toString());
    return null;
  }
}

function listarRelatoriosIdentificacao(token) {
  try {
    // Filtro por departamento (null = sem restrição)
    var slugsUsuario = _obterSlugsDepsUsuarioParaRelatorio(token);

    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const arquivos = pasta.getFiles();
    const relatorios = [];

    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nomeArquivo = arquivo.getName();

      // Filtrar apenas arquivos de relatório
      if (nomeArquivo.startsWith(PREFIXO_ARQUIVO_RELATORIO) && nomeArquivo.endsWith(EXTENSAO_RELATORIO)) {

        // ─── Filtro de departamento ───────────────────────────────────────
        if (slugsUsuario !== null) {
          var temDep = nomeArquivo.indexOf(MARCADOR_DEP_RELATORIO) !== -1;
          if (temDep) {
            // Arquivo tem marcação de departamento: verifica se pertence ao usuário
            var pertence = slugsUsuario.some(function(slug) {
              return nomeArquivo.indexOf(MARCADOR_DEP_RELATORIO + slug + '_') !== -1;
            });
            if (!pertence) continue; // pula arquivo de outro departamento
          }
          // Arquivo sem marcação de departamento (legado): visível a todos
        }
        // ─────────────────────────────────────────────────────────────────

        // Extrair data do nome do arquivo
        // Formato novo: Relatorio_Identificacoes_DEP-SLUG_YYYYMMDD_HHMMSS.md
        // Formato antigo: Relatorio_Identificacoes_YYYYMMDD_HHMMSS.md
        var nomeBase = nomeArquivo
          .replace(PREFIXO_ARQUIVO_RELATORIO, '')
          .replace(EXTENSAO_RELATORIO, '');

        // Remove o segmento _DEP-SLUG se existir
        var matchDep = nomeBase.match(/^DEP-[^_]+_/);
        if (matchDep) nomeBase = nomeBase.substring(matchDep[0].length);

        const partes = nomeBase.split('_');
        let dataFormatada = '-';
        let dataCriacaoMs = 0;

        // Data real do Drive (para ordenação)
        const dt = arquivo.getDateCreated();
        dataCriacaoMs = dt ? dt.getTime() : 0;

        // Data do nome do arquivo (para exibição)
        if (partes.length >= 2) {
          const dataStr = partes[0]; // YYYYMMDD
          const horaStr = partes[1]; // HHMMSS

          if (dataStr.length === 8 && horaStr.length === 6) {
            const ano = dataStr.substring(0, 4);
            const mes = dataStr.substring(4, 6);
            const dia = dataStr.substring(6, 8);
            const hora = horaStr.substring(0, 2);
            const minuto = horaStr.substring(2, 4);
            dataFormatada = `${dia}/${mes}/${ano} ${hora}:${minuto}`;
          }
        }

        relatorios.push({
          id: arquivo.getId(),
          nome: nomeArquivo,
          dataCriacaoMs: dataCriacaoMs,
          dataFormatada: dataFormatada,
          tamanho: arquivo.getSize(),
          tamanhoFormatado: formatarTamanhoArquivoBackend(arquivo.getSize()),
          link: arquivo.getUrl()
        });
      }
    }

    // Ordenar por data (mais recente primeiro)
    relatorios.sort((a, b) => (b.dataCriacaoMs || 0) - (a.dataCriacaoMs || 0));

    return {
      sucesso: true,
      relatorios: relatorios,
      total: relatorios.length
    };

  } catch (erro) {
    Logger.log('ERRO listarRelatoriosIdentificacao: ' + erro.toString());
    return {
      sucesso: false,
      mensagem: erro.message,
      relatorios: []
    };
  }
}

function inserirProjetoDaIdentificacao(dadosProjeto) {
  try {
    exigirPermissaoEdicao();

    const aba = obterAba(NOME_ABA_PROJETOS);

    const novoId = (dadosProjeto.id && dadosProjeto.id.trim() !== '')
      ? dadosProjeto.id.trim()
      : gerarId();

    // Monta a linha respeitando exatamente COLUNAS_PROJETOS (índices 0 a 16)
    const linha = [
      novoId,                              // 0  → ID
      dadosProjeto.nome         || '',     // 1  → NOME
      dadosProjeto.descricao    || '',     // 2  → DESCRICAO
      dadosProjeto.tipo         || '',     // 3  → TIPO
      dadosProjeto.paraQuem     || '',     // 4  → PARA_QUEM
      dadosProjeto.status       || 'A Fazer', // 5 → STATUS
      dadosProjeto.prioridade   || '',     // 6  → PRIORIDADE
      dadosProjeto.link         || '',     // 7  → LINK
      dadosProjeto.gravidade    || '',     // 8  → GRAVIDADE
      dadosProjeto.urgencia     || '',     // 9  → URGENCIA
      dadosProjeto.esforco      || '',     // 10 → ESFORCO
      dadosProjeto.setor        || '',     // 11 → SETOR
      dadosProjeto.pilar        || '',     // 12 → PILAR
      dadosProjeto.responsaveisIds || '',  // 13 → RESPONSAVEIS_IDS
      dadosProjeto.valorPrioridade || '', // 14 → VALOR_PRIORIDADE
      dadosProjeto.dataInicio   || '',     // 15 → DATA_INICIO
      dadosProjeto.dataFim      || ''      // 16 → DATA_FIM
    ];

    aba.appendRow(linha);

    Logger.log('Projeto inserido — ID: ' + novoId + ' | Nome: ' + dadosProjeto.nome
      + ' | Gravidade: ' + dadosProjeto.gravidade
      + ' | Urgência: '  + dadosProjeto.urgencia
      + ' | Esforço: '   + dadosProjeto.esforco
      + ' | Valor Prio: '+ dadosProjeto.valorPrioridade
      + ' | Início: '    + dadosProjeto.dataInicio
      + ' | Fim: '       + dadosProjeto.dataFim);

    return {
      sucesso: true,
      mensagem: 'Projeto inserido com sucesso!',
      id: novoId,
      nome: dadosProjeto.nome
    };

  } catch (erro) {
    Logger.log('ERRO inserirProjetoDaIdentificacao: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

/**
 * Formata tamanho do arquivo em bytes para formato legível
 */
function formatarTamanhoArquivoBackend(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1024 / 1024).toFixed(2) + ' MB';
}

/** =====================================================================
 *                    FUNÇÕES DE LEITURA E PARSING
 * ===================================================================== */

/**
 * Obtém o conteúdo de um relatório e faz o parsing das identificações
 * @param {string} arquivoId - ID do arquivo no Drive
 * @returns {Object} - Dados parseados do relatório
 */
/**
 * Salva preferências de personalização do relatório para o usuário atual.
 * Armazenamento permanente via PropertiesService.getUserProperties().
 */
function salvarPrefsRelatorio(token, prefs) {
  try {
    const sessao = _obterSessao(token);
    if (!sessao) return { sucesso: false, mensagem: 'Sessão inválida.' };
    PropertiesService.getUserProperties().setProperty('mrel_prefs_v1', JSON.stringify(prefs));
    return { sucesso: true };
  } catch (e) {
    Logger.log('ERRO salvarPrefsRelatorio: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Carrega preferências de personalização do relatório do usuário atual.
 */
function carregarPrefsRelatorio(token) {
  try {
    const sessao = _obterSessao(token);
    if (!sessao) return { sucesso: false, mensagem: 'Sessão inválida.' };
    const raw = PropertiesService.getUserProperties().getProperty('mrel_prefs_v1');
    return { sucesso: true, prefs: raw ? JSON.parse(raw) : null };
  } catch (e) {
    Logger.log('ERRO carregarPrefsRelatorio: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Retorna todas as etapas para uso no preview e envio do relatório.
 */
function carregarEtapasParaRelatorio(token) {
  try {
    const sessao = _obterSessao(token);
    if (!sessao) return { sucesso: false, etapas: [] };
    const etapas = obterTodasEtapas();
    return { sucesso: true, etapas: etapas };
  } catch (e) {
    Logger.log('ERRO carregarEtapasParaRelatorio: ' + e.toString());
    return { sucesso: false, etapas: [] };
  }
}

function obterConteudoRelatorio(arquivoId) {
  try {
    const arquivo = DriveApp.getFileById(arquivoId);
    const conteudo = arquivo.getBlob().getDataAsString();
    
    Logger.log('=== INÍCIO DO PARSING DO RELATÓRIO ===');
    Logger.log('Tamanho do conteúdo: ' + conteudo.length + ' caracteres');
    
    // Fazer parsing do conteúdo markdown
    const dadosParseados = parsearRelatorioIdentificacao(conteudo);
    
    Logger.log('=== RESULTADO DO PARSING ===');
    Logger.log('Setores encontrados: ' + dadosParseados.setores.length);
    Logger.log('Projetos encontrados: ' + dadosParseados.projetos.length);
    Logger.log('Etapas encontradas: ' + dadosParseados.etapas.length);
    
    return {
      sucesso: true,
      arquivoId: arquivoId,
      nomeArquivo: arquivo.getName(),
      conteudoOriginal: conteudo,
      dadosParseados: dadosParseados
    };
    
  } catch (erro) {
    Logger.log('ERRO obterConteudoRelatorio: ' + erro.toString());
    return {
      sucesso: false,
      mensagem: erro.message
    };
  }
}

/**
 * Faz o parsing do conteúdo markdown do relatório
 * Extrai setores, projetos e etapas identificados
 * 
 * VERSÃO ROBUSTA - Compatível com variações do formato da IA
 */
function parsearRelatorioIdentificacao(conteudoMarkdown) {
  const resultado = {
    informacoesGerais: {},
    setores: [],
    projetos: [],
    etapas: [],
    setoresExcluidos: [],
    resumo: {
      totalSetores: 0,
      totalProjetos: 0,
      totalEtapas: 0,
      novosSetores: 0,
      novosProjetos: 0,
      novasEtapas: 0
    }
  };
  
  try {
    // Normalizar quebras de linha
    const conteudo = conteudoMarkdown.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    
    // Extrair informações gerais
    resultado.informacoesGerais = extrairInformacoesGerais(conteudo);
    
    // Extrair setores
    resultado.setores = extrairSetoresDoRelatorioV2(conteudo);
    Logger.log('Setores extraídos: ' + resultado.setores.length);
    
    // Extrair projetos
    resultado.projetos = extrairProjetosDoRelatorioV2(conteudo);
    Logger.log('Projetos extraídos: ' + resultado.projetos.length);
    
    // Extrair etapas
    resultado.etapas = extrairEtapasDoRelatorioV2(conteudo);
    Logger.log('Etapas extraídas: ' + resultado.etapas.length);
    
    // Extrair setores excluídos
    resultado.setoresExcluidos = extrairSetoresExcluidosV2(conteudo);
    
    // Calcular resumo
    resultado.resumo.totalSetores = resultado.setores.length;
    resultado.resumo.totalProjetos = resultado.projetos.length;
    resultado.resumo.totalEtapas = resultado.etapas.length;
    resultado.resumo.novosSetores = resultado.setores.filter(s => s.ehNovo).length;
    resultado.resumo.novosProjetos = resultado.projetos.filter(p => p.ehNovo).length;
    resultado.resumo.novasEtapas = resultado.etapas.filter(e => e.ehNovo).length;
    
  } catch (erro) {
    Logger.log('ERRO parsearRelatorioIdentificacao: ' + erro.toString());
    Logger.log('Stack: ' + erro.stack);
  }
  
  return resultado;
}

/**
 * Divide uma seção em itens individuais (cada ### é um item)
 */
function dividirPorItens(secao) {
  const itens = [];
  
  // Dividir por ### (início de item)
  const partes = secao.split(/\n(?=###\s)/);
  
  for (const parte of partes) {
    const parteTrimmada = parte.trim();
    if (parteTrimmada && parteTrimmada.startsWith('###')) {
      itens.push(parteTrimmada);
    }
  }
  
  return itens;
}

/**
 * Parseia um item individual (setor, projeto ou etapa)
 * Extrai nome, campos da tabela e justificativa
 */
function parsearItemGenerico(textoItem, tipo) {
  const item = {
    nomeExtraido: '',
    ehNovo: false,
    idExistente: null,
    idSugerido: null,
    campos: {},
    justificativa: ''
  };
  
  try {
    // Extrair nome do item (### Nome do Item)
    const matchNome = textoItem.match(/^###\s*(.+?)(?:\n|$)/);
    if (matchNome) {
      item.nomeExtraido = matchNome[1].trim();
    }
    
    // Extrair campos da tabela Markdown
    // Procurar todas as linhas que parecem ser linhas de tabela com dados
    const linhas = textoItem.split('\n');
    let dentroTabela = false;
    let cabecalhoEncontrado = false;
    
    for (const linha of linhas) {
      const linhaTrim = linha.trim();
      
      // Detectar cabeçalho da tabela
      if (linhaTrim.match(/^\|\s*Campo\s*\|\s*Valor\s*\|/i)) {
        dentroTabela = true;
        cabecalhoEncontrado = true;
        continue;
      }
      
      // Pular linha de separação (|---|---|)
      if (linhaTrim.match(/^\|[\s-:|]+\|$/)) {
        continue;
      }
      
      // Extrair dados de linha da tabela
      if (dentroTabela && linhaTrim.startsWith('|') && linhaTrim.endsWith('|')) {
        const colunas = linhaTrim.split('|').map(c => c.trim()).filter(c => c !== '');
        
        if (colunas.length >= 2) {
          const campo = colunas[0].toUpperCase().replace(/\s+/g, '_');
          const valor = colunas[1];
          
          // Ignorar campos vazios ou de cabeçalho
          if (campo && valor && campo !== 'CAMPO' && campo !== '---') {
            item.campos[campo] = converterValorCampo(campo, valor);
            
            // Verificar se é novo ou existente pelo campo ID
            if (campo === 'ID') {
              const valorUpper = valor.toUpperCase();

              if (valorUpper.includes('EXISTENTE')) {
                // Item existente — extrair ID real
                item.ehNovo = false;
                const matchId = valor.match(/ID:\s*([^)\s,]+)/i);
                if (matchId) item.idExistente = matchId[1].trim();

              } else if (
                // ✅ CORRIGIDO: reconhece atv_, etp_, proj_, setor_ SEM prefixo "NOVO_"
                valor.match(/^(atv_|etp_|proj_|setor_)\d+$/i)
              ) {
                item.ehNovo = true;
                item.idSugerido = valor;

              } else {
                // Sem marcação clara → considera novo como fallback
                item.ehNovo = true;
                item.idSugerido = valor;
              }
            }
          }
        }
      }
      
      // Fim da tabela quando encontrar linha sem pipe no início
      if (dentroTabela && !linhaTrim.startsWith('|') && linhaTrim.length > 0) {
        dentroTabela = false;
      }
    }
    
    // Extrair justificativa
    const matchJustificativa = textoItem.match(/\*\*Justificativa[:\s]*\*\*\s*(.+?)(?=\n\n|\n###|$)/is);
    if (matchJustificativa) {
      item.justificativa = matchJustificativa[1].trim();
    }
    
    // Se não encontrou nome no campo NOME, usar o extraído do título
    if (!item.campos.NOME && item.nomeExtraido) {
      item.campos.NOME = item.nomeExtraido;
    }
    
  } catch (erro) {
    Logger.log('ERRO parsearItemGenerico (' + tipo + '): ' + erro.toString());
  }
  
  return item;
}

/**
 * Extrai setores mencionados mas não incluídos
 */
function extrairSetoresExcluidosV2(conteudo) {
  const excluidos = [];
  
  try {
    // Encontrar seção
    const padroes = [
      /##\s*❌\s*SETORES MENCIONADOS[^\n]*\n([\s\S]*?)(?=\n##|$)/i,
      /##\s*SETORES MENCIONADOS MAS NÃO INCLUÍDOS[^\n]*\n([\s\S]*?)(?=\n##|$)/i
    ];
    
    let secao = '';
    for (const padrao of padroes) {
      const match = conteudo.match(padrao);
      if (match) {
        secao = match[1];
        break;
      }
    }
    
    if (!secao) return excluidos;
    
    // Extrair da tabela
    const linhas = secao.split('\n');
    for (const linha of linhas) {
      const linhaTrim = linha.trim();
      
      if (linhaTrim.startsWith('|') && linhaTrim.endsWith('|')) {
        const colunas = linhaTrim.split('|').map(c => c.trim()).filter(c => c !== '');
        
        if (colunas.length >= 2) {
          const setor = colunas[0];
          const motivo = colunas[1];
          
          // Ignorar cabeçalho e separador
          if (setor && !setor.toLowerCase().includes('setor mencionado') && 
              !setor.includes('---') && setor !== '[Nome]') {
            excluidos.push({ setor, motivo });
          }
        }
      }
    }
    
  } catch (erro) {
    Logger.log('ERRO extrairSetoresExcluidosV2: ' + erro.toString());
  }
  
  return excluidos;
}

/** =====================================================================
 *                    FUNÇÕES DE CONTEXTO ATUAL
 * ===================================================================== */

/**
 * Obtém o estado atual dos dados para pré-visualização
 * Retorna setores, projetos e etapas existentes
 */
function obterContextoAtualParaPreview() {
  try {
    // Carregar setores
    const abaSetores = obterAba(NOME_ABA_SETORES);
    const dadosSetores = abaSetores.getDataRange().getValues();
    const setores = [];
    
    for (let i = 1; i < dadosSetores.length; i++) {
      if (dadosSetores[i][COLUNAS_SETORES.ID]) {
        setores.push({
          id: dadosSetores[i][COLUNAS_SETORES.ID],
          nome: dadosSetores[i][COLUNAS_SETORES.NOME],
          descricao: dadosSetores[i][COLUNAS_SETORES.DESCRICAO] || ''
        });
      }
    }
    
    // Carregar projetos
    const abaProjetos = obterAba(NOME_ABA_PROJETOS);
    const dadosProjetos = abaProjetos.getDataRange().getValues();
    const projetos = [];
    
    for (let i = 1; i < dadosProjetos.length; i++) {
      if (dadosProjetos[i][COLUNAS_PROJETOS.ID]) {
        projetos.push({
          id: dadosProjetos[i][COLUNAS_PROJETOS.ID],
          nome: dadosProjetos[i][COLUNAS_PROJETOS.NOME],
          descricao: dadosProjetos[i][COLUNAS_PROJETOS.DESCRICAO] || '',
          status: dadosProjetos[i][COLUNAS_PROJETOS.STATUS] || 'Ativo',
          setor: dadosProjetos[i][COLUNAS_PROJETOS.SETOR] || '',
          prioridade: dadosProjetos[i][COLUNAS_PROJETOS.PRIORIDADE] || 'Média'
        });
      }
    }
    
    // Carregar etapas
    const abaEtapas = obterAba(NOME_ABA_ETAPAS);
    const dadosEtapas = abaEtapas.getDataRange().getValues();
    const etapas = [];
    
    for (let i = 1; i < dadosEtapas.length; i++) {
      if (dadosEtapas[i][COLUNAS_ETAPAS.ID]) {
        etapas.push({
          id: dadosEtapas[i][COLUNAS_ETAPAS.ID],
          projetoId: dadosEtapas[i][COLUNAS_ETAPAS.PROJETO_ID],
          nome: dadosEtapas[i][COLUNAS_ETAPAS.NOME],
          descricao: dadosEtapas[i][COLUNAS_ETAPAS.DESCRICAO] || '',
          status: dadosEtapas[i][COLUNAS_ETAPAS.STATUS] || STATUS_ETAPAS.A_FAZER,
          etapaPaiId: dadosEtapas[i][COLUNAS_ETAPAS.ETAPA_PAI_ID] || ''
        });
      }
    }
    
    // Carregar responsáveis
    const abaResp = obterAba(NOME_ABA_RESPONSAVEIS);
    const dadosResp = abaResp.getDataRange().getValues();
    const responsaveis = [];
    
    for (let i = 1; i < dadosResp.length; i++) {
      if (dadosResp[i][COLUNAS_RESPONSAVEIS.ID]) {
        responsaveis.push({
          id: dadosResp[i][COLUNAS_RESPONSAVEIS.ID],
          nome: dadosResp[i][COLUNAS_RESPONSAVEIS.NOME],
          email: dadosResp[i][COLUNAS_RESPONSAVEIS.EMAIL] || ''
        });
      }
    }
    
    return {
      sucesso: true,
      setores: setores,
      projetos: projetos,
      etapas: etapas,
      responsaveis: responsaveis
    };
    
  } catch (erro) {
    Logger.log('ERRO obterContextoAtualParaPreview: ' + erro.toString());
    return {
      sucesso: false,
      mensagem: erro.message
    };
  }
}

/** =====================================================================
 *                    FUNÇÕES DE INSERÇÃO
 * ===================================================================== */

/**
 * Insere um novo setor na planilha
 * @param {Object} dadosSetor - Dados do setor a inserir
 */
function inserirSetorDaIdentificacao(dadosSetor) {
  try {
    exigirPermissaoEdicao();
    
    const aba = obterAba(NOME_ABA_SETORES);
    
    // CORREÇÃO: Usar ID sugerido se disponível, senão gerar novo
    const novoId = dadosSetor.id && dadosSetor.id.trim() !== '' 
      ? dadosSetor.id.trim() 
      : gerarId();
    
    const linha = [
      novoId,
      dadosSetor.nome || '',
      dadosSetor.descricao || '',
      dadosSetor.responsaveisIds || ''
    ];
    
    aba.appendRow(linha);
    
    return {
      sucesso: true,
      mensagem: 'Setor inserido com sucesso!',
      id: novoId,
      nome: dadosSetor.nome
    };
    
  } catch (erro) {
    Logger.log('ERRO inserirSetorDaIdentificacao: ' + erro.toString());
    return {
      sucesso: false,
      mensagem: erro.message
    };
  }
}

/**
 * Converte valor de campo para o tipo apropriado
 */
function converterValorCampo(nomeCampo, valor) {
  if (!valor || valor === 'N/A') return '';
  
  const valorTrim = valor.toString().trim();
  
  // Campos numéricos
  if (nomeCampo === 'OFFSET_X' || nomeCampo === 'OFFSET_Y' || 
      nomeCampo === 'POS_X' || nomeCampo === 'POS_Y') {
    // Remover pontos de milhar e converter vírgula para ponto
    const numeroLimpo = valorTrim.replace(/\./g, '').replace(',', '.');
    const numero = parseFloat(numeroLimpo);
    return isNaN(numero) ? 0 : numero;
  }
  
  // Campos booleanos
  if (nomeCampo === 'BLOQUEADO') {
    return valorTrim.toLowerCase() === 'true' || valorTrim === '1';
  }
  
  // Campos JSON (arrays ou objetos)
  if (nomeCampo === 'PENDENCIAS') {
    // Se já parece JSON, retorna como está
    if (valorTrim.startsWith('[') || valorTrim.startsWith('{')) {
      return valorTrim;
    }
    // Se estiver vazio, retorna array vazio
    if (valorTrim === '' || valorTrim === 'N/A') {
      return '[]';
    }
  }
  
  // Campos de texto que podem ser N/A
  if (valorTrim === 'N/A') {
    return '';
  }
  
  return valorTrim;
}

/**
 * Obtém listas de opções para dropdowns do modal de edição
 * @returns {Object} Listas de responsáveis, setores, status e prioridades
 */
function obterOpcoesParaDropdowns() {
  try {
    // Carregar responsáveis
    const abaResp = obterAba(NOME_ABA_RESPONSAVEIS);
    const dadosResp = abaResp.getDataRange().getValues();
    const responsaveis = [];
    
    for (let i = 1; i < dadosResp.length; i++) {
      if (dadosResp[i][COLUNAS_RESPONSAVEIS.ID]) {
        responsaveis.push({
          id: dadosResp[i][COLUNAS_RESPONSAVEIS.ID],
          nome: dadosResp[i][COLUNAS_RESPONSAVEIS.NOME],
          email: dadosResp[i][COLUNAS_RESPONSAVEIS.EMAIL] || ''
        });
      }
    }
    
    // Carregar setores
    const abaSetores = obterAba(NOME_ABA_SETORES);
    const dadosSetores = abaSetores.getDataRange().getValues();
    const setores = [];
    
    for (let i = 1; i < dadosSetores.length; i++) {
      if (dadosSetores[i][COLUNAS_SETORES.ID]) {
        setores.push({
          id: dadosSetores[i][COLUNAS_SETORES.ID],
          nome: dadosSetores[i][COLUNAS_SETORES.NOME]
        });
      }
    }
    
    // Opções fixas de status
    const statusEtapas = ['A Fazer', 'Em Andamento', 'Bloqueada', 'Concluída'];
    const statusProjetos = ['Ativo', 'Pausado', 'Concluído', 'Cancelado'];
    
    // Opções fixas de prioridade
    const prioridades = ['Alta', 'Média', 'Baixa'];
    
    // Opções fixas de urgência
    const urgencias = ['alta', 'media', 'baixa'];
    
    return {
      sucesso: true,
      responsaveis: responsaveis,
      setores: setores,
      statusEtapas: statusEtapas,
      statusProjetos: statusProjetos,
      prioridades: prioridades,
      urgencias: urgencias
    };
    
  } catch (erro) {
    Logger.log('ERRO obterOpcoesParaDropdowns: ' + erro.toString());
    return {
      sucesso: false,
      mensagem: erro.message
    };
  }
}

function inserirEtapaDaIdentificacao(dadosEtapa) {
  try {
    exigirPermissaoEdicao();

    const aba = obterAba(NOME_ABA_ETAPAS);

    const novoId = (dadosEtapa.id && dadosEtapa.id.trim() !== '')
      ? dadosEtapa.id.trim()
      : gerarId();

    // Linha com exatamente 6 colunas — alinhada com COLUNAS_ETAPAS (0–5)
    const linha = [
      novoId,                                      // 0 → ID
      dadosEtapa.projetoId      || '',             // 1 → PROJETO_ID
      dadosEtapa.responsaveisIds|| '',             // 2 → RESPONSAVEIS_IDS
      dadosEtapa.nome           || '',             // 3 → NOME
      dadosEtapa.oQueFazer      || '',             // 4 → O_QUE_FAZER
      dadosEtapa.status         || STATUS_ETAPAS.A_FAZER // 5 → STATUS
    ];

    aba.appendRow(linha);

    Logger.log('Etapa inserida — ID: ' + novoId
      + ' | ProjetoId: ' + dadosEtapa.projetoId
      + ' | Nome: '      + dadosEtapa.nome
      + ' | Status: '    + dadosEtapa.status);

    return {
      sucesso: true,
      mensagem: 'Atividade inserida com sucesso!',
      id: novoId,
      nome: dadosEtapa.nome
    };

  } catch (erro) {
    Logger.log('ERRO inserirEtapaDaIdentificacao: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

/**
 * Insere múltiplos itens de uma vez (batch)
 */
function inserirItensEmLote(itens) {
  const resultados = {
    sucesso: true,
    inseridos: 0,
    erros: [],
    detalhes: []
  };
  
  try {
    for (const item of itens) {
      let resultado;
      
      switch (item.tipo) {
        case 'setor':
          resultado = inserirSetorDaIdentificacao(item.dados);
          break;
        case 'projeto':
          resultado = inserirProjetoDaIdentificacao(item.dados);
          break;
        case 'etapa':
          resultado = inserirEtapaDaIdentificacao(item.dados);
          break;
        default:
          resultado = { sucesso: false, mensagem: 'Tipo desconhecido: ' + item.tipo };
      }
      
      if (resultado.sucesso) {
        resultados.inseridos++;
        resultados.detalhes.push({
          tipo: item.tipo,
          nome: item.dados.nome,
          id: resultado.id,
          sucesso: true
        });
      } else {
        resultados.erros.push({
          tipo: item.tipo,
          nome: item.dados.nome,
          erro: resultado.mensagem
        });
        resultados.detalhes.push({
          tipo: item.tipo,
          nome: item.dados.nome,
          sucesso: false,
          erro: resultado.mensagem
        });
      }
    }
    
    if (resultados.erros.length > 0) {
      resultados.sucesso = resultados.inseridos > 0;
      resultados.mensagem = `${resultados.inseridos} item(ns) inserido(s), ${resultados.erros.length} erro(s)`;
    } else {
      resultados.mensagem = `${resultados.inseridos} item(ns) inserido(s) com sucesso!`;
    }
    
  } catch (erro) {
    Logger.log('ERRO inserirItensEmLote: ' + erro.toString());
    resultados.sucesso = false;
    resultados.mensagem = erro.message;
  }
  
  return resultados;
}

/**
 * Exclui um relatório do Drive
 */
function excluirRelatorioIdentificacao(arquivoId) {
  try {
    exigirPermissaoEdicao();
    
    const arquivo = DriveApp.getFileById(arquivoId);
    const nomeArquivo = arquivo.getName();
    
    arquivo.setTrashed(true);
    
    return {
      sucesso: true,
      mensagem: 'Relatório movido para lixeira',
      nomeArquivo: nomeArquivo
    };
    
  } catch (erro) {
    Logger.log('ERRO excluirRelatorioIdentificacao: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

/**
 * Busca ID de um projeto pelo nome do setor
 */
function buscarProjetoIdPorNomeSetor(nomeSetor, nomeProjeto) {
  try {
    const aba = obterAba(NOME_ABA_PROJETOS);
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      const setorProjeto = (dados[i][COLUNAS_PROJETOS.SETOR] || '').toString().toLowerCase();
      const nomeProjAtual = (dados[i][COLUNAS_PROJETOS.NOME] || '').toString().toLowerCase();
      
      if (setorProjeto.includes(nomeSetor.toLowerCase()) || 
          nomeProjAtual.includes(nomeProjeto.toLowerCase())) {
        return {
          sucesso: true,
          projetoId: dados[i][COLUNAS_PROJETOS.ID],
          nomeProjeto: dados[i][COLUNAS_PROJETOS.NOME]
        };
      }
    }
    
    return { sucesso: false, mensagem: 'Projeto não encontrado' };
    
  } catch (erro) {
    Logger.log('ERRO buscarProjetoIdPorNomeSetor: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

/** =====================================================================
 *                    FUNÇÕES DE DEBUG
 * ===================================================================== */

function debugListarRelatorios() {
  Logger.log('========================================');
  Logger.log('   DEBUG DE RELATÓRIOS');
  Logger.log('========================================');
  
  try {
    Logger.log('ID_PASTA_DRIVE_REUNIOES: ' + ID_PASTA_DRIVE_REUNIOES);
    
    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    Logger.log('Pasta: ' + pasta.getName());
    
    const arquivos = pasta.getFiles();
    let count = 0;
    
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nome = arquivo.getName();
      count++;
      
      Logger.log(count + '. ' + nome);
      Logger.log('   Começa com prefixo: ' + nome.startsWith(PREFIXO_ARQUIVO_RELATORIO));
      Logger.log('   Termina com .md: ' + nome.endsWith('.md'));
    }
    
    Logger.log('Total: ' + count + ' arquivos');
    
  } catch (e) {
    Logger.log('ERRO: ' + e.toString());
  }
}

function testarParsingRelatorio(arquivoId) {
  Logger.log('========================================');
  Logger.log('   TESTE DE PARSING');
  Logger.log('========================================');
  
  try {
    const resultado = obterConteudoRelatorio(arquivoId);
    
    if (!resultado.sucesso) {
      Logger.log('ERRO: ' + resultado.mensagem);
      return;
    }
    
    const dados = resultado.dadosParseados;
    
    Logger.log('--- Informações Gerais ---');
    Logger.log(JSON.stringify(dados.informacoesGerais, null, 2));
    
    Logger.log('--- Setores (' + dados.setores.length + ') ---');
    for (const setor of dados.setores) {
      Logger.log('  - ' + setor.nomeExtraido + ' (novo: ' + setor.ehNovo + ')');
      Logger.log('    Campos: ' + JSON.stringify(setor.campos));
    }
    
    Logger.log('--- Projetos (' + dados.projetos.length + ') ---');
    for (const projeto of dados.projetos) {
      Logger.log('  - ' + projeto.nomeExtraido + ' (novo: ' + projeto.ehNovo + ')');
      Logger.log('    Campos: ' + JSON.stringify(projeto.campos));
    }
    
    Logger.log('--- Etapas (' + dados.etapas.length + ') ---');
    for (const etapa of dados.etapas) {
      Logger.log('  - ' + etapa.nomeExtraido + ' (novo: ' + etapa.ehNovo + ')');
      Logger.log('    Campos: ' + JSON.stringify(etapa.campos));
    }
    
    Logger.log('--- Resumo ---');
    Logger.log(JSON.stringify(dados.resumo, null, 2));
    
  } catch (e) {
    Logger.log('ERRO: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
  }
}






// ============================================================
//   FUNÇÕES DE ATUALIZAÇÃO — ADICIONAR AO Relatorios.gs
//   Estas funções encontram a linha existente pelo ID e 
//   atualizam os valores IN-PLACE (sem criar nova linha)
// ============================================================

/**
 * Atualiza um PROJETO existente na planilha pelo ID.
 * Busca a linha onde o ID corresponde e atualiza apenas os campos informados.
 */
function atualizarProjetoDaIdentificacao(idExistente, dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName('Projetos');
    if (!aba) return { sucesso: false, mensagem: 'Aba "Projetos" não encontrada' };

    var dataRange = aba.getDataRange();
    var valores = dataRange.getValues();
    var linhaEncontrada = -1;

    // Coluna ID é a coluna 0 (COLUNAS_PROJETOS.ID = 0)
    for (var i = 1; i < valores.length; i++) { // começa em 1 para pular cabeçalho
      if (String(valores[i][COLUNAS_PROJETOS.ID]).trim() === String(idExistente).trim()) {
        linhaEncontrada = i + 1; // +1 porque getRange é 1-based
        break;
      }
    }

    if (linhaEncontrada === -1) {
      return { sucesso: false, mensagem: 'Projeto com ID "' + idExistente + '" não encontrado na planilha' };
    }

    // Atualiza em memória e grava a linha inteira em 1 operação.
    const indice = linhaEncontrada - 1;
    const linhaAtualizada = valores[indice].slice();
    const mapa = {
      nome: COLUNAS_PROJETOS.NOME,
      descricao: COLUNAS_PROJETOS.DESCRICAO,
      tipo: COLUNAS_PROJETOS.TIPO,
      paraQuem: COLUNAS_PROJETOS.PARA_QUEM,
      status: COLUNAS_PROJETOS.STATUS,
      prioridade: COLUNAS_PROJETOS.PRIORIDADE,
      link: COLUNAS_PROJETOS.LINK,
      gravidade: COLUNAS_PROJETOS.GRAVIDADE,
      urgencia: COLUNAS_PROJETOS.URGENCIA,
      bloqueado: COLUNAS_PROJETOS.BLOQUEADO,
      setor: COLUNAS_PROJETOS.SETOR,
      pilar: COLUNAS_PROJETOS.PILAR,
      responsaveisIds: COLUNAS_PROJETOS.RESPONSAVEIS_IDS,
      valorPrioridade: COLUNAS_PROJETOS.VALOR_PRIORIDADE
    };

    Object.keys(mapa).forEach(function(campo) {
      if (dados[campo] !== undefined && dados[campo] !== null) {
        linhaAtualizada[mapa[campo]] = dados[campo];
      }
    });

    aba.getRange(linhaEncontrada, 1, 1, linhaAtualizada.length).setValues([linhaAtualizada]);

    return { sucesso: true, mensagem: 'Projeto "' + idExistente + '" atualizado com sucesso (linha ' + linhaEncontrada + ')' };

  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao atualizar projeto: ' + e.message };
  }
}

function atualizarEtapaDaIdentificacao(idExistente, dados) {
  try {
    var aba = obterAba(NOME_ABA_ETAPAS);
    if (!aba) return { sucesso: false, mensagem: 'Aba de atividades não encontrada' };

    var valores = aba.getDataRange().getValues();
    var linhaEncontrada = -1;

    for (var i = 1; i < valores.length; i++) {
      if (String(valores[i][COLUNAS_ETAPAS.ID]).trim() === String(idExistente).trim()) {
        linhaEncontrada = i + 1; // +1: getRange é 1-based
        break;
      }
    }

    if (linhaEncontrada === -1) {
      return { sucesso: false, mensagem: 'Atividade com ID "' + idExistente + '" não encontrada' };
    }

    const indice = linhaEncontrada - 1;
    const linhaAtualizada = valores[indice].slice();
    const mapa = {
      projetoId: COLUNAS_ETAPAS.PROJETO_ID,
      responsaveisIds: COLUNAS_ETAPAS.RESPONSAVEIS_IDS,
      nome: COLUNAS_ETAPAS.NOME,
      oQueFazer: COLUNAS_ETAPAS.O_QUE_FAZER,
      status: COLUNAS_ETAPAS.STATUS
    };
    Object.keys(mapa).forEach(function(campo) {
      if (dados[campo] !== undefined && dados[campo] !== null) {
        linhaAtualizada[mapa[campo]] = dados[campo];
      }
    });
    aba.getRange(linhaEncontrada, 1, 1, linhaAtualizada.length).setValues([linhaAtualizada]);

    return {
      sucesso: true,
      mensagem: 'Atividade "' + idExistente + '" atualizada com sucesso (linha ' + linhaEncontrada + ')'
    };

  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao atualizar atividade: ' + e.message };
  }
}


/**
 * Atualiza um SETOR existente na planilha pelo ID.
 */
function atualizarSetorDaIdentificacao(idExistente, dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName('Setores');
    if (!aba) return { sucesso: false, mensagem: 'Aba "Setores" não encontrada' };

    var dataRange = aba.getDataRange();
    var valores = dataRange.getValues();
    var linhaEncontrada = -1;

    for (var i = 1; i < valores.length; i++) {
      if (String(valores[i][COLUNAS_SETORES.ID]).trim() === String(idExistente).trim()) {
        linhaEncontrada = i + 1;
        break;
      }
    }

    if (linhaEncontrada === -1) {
      return { sucesso: false, mensagem: 'Setor com ID "' + idExistente + '" não encontrado na planilha' };
    }

    const indice = linhaEncontrada - 1;
    const linhaAtualizada = valores[indice].slice();
    const mapa = {
      nome: COLUNAS_SETORES.NOME,
      descricao: COLUNAS_SETORES.DESCRICAO,
      responsaveisIds: COLUNAS_SETORES.RESPONSAVEIS_IDS
    };
    Object.keys(mapa).forEach(function(campo) {
      if (dados[campo] !== undefined && dados[campo] !== null) {
        linhaAtualizada[mapa[campo]] = dados[campo];
      }
    });
    aba.getRange(linhaEncontrada, 1, 1, linhaAtualizada.length).setValues([linhaAtualizada]);

    return { sucesso: true, mensagem: 'Setor "' + idExistente + '" atualizado com sucesso (linha ' + linhaEncontrada + ')' };

  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao atualizar setor: ' + e.message };
  }
}

function extrairEtapasDoRelatorioV2(conteudo) {
  const atividades = [];

  try {
    // ✅ CORRIGIDO: regex agora reconhece "ATIVIDADES" e "ETAPAS"
    // e o fim da seção pode ser o checklist ⚠️ ou outra seção ##
    const padroeSecao = [
      // Novo formato: "ATIVIDADES IDENTIFICADAS"
      /##\s*📋\s*ATIVIDADES IDENTIFICADAS[^\n]*\n([\s\S]*?)(?=\n---\n##\s*⚠️|\n##\s*⚠️|\n---\n##|\n##\s*❌|$)/i,
      // Formato antigo: "ETAPAS IDENTIFICADAS"
      /##\s*📋\s*ETAPAS IDENTIFICADAS[^\n]*\n([\s\S]*?)(?=\n---\n##\s*⚠️|\n##\s*⚠️|\n---\n##|\n##\s*❌|$)/i,
      // Fallback genérico com 📋
      /##\s*📋[^\n]*\n([\s\S]*?)(?=\n---\n##\s*⚠️|\n##\s*⚠️|\n---\n##|$)/i
    ];

    let secao = '';
    for (const padrao of padroeSecao) {
      const match = conteudo.match(padrao);
      if (match) {
        secao = match[1];
        Logger.log('Seção de atividades encontrada, tamanho: ' + secao.length);
        break;
      }
    }

    if (!secao) {
      Logger.log('Seção de atividades/etapas NÃO encontrada');
      return atividades;
    }

    const itens = dividirPorItens(secao);
    Logger.log('Itens de atividade encontrados: ' + itens.length);

    for (const item of itens) {
      const atividade = parsearItemGenerico(item, 'etapa');
      if (atividade && atividade.campos && Object.keys(atividade.campos).length > 0) {
        atividades.push(atividade);
      }
    }

  } catch (erro) {
    Logger.log('ERRO extrairEtapasDoRelatorioV2: ' + erro.toString());
  }

  return atividades;
}

function extrairSetoresDoRelatorioV2(conteudo) {
  const setores = [];

  try {
    const padroeSecao = [
      // Novo formato: "SETOR BENEFICIADO" (singular, novo padrão)
      /##\s*🏢\s*SETOR BENEFICIADO[^\n]*\n([\s\S]*?)(?=\n---\n##\s*📁|\n##\s*📁|$)/i,
      // Formato antigo: "SETORES COM PROJETOS"
      /##\s*🏢\s*SETORES COM PROJETOS[^\n]*\n([\s\S]*?)(?=\n---\n##\s*📁|\n##\s*📁|$)/i,
      // Fallback genérico 🏢
      /##\s*🏢[^\n]*\n([\s\S]*?)(?=\n---\n##\s*📁|\n##\s*📁|$)/i
    ];

    let secao = '';
    for (const padrao of padroeSecao) {
      const match = conteudo.match(padrao);
      if (match) {
        secao = match[1];
        Logger.log('Seção de setores encontrada, tamanho: ' + secao.length);
        break;
      }
    }

    if (!secao) {
      Logger.log('Seção de setores NÃO encontrada');
      return setores;
    }

    const itens = dividirPorItens(secao);
    Logger.log('Itens de setor encontrados: ' + itens.length);

    for (const item of itens) {
      const setor = parsearItemGenerico(item, 'setor');
      if (setor && setor.campos && Object.keys(setor.campos).length > 0) {
        setores.push(setor);
      }
    }

  } catch (erro) {
    Logger.log('ERRO extrairSetoresDoRelatorioV2: ' + erro.toString());
  }

  return setores;
}

function extrairProjetosDoRelatorioV2(conteudo) {
  const projetos = [];

  try {
    const padroeSecao = [
      /##\s*📁\s*PROJETOS IDENTIFICADOS[^\n]*\n([\s\S]*?)(?=\n---\n##\s*📋|\n##\s*📋|$)/i,
      /##\s*PROJETOS IDENTIFICADOS[^\n]*\n([\s\S]*?)(?=\n---\n##\s*📋|\n##\s*📋|$)/i,
      /##\s*📁[^\n]*PROJETOS[^\n]*\n([\s\S]*?)(?=\n##\s*📋|$)/i
    ];

    let secao = '';
    for (const padrao of padroeSecao) {
      const match = conteudo.match(padrao);
      if (match) {
        secao = match[1];
        Logger.log('Seção de projetos encontrada, tamanho: ' + secao.length);
        break;
      }
    }

    if (!secao) {
      Logger.log('Seção de projetos NÃO encontrada');
      return projetos;
    }

    const itens = dividirPorItens(secao);
    Logger.log('Itens de projeto encontrados: ' + itens.length);

    for (const item of itens) {
      const projeto = parsearItemGenerico(item, 'projeto');
      if (projeto && projeto.campos && Object.keys(projeto.campos).length > 0) {
        projetos.push(projeto);
      }
    }

  } catch (erro) {
    Logger.log('ERRO extrairProjetosDoRelatorioV2: ' + erro.toString());
  }

  return projetos;
}



function salvarRelatorioNoDrive(conteudoRelatorio, tituloReuniao, nomeDepartamento) {
  try {
    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMdd_HHmmss');
    // Embute o departamento no nome do arquivo para filtragem por permissão
    var slugDep = _slugificarNomeDep(nomeDepartamento || '');
    var sufixoDep = slugDep ? (MARCADOR_DEP_RELATORIO + slugDep) : '';
    const nomeArquivo = PREFIXO_ARQUIVO_RELATORIO + sufixoDep.replace(/^_/, '') + (sufixoDep ? '_' : '') + timestamp + EXTENSAO_RELATORIO;
    var arquivo = pasta.createFile(nomeArquivo, conteudoRelatorio, MimeType.PLAIN_TEXT);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('Relatório salvo: ' + nomeArquivo);

    return {
      nomeArquivo: nomeArquivo,
      idArquivo: arquivo.getId(),
      linkArquivo: arquivo.getUrl(),
      sucesso: true
    };

  } catch (erro) {
    Logger.log('ERRO salvarRelatorioNoDrive: ' + erro.toString());
    return { sucesso: false, nomeArquivo: '', linkArquivo: '' };
  }
}

function extrairTemaParaNomeArquivo(conteudo) {
  try {
    // Procura a linha de tema central
    const matchTema = conteudo.match(/\*\*Tema Central da Reuni[aã]o:\*\*\s*([^\n]+)/i);
    if (!matchTema) return '';

    let tema = matchTema[1].trim();

    // Pegar apenas as primeiras palavras significativas (máx 60 chars antes de slugificar)
    if (tema.length > 80) {
      // Cortar na última palavra completa antes de 80 chars
      tema = tema.substring(0, 80).replace(/\s+\S*$/, '');
    }

    // Slugificar: remover acentos, manter apenas letras/números/espaços, trocar espaços por hífen
    const slug = tema
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')        // remove acentos
      .replace(/[^a-zA-Z0-9\s]/g, '')         // remove caracteres especiais
      .trim()
      .replace(/\s+/g, '-')                   // espaços → hífens
      .replace(/-{2,}/g, '-')                 // múltiplos hífens → 1
      .substring(0, 60);                      // máx 60 chars

    return slug || '';

  } catch (e) {
    Logger.log('ERRO extrairTemaParaNomeArquivo: ' + e.toString());
    return '';
  }
}

function extrairInformacoesGerais(conteudo) {
  const info = {};

  try {
    const padroes = [
      /##\s*📅\s*Informações Gerais([\s\S]*?)(?=\n##|\n---\n##|$)/i,
      /##\s*Informações Gerais([\s\S]*?)(?=\n##|\n---\n##|$)/i,
    ];

    let secaoInfo = '';
    for (const padrao of padroes) {
      const match = conteudo.match(padrao);
      if (match) { secaoInfo = match[1]; break; }
    }

    if (secaoInfo) {
      const regexCampo = /\*\*([^*:]+)[:\*]+\s*([^\n*]+)/g;
      let match;
      while ((match = regexCampo.exec(secaoInfo)) !== null) {
        const chave = match[1].trim().toLowerCase()
          .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, '_');
        info[chave] = match[2].trim();
      }
    }
  } catch (e) {
    Logger.log('Erro extrairInformacoesGerais: ' + e.toString());
  }

  return info;
}
