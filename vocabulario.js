// ═══════════════════════════════════════════════════════════════════════
//  VOCABULARIO.GS — Sistema de Correção de Transcrição em 3 Camadas
//
//  Camada 1: Glossário no Prompt (preventiva — resolve ~85-90%)
//  Camada 2: Similaridade Levenshtein (corretiva inteligente — ~8-10%)
//  Camada 3: Substituição Direta por Variações (corretiva bruta — ~2-3%)
// ═══════════════════════════════════════════════════════════════════════

// ──────────────────────────────────────────────────────────────────────
//  CONSTANTES — ABA E COLUNAS
// ──────────────────────────────────────────────────────────────────────

const NOME_ABA_VOCABULARIO = 'Vocabulário';

const COLUNAS_VOCABULARIO = {
  ID: 0,
  TERMO_CORRETO: 1,
  CATEGORIA: 2,
  PRONUNCIA: 3,
  VARIACOES: 4,
  DESCRICAO: 5,
  PALAVRAS_VIZINHAS: 6,
  ATIVO: 7,
  PRIORIDADE: 8,
  DATA_CRIACAO: 9,
  ULTIMA_CORRECAO: 10
};

const CATEGORIAS_VOCABULARIO = {
  PESSOA: 'Pessoa',
  SISTEMA: 'Sistema',
  SETOR: 'Setor',
  LOCAL: 'Local',
  SIGLA: 'Sigla',
  TECNICO: 'Técnico',
  OUTRO: 'Outro'
};

// ──────────────────────────────────────────────────────────────────────
//  CONSTANTES — CONFIGURAÇÃO DA VALIDAÇÃO PÓS-TRANSCRIÇÃO
// ──────────────────────────────────────────────────────────────────────

const CONFIG_VALIDACAO_TRANSCRICAO = {
  LIMIAR_SIMILARIDADE_MINIMO: 0.82,   // abaixo disso → NÃO corrige
  LIMIAR_SIMILARIDADE_AUTO: 0.92,     // acima disso → corrige sem verificar contexto
  TAMANHO_MINIMO_PALAVRA: 3,          // palavras menores são ignoradas
  IGNORAR_PALAVRAS_COMUNS: true,
  SUBSTITUICAO_CASE_INSENSITIVE: true,
  RESPEITAR_LIMITES_PALAVRA: true,    // nunca substitui substrings
  MAX_CORRECOES_POR_TRANSCRICAO: 500, // trava de segurança
  REGISTRAR_CORRECOES_NO_LOG: true
};

// ──────────────────────────────────────────────────────────────────────
//  CONSTANTES — PALAVRAS COMUNS QUE NUNCA DEVEM SER CORRIGIDAS
//  (usa objeto para busca O(1) ao invés de array.indexOf)
// ──────────────────────────────────────────────────────────────────────

const PALAVRAS_COMUNS_IGNORAR = (function() {
  const lista = [
    'o','a','os','as','um','uma','uns','umas',
    'de','do','da','dos','das','em','no','na','nos','nas',
    'por','para','com','sem','sob','sobre','entre','ate',
    'ao','aos','pelo','pela','pelos','pelas',
    'e','ou','mas','que','se','como','quando','porque',
    'pois','nem','porem','contudo','todavia','embora',
    'eu','tu','ele','ela','voce','voces',
    'me','te','lhe','lhes',
    'meu','minha','seu','sua','nosso','nossa',
    'este','esta','esse','essa','aquele','aquela',
    'isto','isso','aquilo',
    'sao','foi','era','ser','estar','estao',
    'tem','tinha','ter','vai','vao','ir',
    'faz','fazer','dar','pode','poder',
    'quer','querer','sabe','saber','vem','vir',
    'houve','havia','disse','dizer','falou','falar',
    'nao','sim','ja','mais','muito','tambem','aqui','ali',
    'la','bem','mal','ainda','sempre','nunca',
    'entao','depois','antes','agora','hoje','ontem','amanha',
    'ai','assim','so','tudo','nada','algo','cada',
    'inaudivel','pausa','participante',
    'dois','tres','quatro','cinco','seis','sete',
    'oito','nove','dez','vinte','trinta','cem','mil'
  ];
  const mapa = {};
  lista.forEach(function(p) { mapa[p] = true; });
  return mapa;
})();

// ══════════════════════════════════════════════════════════════════════
//  CRIAÇÃO AUTOMÁTICA DA ABA
// ══════════════════════════════════════════════════════════════════════

/**
 * Cria a aba "Vocabulário" com cabeçalho e validações se ela não existir.
 * Chamada automaticamente por carregarVocabularioCompleto().
 */
function criarAbaVocabularioSeNecessario() {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    let aba = planilha.getSheetByName(NOME_ABA_VOCABULARIO);
    if (aba) return aba;

    aba = planilha.insertSheet(NOME_ABA_VOCABULARIO);

    const cabecalho = [
      'ID', 'TERMO_CORRETO', 'CATEGORIA', 'PRONUNCIA', 'VARIACOES',
      'DESCRICAO', 'PALAVRAS_VIZINHAS', 'ATIVO', 'PRIORIDADE',
      'DATA_CRIACAO', 'ULTIMA_CORRECAO'
    ];
    const rangeCabecalho = aba.getRange(1, 1, 1, cabecalho.length);
    rangeCabecalho.setValues([cabecalho]).setFontWeight('bold').setBackground('#2C1810').setFontColor('#F5EFE0');
    aba.setFrozenRows(1);

    // Validação de lista para CATEGORIA
    const validacaoCategoria = SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.values(CATEGORIAS_VOCABULARIO), true)
      .setAllowInvalid(false)
      .build();
    aba.getRange(2, COLUNAS_VOCABULARIO.CATEGORIA + 1, 1000, 1).setDataValidation(validacaoCategoria);

    // Validação de booleano para ATIVO
    const validacaoAtivo = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE', 'FALSE'], true)
      .setAllowInvalid(false)
      .build();
    aba.getRange(2, COLUNAS_VOCABULARIO.ATIVO + 1, 1000, 1).setDataValidation(validacaoAtivo);

    // Validação de range para PRIORIDADE (1-5)
    const validacaoPrioridade = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 5)
      .setAllowInvalid(false)
      .build();
    aba.getRange(2, COLUNAS_VOCABULARIO.PRIORIDADE + 1, 1000, 1).setDataValidation(validacaoPrioridade);

    // Larguras das colunas
    aba.setColumnWidth(COLUNAS_VOCABULARIO.TERMO_CORRETO + 1, 160);
    aba.setColumnWidth(COLUNAS_VOCABULARIO.PRONUNCIA + 1, 120);
    aba.setColumnWidth(COLUNAS_VOCABULARIO.VARIACOES + 1, 200);
    aba.setColumnWidth(COLUNAS_VOCABULARIO.PALAVRAS_VIZINHAS + 1, 220);
    aba.setColumnWidth(COLUNAS_VOCABULARIO.DESCRICAO + 1, 220);

    Logger.log('[criarAbaVocabularioSeNecessario] Aba "Vocabulário" criada com sucesso.');
    return aba;

  } catch (erro) {
    Logger.log('ERRO criarAbaVocabularioSeNecessario: ' + erro.toString());
    return null;
  }
}

// ══════════════════════════════════════════════════════════════════════
//  CAMADA 1 — CARREGAMENTO E MONTAGEM DO GLOSSÁRIO
// ══════════════════════════════════════════════════════════════════════

/**
 * Lê toda a aba "Vocabulário" e retorna um objeto estruturado com:
 * - lista de termos ativos
 * - mapa de variações → termo correto (para Camada 3)
 * - mapa de termos em lowercase (para skip na Camada 2)
 *
 * Deve ser chamada UMA VEZ por fluxo de processamento e o resultado
 * reutilizado em todas as etapas.
 */
function carregarVocabularioCompleto() {
  try {
    criarAbaVocabularioSeNecessario();

    const aba = obterAba(NOME_ABA_VOCABULARIO);
    if (!aba || aba.getLastRow() <= 1) {
      return _vocabularioVazio();
    }

    const dados = aba.getDataRange().getValues();
    const termos = [];
    const mapaVariacoes = {};    // { 'variacao_normalizada': 'TermoCorreto' }
    const mapaTermosLower = {};  // { 'termo_normalizado': { termoCorreto, categoria, palavrasVizinhas } }

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];

      // Ignora linhas sem ID ou desativadas
      const ativoRaw = linha[COLUNAS_VOCABULARIO.ATIVO];
      const ativo = ativoRaw === true || String(ativoRaw).toUpperCase() === 'TRUE';
      if (!ativo) continue;

      const termoCorreto = linha[COLUNAS_VOCABULARIO.TERMO_CORRETO]
        ? linha[COLUNAS_VOCABULARIO.TERMO_CORRETO].toString().trim() : '';
      if (!termoCorreto) continue;

      // Processa variações (separadas por vírgula)
      const variacoesRaw = linha[COLUNAS_VOCABULARIO.VARIACOES]
        ? linha[COLUNAS_VOCABULARIO.VARIACOES].toString().trim() : '';
      const variacoes = variacoesRaw
        ? variacoesRaw.split(',').map(function(v) { return v.trim(); }).filter(function(v) { return v.length > 0; })
        : [];

      // Processa palavras vizinhas (separadas por vírgula)
      const vizinhasRaw = linha[COLUNAS_VOCABULARIO.PALAVRAS_VIZINHAS]
        ? linha[COLUNAS_VOCABULARIO.PALAVRAS_VIZINHAS].toString().trim() : '';
      const palavrasVizinhas = vizinhasRaw
        ? vizinhasRaw.split(',').map(function(p) { return p.trim(); }).filter(function(p) { return p.length > 0; })
        : [];

      const termo = {
        id: linha[COLUNAS_VOCABULARIO.ID] ? linha[COLUNAS_VOCABULARIO.ID].toString() : '',
        termoCorreto: termoCorreto,
        categoria: linha[COLUNAS_VOCABULARIO.CATEGORIA]
          ? linha[COLUNAS_VOCABULARIO.CATEGORIA].toString().trim() : '',
        pronuncia: linha[COLUNAS_VOCABULARIO.PRONUNCIA]
          ? linha[COLUNAS_VOCABULARIO.PRONUNCIA].toString().trim() : '',
        variacoes: variacoes,
        descricao: linha[COLUNAS_VOCABULARIO.DESCRICAO]
          ? linha[COLUNAS_VOCABULARIO.DESCRICAO].toString().trim() : '',
        palavrasVizinhas: palavrasVizinhas,
        prioridade: parseInt(linha[COLUNAS_VOCABULARIO.PRIORIDADE]) || 3,
        linhaIndice: i + 1  // linha real na planilha (para atualização futura)
      };

      termos.push(termo);

      // Popula mapa de termos corretos (para skip na Camada 2)
      const termoNorm = normalizarParaComparacao(termoCorreto);
      mapaTermosLower[termoNorm] = {
        termoCorreto: termoCorreto,
        categoria: termo.categoria,
        palavrasVizinhas: palavrasVizinhas
      };

      // Popula mapa de variações → termo correto (para Camada 3)
      variacoes.forEach(function(v) {
        const vNorm = normalizarParaComparacao(v);
        if (vNorm && vNorm.length >= 2 && !mapaVariacoes[vNorm]) {
          mapaVariacoes[vNorm] = termoCorreto;
        }
      });
    }

    Logger.log('[carregarVocabularioCompleto] ' + termos.length + ' termos ativos carregados.');

    return {
      termos: termos,
      mapaVariacoes: mapaVariacoes,
      mapaTermosLower: mapaTermosLower,
      totalTermos: termos.length,
      dataCarregamento: Date.now()
    };

  } catch (erro) {
    Logger.log('ERRO carregarVocabularioCompleto: ' + erro.toString());
    return _vocabularioVazio();
  }
}

/** Retorna estrutura vazia (evita null-checks espalhados pelo código) */
function _vocabularioVazio() {
  return { termos: [], mapaVariacoes: {}, mapaTermosLower: {}, totalTermos: 0, dataCarregamento: Date.now() };
}

// ──────────────────────────────────────────────────────────────────────

/**
 * Gera o bloco de texto do glossário para injeção no prompt de transcrição.
 * Agrupa por categoria, ordena por prioridade, inclui pronúncias.
 * Se houver mais de 150 termos ativos, filtra apenas prioridade >= 3.
 */
function montarGlossarioParaPrompt(vocabulario) {
  if (!vocabulario || vocabulario.totalTermos === 0) return '';

  // Ordena por prioridade decrescente e filtra se necessário
  var termos = vocabulario.termos.slice().sort(function(a, b) { return b.prioridade - a.prioridade; });
  if (termos.length > 150) {
    termos = termos.filter(function(t) { return t.prioridade >= 3; });
  }

  // Mapa de categoria → nome amigável no glossário (une categorias afins)
  var NOME_GRUPO = {
    'Pessoa': 'Pessoas',
    'Sistema': 'Sistemas e Siglas',
    'Sigla': 'Sistemas e Siglas',
    'Setor': 'Setores e Locais',
    'Local': 'Setores e Locais',
    'Técnico': 'Termos Técnicos',
    'Outro': 'Outros Termos'
  };
  var ORDEM_GRUPOS = ['Pessoas', 'Sistemas e Siglas', 'Setores e Locais', 'Termos Técnicos', 'Outros Termos'];

  // Agrupa termos por nome de grupo amigável
  var grupos = {};
  termos.forEach(function(t) {
    var grupo = NOME_GRUPO[t.categoria] || 'Outros Termos';
    if (!grupos[grupo]) grupos[grupo] = [];
    grupos[grupo].push(t);
  });

  var bloco = '## 🔤 VOCABULÁRIO OBRIGATÓRIO — Use estas grafias EXATAS\n\n';
  bloco += 'Ao transcrever, sempre que ouvir algo foneticamente SIMILAR aos termos abaixo,\n';
  bloco += 'use a grafia EXATA indicada. NÃO invente ocorrências — use APENAS quando\n';
  bloco += 'realmente ouvir o termo no áudio.\n\n';

  ORDEM_GRUPOS.forEach(function(grupo) {
    if (!grupos[grupo] || grupos[grupo].length === 0) return;
    bloco += '### ' + grupo + ':\n';
    grupos[grupo].forEach(function(t) {
      var linha = '- ' + t.termoCorreto;
      if (t.pronuncia) linha += ' (pronunciado "' + t.pronuncia + '")';
      if (t.descricao) linha += ' — ' + t.descricao;
      bloco += linha + '\n';
    });
    bloco += '\n';
  });

  bloco += '## ⚠️ REGRAS CRÍTICAS DO VOCABULÁRIO:\n';
  bloco += '1. Use a grafia EXATA listada acima (incluindo maiúsculas/minúsculas e acentos)\n';
  bloco += '2. NÃO insira estes termos se eles NÃO foram falados no áudio\n';
  bloco += '3. Nomes de pessoas SEMPRE com inicial maiúscula\n';
  bloco += '4. Siglas de sistemas SEMPRE em MAIÚSCULAS\n\n';

  return bloco;
}

// ──────────────────────────────────────────────────────────────────────

/**
 * Gera bloco de exemplos "anchoring" para o prompt.
 * Usa a coluna VARIACOES para mostrar a IA exemplos de erros comuns
 * e a grafia correta. Limitado a 8 exemplos (máximo 1 por termo prioritário).
 */
function montarExemplosAnchoring(vocabulario) {
  if (!vocabulario || vocabulario.totalTermos === 0) return '';

  var termosComExemplos = vocabulario.termos
    .filter(function(t) { return t.prioridade >= 4 && t.variacoes.length > 0; })
    .sort(function(a, b) { return b.prioridade - a.prioridade; })
    .slice(0, 8);

  if (termosComExemplos.length === 0) return '';

  var bloco = '## 📝 EXEMPLOS DE TRANSCRIÇÃO CORRETA (siga estes padrões):\n\n';

  termosComExemplos.forEach(function(t) {
    var variacao = t.variacoes[0]; // usa a primeira variação como exemplo do erro
    bloco += 'INCORRETO: "...o ' + variacao + ' apresentou problema..."\n';
    bloco += 'CORRETO:   "...o ' + t.termoCorreto + ' apresentou problema..."\n\n';
  });

  return bloco;
}

// ══════════════════════════════════════════════════════════════════════
//  CAMADAS 2 E 3 — VALIDAÇÃO PÓS-TRANSCRIÇÃO
// ══════════════════════════════════════════════════════════════════════

/**
 * Função principal de validação pós-transcrição.
 * Executa Camada 3 (substituição direta) primeiro, depois Camada 2 (similaridade).
 *
 * @param {string} transcricaoBruta - Texto bruto retornado pelo Gemini
 * @param {Object} vocabulario - Objeto retornado por carregarVocabularioCompleto()
 * @returns {{ transcricaoCorrigida, totalCorrecoes, correcoes[] }}
 */
function validarECorrigirTranscricao(transcricaoBruta, vocabulario) {
  var resultado = {
    transcricaoCorrigida: transcricaoBruta,
    totalCorrecoes: 0,
    correcoes: []
  };

  // Proteção: transcrições muito curtas ou vocabulário vazio → retorna sem alteração
  if (!transcricaoBruta || transcricaoBruta.length < 100) return resultado;
  if (!vocabulario || vocabulario.totalTermos === 0) return resultado;

  var texto = transcricaoBruta;

  // ────────────────────────────────────────────
  //  CAMADA 3 — Substituição Direta por Variações
  //  (executa PRIMEIRO — mais rápida e mais precisa)
  // ────────────────────────────────────────────

  // Ordena por comprimento decrescente para substituir "sipaque" antes de "sipa"
  var entradasVariacoes = Object.keys(vocabulario.mapaVariacoes)
    .sort(function(a, b) { return b.length - a.length; });

  for (var vi = 0; vi < entradasVariacoes.length; vi++) {
    if (resultado.totalCorrecoes >= CONFIG_VALIDACAO_TRANSCRICAO.MAX_CORRECOES_POR_TRANSCRICAO) {
      Logger.log('[validarECorrigirTranscricao] AVISO: Limite de ' + CONFIG_VALIDACAO_TRANSCRICAO.MAX_CORRECOES_POR_TRANSCRICAO + ' correções atingido!');
      break;
    }

    var variacaoNorm = entradasVariacoes[vi];
    var termoCorreto = vocabulario.mapaVariacoes[variacaoNorm];

    if (!variacaoNorm || variacaoNorm.length < CONFIG_VALIDACAO_TRANSCRICAO.TAMANHO_MINIMO_PALAVRA) continue;

    var posicao = 0;
    while (posicao < texto.length) {
      var textoLower = texto.toLowerCase();
      var idx = textoLower.indexOf(variacaoNorm, posicao);
      if (idx === -1) break;

      // Verifica limites de palavra (anti-substring)
      var limiteEsquerda = idx === 0 || ehLimiteDePalavra(texto, idx - 1);
      var limiteDireita  = (idx + variacaoNorm.length >= texto.length) || ehLimiteDePalavra(texto, idx + variacaoNorm.length);

      if (limiteEsquerda && limiteDireita) {
        var trechoOriginal = texto.substring(idx, idx + variacaoNorm.length);
        texto = texto.substring(0, idx) + termoCorreto + texto.substring(idx + variacaoNorm.length);

        resultado.correcoes.push({
          original: trechoOriginal,
          corrigido: termoCorreto,
          camada: 3,
          confianca: 1.0,
          posicao: idx,
          contexto: _obterContexto(transcricaoBruta, idx, 30)
        });
        resultado.totalCorrecoes++;
        posicao = idx + termoCorreto.length; // avança após a substituição
      } else {
        posicao = idx + 1;
      }
    }
  }

  // ────────────────────────────────────────────
  //  CAMADA 2 — Similaridade (Levenshtein)
  //  Executa sobre o texto já parcialmente corrigido pela Camada 3
  // ────────────────────────────────────────────

  // Coleta palavras únicas do texto para evitar recalcular Levenshtein
  // para a mesma palavra que aparece múltiplas vezes
  var tokenRegex = /[^\s\[\].,;:!?()"'\-—\d:]+/g;
  var palavrasUnicas = {};
  var matchToken;

  while ((matchToken = tokenRegex.exec(texto)) !== null) {
    var palavra = matchToken[0];
    if (palavra.length < CONFIG_VALIDACAO_TRANSCRICAO.TAMANHO_MINIMO_PALAVRA) continue;
    var palavraNorm = normalizarParaComparacao(palavra);
    if (!palavraNorm || palavraNorm.length < CONFIG_VALIDACAO_TRANSCRICAO.TAMANHO_MINIMO_PALAVRA) continue;
    if (PALAVRAS_COMUNS_IGNORAR[palavraNorm]) continue;
    if (vocabulario.mapaTermosLower[palavraNorm]) continue; // já é um termo correto
    if (!palavrasUnicas[palavra]) {
      palavrasUnicas[palavra] = { norm: palavraNorm, primeiraPos: matchToken.index };
    }
  }

  // Para cada palavra única, calcula a melhor correção possível
  var mapaCorrecoesCamada2 = {};

  for (var palavraOriginal in palavrasUnicas) {
    if (resultado.totalCorrecoes >= CONFIG_VALIDACAO_TRANSCRICAO.MAX_CORRECOES_POR_TRANSCRICAO) break;

    var info = palavrasUnicas[palavraOriginal];
    var melhorSim = 0;
    var melhorTermo = null;

    for (var ti = 0; ti < vocabulario.termos.length; ti++) {
      var termo = vocabulario.termos[ti];
      var termoNorm = normalizarParaComparacao(termo.termoCorreto);

      // Pré-filtro de comprimento: impossível ter sim >= 0.82 se diferença > 3 chars
      if (Math.abs(info.norm.length - termoNorm.length) > 3) continue;

      var sim = calcularSimilaridade(info.norm, termoNorm);
      if (sim > melhorSim) {
        melhorSim = sim;
        melhorTermo = termo;
      }
    }

    if (!melhorTermo) continue;

    var deveCorrigir = false;

    if (melhorSim >= CONFIG_VALIDACAO_TRANSCRICAO.LIMIAR_SIMILARIDADE_AUTO) {
      // Alta confiança → corrige automaticamente
      deveCorrigir = true;
    } else if (melhorSim >= CONFIG_VALIDACAO_TRANSCRICAO.LIMIAR_SIMILARIDADE_MINIMO) {
      // Confiança média → verifica contexto (palavras vizinhas)
      var vizinhas = _extrairPalavrasVizinhas(texto, info.primeiraPos, palavraOriginal.length, 5);
      var temContexto = melhorTermo.palavrasVizinhas.some(function(pv) {
        var pvNorm = normalizarParaComparacao(pv);
        return vizinhas.some(function(v) {
          var vNorm = normalizarParaComparacao(v);
          return vNorm === pvNorm || vNorm.indexOf(pvNorm) !== -1;
        });
      });
      deveCorrigir = temContexto;
    }

    if (deveCorrigir) {
      mapaCorrecoesCamada2[palavraOriginal] = {
        corrigido: melhorTermo.termoCorreto,
        confianca: melhorSim
      };
    }
  }

  // Aplica correções da Camada 2 com verificação de palavra inteira
  for (var palavraCamada2 in mapaCorrecoesCamada2) {
    if (resultado.totalCorrecoes >= CONFIG_VALIDACAO_TRANSCRICAO.MAX_CORRECOES_POR_TRANSCRICAO) break;

    var infoCor = mapaCorrecoesCamada2[palavraCamada2];
    var pos = 0;

    while (pos < texto.length) {
      var idxC2 = texto.indexOf(palavraCamada2, pos);
      if (idxC2 === -1) break;

      var limEsqC2 = idxC2 === 0 || ehLimiteDePalavra(texto, idxC2 - 1);
      var limDirC2 = (idxC2 + palavraCamada2.length >= texto.length) || ehLimiteDePalavra(texto, idxC2 + palavraCamada2.length);

      if (limEsqC2 && limDirC2) {
        texto = texto.substring(0, idxC2) + infoCor.corrigido + texto.substring(idxC2 + palavraCamada2.length);

        resultado.correcoes.push({
          original: palavraCamada2,
          corrigido: infoCor.corrigido,
          camada: 2,
          confianca: infoCor.confianca,
          posicao: idxC2,
          contexto: _obterContexto(transcricaoBruta, idxC2, 30)
        });
        resultado.totalCorrecoes++;
        pos = idxC2 + infoCor.corrigido.length;
      } else {
        pos = idxC2 + 1;
      }
    }
  }

  resultado.transcricaoCorrigida = texto;

  if (CONFIG_VALIDACAO_TRANSCRICAO.REGISTRAR_CORRECOES_NO_LOG) {
    Logger.log('[validarECorrigirTranscricao] Total de correções aplicadas: ' + resultado.totalCorrecoes);
    if (resultado.totalCorrecoes > 0 && resultado.totalCorrecoes <= 20) {
      resultado.correcoes.forEach(function(c) {
        Logger.log('  [Camada ' + c.camada + '] "' + c.original + '" → "' + c.corrigido + '" (conf: ' + c.confianca.toFixed(2) + ')');
      });
    }
  }

  return resultado;
}

// ══════════════════════════════════════════════════════════════════════
//  LEVENSHTEIN E SIMILARIDADE
// ══════════════════════════════════════════════════════════════════════

/**
 * Calcula a distância de edição entre duas strings.
 * Usa apenas 2 linhas da matriz DP → O(min(m,n)) de memória.
 * SEMPRE comparar strings já normalizadas (sem acentos, lowercase).
 */
function calcularDistanciaLevenshtein(a, b) {
  if (!a) return b ? b.length : 0;
  if (!b) return a.length;
  if (a === b) return 0;

  var m = a.length, n = b.length;
  var anterior = new Array(n + 1);
  var atual    = new Array(n + 1);

  for (var j = 0; j <= n; j++) anterior[j] = j;

  for (var i = 1; i <= m; i++) {
    atual[0] = i;
    for (var jj = 1; jj <= n; jj++) {
      var custo = a[i - 1] === b[jj - 1] ? 0 : 1;
      atual[jj] = Math.min(
        anterior[jj] + 1,          // deleção
        atual[jj - 1] + 1,         // inserção
        anterior[jj - 1] + custo   // substituição
      );
    }
    var temp = anterior; anterior = atual; atual = temp;
  }

  return anterior[n];
}

/**
 * Retorna similaridade normalizada entre 0.0 e 1.0.
 * 1.0 = strings idênticas, 0.0 = completamente diferentes.
 */
function calcularSimilaridade(a, b) {
  if (!a && !b) return 1.0;
  if (!a || !b) return 0.0;
  var maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1.0;
  return 1 - (calcularDistanciaLevenshtein(a, b) / maxLen);
}

// ══════════════════════════════════════════════════════════════════════
//  NORMALIZAÇÃO E LIMITES DE PALAVRA
// ══════════════════════════════════════════════════════════════════════

/**
 * Normaliza string para comparação fonética/similaridade.
 * Remove acentos, converte para minúsculas.
 * NUNCA usar o resultado para substituição — só para comparar.
 */
function normalizarParaComparacao(texto) {
  if (!texto) return '';
  return texto
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')   // remove diacríticos (á→a, ç→c, etc.)
    .replace(/[^a-z0-9\s]/g, '')       // remove pontuação/especiais
    .trim();
}

/**
 * Retorna true se o caractere na posição for um separador de palavras.
 * Usado para garantir que só substituímos palavras inteiras (nunca substrings).
 * ⚠️ Não usar \b do regex: não reconhece acentos em JS/GAS.
 */
function ehLimiteDePalavra(texto, posicao) {
  if (posicao < 0 || posicao >= texto.length) return true; // bordas do texto
  var char = texto.charAt(posicao);
  var separadores = ' \t\n\r.,;:!?()[]{}"\'\u2014-/\\';  // \u2014 = travessão —
  return separadores.indexOf(char) !== -1;
}

// ══════════════════════════════════════════════════════════════════════
//  ATUALIZAÇÃO DA PLANILHA
// ══════════════════════════════════════════════════════════════════════

/**
 * Atualiza a coluna ULTIMA_CORRECAO para os termos que foram corrigidos.
 * Chamada em batch após a validação para minimizar operações na planilha.
 */
function atualizarUltimaCorrecaoVocabulario(idsTermosCorrigidos) {
  if (!idsTermosCorrigidos || idsTermosCorrigidos.length === 0) return;
  try {
    var aba = obterAba(NOME_ABA_VOCABULARIO);
    if (!aba || aba.getLastRow() <= 1) return;

    var idsSet = {};
    idsTermosCorrigidos.forEach(function(id) { idsSet[id] = true; });

    var colunaIds = aba.getRange(2, COLUNAS_VOCABULARIO.ID + 1, aba.getLastRow() - 1, 1).getValues();
    var dataHoje = new Date();

    for (var i = 0; i < colunaIds.length; i++) {
      var idCelula = colunaIds[i][0] ? colunaIds[i][0].toString() : '';
      if (idsSet[idCelula]) {
        aba.getRange(i + 2, COLUNAS_VOCABULARIO.ULTIMA_CORRECAO + 1).setValue(dataHoje);
      }
    }
  } catch (erro) {
    Logger.log('ERRO atualizarUltimaCorrecaoVocabulario: ' + erro.toString());
  }
}

// ══════════════════════════════════════════════════════════════════════
//  AUXILIARES PRIVADOS (prefixo _ indica uso interno)
// ══════════════════════════════════════════════════════════════════════

/** Retorna trecho do texto ao redor de uma posição (para log de contexto) */
function _obterContexto(texto, posicao, raio) {
  var ini = Math.max(0, posicao - raio);
  var fim = Math.min(texto.length, posicao + raio);
  return '...' + texto.substring(ini, fim) + '...';
}

/**
 * Extrai as N palavras antes e depois de uma posição no texto.
 * Usada pela Camada 2 para verificar contexto (palavras vizinhas).
 */
function _extrairPalavrasVizinhas(texto, posicaoPalavra, tamanhoPalavra, raio) {
  var tokenRegex = /[^\s\[\].,;:!?()"'\-—\d]+/g;
  var todasPalavras = [];
  var m;

  while ((m = tokenRegex.exec(texto)) !== null) {
    todasPalavras.push({ palavra: m[0], pos: m.index });
  }

  // Encontra o índice da palavra-alvo pela posição
  var idxAlvo = -1;
  for (var i = 0; i < todasPalavras.length; i++) {
    if (Math.abs(todasPalavras[i].pos - posicaoPalavra) <= 2) {
      idxAlvo = i;
      break;
    }
  }
  if (idxAlvo === -1) return [];

  var vizinhas = [];
  var ini = Math.max(0, idxAlvo - raio);
  var fim = Math.min(todasPalavras.length - 1, idxAlvo + raio);
  for (var j = ini; j <= fim; j++) {
    if (j !== idxAlvo) vizinhas.push(todasPalavras[j].palavra);
  }
  return vizinhas;
}

// ══════════════════════════════════════════════════════════════════════
//  TESTES (executar manualmente pelo Editor para validar o sistema)
// ══════════════════════════════════════════════════════════════════════

/**
 * Executa todos os casos de teste obrigatórios da especificação.
 * Para usar: abrir o Editor GAS → selecionar testarValidacaoTranscricao → Executar.
 */
function testarValidacaoTranscricao() {
  Logger.log('════════════════════════════════════════════');
  Logger.log('TESTES DE VALIDAÇÃO DE TRANSCRIÇÃO');
  Logger.log('════════════════════════════════════════════');

  // Monta vocabulário de teste sem depender da planilha real
  var vocTeste = _montarVocabularioTeste();

  var testes = [
    {
      nome: 'T01 — Substituição básica (Camada 3)',
      entrada: 'o ítaro falou que o cipaque está fora',
      esperado: 'o Ítalo falou que o SIPAC está fora'
    },
    {
      nome: 'T02 — NÃO substituir substring',
      entrada: 'compartilhar os dados',
      esperado: 'compartilhar os dados'
    },
    {
      nome: 'T03 — Case-insensitive',
      entrada: 'o ÍTARO falou',
      esperado: 'o Ítalo falou'
    },
    {
      nome: 'T04 — Múltiplas ocorrências',
      entrada: 'ítaro perguntou pro ítaro se ítaro tinha feito',
      esperado: 'Ítalo perguntou pro Ítalo se Ítalo tinha feito'
    },
    {
      nome: 'T05 — Verbo "siga" NÃO vira SIGAA',
      entrada: 'siga em frente com o relatório',
      esperado: 'siga em frente com o relatório'
    },
    {
      nome: 'T06 — "marca" NÃO vira "Marcos"',
      entrada: 'essa marca de equipamento é boa',
      esperado: 'essa marca de equipamento é boa'
    },
    {
      nome: 'T07 — Palavra já correta não é alterada',
      entrada: 'o SIGAA está funcionando normalmente',
      esperado: 'o SIGAA está funcionando normalmente'
    },
    {
      nome: 'T08 — BI não afeta "possibilidade"',
      entrada: 'a possibilidade de usar o BI é grande',
      esperado: 'a possibilidade de usar o BI é grande'
    },
    {
      nome: 'T09 — Variação início de frase',
      entrada: 'Hitalo começou a reunião',
      esperado: 'Ítalo começou a reunião'
    },
    {
      nome: 'T10 — cigá com contexto SIGAA',
      entrada: 'abrir o ciga pra ver as notas dos alunos',
      esperado: 'abrir o SIGAA pra ver as notas dos alunos'
    }
  ];

  var acertos = 0;
  testes.forEach(function(t) {
    // Pula testes que precisam de transcrição muito longa (< 100 chars é ignorado na função real)
    // Para testar, precisa de textos com 100+ chars. Aqui adaptamos para testes curtos.
    var textoTeste = t.entrada + ' (teste de validação do sistema de correção de transcrição automatizada)';
    var resultado = validarECorrigirTranscricao(textoTeste, vocTeste);
    var saidaReal = resultado.transcricaoCorrigida.replace(' (teste de validação do sistema de correção de transcrição automatizada)', '');
    var ok = saidaReal === t.esperado;
    if (ok) acertos++;
    Logger.log((ok ? '✅' : '❌') + ' ' + t.nome);
    if (!ok) {
      Logger.log('   Entrada:  "' + t.entrada + '"');
      Logger.log('   Esperado: "' + t.esperado + '"');
      Logger.log('   Obtido:   "' + saidaReal + '"');
    }
  });

  Logger.log('────────────────────────────────────────────');
  Logger.log('RESULTADO: ' + acertos + '/' + testes.length + ' testes passaram');
  Logger.log('════════════════════════════════════════════');
  return { acertos: acertos, total: testes.length };
}

/** Monta vocabulário de teste in-memory (não usa a planilha) */
function _montarVocabularioTeste() {
  var termosBase = [
    { id: 'v001', termoCorreto: 'Ítalo', categoria: 'Pessoa', pronuncia: 'ítalo',
      variacoes: ['ítaro', 'hitalo', 'hítalo', 'itálo'], descricao: 'Coordenador BI',
      palavrasVizinhas: ['BI', 'coordenador', 'equipe'], prioridade: 5 },
    { id: 'v002', termoCorreto: 'SIGAA', categoria: 'Sistema', pronuncia: 'sigá',
      variacoes: ['cigá', 'ciga', 'sigah'],  descricao: 'Sistema acadêmico',
      palavrasVizinhas: ['aluno', 'matrícula', 'acadêmico', 'notas', 'disciplina'], prioridade: 5 },
    { id: 'v003', termoCorreto: 'SIPAC', categoria: 'Sistema', pronuncia: 'sipáqui',
      variacoes: ['cipaque', 'sipaque', 'cipac'], descricao: 'Sistema patrimonial',
      palavrasVizinhas: ['patrimônio', 'almoxarifado', 'compra'], prioridade: 5 },
    { id: 'v004', termoCorreto: 'Power BI', categoria: 'Sistema', pronuncia: 'pauer bi',
      variacoes: ['pau BI', 'pauer bi'], descricao: 'Dashboards',
      palavrasVizinhas: ['dashboard', 'relatório', 'gráfico'], prioridade: 4 },
    { id: 'v005', termoCorreto: 'Marcos', categoria: 'Pessoa', pronuncia: 'marcos',
      variacoes: [], descricao: 'Analista',
      palavrasVizinhas: ['analista', 'equipe', 'relatório'], prioridade: 3 }
  ];

  var vocab = { termos: termosBase, mapaVariacoes: {}, mapaTermosLower: {}, totalTermos: termosBase.length, dataCarregamento: Date.now() };
  termosBase.forEach(function(t) {
    vocab.mapaTermosLower[normalizarParaComparacao(t.termoCorreto)] = { termoCorreto: t.termoCorreto, categoria: t.categoria, palavrasVizinhas: t.palavrasVizinhas };
    t.variacoes.forEach(function(v) {
      var vNorm = normalizarParaComparacao(v);
      if (vNorm) vocab.mapaVariacoes[vNorm] = t.termoCorreto;
    });
  });
  return vocab;
}