/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║                     PROMPTS DE IA - Smart Meeting                ║
 * ║  Centralize aqui todos os prompts enviados ao Gemini.            ║
 * ║  Para alterar o comportamento da IA, edite apenas este arquivo.  ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

// ─────────────────────────────────────────────────────────────────────────────
//  CONFIGURAÇÃO DE MODELOS E PARÂMETROS
//  Ajuste modelo, temperatura e tokens de cada etapa aqui.
// ─────────────────────────────────────────────────────────────────────────────
const CONFIG_PROMPTS_REUNIAO = {
  TRANSCRICAO: {
    modelo: 'gemini-2.5-flash',
    temperatura: 0.0,
    maxTokens: 200000,
    pensamento: 1
  },
  EXTRACAO: {
    modelo: 'gemini-2.5-flash',
    temperatura: 0.2,
    maxTokens: 100000,
    pensamento: 0
  },
  ALTERACOES: {
    modelo: 'gemini-2.5-flash',
    temperatura: 0.1,
    maxTokens: 100000,
    pensamento: 0
  },

  // ─── ESTILOS DE ATA ────────────────────────────────────────────────────────
  // Cada estilo tem seus segmentos, e cada segmento tem configuração própria
  // de modelo, temperatura e tokens — dando controle granular sobre cada parte.
  ESTILOS: {

    // ── ATA EXECUTIVA (2 segmentos) ──────────────────────────────────────────
    executiva: {
      segmentos: [
        {
          id: 'decisoes',
          nome: 'Cabeçalho e Decisões Tomadas',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.2,
          maxTokens: 10000,
          pensamento: 0
        },
        {
          id: 'encaminhamentos',
          nome: 'Encaminhamentos e Próximos Passos',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.2,
          maxTokens: 8000,
          pensamento: 0
        }
      ]
    },

    // ── ATA DETALHADA (4 segmentos) ──────────────────────────────────────────
    detalhada: {
      segmentos: [
        {
          id: 'cabecalho',
          nome: 'Cabeçalho e Síntese Executiva',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.2,
          maxTokens: 8000,
          pensamento: 0
        },
        {
          id: 'topicos',
          nome: 'Tópicos de Gestão e Matriz de Ação',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.3,
          maxTokens: 25000,
          pensamento: 5000
        },
        {
          id: 'log_compliance',
          nome: 'Log Operacional e Compliance',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.2,
          maxTokens: 10000,
          pensamento: 0
        },
        {
          id: 'feedback',
          nome: 'Relatório de Feedback (Privado)',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.4,
          maxTokens: 8000,
          pensamento: 0
        }
      ]
    },

    // ── ATA POR RESPONSÁVEL (3 segmentos) ────────────────────────────────────
    por_responsavel: {
      segmentos: [
        {
          id: 'cabecalho',
          nome: 'Cabeçalho',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.2,
          maxTokens: 4000,
          pensamento: 0
        },
        {
          id: 'participacao',
          nome: 'Participação por Pessoa',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.3,
          maxTokens: 18000,
          pensamento: 0
        },
        {
          id: 'consolidado',
          nome: 'Responsabilidades Consolidadas',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.2,
          maxTokens: 8000,
          pensamento: 0
        }
      ]
    },

    // ── ATA DE ALINHAMENTO RÁPIDO (1 segmento — concisa por definição) ───────
    alinhamento: {
      segmentos: [
        {
          id: 'completo',
          nome: 'Ata de Alinhamento Rápido',
          modelo: 'gemini-2.5-flash',
          temperatura: 0.2,
          maxTokens: 5000,
          pensamento: 0
        }
      ]
    }
  }
};


// ─────────────────────────────────────────────────────────────────────────────
//  PROMPT 1: TRANSCRIÇÃO (áudio inline / base64)
//  Usado quando o arquivo de áudio é enviado diretamente pelo conteúdo.
// ─────────────────────────────────────────────────────────────────────────────
function montarPromptTranscricaoInline(blocoGlossario, blocoAnchoring) {
  blocoGlossario  = blocoGlossario  || '';
  blocoAnchoring  = blocoAnchoring  || '';
  return (
    'Você é um transcritor profissional especializado em reuniões corporativas em português brasileiro.\n\n' +
    '## SUA TAREFA:\n' +
    'Transcreva o áudio da reunião de forma completa e precisa.\n\n' +
    '## INSTRUÇÕES:\n' +
    '1. Transcreva TODO o conteúdo falado no áudio\n' +
    '2. Identifique diferentes falantes quando possível (Participante 1, Participante 2, etc.)\n' +
    '3. Inclua marcações de tempo aproximadas a cada mudança significativa de tópico [MM:SS]\n' +
    '4. Preserve termos técnicos, nomes de projetos, pessoas e sistemas mencionados\n' +
    '5. Indique pausas longas com [pausa] e trechos inaudíveis com [inaudível]\n' +
    '6. Quando dois ou mais participantes falarem SIMULTANEAMENTE:\n' +
    '   - NÃO tente combinar ou interpretar os áudios sobrepostos\n' +
    '   - NÃO invente palavras que "pareçam fazer sentido" com o ruído resultante\n' +
    '   - Use exatamente: [falas simultâneas — trecho inaudível]\n' +
    '   - Retome a transcrição assim que um único falante estiver claro\n' +
    '   - Exemplo correto: [00:14] Participante 1: Com certeza, [falas simultâneas — trecho inaudível] [00:17] Kauã: ...então vamos seguir assim.\n' +
    '7. Mantenha interjeições e expressões que indiquem concordância/discordância\n\n' +
    '## FORMATO DE SAÍDA:\n' +
    'Retorne APENAS a transcrição em texto corrido, com identificação de falantes e marcações de tempo.\n' +
    'Exemplo:\n' +
    '[00:00] Participante 1: Bom dia a todos, vamos começar a reunião...\n' +
    '[00:15] Participante 2: Bom dia! Sobre o projeto X...\n\n' +
    '## IMPORTANTE:\n' +
    '- NÃO resuma ou interprete, apenas transcreva fielmente\n' +
    '- NÃO adicione formatação markdown além da identificação de falantes\n' +
    '- NÃO inclua comentários ou análises\n\n' +
    blocoGlossario +
    blocoAnchoring +
    'Transcreva o áudio a seguir:'
  );
}


// ─────────────────────────────────────────────────────────────────────────────
//  PROMPT 2: TRANSCRIÇÃO (File URI / Gemini File API)
//  Usado quando o arquivo já está hospedado na Gemini File API.
// ─────────────────────────────────────────────────────────────────────────────
function montarPromptTranscricaoFileUri(blocoGlossario, blocoAnchoring) {
  blocoGlossario  = blocoGlossario  || '';
  blocoAnchoring  = blocoAnchoring  || '';
  return (
    'Você é um transcritor profissional especializado em reuniões corporativas em português brasileiro.\n\n' +
    '## SUA TAREFA:\n' +
    'Transcreva o áudio da reunião de forma completa e precisa.\n\n' +
    '## INSTRUÇÕES:\n' +
    '1. Transcreva TODO o conteúdo falado no áudio\n' +
    '2. Identifique diferentes falantes quando possível (Participante 1, Participante 2, etc.)\n' +
    '3. Inclua marcações de tempo aproximadas a cada mudança significativa de tópico [MM:SS]\n' +
    '4. Preserve termos técnicos, nomes de projetos, pessoas e sistemas mencionados\n' +
    '5. Indique pausas longas com [pausa] e trechos inaudíveis com [inaudível]\n' +
    '6. Mantenha interjeições e expressões que indiquem concordância/discordância\n\n' +
    '## FORMATO DE SAÍDA:\n' +
    'Retorne APENAS a transcrição em texto corrido, com identificação de falantes e marcações de tempo.\n\n' +
    '## IMPORTANTE:\n' +
    '- NÃO resuma ou interprete, apenas transcreva fielmente\n' +
    '- NÃO adicione formatação markdown além da identificação de falantes\n' +
    '- NÃO inclua comentários ou análises\n\n' +
    blocoGlossario +
    blocoAnchoring +
    'Transcreva o áudio a seguir:'
  );
}


// ─────────────────────────────────────────────────────────────────────────────
//  PROMPT 3: EXTRAÇÃO DE PONTOS-CHAVE (Map-Reduce, por segmento)
//  Usado no primeiro passo do Map-Reduce para transcrições longas.
// ─────────────────────────────────────────────────────────────────────────────
function montarPromptExtracao(numSegmento, totalSegmentos, segmento) {
  return (
    'Você é um analista especializado em extrair informações estruturadas de transcrições de reunião.\n\n' +
    '## CONTEXTO:\n' +
    'Esta é a PARTE ' + numSegmento + ' de ' + totalSegmentos + ' de uma transcrição de reunião.\n\n' +
    '## SUA TAREFA:\n' +
    'Extraia TODOS os pontos importantes desta parte da transcrição, sem omitir nada.\n\n' +
    '## EXTRAIA OBRIGATORIAMENTE:\n' +
    '1. **TÓPICOS DISCUTIDOS**: Liste cada tema/assunto abordado com detalhes\n' +
    '2. **DECISÕES TOMADAS**: Qualquer decisão ou resolução mencionada\n' +
    '3. **AÇÕES DELEGADAS**: Quem ficou responsável por quê (nome → tarefa → prazo)\n' +
    '4. **PROBLEMAS/DIFICULDADES**: Bloqueios, reclamações, pendências mencionadas\n' +
    '5. **PROCESSOS DESCRITOS**: Fluxos de trabalho, procedimentos, regras explicadas\n' +
    '6. **MUDANÇAS ORGANIZACIONAIS**: Alterações em equipe, estrutura, responsabilidades\n' +
    '7. **PRAZOS E DATAS**: Qualquer menção a prazos, deadlines, datas\n' +
    '8. **NOMES E FUNÇÕES**: Pessoas mencionadas e seus papéis\n' +
    '9. **NÚMEROS E DADOS**: Valores, métricas, quantidades mencionadas\n\n' +
    '## REGRAS:\n' +
    '- Seja EXAUSTIVO, extraia TUDO, não resuma demais\n' +
    '- Mantenha detalhes específicos (nomes, números, datas)\n' +
    '- Use citações diretas quando relevante\n' +
    '- Não invente informações\n' +
    '- Organize por tópico\n\n' +
    '## TRANSCRIÇÃO (PARTE ' + numSegmento + '/' + totalSegmentos + '):\n' +
    segmento + '\n\n' +
    'Extraia todos os pontos-chave:'
  );
}


// ─────────────────────────────────────────────────────────────────────────────
//  PROMPT 4: GERAÇÃO DE ATA (LEGADO — mantido apenas para compatibilidade)
//  Use montarPromptSegmentoAta() para novos fluxos.
// ─────────────────────────────────────────────────────────────────────────────
function montarPromptAta(titulo, participantes, data, instrucaoExtra, tipoFonte, textoParaAta) {
  titulo         = titulo         || 'Não informado';
  participantes  = participantes  || 'Não informados';
  data           = data           || new Date().toLocaleDateString('pt-BR');
  instrucaoExtra = instrucaoExtra || '';
  tipoFonte      = tipoFonte      || 'TRANSCRIÇÃO';
  textoParaAta   = textoParaAta   || '';

  return (
    'Você é um secretário executivo especializado em redigir atas de reunião profissionais e COMPLETAS.\n\n' +
    '## DADOS DA REUNIÃO:\n' +
    '- **Título informado:** ' + titulo + '\n' +
    '- **Participantes informados:** ' + participantes + '\n' +
    '- **Data:** ' + data + '\n' +
    instrucaoExtra + '\n\n' +
    '## ' + tipoFonte + ':\n' +
    textoParaAta + '\n\n' +
    '## SUA TAREFA:\n' +
    'Analise o conteúdo acima e redija uma ATA DE REUNIÃO completa e profissional.\n' +
    'Você DEVE preencher TODAS as seções abaixo. NÃO deixe nenhuma seção em branco ou incompleta.\n\n' +
    '## ESTRUTURA OBRIGATÓRIA DA ATA:\n\n' +
    '### 1. CABEÇALHO\n' +
    '- Nome da Reunião (extraia do contexto ou use o título informado)\n' +
    '- Data e Hora (Início/Término estimados)\n' +
    '- Duração aproximada\n' +
    '- Local (presencial/virtual)\n' +
    '- Participantes Presentes\n' +
    '- Outros Funcionários Citados\n' +
    '- Pauta Principal\n\n' +
    '### 2. TEMA PRINCIPAL E OBJETIVOS\n' +
    'Um parágrafo resumindo o propósito central da reunião.\n\n' +
    '### 3. DETALHES DA DISCUSSÃO POR TÓPICO\n' +
    'Liste os principais pontos debatidos de forma numerada:\n' +
    '- Decisões tomadas\n' +
    '- Processos mencionados ou mapeados\n' +
    '- Mudanças organizacionais\n' +
    '- Dificuldades ou limitações declaradas\n\n' +
    '### 4. MATRIZ DE AÇÃO (PLANO DE AÇÃO)\n' +
    'Tabela com: Nº | Ação | Responsável | Prazo | Status\n' +
    'Liste Principais Pendências e Próximos Passos.\n\n' +
    '### 5. OUTROS PONTOS LEVANTADOS\n' +
    'Observações secundárias, avisos, prazos futuros.\n\n' +
    '### 6. CONSIDERAÇÕES FINAIS\n' +
    'Fechamento sintetizando o clima da reunião e principais pendências.\n\n' +
    '## FORMATO DE SAÍDA:\n' +
    'Retorne a ATA em formato Markdown bem estruturado.\n' +
    'Use tabelas markdown para a Matriz de Ação.\n' +
    'NÃO inclua JSON, apenas a ATA formatada.\n' +
    'NÃO repita separadores, traços ou qualquer padrão.\n\n' +
    '## ⚠️ REGRA CRÍTICA:\n' +
    '- PREENCHA TODAS AS 6 SEÇÕES COMPLETAMENTE\n' +
    '- Se uma seção não tem informação, escreva "Não identificado na reunião" em vez de deixar em branco\n' +
    '- Seja objetivo e profissional\n' +
    '- Não invente informações que não estejam no conteúdo fornecido\n' +
    '- Use linguagem formal e clara\n' +
    '- Destaque decisões importantes em negrito'
  );
}


// ─────────────────────────────────────────────────────────────────────────────
//  PROMPT 5: RELATÓRIO DE IDENTIFICAÇÃO DE PROJETOS E ATIVIDADES
//  Parâmetros:
//    setoresExistentes        – array de setores cadastrados
//    projetosExistentes       – array de projetos existentes
//    etapasExistentes         – array de atividades/etapas existentes
//    responsaveisExistentes   – array de responsáveis cadastrados
//    transcricaoParaRelatorio – texto da transcrição a ser analisada
// ─────────────────────────────────────────────────────────────────────────────
function montarPromptRelatorio(setoresExistentes, projetosExistentes, etapasExistentes, responsaveisExistentes, transcricaoParaRelatorio) {
  return (
    'Você é um analista de processos especializado em identificar projetos e atividades a partir de reuniões.\n\n' +
    '## ⚠️ REGRA ABSOLUTA — PROJETO ÚNICO:\n\n' +
    'Esta reunião trata de UM ÚNICO PROJETO. Toda discussão gira em torno de resolver um problema central.\n' +
    'Você DEVE identificar e gerar EXATAMENTE 1 (UM) projeto no relatório, não importa quantos temas\n' +
    'pareçam diferentes — eles são facetas do mesmo projeto central.\n\n' +
    'Se a reunião discutiu várias funcionalidades ou sub-temas, eles devem se tornar\n' +
    'ATIVIDADES do projeto único, não projetos separados.\n\n' +
    'Pergunte-se: "Qual é o projeto que resume toda esta reunião em poucas palavras?"\n' +
    'Essa é a resposta. Um projeto. Ponto final.\n\n' +
    '## ⚠️ REGRA FUNDAMENTAL — APENAS O QUE FOI DISCUTIDO:\n\n' +
    'Você DEVE identificar APENAS o que foi REALMENTE DISCUTIDO nesta reunião.\n' +
    'Não liste atividades existentes que NÃO foram mencionadas na transcrição.\n' +
    'O contexto existe APENAS para verificar se o projeto/atividade já existe.\n\n' +
    '## COMO DECIDIR O QUE INCLUIR:\n\n' +
    '### O Projeto Único:\n' +
    '1. Leia a transcrição inteira\n' +
    '2. Identifique O PROBLEMA CENTRAL que motivou esta reunião\n' +
    '3. Verifique se esse projeto já existe no contexto:\n' +
    '   - Se SIM → liste como EXISTENTE com ID real\n' +
    '   - Se NÃO → crie NOVO projeto com ID no formato proj_001\n' +
    '4. O nome do projeto deve ser autoexplicativo (ex: "Automação do Processo de Aprovação de Pagamentos da Pós-graduação")\n\n' +
    '### As Atividades:\n' +
    '- São as AÇÕES CONCRETAS que precisam ser executadas para concluir o projeto\n' +
    '- Pense: "O que precisa ser FEITO, passo a passo, para resolver o que foi discutido?"\n' +
    '- Cada atividade deve ser uma TAREFA EXECUTÁVEL com resultado mensurável\n' +
    '- Mínimo 3, máximo 12 atividades (todas as ações relevantes discutidas na reunião)\n' +
    '- Não liste atividades existentes que não foram mencionadas\n\n' +
    '### O Setor:\n' +
    '- É o setor/unidade que será BENEFICIADO, NÃO o executor\n' +
    '- Exemplo: BI cria dashboard para RH → setor é RH (beneficiado), não BI\n' +
    '- Use setores existentes quando possível\n\n' +
    '---\n\n' +
    '## 🧮 SISTEMA DE CÁLCULO DE PRIORIDADE — MUITO IMPORTANTE!\n\n' +
    'Os campos GRAVIDADE, URGENCIA, TIPO, PARA_QUEM e ESFORCO são usados para calcular\n' +
    'automaticamente o VALOR_PRIORIDADE usando a fórmula:\n\n' +
    '  VALOR_PRIORIDADE = GRAVIDADE × URGENCIA × TIPO × PARA_QUEM × ESFORCO\n\n' +
    'Use os textos EXATAMENTE como nas tabelas abaixo (o sistema faz lookup por string):\n\n' +
    '### GRAVIDADE — "O que acontece se este projeto não for feito?"\n' +
    '| Valor a usar (exato) | Peso | Quando usar |\n' +
    '|---|---|---|\n' +
    '| Crítico - Não é possível cumprir as atividades | 5 | Paralisa operações, sem alternativa |\n' +
    '| Alto - É possível cumprir parcialmente | 4 | Impacto severo, há alternativa precária |\n' +
    '| Médio - É possível mas demora muito | 3 | Impacto moderado, há alternativa razoável |\n\n' +
    '### URGENCIA — "Para quando este projeto precisa estar pronto?"\n' +
    '| Valor a usar (exato) | Peso | Quando usar |\n' +
    '|---|---|---|\n' +
    '| Imediata - Executar imediatamente | 5 | Prazo vencido ou crítico agora |\n' +
    '| Muito urgente - Prazo curto (5 dias) | 4 | Precisa ser feito em até 5 dias |\n' +
    '| Urgente - Curto prazo (10 dias) | 3 | Precisa ser feito em até 10 dias |\n' +
    '| Pouco urgente - Mais de 10 dias | 2 | Prazo confortável |\n' +
    '| Pode esperar | 1 | Sem prazo definido ou muito distante |\n\n' +
    '### TIPO — "Qual é a natureza deste projeto?"\n' +
    '| Valor a usar (exato) | Peso | Quando usar |\n' +
    '|---|---|---|\n' +
    '| Correção | 5 | Corrigir algo que está errado ou quebrado |\n' +
    '| Nova Implementação | 4 | Criar algo que ainda não existe |\n' +
    '| Melhoria | 3 | Melhorar algo que já funciona |\n\n' +
    '### PARA_QUEM — "Quem solicitou ou será beneficiado?"\n' +
    '| Valor a usar (exato) | Peso | Quando usar |\n' +
    '|---|---|---|\n' +
    '| Diretoria | 5 | Solicitado pela diretoria ou beneficia diretamente a diretoria |\n' +
    '| Demais áreas | 4 | Beneficia outras áreas operacionais |\n\n' +
    '### ESFORCO — "Quanto tempo de desenvolvimento este projeto exige?"\n' +
    '| Valor a usar (exato) | Peso | Quando usar |\n' +
    '|---|---|---|\n' +
    '| 1 turno ou menos (4 horas) | 5 | Resolve em meio período |\n' +
    '| 1 dia ou menos (8 horas) | 4 | Resolve em até um dia |\n' +
    '| uma semana (40h) | 3 | Resolve em até uma semana |\n' +
    '| mais de uma semana (40h) | 2 | Projeto longo, mais de uma semana |\n\n' +
    '### Escala de classificação do VALOR_PRIORIDADE resultante:\n' +
    '- 🔴 ALTA: resultado ≥ 2102\n' +
    '- 🟡 MÉDIA: resultado entre 1078 e 2101\n' +
    '- 🟢 BAIXA: resultado ≤ 1077\n\n' +
    'Exemplo: Correção(5) × Imediata(5) × Crítico(5) × Diretoria(5) × 1 turno(5) = 3125 → ALTA\n\n' +
    '---\n\n' +
    '## CONTEXTO DOS SETORES EXISTENTES:\n' +
    '```json\n' +
    JSON.stringify(setoresExistentes, null, 2) + '\n' +
    '```\n\n' +
    '## CONTEXTO DOS PROJETOS EXISTENTES (use APENAS para verificar se já existe):\n' +
    '```json\n' +
    JSON.stringify(projetosExistentes, null, 2) + '\n' +
    '```\n\n' +
    '## ATIVIDADES EXISTENTES (use APENAS para verificar se já existe):\n' +
    '```json\n' +
    JSON.stringify(etapasExistentes, null, 2) + '\n' +
    '```\n\n' +
    '## RESPONSÁVEIS DA EQUIPE:\n' +
    '```json\n' +
    JSON.stringify(responsaveisExistentes, null, 2) + '\n' +
    '```\n\n' +
    '## TRANSCRIÇÃO DA REUNIÃO:\n' +
    transcricaoParaRelatorio + '\n\n' +
    '## INSTRUÇÕES PARA O RELATÓRIO:\n\n' +
    '### 1. ANALISE PRIMEIRO (antes de escrever qualquer coisa):\n' +
    '- Qual é o PROBLEMA CENTRAL desta reunião? (será o nome do projeto)\n' +
    '- Qual SETOR/UNIDADE será beneficiado?\n' +
    '- Quais são as ATIVIDADES CONCRETAS discutidas para resolver este problema?\n' +
    '- Analise GRAVIDADE, URGENCIA, TIPO, PARA_QUEM e ESFORCO do projeto\n\n' +
    '### 2. REGRAS PARA O PROJETO:\n' +
    '- Crie um nome DESCRITIVO e ESPECÍFICO\n' +
    '- O projeto deve resolver o PROBLEMA CENTRAL identificado na reunião\n' +
    '- APENAS 1 projeto — sem exceções\n' +
    '- Preencha os campos de prioridade com os valores EXATOS das tabelas acima\n' +
    '- Calcule VALOR_PRIORIDADE multiplicando os 5 pesos\n' +
    '- Para itens NOVOS, use IDs no formato: proj_001, setor_001, atv_001 (SEM prefixo "NOVO_")\n' +
    '- Para itens EXISTENTES, use: EXISTENTE (ID: xxx)\n\n' +
    '### 3. REGRAS PARA ATIVIDADES:\n' +
    '- Atividades são AÇÕES para concluir o projeto\n' +
    '- Mínimo 3, máximo 12 atividades por projeto\n' +
    '- Use o ID no formato atv_001, atv_002, etc. (SEM prefixo "NOVO_")\n' +
    '- Foque em atividades PENDENTES (o que ainda precisa ser feito)\n\n' +
    '### 4. REGRAS PARA SETORES:\n' +
    '- Identifique o setor BENEFICIADO, não o executor\n' +
    '- Use setores existentes quando possível (verifique pelo nome/descrição)\n' +
    '- Só inclua setores com projetos vinculados nesta reunião\n' +
    '- Para setores novos, use ID no formato: setor_001 (SEM prefixo "NOVO_")\n\n' +
    '## ESTRUTURA DO RELATÓRIO (use este formato EXATO em Markdown):\n\n' +
    '# 📊 RELATÓRIO DE IDENTIFICAÇÕES DA REUNIÃO\n\n' +
    '## 📅 Informações Gerais\n\n' +
    '- **Data da Análise:** [data atual no formato DD/MM/AAAA]\n' +
    '- **Baseado na Transcrição:** Sim\n' +
    '- **Tema Central da Reunião:** [descreva em 1 frase curta e direta o problema central — este texto será usado no nome do arquivo]\n' +
    '- **Setor Beneficiado Principal:** [nome do setor]\n' +
    '- **Total de Projetos Identificados:** 1\n' +
    '- **Total de Atividades Identificadas:** [número]\n\n' +
    '---\n\n' +
    '## 🏢 SETOR BENEFICIADO\n\n' +
    '### [Nome do Setor]\n\n' +
    '| Campo | Valor |\n' +
    '|-------|-------|\n' +
    '| ID | [EXISTENTE (ID: xxx) OU setor_001] |\n' +
    '| NOME | [nome do setor] |\n' +
    '| DESCRICAO | [descrição do setor] |\n' +
    '| RESPONSAVEIS_IDS | [IDs dos responsáveis separados por vírgula] |\n\n' +
    '**Justificativa:** [Por que este setor é o beneficiado]\n\n' +
    '---\n\n' +
    '## 📁 PROJETOS IDENTIFICADOS\n\n' +
    '### [Nome Descritivo do Projeto Único]\n\n' +
    '| Campo | Valor |\n' +
    '|-------|-------|\n' +
    '| ID | [EXISTENTE (ID: xxx) OU proj_001] |\n' +
    '| NOME | [nome descritivo e específico] |\n' +
    '| DESCRICAO | [descrição detalhada do que o projeto resolve, baseada na discussão] |\n' +
    '| TIPO | [Correção / Nova Implementação / Melhoria — use exatamente um destes] |\n' +
    '| PARA_QUEM | [Diretoria / Demais áreas — use exatamente um destes] |\n' +
    '| STATUS | [A Fazer / Em Andamento / Concluída] |\n' +
    '| PRIORIDADE | [Alta / Média / Baixa conforme análise] |\n' +
    '| LINK | [N/A ou URL se mencionado] |\n' +
    '| GRAVIDADE | [use exatamente um dos labels da tabela de GRAVIDADE acima] |\n' +
    '| URGENCIA | [use exatamente um dos labels da tabela de URGENCIA acima] |\n' +
    '| ESFORCO | [use exatamente um dos labels da tabela de ESFORCO acima] |\n' +
    '| SETOR | [ID ou nome do setor beneficiado] |\n' +
    '| PILAR | [pilar estratégico relacionado] |\n' +
    '| RESPONSAVEIS_IDS | [IDs dos responsáveis pela execução separados por vírgula] |\n' +
    '| VALOR_PRIORIDADE | [** GRAVIDADE([peso]) × URGENCIA([peso]) × TIPO([peso]) × PARA_QUEM([peso]) × ESFORCO([peso]) = [resultado] → [Alta/Média/Baixa]] |\n' +
    '| DATA_INICIO | [DD/MM/AAAA se mencionada, senão deixar vazio] |\n' +
    '| DATA_FIM | [DD/MM/AAAA se mencionada, senão deixar vazio] |\n\n' +
    '**Cálculo de prioridade:** GRAVIDADE([peso]) × URGENCIA([peso]) × TIPO([peso]) × PARA_QUEM([peso]) × ESFORCO([peso]) = [resultado] → [Alta/Média/Baixa]\n\n' +
    '⚠️ **REGRA CRÍTICA DO VALOR_PRIORIDADE:** O número no campo VALOR_PRIORIDADE da tabela acima DEVE SER IDÊNTICO ao resultado da multiplicação mostrada na linha "Cálculo de prioridade". Se o cálculo dá 384, o campo DEVE ser 384. NUNCA coloque um valor diferente do resultado real da multiplicação.\n\n' +
    '**O que motivou este projeto na reunião:** [Cite 2-3 trechos ou situações específicas da transcrição que justificam este projeto]\n\n' +
    '---\n\n' +
    '## 📋 ATIVIDADES IDENTIFICADAS\n\n' +
    '### [Nome da Atividade — deve ser uma AÇÃO concreta]\n\n' +
    '| Campo | Valor |\n' +
    '|-------|-------|\n' +
    '| ID | [EXISTENTE (ID: xxx) OU atv_001] |\n' +
    '| PROJETO_ID | [ID do projeto acima] |\n' +
    '| RESPONSAVEIS_IDS | [IDs dos responsáveis separados por vírgula] |\n' +
    '| NOME | [nome da atividade — verbo no infinitivo, ex: "Levantar requisitos do sistema"] |\n' +
    '| O_QUE_FAZER | [instruções claras e objetivas de como executar esta atividade] |\n' +
    '| STATUS | [A Fazer / Em Andamento / Bloqueada / Concluída] |\n\n' +
    '**Justificativa:** [Por que esta atividade é necessária para o projeto]\n\n' +
    '[repita o bloco acima para cada atividade identificada]\n\n' +
    '---\n\n' +
    '## ⚠️ CHECKLIST FINAL (verifique antes de entregar):\n' +
    '- [ ] O relatório tem EXATAMENTE 1 projeto?\n' +
    '- [ ] O projeto foi realmente discutido na transcrição?\n' +
    '- [ ] As atividades são AÇÕES CONCRETAS (não descrições)?\n' +
    '- [ ] O setor identificado é o BENEFICIADO (não o executor)?\n' +
    '- [ ] O projeto tem entre 3-12 atividades relevantes?\n' +
    '- [ ] Os campos GRAVIDADE, URGENCIA, TIPO, PARA_QUEM e ESFORCO usam os textos EXATOS das tabelas?\n' +
    '- [ ] O VALOR_PRIORIDADE foi calculado corretamente (multiplicação dos 5 pesos)?\n' +
    '- [ ] O VALOR_PRIORIDADE na tabela é IDÊNTICO ao resultado mostrado no "Cálculo de prioridade"?\n' +
    '- [ ] Os IDs de itens novos usam formato limpo (proj_001, atv_001, setor_001) SEM prefixo "NOVO_"?\n' +
    '- [ ] O campo "Tema Central da Reunião" é uma frase curta e direta?\n' +
    '- [ ] Os campos DATA_INICIO e DATA_FIM foram preenchidos quando mencionados na transcrição?\n\n' +
    'Gere o relatório completo em Markdown:'
  );
}


// ─────────────────────────────────────────────────────────────────────────────
//  PROMPT 6: SEGMENTO DE ATA POR ESTILO
//  Gera o prompt para UM segmento específico de um estilo de ata.
//  Cada estilo é dividido em segmentos (ver CONFIG_PROMPTS_REUNIAO.ESTILOS),
//  e cada segmento é processado em uma chamada separada ao Gemini,
//  com configuração própria de modelo e tokens.
//
//  Parâmetros:
//    estilo         – 'executiva' | 'detalhada' | 'por_responsavel' | 'alinhamento'
//    segmentoId     – id do segmento conforme definido em ESTILOS[estilo].segmentos
//    titulo         – título da reunião
//    participantes  – lista de participantes (texto livre)
//    data           – data da reunião (string formatada)
//    instrucaoExtra – instruções adicionais opcionais
//    transcricao    – transcrição completa da reunião
// ─────────────────────────────────────────────────────────────────────────────
function montarPromptSegmentoAta(estilo, segmentoId, titulo, participantes, data, instrucaoExtra, transcricao) {
  titulo         = titulo         || 'Não informado';
  participantes  = participantes  || 'Não informados';
  data           = data           || new Date().toLocaleDateString('pt-BR');
  instrucaoExtra = instrucaoExtra ? '\n**Instruções adicionais:** ' + instrucaoExtra + '\n' : '';
  transcricao    = transcricao    || '';

  var cabecalhoComum =
    '## DADOS DA REUNIÃO:\n' +
    '- **Título:** ' + titulo + '\n' +
    '- **Participantes:** ' + participantes + '\n' +
    '- **Data:** ' + data + '\n' +
    instrucaoExtra + '\n' +
    '## TRANSCRIÇÃO COMPLETA:\n' +
    transcricao + '\n\n';


  // ══════════════════════════════════════════════════════════════════
  //  ESTILO: EXECUTIVA
  // ══════════════════════════════════════════════════════════════════

  if (estilo === 'executiva') {

    if (segmentoId === 'decisoes') {
      return (
        'Você é um secretário executivo especializado em atas de alto nível para diretores e gestores.\n\n' +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS as seções 1 e 2 da Ata Executiva. Seja EXAUSTIVO — extraia cada decisão sem omitir nenhuma.\n\n' +
        '### 1. CABEÇALHO\n' +
        '- Reunião, Data, Participantes Presentes, Pauta\n\n' +
        '### 2. DECISÕES TOMADAS\n' +
        'Lista numerada com CADA DECISÃO firmada, quem decidiu e o racional resumido em 1-2 frases.\n' +
        'Se houver 10 decisões, liste 10. Se houver 20, liste 20. Não omita nenhuma.\n\n' +
        '⚠️ Retorne APENAS essas 2 seções em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }

    if (segmentoId === 'encaminhamentos') {
      return (
        'Você é um secretário executivo especializado em atas de alto nível para diretores e gestores.\n\n' +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS as seções 3 e 4 da Ata Executiva. Seja EXAUSTIVO — não omita nenhum encaminhamento.\n\n' +
        '### 3. ENCAMINHAMENTOS\n' +
        'Tabela completa: Nº | Ação | Responsável | Prazo\n' +
        'Liste TODOS os encaminhamentos, tarefas e compromissos assumidos na reunião.\n\n' +
        '### 4. PRÓXIMA REUNIÃO / PRÓXIMOS PASSOS\n' +
        'Data ou prazo da próxima revisão, se mencionado. Outros próximos passos estratégicos.\n\n' +
        '⚠️ Retorne APENAS essas 2 seções em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }
  }


  // ══════════════════════════════════════════════════════════════════
  //  ESTILO: DETALHADA
  // ══════════════════════════════════════════════════════════════════

  if (estilo === 'detalhada') {
    var personaDetalhada =
      'Aja como uma Assistente Executiva Sênior e Especialista em Governança Corporativa. ' +
      'Sua missão é processar transcrições de reuniões de um grupo educacional e veterinário em Fortaleza.\n\n' +
      '**Dicionário de Nomenclatura** — corrija erros fonéticos usando estes nomes:\n' +
      '- Pessoas: Murilo, Adeline, Ana, Anne, Bráulio, Dr. André, Cristina, Cymara, Dânia, David, ' +
      'Dayanne, Diana, Diógenes, Elisa, Fernando, Filipe, Gil, Ítalo, Jaiane, Josi, Kauã, Luzia, ' +
      'Dr. Marcos, Mayra, Monalisa, Melissa, Priscila, Ramon, Raquel, Rayanne, Rayssa, Rebecca, ' +
      'Renan, Sânia, Saskia, Silvaney, Thaís, Torquato, Werysson, Wesley.\n' +
      '- Unidades: Christus, Unichristus, CESIU, LabVet, CVU, CES, LEAC.\n\n';

    if (segmentoId === 'cabecalho') {
      return (
        personaDetalhada +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS as seções 1 e 2 da Ata Detalhada. Extraia ao máximo os dados do cabeçalho e síntese.\n\n' +
        '### 1. CABEÇALHO\n' +
        '- Nome da Reunião, Data/Hora de início e término, Local, Participantes presentes.\n\n' +
        '### 2. PAUTA E SÍNTESE EXECUTIVA\n' +
        'Resumo executivo em 4-6 linhas cobrindo os temas centrais discutidos.\n\n' +
        '⚠️ Retorne APENAS essas 2 seções em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }

    if (segmentoId === 'topicos') {
      return (
        personaDetalhada +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS as seções 3 e 4 da Ata Detalhada. Seja EXTREMAMENTE EXAUSTIVO:\n' +
        'não omita nenhum assunto, nenhuma decisão, nenhum dado, nenhum número mencionado.\n\n' +
        '### 3. TÓPICOS DE GESTÃO\n' +
        'Para CADA assunto estratégico discutido, crie um bloco:\n' +
        '**[Título do Tópico]**\n' +
        '- Pontos-chave: bullet points com todos os dados, argumentos e contextos mencionados.\n' +
        '- Decisões Tomadas: o que foi formalizado ou acordado neste tópico.\n\n' +
        '### 4. MATRIZ DE AÇÃO\n' +
        '| Ação | Responsável | Prazo | Unidade | Status |\n' +
        '| :--- | :--- | :--- | :--- | :--- |\n' +
        'Liste TODAS as ações, tarefas e responsabilidades mencionadas. Não omita nenhuma.\n\n' +
        '⚠️ Retorne APENAS essas 2 seções em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }

    if (segmentoId === 'log_compliance') {
      return (
        personaDetalhada +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS as seções 5 e 6 da Ata Detalhada. ' +
        'Faça uma varredura PROFUNDA para extrair TODOS os itens técnicos e riscos, ' +
        'mesmo os mencionados de passagem ou informalmente.\n\n' +
        '### 5. LOG OPERACIONAL E INFRAESTRUTURA\n' +
        'Liste TODOS os itens técnicos, reparos e logísticas mencionados:\n' +
        '- Entraves Técnicos e Manutenção: infraestrutura (pisos, grades, vazamentos), TI (ataques, bugs), insumos.\n' +
        '- Logística de Bastidor: detalhes operacionais que afetam o dia a dia (ex: animais, aulas, escalas).\n' +
        '- Incidentes e Riscos: crises de imagem, problemas com fornecedores, falhas recorrentes de processos.\n\n' +
        '### 6. FILTRO DE COMPLIANCE E RISCO (Anexo Consultivo)\n' +
        'Identifique falas ou decisões com risco jurídico (Trabalhista, LGPD, Ético).\n' +
        'Descreva de forma técnica, neutra e puramente factual como alertas para a diretoria.\n\n' +
        '⚠️ Retorne APENAS essas 2 seções em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }

    if (segmentoId === 'feedback') {
      return (
        personaDetalhada +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS a seção 7 da Ata Detalhada: o Relatório de Feedback privado para Murilo.\n\n' +
        '### 7. RELATÓRIO DE FEEDBACK (Privado para Murilo)\n' +
        'Analise a dinâmica da reunião sob a ótica de liderança:\n' +
        '- **Clareza e Direcionamento:** O foco foi mantido ou houve dispersão em detalhes técnicos?\n' +
        '- **Equilíbrio de Voz:** Murilo deu espaço ou monopolizou? A equipe foi proativa ou passiva?\n' +
        '- **Gestão de Conflitos e Tom:** Como reagiu a tensões? O tom foi assertivo, autoritário ou colaborativo?\n' +
        '- **Oportunidades de Melhoria:** Aponte 2 comportamentos específicos para ajustar a eficiência da gestão.\n' +
        '- **Uso de IA Mencionado:** Se houver menção a IAs (ex: Gemini), descreva a metodologia e dificuldades citadas.\n\n' +
        '⚠️ Retorne APENAS esta seção em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }
  }


  // ══════════════════════════════════════════════════════════════════
  //  ESTILO: POR RESPONSÁVEL
  // ══════════════════════════════════════════════════════════════════

  if (estilo === 'por_responsavel') {

    if (segmentoId === 'cabecalho') {
      return (
        'Você é um secretário especializado em atas organizadas por pessoa.\n\n' +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS a seção 1 da Ata por Responsável.\n\n' +
        '### 1. CABEÇALHO\n' +
        '- Reunião, Data, Participantes Presentes, Pauta, Duração estimada.\n\n' +
        '⚠️ Retorne APENAS esta seção em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }

    if (segmentoId === 'participacao') {
      return (
        'Você é um secretário especializado em atas organizadas por pessoa.\n\n' +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS a seção 2 da Ata por Responsável. ' +
        'Seja EXAUSTIVO — capture TUDO que cada pessoa disse, decidiu ou assumiu. ' +
        'Não agrupe pessoas diferentes. Não resuma demais.\n\n' +
        '### 2. PARTICIPAÇÃO POR PESSOA\n' +
        'Para CADA participante que contribuiu ativamente:\n\n' +
        '**[Nome do Participante]**\n' +
        '- Posicionamentos e declarações relevantes (com detalhes, não apenas "concordou")\n' +
        '- Decisões que tomou ou co-participou\n' +
        '- Tarefas ou responsabilidades assumidas (com prazo, se mencionado)\n' +
        '- Questões ou dúvidas levantadas\n\n' +
        'Inclua apenas participantes com contribuição identificável.\n\n' +
        '⚠️ Retorne APENAS esta seção em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }

    if (segmentoId === 'consolidado') {
      return (
        'Você é um secretário especializado em atas organizadas por pessoa.\n\n' +
        cabecalhoComum +
        '## SUA TAREFA:\n' +
        'Gere APENAS as seções 3 e 4 da Ata por Responsável. ' +
        'Seja COMPLETO — liste TODAS as pendências e responsabilidades, incluindo as implícitas.\n\n' +
        '### 3. RESPONSABILIDADES CONSOLIDADAS\n' +
        'Tabela completa: Responsável | Tarefa | Prazo | Status\n' +
        'Liste TODAS as responsabilidades atribuídas na reunião.\n\n' +
        '### 4. PONTOS SEM RESPONSÁVEL DEFINIDO\n' +
        'Pendências, problemas ou questões levantadas sem atribuição clara de responsável.\n\n' +
        '⚠️ Retorne APENAS essas 2 seções em Markdown. Não adicione introdução, conclusão ou outras seções.'
      );
    }
  }


  // ══════════════════════════════════════════════════════════════════
  //  ESTILO: ALINHAMENTO RÁPIDO
  // ══════════════════════════════════════════════════════════════════

  if (estilo === 'alinhamento') {

    if (segmentoId === 'completo') {
      return (
        'Você é um secretário especializado em atas de alinhamento rápido para dailies, syncs e reuniões informais.\n\n' +
        cabecalhoComum +
        '## SUA TAREFA — ATA DE ALINHAMENTO RÁPIDO:\n' +
        'Redija uma ata em TÓPICOS CURTOS. Foco em: o que foi alinhado, próximos passos e prazos.\n\n' +
        '## ESTRUTURA OBRIGATÓRIA:\n\n' +
        '### 1. CABEÇALHO\n' +
        '- Reunião, Data, Participantes, Duração estimada\n\n' +
        '### 2. O QUE FOI ALINHADO\n' +
        'Lista de bullet points com cada alinhamento ou acordo. Máximo 2 linhas por item.\n\n' +
        '### 3. PRÓXIMOS PASSOS\n' +
        'Lista: ✅ [Ação] — [Responsável] — [Prazo]\n\n' +
        '### 4. PONTOS DE ATENÇÃO\n' +
        'Riscos, bloqueios ou dúvidas levantadas (se houver). Máximo 3 itens.\n\n' +
        '## REGRAS:\n' +
        '- Seja MUITO conciso. A ata inteira não deve ultrapassar 1 página.\n' +
        '- Use linguagem direta.\n' +
        '- Se não há próximos passos definidos, escreva "Nenhum próximo passo definido".\n\n' +
        'Gere a Ata de Alinhamento Rápido em formato Markdown:'
      );
    }
  }


  // Segmento desconhecido — retorna string vazia
  Logger.log('[montarPromptSegmentoAta] Segmento desconhecido: estilo=' + estilo + ', segmentoId=' + segmentoId);
  return '';
}

