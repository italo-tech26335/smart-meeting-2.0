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
    modelo: 'gemini-2.5-pro',
    temperatura: 0.0,
    maxTokens: 200000,
    pensamento: 1
  },
  ATA: {
    modelo: 'gemini-2.5-flash',
    temperatura: 0.3,
    maxTokens: 20000,
    pensamento: 0,
    maxTokensPorSecao: 20000
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
//  PROMPT 4: GERAÇÃO DE ATA
//  Parâmetros:
//    titulo         – título da reunião
//    participantes  – lista de participantes (texto livre)
//    data           – data da reunião (string formatada)
//    instrucaoExtra – instruções adicionais opcionais (ex: foco em decisões)
//    tipoFonte      – rótulo da fonte (ex: "TRANSCRIÇÃO" ou "PONTOS EXTRAÍDOS")
//    textoParaAta   – conteúdo a ser transformado em ATA
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
