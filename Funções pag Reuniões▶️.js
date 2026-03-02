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

/**
 * Interpreta uma célula de IDs de responsáveis,
 * suportando dois formatos:
 *   • Puro:        resp_Italo
 *   • JSON array:  ["resp_Guilherme"]  ou  ["resp_A","resp_B"]
 * Retorna sempre um Array de strings limpas.
 */
function parsearIdsColuna(valorCelula) {
  if (!valorCelula) return [];
  var raw = valorCelula.toString().trim();
  if (!raw) return [];

  // Tenta JSON (cobre ["resp_X"] e ["resp_A","resp_B"])
  if (raw.charAt(0) === '[') {
    try {
      var parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        return parsed
          .map(function(id) { return id ? id.toString().trim() : ''; })
          .filter(function(id) { return id !== ''; });
      }
    } catch (e) { /* falha silenciosa, cai no split */ }
  }

  // Fallback: string simples separada por vírgula
  return raw.split(',')
    .map(function(id) { return id.trim(); })
    .filter(function(id) { return id !== ''; });
}

function etapa1a_IniciarSessoesUpload(dados) {
  const logs = [];
  const log = function(tipo, msg) {
    logs.push({ tipo: tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log('[etapa1a][' + tipo + '] ' + msg);
  };

  try {
    log('INFO', '☁️ [ETAPA 1a] Iniciando sessões de upload...');
    const chaveApi = obterChaveGeminiProjeto();
    if (!chaveApi) throw new Error('Chave API do Gemini não configurada.');

    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const tipoMime = dados.tipoMime || 'audio/webm';
    var tituloLimpo = limparTituloParaNomeArquivo(dados.titulo || 'Reuniao');

    // ── 1. Contar chunks e calcular tamanho binário total ──
    const arquivosChunks = [];
    let idx = 0;
    while (true) {
      const it = pasta.getFilesByName('chunk_' + dados.idUpload + '_' + idx);
      if (!it.hasNext()) break;
      arquivosChunks.push(it.next());
      idx++;
    }
    if (arquivosChunks.length === 0) {
      throw new Error('Nenhum chunk encontrado para idUpload: ' + dados.idUpload);
    }
    log('INFO', '  📦 ' + arquivosChunks.length + ' chunks encontrados no Drive');

    let tamBinTotal = 0;
    for (let i = 0; i < arquivosChunks.length; i++) {
      if (i < arquivosChunks.length - 1) {
        var tamBase64 = arquivosChunks[i].getSize();
        tamBinTotal += Math.floor(tamBase64 * 3 / 4);
      } else {
        var ultimoBase64 = arquivosChunks[i].getBlob().getDataAsString();
        var ultimoBytes = Utilities.base64Decode(ultimoBase64);
        tamBinTotal += ultimoBytes.length;
      }
    }
    log('SUCESSO', '✅ Tamanho binário total: ' + (tamBinTotal / 1024 / 1024).toFixed(2) + ' MB');

    // ✅ CORRIGIDO: Declara extensao, nomeAudioReal e nomeArquivoGemini ANTES de usá-los
    var extensao = obterExtensaoDoMime(tipoMime);
    var timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMdd_HHmmss');
    var nomeAudioReal = tituloLimpo + extensao;
    var nomeArquivoGemini = tituloLimpo + '_' + timestamp; // display_name para o Gemini

    // ── 2. Iniciar sessão resumível no Gemini ──
    log('INFO', '📤 Iniciando sessão Gemini...');

    var respInicioGemini = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/upload/v1beta/files?key=' + chaveApi, {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'X-Goog-Upload-Protocol': 'resumable',
          'X-Goog-Upload-Command': 'start',
          'X-Goog-Upload-Header-Content-Length': tamBinTotal.toString(),
          'X-Goog-Upload-Header-Content-Type': tipoMime
        },
        payload: JSON.stringify({ file: { display_name: nomeArquivoGemini } }),
        muteHttpExceptions: true
      }
    );

    if (respInicioGemini.getResponseCode() !== 200) {
      throw new Error('Falha ao iniciar sessão Gemini: HTTP ' +
        respInicioGemini.getResponseCode() + ' → ' +
        respInicioGemini.getContentText().substring(0, 200));
    }

    var headersGemini = respInicioGemini.getHeaders();
    var geminiUrl = headersGemini['x-goog-upload-url'] || headersGemini['X-Goog-Upload-URL'];
    if (!geminiUrl) throw new Error('URL de upload Gemini não retornada nos headers');
    log('SUCESSO', '✅ Sessão Gemini criada');

    // ── 3. Iniciar sessão resumível no Drive ──
    var tokenOAuth = ScriptApp.getOAuthToken();
    var driveUrl = null;

    try {
      var respInicioDrive = UrlFetchApp.fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
          method: 'POST',
          headers: {
            'Authorization': 'Bearer ' + tokenOAuth,
            'Content-Type': 'application/json',
            'X-Upload-Content-Type': tipoMime,
            'X-Upload-Content-Length': tamBinTotal.toString()
          },
          payload: JSON.stringify({ name: nomeAudioReal, parents: [ID_PASTA_DRIVE_REUNIOES] }),
          muteHttpExceptions: true
        }
      );

      if (respInicioDrive.getResponseCode() === 200) {
        var hdDrive = respInicioDrive.getHeaders();
        driveUrl = hdDrive['Location'] || hdDrive['location'];
        log(driveUrl ? 'SUCESSO' : 'ALERTA',
          driveUrl ? '✅ Sessão Drive criada' : '⚠️ URL Drive não retornada');
      } else {
        log('ALERTA', '⚠️ Sessão Drive falhou (HTTP ' + respInicioDrive.getResponseCode() + ')');
      }
    } catch (eDrive) {
      log('ALERTA', '⚠️ Erro ao criar sessão Drive: ' + eDrive.message);
    }

    log('SUCESSO', '✅ ETAPA 1a CONCLUÍDA');

    return {
      sucesso: true,
      logs: logs,
      geminiUrl: geminiUrl,
      driveUrl: driveUrl || '',
      totalBytes: tamBinTotal,
      numChunks: arquivosChunks.length,
      nomeAudio: nomeAudioReal
    };

  } catch (erro) {
    log('ERRO', '❌ ' + erro.message);
    return { sucesso: false, logs: logs, mensagem: erro.message };
  }
}

function etapa1b_EnviarLoteParaGemini(dados) {
  const logs = [];
  const log = function(tipo, msg) {
    logs.push({ tipo: tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log('[etapa1b][' + tipo + '] ' + msg);
  };

  try {
    var GRANULARIDADE_GEMINI = 8 * 1024 * 1024; // 8MB exatos — NÃO alterar
    var pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    var offsetGemini = dados.offsetGemini || 0;
    var offsetDrive = dados.offsetDrive || 0; // mantido para compatibilidade
    var pacotesEnviados = 0;
    var tempoInicio = Date.now();

    log('INFO', '📤 [ETAPA 1b] Lote: chunks ' +
      dados.chunkInicio + ' a ' + dados.chunkFim +
      ' | offset: ' + (offsetGemini / 1024 / 1024).toFixed(1) + ' MB' +
      ' | último: ' + (dados.ehUltimoLote ? 'SIM' : 'NÃO'));

    // ── 1. Ler buffer residual do lote anterior (se existir) ──
    var bufferLista = [];
    var tamBufferTotal = 0;

    var nomeBufferResidual = 'buffer_residual_' + dados.idUpload;
    var itBuffer = pasta.getFilesByName(nomeBufferResidual);
    if (itBuffer.hasNext()) {
      var arquivoBuffer = itBuffer.next();
      var bufferBase64 = arquivoBuffer.getBlob().getDataAsString();
      if (bufferBase64 && bufferBase64.length > 0) {
        var bytesResidual = Utilities.base64Decode(bufferBase64);
        bufferLista.push(bytesResidual);
        tamBufferTotal += bytesResidual.length;
        log('INFO', '  📦 Buffer residual: ' + (bytesResidual.length / 1024 / 1024).toFixed(2) + ' MB');
      }
      arquivoBuffer.setTrashed(true);
    }

    // ── 2. Ler e decodificar chunks do lote ──
    for (var i = dados.chunkInicio; i <= dados.chunkFim; i++) {
      var itChunk = pasta.getFilesByName('chunk_' + dados.idUpload + '_' + i);
      if (!itChunk.hasNext()) {
        log('ALERTA', '  ⚠️ Chunk ' + i + ' não encontrado, pulando');
        continue;
      }

      var arquivoChunk = itChunk.next();
      var chunkBase64 = arquivoChunk.getBlob().getDataAsString();
      var bytesChunk = Utilities.base64Decode(chunkBase64);

      bufferLista.push(bytesChunk);
      tamBufferTotal += bytesChunk.length;

      // Apaga chunk processado imediatamente (libera espaço no Drive)
      try { arquivoChunk.setTrashed(true); } catch (e) { /* não crítico */ }

      log('INFO', '  🔄 Chunk ' + i + ' (' +
        (bytesChunk.length / 1024 / 1024).toFixed(1) + ' MB) | buffer: ' +
        (tamBufferTotal / 1024 / 1024).toFixed(1) + ' MB');

      // ── 3. Enquanto buffer >= 8MB, enviar pacote ao Gemini ──
      while (tamBufferTotal >= GRANULARIDADE_GEMINI) {
        // Se é o último lote e buffer tem EXATAMENTE 8MB ou um pouco mais,
        // verifica se restam mais chunks para não enviar 'finalize' cedo demais
        var ehRealmenteUltimo = dados.ehUltimoLote && (i === dados.chunkFim) && (tamBufferTotal < GRANULARIDADE_GEMINI * 2);
        if (ehRealmenteUltimo) break; // deixa pro bloco final abaixo

        pacotesEnviados++;
        var pacote8MB = extrairPrimeirosNBytesDoBuffer(bufferLista, GRANULARIDADE_GEMINI);
        tamBufferTotal -= GRANULARIDADE_GEMINI;

        log('INFO', '  📤 Pacote ' + pacotesEnviados + ' (8 MB) → Gemini...');
        var tPacote = Date.now();

        var respGemini = UrlFetchApp.fetch(dados.geminiUrl, {
          method: 'post',
          contentType: 'application/octet-stream',
          headers: {
            'X-Goog-Upload-Command': 'upload',
            'X-Goog-Upload-Offset': offsetGemini.toString()
          },
          payload: pacote8MB,
          muteHttpExceptions: true
        });

        if (respGemini.getResponseCode() !== 200) {
          throw new Error('Erro Gemini pacote ' + pacotesEnviados +
            ': HTTP ' + respGemini.getResponseCode() +
            ' → ' + respGemini.getContentText().substring(0, 200));
        }

        offsetGemini += GRANULARIDADE_GEMINI;
        log('SUCESSO', '  ✅ Pacote ' + pacotesEnviados + ' OK (' +
          ((Date.now() - tPacote) / 1000).toFixed(1) + 's)');

        // Verificação de segurança: se está chegando perto do limite de 5 min,
        // salva buffer e retorna para o client chamar outro lote
        var tempoDecorrido = (Date.now() - tempoInicio) / 1000;
        if (tempoDecorrido > 240 && !dados.ehUltimoLote) { // 4 minutos
          log('ALERTA', '⏰ Tempo de segurança atingido (' +
            tempoDecorrido.toFixed(0) + 's). Salvando buffer e retornando...');

          // Salvar buffer residual
          if (tamBufferTotal > 0) {
            var bufferJs = combinarBufferCompleto(bufferLista);
            var bufferB64 = Utilities.base64Encode(bufferJs);
            pasta.createFile(nomeBufferResidual, bufferB64, MimeType.PLAIN_TEXT);
            log('INFO', '  💾 Buffer residual salvo: ' + (tamBufferTotal / 1024 / 1024).toFixed(2) + ' MB');
          }

          return {
            sucesso: true,
            logs: logs,
            offsetGemini: offsetGemini,
            offsetDrive: offsetDrive,
            pacotesEnviados: pacotesEnviados,
            finalizado: false,
            // Informa ao client quais chunks AINDA não foram processados
            proximoChunkInicio: i + 1,
            interrompidoPorTempo: true
          };
        }
      }
    }

    // ── 4. Se é o último lote: enviar TUDO que resta com 'finalize' ──
    if (dados.ehUltimoLote && tamBufferTotal > 0) {

      // Primeiro: enviar todos os pacotes completos de 8MB como 'upload'
      while (tamBufferTotal > GRANULARIDADE_GEMINI) {
        pacotesEnviados++;
        var pacoteInterm = extrairPrimeirosNBytesDoBuffer(bufferLista, GRANULARIDADE_GEMINI);
        tamBufferTotal -= GRANULARIDADE_GEMINI;

        log('INFO', '  📤 Pacote intermediário ' + pacotesEnviados + ' (8 MB)...');

        var respInterm = UrlFetchApp.fetch(dados.geminiUrl, {
          method: 'post',
          contentType: 'application/octet-stream',
          headers: {
            'X-Goog-Upload-Command': 'upload',
            'X-Goog-Upload-Offset': offsetGemini.toString()
          },
          payload: pacoteInterm,
          muteHttpExceptions: true
        });

        if (respInterm.getResponseCode() !== 200) {
          throw new Error('Erro Gemini intermediário: HTTP ' + respInterm.getResponseCode());
        }
        offsetGemini += GRANULARIDADE_GEMINI;
        log('SUCESSO', '  ✅ Pacote intermediário ' + pacotesEnviados + ' OK');
      }

      // Agora: pacote final (qualquer tamanho, com 'finalize')
      pacotesEnviados++;
      var pacoteFinal = combinarBufferCompleto(bufferLista);
      bufferLista = [];
      tamBufferTotal = 0;

      log('INFO', '  📤 Pacote FINAL ' + pacotesEnviados + ' (' +
        (pacoteFinal.length / 1024 / 1024).toFixed(2) + ' MB, finalize)...');

      var respFinal = UrlFetchApp.fetch(dados.geminiUrl, {
        method: 'post',
        contentType: 'application/octet-stream',
        headers: {
          'X-Goog-Upload-Command': 'upload, finalize',
          'X-Goog-Upload-Offset': offsetGemini.toString()
        },
        payload: pacoteFinal,
        muteHttpExceptions: true
      });

      if (respFinal.getResponseCode() !== 200) {
        throw new Error('Erro Gemini finalize: HTTP ' + respFinal.getResponseCode() +
          ' → ' + respFinal.getContentText().substring(0, 300));
      }

      offsetGemini += pacoteFinal.length;
      var respostaGemini = JSON.parse(respFinal.getContentText());

      var tempoTotal = ((Date.now() - tempoInicio) / 1000).toFixed(1);
      log('SUCESSO', '✅ Upload Gemini FINALIZADO em ' + tempoTotal + 's! (' +
        pacotesEnviados + ' pacotes)');

      return {
        sucesso: true,
        logs: logs,
        offsetGemini: offsetGemini,
        offsetDrive: offsetDrive,
        pacotesEnviados: pacotesEnviados,
        finalizado: true,
        geminiFileUri: respostaGemini.file.uri,
        geminiFileName: respostaGemini.file.name
      };
    }

    // ── 5. Não é último lote: salvar buffer residual no Drive ──
    if (tamBufferTotal > 0) {
      var bufferCompletoJs = combinarBufferCompleto(bufferLista);
      var bufferResidualBase64 = Utilities.base64Encode(bufferCompletoJs);
      pasta.createFile(nomeBufferResidual, bufferResidualBase64, MimeType.PLAIN_TEXT);
      log('INFO', '  💾 Buffer residual salvo: ' + (tamBufferTotal / 1024 / 1024).toFixed(2) + ' MB');
    }

    var tempoLote = ((Date.now() - tempoInicio) / 1000).toFixed(1);
    log('SUCESSO', '✅ Lote concluído em ' + tempoLote + 's (' + pacotesEnviados + ' pacotes)');

    return {
      sucesso: true,
      logs: logs,
      offsetGemini: offsetGemini,
      offsetDrive: offsetDrive,
      pacotesEnviados: pacotesEnviados,
      finalizado: false,
      interrompidoPorTempo: false
    };

  } catch (erro) {
    log('ERRO', '❌ ' + erro.message);
    Logger.log('[STACK etapa1b] ' + (erro.stack || erro.toString()));
    return { sucesso: false, logs: logs, mensagem: erro.message };
  }
}

function etapa1c_AguardarProcessamentoGemini(dados) {
  const logs = [];
  const log = function(tipo, msg) {
    logs.push({ tipo: tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log('[etapa1c][' + tipo + '] ' + msg);
  };

  try {
    log('INFO', '⏳ [ETAPA 1c] Aguardando Gemini processar o arquivo...');
    log('INFO', '  📌 FileName: ' + dados.geminiFileName);

    var chaveApi = obterChaveGeminiProjeto();

    // ── 1. Aguardar processamento (polling) ──
    aguardarProcessamentoArquivoGemini(dados.geminiFileName, chaveApi);
    log('SUCESSO', '✅ Arquivo Gemini processado e ATIVO!');

    // ── 2. Salvar áudio no Drive ──
    // Como removemos o upload ao Drive da etapa1b para ganhar velocidade,
    // agora salvamos o áudio reconstruindo dos chunks restantes.
    // Se os chunks já foram apagados, usamos o que temos.
    var linkAudioReal = '';
    var idArquivoDrive = '';

    try {
      log('INFO', '💾 Salvando áudio no Drive...');
      var pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);

      // Tenta reconstruir o áudio dos chunks que ainda existem
      // (os chunks são movidos para lixeira, mas ainda acessíveis no curto prazo)
      // Estratégia alternativa: como temos o fileUri do Gemini, podemos
      // simplesmente registrar o link do Gemini e pular o Drive
      // para arquivos muito grandes.

      // Verificar se algum chunk ainda existe (pode ter sido apagado na etapa1b)
      var itTeste = pasta.getFilesByName('chunk_' + dados.idUpload + '_0');
      var chunksDisponiveis = itTeste.hasNext();

      if (!chunksDisponiveis) {
        // Chunks já apagados — cria um arquivo placeholder com metadados
        log('INFO', '  📝 Chunks já processados. Criando registro no Drive...');
        var conteudoMeta = 'Áudio processado pelo Smart Meeting\n' +
          'Data: ' + new Date().toLocaleString('pt-BR') + '\n' +
          'Gemini FileURI: ' + dados.geminiFileUri + '\n' +
          'Tamanho original: ' + ((dados.totalBytes || 0) / 1024 / 1024).toFixed(2) + ' MB\n' +
          'Nome: ' + (dados.nomeAudio || 'audio') + '\n\n' +
          'O áudio original foi enviado diretamente ao Gemini para processamento.\n' +
          'Para ouvir, use o áudio original que foi carregado na reunião.';

        var nomeMetaArquivo = (dados.nomeAudio || 'Reuniao_audio') + '_info.txt';
        var arquivoMeta = pasta.createFile(nomeMetaArquivo, conteudoMeta, MimeType.PLAIN_TEXT);
        arquivoMeta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        linkAudioReal = arquivoMeta.getUrl();
        log('SUCESSO', '✅ Registro criado no Drive');
      }
    } catch (eDrive) {
      log('ALERTA', '⚠️ Erro ao salvar no Drive: ' + eDrive.message);
    }

    // ── 3. Limpar buffer residual e outros temporários ──
    if (dados.idUpload) {
      try {
        var pastaLimpar = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
        var itLimpar = pastaLimpar.getFilesByName('buffer_residual_' + dados.idUpload);
        while (itLimpar.hasNext()) {
          itLimpar.next().setTrashed(true);
        }
      } catch (e) { /* não crítico */ }
    }

    log('SUCESSO', '✅ ETAPA 1c CONCLUÍDA');

    return {
      sucesso: true,
      logs: logs,
      fileUri: dados.geminiFileUri,
      fileName: dados.geminiFileName,
      linkAudio: linkAudioReal
    };

  } catch (erro) {
    log('ERRO', '❌ ' + erro.message);
    return { sucesso: false, logs: logs, mensagem: erro.message };
  }
}

function obterUrlBaseWebApp() {
  return ScriptApp.getService().getUrl();
}

const CONFIGURACAO_REUNIOES = {
  TAMANHO_MAXIMO_AUDIO_MB: 1000,
  TEMPO_TIMEOUT_GEMINI_MS: 800000,
  FORMATOS_AUDIO_ACEITOS: ['audio/webm', 'audio/mp3', 'audio/mpeg', 'audio/wav', 'audio/ogg', 'audio/mp4', 'audio/x-m4a'],
  EXTENSAO_PADRAO: '.webm',
  LIMITE_CHARS_MAP_REDUCE: 70000
};

const TAMANHO_LOTE_UPLOAD = 3;

function extrairTextoRespostaGemini(respostaJson) {
  if (!respostaJson.candidates || respostaJson.candidates.length === 0) {
    throw new Error('Gemini não retornou resposta válida (sem candidates)');
  }

  const partes = respostaJson.candidates[0].content.parts;
  if (!partes || partes.length === 0) {
    throw new Error('Gemini não retornou resposta válida (sem parts)');
  }

  // Filtra partes que NÃO são thinking
  const partesReais = partes.filter(function(p) { return !p.thought; });

  if (partesReais.length > 0) {
    return partesReais.map(function(p) { return p.text || ''; }).join('');
  }

  // Fallback: se todas as partes são thinking, pega a última
  Logger.log('AVISO: Todas as partes são thinking, usando última parte como fallback');
  return partes[partes.length - 1].text || '';
}

function montarConfigGeracao(configPrompt, nomeEtapa) {
  nomeEtapa = nomeEtapa || 'desconhecida';

  const config = {
    temperature: configPrompt.temperatura,
    maxOutputTokens: configPrompt.maxTokens
  };

  // Pensamento desabilitado (0) → força thinkingBudget = 0
  if (configPrompt.pensamento === 0) {
    config.thinkingConfig = { thinkingBudget: 0 };
    Logger.log('[montarConfigGeracao] ' + nomeEtapa + ': thinking DESABILITADO, maxOutputTokens=' + configPrompt.maxTokens);
  }
  // Pensamento com budget explícito
  else if (configPrompt.pensamentoBudget && configPrompt.pensamentoBudget > 0) {
    config.thinkingConfig = { thinkingBudget: configPrompt.pensamentoBudget };
    Logger.log('[montarConfigGeracao] ' + nomeEtapa + ': thinkingBudget=' + configPrompt.pensamentoBudget + ', maxOutputTokens=' + configPrompt.maxTokens);
  }
  // Pensamento habilitado sem limite (padrão Gemini) — CUIDADO
  else {
    Logger.log('[montarConfigGeracao] ' + nomeEtapa + ': thinking HABILITADO SEM LIMITE, maxOutputTokens=' + configPrompt.maxTokens);
  }

  return config;
}

function validarSecaoGerada(textoSecao, nomeSecao) {
  const resultado = { valida: true, motivo: '' };

  // 1. Seção vazia ou muito curta
  if (!textoSecao || textoSecao.trim().length < 30) {
    resultado.valida = false;
    resultado.motivo = 'Seção "' + nomeSecao + '" vazia ou muito curta (' + (textoSecao ? textoSecao.trim().length : 0) + ' chars)';
    return resultado;
  }

  // 2. Detectar padrão de repetição de "---" ou "- - -" (bug principal!)
  const contagemTracos = (textoSecao.match(/-{3,}/g) || []).length;
  const proporcaoTracos = contagemTracos / (textoSecao.length / 100);
  if (contagemTracos > 20 || proporcaoTracos > 5) {
    resultado.valida = false;
    resultado.motivo = 'Seção "' + nomeSecao + '" contém padrão repetitivo de traços (' + contagemTracos + ' ocorrências de "---")';
    return resultado;
  }

  // 3. Detectar repetição excessiva de qualquer caractere/padrão
  // Verifica se mais de 40% do texto é o mesmo padrão de 3+ chars repetido
  const textoLimpo = textoSecao.replace(/\s+/g, ' ');
  if (textoLimpo.length > 200) {
    const amostra = textoLimpo.substring(0, 500);
    const palavras = amostra.split(' ');
    const contagem = {};
    palavras.forEach(function(p) {
      if (p.length >= 3) contagem[p] = (contagem[p] || 0) + 1;
    });
    const maxRepeticoes = Math.max.apply(null, Object.values(contagem).concat([0]));
    if (maxRepeticoes > palavras.length * 0.4 && palavras.length > 10) {
      resultado.valida = false;
      resultado.motivo = 'Seção "' + nomeSecao + '" contém palavra repetida ' + maxRepeticoes + '/' + palavras.length + ' vezes';
      return resultado;
    }
  }

  // 4. Seção absurdamente grande (> 50K chars para uma seção individual)
  if (textoSecao.length > 50000) {
    resultado.valida = false;
    resultado.motivo = 'Seção "' + nomeSecao + '" excede limite seguro (' + textoSecao.length + ' chars > 50K)';
    return resultado;
  }

  // 5. Verificar se o finishReason era MAX_TOKENS (truncamento)
  // Isso é verificado externamente, mas checamos padrão de texto cortado
  const ultimosChars = textoSecao.trim().slice(-20);
  if (ultimosChars.match(/[a-záéíóú]{1,3}$/i) && !ultimosChars.match(/[.!?)\]}"]\s*$/)) {
    // Texto parece cortado no meio de uma palavra — possível truncamento
    Logger.log('[validarSecaoGerada] AVISO: Seção "' + nomeSecao + '" pode ter sido truncada. Últimos chars: "' + ultimosChars + '"');
    // Não invalida, apenas loga (pode ser legítimo)
  }

  return resultado;
}

function processarAudioReuniao(dadosAudio) {
  const logs = [];
  const adicionarLog = (tipo, mensagem) => {
    const timestamp = new Date().toLocaleTimeString('pt-BR');
    logs.push({ tipo, mensagem, timestamp });
    Logger.log(`[${tipo}] ${mensagem}`);
  };

  try {
    adicionarLog('INFO', '🚀 Iniciando processamento do áudio...');

    if (!dadosAudio || !dadosAudio.audioBase64) {
      throw new Error('Dados do áudio não fornecidos');
    }

    adicionarLog('INFO', `📊 Tamanho do áudio: ${(dadosAudio.tamanhoBytes / 1024 / 1024).toFixed(2)} MB`);
    adicionarLog('INFO', `🎵 Tipo do arquivo: ${dadosAudio.tipoMime || 'audio/webm'}`);

    const chaveGemini = obterChaveGeminiProjeto();
    if (!chaveGemini) {
      throw new Error('Nenhuma chave API do Gemini configurada. Configure em Configurações.');
    }
    adicionarLog('SUCESSO', '🔑 Chave API do Gemini validada');

    // ── Carrega vocabulário UMA VEZ — reutilizado em toda a função ──
    adicionarLog('INFO', '📖 Carregando vocabulário personalizado...');
    const vocabulario = carregarVocabularioCompleto();
    if (vocabulario.totalTermos > 0) {
      adicionarLog('SUCESSO', `✅ ${vocabulario.totalTermos} termos de vocabulário carregados`);
    } else {
      adicionarLog('INFO', '⚠️ Aba Vocabulário vazia — processando sem glossário personalizado');
    }

    // ========== ETAPA 0: SALVAR ÁUDIO NO DRIVE ==========
    adicionarLog('INFO', '💾 Salvando áudio no Google Drive...');
    const resultadoDrive = salvarAudioNoDrive(dadosAudio);
    if (!resultadoDrive.sucesso) {
      throw new Error('Falha ao salvar no Drive: ' + resultadoDrive.mensagem);
    }
    adicionarLog('SUCESSO', `✅ Áudio salvo: ${resultadoDrive.nomeArquivo}`);

    // ========== ETAPA 1: TRANSCRIÇÃO ==========
    adicionarLog('INFO', '🎙️ [ETAPA 1/3] Iniciando transcrição do áudio...');
    adicionarLog('ALERTA', '⏳ Isso pode levar alguns minutos...');

    let resultadoTranscricao;
    const tamanhoBase64MB = (dadosAudio.audioBase64.length / 1024 / 1024);

    if (tamanhoBase64MB > 15) {
      adicionarLog('INFO', `📦 Áudio grande (${tamanhoBase64MB.toFixed(1)}MB base64), usando File API...`);
      const resultadoUpload = uploadParaFileApiGemini(
        resultadoDrive.arquivoId, dadosAudio.tipoMime || 'audio/webm', chaveGemini
      );
      if (!resultadoUpload.sucesso) {
        throw new Error('Falha no upload para Gemini File API: ' + resultadoUpload.mensagem);
      }
      adicionarLog('SUCESSO', '✅ Upload para File API concluído');
      // Passa vocabulário para injeção no prompt e validação pós-transcrição
      resultadoTranscricao = executarTranscricaoViaFileUri(
        resultadoUpload.fileUri, dadosAudio.tipoMime || 'audio/webm', chaveGemini, vocabulario
      );
      try { limparArquivoGemini(resultadoUpload.fileName, chaveGemini); } catch (e) { }
    } else {
      adicionarLog('INFO', `📦 Áudio pequeno (${tamanhoBase64MB.toFixed(1)}MB), usando inline_data...`);
      // Passa vocabulário para injeção no prompt e validação pós-transcrição
      resultadoTranscricao = executarEtapaTranscricao(dadosAudio, chaveGemini, vocabulario);
    }

    if (!resultadoTranscricao.sucesso) {
      throw new Error('Falha na transcrição: ' + resultadoTranscricao.mensagem);
    }

    // ========== SALVAR TRANSCRIÇÃO NO DRIVE ==========
    adicionarLog('INFO', '📝 Salvando transcrição no Drive...');
    const resultadoTranscricaoDrive = salvarTranscricaoNoDrive(resultadoTranscricao.transcricao, dadosAudio.titulo);
    if (resultadoTranscricaoDrive.sucesso) {
      adicionarLog('SUCESSO', `✅ Transcrição salva: ${resultadoTranscricaoDrive.nomeArquivo}`);
    }

    // ========== ETAPA 2: GERAÇÃO DA ATA ==========
    adicionarLog('INFO', '📋 [ETAPA 2/3] Gerando ATA da reunião...');
    const resultadoAta = executarEtapaGeracaoAta(
      resultadoTranscricao.transcricao, dadosAudio, chaveGemini
    );
    if (!resultadoAta.sucesso) {
      throw new Error('Falha na geração da ATA: ' + resultadoAta.mensagem);
    }
    adicionarLog('SUCESSO', '✅ ATA gerada com sucesso!');

    // ========== ETAPA 3: RELATÓRIO ==========
    adicionarLog('INFO', '🔍 [ETAPA 3/3] Gerando relatório de identificações...');
    const contexto = obterContextoProjetosParaGemini();
    const resultadoRelatorio = executarEtapaIdentificacaoAlteracoes(
      resultadoTranscricao.transcricao, contexto, chaveGemini, dadosAudio.titulo
    );

    let linkRelatorio = '', nomeArquivoRelatorio = '';
    let totalProjetosIdentificados = 0, totalEtapasIdentificadas = 0;
    let novosProjetosSugeridos = 0, novasEtapasSugeridas = 0, relatorioTexto = '';

    if (resultadoRelatorio.sucesso) {
      linkRelatorio = resultadoRelatorio.linkRelatorio || '';
      nomeArquivoRelatorio = resultadoRelatorio.nomeArquivoRelatorio || '';
      relatorioTexto = resultadoRelatorio.relatorio || '';
      totalProjetosIdentificados = Array.isArray(resultadoRelatorio.projetosIdentificados) ? resultadoRelatorio.projetosIdentificados.length : 0;
      totalEtapasIdentificadas = Array.isArray(resultadoRelatorio.etapasIdentificadas) ? resultadoRelatorio.etapasIdentificadas.length : 0;
      novosProjetosSugeridos = resultadoRelatorio.novosProjetosSugeridos || 0;
      novasEtapasSugeridas = resultadoRelatorio.novasEtapasSugeridas || 0;
      adicionarLog('SUCESSO', '✅ Relatório gerado!');
    } else {
      adicionarLog('ALERTA', '⚠️ Relatório não gerado: ' + resultadoRelatorio.mensagem);
    }

    // ========== SALVAR REUNIÃO ==========
    adicionarLog('INFO', '📊 Salvando reunião na planilha...');
    const reuniaoId = salvarReuniaoNaPlanilha({
      titulo: dadosAudio.titulo || 'Reunião ' + new Date().toLocaleDateString('pt-BR'),
      dataInicio: dadosAudio.dataInicio || new Date(),
      dataFim: dadosAudio.dataFim || new Date(),
      duracao: dadosAudio.duracaoMinutos || 0,
      participantes: dadosAudio.participantes || '',
      transcricao: resultadoTranscricao.transcricao || '',
      ata: resultadoAta.ata || '',
      sugestoesIA: resultadoAta.sugestoes || '',
      linkAudio: resultadoDrive.linkArquivo,
      projetosImpactados: '',
      etapasImpactadas: ''
    });
    adicionarLog('SUCESSO', `✅ Reunião salva com ID: ${reuniaoId}`);
    adicionarLog('SUCESSO', '🎉 Processamento concluído com sucesso!');

    return {
      sucesso: true, logs, reuniaoId,
      transcricao: resultadoTranscricao.transcricao,
      ata: resultadoAta.ata, sugestoes: resultadoAta.sugestoes || '',
      relatorioIdentificacoes: relatorioTexto,
      linkRelatorioIdentificacoes: linkRelatorio,
      nomeArquivoRelatorio,
      totalProjetosIdentificados, totalEtapasIdentificadas,
      novosProjetosSugeridos, novasEtapasSugeridas,
      linkAudio: resultadoDrive.linkArquivo
    };

  } catch (erro) {
    adicionarLog('ERRO', `❌ ${erro.message}`);
    Logger.log('ERRO processarAudioReuniao: ' + erro.toString());
    return { sucesso: false, logs, mensagem: erro.message };
  }
}

function executarEtapaTranscricao(dadosAudio, chaveApi, vocabulario) {
  try {
    // Carrega vocabulário se não foi passado externamente (evita leitura dupla da planilha)
    var vocab = vocabulario || carregarVocabularioCompleto();
    var blocoGlossario  = montarGlossarioParaPrompt(vocab);
    var blocoAnchoring  = montarExemplosAnchoring(vocab);
    var temVocabulario  = vocab.totalTermos > 0;

    if (temVocabulario) {
      Logger.log('[executarEtapaTranscricao] Glossário montado: ' + vocab.totalTermos + ' termos → injetando no prompt');
    }

    var audioBase64 = dadosAudio.audioBase64.split(',')[1] || dadosAudio.audioBase64;

    var promptTranscricao = 'Você é um transcritor profissional especializado em reuniões corporativas em português brasileiro.\n\n' +
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
      'Transcreva o áudio a seguir:';

    var urlApi = 'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG_PROMPTS_REUNIAO.TRANSCRICAO.modelo + ':generateContent?key=' + chaveApi;

    var corpoRequisicao = {
      contents: [{
        parts: [
          { text: promptTranscricao },
          {
            inline_data: {
              mime_type: dadosAudio.tipoMime || 'audio/webm',
              data: audioBase64
            }
          }
        ]
      }],
      generationConfig: montarConfigGeracao(CONFIG_PROMPTS_REUNIAO.TRANSCRICAO)
    };

    var opcoes = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(corpoRequisicao),
      muteHttpExceptions: true
    };

    var resposta = UrlFetchApp.fetch(urlApi, opcoes);
    var codigoStatus = resposta.getResponseCode();
    var textoResposta = resposta.getContentText();

    if (codigoStatus !== 200) {
      Logger.log('Erro Gemini Transcrição - Status: ' + codigoStatus);
      throw new Error('Erro na API (' + codigoStatus + '): ' + textoResposta.substring(0, 200));
    }

    var respostaJson = JSON.parse(textoResposta);
    var transcricaoBruta = extrairTextoRespostaGemini(respostaJson);

    // ── CAMADAS 2 e 3: Validação pós-transcrição ──
    var resultadoValidacao = validarECorrigirTranscricao(transcricaoBruta, vocab);
    if (resultadoValidacao.totalCorrecoes > 0) {
      Logger.log('[executarEtapaTranscricao] Camadas 2/3: ' + resultadoValidacao.totalCorrecoes + ' correção(ões) aplicada(s)');
    }

    return { sucesso: true, transcricao: resultadoValidacao.transcricaoCorrigida };

  } catch (erro) {
    Logger.log('ERRO executarEtapaTranscricao: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

// =====================================================================
//  ETAPA 1 (alternativa): TRANSCRIÇÃO VIA FILE URI (áudio grande)
//  ✅ FIX: usa extrairTextoRespostaGemini + desabilita thinking
// =====================================================================

function executarTranscricaoViaFileUri(fileUri, tipoMime, chaveApi, vocabulario) {
  try {
    // Carrega vocabulário se não foi passado externamente
    var vocab = vocabulario || carregarVocabularioCompleto();
    var blocoGlossario = montarGlossarioParaPrompt(vocab);
    var blocoAnchoring = montarExemplosAnchoring(vocab);

    if (vocab.totalTermos > 0) {
      Logger.log('[executarTranscricaoViaFileUri] Glossário: ' + vocab.totalTermos + ' termos → injetando no prompt');
    }

    var promptTranscricao = 'Você é um transcritor profissional especializado em reuniões corporativas em português brasileiro.\n\n' +
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
      'Transcreva o áudio a seguir:';

    var urlApi = 'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG_PROMPTS_REUNIAO.TRANSCRICAO.modelo + ':generateContent?key=' + chaveApi;

    var corpoRequisicao = {
      contents: [{
        parts: [
          { text: promptTranscricao },
          {
            file_data: {
              mime_type: tipoMime,
              file_uri: fileUri
            }
          }
        ]
      }],
      generationConfig: montarConfigGeracao(CONFIG_PROMPTS_REUNIAO.TRANSCRICAO)
    };

    var opcoes = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(corpoRequisicao),
      muteHttpExceptions: true
    };

    var resposta = UrlFetchApp.fetch(urlApi, opcoes);
    var codigoStatus = resposta.getResponseCode();
    var textoResposta = resposta.getContentText();

    if (codigoStatus !== 200) {
      Logger.log('Erro Gemini Transcrição via FileUri - Status: ' + codigoStatus);
      throw new Error('Erro na API (' + codigoStatus + '): ' + textoResposta.substring(0, 200));
    }

    var respostaJson = JSON.parse(textoResposta);
    var transcricaoBruta = extrairTextoRespostaGemini(respostaJson);

    // ── CAMADAS 2 e 3: Validação pós-transcrição ──
    var resultadoValidacao = validarECorrigirTranscricao(transcricaoBruta, vocab);
    if (resultadoValidacao.totalCorrecoes > 0) {
      Logger.log('[executarTranscricaoViaFileUri] Camadas 2/3: ' + resultadoValidacao.totalCorrecoes + ' correção(ões) aplicada(s)');
    }

    return { sucesso: true, transcricao: resultadoValidacao.transcricaoCorrigida };

  } catch (erro) {
    Logger.log('ERRO executarTranscricaoViaFileUri: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

// =====================================================================
//  ETAPA 2: GERAÇÃO DA ATA
//  ✅ FIX: Map-Reduce para transcrições longas
//  ✅ FIX: usa extrairTextoRespostaGemini + desabilita thinking
//  ✅ FIX: maxTokens aumentado para 65K
// =====================================================================

function executarEtapaGeracaoAta(transcricao, dadosAudio, chaveApi) {
  try {
    return gerarAtaSegmentada(transcricao, dadosAudio, chaveApi, false);
  } catch (erro) {
    Logger.log('ERRO executarEtapaGeracaoAta: ' + erro.toString());
    // Fallback: tenta a versão direta caso a segmentada falhe totalmente
    Logger.log('Tentando fallback com gerarAtaDireta...');
    return gerarAtaDireta(transcricao, dadosAudio, chaveApi, false);
  }
}

// =====================================================================
//  ✅ NOVO: MAP-REDUCE para ATAs de transcrições longas
//  Fase MAP: extrai pontos-chave de cada segmento
//  Fase REDUCE: gera ATA completa a partir dos pontos combinados
// =====================================================================

function executarAtaMapReduce(transcricao, dadosAudio, chaveApi) {
  try {
    const limiteChars = CONFIGURACAO_REUNIOES.LIMITE_CHARS_MAP_REDUCE;

    // ── FASE MAP: dividir e extrair ──
    const segmentos = dividirTranscricaoEmSegmentos(transcricao, limiteChars);
    Logger.log(`Map-Reduce: ${segmentos.length} segmentos de ~${limiteChars} chars`);

    const pontosExtraidos = [];

    for (let i = 0; i < segmentos.length; i++) {
      Logger.log(`Extraindo pontos-chave do segmento ${i + 1}/${segmentos.length}...`);
      const resultado = extrairPontosChaveSegmento(segmentos[i], i + 1, segmentos.length, chaveApi);

      if (resultado.sucesso) {
        pontosExtraidos.push(resultado.pontosChave);
      } else {
        Logger.log(`AVISO: Falha na extração do segmento ${i + 1}: ${resultado.mensagem}`);
        // Usa o segmento bruto como fallback (truncado)
        pontosExtraidos.push(`[SEGMENTO ${i + 1} - extração falhou, resumo bruto]\n${segmentos[i].substring(0, 15000)}`);
      }
    }

    // ── FASE REDUCE: gerar ATA a partir dos pontos combinados ──
    const pontosConsolidados = pontosExtraidos.join('\n\n========== PRÓXIMO SEGMENTO ==========\n\n');
    Logger.log(`Reduce: gerando ATA a partir de ${pontosConsolidados.length} chars de pontos-chave`);

    return gerarAtaDireta(pontosConsolidados, dadosAudio, chaveApi, true);

  } catch (erro) {
    Logger.log('ERRO executarAtaMapReduce: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

/** Divide transcrição em segmentos respeitando quebras de linha */
function dividirTranscricaoEmSegmentos(texto, tamanhoMaximo) {
  const segmentos = [];
  let inicio = 0;

  while (inicio < texto.length) {
    let fim = Math.min(inicio + tamanhoMaximo, texto.length);

    // Se não é o último segmento, tenta cortar em uma quebra de linha
    if (fim < texto.length) {
      const ultimaQuebra = texto.lastIndexOf('\n', fim);
      if (ultimaQuebra > inicio + tamanhoMaximo * 0.5) {
        fim = ultimaQuebra + 1;
      }
    }

    segmentos.push(texto.substring(inicio, fim));
    inicio = fim;
  }

  return segmentos;
}

/** ✅ NOVO: Extrai pontos-chave de um segmento da transcrição */
function extrairPontosChaveSegmento(segmento, numSegmento, totalSegmentos, chaveApi) {
  try {
    const promptExtracao = `Você é um analista especializado em extrair informações estruturadas de transcrições de reunião.

## CONTEXTO:
Esta é a PARTE ${numSegmento} de ${totalSegmentos} de uma transcrição de reunião.

## SUA TAREFA:
Extraia TODOS os pontos importantes desta parte da transcrição, sem omitir nada.

## EXTRAIA OBRIGATORIAMENTE:
1. **TÓPICOS DISCUTIDOS**: Liste cada tema/assunto abordado com detalhes
2. **DECISÕES TOMADAS**: Qualquer decisão ou resolução mencionada
3. **AÇÕES DELEGADAS**: Quem ficou responsável por quê (nome → tarefa → prazo)
4. **PROBLEMAS/DIFICULDADES**: Bloqueios, reclamações, pendências mencionadas
5. **PROCESSOS DESCRITOS**: Fluxos de trabalho, procedimentos, regras explicadas
6. **MUDANÇAS ORGANIZACIONAIS**: Alterações em equipe, estrutura, responsabilidades
7. **PRAZOS E DATAS**: Qualquer menção a prazos, deadlines, datas
8. **NOMES E FUNÇÕES**: Pessoas mencionadas e seus papéis
9. **NÚMEROS E DADOS**: Valores, métricas, quantidades mencionadas

## REGRAS:
- Seja EXAUSTIVO, extraia TUDO, não resuma demais
- Mantenha detalhes específicos (nomes, números, datas)
- Use citações diretas quando relevante
- Não invente informações
- Organize por tópico

## TRANSCRIÇÃO (PARTE ${numSegmento}/${totalSegmentos}):
${segmento}

Extraia todos os pontos-chave:`;

    const urlApi = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG_PROMPTS_REUNIAO.EXTRACAO.modelo}:generateContent?key=${chaveApi}`;

    const corpoRequisicao = {
      contents: [{ parts: [{ text: promptExtracao }] }],
      generationConfig: montarConfigGeracao(CONFIG_PROMPTS_REUNIAO.EXTRACAO)
    };

    const resposta = UrlFetchApp.fetch(urlApi, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(corpoRequisicao),
      muteHttpExceptions: true
    });

    if (resposta.getResponseCode() !== 200) {
      throw new Error(`Erro API (${resposta.getResponseCode()})`);
    }

    const respostaJson = JSON.parse(resposta.getContentText());
    const pontosChave = extrairTextoRespostaGemini(respostaJson);

    return { sucesso: true, pontosChave: pontosChave };

  } catch (erro) {
    Logger.log('ERRO extrairPontosChaveSegmento: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function gerarAtaDireta(textoBase, dadosAudio, chaveApi, ehMapReduce) {
  try {
    let textoParaAta = textoBase;
    const LIMITE_DIRETO = 200000;

    if (!ehMapReduce && textoParaAta.length > LIMITE_DIRETO) {
      const metade = Math.floor(LIMITE_DIRETO / 2);
      textoParaAta = textoParaAta.substring(0, metade) +
        '\n\n[... seção central omitida por tamanho — ' +
        (textoParaAta.length - LIMITE_DIRETO) + ' caracteres ...]\n\n' +
        textoParaAta.substring(textoParaAta.length - metade);
      Logger.log('[gerarAtaDireta] Truncado de ' + textoBase.length + ' para ' + textoParaAta.length + ' chars');
    }

    const tipoFonte = ehMapReduce
      ? 'PONTOS-CHAVE EXTRAÍDOS (consolidados de múltiplos segmentos)'
      : 'TRANSCRIÇÃO DA REUNIÃO';

    const instrucaoExtra = ehMapReduce
      ? `
## ⚠️ ATENÇÃO ESPECIAL:
Os dados abaixo são PONTOS-CHAVE já extraídos de uma transcrição longa, dividida em segmentos.
Você DEVE incluir informações de TODOS os segmentos na ATA final.
NÃO omita nenhum segmento. A ATA deve cobrir 100% dos tópicos discutidos.`
      : '';

    const promptAta = `Você é um secretário executivo especializado em redigir atas de reunião profissionais e COMPLETAS.

## DADOS DA REUNIÃO:
- **Título informado:** ${dadosAudio.titulo || 'Não informado'}
- **Participantes informados:** ${dadosAudio.participantes || 'Não informados'}
- **Data:** ${new Date().toLocaleDateString('pt-BR')}
${instrucaoExtra}

## ${tipoFonte}:
${textoParaAta}

## SUA TAREFA:
Analise o conteúdo acima e redija uma ATA DE REUNIÃO completa e profissional.
Você DEVE preencher TODAS as seções abaixo. NÃO deixe nenhuma seção em branco ou incompleta.

## ESTRUTURA OBRIGATÓRIA DA ATA:

### 1. CABEÇALHO
- Nome da Reunião (extraia do contexto ou use o título informado)
- Data e Hora (Início/Término estimados)
- Duração aproximada
- Local (presencial/virtual)
- Participantes Presentes
- Outros Funcionários Citados
- Pauta Principal

### 2. TEMA PRINCIPAL E OBJETIVOS
Um parágrafo resumindo o propósito central da reunião.

### 3. DETALHES DA DISCUSSÃO POR TÓPICO
Liste os principais pontos debatidos de forma numerada:
- Decisões tomadas
- Processos mencionados ou mapeados
- Mudanças organizacionais
- Dificuldades ou limitações declaradas

### 4. MATRIZ DE AÇÃO (PLANO DE AÇÃO)
Tabela com: Nº | Ação | Responsável | Prazo | Status
Liste Principais Pendências e Próximos Passos.

### 5. OUTROS PONTOS LEVANTADOS
Observações secundárias, avisos, prazos futuros.

### 6. CONSIDERAÇÕES FINAIS
Fechamento sintetizando o clima da reunião e principais pendências.

## FORMATO DE SAÍDA:
Retorne a ATA em formato Markdown bem estruturado.
Use tabelas markdown para a Matriz de Ação.
NÃO inclua JSON, apenas a ATA formatada.
NÃO repita separadores, traços ou qualquer padrão.

## ⚠️ REGRA CRÍTICA:
- PREENCHA TODAS AS 6 SEÇÕES COMPLETAMENTE
- Se uma seção não tem informação, escreva "Não identificado na reunião" em vez de deixar em branco
- Seja objetivo e profissional
- Não invente informações que não estejam no conteúdo fornecido
- Use linguagem formal e clara
- Destaque decisões importantes em negrito`;

    Logger.log('[gerarAtaDireta] Tamanho prompt: ' + promptAta.length + ' chars');

    const urlApi = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG_PROMPTS_REUNIAO.ATA.modelo}:generateContent?key=${chaveApi}`;

    const configGeracao = montarConfigGeracao(CONFIG_PROMPTS_REUNIAO.ATA, 'ATA-direta');

    // Para geração direta (não segmentada), usa um budget maior
    configGeracao.maxOutputTokens = 65000;

    const corpoRequisicao = {
      contents: [{ parts: [{ text: promptAta }] }],
      generationConfig: configGeracao
    };

    const tempoInicio = Date.now();

    const resposta = UrlFetchApp.fetch(urlApi, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(corpoRequisicao),
      muteHttpExceptions: true
    });

    const codigoStatus = resposta.getResponseCode();
    const tempoMs = Date.now() - tempoInicio;

    Logger.log('[gerarAtaDireta] HTTP ' + codigoStatus + ' em ' + (tempoMs / 1000).toFixed(1) + 's');

    if (codigoStatus !== 200) {
      throw new Error('Erro na API (' + codigoStatus + '): ' + resposta.getContentText().substring(0, 200));
    }

    const respostaJson = JSON.parse(resposta.getContentText());

    // ✅ Log de tokens usados
    const metadadosUso = respostaJson.usageMetadata || {};
    Logger.log('[gerarAtaDireta] Tokens: prompt=' + (metadadosUso.promptTokenCount || '?') +
      ', output=' + (metadadosUso.candidatesTokenCount || '?') +
      ', thinking=' + (metadadosUso.thoughtsTokenCount || '?'));

    const candidato = respostaJson.candidates && respostaJson.candidates[0];
    Logger.log('[gerarAtaDireta] finishReason: ' + (candidato ? candidato.finishReason : 'N/A'));

    const ataGerada = extrairTextoRespostaGemini(respostaJson);

    Logger.log('[gerarAtaDireta] ATA gerada: ' + ataGerada.length + ' chars');

    // ✅ Validação
    const validacao = validarSecaoGerada(ataGerada, 'ATA_DIRETA');
    if (!validacao.valida) {
      Logger.log('[gerarAtaDireta] ⚠️ ATA INVÁLIDA: ' + validacao.motivo);
      return { sucesso: false, mensagem: 'ATA gerada inválida: ' + validacao.motivo };
    }

    const sugestoes = extrairSugestoesDoContexto(textoBase.substring(0, 50000));

    return { sucesso: true, ata: ataGerada, sugestoes: sugestoes };

  } catch (erro) {
    Logger.log('[gerarAtaDireta] ERRO: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

// =====================================================================
//  GERAÇÃO SEGMENTADA DA ATA — 4 chamadas independentes
//  Cada chamada gera 1-2 seções, evitando limite de tokens de saída
// =====================================================================

function gerarAtaSegmentada(textoBase, dadosAudio, chaveApi, ehMapReduce) {
  try {
    // ── Truncagem de segurança ──
    let textoParaAta = textoBase;
    const LIMITE_DIRETO = 200000;

    if (!ehMapReduce && textoParaAta.length > LIMITE_DIRETO) {
      const metade = Math.floor(LIMITE_DIRETO / 2);
      textoParaAta = textoParaAta.substring(0, metade) +
        '\n\n[... seção central omitida por tamanho — ' +
        (textoParaAta.length - LIMITE_DIRETO) + ' caracteres ...]\n\n' +
        textoParaAta.substring(textoParaAta.length - metade);
      Logger.log('[gerarAtaSegmentada] Truncado de ' + textoBase.length + ' para ' + textoParaAta.length + ' chars');
    }

    const tipoFonte = ehMapReduce
      ? 'PONTOS-CHAVE EXTRAÍDOS (consolidados de múltiplos segmentos)'
      : 'TRANSCRIÇÃO DA REUNIÃO';

    const cabecalhoContexto = montarCabecalhoContextoAta(dadosAudio, tipoFonte, textoParaAta, ehMapReduce);

    // ── DEFINIÇÃO DAS 4 SEÇÕES (sem alteração nos prompts) ──
    const definicoesSecoes = [
      {
        nome: 'Cabeçalho + Tema Principal',
        instrucao: `Gere APENAS as seções 1 e 2 da ATA:

### 1. CABEÇALHO
- Nome da Reunião (extraia do contexto ou use o título informado)
- Data e Hora (Início/Término estimados)
- Local (presencial/virtual)
- Participantes Presentes
- Outros Funcionários Citados
- Pauta Principal

### 2. TEMA PRINCIPAL E OBJETIVOS
Um parágrafo resumindo o propósito central da reunião.

⚠️ REGRAS:
- Retorne APENAS essas 2 seções, nada mais
- Use Markdown bem formatado
- Não inclua introduções ou explicações, vá direto ao conteúdo
- NÃO inclua duração da reunião, links de áudio, links do Drive ou referências a arquivos
- Preencha TODOS os campos; se não encontrar informação, escreva "Não identificado na reunião"`
      },
      {
        nome: 'Detalhes da Discussão',
        instrucao: `Gere APENAS a seção 3 da ATA:

### 3. DETALHES DA DISCUSSÃO POR TÓPICO
Liste TODOS os principais pontos debatidos de forma numerada e detalhada:
- Decisões tomadas
- Processos mencionados ou mapeados
- Mudanças organizacionais
- Dificuldades ou limitações declaradas
- Cada tópico deve ter um subtítulo descritivo

⚠️ REGRAS:
- Retorne APENAS esta seção, nada mais
- Seja EXAUSTIVO: inclua TODOS os tópicos discutidos, não resuma
- Use Markdown bem formatado
- Não inclua introduções ou explicações, vá direto ao conteúdo
- NÃO inclua links externos, URLs ou referências a arquivos anexados
- Cada tópico deve ter detalhes suficientes para quem não participou entender o que foi discutido`
      },
      {
        nome: 'Matriz de Ação',
        instrucao: `Gere APENAS a seção 4 da ATA:

### 4. MATRIZ DE AÇÃO (PLANO DE AÇÃO)

Crie uma tabela Markdown com Principais Pendências e Próximos Passos.
Formato da tabela:

| Nº | Ação | Responsável | Prazo | Status |
|----|------|-------------|-------|--------|
| 1  | ...  | ...         | ...   | ...    |

⚠️ REGRAS CRÍTICAS:
- Retorne APENAS esta seção (a tabela), nada mais
- Evite adicionar tarefas que fujam do objetivo central da reunião!
- Liste SOMENTE ações PENDENTES ou FUTURAS — o que AINDA PRECISA SER FEITO
- NÃO inclua tarefas já implementadas, concluídas ou resolvidas antes/durante a reunião
- Se alguém mencionou que "já fez X" ou "X já está pronto", NÃO liste X na tabela
- Se o responsável não foi mencionado, escreva "A definir"
- Se o prazo não foi mencionado, escreva "A definir"
- Status padrão: "Pendente" (use "Em Andamento" apenas se explicitamente iniciado mas não concluído)
- Não inclua introduções ou explicações, vá direto à tabela
- NÃO inclua links externos ou URLs na tabela
- Se nenhuma ação pendente foi identificada, escreva "Nenhuma ação pendente foi identificada na reunião"
- NÃO repita separadores ou traços além dos necessários para a tabela Markdown`
      },
      {
        nome: 'Outros Pontos + Considerações Finais',
        instrucao: `Gere APENAS as seções 5 e 6 da ATA:

### 5. OUTROS PONTOS LEVANTADOS
Observações secundárias, avisos, prazos futuros, informações adicionais que não se encaixaram nos tópicos anteriores.

### 6. CONSIDERAÇÕES FINAIS
Fechamento sintetizando o clima da reunião, principais pendências e próximos passos.

⚠️ REGRAS:
- Retorne APENAS essas 2 seções, nada mais
- Use Markdown bem formatado
- Não inclua introduções ou explicações, vá direto ao conteúdo
- NÃO inclua links externos, URLs, links do Drive ou referências a arquivos/gravações
- Se não houver informação para alguma seção, escreva "Não identificado na reunião"`
      }
    ];

    // ── CONSTANTES DE CONTROLE ──
    const MAX_TENTATIVAS_POR_SECAO = 2; // 1 tentativa + 1 retry
    const LIMITE_CHARS_POR_SECAO = 50000; // Segurança: nenhuma seção deve ter mais que isso
    const LIMITE_TOTAL_ATA = 200000; // ATA final não pode exceder 200K chars

    // ── EXECUÇÃO DAS 4 CHAMADAS COM VALIDAÇÃO ──
    const secoesGeradas = [];
    let tempoTotalGeracaoMs = 0;

    Logger.log('====================================================');
    Logger.log('[gerarAtaSegmentada] INICIANDO geração de 4 seções');
    Logger.log('[gerarAtaSegmentada] Modelo: ' + CONFIG_PROMPTS_REUNIAO.ATA.modelo);
    Logger.log('[gerarAtaSegmentada] maxTokens/seção: ' + CONFIG_PROMPTS_REUNIAO.ATA.maxTokens);
    Logger.log('[gerarAtaSegmentada] Pensamento: ' + (CONFIG_PROMPTS_REUNIAO.ATA.pensamento === 0 ? 'DESABILITADO' : 'HABILITADO'));
    Logger.log('[gerarAtaSegmentada] Tamanho contexto: ' + textoParaAta.length + ' chars');
    Logger.log('[gerarAtaSegmentada] ehMapReduce: ' + ehMapReduce);
    Logger.log('====================================================');

    for (let i = 0; i < definicoesSecoes.length; i++) {
      const secao = definicoesSecoes[i];
      let textoSecaoFinal = '';
      let secaoValida = false;

      for (let tentativa = 1; tentativa <= MAX_TENTATIVAS_POR_SECAO; tentativa++) {
        const tempoInicioSecao = Date.now();

        Logger.log('----------------------------------------------------');
        Logger.log('[gerarAtaSegmentada] Seção ' + (i + 1) + '/4: "' + secao.nome + '" (tentativa ' + tentativa + '/' + MAX_TENTATIVAS_POR_SECAO + ')');

        const promptSecao = cabecalhoContexto + '\n\n## SUA TAREFA AGORA:\n' + secao.instrucao;

        Logger.log('[gerarAtaSegmentada] Tamanho do prompt: ' + promptSecao.length + ' chars');

        const urlApi = 'https://generativelanguage.googleapis.com/v1beta/models/' +
          CONFIG_PROMPTS_REUNIAO.ATA.modelo + ':generateContent?key=' + chaveApi;

        const configGeracao = montarConfigGeracao(CONFIG_PROMPTS_REUNIAO.ATA, 'ATA-secao-' + (i + 1));

        const corpoRequisicao = {
          contents: [{ parts: [{ text: promptSecao }] }],
          generationConfig: configGeracao
        };

        let resposta;
        try {
          resposta = UrlFetchApp.fetch(urlApi, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(corpoRequisicao),
            muteHttpExceptions: true
          });
        } catch (erroFetch) {
          Logger.log('[gerarAtaSegmentada] ERRO de rede na seção ' + (i + 1) + ': ' + erroFetch.toString());
          if (tentativa < MAX_TENTATIVAS_POR_SECAO) {
            Logger.log('[gerarAtaSegmentada] Aguardando 3s antes do retry...');
            Utilities.sleep(3000);
            continue;
          }
          textoSecaoFinal = '\n\n### ⚠️ [Seção "' + secao.nome + '" não gerada — erro de rede]\n\n';
          break;
        }

        const codigoStatus = resposta.getResponseCode();
        const tempoSecaoMs = Date.now() - tempoInicioSecao;
        tempoTotalGeracaoMs += tempoSecaoMs;

        Logger.log('[gerarAtaSegmentada] HTTP ' + codigoStatus + ' em ' + (tempoSecaoMs / 1000).toFixed(1) + 's');

        if (codigoStatus !== 200) {
          Logger.log('[gerarAtaSegmentada] ERRO HTTP seção ' + (i + 1) + ': ' + resposta.getContentText().substring(0, 300));

          if (tentativa < MAX_TENTATIVAS_POR_SECAO) {
            Logger.log('[gerarAtaSegmentada] Retry em 3s...');
            Utilities.sleep(3000);
            continue;
          }

          textoSecaoFinal = '\n\n### ⚠️ [Seção "' + secao.nome + '" não gerada — erro HTTP ' + codigoStatus + ']\n\n';
          break;
        }

        // ── Parse da resposta ──
        const respostaJson = JSON.parse(resposta.getContentText());

        // ✅ NOVO: Log detalhado do finishReason e tokens usados
        const candidato = respostaJson.candidates && respostaJson.candidates[0];
        const motivoParada = candidato ? (candidato.finishReason || 'N/A') : 'SEM_CANDIDATO';
        const metadadosUso = respostaJson.usageMetadata || {};

        Logger.log('[gerarAtaSegmentada] finishReason: ' + motivoParada);
        Logger.log('[gerarAtaSegmentada] usageMetadata: promptTokenCount=' + (metadadosUso.promptTokenCount || '?') +
          ', candidatesTokenCount=' + (metadadosUso.candidatesTokenCount || '?') +
          ', totalTokenCount=' + (metadadosUso.totalTokenCount || '?') +
          ', thoughtsTokenCount=' + (metadadosUso.thoughtsTokenCount || '?'));

        // ✅ NOVO: Alerta se atingiu limite de tokens (MAX_TOKENS = conteúdo truncado)
        if (motivoParada === 'MAX_TOKENS') {
          Logger.log('[gerarAtaSegmentada] ⚠️ ALERTA: Seção ' + (i + 1) + ' TRUNCADA pelo limite de tokens!');
        }

        const textoSecao = extrairTextoRespostaGemini(respostaJson);

        Logger.log('[gerarAtaSegmentada] Seção ' + (i + 1) + ' bruta: ' + textoSecao.length + ' chars');

        // ✅ NOVO: Validação da seção gerada
        const validacao = validarSecaoGerada(textoSecao, secao.nome);

        if (!validacao.valida) {
          Logger.log('[gerarAtaSegmentada] ❌ SEÇÃO INVÁLIDA: ' + validacao.motivo);

          if (tentativa < MAX_TENTATIVAS_POR_SECAO) {
            Logger.log('[gerarAtaSegmentada] Retry da seção ' + (i + 1) + '...');
            Utilities.sleep(2000);
            continue;
          }

          // Último retry falhou → usa placeholder
          textoSecaoFinal = '\n\n### ⚠️ [Seção "' + secao.nome + '" não gerada corretamente — ' + validacao.motivo + ']\n\n';
          Logger.log('[gerarAtaSegmentada] Seção ' + (i + 1) + ' SUBSTITUÍDA por placeholder após ' + MAX_TENTATIVAS_POR_SECAO + ' tentativas');
          break;
        }

        // ✅ Seção válida!
        textoSecaoFinal = textoSecao;
        secaoValida = true;
        Logger.log('[gerarAtaSegmentada] ✅ Seção ' + (i + 1) + '/4 OK (' + textoSecao.length + ' chars, ' + (tempoSecaoMs / 1000).toFixed(1) + 's)');
        break; // Sai do loop de tentativas
      }

      secoesGeradas.push(textoSecaoFinal);
    }

    // ── CONCATENAÇÃO FINAL COM VALIDAÇÃO GLOBAL ──
    let ataCompleta = secoesGeradas.join('\n\n---\n\n');

    Logger.log('====================================================');
    Logger.log('[gerarAtaSegmentada] ATA montada: ' + ataCompleta.length + ' chars');
    Logger.log('[gerarAtaSegmentada] Tempo total geração: ' + (tempoTotalGeracaoMs / 1000).toFixed(1) + 's');

    // ✅ NOVO: Validação final — ATA não pode exceder limite
    if (ataCompleta.length > LIMITE_TOTAL_ATA) {
      Logger.log('[gerarAtaSegmentada] ⚠️ ATA excedeu limite! ' + ataCompleta.length + ' > ' + LIMITE_TOTAL_ATA);
      Logger.log('[gerarAtaSegmentada] Truncando ATA para ' + LIMITE_TOTAL_ATA + ' chars');

      ataCompleta = ataCompleta.substring(0, LIMITE_TOTAL_ATA) +
        '\n\n---\n\n### ⚠️ ATA truncada\nA ATA excedeu o limite seguro de ' +
        LIMITE_TOTAL_ATA + ' caracteres e foi truncada. ' +
        'Total original: ' + secoesGeradas.join('').length + ' chars.';
    }

    // ✅ NOVO: Detectar se a ATA final ainda contém padrões corrompidos
    const validacaoFinal = validarSecaoGerada(ataCompleta, 'ATA_COMPLETA');
    if (!validacaoFinal.valida) {
      Logger.log('[gerarAtaSegmentada] ❌ ATA FINAL CORROMPIDA: ' + validacaoFinal.motivo);
      // Não retorna erro — retorna o que tem, mas com aviso
      ataCompleta = '### ⚠️ AVISO: Esta ATA pode conter erros de geração\n\n' +
        'Motivo: ' + validacaoFinal.motivo + '\n\n---\n\n' + ataCompleta;
    }

    Logger.log('[gerarAtaSegmentada] ✅ CONCLUÍDO. ATA final: ' + ataCompleta.length + ' chars');
    Logger.log('====================================================');

    const sugestoes = extrairSugestoesDoContexto(textoParaAta.substring(0, 50000));

    return { sucesso: true, ata: ataCompleta, sugestoes: sugestoes };

  } catch (erro) {
    Logger.log('[gerarAtaSegmentada] ERRO FATAL: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}


/**
 * Monta o cabeçalho de contexto compartilhado entre todas as chamadas segmentadas.
 * Evita repetir a transcrição inteira no prompt — cada seção recebe o mesmo contexto.
 */
function montarCabecalhoContextoAta(dadosAudio, tipoFonte, textoParaAta, ehMapReduce) {
  const instrucaoExtra = ehMapReduce
    ? '\n## ⚠️ ATENÇÃO: Os dados abaixo são PONTOS-CHAVE extraídos de múltiplos segmentos. Cubra TODOS os segmentos.'
    : '';

  return 'Você é um secretário executivo especializado em redigir atas de reunião profissionais e COMPLETAS.\n\n' +
    '## DADOS DA REUNIÃO:\n' +
    '- **Título:** ' + (dadosAudio.titulo || 'Não informado') + '\n' +
    '- **Participantes:** ' + (dadosAudio.participantes || 'Não informados') + '\n' +
    '- **Data:** ' + new Date().toLocaleDateString('pt-BR') + '\n' +
    instrucaoExtra + '\n\n' +
    '## ' + tipoFonte + ':\n' + textoParaAta + '\n\n' +
    '## REGRAS GERAIS:\n' +
    '- Use linguagem formal e clara\n' +
    '- Destaque decisões importantes em **negrito**\n' +
    '- Não invente informações que não estejam no conteúdo fornecido\n' +
    '- Use formato Markdown\n' +
    '- NÃO inclua JSON';
}


// =====================================================================
//  SUBSTITUIR: etapa3_GerarAta
//  Mudança: Adiciona log sobre truncagem, sem map-reduce
// =====================================================================

function etapa3_GerarAta(dados) {
  const logs = [];
  const log = function(tipo, msg) {
    logs.push({ tipo: tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log('[etapa3_GerarAta][' + tipo + '] ' + msg);
  };

  try {
    log('INFO', '📋 Gerando ATA da reunião (modo segmentado: 4 seções)...');
    log('INFO', '  📝 Transcrição: ' + dados.transcricao.length + ' chars');
    log('INFO', '  🤖 Modelo: ' + CONFIG_PROMPTS_REUNIAO.ATA.modelo);
    log('INFO', '  🧠 Pensamento: ' + (CONFIG_PROMPTS_REUNIAO.ATA.pensamento === 0 ? 'DESABILITADO' : 'HABILITADO'));
    log('INFO', '  📊 maxTokens/seção: ' + CONFIG_PROMPTS_REUNIAO.ATA.maxTokens);

    if (dados.transcricao.length > 200000) {
      log('ALERTA', '  ✂️ Transcrição será truncada (' + dados.transcricao.length + ' > 200K chars)');
      log('INFO', '  💡 Para resultado completo, use Map-Reduce no client');
    }

    const tempoInicio = Date.now();
    const chaveApi = obterChaveGeminiProjeto();

    if (!chaveApi) {
      log('ERRO', '❌ Chave API do Gemini não configurada!');
      return { sucesso: false, logs: logs, mensagem: 'Chave API não configurada' };
    }

    const resultado = executarEtapaGeracaoAta(
      dados.transcricao,
      {
        titulo: dados.titulo || 'Reunião ' + new Date().toLocaleDateString('pt-BR'),
        participantes: dados.participantes || ''
      },
      chaveApi
    );

    const tempoSeg = ((Date.now() - tempoInicio) / 1000).toFixed(1);

    if (!resultado.sucesso) {
      log('ERRO', '❌ ATA falhou (' + tempoSeg + 's): ' + resultado.mensagem);
      return { sucesso: false, logs: logs, mensagem: resultado.mensagem };
    }

    // ✅ NOVO: Validação do resultado antes de retornar
    const tamanhoAta = resultado.ata ? resultado.ata.length : 0;
    log('SUCESSO', '✅ ATA gerada em ' + tempoSeg + 's! (' + tamanhoAta + ' chars)');

    if (tamanhoAta > 150000) {
      log('ALERTA', '⚠️ ATA muito grande (' + tamanhoAta + ' chars). Possível conteúdo repetitivo.');
    }

    if (tamanhoAta < 100) {
      log('ALERTA', '⚠️ ATA muito curta (' + tamanhoAta + ' chars). Possível falha parcial.');
    }

    // ✅ NOVO: Log com preview do início e fim da ATA
    if (resultado.ata) {
      log('INFO', '  📄 Preview início: "' + resultado.ata.substring(0, 120).replace(/\n/g, ' ') + '..."');
      log('INFO', '  📄 Preview fim: "...' + resultado.ata.substring(resultado.ata.length - 80).replace(/\n/g, ' ') + '"');
    }

    return { sucesso: true, logs: logs, ata: resultado.ata, sugestoes: resultado.sugestoes || '' };

  } catch (erro) {
    log('ERRO', '❌ Exceção: ' + erro.message);
    return { sucesso: false, logs: logs, mensagem: erro.message };
  }
}


// =====================================================================
//  ✅ ADICIONAR: extrairPontosSegmentoServidor
//  Chamada ATÔMICA (1 segmento por vez, < 2 min de execução)
//  O client chama essa função N vezes em PARALELO
// =====================================================================

function extrairPontosSegmentoServidor(dados) {
  const logs = [];
  const log = (tipo, msg) => {
    logs.push({ tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log(`[${tipo}] ${msg}`);
  };

  try {
    const numSegmento = dados.numSegmento || 1;
    const totalSegmentos = dados.totalSegmentos || 1;
    const segmento = dados.segmento || '';

    log('INFO', `📝 Extraindo pontos-chave do segmento ${numSegmento}/${totalSegmentos} (${segmento.length} chars)...`);
    const tempoInicio = Date.now();
    const chaveApi = obterChaveGeminiProjeto();

    const promptExtracao = `Você é um analista especializado em extrair informações estruturadas de transcrições de reunião.

## CONTEXTO:
Esta é a PARTE ${numSegmento} de ${totalSegmentos} de uma transcrição de reunião.

## SUA TAREFA:
Extraia TODOS os pontos importantes desta parte da transcrição, sem omitir nada.

## EXTRAIA OBRIGATORIAMENTE:
1. **TÓPICOS DISCUTIDOS**: Liste cada tema/assunto abordado com detalhes
2. **DECISÕES TOMADAS**: Qualquer decisão ou resolução mencionada
3. **AÇÕES DELEGADAS**: Quem ficou responsável por quê (nome → tarefa → prazo)
4. **PROBLEMAS/DIFICULDADES**: Bloqueios, reclamações, pendências mencionadas
5. **PROCESSOS DESCRITOS**: Fluxos de trabalho, procedimentos, regras explicadas
6. **MUDANÇAS ORGANIZACIONAIS**: Alterações em equipe, estrutura, responsabilidades
7. **PRAZOS E DATAS**: Qualquer menção a prazos, deadlines, datas
8. **NOMES E FUNÇÕES**: Pessoas mencionadas e seus papéis
9. **NÚMEROS E DADOS**: Valores, métricas, quantidades mencionadas

## REGRAS:
- Seja EXAUSTIVO, extraia TUDO, não resuma demais
- Mantenha detalhes específicos (nomes, números, datas)
- Use citações diretas quando relevante
- Não invente informações
- Organize por tópico

## TRANSCRIÇÃO (PARTE ${numSegmento}/${totalSegmentos}):
${segmento}

Extraia todos os pontos-chave:`;

    const urlApi = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG_PROMPTS_REUNIAO.EXTRACAO.modelo}:generateContent?key=${chaveApi}`;

    const corpoRequisicao = {
      contents: [{ parts: [{ text: promptExtracao }] }],
      generationConfig: montarConfigGeracao(CONFIG_PROMPTS_REUNIAO.EXTRACAO)
    };

    const resposta = UrlFetchApp.fetch(urlApi, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(corpoRequisicao),
      muteHttpExceptions: true
    });

    if (resposta.getResponseCode() !== 200) {
      throw new Error(`Erro API (${resposta.getResponseCode()}): ${resposta.getContentText().substring(0, 200)}`);
    }

    const respostaJson = JSON.parse(resposta.getContentText());
    const pontosChave = extrairTextoRespostaGemini(respostaJson);

    const tempoSeg = ((Date.now() - tempoInicio) / 1000).toFixed(1);
    log('SUCESSO', `✅ Segmento ${numSegmento}/${totalSegmentos} extraído em ${tempoSeg}s (${pontosChave.length} chars)`);

    return { sucesso: true, logs: logs, pontosChave: pontosChave };

  } catch (erro) {
    log('ERRO', `❌ Segmento ${dados.numSegmento || '?'}: ${erro.message}`);
    return { sucesso: false, logs: logs, mensagem: erro.message, pontosChave: '' };
  }
}


function gerarAtaDesdeExtracoes(dados) {
  const logs = [];
  const log = function(tipo, msg) {
    logs.push({ tipo: tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log('[gerarAtaDesdeExtracoes][' + tipo + '] ' + msg);
  };

  try {
    const pontosConsolidados = dados.pontosConsolidados || '';
    const titulo = dados.titulo || 'Reunião';
    const participantes = dados.participantes || '';

    log('INFO', '📋 Gerando ATA final SEGMENTADA a partir de ' + pontosConsolidados.length + ' chars de pontos-chave...');
    log('INFO', '  🤖 Modelo: ' + CONFIG_PROMPTS_REUNIAO.ATA.modelo);
    log('INFO', '  🧠 Pensamento: ' + (CONFIG_PROMPTS_REUNIAO.ATA.pensamento === 0 ? 'DESABILITADO' : 'HABILITADO'));

    const tempoInicio = Date.now();
    const chaveApi = obterChaveGeminiProjeto();

    const dadosAudio = { titulo: titulo, participantes: participantes };

    const resultado = gerarAtaSegmentada(pontosConsolidados, dadosAudio, chaveApi, true);

    const tempoSeg = ((Date.now() - tempoInicio) / 1000).toFixed(1);

    if (!resultado.sucesso) {
      log('ERRO', '❌ ATA final falhou (' + tempoSeg + 's): ' + resultado.mensagem);
      return { sucesso: false, logs: logs, mensagem: resultado.mensagem };
    }

    const tamanhoAta = resultado.ata ? resultado.ata.length : 0;
    log('SUCESSO', '✅ ATA final segmentada gerada em ' + tempoSeg + 's! (' + tamanhoAta + ' chars)');

    // ✅ NOVO: Alerta se ATA parecer suspeita
    if (tamanhoAta > 100000) {
      log('ALERTA', '⚠️ ATA muito grande (' + tamanhoAta + ' chars). Verifique se há repetições.');
    }

    return {
      sucesso: true,
      logs: logs,
      ata: resultado.ata,
      sugestoes: resultado.sugestoes || ''
    };

  } catch (erro) {
    log('ERRO', '❌ ' + erro.message);
    return { sucesso: false, logs: logs, mensagem: erro.message };
  }
}

function extrairSugestoesDoContexto(transcricao) {
  const sugestoes = [];
  const palavrasProblema = ['problema', 'dificuldade', 'bloqueio', 'atraso', 'pendente'];
  palavrasProblema.forEach(palavra => {
    if (transcricao.toLowerCase().includes(palavra)) {
      sugestoes.push(`⚠️ Foram mencionados "${palavra}s" na reunião - verificar acompanhamento`);
    }
  });
  if (transcricao.match(/até (segunda|terça|quarta|quinta|sexta|sábado|domingo)/i)) {
    sugestoes.push('📅 Prazos com dias da semana foram mencionados - criar lembretes');
  }
  return sugestoes.length > 0 ? sugestoes.join('\n') : '';
}


function executarEtapaIdentificacaoAlteracoes(transcricao, contexto, chaveApi, tituloReuniao) {
  try {
    const setoresExistentes = obterSetoresParaContexto();

    // ✅ OTIMIZAÇÃO: se a transcrição é muito longa, usa apenas partes do início e fim
    let transcricaoParaRelatorio = transcricao;
    if (transcricao.length > 100000) {
      transcricaoParaRelatorio =
        transcricao.substring(0, 50000) +
        '\n\n[... parte central omitida por tamanho ...]\n\n' +
        transcricao.substring(transcricao.length - 50000);
      Logger.log(`Relatório: transcrição truncada de ${transcricao.length} para ${transcricaoParaRelatorio.length} chars`);
    }

    const promptRelatorio = `Você é um analista de processos especializado em identificar projetos e atividades a partir de reuniões.

## ⚠️ REGRA ABSOLUTA — PROJETO ÚNICO:

Esta reunião trata de UM ÚNICO PROJETO. Toda discussão gira em torno de resolver um problema central.
Você DEVE identificar e gerar EXATAMENTE 1 (UM) projeto no relatório, não importa quantos temas
pareçam diferentes — eles são facetas do mesmo projeto central.

Se a reunião discutiu várias funcionalidades ou sub-temas, eles devem se tornar
ATIVIDADES do projeto único, não projetos separados.

Pergunte-se: "Qual é o projeto que resume toda esta reunião em poucas palavras?"
Essa é a resposta. Um projeto. Ponto final.

## ⚠️ REGRA FUNDAMENTAL — APENAS O QUE FOI DISCUTIDO:

Você DEVE identificar APENAS o que foi REALMENTE DISCUTIDO nesta reunião.
Não liste atividades existentes que NÃO foram mencionadas na transcrição.
O contexto existe APENAS para verificar se o projeto/atividade já existe.

## COMO DECIDIR O QUE INCLUIR:

### O Projeto Único:
1. Leia a transcrição inteira
2. Identifique O PROBLEMA CENTRAL que motivou esta reunião
3. Verifique se esse projeto já existe no contexto:
   - Se SIM → liste como EXISTENTE com ID real
   - Se NÃO → crie NOVO projeto com ID no formato proj_001
4. O nome do projeto deve ser autoexplicativo (ex: "Automação do Processo de Aprovação de Pagamentos da Pós-graduação")

### As Atividades:
- São as AÇÕES CONCRETAS que precisam ser executadas para concluir o projeto
- Pense: "O que precisa ser FEITO, passo a passo, para resolver o que foi discutido?"
- Cada atividade deve ser uma TAREFA EXECUTÁVEL com resultado mensurável
- Mínimo 3, máximo 12 atividades (todas as ações relevantes discutidas na reunião)
- Não liste atividades existentes que não foram mencionadas

### O Setor:
- É o setor/unidade que será BENEFICIADO, NÃO o executor
- Exemplo: BI cria dashboard para RH → setor é RH (beneficiado), não BI
- Use setores existentes quando possível

---

## 🧮 SISTEMA DE CÁLCULO DE PRIORIDADE — MUITO IMPORTANTE!

Os campos GRAVIDADE, URGENCIA, TIPO, PARA_QUEM e ESFORCO são usados para calcular
automaticamente o VALOR_PRIORIDADE usando a fórmula:

  VALOR_PRIORIDADE = GRAVIDADE × URGENCIA × TIPO × PARA_QUEM × ESFORCO

Use os textos EXATAMENTE como nas tabelas abaixo (o sistema faz lookup por string):

### GRAVIDADE — "O que acontece se este projeto não for feito?"
| Valor a usar (exato) | Peso | Quando usar |
|---|---|---|
| Crítico - Não é possível cumprir as atividades | 5 | Paralisa operações, sem alternativa |
| Alto - É possível cumprir parcialmente | 4 | Impacto severo, há alternativa precária |
| Médio - É possível mas demora muito | 3 | Impacto moderado, há alternativa razoável |

### URGENCIA — "Para quando este projeto precisa estar pronto?"
| Valor a usar (exato) | Peso | Quando usar |
|---|---|---|
| Imediata - Executar imediatamente | 5 | Prazo vencido ou crítico agora |
| Muito urgente - Prazo curto (5 dias) | 4 | Precisa ser feito em até 5 dias |
| Urgente - Curto prazo (10 dias) | 3 | Precisa ser feito em até 10 dias |
| Pouco urgente - Mais de 10 dias | 2 | Prazo confortável |
| Pode esperar | 1 | Sem prazo definido ou muito distante |

### TIPO — "Qual é a natureza deste projeto?"
| Valor a usar (exato) | Peso | Quando usar |
|---|---|---|
| Correção | 5 | Corrigir algo que está errado ou quebrado |
| Nova Implementação | 4 | Criar algo que ainda não existe |
| Melhoria | 3 | Melhorar algo que já funciona |

### PARA_QUEM — "Quem solicitou ou será beneficiado?"
| Valor a usar (exato) | Peso | Quando usar |
|---|---|---|
| Diretoria | 5 | Solicitado pela diretoria ou beneficia diretamente a diretoria |
| Demais áreas | 4 | Beneficia outras áreas operacionais |

### ESFORCO — "Quanto tempo de desenvolvimento este projeto exige?"
| Valor a usar (exato) | Peso | Quando usar |
|---|---|---|
| 1 turno ou menos (4 horas) | 5 | Resolve em meio período |
| 1 dia ou menos (8 horas) | 4 | Resolve em até um dia |
| uma semana (40h) | 3 | Resolve em até uma semana |
| mais de uma semana (40h) | 2 | Projeto longo, mais de uma semana |

### Escala de classificação do VALOR_PRIORIDADE resultante:
- 🔴 ALTA: resultado ≥ 2102
- 🟡 MÉDIA: resultado entre 1078 e 2101
- 🟢 BAIXA: resultado ≤ 1077

Exemplo: Correção(5) × Imediata(5) × Crítico(5) × Diretoria(5) × 1 turno(5) = 3125 → ALTA

---

## CONTEXTO DOS SETORES EXISTENTES:
\`\`\`json
${JSON.stringify(setoresExistentes, null, 2)}
\`\`\`

## CONTEXTO DOS PROJETOS EXISTENTES (use APENAS para verificar se já existe):
\`\`\`json
${JSON.stringify(contexto.projetos, null, 2)}
\`\`\`

## ATIVIDADES EXISTENTES (use APENAS para verificar se já existe):
\`\`\`json
${JSON.stringify(contexto.etapas, null, 2)}
\`\`\`

## RESPONSÁVEIS DA EQUIPE:
\`\`\`json
${JSON.stringify(contexto.responsaveis, null, 2)}
\`\`\`

## TRANSCRIÇÃO DA REUNIÃO:
${transcricaoParaRelatorio}

## INSTRUÇÕES PARA O RELATÓRIO:

### 1. ANALISE PRIMEIRO (antes de escrever qualquer coisa):
- Qual é o PROBLEMA CENTRAL desta reunião? (será o nome do projeto)
- Qual SETOR/UNIDADE será beneficiado?
- Quais são as ATIVIDADES CONCRETAS discutidas para resolver este problema?
- Analise GRAVIDADE, URGENCIA, TIPO, PARA_QUEM e ESFORCO do projeto

### 2. REGRAS PARA O PROJETO:
- Crie um nome DESCRITIVO e ESPECÍFICO
- O projeto deve resolver o PROBLEMA CENTRAL identificado na reunião
- APENAS 1 projeto — sem exceções
- Preencha os campos de prioridade com os valores EXATOS das tabelas acima
- Calcule VALOR_PRIORIDADE multiplicando os 5 pesos
- Para itens NOVOS, use IDs no formato: proj_001, setor_001, atv_001 (SEM prefixo "NOVO_")
- Para itens EXISTENTES, use: EXISTENTE (ID: xxx)

### 3. REGRAS PARA ATIVIDADES:
- Atividades são AÇÕES para concluir o projeto
- Mínimo 3, máximo 12 atividades por projeto
- Use o ID no formato atv_001, atv_002, etc. (SEM prefixo "NOVO_")
- Foque em atividades PENDENTES (o que ainda precisa ser feito)

### 4. REGRAS PARA SETORES:
- Identifique o setor BENEFICIADO, não o executor
- Use setores existentes quando possível (verifique pelo nome/descrição)
- Só inclua setores com projetos vinculados nesta reunião
- Para setores novos, use ID no formato: setor_001 (SEM prefixo "NOVO_")

## ESTRUTURA DO RELATÓRIO (use este formato EXATO em Markdown):

# 📊 RELATÓRIO DE IDENTIFICAÇÕES DA REUNIÃO

## 📅 Informações Gerais

- **Data da Análise:** [data atual no formato DD/MM/AAAA]
- **Baseado na Transcrição:** Sim
- **Tema Central da Reunião:** [descreva em 1 frase curta e direta o problema central — este texto será usado no nome do arquivo]
- **Setor Beneficiado Principal:** [nome do setor]
- **Total de Projetos Identificados:** 1
- **Total de Atividades Identificadas:** [número]

---

## 🏢 SETOR BENEFICIADO

### [Nome do Setor]

| Campo | Valor |
|-------|-------|
| ID | [EXISTENTE (ID: xxx) OU setor_001] |
| NOME | [nome do setor] |
| DESCRICAO | [descrição do setor] |
| RESPONSAVEIS_IDS | [IDs dos responsáveis separados por vírgula] |

**Justificativa:** [Por que este setor é o beneficiado]

---

## 📁 PROJETOS IDENTIFICADOS

### [Nome Descritivo do Projeto Único]

| Campo | Valor |
|-------|-------|
| ID | [EXISTENTE (ID: xxx) OU proj_001] |
| NOME | [nome descritivo e específico] |
| DESCRICAO | [descrição detalhada do que o projeto resolve, baseada na discussão] |
| TIPO | [Correção / Nova Implementação / Melhoria — use exatamente um destes] |
| PARA_QUEM | [Diretoria / Demais áreas — use exatamente um destes] |
| STATUS | [A Fazer / Em Andamento / Concluída] |
| PRIORIDADE | [Alta / Média / Baixa conforme análise] |
| LINK | [N/A ou URL se mencionado] |
| GRAVIDADE | [use exatamente um dos labels da tabela de GRAVIDADE acima] |
| URGENCIA | [use exatamente um dos labels da tabela de URGENCIA acima] |
| ESFORCO | [use exatamente um dos labels da tabela de ESFORCO acima] |
| SETOR | [ID ou nome do setor beneficiado] |
| PILAR | [pilar estratégico relacionado] |
| RESPONSAVEIS_IDS | [IDs dos responsáveis pela execução separados por vírgula] |
| VALOR_PRIORIDADE | [** GRAVIDADE([peso]) × URGENCIA([peso]) × TIPO([peso]) × PARA_QUEM([peso]) × ESFORCO([peso]) = [resultado] → [Alta/Média/Baixa]] |
| DATA_INICIO | [DD/MM/AAAA se mencionada, senão deixar vazio] |
| DATA_FIM | [DD/MM/AAAA se mencionada, senão deixar vazio] |

**Cálculo de prioridade:** GRAVIDADE([peso]) × URGENCIA([peso]) × TIPO([peso]) × PARA_QUEM([peso]) × ESFORCO([peso]) = [resultado] → [Alta/Média/Baixa]

⚠️ **REGRA CRÍTICA DO VALOR_PRIORIDADE:** O número no campo VALOR_PRIORIDADE da tabela acima DEVE SER IDÊNTICO ao resultado da multiplicação mostrada na linha "Cálculo de prioridade". Se o cálculo dá 384, o campo DEVE ser 384. NUNCA coloque um valor diferente do resultado real da multiplicação.

**O que motivou este projeto na reunião:** [Cite 2-3 trechos ou situações específicas da transcrição que justificam este projeto]

---

## 📋 ATIVIDADES IDENTIFICADAS

### [Nome da Atividade — deve ser uma AÇÃO concreta]

| Campo | Valor |
|-------|-------|
| ID | [EXISTENTE (ID: xxx) OU atv_001] |
| PROJETO_ID | [ID do projeto acima] |
| RESPONSAVEIS_IDS | [IDs dos responsáveis separados por vírgula] |
| NOME | [nome da atividade — verbo no infinitivo, ex: "Levantar requisitos do sistema"] |
| O_QUE_FAZER | [instruções claras e objetivas de como executar esta atividade] |
| STATUS | [A Fazer / Em Andamento / Bloqueada / Concluída] |

**Justificativa:** [Por que esta atividade é necessária para o projeto]

[repita o bloco acima para cada atividade identificada]

---

## ⚠️ CHECKLIST FINAL (verifique antes de entregar):
- [ ] O relatório tem EXATAMENTE 1 projeto?
- [ ] O projeto foi realmente discutido na transcrição?
- [ ] As atividades são AÇÕES CONCRETAS (não descrições)?
- [ ] O setor identificado é o BENEFICIADO (não o executor)?
- [ ] O projeto tem entre 3-12 atividades relevantes?
- [ ] Os campos GRAVIDADE, URGENCIA, TIPO, PARA_QUEM e ESFORCO usam os textos EXATOS das tabelas?
- [ ] O VALOR_PRIORIDADE foi calculado corretamente (multiplicação dos 5 pesos)?
- [ ] O VALOR_PRIORIDADE na tabela é IDÊNTICO ao resultado mostrado no "Cálculo de prioridade"?
- [ ] Os IDs de itens novos usam formato limpo (proj_001, atv_001, setor_001) SEM prefixo "NOVO_"?
- [ ] O campo "Tema Central da Reunião" é uma frase curta e direta?
- [ ] Os campos DATA_INICIO e DATA_FIM foram preenchidos quando mencionados na transcrição?

Gere o relatório completo em Markdown:`;

    const urlApi = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG_PROMPTS_REUNIAO.ALTERACOES.modelo}:generateContent?key=${chaveApi}`;

    const corpoRequisicao = {
      contents: [{ parts: [{ text: promptRelatorio }] }],
      generationConfig: montarConfigGeracao(CONFIG_PROMPTS_REUNIAO.ALTERACOES)
    };

    const resposta = UrlFetchApp.fetch(urlApi, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(corpoRequisicao),
      muteHttpExceptions: true
    });

    const codigoStatus = resposta.getResponseCode();
    if (codigoStatus !== 200) {
      throw new Error(`Erro na API (${codigoStatus}): ${resposta.getContentText().substring(0, 200)}`);
    }

    const respostaJson = JSON.parse(resposta.getContentText());

    // ✅ usa helper que ignora thinking tokens
    const relatorioGerado = extrairTextoRespostaGemini(respostaJson);

    // ✅ salvarRelatorioNoDrive agora usa o título da reunião
    const arquivoRelatorio = salvarRelatorioNoDrive(relatorioGerado, tituloReuniao);
    const contagens = extrairContagensDoRelatorio(relatorioGerado);

    return {
      sucesso: true,
      relatorio: relatorioGerado,
      linkRelatorio: arquivoRelatorio.linkArquivo,
      nomeArquivoRelatorio: arquivoRelatorio.nomeArquivo,
      projetosIdentificados: contagens.projetos,
      etapasIdentificadas: contagens.etapas,
      setoresIdentificados: contagens.setores,
      novosProjetosSugeridos: contagens.novosProjetos,
      novasEtapasSugeridas: contagens.novasEtapas,
      novosSetoresSugeridos: contagens.novosSetores
    };

  } catch (erro) {
    Logger.log('ERRO executarEtapaIdentificacaoAlteracoes: ' + erro.toString());
    return {
      sucesso: false,
      mensagem: erro.message,
      relatorio: '',
      projetosIdentificados: [],
      etapasIdentificadas: [],
      setoresIdentificados: []
    };
  }
}

// =====================================================================
//  ✅ NOVO: Funções separadas para PARALELIZAÇÃO client-side
//  Etapa 3 (ATA) e Relatório rodam SIMULTANEAMENTE no client
// =====================================================================

/** Apenas gera o relatório de identificações (sem salvar reunião) */
function etapa_SoRelatorioIdentificacoes(dados) {
  const logs = [];
  const log = (tipo, msg) => {
    logs.push({ tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log(`[${tipo}] ${msg}`);
  };

  try {
    log('INFO', '🔍 Gerando relatório de identificações...');
    const tempoInicio = Date.now();
    const chaveApi = obterChaveGeminiProjeto();
    const contexto = obterContextoProjetosParaGemini();
    log('INFO', `📊 Contexto: ${contexto.totalProjetos} projetos, ${contexto.totalEtapas} etapas`);

    const resultado = executarEtapaIdentificacaoAlteracoes(dados.transcricao, contexto, chaveApi, dados.titulo || '');

    const tempoSeg = ((Date.now() - tempoInicio) / 1000).toFixed(1);

    if (!resultado.sucesso) {
      log('ALERTA', `⚠️ Relatório não gerado (${tempoSeg}s): ${resultado.mensagem}`);
      return { sucesso: false, logs: logs, mensagem: resultado.mensagem };
    }

    log('SUCESSO', `✅ Relatório gerado em ${tempoSeg}s!`);
    log('INFO', `📊 Projetos: ${Array.isArray(resultado.projetosIdentificados) ? resultado.projetosIdentificados.length : 0} | Novos: ${resultado.novosProjetosSugeridos || 0}`);
    log('INFO', `📋 Etapas: ${Array.isArray(resultado.etapasIdentificadas) ? resultado.etapasIdentificadas.length : 0} | Novas: ${resultado.novasEtapasSugeridas || 0}`);

    return {
      sucesso: true, logs: logs,
      relatorio: resultado.relatorio,
      linkRelatorio: resultado.linkRelatorio,
      nomeArquivoRelatorio: resultado.nomeArquivoRelatorio,
      projetosIdentificados: resultado.projetosIdentificados,
      etapasIdentificadas: resultado.etapasIdentificadas,
      novosProjetosSugeridos: resultado.novosProjetosSugeridos,
      novasEtapasSugeridas: resultado.novasEtapasSugeridas
    };

  } catch (erro) {
    log('ERRO', `❌ ${erro.message}`);
    return { sucesso: false, logs: logs, mensagem: erro.message };
  }
}

/** Apenas salva a reunião na planilha (chamada após ATA e Relatório) */
function etapa_SalvarReuniao(dados) {
  const logs = [];
  const log = (tipo, msg) => {
    logs.push({ tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log(`[${tipo}] ${msg}`);
  };

  try {
    log('INFO', '📊 Salvando reunião na planilha...');

    // ✅ O linkAudio já vem do etapa1 (áudio real salvo no Drive)
    const linkAudioFinal = dados.linkAudio || '';

    if (linkAudioFinal) {
      log('INFO', `🔗 Link do áudio: ${linkAudioFinal.substring(0, 60)}...`);
    } else {
      log('ALERTA', '⚠️ Link do áudio não disponível');
    }

    // Salvar transcrição no Drive
    log('INFO', '📝 Salvando transcrição no Drive...');
    const resultadoTranscricaoDrive = salvarTranscricaoNoDrive(dados.transcricao, dados.titulo);
    if (resultadoTranscricaoDrive.sucesso) {
      log('SUCESSO', `✅ Transcrição salva: ${resultadoTranscricaoDrive.nomeArquivo}`);
    }

    // Salvar reunião
    const reuniaoId = salvarReuniaoNaPlanilha({
      titulo: dados.titulo || 'Reunião ' + new Date().toLocaleDateString('pt-BR'),
      dataInicio: dados.dataInicio || new Date(),
      dataFim: new Date(),
      duracao: dados.duracaoMinutos || 0,
      participantes: dados.participantes || '',
      transcricao: dados.transcricao || '',
      ata: dados.ata || '',
      sugestoesIA: dados.sugestoes || '',
      linkAudio: linkAudioFinal,
      projetosImpactados: '',
      etapasImpactadas: ''
    });

    log('SUCESSO', `✅ Reunião salva com ID: ${reuniaoId}`);

    // Limpar arquivo do Gemini
    if (dados.fileName) {
      try {
        limparArquivoGemini(dados.fileName, obterChaveGeminiProjeto());
        log('INFO', '🗑️ Arquivo temporário do Gemini removido');
      } catch (e) { }
    }

    return { sucesso: true, logs, reuniaoId, linkAudio: linkAudioFinal };

  } catch (erro) {
    log('ERRO', `❌ ${erro.message}`);
    return { sucesso: false, logs, mensagem: erro.message };
  }
}

// =====================================================================
//  FUNÇÕES MANTIDAS (sem alteração significativa)
// =====================================================================

function extrairJsonDaResposta(textoResposta) {
  try {
    let jsonString = textoResposta;
    const regexJson = /```json\s*([\s\S]*?)\s*```/;
    const match = textoResposta.match(regexJson);
    if (match) {
      jsonString = match[1];
    } else {
      const inicioJson = textoResposta.indexOf('{');
      const fimJson = textoResposta.lastIndexOf('}');
      if (inicioJson !== -1 && fimJson !== -1) {
        jsonString = textoResposta.substring(inicioJson, fimJson + 1);
      }
    }
    return JSON.parse(jsonString);
  } catch (erro) {
    Logger.log('ERRO extrairJsonDaResposta: ' + erro.toString());
    return { alteracoes: [], projetosIdentificados: [], etapasIdentificadas: [] };
  }
}

function salvarAudioNoDrive(dadosAudio) {
  try {
    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const audioBase64 = dadosAudio.audioBase64.split(',')[1] || dadosAudio.audioBase64;
    const audioBlob = Utilities.newBlob(
      Utilities.base64Decode(audioBase64),
      dadosAudio.tipoMime || 'audio/webm'
    );
    const extensao = obterExtensaoDoMime(dadosAudio.tipoMime);
    const tituloLimpo = limparTituloParaNomeArquivo(dadosAudio.titulo || 'Reuniao');
    const nomeArquivo = tituloLimpo + extensao;
    audioBlob.setName(nomeArquivo);
    const arquivo = pasta.createFile(audioBlob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      sucesso: true, arquivoId: arquivo.getId(),
      nomeArquivo: nomeArquivo, linkArquivo: arquivo.getUrl()
    };
  } catch (erro) {
    Logger.log('ERRO salvarAudioNoDrive: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function limparTituloParaNomeArquivo(titulo) {
  if (!titulo || titulo.trim() === '') return 'Reuniao';

  var limpo = titulo
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')        // remove acentos
    .replace(/[^a-zA-Z0-9\s]/g, '')         // remove caracteres especiais, mantém espaços
    .trim()
    .replace(/\s+/g, ' ')                   // múltiplos espaços → 1
    .substring(0, 80);                      // máx 80 chars

  return limpo || 'Reuniao';
}

function salvarTranscricaoNoDrive(textoTranscricao, tituloReuniao) {
  try {
    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const tituloLimpo = limparTituloParaNomeArquivo(tituloReuniao || 'Reuniao');
    const nomeArquivo = 'Transcricao ' + tituloLimpo + '.txt';
    const arquivo = pasta.createFile(nomeArquivo, textoTranscricao, MimeType.PLAIN_TEXT);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { sucesso: true, arquivoId: arquivo.getId(), nomeArquivo: nomeArquivo, linkArquivo: arquivo.getUrl() };
  } catch (erro) {
    Logger.log('ERRO salvarTranscricaoNoDrive: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function obterExtensaoDoMime(tipoMime) {
  const mapeamento = {
    'audio/webm': '.webm', 'audio/mp3': '.mp3', 'audio/mpeg': '.mp3',
    'audio/wav': '.wav', 'audio/ogg': '.ogg', 'audio/mp4': '.m4a', 'audio/x-m4a': '.m4a'
  };
  return mapeamento[tipoMime] || CONFIGURACAO_REUNIOES.EXTENSAO_PADRAO;
}

function obterContextoProjetosParaGemini() {
  try {
    const abaProjetos = obterAba(NOME_ABA_PROJETOS);
    const abaEtapas = obterAba(NOME_ABA_ETAPAS);
    const abaResponsaveis = obterAba(NOME_ABA_RESPONSAVEIS);

    const dadosProjetos = abaProjetos.getDataRange().getValues();
    const projetos = [];
    for (let i = 1; i < dadosProjetos.length; i++) {
      if (dadosProjetos[i][COLUNAS_PROJETOS.ID]) {
        projetos.push({
          id: dadosProjetos[i][COLUNAS_PROJETOS.ID],
          nome: dadosProjetos[i][COLUNAS_PROJETOS.NOME],
          descricao: dadosProjetos[i][COLUNAS_PROJETOS.DESCRICAO],
          status: dadosProjetos[i][COLUNAS_PROJETOS.STATUS],
          prioridade: dadosProjetos[i][COLUNAS_PROJETOS.PRIORIDADE]
        });
      }
    }

    const dadosEtapas = abaEtapas.getDataRange().getValues();
    const etapas = [];
    for (let i = 1; i < dadosEtapas.length; i++) {
      if (dadosEtapas[i][COLUNAS_ETAPAS.ID]) {
        etapas.push({
          id: dadosEtapas[i][COLUNAS_ETAPAS.ID],
          projetoId: dadosEtapas[i][COLUNAS_ETAPAS.PROJETO_ID],
          nome: dadosEtapas[i][COLUNAS_ETAPAS.NOME],
          descricao: dadosEtapas[i][COLUNAS_ETAPAS.DESCRICAO],
          status: dadosEtapas[i][COLUNAS_ETAPAS.STATUS],
          pendencias: dadosEtapas[i][COLUNAS_ETAPAS.PENDENCIAS]
        });
      }
    }

    const dadosResponsaveis = abaResponsaveis.getDataRange().getValues();
    const responsaveis = [];
    for (let i = 1; i < dadosResponsaveis.length; i++) {
      if (dadosResponsaveis[i][COLUNAS_RESPONSAVEIS.ID]) {
        responsaveis.push({
          id: dadosResponsaveis[i][COLUNAS_RESPONSAVEIS.ID],
          nome: dadosResponsaveis[i][COLUNAS_RESPONSAVEIS.NOME]
        });
      }
    }

    return { projetos, etapas, responsaveis, totalProjetos: projetos.length, totalEtapas: etapas.length };
  } catch (erro) {
    Logger.log('ERRO obterContextoProjetosParaGemini: ' + erro.toString());
    return { projetos: [], etapas: [], responsaveis: [], totalProjetos: 0, totalEtapas: 0 };
  }
}

function salvarReuniaoNaPlanilha(dadosReuniao) {
  try {
    const aba = obterAba(NOME_ABA_REUNIOES);
    const reuniaoId = gerarId();
    const linha = [
      reuniaoId, dadosReuniao.titulo, dadosReuniao.dataInicio, dadosReuniao.dataFim,
      dadosReuniao.duracao, 'Processada', dadosReuniao.participantes,
      dadosReuniao.transcricao, dadosReuniao.ata, dadosReuniao.sugestoesIA,
      dadosReuniao.linkAudio, '', '', dadosReuniao.projetosImpactados,
      dadosReuniao.etapasImpactadas
    ];
    aba.appendRow(linha);
    return reuniaoId;
  } catch (erro) {
    Logger.log('ERRO salvarReuniaoNaPlanilha: ' + erro.toString());
    throw erro;
  }
}

function atualizarCampoEtapa(etapaId, campo, valor) {
  const aba = obterAba(NOME_ABA_ETAPAS);
  const dados = aba.getDataRange().getValues();
  const mapeamentoCampos = {
    'status': COLUNAS_ETAPAS.STATUS, 'descricao': COLUNAS_ETAPAS.DESCRICAO,
    'pendencias': COLUNAS_ETAPAS.PENDENCIAS, 'oQueFazer': COLUNAS_ETAPAS.O_QUE_FAZER
  };
  const indiceCampo = mapeamentoCampos[campo];
  if (indiceCampo === undefined) return;
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_ETAPAS.ID] === etapaId) {
      aba.getRange(i + 1, indiceCampo + 1).setValue(valor);
      return;
    }
  }
}

function atualizarCampoProjeto(projetoId, campo, valor) {
  const aba = obterAba(NOME_ABA_PROJETOS);
  const dados = aba.getDataRange().getValues();
  const mapeamentoCampos = {
    'status': COLUNAS_PROJETOS.STATUS, 'descricao': COLUNAS_PROJETOS.DESCRICAO,
    'prioridade': COLUNAS_PROJETOS.PRIORIDADE
  };
  const indiceCampo = mapeamentoCampos[campo];
  if (indiceCampo === undefined) return;
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_PROJETOS.ID] === projetoId) {
      aba.getRange(i + 1, indiceCampo + 1).setValue(valor);
      return;
    }
  }
}

function criarNovaEtapaSimples(dadosEtapa) {
  const aba = obterAba(NOME_ABA_ETAPAS);
  const etapaId = gerarId();
  let responsavelId = '';
  if (dadosEtapa.responsavelNome) {
    const abaResp = obterAba(NOME_ABA_RESPONSAVEIS);
    const dadosResp = abaResp.getDataRange().getValues();
    for (let i = 1; i < dadosResp.length; i++) {
      if (dadosResp[i][COLUNAS_RESPONSAVEIS.NOME] &&
        dadosResp[i][COLUNAS_RESPONSAVEIS.NOME].toString().toLowerCase().includes(dadosEtapa.responsavelNome.toLowerCase())) {
        responsavelId = dadosResp[i][COLUNAS_RESPONSAVEIS.ID];
        break;
      }
    }
  }
  const linha = [
    etapaId, dadosEtapa.projetoId, responsavelId,
    dadosEtapa.nome || 'Nova Etapa', dadosEtapa.descricao || '',
    '', '', STATUS_ETAPAS.A_FAZER, 0, 0, false, ''
  ];
  aba.appendRow(linha);
}

function adicionarPendenciaEtapa(etapaId, textoPendencia, urgencia) {
  urgencia = urgencia || 'media';
  const aba = obterAba(NOME_ABA_ETAPAS);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_ETAPAS.ID] === etapaId) {
      const pendenciasAtuais = dados[i][COLUNAS_ETAPAS.PENDENCIAS];
      let listaPendencias = [];
      if (pendenciasAtuais && pendenciasAtuais.toString().trim() !== '') {
        try {
          listaPendencias = JSON.parse(pendenciasAtuais);
          if (!Array.isArray(listaPendencias)) listaPendencias = [];
        } catch (e) { listaPendencias = []; }
      }
      listaPendencias.push({
        urgencia: urgencia, dataCriacao: new Date().toISOString(),
        id: 'pnd_' + gerarId(), texto: textoPendencia, concluido: false
      });
      aba.getRange(i + 1, COLUNAS_ETAPAS.PENDENCIAS + 1).setValue(JSON.stringify(listaPendencias));
      return;
    }
  }
}

function enviarAtaPorEmail(reuniaoId, destinatarios) {
  try {
    if (!destinatarios || destinatarios.length === 0) {
      return { sucesso: false, mensagem: 'Nenhum destinatário informado' };
    }
    const aba = obterAba(NOME_ABA_REUNIOES);
    const dados = aba.getDataRange().getValues();
    let reuniao = null, linhaReuniao = -1;
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_REUNIOES.ID] === reuniaoId) {
        reuniao = {
          id: dados[i][COLUNAS_REUNIOES.ID], titulo: dados[i][COLUNAS_REUNIOES.TITULO],
          dataInicio: dados[i][COLUNAS_REUNIOES.DATA_INICIO], duracao: dados[i][COLUNAS_REUNIOES.DURACAO],
          participantes: dados[i][COLUNAS_REUNIOES.PARTICIPANTES], ata: dados[i][COLUNAS_REUNIOES.ATA],
          linkAudio: dados[i][COLUNAS_REUNIOES.LINK_AUDIO]
        };
        linhaReuniao = i;
        break;
      }
    }
    if (!reuniao) return { sucesso: false, mensagem: 'Reunião não encontrada' };

    const assunto = `📋 Ata da Reunião: ${reuniao.titulo}`;
    const ataHtml = converterMarkdownParaHtmlEmail(reuniao.ata);
    const corpoHtml = `
      <!DOCTYPE html><html><head><meta charset="UTF-8">
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto; padding: 20px; }
        h1 { color: #6366f1; border-bottom: 2px solid #6366f1; padding-bottom: 10px; }
        h2 { color: #4f46e5; margin-top: 25px; } h3 { color: #6366f1; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background-color: #6366f1; color: white; }
        tr:nth-child(even) { background-color: #f9fafb; }
        .info-box { background-color: #f0f9ff; border-left: 4px solid #6366f1; padding: 15px; margin: 20px 0; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 0.9em; color: #666; }
        a { color: #6366f1; text-decoration: none; } a:hover { text-decoration: underline; }
      </style></head><body>
        <div class="info-box">
          <strong>📅 Data:</strong> ${reuniao.dataInicio ? new Date(reuniao.dataInicio).toLocaleString('pt-BR') : '-'}<br>
          <strong>👥 Participantes:</strong> ${reuniao.participantes || '-'}
        </div>
        ${ataHtml}
        <div class="footer"><p>Mensagem automática gerada pelo Smart Meeting.</p></div>
      </body></html>`;

    MailApp.sendEmail({ to: destinatarios.join(','), subject: assunto, htmlBody: corpoHtml });

    const emailsEnviados = destinatarios.join(', ') + ' (' + new Date().toLocaleString('pt-BR') + ')';
    aba.getRange(linhaReuniao + 1, COLUNAS_REUNIOES.EMAILS_ENVIADOS + 1).setValue(emailsEnviados);

    return { sucesso: true, mensagem: `Email enviado para ${destinatarios.length} destinatário(s)`, destinatarios };
  } catch (erro) {
    Logger.log('ERRO enviarAtaPorEmail: ' + erro.toString());
    return { sucesso: false, mensagem: 'Erro ao enviar email: ' + erro.message };
  }
}

function converterMarkdownParaHtmlEmail(markdown) {
  if (!markdown) return '';
  let html = markdown;
  html = html.replace(/^### (.*$)/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.*$)/gm, '<h2>$1</h2>');
  html = html.replace(/^# (.*$)/gm, '<h1>$1</h1>');
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
  html = html.replace(/^\- (.*$)/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>)/gm, '<ul>$1</ul>');
  html = html.replace(/<\/ul>\s*<ul>/g, '');

  const linhas = html.split('\n');
  let dentroTabela = false;
  const novasLinhas = [];
  for (let i = 0; i < linhas.length; i++) {
    const linha = linhas[i];
    if (linha.includes('|')) {
      if (!dentroTabela) { novasLinhas.push('<table>'); dentroTabela = true; }
      const celulas = linha.split('|').map(c => c.trim()).filter(c => c !== '');
      if (celulas.every(c => c.match(/^-+$/))) continue;
      const ehCabecalho = !novasLinhas.some(l => l.includes('<tr>'));
      novasLinhas.push('<tr>');
      celulas.forEach(celula => {
        novasLinhas.push(ehCabecalho ? `<th>${celula}</th>` : `<td>${celula}</td>`);
      });
      novasLinhas.push('</tr>');
    } else {
      if (dentroTabela) { novasLinhas.push('</table>'); dentroTabela = false; }
      novasLinhas.push(linha);
    }
  }
  if (dentroTabela) novasLinhas.push('</table>');
  html = novasLinhas.join('\n');
  html = html.replace(/\n\n/g, '</p><p>');
  html = '<p>' + html + '</p>';
  html = html.replace(/<p>\s*<\/p>/g, '');
  return html;
}

function verificarConfiguracaoReunioes() {
  try {
    let temChave = false;
    try {
      if (typeof temChaveApiConfigurada === 'function') {
        temChave = temChaveApiConfigurada();
      } else {
        const props = PropertiesService.getScriptProperties();
        for (let i = 1; i <= 5; i++) {
          const chave = props.getProperty('GEMINI_API_KEY_' + i);
          if (chave && chave.trim() !== '') { temChave = true; break; }
        }
        const chaveProjeto = props.getProperty('GEMINI_API_KEY_PROJETO_EDITOR');
        if (chaveProjeto && chaveProjeto.trim() !== '') temChave = true;
      }
    } catch (e) { temChave = false; }

    let temPasta = false, nomePasta = '';
    try {
      const idPasta = typeof ID_PASTA_DRIVE_REUNIOES !== 'undefined' ? ID_PASTA_DRIVE_REUNIOES : '';
      if (idPasta) {
        const pasta = DriveApp.getFolderById(idPasta);
        temPasta = true;
        nomePasta = pasta.getName();
      }
    } catch (e) { temPasta = false; }

    return {
      sucesso: true, temChaveApi: temChave, temPastaDrive: temPasta,
      nomePastaDrive: nomePasta,
      modeloGemini: typeof MODELO_GEMINI !== 'undefined' ? MODELO_GEMINI : 'gemini-2.5-flash',
      participantesCadastrados: typeof PARTICIPANTES_CADASTRADOS !== 'undefined' ? PARTICIPANTES_CADASTRADOS : []
    };
  } catch (erro) {
    return { sucesso: false, mensagem: erro.message, temChaveApi: false, temPastaDrive: false, nomePastaDrive: '', modeloGemini: '', participantesCadastrados: [] };
  }
}

function listarReunioesRecentes(limite) {
  limite = limite || 20;
  try {
    if (typeof obterAba !== 'function') return { sucesso: true, reunioes: [] };
    
    const nomeAba = typeof NOME_ABA_REUNIOES !== 'undefined' ? NOME_ABA_REUNIOES : 'Reuniões';
    const colunas = typeof COLUNAS_REUNIOES !== 'undefined' ? COLUNAS_REUNIOES : {
      ID: 0, TITULO: 1, DATA_INICIO: 2, DATA_FIM: 3, DURACAO: 4, STATUS: 5,
      PARTICIPANTES: 6, TRANSCRICAO: 7, ATA: 8, SUGESTOES_IA: 9, LINK_AUDIO: 10,
      LINK_ATA: 11, EMAILS_ENVIADOS: 12
    };
    
    const aba = obterAba(nomeAba);
    if (!aba || aba.getLastRow() <= 1) return { sucesso: true, reunioes: [] };
    
    const dados = aba.getDataRange().getValues();
    const reunioes = [];
    
    for (let i = dados.length - 1; i >= 1 && reunioes.length < limite; i--) {
      const idCelula = dados[i][colunas.ID];
      
      // ✅ FIX: Verifica se tem ID válido (qualquer tipo)
      if (idCelula !== null && idCelula !== undefined && idCelula.toString().trim() !== '') {
        
        // ✅ FIX: Converte para string antes de verificar
        const ataTexto = dados[i][colunas.ATA] ? dados[i][colunas.ATA].toString().trim() : '';
        const transcricaoTexto = dados[i][colunas.TRANSCRICAO] ? dados[i][colunas.TRANSCRICAO].toString().trim() : '';
        
        const temAta = ataTexto.length > 10;
        const temTranscricao = transcricaoTexto.length > 10;
        
        reunioes.push({
          id: idCelula.toString().trim(),
          titulo: dados[i][colunas.TITULO] ? dados[i][colunas.TITULO].toString() : '',
          dataInicio: dados[i][colunas.DATA_INICIO],
          duracao: dados[i][colunas.DURACAO],
          status: dados[i][colunas.STATUS] ? dados[i][colunas.STATUS].toString() : '',
          participantes: dados[i][colunas.PARTICIPANTES] ? dados[i][colunas.PARTICIPANTES].toString() : '',
          linkAudio: dados[i][colunas.LINK_AUDIO] ? dados[i][colunas.LINK_AUDIO].toString() : '',
          temAta: temAta,
          temTranscricao: temTranscricao,
          emailsEnviados: dados[i][colunas.EMAILS_ENVIADOS] ? dados[i][colunas.EMAILS_ENVIADOS].toString() : ''
        });
      }
    }
    
    return { sucesso: true, reunioes: reunioes };
    
  } catch (erro) {
    Logger.log('ERRO listarReunioesRecentes: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message, reunioes: [] };
  }
}

/**
 * Retorna metadados leves da reunião (sem ata/transcrição).
 * Chamada rápida e leve para abrir o modal.
 */
function obterMetadadosReuniao(reuniaoId) {
  try {
    if (!reuniaoId) return { sucesso: false, mensagem: 'ID não fornecido' };

    const aba = obterAba(NOME_ABA_REUNIOES);
    if (!aba) return { sucesso: false, mensagem: 'Aba não encontrada' };

    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) return { sucesso: false, mensagem: 'Nenhuma reunião cadastrada' };

    const idBuscado = String(reuniaoId).trim();

    // Busca APENAS coluna de IDs — não carrega dados pesados
    const idsColuna = aba.getRange(2, COLUNAS_REUNIOES.ID + 1, ultimaLinha - 1, 1).getValues();

    let linhaEncontrada = -1;
    for (let i = 0; i < idsColuna.length; i++) {
      if (String(idsColuna[i][0]).trim() === idBuscado) {
        linhaEncontrada = i + 2;
        break;
      }
    }

    if (linhaEncontrada === -1) {
      Logger.log('obterMetadadosReuniao: ID "' + idBuscado + '" NÃO encontrado');
      return { sucesso: false, mensagem: 'Reunião não encontrada' };
    }

    // Carrega APENAS as colunas leves (ID até Participantes = colunas 1-7, e LinkAudio = coluna 11)
    const dadosLeves = aba.getRange(linhaEncontrada, 1, 1, 15).getValues()[0];

    // Verifica se tem ata/transcrição sem carregar o conteúdo completo
    const ataRaw = dadosLeves[COLUNAS_REUNIOES.ATA] ? String(dadosLeves[COLUNAS_REUNIOES.ATA]) : '';
    const transcricaoRaw = dadosLeves[COLUNAS_REUNIOES.TRANSCRICAO] ? String(dadosLeves[COLUNAS_REUNIOES.TRANSCRICAO]) : '';

    return {
      sucesso: true,
      reuniao: {
        id: idBuscado,
        titulo: dadosLeves[COLUNAS_REUNIOES.TITULO] ? String(dadosLeves[COLUNAS_REUNIOES.TITULO]) : '',
        dataInicio: dadosLeves[COLUNAS_REUNIOES.DATA_INICIO],
        dataFim: dadosLeves[COLUNAS_REUNIOES.DATA_FIM],
        duracao: dadosLeves[COLUNAS_REUNIOES.DURACAO],
        status: dadosLeves[COLUNAS_REUNIOES.STATUS] ? String(dadosLeves[COLUNAS_REUNIOES.STATUS]) : '',
        participantes: dadosLeves[COLUNAS_REUNIOES.PARTICIPANTES] ? String(dadosLeves[COLUNAS_REUNIOES.PARTICIPANTES]) : '',
        linkAudio: dadosLeves[COLUNAS_REUNIOES.LINK_AUDIO] ? String(dadosLeves[COLUNAS_REUNIOES.LINK_AUDIO]) : '',
        sugestoesIA: dadosLeves[COLUNAS_REUNIOES.SUGESTOES_IA] ? String(dadosLeves[COLUNAS_REUNIOES.SUGESTOES_IA]) : '',
        temAta: ataRaw.length > 10,
        temTranscricao: transcricaoRaw.length > 10,
        tamanhoAta: ataRaw.length,
        tamanhoTranscricao: transcricaoRaw.length
      }
    };

  } catch (erro) {
    Logger.log('ERRO obterMetadadosReuniao: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

/**
 * Retorna APENAS o conteúdo de um campo específico (ata ou transcricao).
 * Chamada separada para não estourar o limite do google.script.run.
 */
function obterConteudoReuniao(reuniaoId, campo) {
  try {
    if (!reuniaoId || !campo) return { sucesso: false, mensagem: 'Parâmetros inválidos' };

    const aba = obterAba(NOME_ABA_REUNIOES);
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) return { sucesso: false, mensagem: 'Sem dados' };

    const idBuscado = String(reuniaoId).trim();
    const idsColuna = aba.getRange(2, COLUNAS_REUNIOES.ID + 1, ultimaLinha - 1, 1).getValues();

    let linhaEncontrada = -1;
    for (let i = 0; i < idsColuna.length; i++) {
      if (String(idsColuna[i][0]).trim() === idBuscado) {
        linhaEncontrada = i + 2;
        break;
      }
    }

    if (linhaEncontrada === -1) return { sucesso: false, mensagem: 'Reunião não encontrada' };

    // Busca APENAS a coluna solicitada
    const indiceCampo = campo === 'ata' ? COLUNAS_REUNIOES.ATA : COLUNAS_REUNIOES.TRANSCRICAO;
    const valorCelula = aba.getRange(linhaEncontrada, indiceCampo + 1).getValue();
    let conteudo = valorCelula ? String(valorCelula) : '';

    // Trunca se necessário (limite seguro para google.script.run)
    const LIMITE_CHARS = 120000;
    let truncado = false;
    if (conteudo.length > LIMITE_CHARS) {
      conteudo = conteudo.substring(0, LIMITE_CHARS) +
        '\n\n[... ' + campo + ' truncada para exibição. Total: ' + String(valorCelula).length + ' caracteres ...]';
      truncado = true;
    }

    return { sucesso: true, conteudo: conteudo, truncado: truncado };

  } catch (erro) {
    Logger.log('ERRO obterConteudoReuniao: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}


function obterSetoresParaContexto() {
  try {
    const abaSetores = obterAba(NOME_ABA_SETORES);
    if (!abaSetores || abaSetores.getLastRow() <= 1) return [];
    const dados = abaSetores.getDataRange().getValues();
    const setores = [];
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_SETORES.ID]) {
        setores.push({
          id: dados[i][COLUNAS_SETORES.ID], nome: dados[i][COLUNAS_SETORES.NOME],
          descricao: dados[i][COLUNAS_SETORES.DESCRICAO] || '',
          responsaveisIds: dados[i][COLUNAS_SETORES.RESPONSAVEIS_IDS] || ''
        });
      }
    }
    return setores;
  } catch (erro) {
    Logger.log('ERRO obterSetoresParaContexto: ' + erro.toString());
    return [];
  }
}

function extrairContagensDoRelatorio(relatorio) {
  const contagens = {
    projetos: [],
    etapas: [],
    setores: [],
    novosProjetos: 0,
    novasEtapas: 0,
    novosSetores: 0
  };

  try {
    let match;

    // --- Projetos existentes ---
    const secaoProjetos = relatorio.match(/## 📁 PROJETOS IDENTIFICADOS[\s\S]*?(?=## 📋|## ❌|$)/);
    if (secaoProjetos) {
      const regexProj = /EXISTENTE \(ID: ([^)]+)\)/g;
      while ((match = regexProj.exec(secaoProjetos[0])) !== null) {
        if (!contagens.projetos.includes(match[1])) contagens.projetos.push(match[1]);
      }
    }

    // --- Etapas/Atividades existentes ---
    const secaoEtapas = relatorio.match(/## 📋 (?:ETAPAS|ATIVIDADES) IDENTIFICAD[AO]S[\s\S]*?(?=## ❌|$)/);
    if (secaoEtapas) {
      const regexEtp = /EXISTENTE \(ID: ([^)]+)\)/g;
      while ((match = regexEtp.exec(secaoEtapas[0])) !== null) {
        if (!contagens.etapas.includes(match[1])) contagens.etapas.push(match[1]);
      }
    }

    // --- Setores existentes ---
    const secaoSetores = relatorio.match(/## 🏢 SETOR(?:ES)? (?:BENEFICIADO|COM PROJETOS)[\s\S]*?(?=## 📁|$)/);
    if (secaoSetores) {
      const regexSet = /EXISTENTE \(ID: ([^)]+)\)/g;
      while ((match = regexSet.exec(secaoSetores[0])) !== null) {
        if (!contagens.setores.includes(match[1])) contagens.setores.push(match[1]);
      }
    }

    // --- Novos (IDs sem prefixo "NOVO_") ---
    // ✅ CORRIGIDO: reconhece proj_001, atv_001, setor_001 SEM "NOVO_"
    const matchNP = relatorio.match(/\bproj_\d+\b/g);
    if (matchNP) contagens.novosProjetos = new Set(matchNP).size;

    const matchNE = relatorio.match(/\b(?:etp|atv)_\d+\b/g);
    if (matchNE) contagens.novasEtapas = new Set(matchNE).size;

    const matchNS = relatorio.match(/\bsetor_\d+\b/g);
    if (matchNS) contagens.novosSetores = new Set(matchNS).size;

  } catch (erro) {
    Logger.log('Erro contagens: ' + erro.toString());
  }

  return contagens;
}

function salvarRelatorioNoDrive(conteudoRelatorio) {
  try {
    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMdd_HHmmss');
    const nomeArquivo = `Relatorio_Identificacoes_${timestamp}.md`;
    const arquivo = pasta.createFile(nomeArquivo, conteudoRelatorio, MimeType.PLAIN_TEXT);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { sucesso: true, arquivoId: arquivo.getId(), nomeArquivo, linkArquivo: arquivo.getUrl() };
  } catch (erro) {
    Logger.log('ERRO salvarRelatorioNoDrive: ' + erro.toString());
    return { sucesso: false, nomeArquivo: '', linkArquivo: '' };
  }
}

// =====================================================================
//  FUNÇÕES DE UPLOAD EM CHUNKS (mantidas)
// =====================================================================

function receberChunkAudio(dadosChunk) {
  try {
    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const nomeArquivoChunk = `chunk_${dadosChunk.idUpload}_${dadosChunk.indiceChunk}`;
    pasta.createFile(nomeArquivoChunk, dadosChunk.chunkBase64, MimeType.PLAIN_TEXT);
    Logger.log(`Chunk ${dadosChunk.indiceChunk + 1}/${dadosChunk.totalChunks} salvo: ${nomeArquivoChunk}`);
    return { sucesso: true, mensagem: `Chunk ${dadosChunk.indiceChunk + 1}/${dadosChunk.totalChunks} recebido` };
  } catch (erro) {
    Logger.log('ERRO receberChunkAudio: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function uploadParaFileApiGemini(idArquivoDrive, tipoMime, chaveApi) {
  try {
    const arquivo = DriveApp.getFileById(idArquivoDrive);
    const blob = arquivo.getBlob();
    const tamanhoBytes = blob.getBytes().length;
    const tamanhoMB = tamanhoBytes / 1024 / 1024;
    Logger.log(`Upload File API (resumível): ${arquivo.getName()} (${tamanhoMB.toFixed(2)} MB)`);

    // ✅ FIX: SEMPRE usa upload resumível — evita o erro
    // "Metadata part is too large" que ocorre com multipart/related direto.
    // O upload resumível separa metadados dos dados binários e funciona
    // para QUALQUER tamanho de arquivo (de 1 KB a 2 GB).
    return uploadResumivelFileApi(blob, tipoMime, arquivo.getName(), chaveApi, tamanhoBytes);
  } catch (erro) {
    Logger.log('ERRO uploadParaFileApiGemini: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function uploadDiretoFileApi(blob, tipoMime, nomeArquivo, chaveApi) {
  try {
    const urlUpload = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${chaveApi}`;
    const metadados = JSON.stringify({ file: { display_name: nomeArquivo } });
    const boundary = 'BOUNDARY_' + Date.now();
    const bytesArquivo = blob.getBytes();
    const parteMetadata = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + metadados + '\r\n';
    const parteArquivoHeader = '--' + boundary + '\r\nContent-Type: ' + tipoMime + '\r\n\r\n';
    const parteEncerramento = '\r\n--' + boundary + '--';
    const bytesMetadata = Utilities.newBlob(parteMetadata).getBytes();
    const bytesArquivoHeader = Utilities.newBlob(parteArquivoHeader).getBytes();
    const bytesEncerramento = Utilities.newBlob(parteEncerramento).getBytes();
    const corpoCompleto = [].concat(bytesMetadata, bytesArquivoHeader, bytesArquivo, bytesEncerramento);
    const resposta = UrlFetchApp.fetch(urlUpload, {
      method: 'post', contentType: 'multipart/related; boundary=' + boundary,
      payload: corpoCompleto, muteHttpExceptions: true
    });
    if (resposta.getResponseCode() !== 200) {
      throw new Error(`Erro no upload (${resposta.getResponseCode()}): ${resposta.getContentText().substring(0, 300)}`);
    }
    const respostaJson = JSON.parse(resposta.getContentText());
    aguardarProcessamentoArquivoGemini(respostaJson.file.name, chaveApi);
    return { sucesso: true, fileUri: respostaJson.file.uri, fileName: respostaJson.file.name };
  } catch (erro) {
    Logger.log('ERRO uploadDiretoFileApi: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function uploadResumivelFileApi(blob, tipoMime, nomeArquivo, chaveApi, tamanhoTotal) {
  try {
    const TAMANHO_PARTE = 40 * 1024 * 1024;
    const bytesArquivo = blob.getBytes();
    const urlIniciar = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${chaveApi}`;
    const respInicio = UrlFetchApp.fetch(urlIniciar, {
      method: 'post', contentType: 'application/json',
      headers: {
        'X-Goog-Upload-Protocol': 'resumable', 'X-Goog-Upload-Command': 'start',
        'X-Goog-Upload-Header-Content-Length': tamanhoTotal.toString(),
        'X-Goog-Upload-Header-Content-Type': tipoMime
      },
      payload: JSON.stringify({ file: { display_name: nomeArquivo } }),
      muteHttpExceptions: true
    });
    if (respInicio.getResponseCode() !== 200) {
      throw new Error('Falha ao iniciar upload resumível: ' + respInicio.getContentText().substring(0, 200));
    }
    const headersResp = respInicio.getHeaders();
    const urlUpload = headersResp['x-goog-upload-url'] || headersResp['X-Goog-Upload-URL'];
    if (!urlUpload) throw new Error('URL de upload resumível não retornada');

    let offsetAtual = 0, respFinal = null;
    while (offsetAtual < tamanhoTotal) {
      const fimParte = Math.min(offsetAtual + TAMANHO_PARTE, tamanhoTotal);
      const ehUltimaParte = fimParte >= tamanhoTotal;
      const bytesParteAtual = bytesArquivo.slice(offsetAtual, fimParte);

      // ✅ FIX: Removido 'Content-Length' — UrlFetchApp calcula automaticamente
      const respParte = UrlFetchApp.fetch(urlUpload, {
        method: 'post',
        headers: {
          'X-Goog-Upload-Command': ehUltimaParte ? 'upload, finalize' : 'upload',
          'X-Goog-Upload-Offset': offsetAtual.toString()
        },
        payload: bytesParteAtual,
        muteHttpExceptions: true
      });
      if (respParte.getResponseCode() !== 200) {
        throw new Error(`Erro chunk offset ${offsetAtual}: ${respParte.getContentText().substring(0, 200)}`);
      }
      if (ehUltimaParte) respFinal = respParte;
      offsetAtual = fimParte;
    }

    const respostaJson = JSON.parse(respFinal.getContentText());
    aguardarProcessamentoArquivoGemini(respostaJson.file.name, chaveApi);
    return { sucesso: true, fileUri: respostaJson.file.uri, fileName: respostaJson.file.name };
  } catch (erro) {
    Logger.log('ERRO uploadResumivelFileApi: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function aguardarProcessamentoArquivoGemini(fileName, chaveApi) {
  const urlStatus = `https://generativelanguage.googleapis.com/v1beta/${fileName}?key=${chaveApi}`;
  const MAX_TENTATIVAS = 60;
  let tentativa = 0;
  while (tentativa < MAX_TENTATIVAS) {
    const resposta = UrlFetchApp.fetch(urlStatus, { muteHttpExceptions: true });
    if (resposta.getResponseCode() === 200) {
      const dados = JSON.parse(resposta.getContentText());
      if (dados.state === 'ACTIVE') { Logger.log('Arquivo ativo!'); return true; }
      if (dados.state === 'FAILED') throw new Error('Processamento falhou no Gemini');
      Logger.log(`Processando... Estado: ${dados.state} (tentativa ${tentativa + 1})`);
    }
    Utilities.sleep(5000);
    tentativa++;
  }
  throw new Error('Timeout aguardando processamento');
}

function limparArquivoGemini(fileName, chaveApi) {
  try {
    UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/${fileName}?key=${chaveApi}`, {
      method: 'delete', muteHttpExceptions: true
    });
  } catch (e) { Logger.log('Aviso limpeza Gemini: ' + e.toString()); }
}

function obterProjetosParaAssociacao() {
  try {
    const abaProjetos = obterAba(NOME_ABA_PROJETOS);
    if (!abaProjetos || abaProjetos.getLastRow() <= 1) return { sucesso: true, projetos: [] };

    // ── Mapa de responsáveis: ID → Nome ──
    const mapaResponsavelNome = {};
    const abaResp = obterAba(NOME_ABA_RESPONSAVEIS);
    if (abaResp && abaResp.getLastRow() > 1) {
      const dadosResp = abaResp.getDataRange().getValues();
      for (let i = 1; i < dadosResp.length; i++) {
        const idResp = dadosResp[i][COLUNAS_RESPONSAVEIS.ID];
        if (idResp) {
          mapaResponsavelNome[idResp.toString().trim()] = dadosResp[i][COLUNAS_RESPONSAVEIS.NOME]
            ? dadosResp[i][COLUNAS_RESPONSAVEIS.NOME].toString()
            : 'Sem nome';
        }
      }
    }

    const dados = abaProjetos.getDataRange().getValues();
    const projetos = [];

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID]) {
        // ✅ USA parsearIdsColuna — suporta string pura E JSON array
        const listaIds = parsearIdsColuna(dados[i][COLUNAS_PROJETOS.RESPONSAVEIS_IDS]);
        const nomes = [];
        for (let j = 0; j < listaIds.length; j++) {
          const nomeResp = mapaResponsavelNome[listaIds[j]];
          if (nomeResp) nomes.push(nomeResp);
        }

        projetos.push({
          id: dados[i][COLUNAS_PROJETOS.ID],
          nome: dados[i][COLUNAS_PROJETOS.NOME] || 'Sem nome',
          status: dados[i][COLUNAS_PROJETOS.STATUS] || '',
          setor: dados[i][COLUNAS_PROJETOS.SETOR] || '',
          responsavelNomes: nomes
        });
      }
    }

    projetos.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });

    return { sucesso: true, projetos: projetos };
  } catch (e) {
    Logger.log('ERRO obterProjetosParaAssociacao: ' + e.toString());
    return { sucesso: false, projetos: [], mensagem: e.message };
  }
}

function obterReunioesCatalogadas() {
  try {
    const nomeAba = typeof NOME_ABA_REUNIOES !== 'undefined' ? NOME_ABA_REUNIOES : 'Reuniões';
    const aba = obterAba(nomeAba);
    if (!aba || aba.getLastRow() <= 1) {
      return { sucesso: true, porProjeto: {}, semCatalogo: [] };
    }

    const dados = aba.getDataRange().getValues();
    const reunioesTodas = [];

    for (let i = dados.length - 1; i >= 1; i--) {
      const idCelula = dados[i][COLUNAS_REUNIOES.ID];
      if (!idCelula || idCelula.toString().trim() === '') continue;

      const ataTexto = dados[i][COLUNAS_REUNIOES.ATA] ? dados[i][COLUNAS_REUNIOES.ATA].toString().trim() : '';
      const transcricaoTexto = dados[i][COLUNAS_REUNIOES.TRANSCRICAO] ? dados[i][COLUNAS_REUNIOES.TRANSCRICAO].toString().trim() : '';
      const projetoId = dados[i][COLUNAS_REUNIOES.PROJETOS_IMPACTADOS]
        ? dados[i][COLUNAS_REUNIOES.PROJETOS_IMPACTADOS].toString().trim()
        : '';

      reunioesTodas.push({
        id: idCelula.toString().trim(),
        titulo: dados[i][COLUNAS_REUNIOES.TITULO] ? dados[i][COLUNAS_REUNIOES.TITULO].toString() : '',
        dataInicio: dados[i][COLUNAS_REUNIOES.DATA_INICIO],
        duracao: dados[i][COLUNAS_REUNIOES.DURACAO],
        status: dados[i][COLUNAS_REUNIOES.STATUS] ? dados[i][COLUNAS_REUNIOES.STATUS].toString() : '',
        participantes: dados[i][COLUNAS_REUNIOES.PARTICIPANTES] ? dados[i][COLUNAS_REUNIOES.PARTICIPANTES].toString() : '',
        linkAudio: dados[i][COLUNAS_REUNIOES.LINK_AUDIO] ? dados[i][COLUNAS_REUNIOES.LINK_AUDIO].toString() : '',
        emailsEnviados: dados[i][COLUNAS_REUNIOES.EMAILS_ENVIADOS] ? dados[i][COLUNAS_REUNIOES.EMAILS_ENVIADOS].toString() : '',
        temAta: ataTexto.length > 10,
        temTranscricao: transcricaoTexto.length > 10,
        projetoId: projetoId
      });
    }

    // ── Mapa de responsáveis: ID → Nome ──
    const mapaResponsavelNome = {};
    const abaResp = obterAba(NOME_ABA_RESPONSAVEIS);
    if (abaResp && abaResp.getLastRow() > 1) {
      const dadosResp = abaResp.getDataRange().getValues();
      for (let i = 1; i < dadosResp.length; i++) {
        const idResp = dadosResp[i][COLUNAS_RESPONSAVEIS.ID];
        if (idResp) {
          mapaResponsavelNome[idResp.toString().trim()] = dadosResp[i][COLUNAS_RESPONSAVEIS.NOME]
            ? dadosResp[i][COLUNAS_RESPONSAVEIS.NOME].toString()
            : 'Sem nome';
        }
      }
    }

    // ── Buscar nomes dos projetos + responsáveis ──
    const abaProjetos = obterAba(NOME_ABA_PROJETOS);
    const mapaProjetoNome = {};
    const mapaProjetoStatus = {};
    const mapaProjetoRespNomes = {};

    if (abaProjetos && abaProjetos.getLastRow() > 1) {
      const dadosProjetos = abaProjetos.getDataRange().getValues();
      for (let i = 1; i < dadosProjetos.length; i++) {
        if (dadosProjetos[i][COLUNAS_PROJETOS.ID]) {
          const pid = dadosProjetos[i][COLUNAS_PROJETOS.ID].toString();
          mapaProjetoNome[pid] = dadosProjetos[i][COLUNAS_PROJETOS.NOME] || 'Sem nome';
          mapaProjetoStatus[pid] = dadosProjetos[i][COLUNAS_PROJETOS.STATUS] || '';

          // Resolver IDs de responsáveis para nomes
const listaIds = parsearIdsColuna(dadosProjetos[i][COLUNAS_PROJETOS.RESPONSAVEIS_IDS]);
const nomes = [];
for (let j = 0; j < listaIds.length; j++) {
  var nomeResp = mapaResponsavelNome[listaIds[j]];
  if (nomeResp) nomes.push(nomeResp);
}
          mapaProjetoRespNomes[pid] = nomes;
        }
      }
    }

    // ── Agrupar reuniões ──
    const porProjeto = {};
    const semCatalogo = [];

    for (var k = 0; k < reunioesTodas.length; k++) {
      var r = reunioesTodas[k];
      if (r.projetoId && mapaProjetoNome[r.projetoId]) {
        if (!porProjeto[r.projetoId]) {
          porProjeto[r.projetoId] = {
            id: r.projetoId,
            nome: mapaProjetoNome[r.projetoId],
            status: mapaProjetoStatus[r.projetoId],
            responsavelNomes: mapaProjetoRespNomes[r.projetoId] || [],
            reunioes: []
          };
        }
        porProjeto[r.projetoId].reunioes.push(r);
      } else {
        semCatalogo.push(r);
      }
    }

    return { sucesso: true, porProjeto: porProjeto, semCatalogo: semCatalogo };

  } catch (e) {
    Logger.log('ERRO obterReunioesCatalogadas: ' + e.toString());
    return { sucesso: false, mensagem: e.message, porProjeto: {}, semCatalogo: [] };
  }
}

function associarReuniaoAoProjeto(reuniaoId, projetoId) {
  try {
    const aba = obterAba(NOME_ABA_REUNIOES);
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) return { sucesso: false, mensagem: 'Sem reuniões' };

    const idsColuna = aba.getRange(2, COLUNAS_REUNIOES.ID + 1, ultimaLinha - 1, 1).getValues();

    for (let i = 0; i < idsColuna.length; i++) {
      if (String(idsColuna[i][0]).trim() === String(reuniaoId).trim()) {
        const linhaReal = i + 2;
        aba.getRange(linhaReal, COLUNAS_REUNIOES.PROJETOS_IMPACTADOS + 1).setValue(projetoId || '');
        return { sucesso: true, mensagem: 'Associação salva!' };
      }
    }

    return { sucesso: false, mensagem: 'Reunião não encontrada' };
  } catch (e) {
    Logger.log('ERRO associarReuniaoAoProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * ETAPA 1: Consolida chunks do Drive e faz upload para Gemini + Drive.
 * ✅ CORREÇÃO: Acumula chunks decodificados até ≥8MB antes de enviar ao Gemini
 *    (Gemini exige granularidade OBRIGATÓRIA de 8,388,608 bytes nos chunks intermediários)
 * ✅ MEMÓRIA: Nunca cria array JS > ~12MB (3 chunks × 3MB) — sem "Invalid array length"
 */
function etapa1_UploadChunksParaGemini(dados) {
  const logs = [];
  const log = function(tipo, msg) {
    logs.push({ tipo: tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log('[' + tipo + '] ' + msg);
  };

  try {
    log('INFO', '☁️ [ETAPA 1/4] Processando chunks e enviando para Gemini...');
    const tempoInicio = Date.now();
    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const chaveApi = obterChaveGeminiProjeto();
    if (!chaveApi) throw new Error('Chave API do Gemini não configurada.');
    const tipoMime = dados.tipoMime || 'audio/webm';

    // ✅ Gemini exige que todos os chunks intermediários sejam múltiplos exatos de 8MB
    // Drive aceita múltiplos de 256KB; 8MB é múltiplo de 256KB, então usamos a mesma granularidade
    const GRANULARIDADE_GEMINI = 8 * 1024 * 1024; // 8,388,608 bytes — NÃO alterar

    // ── 1. Coletar arquivos de chunk em ordem ──
    const arquivosChunks = [];
    let idx = 0;
    while (true) {
      const it = pasta.getFilesByName('chunk_' + dados.idUpload + '_' + idx);
      if (!it.hasNext()) break;
      arquivosChunks.push(it.next());
      idx++;
    }
    if (arquivosChunks.length === 0) throw new Error('Nenhum chunk encontrado: ' + dados.idUpload);
    log('INFO', '  📦 ' + arquivosChunks.length + ' chunks encontrados');

    // ── 2. Calcular tamanho binário total ──
    // Chunks intermediários: TAMANHO_CHUNK_BYTES=4194304 é divisível por 3 → sem padding base64
    // Fórmula exata: tamBin = tamBase64 × 3/4
    // Último chunk: decodifica para obter tamanho real (considera padding '=')
    let tamBinTotal = 0;
    let ultimoChunkBytes = null;

    for (let i = 0; i < arquivosChunks.length; i++) {
      if (i < arquivosChunks.length - 1) {
        const tamBase64 = arquivosChunks[i].getSize(); // ASCII: 1 char = 1 byte
        tamBinTotal += Math.floor(tamBase64 * 3 / 4);
        log('INFO', '  📄 Chunk ' + i + ': ' + (tamBase64 / 1024 / 1024).toFixed(2) + ' MB base64');
      } else {
        const ultimoBase64 = arquivosChunks[i].getBlob().getDataAsString();
        ultimoChunkBytes = Utilities.base64Decode(ultimoBase64); // Byte[] Java, ~2-3MB
        tamBinTotal += ultimoChunkBytes.length;
        log('INFO', '  📄 Chunk ' + i + ' (último): ' + (ultimoChunkBytes.length / 1024 / 1024).toFixed(2) + ' MB binário');
      }
    }
    log('SUCESSO', '✅ Tamanho total calculado: ' + (tamBinTotal / 1024 / 1024).toFixed(2) + ' MB');

    // ── 3. Iniciar sessão de upload resumível no Gemini ──
    log('INFO', '📤 Iniciando sessão de upload no Gemini...');
    const timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMdd_HHmmss');
    const nomeArquivoGemini = 'Reuniao_' + timestamp;

    const respInicioGemini = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/upload/v1beta/files?key=' + chaveApi, {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'X-Goog-Upload-Protocol': 'resumable',
          'X-Goog-Upload-Command': 'start',
          'X-Goog-Upload-Header-Content-Length': tamBinTotal.toString(),
          'X-Goog-Upload-Header-Content-Type': tipoMime
        },
        payload: JSON.stringify({ file: { display_name: nomeArquivoGemini } }),
        muteHttpExceptions: true
      }
    );

    if (respInicioGemini.getResponseCode() !== 200) {
      throw new Error('Falha Gemini: HTTP ' + respInicioGemini.getResponseCode() +
        ' → ' + respInicioGemini.getContentText().substring(0, 200));
    }
    const headersGemini = respInicioGemini.getHeaders();
    const urlUploadGemini = headersGemini['x-goog-upload-url'] || headersGemini['X-Goog-Upload-URL'];
    if (!urlUploadGemini) throw new Error('URL de upload Gemini não retornada');
    log('SUCESSO', '✅ Sessão Gemini criada');

    // ── 4. Iniciar sessão de upload resumível no Drive via API REST ──
    const extensao = obterExtensaoDoMime(tipoMime);
    const nomeAudioReal = 'Reuniao_' + timestamp + extensao;
    const tokenOAuth = ScriptApp.getOAuthToken();
    let urlUploadDrive = null;
    let idArquivoDrive = null;

    try {
      const respInicioDrive = UrlFetchApp.fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
          method: 'POST',
          headers: {
            'Authorization': 'Bearer ' + tokenOAuth,
            'Content-Type': 'application/json',
            'X-Upload-Content-Type': tipoMime,
            'X-Upload-Content-Length': tamBinTotal.toString()
          },
          payload: JSON.stringify({ name: nomeAudioReal, parents: [ID_PASTA_DRIVE_REUNIOES] }),
          muteHttpExceptions: true
        }
      );

      if (respInicioDrive.getResponseCode() === 200) {
        const hdDrive = respInicioDrive.getHeaders();
        urlUploadDrive = hdDrive['Location'] || hdDrive['location'];
        log(urlUploadDrive ? 'SUCESSO' : 'ALERTA',
          urlUploadDrive ? '✅ Sessão Drive criada' : '⚠️ URL Drive não retornada');
      } else {
        log('ALERTA', '⚠️ Sessão Drive falhou (HTTP ' + respInicioDrive.getResponseCode() + '). Gemini continua normalmente.');
      }
    } catch (eDrive) {
      log('ALERTA', '⚠️ Erro ao criar sessão Drive: ' + eDrive.message + ' (não crítico)');
    }

    // ── 5. ENVIO COM BUFFER ACUMULADOR (granularidade 8MB) ──
    //
    // LÓGICA:
    //   - Cada chunk decodificado tem ~3MB binário (de 4MB base64)
    //   - Acumulamos em bufferLista até ter ≥ 8MB
    //   - Então extraímos EXATAMENTE 8MB e enviamos (Gemini aceita)
    //   - Resto fica no buffer para o próximo ciclo
    //   - Último pacote pode ter qualquer tamanho (vai com 'upload, finalize')
    //
    // MEMÓRIA MÁXIMA: ~9MB em bufferLista (3 chunks de ~3MB) — seguro, nunca cria array de 146MB

    log('INFO', '📤 Enviando para Gemini' + (urlUploadDrive ? ' + Drive' : '') +
      ' (pacotes de 8 MB)...');

    let bufferLista = [];       // Lista de Byte[] Java acumulados aguardando envio
    let tamBufferTotal = 0;     // Bytes acumulados no buffer
    let offsetGemini = 0;       // Offset atual no upload Gemini
    let offsetDrive = 0;        // Offset atual no upload Drive
    let respFinalGemini = null; // Resposta do chunk 'finalize' do Gemini
    let numeroPacote = 0;       // Contador de pacotes enviados ao Gemini

    for (let i = 0; i < arquivosChunks.length; i++) {
      const tChunk = Date.now();
      const ehUltimo = (i === arquivosChunks.length - 1);

      // Decodifica apenas este chunk (~3MB Byte[] Java em memória, nunca um array JS gigante)
      let bytesChunk;
      if (ehUltimo && ultimoChunkBytes !== null) {
        bytesChunk = ultimoChunkBytes; // Já decodificado na etapa 2 (reutiliza)
        ultimoChunkBytes = null;
      } else {
        bytesChunk = Utilities.base64Decode(arquivosChunks[i].getBlob().getDataAsString());
      }

      // Adiciona ao buffer de acumulação
      bufferLista.push(bytesChunk);
      tamBufferTotal += bytesChunk.length;

      // Apaga chunk temporário do Drive imediatamente (libera espaço)
      try { arquivosChunks[i].setTrashed(true); } catch (e) { /* não crítico */ }

      if (ehUltimo) {
        // ── Envio final: todo o restante do buffer (tamanho livre, sem restrição) ──
        numeroPacote++;
        const pacoteFinal = combinarBufferCompleto(bufferLista);
        bufferLista = [];
        tamBufferTotal = 0;

        log('INFO', '  📤 Pacote final ' + numeroPacote + ' (' +
          (pacoteFinal.length / 1024 / 1024).toFixed(2) + ' MB, finalize)...');
        const tEnvio = Date.now();

        const respGeminiFinal = UrlFetchApp.fetch(urlUploadGemini, {
          method: 'post',
          contentType: 'application/octet-stream',
          headers: {
            'X-Goog-Upload-Command': 'upload, finalize',
            'X-Goog-Upload-Offset': offsetGemini.toString()
          },
          payload: pacoteFinal,
          muteHttpExceptions: true
        });

        if (respGeminiFinal.getResponseCode() !== 200) {
          throw new Error('Erro Gemini pacote final: HTTP ' + respGeminiFinal.getResponseCode() +
            ' → ' + respGeminiFinal.getContentText().substring(0, 300));
        }
        respFinalGemini = respGeminiFinal;
        offsetGemini += pacoteFinal.length;

        // Envia último pacote ao Drive
        if (urlUploadDrive) {
          try {
            const fimDrive = offsetDrive + pacoteFinal.length - 1;
            const respDriveFinal = UrlFetchApp.fetch(urlUploadDrive, {
              method: 'PUT',
              headers: { 'Content-Range': 'bytes ' + offsetDrive + '-' + fimDrive + '/' + tamBinTotal },
              payload: pacoteFinal,
              muteHttpExceptions: true
            });
            const stDrive = respDriveFinal.getResponseCode();
            if (stDrive === 200 || stDrive === 201) {
              idArquivoDrive = JSON.parse(respDriveFinal.getContentText()).id;
            } else {
              log('ALERTA', '⚠️ Drive pacote final HTTP ' + stDrive);
            }
            offsetDrive += pacoteFinal.length;
          } catch (eDrive) {
            log('ALERTA', '⚠️ Erro Drive pacote final: ' + eDrive.message);
          }
        }

        log('SUCESSO', '  ✅ Pacote final enviado em ' + ((Date.now() - tEnvio) / 1000).toFixed(1) + 's');

      } else {
        // ── Enquanto buffer ≥ 8MB, extrai e envia exatamente 8MB ──
        while (tamBufferTotal >= GRANULARIDADE_GEMINI) {
          numeroPacote++;
          // Extrai primeiros 8MB do buffer; modifica bufferLista in-place (remove bytes usados)
          const pacote8MB = extrairPrimeirosNBytesDoBuffer(bufferLista, GRANULARIDADE_GEMINI);
          tamBufferTotal -= GRANULARIDADE_GEMINI;

          log('INFO', '  📤 Pacote ' + numeroPacote + ' (8 MB exatos | buffer restante: ' +
            (tamBufferTotal / 1024 / 1024).toFixed(2) + ' MB)...');
          const tEnvio = Date.now();

          const respGemini = UrlFetchApp.fetch(urlUploadGemini, {
            method: 'post',
            contentType: 'application/octet-stream',
            headers: {
              'X-Goog-Upload-Command': 'upload', // NÃO finalizar — há mais dados
              'X-Goog-Upload-Offset': offsetGemini.toString()
            },
            payload: pacote8MB,
            muteHttpExceptions: true
          });

          if (respGemini.getResponseCode() !== 200) {
            throw new Error('Erro Gemini pacote ' + numeroPacote + ': HTTP ' +
              respGemini.getResponseCode() + ' → ' + respGemini.getContentText().substring(0, 300));
          }
          offsetGemini += GRANULARIDADE_GEMINI;

          // Envia mesmo pacote 8MB ao Drive
          if (urlUploadDrive) {
            try {
              const fimDrive = offsetDrive + GRANULARIDADE_GEMINI - 1;
              const respDrive = UrlFetchApp.fetch(urlUploadDrive, {
                method: 'PUT',
                headers: { 'Content-Range': 'bytes ' + offsetDrive + '-' + fimDrive + '/' + tamBinTotal },
                payload: pacote8MB,
                muteHttpExceptions: true
              });
              const stDrive = respDrive.getResponseCode();
              // 308 = "Resume Incomplete" = resposta ESPERADA para chunks intermediários do Drive ✅
              if (stDrive !== 308 && stDrive !== 200 && stDrive !== 201) {
                log('ALERTA', '⚠️ Drive pacote ' + numeroPacote + ' HTTP inesperado: ' + stDrive);
              }
              offsetDrive += GRANULARIDADE_GEMINI;
            } catch (eDrive) {
              log('ALERTA', '⚠️ Erro Drive pacote ' + numeroPacote + ': ' + eDrive.message + ' (Gemini continua)');
            }
          }

          log('SUCESSO', '  ✅ Pacote ' + numeroPacote + ' enviado em ' + ((Date.now() - tEnvio) / 1000).toFixed(1) + 's');
        }
      }

      log('INFO', '  🔄 Chunk ' + (i + 1) + '/' + arquivosChunks.length +
        ' processado (' + ((Date.now() - tChunk) / 1000).toFixed(1) + 's) | buffer: ' +
        (tamBufferTotal / 1024 / 1024).toFixed(2) + ' MB');
    }

    // ── 6. Configurar permissão pública no arquivo Drive e obter link ──
    let linkAudioReal = '';
    if (idArquivoDrive) {
      try {
        UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files/' + idArquivoDrive + '/permissions', {
            method: 'POST',
            headers: { 'Authorization': 'Bearer ' + tokenOAuth, 'Content-Type': 'application/json' },
            payload: JSON.stringify({ role: 'reader', type: 'anyone' }),
            muteHttpExceptions: true
          }
        );
        linkAudioReal = 'https://drive.google.com/file/d/' + idArquivoDrive + '/view';
        log('SUCESSO', '✅ Áudio salvo no Drive: ' + nomeAudioReal);
        log('INFO', '  🔗 ' + linkAudioReal);
      } catch (ePerm) {
        linkAudioReal = 'https://drive.google.com/file/d/' + idArquivoDrive + '/view';
        log('ALERTA', '⚠️ Permissão pública não configurada: ' + ePerm.message);
      }
    } else {
      log('ALERTA', '⚠️ Áudio não salvo no Drive (upload Drive falhou ou não iniciou).');
    }

    // ── 7. Aguardar Gemini processar o arquivo ──
    if (!respFinalGemini) throw new Error('Upload Gemini não foi finalizado corretamente');
    const geminiArquivo = JSON.parse(respFinalGemini.getContentText());
    log('INFO', '⏳ Aguardando Gemini processar o arquivo...');
    log('INFO', '  📌 FileURI: ' + geminiArquivo.file.uri);
    aguardarProcessamentoArquivoGemini(geminiArquivo.file.name, chaveApi);

    const tempoTotal = ((Date.now() - tempoInicio) / 1000).toFixed(1);
    log('SUCESSO', '✅ ETAPA 1 CONCLUÍDA em ' + tempoTotal + 's! (' + numeroPacote + ' pacotes enviados ao Gemini)');

    return {
      sucesso: true, logs: logs,
      fileUri: geminiArquivo.file.uri,
      fileName: geminiArquivo.file.name,
      idUpload: dados.idUpload,
      linkAudio: linkAudioReal,
      nomeAudio: nomeAudioReal
    };

  } catch (erro) {
    log('ERRO', '❌ ' + erro.message);
    Logger.log('[STACK etapa1_UploadChunksParaGemini] ' + (erro.stack || erro.toString()));
    return { sucesso: false, logs: logs, mensagem: erro.message };
  }
}

// ============================================================
//  AUXILIARES DO BUFFER DE ACUMULAÇÃO (etapa1)
// ============================================================

/**
 * Converte um Byte[] Java para array JS de forma eficiente.
 * Array.from() é nativo e muito mais rápido que loop push no V8 (GAS moderno).
 */
function byteArrayParaJsArray(bytes) {
  if (typeof Array.from === 'function') {
    return Array.from(bytes); // V8 runtime (GAS atual) — rápido
  }
  // Fallback Rhino (legado)
  const r = new Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) r[i] = bytes[i];
  return r;
}

/**
 * Extrai exatamente N bytes do início da lista de Byte[] arrays.
 * Modifica bufferLista in-place: remove os bytes usados e mantém o restante.
 * 
 * @param {Array} bufferLista  - Lista de Byte[] Java (ou arrays JS) acumulados
 * @param {number} n           - Número exato de bytes a extrair (múltiplo de 8MB para Gemini)
 * @returns {number[]}         - Array JS com exatamente N bytes para uso como payload
 */
function extrairPrimeirosNBytesDoBuffer(bufferLista, n) {
  // Converte todos os arrays acumulados para JS e concatena em um único array plano
  // Tamanho máximo: ~9MB (3 chunks × ~3MB) — seguro para V8
  const listaJs = bufferLista.map(byteArrayParaJsArray);
  const completo = [].concat.apply([], listaJs); // Flatten: rápido com .apply para poucos arrays

  // Divide em: pacote de N bytes + restante
  const extraido = completo.slice(0, n);
  const restante = completo.slice(n);

  // Atualiza o buffer in-place
  bufferLista.length = 0;
  if (restante.length > 0) bufferLista.push(restante);

  return extraido;
}

/**
 * Combina todos os Byte[] do buffer em um único array JS.
 * Usado apenas para o chunk FINAL (tamanho livre, sem restrição de granularidade).
 * 
 * @param {Array} bufferLista  - Lista de Byte[] Java ou arrays JS pendentes
 * @returns {number[]}         - Array JS flat com todos os bytes
 */
function combinarBufferCompleto(bufferLista) {
  const listaJs = bufferLista.map(byteArrayParaJsArray);
  return [].concat.apply([], listaJs);
}

function etapa2_TranscreverAudio(dados) {
  const logs = [];
  const log = (tipo, msg) => {
    logs.push({ tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log(`[${tipo}] ${msg}`);
  };

  try {
    log('INFO', '🎙️ [ETAPA 2/4] Transcrevendo áudio...');
    log('INFO', `  📌 FileURI: ${dados.fileUri}`);
    const tempoInicio = Date.now();
    const chaveApi = obterChaveGeminiProjeto();

    // ── Carrega vocabulário UMA VEZ para Camadas 1, 2 e 3 ──
    log('INFO', '📖 Carregando vocabulário de termos...');
    const vocabulario = carregarVocabularioCompleto();
    if (vocabulario.totalTermos > 0) {
      log('SUCESSO', `✅ Vocabulário: ${vocabulario.totalTermos} termos ativos`);
    } else {
      log('INFO', '⚠️ Vocabulário vazio — transcrição sem glossário personalizado');
    }

    // Passa vocabulário para injeção no prompt (Camada 1) e validação posterior (Camadas 2 e 3)
    const resultado = executarTranscricaoViaFileUri(
      dados.fileUri,
      dados.tipoMime || 'audio/webm',
      chaveApi,
      vocabulario
    );

    const tempoSeg = ((Date.now() - tempoInicio) / 1000).toFixed(1);

    if (!resultado.sucesso) {
      log('ERRO', `❌ Transcrição falhou (${tempoSeg}s): ${resultado.mensagem}`);
      return { sucesso: false, logs, mensagem: resultado.mensagem };
    }

    log('SUCESSO', `✅ ETAPA 2 CONCLUÍDA em ${tempoSeg}s!`);
    log('INFO', `  📝 Tamanho: ${resultado.transcricao.length} chars`);
    log('INFO', `  📝 Preview: "${resultado.transcricao.substring(0, 150)}..."`);

    return { sucesso: true, logs, transcricao: resultado.transcricao };

  } catch (erro) {
    log('ERRO', `❌ ${erro.message}`);
    return { sucesso: false, logs, mensagem: erro.message };
  }
}

// ✅ MANTIDA (sem alteração na etapa4, mas agora é opcional — client pode usar as funções separadas)
function etapa4_GerarRelatorioESalvar(dados) {
  const logs = [];
  const log = (tipo, msg) => {
    logs.push({ tipo, mensagem: msg, timestamp: new Date().toLocaleTimeString('pt-BR') });
    Logger.log(`[${tipo}] ${msg}`);
  };

  try {
    log('INFO', '🔍 [ETAPA 4] Relatório + salvando...');
    const chaveApi = obterChaveGeminiProjeto();
    const tempoInicio = Date.now();

    let linkAudioFinal = '';
    if (dados.idUpload) {
      try {
        const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
        const it = pasta.getFilesByName(`chunk_${dados.idUpload}_0`);
        if (it.hasNext()) {
          const arq = it.next();
          arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          linkAudioFinal = arq.getUrl();
          log('SUCESSO', '✅ Link do áudio obtido');
        }
      } catch (e) { log('ALERTA', '⚠️ Link áudio não obtido'); }
    }

    log('INFO', '📝 Salvando transcrição no Drive...');
    const rTrans = salvarTranscricaoNoDrive(dados.transcricao, dados.titulo);
    if (rTrans.sucesso) log('SUCESSO', `✅ Transcrição salva: ${rTrans.nomeArquivo}`);

    log('INFO', '🔍 Gerando relatório...');
    const contexto = obterContextoProjetosParaGemini();
    const rRelatorio = executarEtapaIdentificacaoAlteracoes(dados.transcricao, contexto, chaveApi, dados.titulo || '');

    let linkRelatorio = '', nomeArquivoRelatorio = '', relatorioTexto = '';
    let totalProj = 0, totalEtp = 0, novosProj = 0, novasEtp = 0;

    if (rRelatorio.sucesso) {
      linkRelatorio = rRelatorio.linkRelatorio || '';
      nomeArquivoRelatorio = rRelatorio.nomeArquivoRelatorio || '';
      relatorioTexto = rRelatorio.relatorio || '';
      totalProj = Array.isArray(rRelatorio.projetosIdentificados) ? rRelatorio.projetosIdentificados.length : 0;
      totalEtp = Array.isArray(rRelatorio.etapasIdentificadas) ? rRelatorio.etapasIdentificadas.length : 0;
      novosProj = rRelatorio.novosProjetosSugeridos || 0;
      novasEtp = rRelatorio.novasEtapasSugeridas || 0;
      log('SUCESSO', '✅ Relatório gerado!');
    } else {
      log('ALERTA', '⚠️ Relatório: ' + rRelatorio.mensagem);
    }

    log('INFO', '📊 Salvando reunião...');
    const reuniaoId = salvarReuniaoNaPlanilha({
      titulo: dados.titulo || 'Reunião ' + new Date().toLocaleDateString('pt-BR'),
      dataInicio: dados.dataInicio || new Date(), dataFim: new Date(),
      duracao: dados.duracaoMinutos || 0, participantes: dados.participantes || '',
      transcricao: dados.transcricao || '', ata: dados.ata || '',
      sugestoesIA: dados.sugestoes || '', linkAudio: linkAudioFinal,
      projetosImpactados: '', etapasImpactadas: ''
    });
    log('SUCESSO', `✅ Reunião salva: ${reuniaoId}`);

    if (dados.fileName) {
      try { limparArquivoGemini(dados.fileName, chaveApi); } catch (e) { }
    }

    log('SUCESSO', `✅ ETAPA 4 CONCLUÍDA em ${((Date.now() - tempoInicio) / 1000).toFixed(1)}s!`);
    log('SUCESSO', '🎉 PROCESSAMENTO COMPLETO!');

    return {
      sucesso: true, logs, reuniaoId,
      relatorioIdentificacoes: relatorioTexto, linkRelatorioIdentificacoes: linkRelatorio,
      nomeArquivoRelatorio, totalProjetosIdentificados: totalProj,
      totalEtapasIdentificadas: totalEtp, novosProjetosSugeridos: novosProj,
      novasEtapasSugeridas: novasEtp, linkAudio: linkAudioFinal
    };
  } catch (erro) {
    log('ERRO', `❌ ${erro.message}`);
    return { sucesso: false, logs, mensagem: erro.message };
  }
}

/**
 * @deprecated Usar etapa1_UploadChunksParaGemini que processa sem concatenação em memória.
 * Esta função falha para arquivos > ~50MB (estouro de memória V8).
 */

function consolidarChunksEmAudioNoDrive(idUpload, tipoMime, deletarChunksAposConcluir) {
  try {
    const pasta = DriveApp.getFolderById(ID_PASTA_DRIVE_REUNIOES);
    const arquivosChunks = [];
    let idx = 0;
    while (true) {
      const it = pasta.getFilesByName(`chunk_${idUpload}_${idx}`);
      if (!it.hasNext()) break;
      arquivosChunks.push(it.next());
      idx++;
    }
    if (arquivosChunks.length === 0) return { sucesso: false, mensagem: 'Nenhum chunk encontrado' };

    const partesDecodificadas = [];
    let tamanhoTotal = 0;
    for (let i = 0; i < arquivosChunks.length; i++) {
      const bytesDecod = Utilities.base64Decode(arquivosChunks[i].getBlob().getDataAsString());
      partesDecodificadas.push(bytesDecod);
      tamanhoTotal += bytesDecod.length;
    }

    const todosBytes = [];
    const TAMANHO_LOTE = 50000;
    for (let i = 0; i < partesDecodificadas.length; i++) {
      const parte = partesDecodificadas[i];
      for (let j = 0; j < parte.length; j += TAMANHO_LOTE) {
        Array.prototype.push.apply(todosBytes, Array.prototype.slice.call(parte, j, Math.min(j + TAMANHO_LOTE, parte.length)));
      }
    }

    const extensao = obterExtensaoDoMime(tipoMime);
    const timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMdd_HHmmss');
    const nomeArquivo = `Reuniao_${timestamp}${extensao}`;
    const audioBlob = Utilities.newBlob(todosBytes, tipoMime || 'audio/webm', nomeArquivo);
    const arquivo = pasta.createFile(audioBlob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    if (deletarChunksAposConcluir) {
      arquivosChunks.forEach(function(arq) { try { arq.setTrashed(true); } catch (e) { } });
    }

    return { sucesso: true, arquivoId: arquivo.getId(), nomeArquivo, linkArquivo: arquivo.getUrl() };
  } catch (erro) {
    Logger.log('ERRO consolidarChunksEmAudioNoDrive: ' + erro.toString());
    return { sucesso: false, mensagem: erro.message };
  }
}

function calcularTamanhoBinarioExato(infoChunks) {
  const totalChunks = infoChunks.length;
  const ultimoArquivo = DriveApp.getFileById(infoChunks[totalChunks - 1].id);
  const textoUltimoChunk = ultimoArquivo.getBlob().getDataAsString();
  let paddingUltimoChunk = 0;
  if (textoUltimoChunk.endsWith('==')) paddingUltimoChunk = 2;
  else if (textoUltimoChunk.endsWith('=')) paddingUltimoChunk = 1;
  let tamBin = 0;
  for (let i = 0; i < totalChunks; i++) {
    const tamBase64 = infoChunks[i].tamanhoBytes;
    tamBin += Math.floor(tamBase64 * 3 / 4) - (i === totalChunks - 1 ? paddingUltimoChunk : 0);
  }
  return tamBin;
}