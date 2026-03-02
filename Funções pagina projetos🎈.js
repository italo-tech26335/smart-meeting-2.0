function obterDadosDiagrama(modoAdminParam) {
  try {
    Logger.log('Iniciando obterDadosDiagrama');

    // Obtém permissões do usuário atual
    const permissoesUsuario = obterPermissoesUsuarioAtual();
    Logger.log('Permissões do usuário: ' + JSON.stringify(permissoesUsuario));

    // Modo admin: só ativo se o usuário realmente é admin no sistema
    const modoAdmin = modoAdminParam === true && ehAdministrador(permissoesUsuario);

    // Carrega todos os dados brutos
    const todosProjetos = obterTodosProjetos();
    const responsaveis = obterTodosResponsaveis();
    const todasEtapas = obterTodasEtapas();
    const dependencias = obterTodasDependencias();
    const setoresResponsaveis = obterSetoresComResponsaveis();

    // Filtra projetos baseado nas permissões
    let projetosFiltrados;
    if (modoAdmin || permissoesUsuario.nivelAcesso === NIVEIS_ACESSO.ADMIN) {
      // Admin (ou modo admin URL) vê tudo sem restrições
      projetosFiltrados = todosProjetos;
    } else if (permissoesUsuario.nivelAcesso === 'visitante' || permissoesUsuario.nivelAcesso === 'inativo') {
      // Visitante/Inativo não vê nada
      projetosFiltrados = [];
    } else {
      // Filtra baseado nas permissões
      projetosFiltrados = todosProjetos.filter(projeto => {
        return podeVerProjeto(projeto.id, permissoesUsuario);
      });
    }
    
    // Filtro por responsável: exibe só projetos onde o usuário tem etapas atribuídas
    if (permissoesUsuario.filtrarPorResponsavel && permissoesUsuario.nivelAcesso !== NIVEIS_ACESSO.ADMIN) {
      const emailUsuarioLower = (permissoesUsuario.email || '').toLowerCase();
      const respUsuario = responsaveis.find(r =>
        r.email && r.email.toLowerCase() === emailUsuarioLower
      );
      if (respUsuario) {
        const projetosComEtapas = new Set(
          todasEtapas
            .filter(e => (e.responsaveisIds || []).includes(respUsuario.id) || e.responsavelId === respUsuario.id)
            .map(e => e.projetoId)
        );
        projetosFiltrados = projetosFiltrados.filter(p => projetosComEtapas.has(p.id));
      }
    }

    // Obtém IDs dos projetos visíveis
    const projetosVisiveis = new Set(projetosFiltrados.map(p => p.id));
    
    // Filtra etapas apenas dos projetos visíveis
    const etapasFiltradas = todasEtapas.filter(e => projetosVisiveis.has(e.projetoId));
    
    // Filtra dependências apenas das etapas visíveis
    const etapasVisiveis = new Set(etapasFiltradas.map(e => e.id));
    const dependenciasFiltradas = dependencias.filter(d => 
      etapasVisiveis.has(d.etapaOrigemId) && etapasVisiveis.has(d.etapaDestinoId)
    );
    
    // Filtra setores (apenas os que têm projetos visíveis)
    const setoresComProjetos = new Set(projetosFiltrados.map(p => p.setor).filter(s => s));
    const setoresFiltrados = setoresResponsaveis.filter(s => 
      permissoesUsuario.nivelAcesso === NIVEIS_ACESSO.ADMIN || setoresComProjetos.has(s.nomeSetor)
    );

    Logger.log('Dados filtrados. Projetos visíveis: ' + projetosFiltrados.length);

    return {
      sucesso: true,
      projetos: projetosFiltrados,
      responsaveis: responsaveis,
      vinculos: [],
      etapas: etapasFiltradas,
      dependencias: dependenciasFiltradas,
      setoresResponsaveis: setoresFiltrados,
      podeEditar: permissoesUsuario.ativo && permissoesUsuario.nivelAcesso !== 'visitante',
      permissoes: {
        nivelAcesso: permissoesUsuario.nivelAcesso,
        podeCriarProjeto: podeCriarProjeto(permissoesUsuario),
        podeCriarEtapa: permissoesUsuario.podeCriarEtapa,
        ehAdmin: ehAdministrador(permissoesUsuario),
        setoresPermitidos: permissoesUsuario.setoresPermitidos,
        projetosPermitidos: permissoesUsuario.projetosPermitidos,
        filtrarPorResponsavel: permissoesUsuario.filtrarPorResponsavel || false,
        email: permissoesUsuario.email || ''
      }
    };
  } catch (erro) {
    Logger.log('ERRO em obterDadosDiagrama: ' + erro.toString());
    return { 
      sucesso: false, 
      mensagem: erro.message, 
      projetos: [], 
      responsaveis: [], 
      vinculos: [], 
      etapas: [], 
      dependencias: [], 
      setoresResponsaveis: [], 
      podeEditar: false,
      permissoes: criarPermissaoPadrao('erro')
    };
  }
}

function obterProjetoPorId(projetoId) {
  try {
    const projetos = obterTodosProjetos();
    const projeto = projetos.find(p => p.id === projetoId);
    if (!projeto) return { sucesso: false, mensagem: 'Projeto não encontrado' };
    
    const etapas = obterTodasEtapas().filter(e => e.projetoId === projetoId);
    const responsaveis = obterTodosResponsaveis();
    
    return {
      sucesso: true,
      projeto: projeto,
      etapas: etapas,
      responsaveis: responsaveis
    };
  } catch (e) {
    return { sucesso: false, mensagem: e.message };
  }
}

function obterTodosProjetos() {
  const aba = obterAba(NOME_ABA_PROJETOS);
  if (!aba || aba.getLastRow() <= 1) return [];
  const dados = aba.getDataRange().getValues();
  const projetos = [];
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_PROJETOS.ID]) {
      // Processa múltiplos IDs separando por vírgula
      const rawIds = dados[i][COLUNAS_PROJETOS.RESPONSAVEIS_IDS] || '';
      const responsaveisIds = rawIds.toString().split(',').map(id => id.trim()).filter(id => id !== '');

      projetos.push({
        id: dados[i][COLUNAS_PROJETOS.ID] || '',
        nome: dados[i][COLUNAS_PROJETOS.NOME] || '',
        descricao: dados[i][COLUNAS_PROJETOS.DESCRICAO] || '',
        finalidade: dados[i][COLUNAS_PROJETOS.FINALIDADE] || '',
        paraQuem: dados[i][COLUNAS_PROJETOS.PARA_QUEM] || '',
        status: dados[i][COLUNAS_PROJETOS.STATUS] || 'Ativo',
        prioridade: dados[i][COLUNAS_PROJETOS.PRIORIDADE] || 'Média',
        link: dados[i][COLUNAS_PROJETOS.LINK] || '',
        posX: dados[i][COLUNAS_PROJETOS.POS_X] || null,
        posY: dados[i][COLUNAS_PROJETOS.POS_Y] || null,
        bloqueado: dados[i][COLUNAS_PROJETOS.BLOQUEADO] === true || dados[i][COLUNAS_PROJETOS.BLOQUEADO] === 'true',
        setor: dados[i][COLUNAS_PROJETOS.SETOR] || '', 
        pilar: dados[i][COLUNAS_PROJETOS.PILAR] || '',
        responsaveisIds: responsaveisIds // Retorna Array
      });
    }
  }
  return projetos;
}

function adicionarProjeto(dados) {
  try {
    Logger.log('Adicionando projeto: ' + JSON.stringify(dados));
    
    // Verifica permissão para criar projetos
    if (!podeCriarProjeto()) {
      throw new Error('PERMISSAO_NEGADA: Você não tem permissão para criar projetos.');
    }
    
    const aba = obterAba(NOME_ABA_PROJETOS);
    const id = gerarId();
    
    const responsaveisString = Array.isArray(dados.responsaveisIds) 
      ? dados.responsaveisIds.join(',') 
      : (dados.responsaveisIds || '');

    const novaLinha = [
      id,
      dados.nome || '',
      dados.descricao || '',
      dados.finalidade || '',
      dados.paraQuem || '',
      dados.status || 'Ativo',
      dados.prioridade || 'Média',
      dados.link || '',
      dados.posX || 100,
      dados.posY || 80,
      false,
      dados.setor || '',
      dados.pilar || '',
      responsaveisString
    ];
    aba.appendRow(novaLinha);
    const linhaNova = aba.getLastRow();
    gravarTimestamp(aba, linhaNova, COLUNAS_PROJETOS.DATA_CRIACAO, COLUNAS_PROJETOS.DATA_ULTIMA_MODIFICACAO, true);
    Logger.log('Projeto criado com ID: ' + id);
    return { sucesso: true, id: id, mensagem: 'Projeto criado!' };
  } catch (e) {
    Logger.log('ERRO adicionarProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function atualizarProjeto(id, campos) {
  try {
    Logger.log('Atualizando projeto ' + id);
    
    // Verifica permissão específica para este projeto
    if (!podeEditarProjeto(id)) {
      throw new Error('PERMISSAO_NEGADA: Você não tem permissão para editar este projeto.');
    }
    
    const aba = obterAba(NOME_ABA_PROJETOS);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID] === id) {
        if ((dados[i][COLUNAS_PROJETOS.BLOQUEADO] === true || dados[i][COLUNAS_PROJETOS.BLOQUEADO] === 'true') && campos.bloqueado === undefined) {
          throw new Error('Projeto bloqueado.');
        }
        const linha = i + 1;
        if (campos.nome !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.NOME + 1).setValue(campos.nome);
        if (campos.descricao !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.DESCRICAO + 1).setValue(campos.descricao);
        if (campos.finalidade !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.FINALIDADE + 1).setValue(campos.finalidade);
        if (campos.paraQuem !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.PARA_QUEM + 1).setValue(campos.paraQuem);
        if (campos.status !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.STATUS + 1).setValue(campos.status);
        if (campos.prioridade !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.PRIORIDADE + 1).setValue(campos.prioridade);
        if (campos.link !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.LINK + 1).setValue(campos.link);
        if (campos.posX !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.POS_X + 1).setValue(campos.posX);
        if (campos.posY !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.POS_Y + 1).setValue(campos.posY);
        if (campos.bloqueado !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.BLOQUEADO + 1).setValue(campos.bloqueado);
        if (campos.setor !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.SETOR + 1).setValue(campos.setor);
        if (campos.pilar !== undefined) aba.getRange(linha, COLUNAS_PROJETOS.PILAR + 1).setValue(campos.pilar);
        
        if (campos.responsaveisIds !== undefined) {
          const respString = Array.isArray(campos.responsaveisIds) ? campos.responsaveisIds.join(',') : campos.responsaveisIds;
          aba.getRange(linha, COLUNAS_PROJETOS.RESPONSAVEIS_IDS + 1).setValue(respString);
        }

        gravarTimestamp(aba, linha, -1, COLUNAS_PROJETOS.DATA_ULTIMA_MODIFICACAO, false);
        return { sucesso: true, mensagem: 'Projeto atualizado!' };
      }
    }
    return { sucesso: false, mensagem: 'Projeto não encontrado' };
  } catch (e) {
    Logger.log('ERRO atualizarProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function atualizarPosicaoProjeto(id, posX, posY) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_PROJETOS);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID] === id) {
        if (dados[i][COLUNAS_PROJETOS.BLOQUEADO] === true || dados[i][COLUNAS_PROJETOS.BLOQUEADO] === 'true') {
          return { sucesso: false, mensagem: 'Projeto bloqueado' };
        }
        aba.getRange(i + 1, COLUNAS_PROJETOS.POS_X + 1).setValue(posX);
        aba.getRange(i + 1, COLUNAS_PROJETOS.POS_Y + 1).setValue(posY);
        return { sucesso: true };
      }
    }
    return { sucesso: false, mensagem: 'Projeto não encontrado' };
  } catch (e) {
    Logger.log('ERRO atualizarPosicaoProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function alternarBloqueioProjeto(id) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_PROJETOS);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PROJETOS.ID] === id) {
        const estadoAtual = dados[i][COLUNAS_PROJETOS.BLOQUEADO] === true || dados[i][COLUNAS_PROJETOS.BLOQUEADO] === 'true';
        const novoEstado = !estadoAtual;
        aba.getRange(i + 1, COLUNAS_PROJETOS.BLOQUEADO + 1).setValue(novoEstado);
        return { sucesso: true, bloqueado: novoEstado, mensagem: novoEstado ? 'Projeto bloqueado' : 'Projeto desbloqueado' };
      }
    }
    return { sucesso: false, mensagem: 'Projeto não encontrado' };
  } catch (e) {
    Logger.log('ERRO alternarBloqueioProjeto: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function obterTodosResponsaveis() {
  const aba = obterAba(NOME_ABA_RESPONSAVEIS);
  if (!aba || aba.getLastRow() <= 1) return [];
  const dados = aba.getDataRange().getValues();
  const responsaveis = [];
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_RESPONSAVEIS.ID]) {
      responsaveis.push({
        id: dados[i][COLUNAS_RESPONSAVEIS.ID] || '',
        nome: dados[i][COLUNAS_RESPONSAVEIS.NOME] || '',
        email: dados[i][COLUNAS_RESPONSAVEIS.EMAIL] || '',
        cargo: dados[i][COLUNAS_RESPONSAVEIS.CARGO] || '',
        posX: dados[i][COLUNAS_RESPONSAVEIS.POS_X] || null,
        posY: dados[i][COLUNAS_RESPONSAVEIS.POS_Y] || null,
        ativo: true
      });
    }
  }
  return responsaveis;
}

function adicionarResponsavel(dados) {
  try {
    Logger.log('Adicionando responsável');
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_RESPONSAVEIS);
    const id = gerarId();
    aba.appendRow([id, dados.nome || '', dados.email || '', dados.cargo || '']);
    return { sucesso: true, id: id, mensagem: 'Responsável criado!' };
  } catch (e) {
    Logger.log('ERRO adicionarResponsavel: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function atualizarResponsavel(id, campos) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_RESPONSAVEIS);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_RESPONSAVEIS.ID] === id) {
        const linha = i + 1;
        if (campos.nome !== undefined) aba.getRange(linha, COLUNAS_RESPONSAVEIS.NOME + 1).setValue(campos.nome);
        if (campos.email !== undefined) aba.getRange(linha, COLUNAS_RESPONSAVEIS.EMAIL + 1).setValue(campos.email);
        if (campos.cargo !== undefined) aba.getRange(linha, COLUNAS_RESPONSAVEIS.CARGO + 1).setValue(campos.cargo);
        return { sucesso: true, mensagem: 'Responsável atualizado!' };
      }
    }
    return { sucesso: false, mensagem: 'Responsável não encontrado' };
  } catch (e) {
    Logger.log('ERRO atualizarResponsavel: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function atualizarPosicaoResponsavel(id, posX, posY) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_RESPONSAVEIS);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_RESPONSAVEIS.ID] === id) {
        aba.getRange(i + 1, COLUNAS_RESPONSAVEIS.POS_X + 1).setValue(posX);
        aba.getRange(i + 1, COLUNAS_RESPONSAVEIS.POS_Y + 1).setValue(posY);
        return { sucesso: true };
      }
    }
    return { sucesso: false, mensagem: 'Responsável não encontrado' };
  } catch (e) {
    Logger.log('ERRO atualizarPosicaoResponsavel: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function obterTodasEtapas() {
  const aba = obterAba(NOME_ABA_ETAPAS);
  if (!aba || aba.getLastRow() <= 1) return [];
  const dados = aba.getDataRange().getValues();
  const etapas = [];
  
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_ETAPAS.ID]) {
      const rawIds = dados[i][COLUNAS_ETAPAS.RESPONSAVEIS_IDS] || '';
      const responsaveisIds = rawIds.toString().split(',').map(id => id.trim()).filter(id => id !== '');
      
      // NOVO: Parsear pendências como JSON ou migrar texto antigo
      const pendenciasRaw = dados[i][COLUNAS_ETAPAS.PENDENCIAS] || '';
      let pendencias = [];
      
      try {
        if (pendenciasRaw.startsWith('[')) {
          pendencias = JSON.parse(pendenciasRaw);
        } else if (pendenciasRaw.trim() !== '') {
          // Migração: texto antigo vira uma pendência única
          pendencias = [{
            id: 'pnd_migrado_' + Date.now(),
            texto: pendenciasRaw,
            urgencia: 'media',
            concluido: false,
            dataCriacao: new Date().toISOString()
          }];
        }
      } catch (e) {
        Logger.log('Erro ao parsear pendências: ' + e.message);
        pendencias = [];
      }

      etapas.push({
        id: dados[i][COLUNAS_ETAPAS.ID] || '',
        projetoId: dados[i][COLUNAS_ETAPAS.PROJETO_ID] || '',
        responsaveisIds: responsaveisIds,
        nome: dados[i][COLUNAS_ETAPAS.NOME] || '',
        descricao: dados[i][COLUNAS_ETAPAS.DESCRICAO] || '',
        oQueFazer: dados[i][COLUNAS_ETAPAS.O_QUE_FAZER] || '',
        pendencias: pendencias, // Agora é um Array de objetos
        status: dados[i][COLUNAS_ETAPAS.STATUS] || STATUS_ETAPAS.A_FAZER,
        offsetX: dados[i][COLUNAS_ETAPAS.OFFSET_X] || 0,
        offsetY: dados[i][COLUNAS_ETAPAS.OFFSET_Y] || 0,
        bloqueado: dados[i][COLUNAS_ETAPAS.BLOQUEADO] === true || dados[i][COLUNAS_ETAPAS.BLOQUEADO] === 'true',
        etapaPaiId: dados[i][COLUNAS_ETAPAS.ETAPA_PAI_ID] || null
      });
    }
  }
  return etapas;
}

function adicionarEtapa(dados) {
  try {
    Logger.log('Adicionando Etapa: ' + JSON.stringify(dados));
    
    if (!dados.projetoId || dados.projetoId.trim() === '') {
      throw new Error('O ID do Projeto é obrigatório para criar uma etapa.');
    }
    
    // Verifica permissão para criar etapa neste projeto
    if (!podeCriarEtapa(dados.projetoId)) {
      throw new Error('PERMISSAO_NEGADA: Você não tem permissão para criar etapas neste projeto.');
    }

    const aba = obterAba(NOME_ABA_ETAPAS);
    const id = gerarId();
    
    const responsaveisString = Array.isArray(dados.responsaveisIds) 
      ? dados.responsaveisIds.join(',') 
      : (dados.responsaveisIds || '');

    let pendenciasString = '';
    if (dados.pendencias) {
      if (Array.isArray(dados.pendencias)) {
        pendenciasString = JSON.stringify(dados.pendencias);
      } else if (typeof dados.pendencias === 'string') {
        pendenciasString = dados.pendencias;
      }
    }

    const novaLinha = [
      id,
      dados.projetoId,
      responsaveisString,
      dados.nome || '',
      dados.descricao || '',
      dados.oQueFazer || '',
      pendenciasString,
      dados.status || STATUS_ETAPAS.A_FAZER,
      dados.offsetX || 0,
      dados.offsetY || 30,
      false,
      dados.etapaPaiId || ''
    ];

    aba.appendRow(novaLinha);
    Logger.log('Etapa criada com ID: ' + id);
    return { sucesso: true, id: id, mensagem: 'Etapa criada!' };
  } catch (e) {
    Logger.log('ERRO adicionarEtapa: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function atualizarEtapa(id, campos) {
  try {
    Logger.log('Atualizando etapa ' + id + ' com campos: ' + JSON.stringify(campos));
    
    // Verifica permissão específica para esta etapa
    if (!podeEditarEtapa(id)) {
      throw new Error('PERMISSAO_NEGADA: Você não tem permissão para editar esta etapa.');
    }
    
    const aba = obterAba(NOME_ABA_ETAPAS);
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_ETAPAS.ID] === id) {
        
        const estaBloqueada = dados[i][COLUNAS_ETAPAS.BLOQUEADO] === true || 
                              dados[i][COLUNAS_ETAPAS.BLOQUEADO] === 'true';
        if (estaBloqueada && campos.bloqueado === undefined) {
          throw new Error('Etapa bloqueada.');
        }
        
        const linha = i + 1;
        
        if (campos.nome !== undefined) {
          aba.getRange(linha, COLUNAS_ETAPAS.NOME + 1).setValue(campos.nome);
        }
        
        if (campos.descricao !== undefined) {
          aba.getRange(linha, COLUNAS_ETAPAS.DESCRICAO + 1).setValue(campos.descricao);
        }
        
        if (campos.oQueFazer !== undefined) {
          aba.getRange(linha, COLUNAS_ETAPAS.O_QUE_FAZER + 1).setValue(campos.oQueFazer);
        }
        
        if (campos.pendencias !== undefined) {
          let pendenciasParaSalvar = '';
          
          if (Array.isArray(campos.pendencias)) {
            pendenciasParaSalvar = JSON.stringify(campos.pendencias);
          } else if (typeof campos.pendencias === 'string') {
            pendenciasParaSalvar = campos.pendencias;
          }
          
          aba.getRange(linha, COLUNAS_ETAPAS.PENDENCIAS + 1).setValue(pendenciasParaSalvar);
        }
        
        if (campos.status !== undefined) {
          aba.getRange(linha, COLUNAS_ETAPAS.STATUS + 1).setValue(campos.status);
        }
        
        if (campos.offsetX !== undefined) {
          aba.getRange(linha, COLUNAS_ETAPAS.OFFSET_X + 1).setValue(campos.offsetX);
        }
        
        if (campos.offsetY !== undefined) {
          aba.getRange(linha, COLUNAS_ETAPAS.OFFSET_Y + 1).setValue(campos.offsetY);
        }
        
        if (campos.bloqueado !== undefined) {
          aba.getRange(linha, COLUNAS_ETAPAS.BLOQUEADO + 1).setValue(campos.bloqueado);
        }
        
        if (campos.etapaPaiId !== undefined) {
          aba.getRange(linha, COLUNAS_ETAPAS.ETAPA_PAI_ID + 1).setValue(campos.etapaPaiId);
        }
        
        if (campos.responsaveisIds !== undefined) {
          const respString = Array.isArray(campos.responsaveisIds) 
            ? campos.responsaveisIds.join(',') 
            : campos.responsaveisIds;
          aba.getRange(linha, COLUNAS_ETAPAS.RESPONSAVEIS_IDS + 1).setValue(respString);
        }
        
        return { sucesso: true, mensagem: 'Etapa atualizada!' };
      }
    }
    
    return { sucesso: false, mensagem: 'Etapa não encontrada' };
  } catch (e) {
    Logger.log('ERRO atualizarEtapa: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function excluirEtapa(id) {
  try {
    Logger.log('Excluindo etapa ' + id);
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_ETAPAS);
    const dados = aba.getDataRange().getValues();
    function encontrarDescendentes(etapaId, todasEtapas) {
      const descendentes = [etapaId];
      const filhas = todasEtapas.filter(e => e.etapaPaiId === etapaId);
      for (const filha of filhas) descendentes.push(...encontrarDescendentes(filha.id, todasEtapas));
      return descendentes;
    }
    const todasEtapas = [];
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_ETAPAS.ID]) todasEtapas.push({ id: dados[i][COLUNAS_ETAPAS.ID], etapaPaiId: dados[i][COLUNAS_ETAPAS.ETAPA_PAI_ID] || null });
    }
    const etapasParaRemover = encontrarDescendentes(id, todasEtapas);
    const abaDeps = obterAba(NOME_ABA_DEPENDENCIAS);
    const dadosDeps = abaDeps.getDataRange().getValues();
    for (let i = dadosDeps.length - 1; i >= 1; i--) {
      if (etapasParaRemover.includes(dadosDeps[i][COLUNAS_DEPENDENCIAS.ETAPA_ORIGEM_ID]) || etapasParaRemover.includes(dadosDeps[i][COLUNAS_DEPENDENCIAS.ETAPA_DESTINO_ID])) abaDeps.deleteRow(i + 1);
    }
    for (let i = dados.length - 1; i >= 1; i--) {
      if (etapasParaRemover.includes(dados[i][COLUNAS_ETAPAS.ID])) aba.deleteRow(i + 1);
    }
    Logger.log('Etapa excluída e descendentes');
    return { sucesso: true, mensagem: etapasParaRemover.length > 1 ? `Etapa e ${etapasParaRemover.length - 1} sub-etapa(s) excluídas!` : 'Etapa excluída!' };
  } catch (e) {
    Logger.log('ERRO excluirEtapa: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function alternarBloqueioEtapa(id) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_ETAPAS);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_ETAPAS.ID] === id) {
        const estadoAtual = dados[i][COLUNAS_ETAPAS.BLOQUEADO] === true || dados[i][COLUNAS_ETAPAS.BLOQUEADO] === 'true';
        const novoEstado = !estadoAtual;
        aba.getRange(i + 1, COLUNAS_ETAPAS.BLOQUEADO + 1).setValue(novoEstado);
        return { sucesso: true, bloqueado: novoEstado };
      }
    }
    return { sucesso: false, mensagem: 'Etapa não encontrada' };
  } catch (e) {
    Logger.log('ERRO alternarBloqueioEtapa: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function atualizarPosicaoEtapa(etapaId, offsetX, offsetY) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_ETAPAS);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_ETAPAS.ID] === etapaId) {
        if (dados[i][COLUNAS_ETAPAS.BLOQUEADO] === true || dados[i][COLUNAS_ETAPAS.BLOQUEADO] === 'true') return { sucesso: false, mensagem: 'Etapa bloqueada' };
        aba.getRange(i + 1, COLUNAS_ETAPAS.OFFSET_X + 1).setValue(offsetX);
        aba.getRange(i + 1, COLUNAS_ETAPAS.OFFSET_Y + 1).setValue(offsetY);
        return { sucesso: true };
      }
    }
    return { sucesso: false, mensagem: 'Etapa não encontrada' };
  } catch (e) {
    Logger.log('ERRO atualizarPosicaoEtapa: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

// Obs: Esta função era usada para setar 1 ID. Vamos mantê-la mas atualizando a string.
// Para UI que suporta múltiplos, o ideal é usar atualizarEtapa, mas mantemos compatibilidade
function vincularResponsavelNaEtapa(dados) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_ETAPAS);
    const dadosEtapas = aba.getDataRange().getValues();
    for (let i = 1; i < dadosEtapas.length; i++) {
      if (dadosEtapas[i][COLUNAS_ETAPAS.ID] === dados.etapaId) {
        // Se já existem, adiciona. Se não, cria. (Comportamento de append)
        // OU substitui? O comportamento padrão anterior era substituir. Manteremos substituir para consistência da função.
        aba.getRange(i + 1, COLUNAS_ETAPAS.RESPONSAVEIS_IDS + 1).setValue(dados.responsavelId);
        return { sucesso: true, mensagem: 'Responsável vinculado!' };
      }
    }
    return { sucesso: false, mensagem: 'Etapa não encontrada' };
  } catch (e) {
    Logger.log('ERRO vincularResponsavelNaEtapa: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function removerResponsavelDaEtapa(dados) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_ETAPAS);
    const dadosEtapas = aba.getDataRange().getValues();
    for (let i = 1; i < dadosEtapas.length; i++) {
      if (dadosEtapas[i][COLUNAS_ETAPAS.ID] === dados.etapaId) {
        aba.getRange(i + 1, COLUNAS_ETAPAS.RESPONSAVEIS_IDS + 1).setValue('');
        return { sucesso: true, mensagem: 'Vínculo removido!' };
      }
    }
    return { sucesso: false, mensagem: 'Etapa não encontrada' };
  } catch (e) {
    Logger.log('ERRO removerResponsavelDaEtapa: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

// Adaptado para suportar múltiplos IDs (append)
function vincularResponsavel(responsavelId, etapaId) {
  try {
    Logger.log('Iniciando vinculação: ' + responsavelId + ' -> ' + etapaId);
    exigirPermissaoEdicao();
    
    const abaResp = obterAba(NOME_ABA_RESPONSAVEIS);
    const responsaveis = abaResp.getDataRange().getValues();
    let responsavelExiste = false;
    for (let i = 1; i < responsaveis.length; i++) {
      if (responsaveis[i][COLUNAS_RESPONSAVEIS.ID] === responsavelId) {
        responsavelExiste = true;
        break;
      }
    }
    if (!responsavelExiste) return { sucesso: false, mensagem: 'Responsável não encontrado' };

    const abaEtapas = obterAba(NOME_ABA_ETAPAS);
    const etapas = abaEtapas.getDataRange().getValues();
    let linhaEtapa = -1;
    for (let i = 1; i < etapas.length; i++) {
      if (etapas[i][COLUNAS_ETAPAS.ID] === etapaId) {
        linhaEtapa = i + 1;
        break;
      }
    }
    if (linhaEtapa === -1) return { sucesso: false, mensagem: 'Etapa não encontrada' };

    const rawIds = etapas[linhaEtapa - 1][COLUNAS_ETAPAS.RESPONSAVEIS_IDS] || '';
    const currentIds = rawIds.toString().split(',').map(id => id.trim()).filter(id => id !== '');

    if (currentIds.includes(responsavelId)) {
      return { sucesso: false, mensagem: 'Responsável já vinculado a esta etapa' };
    }

    currentIds.push(responsavelId);
    const novaString = currentIds.join(',');

    abaEtapas.getRange(linhaEtapa, COLUNAS_ETAPAS.RESPONSAVEIS_IDS + 1).setValue(novaString);

    Logger.log('Vínculo realizado com sucesso.');
    return {
       sucesso: true,
       id: gerarId(),
       mensagem: 'Responsável vinculado à etapa com sucesso!'
     };
  } catch (e) {
    Logger.log('ERRO vincularResponsavel: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function obterTodasDependencias() {
  const aba = obterAba(NOME_ABA_DEPENDENCIAS);
  if (!aba || aba.getLastRow() <= 1) return [];
  const dados = aba.getDataRange().getValues();
  const dependencias = [];
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_DEPENDENCIAS.ID]) {
      dependencias.push({
        id: dados[i][COLUNAS_DEPENDENCIAS.ID] || '',
        etapaOrigemId: dados[i][COLUNAS_DEPENDENCIAS.ETAPA_ORIGEM_ID] || '',
        origemAnchor: dados[i][COLUNAS_DEPENDENCIAS.ORIGEM_ANCHOR] || 'baixo',
        etapaDestinoId: dados[i][COLUNAS_DEPENDENCIAS.ETAPA_DESTINO_ID] || '',
        destinoAnchor: dados[i][COLUNAS_DEPENDENCIAS.DESTINO_ANCHOR] || 'topo'
      });
    }
  }
  return dependencias;
}

function adicionarDependencia(dados) {
  try {
    exigirPermissaoEdicao();
    if (dados.etapaOrigemId === dados.etapaDestinoId) return { sucesso: false, mensagem: 'Uma etapa não pode depender de si mesma' };
    const aba = obterAba(NOME_ABA_DEPENDENCIAS);
    const id = gerarId();
    aba.appendRow([id, dados.etapaOrigemId, dados.origemAnchor || 'baixo', dados.etapaDestinoId, dados.destinoAnchor || 'topo']);
    return { sucesso: true, id: id, mensagem: 'Dependência criada!' };
  } catch (e) {
    Logger.log('ERRO adicionarDependencia: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function removerDependencia(dependenciaId) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_DEPENDENCIAS);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_DEPENDENCIAS.ID] === dependenciaId) {
        aba.deleteRow(i + 1);
        return { sucesso: true, mensagem: 'Dependência removida!' };
      }
    }
    return { sucesso: false, mensagem: 'Dependência não encontrada' };
  } catch (e) {
    Logger.log('ERRO removerDependencia: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function atualizarBloqueioProjeto(id, bloqueado) {
  return alternarBloqueioProjeto(id);
}

function atualizarBloqueioEtapa(id, bloqueado) {
  return alternarBloqueioEtapa(id);
}

// Novas Funções para Gestão de Setores/Responsáveis

function obterSetoresComResponsaveis() {
  const aba = obterAba(NOME_ABA_SETORES);
  if (!aba || aba.getLastRow() <= 1) return [];
  const dados = aba.getDataRange().getValues();
  const setores = [];
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_SETORES.NOME]) {
      setores.push({
        id: dados[i][COLUNAS_SETORES.ID] || '',
        nomeSetor: dados[i][COLUNAS_SETORES.NOME] || '',
        responsavelId: dados[i][COLUNAS_SETORES.RESPONSAVEIS_IDS] || '',
        posX: dados[i][COLUNAS_SETORES.POS_X] || null,
        posY: dados[i][COLUNAS_SETORES.POS_Y] || null
      });
    }
  }
  return setores;
}

/**
 * Adiciona um novo setor na aba Setores
 * @param {Object} dados - {nome, descricao, responsavelId}
 * @returns {Object} - {sucesso, id, mensagem}
 */
function adicionarSetor(dados) {
  try {
    Logger.log('Adicionando setor: ' + JSON.stringify(dados));
    exigirPermissaoEdicao();
    
    if (!dados.nome || dados.nome.trim() === '') {
      return { sucesso: false, mensagem: 'Nome do setor é obrigatório' };
    }
    
    const aba = obterAba(NOME_ABA_SETORES);
    const dadosExistentes = aba.getDataRange().getValues();
    
    // Verifica se já existe setor com mesmo nome
    for (let i = 1; i < dadosExistentes.length; i++) {
      if (dadosExistentes[i][COLUNAS_SETORES.NOME] && 
          dadosExistentes[i][COLUNAS_SETORES.NOME].toString().toLowerCase() === dados.nome.trim().toLowerCase()) {
        return { sucesso: false, mensagem: 'Já existe um setor com este nome' };
      }
    }
    
    const id = gerarId();
    const novaLinha = [
      id,
      dados.nome.trim(),
      dados.descricao || '',
      dados.responsavelId || ''
    ];
    
    aba.appendRow(novaLinha);
    Logger.log('Setor criado com ID: ' + id);
    
    return { sucesso: true, id: id, mensagem: 'Setor criado com sucesso!' };
  } catch (e) {
    Logger.log('ERRO adicionarSetor: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Atualiza um setor existente
 * @param {string} id - ID do setor
 * @param {Object} campos - {nome, descricao, responsavelId}
 * @returns {Object} - {sucesso, mensagem}
 */
function atualizarSetor(id, campos) {
  try {
    Logger.log('Atualizando setor ' + id);
    exigirPermissaoEdicao();
    
    const aba = obterAba(NOME_ABA_SETORES);
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_SETORES.ID] === id) {
        const linha = i + 1;
        
        if (campos.nome !== undefined) {
          // Verifica duplicidade de nome
          for (let j = 1; j < dados.length; j++) {
            if (j !== i && dados[j][COLUNAS_SETORES.NOME] && 
                dados[j][COLUNAS_SETORES.NOME].toString().toLowerCase() === campos.nome.trim().toLowerCase()) {
              return { sucesso: false, mensagem: 'Já existe um setor com este nome' };
            }
          }
          aba.getRange(linha, COLUNAS_SETORES.NOME + 1).setValue(campos.nome.trim());
        }
        
        if (campos.descricao !== undefined) {
          aba.getRange(linha, COLUNAS_SETORES.DESCRICAO + 1).setValue(campos.descricao);
        }
        
        if (campos.responsavelId !== undefined) {
          aba.getRange(linha, COLUNAS_SETORES.RESPONSAVEIS_IDS + 1).setValue(campos.responsavelId);
        }
        
        return { sucesso: true, mensagem: 'Setor atualizado!' };
      }
    }
    
    return { sucesso: false, mensagem: 'Setor não encontrado' };
  } catch (e) {
    Logger.log('ERRO atualizarSetor: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Exclui um setor (apenas se não houver projetos vinculados)
 * @param {string} id - ID do setor
 * @returns {Object} - {sucesso, mensagem}
 */
function excluirSetor(id) {
  try {
    Logger.log('Excluindo setor ' + id);
    exigirPermissaoEdicao();
    
    const aba = obterAba(NOME_ABA_SETORES);
    const dados = aba.getDataRange().getValues();
    
    // Encontra o setor
    let linhaSetor = -1;
    let nomeSetor = '';
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_SETORES.ID] === id) {
        linhaSetor = i + 1;
        nomeSetor = dados[i][COLUNAS_SETORES.NOME];
        break;
      }
    }
    
    if (linhaSetor === -1) {
      return { sucesso: false, mensagem: 'Setor não encontrado' };
    }
    
    // Verifica se há projetos vinculados a este setor
    const abaProjetos = obterAba(NOME_ABA_PROJETOS);
    const dadosProjetos = abaProjetos.getDataRange().getValues();
    
    for (let i = 1; i < dadosProjetos.length; i++) {
      if (dadosProjetos[i][COLUNAS_PROJETOS.SETOR] === nomeSetor) {
        return { 
          sucesso: false, 
          mensagem: 'Não é possível excluir: existem projetos vinculados a este setor' 
        };
      }
    }
    
    aba.deleteRow(linhaSetor);
    return { sucesso: true, mensagem: 'Setor excluído com sucesso!' };
  } catch (e) {
    Logger.log('ERRO excluirSetor: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Obtém todos os setores cadastrados (para datalist e seleção)
 * @returns {Object} - {sucesso, setores: Array}
 */
function obterTodosSetores() {
  try {
    const aba = obterAba(NOME_ABA_SETORES);
    if (!aba || aba.getLastRow() <= 1) {
      return { sucesso: true, setores: [] };
    }
    
    const dados = aba.getDataRange().getValues();
    const setores = [];
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_SETORES.ID] || dados[i][COLUNAS_SETORES.NOME]) {
        setores.push({
          id: dados[i][COLUNAS_SETORES.ID] || '',
          nome: dados[i][COLUNAS_SETORES.NOME] || '',
          descricao: dados[i][COLUNAS_SETORES.DESCRICAO] || '',
          responsavelId: dados[i][COLUNAS_SETORES.RESPONSAVEIS_IDS] || ''
        });
      }
    }
    
    return { sucesso: true, setores: setores };
  } catch (e) {
    Logger.log('ERRO obterTodosSetores: ' + e.toString());
    return { sucesso: false, mensagem: e.message, setores: [] };
  }
}

function vincularResponsavelAoSetor(nomeSetor, responsavelId) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_SETORES);
    const dados = aba.getDataRange().getValues();
    
    // Procurar setor existente
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_SETORES.NOME] === nomeSetor) {
        aba.getRange(i + 1, COLUNAS_SETORES.RESPONSAVEIS_IDS + 1).setValue(responsavelId);
        return { sucesso: true, mensagem: 'Gerente vinculado ao setor!' };
      }
    }
    
    // Se não existe, criar
    const id = gerarId();
    aba.appendRow([id, nomeSetor, responsavelId, null, null]);
    return { sucesso: true, id: id, mensagem: 'Setor criado com gerente!' };
  } catch (e) {
    Logger.log('ERRO vincularResponsavelAoSetor: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

// ==========================================
// CRUD PRIORIDADES POR RESPONSÁVEL
// ==========================================

function obterPrioridadesResponsavel(responsavelId) {
  try {
    const aba = obterAba(NOME_ABA_PRIORIDADES);
    if (!aba || aba.getLastRow() <= 1) return { sucesso: true, prioridades: [] };
    
    const dados = aba.getDataRange().getValues();
    const prioridades = [];
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_PRIORIDADES.RESPONSAVEL_ID] === responsavelId) {
        prioridades.push({
          id: dados[i][COLUNAS_PRIORIDADES.ID],
          responsavelId: dados[i][COLUNAS_PRIORIDADES.RESPONSAVEL_ID],
          tipoItem: dados[i][COLUNAS_PRIORIDADES.TIPO_ITEM],
          itemId: dados[i][COLUNAS_PRIORIDADES.ITEM_ID],
          ordemPrioridade: dados[i][COLUNAS_PRIORIDADES.ORDEM_PRIORIDADE],
          projetoReferencia: dados[i][COLUNAS_PRIORIDADES.PROJETO_REFERENCIA] || null // ← NOVO
        });
      }
    }
    
    // ============================================================
    // NOVO: Ordenação com lógica hierárquica
    // Projetos primeiro (ordenados por prioridade), depois suas etapas
    // ============================================================
    const projetos = prioridades.filter(p => p.tipoItem === 'projeto').sort((a, b) => a.ordemPrioridade - b.ordemPrioridade);
    const resultado = [];
    
    projetos.forEach(projeto => {
      resultado.push(projeto);
      
      // Adicionar etapas deste projeto ordenadas
      const etapasDoProjeto = prioridades
        .filter(p => p.tipoItem === 'etapa' && p.projetoReferencia === projeto.itemId)
        .sort((a, b) => a.ordemPrioridade - b.ordemPrioridade);
      
      resultado.push(...etapasDoProjeto);
    });
    
    return { sucesso: true, prioridades: resultado };
  } catch (e) {
    Logger.log('ERRO obterPrioridadesResponsavel: ' + e.toString());
    return { sucesso: false, mensagem: e.message, prioridades: [] };
  }
}


/**
 * Salva as prioridades de um responsável com hierarquia projeto → etapas
 * @param {string} responsavelId - ID do responsável
 * @param {Array} listaPrioridades - Array de {tipoItem, itemId, projetoReferencia}
 * @returns {Object} - {sucesso, mensagem}
 */
function salvarPrioridadesResponsavel(responsavelId, listaPrioridades) {
  try {
    exigirPermissaoEdicao();
    
    // ============================================================
    // PASSO 1: Obter aba e garantir que ela existe
    // ============================================================
    const aba = obterAba(NOME_ABA_PRIORIDADES);
    const ultimaLinha = aba.getLastRow();
    
    // ============================================================
    // PASSO 2: Identificar linhas a deletar
    // ============================================================
    const linhasParaDeletar = [];
    
    if (ultimaLinha > 1) {
      const dados = aba.getRange(2, 1, ultimaLinha - 1, 6).getValues();
      
      for (let i = 0; i < dados.length; i++) {
        if (dados[i][COLUNAS_PRIORIDADES.RESPONSAVEL_ID] === responsavelId) {
          linhasParaDeletar.push(i + 2);
        }
      }
    }
    
    // ============================================================
    // PASSO 3: Deletar linhas antigas (do fim para o início)
    // ============================================================
    if (linhasParaDeletar.length > 0) {
      linhasParaDeletar.sort((a, b) => b - a);
      
      linhasParaDeletar.forEach(numeroLinha => {
        aba.deleteRow(numeroLinha);
      });
      
      SpreadsheetApp.flush();
      Logger.log(`Deletadas ${linhasParaDeletar.length} prioridades antigas do responsável ${responsavelId}`);
    }
    
    // ============================================================
    // PASSO 4: Validar lista
    // ============================================================
    if (!listaPrioridades || listaPrioridades.length === 0) {
      return { 
        sucesso: true, 
        mensagem: 'Prioridades removidas (lista vazia)' 
      };
    }
    
    // ============================================================
    // PASSO 5: Processar prioridades com contadores separados
    // ============================================================
    const novasLinhas = [];
    let contadorProjetos = 0;
    const contadoresEtapasPorProjeto = {}; // { projetoId: contador }
    
    listaPrioridades.forEach((item) => {
      let ordemPrioridade;
      let projetoReferencia = '';
      
      if (item.tipoItem === 'projeto') {
        // PROJETO: incrementa contador de projetos
        contadorProjetos++;
        ordemPrioridade = contadorProjetos;
        projetoReferencia = ''; // Projetos não têm referência
        
      } else if (item.tipoItem === 'etapa') {
        // ETAPA: incrementa contador específico do projeto pai
        const projetoPai = item.projetoReferencia || '';
        
        if (!projetoPai) {
          Logger.log(`AVISO: Etapa ${item.itemId} sem projetoReferencia!`);
          return; // Pula esta etapa (dados inválidos)
        }
        
        if (!contadoresEtapasPorProjeto[projetoPai]) {
          contadoresEtapasPorProjeto[projetoPai] = 0;
        }
        
        contadoresEtapasPorProjeto[projetoPai]++;
        ordemPrioridade = contadoresEtapasPorProjeto[projetoPai];
        projetoReferencia = projetoPai;
      }
      
      novasLinhas.push([
        gerarId(),                // ID
        responsavelId,            // ResponsavelId
        item.tipoItem,            // TipoItem (projeto ou etapa)
        item.itemId,              // ItemId
        ordemPrioridade,          // OrdemPrioridade (1, 2, 3... independente por tipo)
        projetoReferencia         // ProjetoReferencia (vazio para projetos, ID do projeto para etapas)
      ]);
    });
    
    // ============================================================
    // PASSO 6: Inserir todas as linhas de uma vez
    // ============================================================
    if (novasLinhas.length > 0) {
      const proximaLinha = aba.getLastRow() + 1;
      aba.getRange(proximaLinha, 1, novasLinhas.length, 6).setValues(novasLinhas);
      
      SpreadsheetApp.flush();
      
      Logger.log(`Adicionadas ${novasLinhas.length} novas prioridades para ${responsavelId}`);
      Logger.log(`Projetos: ${contadorProjetos} | Etapas: ${JSON.stringify(contadoresEtapasPorProjeto)}`);
    }
    
    return { 
      sucesso: true, 
      mensagem: `Prioridades salvas! (${contadorProjetos} projetos, ${Object.keys(contadoresEtapasPorProjeto).length} com etapas)` 
    };
    
  } catch (e) {
    Logger.log('ERRO salvarPrioridadesResponsavel: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    return { 
      sucesso: false, 
      mensagem: 'Erro ao salvar: ' + e.message 
    };
  }
}

function testarSalvarPrioridades() {
  const resultado = salvarPrioridadesResponsavel('teste123', [
    { tipoItem: 'projeto', itemId: 'proj1' },
    { tipoItem: 'etapa', itemId: 'etapa1' },
    { tipoItem: 'projeto', itemId: 'proj2' }
  ]);
  
  Logger.log(resultado);
  
  // Verificar se salvou
  const aba = obterAba(NOME_ABA_PRIORIDADES);
  const dados = aba.getDataRange().getValues();
  Logger.log('Total de linhas na aba: ' + dados.length);
  Logger.log('Dados salvos: ' + JSON.stringify(dados));
}

/**
 * Envia email com projetos e prioridades para o colaborador
 * @param {string} responsavelId - ID do responsável
 * @param {Array} projetosIds - IDs dos projetos selecionados para enviar
 * @param {string} mensagemAdicional - Mensagem opcional do gestor
 */
function enviarEmailPrioridadesColaborador(responsavelId, projetosIds, mensagemAdicional) {
  try {
    // Buscar dados do responsável
    const responsaveis = obterTodosResponsaveis();
    const responsavel = responsaveis.find(r => r.id === responsavelId);
    
    if (!responsavel || !responsavel.email) {
      return { sucesso: false, mensagem: 'Colaborador não possui email cadastrado.' };
    }
    
    // Buscar projetos
    const todosProjetos = obterTodosProjetos();
    const todasEtapas = obterTodasEtapas();
    const prioridadesResp = obterPrioridadesResponsavel(responsavelId);
    const mapaPrioridades = new Map();
    
    prioridadesResp.prioridades.forEach(p => {
      mapaPrioridades.set(p.tipoItem + '-' + p.itemId, p.ordemPrioridade);
    });
    
    // Filtrar e ordenar projetos selecionados
    let projetosFiltrados = todosProjetos.filter(p => projetosIds.includes(p.id));
    
    // Ordenar por prioridade
    projetosFiltrados.sort((a, b) => {
      const prioA = mapaPrioridades.get('projeto-' + a.id) || 999;
      const prioB = mapaPrioridades.get('projeto-' + b.id) || 999;
      return prioA - prioB;
    });
    
    // Montar HTML do email
    const htmlEmail = montarHtmlEmailPrioridades(responsavel, projetosFiltrados, todasEtapas, mapaPrioridades, mensagemAdicional);
    
    // Enviar email
    MailApp.sendEmail({
      to: responsavel.email,
      subject: '📋 Suas Prioridades de Projetos - Smart Meeting',
      htmlBody: htmlEmail
    });
    
    return { sucesso: true, mensagem: `Email enviado para ${responsavel.email}!` };
  } catch (e) {
    Logger.log('ERRO enviarEmailPrioridadesColaborador: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Monta o HTML do email de prioridades com pendências
 * @param {Object} responsavel - Dados do responsável
 * @param {Array} projetos - Lista de projetos
 * @param {Array} todasEtapas - Lista de todas as etapas
 * @param {Map} mapaPrioridades - Mapa de prioridades
 * @param {string} mensagemAdicional - Mensagem opcional do gestor
 * @returns {string} HTML do email
 */
function montarHtmlEmailPrioridades(responsavel, projetos, todasEtapas, mapaPrioridades, mensagemAdicional) {
  const dataAtual = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  
  // ============================================================
  // ESTILOS INLINE REUTILIZÁVEIS (compatibilidade com email)
  // ============================================================
  const ESTILOS = {
    corPrimaria: '#b45309',
    corPrimariaClara: '#fef3c7',
    corFundo: '#f5f5f4',
    corTexto: '#44403c',
    corTextoSecundario: '#78716c',
    corBorda: '#e5e7eb',
    corSucesso: '#22c55e',
    corAviso: '#eab308',
    corPerigo: '#ef4444',
    corNeutro: '#6b7280',
    fontePrincipal: "'Segoe UI', Arial, sans-serif"
  };
  
  // ============================================================
  // FUNÇÃO AUXILIAR: Renderiza badge de urgência
  // ============================================================
  function renderizarBadgeUrgencia(urgencia) {
    const cores = {
      alta: { bg: '#fef2f2', cor: '#dc2626', icone: '🔴' },
      media: { bg: '#fefce8', cor: '#ca8a04', icone: '🟡' },
      baixa: { bg: '#f0fdf4', cor: '#16a34a', icone: '🟢' }
    };
    const config = cores[urgencia] || cores.media;
    return `<span style="display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; background: ${config.bg}; color: ${config.cor};">${config.icone} ${urgencia || 'média'}</span>`;
  }
  
  // ============================================================
  // FUNÇÃO AUXILIAR: Renderiza lista de pendências de uma etapa
  // ============================================================
  function renderizarPendenciasEtapa(pendencias) {
    if (!pendencias || !Array.isArray(pendencias) || pendencias.length === 0) {
      return '';
    }
    
    // Separar pendências por status
    const pendentes = pendencias.filter(p => !p.concluido && p.concluido !== 'true');
    const concluidas = pendencias.filter(p => p.concluido === true || p.concluido === 'true');
    
    if (pendentes.length === 0 && concluidas.length === 0) {
      return '';
    }
    
    // Ordenar pendentes por urgência (alta primeiro)
    const ordemUrgencia = { alta: 0, media: 1, baixa: 2 };
    pendentes.sort((a, b) => (ordemUrgencia[a.urgencia] || 1) - (ordemUrgencia[b.urgencia] || 1));
    
    let htmlPendencias = `
      <tr>
        <td colspan="4" style="padding: 0;">
          <div style="margin: 0 10px 10px 10px; background: #fffbeb; border: 1px solid #fde68a; border-radius: 8px; overflow: hidden;">
            <div style="background: linear-gradient(135deg, #fbbf24, #f59e0b); color: white; padding: 8px 12px; font-size: 12px; font-weight: 600;">
              <span style="margin-right: 6px;">⚠️</span>
              Pendências (${pendentes.length} pendente${pendentes.length !== 1 ? 's' : ''}${concluidas.length > 0 ? ' • ' + concluidas.length + ' concluída' + (concluidas.length !== 1 ? 's' : '') : ''})
            </div>
            <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
    `;
    
    // Renderizar pendências pendentes
    pendentes.forEach((pend, idx) => {
      const bgColor = idx % 2 === 0 ? '#fffbeb' : '#fef3c7';
      htmlPendencias += `
        <tr style="background: ${bgColor};">
          <td style="padding: 8px 12px; width: 70px; vertical-align: top;">
            ${renderizarBadgeUrgencia(pend.urgencia)}
          </td>
          <td style="padding: 8px 12px; color: #78350f; line-height: 1.4;">
            ${escaparHtmlEmail(pend.texto || 'Sem descrição')}
          </td>
        </tr>
      `;
    });
    
    // Se houver concluídas, mostrar resumo colapsado
    if (concluidas.length > 0) {
      htmlPendencias += `
        <tr style="background: #f0fdf4;">
          <td colspan="2" style="padding: 8px 12px; color: #166534; font-size: 11px;">
            <span style="margin-right: 4px;">✅</span>
            ${concluidas.length} pendência${concluidas.length !== 1 ? 's' : ''} concluída${concluidas.length !== 1 ? 's' : ''}
          </td>
        </tr>
      `;
    }
    
    htmlPendencias += `
            </table>
          </div>
        </td>
      </tr>
    `;
    
    return htmlPendencias;
  }
  
  // ============================================================
  // FUNÇÃO AUXILIAR: Obter cor por status da etapa
  // ============================================================
  function obterCorStatus(status) {
    const cores = {
      'Concluída': { bg: '#dcfce7', cor: '#166534', borda: '#86efac' },
      'Em Andamento': { bg: '#fef9c3', cor: '#854d0e', borda: '#fde047' },
      'Bloqueada': { bg: '#fee2e2', cor: '#991b1b', borda: '#fca5a5' },
      'A Fazer': { bg: '#f3f4f6', cor: '#374151', borda: '#d1d5db' }
    };
    return cores[status] || cores['A Fazer'];
  }
  
  // ============================================================
  // CONSTRUÇÃO DO HTML DOS PROJETOS
  // ============================================================
  let htmlProjetos = '';
  
  projetos.forEach((projeto, indexProjeto) => {
    const prioridadeProjeto = mapaPrioridades.get('projeto-' + projeto.id) || '-';
    const etapasDoProjeto = todasEtapas.filter(e => e.projetoId === projeto.id);
    
    // Calcular estatísticas do projeto
    const totalEtapas = etapasDoProjeto.length;
    const etapasConcluidas = etapasDoProjeto.filter(e => e.status === 'Concluída').length;
    const percentualConclusao = totalEtapas > 0 ? Math.round((etapasConcluidas / totalEtapas) * 100) : 0;
    
    // Contar pendências totais do projeto
    let totalPendenciasProj = 0;
    let pendenciasUrgentesProj = 0;
    etapasDoProjeto.forEach(e => {
      if (Array.isArray(e.pendencias)) {
        const pendentes = e.pendencias.filter(p => !p.concluido && p.concluido !== 'true');
        totalPendenciasProj += pendentes.length;
        pendenciasUrgentesProj += pendentes.filter(p => p.urgencia === 'alta').length;
      }
    });
    
    // Ordenar etapas por prioridade
    etapasDoProjeto.sort((a, b) => {
      const prioA = mapaPrioridades.get('etapa-' + a.id) || 999;
      const prioB = mapaPrioridades.get('etapa-' + b.id) || 999;
      return prioA - prioB;
    });
    
    // Renderizar etapas
    let htmlEtapas = '';
    etapasDoProjeto.forEach((etapa, idxEtapa) => {
      const prioridadeEtapa = mapaPrioridades.get('etapa-' + etapa.id) || '-';
      const coresStatus = obterCorStatus(etapa.status);
      const bgLinha = idxEtapa % 2 === 0 ? '#ffffff' : '#fafaf9';
      
      // Linha principal da etapa
      htmlEtapas += `
        <tr style="background: ${bgLinha}; border-bottom: 1px solid ${ESTILOS.corBorda};">
          <td style="padding: 12px 10px; text-align: center; font-weight: bold; color: ${ESTILOS.corPrimaria}; width: 50px; vertical-align: top;">
            ${prioridadeEtapa !== '-' ? '#' + prioridadeEtapa : '-'}
          </td>
          <td style="padding: 12px 10px; vertical-align: top;">
            <div style="font-weight: 600; color: ${ESTILOS.corTexto}; margin-bottom: 4px;">
              ${escaparHtmlEmail(etapa.nome)}
            </div>
            ${etapa.oQueFazer ? `
              <div style="font-size: 12px; color: ${ESTILOS.corTextoSecundario}; line-height: 1.4;">
                <strong style="color: ${ESTILOS.corTexto};">O que fazer:</strong> ${escaparHtmlEmail(etapa.oQueFazer)}
              </div>
            ` : ''}
          </td>
          <td style="padding: 12px 10px; text-align: center; width: 110px; vertical-align: top;">
            <span style="display: inline-block; background: ${coresStatus.bg}; color: ${coresStatus.cor}; border: 1px solid ${coresStatus.borda}; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 500;">
              ${etapa.status}
            </span>
          </td>
        </tr>
      `;
      
      // Renderizar pendências da etapa (se houver)
      htmlEtapas += renderizarPendenciasEtapa(etapa.pendencias);
    });
    
    // Determinar cor do progresso
    let corProgresso = '#3b82f6';
    if (percentualConclusao === 100) corProgresso = '#22c55e';
    else if (percentualConclusao >= 50) corProgresso = '#eab308';
    else if (percentualConclusao > 0) corProgresso = '#f97316';
    
    // Montar card do projeto
    htmlProjetos += `
      <div style="margin-bottom: 24px; border: 1px solid ${ESTILOS.corBorda}; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06);">
        
        <!-- Header do Projeto -->
        <div style="background: linear-gradient(135deg, #44403c, #292524); color: white; padding: 16px;">
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="width: 44px; vertical-align: top;">
                <div style="background: ${ESTILOS.corPrimaria}; color: white; width: 40px; height: 40px; border-radius: 50%; text-align: center; line-height: 40px; font-weight: bold; font-size: 14px;">
                  ${prioridadeProjeto !== '-' ? '#' + prioridadeProjeto : '-'}
                </div>
              </td>
              <td style="padding-left: 12px; vertical-align: top;">
                <div style="font-size: 16px; font-weight: 600; margin-bottom: 4px;">
                  ${escaparHtmlEmail(projeto.nome)}
                </div>
                <div style="font-size: 12px; opacity: 0.85;">
                  ${escaparHtmlEmail(projeto.setor || 'Sem setor')}
                  ${projeto.pilar ? ' • ' + escaparHtmlEmail(projeto.pilar) : ''}
                </div>
              </td>
              <td style="text-align: right; vertical-align: top; width: 120px;">
                <div style="font-size: 20px; font-weight: bold;">${percentualConclusao}%</div>
                <div style="font-size: 11px; opacity: 0.8;">${etapasConcluidas}/${totalEtapas} etapas</div>
              </td>
            </tr>
          </table>
          
          <!-- Barra de progresso -->
          <div style="margin-top: 12px; background: rgba(255,255,255,0.2); border-radius: 4px; height: 6px; overflow: hidden;">
            <div style="background: ${corProgresso}; height: 100%; width: ${percentualConclusao}%; border-radius: 4px;"></div>
          </div>
        </div>
        
        <!-- Descrição do Projeto (se houver) -->
        ${projeto.descricao ? `
          <div style="padding: 12px 16px; background: #fffbf2; border-bottom: 1px solid ${ESTILOS.corBorda}; color: ${ESTILOS.corTextoSecundario}; font-size: 13px; line-height: 1.5;">
            ${escaparHtmlEmail(projeto.descricao)}
          </div>
        ` : ''}
        
        <!-- Resumo de Pendências do Projeto -->
        ${totalPendenciasProj > 0 ? `
          <div style="padding: 10px 16px; background: ${pendenciasUrgentesProj > 0 ? '#fef2f2' : '#fffbeb'}; border-bottom: 1px solid ${ESTILOS.corBorda};">
            <span style="font-size: 12px; color: ${pendenciasUrgentesProj > 0 ? '#dc2626' : '#ca8a04'}; font-weight: 500;">
              ⚠️ ${totalPendenciasProj} pendência${totalPendenciasProj !== 1 ? 's' : ''} no projeto
              ${pendenciasUrgentesProj > 0 ? ` (${pendenciasUrgentesProj} urgente${pendenciasUrgentesProj !== 1 ? 's' : ''})` : ''}
            </span>
          </div>
        ` : ''}
        
        <!-- Tabela de Etapas -->
        ${etapasDoProjeto.length > 0 ? `
          <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
            <thead>
              <tr style="background: ${ESTILOS.corFundo}; color: ${ESTILOS.corTextoSecundario};">
                <th style="padding: 10px; text-align: center; width: 50px; font-weight: 600;">Prio</th>
                <th style="padding: 10px; text-align: left; font-weight: 600;">Etapa / Detalhes</th>
                <th style="padding: 10px; text-align: center; width: 110px; font-weight: 600;">Status</th>
              </tr>
            </thead>
            <tbody>
              ${htmlEtapas}
            </tbody>
          </table>
        ` : `
          <div style="padding: 20px; text-align: center; color: ${ESTILOS.corTextoSecundario}; font-style: italic;">
            Nenhuma etapa cadastrada neste projeto
          </div>
        `}
      </div>
    `;
  });
  
  // ============================================================
  // ESTATÍSTICAS GERAIS
  // ============================================================
  const totalProjetosEmail = projetos.length;
  let totalEtapasEmail = 0;
  let totalPendenciasEmail = 0;
  let totalUrgentesEmail = 0;
  
  projetos.forEach(p => {
    const etapasP = todasEtapas.filter(e => e.projetoId === p.id);
    totalEtapasEmail += etapasP.length;
    etapasP.forEach(e => {
      if (Array.isArray(e.pendencias)) {
        const pendentes = e.pendencias.filter(pend => !pend.concluido && pend.concluido !== 'true');
        totalPendenciasEmail += pendentes.length;
        totalUrgentesEmail += pendentes.filter(pend => pend.urgencia === 'alta').length;
      }
    });
  });
  
  // ============================================================
  // HTML FINAL DO EMAIL
  // ============================================================
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Suas Prioridades - Smart Meeting</title>
    </head>
    <body style="font-family: ${ESTILOS.fontePrincipal}; background: ${ESTILOS.corFundo}; margin: 0; padding: 20px; -webkit-font-smoothing: antialiased;">
      
      <!-- Container Principal -->
      <div style="max-width: 700px; margin: 0 auto; background: white; border-radius: 16px; box-shadow: 0 4px 24px rgba(0,0,0,0.12); overflow: hidden;">
        
        <!-- ====== HEADER ====== -->
        <div style="background: linear-gradient(135deg, ${ESTILOS.corPrimaria}, #92400e); color: white; padding: 32px; text-align: center;">
          <div style="font-size: 36px; margin-bottom: 8px;">📋</div>
          <h1 style="margin: 0 0 8px 0; font-size: 24px; font-weight: 700;">Suas Prioridades</h1>
          <p style="margin: 0; opacity: 0.9; font-size: 14px;">Projetos e etapas organizados para você</p>
        </div>
        
        <!-- ====== SAUDAÇÃO ====== -->
        <div style="padding: 24px; border-bottom: 1px solid ${ESTILOS.corBorda};">
          <p style="margin: 0; font-size: 16px; color: ${ESTILOS.corTexto};">
            Olá, <strong>${escaparHtmlEmail(responsavel.nome)}</strong>! 👋
          </p>
          
          ${mensagemAdicional ? `
            <div style="margin-top: 16px; padding: 14px; background: linear-gradient(135deg, #fffbeb, #fef3c7); border-left: 4px solid ${ESTILOS.corPrimaria}; border-radius: 0 8px 8px 0;">
              <div style="font-size: 11px; color: ${ESTILOS.corPrimaria}; font-weight: 600; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.5px;">
                💬 Mensagem do Gestor
              </div>
              <p style="margin: 0; font-size: 14px; color: ${ESTILOS.corTexto}; line-height: 1.5;">${escaparHtmlEmail(mensagemAdicional)}</p>
            </div>
          ` : ''}
          
          <p style="margin: 16px 0 0 0; font-size: 13px; color: ${ESTILOS.corTextoSecundario}; line-height: 1.5;">
            Segue abaixo a lista de projetos e etapas sob sua responsabilidade, ordenados por prioridade.
            ${totalPendenciasEmail > 0 ? `<br><strong style="color: ${totalUrgentesEmail > 0 ? ESTILOS.corPerigo : ESTILOS.corAviso};">Atenção:</strong> Você tem ${totalPendenciasEmail} pendência${totalPendenciasEmail !== 1 ? 's' : ''} aguardando ação${totalUrgentesEmail > 0 ? `, sendo ${totalUrgentesEmail} urgente${totalUrgentesEmail !== 1 ? 's' : ''}` : ''}.` : ''}
          </p>
        </div>
        
        <!-- ====== RESUMO RÁPIDO ====== -->
        <div style="padding: 20px 24px; background: ${ESTILOS.corFundo}; border-bottom: 1px solid ${ESTILOS.corBorda};">
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="text-align: center; padding: 10px;">
                <div style="font-size: 28px; font-weight: 700; color: ${ESTILOS.corPrimaria};">${totalProjetosEmail}</div>
                <div style="font-size: 11px; color: ${ESTILOS.corTextoSecundario}; text-transform: uppercase; letter-spacing: 0.5px;">Projetos</div>
              </td>
              <td style="text-align: center; padding: 10px; border-left: 1px solid ${ESTILOS.corBorda}; border-right: 1px solid ${ESTILOS.corBorda};">
                <div style="font-size: 28px; font-weight: 700; color: #3b82f6;">${totalEtapasEmail}</div>
                <div style="font-size: 11px; color: ${ESTILOS.corTextoSecundario}; text-transform: uppercase; letter-spacing: 0.5px;">Etapas</div>
              </td>
              <td style="text-align: center; padding: 10px;">
                <div style="font-size: 28px; font-weight: 700; color: ${totalUrgentesEmail > 0 ? ESTILOS.corPerigo : ESTILOS.corAviso};">${totalPendenciasEmail}</div>
                <div style="font-size: 11px; color: ${ESTILOS.corTextoSecundario}; text-transform: uppercase; letter-spacing: 0.5px;">Pendências</div>
              </td>
            </tr>
          </table>
        </div>
        
        <!-- ====== PROJETOS ====== -->
        <div style="padding: 24px;">
          ${htmlProjetos}
        </div>
        
        <!-- ====== LEGENDA ====== -->
        <div style="padding: 16px 24px; background: ${ESTILOS.corFundo}; border-top: 1px solid ${ESTILOS.corBorda};">
          <div style="font-size: 11px; color: ${ESTILOS.corTextoSecundario}; margin-bottom: 8px; font-weight: 600;">LEGENDA DE URGÊNCIA:</div>
          <table style="border-collapse: collapse;">
            <tr>
              <td style="padding: 4px 12px 4px 0;">
                <span style="display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; background: #fef2f2; color: #dc2626;">🔴 Alta</span>
              </td>
              <td style="padding: 4px 12px 4px 0;">
                <span style="display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; background: #fefce8; color: #ca8a04;">🟡 Média</span>
              </td>
              <td style="padding: 4px 12px 4px 0;">
                <span style="display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; background: #f0fdf4; color: #16a34a;">🟢 Baixa</span>
              </td>
            </tr>
          </table>
        </div>
        
        <!-- ====== FOOTER ====== -->
        <div style="background: #292524; color: white; padding: 24px; text-align: center;">
          <p style="margin: 0 0 8px 0; font-size: 12px; opacity: 0.8;">
            Enviado em ${dataAtual}
          </p>
          <p style="margin: 0; font-size: 14px; font-weight: 600;">
            Smart Meeting - Gestão de Projetos
          </p>
          <p style="margin: 12px 0 0 0; font-size: 11px; opacity: 0.6;">
            Este é um email automático. Por favor, não responda diretamente.
          </p>
        </div>
      </div>
      
    </body>
    </html>
  `;
}

/**
 * Função auxiliar para escapar HTML em emails
 * @param {string} texto - Texto a ser escapado
 * @returns {string} Texto escapado
 */
function escaparHtmlEmail(texto) {
  if (texto == null) return '';
  return String(texto)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function migrarPrioridadesAntigas() {
  const aba = obterAba(NOME_ABA_PRIORIDADES);
  const dados = aba.getDataRange().getValues();
  
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_PRIORIDADES.TIPO_ITEM] === 'etapa') {
      // Tentar descobrir o projeto pai da etapa
      const etapa = dadosDiagrama.etapas.find(e => e.id === dados[i][COLUNAS_PRIORIDADES.ITEM_ID]);
      if (etapa && etapa.projetoId) {
        aba.getRange(i + 1, COLUNAS_PRIORIDADES.PROJETO_REFERENCIA + 1).setValue(etapa.projetoId);
      }
    }
  }
  
  SpreadsheetApp.flush();
  Logger.log('Migração concluída');
}

function testeEnvioEmail() {
  try {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: 'Teste de Permissão',
      body: 'Se você recebeu este email, a permissão está funcionando!'
    });
    Logger.log('Email enviado com sucesso!');
  } catch (e) {
    Logger.log('ERRO: ' + e.message);
  }
}

/**
 * Obtém todas as pendências de um responsável específico
 * @param {string} responsavelId - ID do responsável
 * @returns {Object} Resumo das pendências do responsável
 */
function obterPendenciasPorResponsavel(responsavelId) {
  try {
    const todasEtapas = obterTodasEtapas();
    const resumo = {
      total: 0,
      pendentes: 0,
      concluidas: 0,
      porUrgencia: { alta: 0, media: 0, baixa: 0 },
      itens: []
    };
    
    todasEtapas.forEach(etapa => {
      const ehResponsavel = etapa.responsaveisIds.includes(responsavelId);
      if (!ehResponsavel) return;
      
      if (Array.isArray(etapa.pendencias)) {
        etapa.pendencias.forEach(pend => {
          resumo.total++;
          if (pend.concluido) {
            resumo.concluidas++;
          } else {
            resumo.pendentes++;
            resumo.porUrgencia[pend.urgencia || 'media']++;
          }
          resumo.itens.push({
            ...pend,
            etapaId: etapa.id,
            etapaNome: etapa.nome,
            projetoId: etapa.projetoId
          });
        });
      }
    });
    
    // Ordenar por urgência (alta primeiro) e depois por não concluídos
    resumo.itens.sort((a, b) => {
      if (a.concluido !== b.concluido) return a.concluido ? 1 : -1;
      const ordemUrgencia = { alta: 0, media: 1, baixa: 2 };
      return (ordemUrgencia[a.urgencia] || 1) - (ordemUrgencia[b.urgencia] || 1);
    });
    
    return { sucesso: true, resumo };
  } catch (e) {
    Logger.log('ERRO obterPendenciasPorResponsavel: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Alterna o status de conclusão de uma pendência específica
 * @param {string} etapaId - ID da etapa
 * @param {string} pendenciaId - ID da pendência
 * @returns {Object} Resultado da operação
 */
function alternarConclusaoPendencia(etapaId, pendenciaId) {
  try {
    exigirPermissaoEdicao();
    const aba = obterAba(NOME_ABA_ETAPAS);
    const dados = aba.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_ETAPAS.ID] === etapaId) {
        const pendenciasRaw = dados[i][COLUNAS_ETAPAS.PENDENCIAS] || '[]';
        let pendencias = [];
        
        try {
          pendencias = JSON.parse(pendenciasRaw);
        } catch (e) {
          return { sucesso: false, mensagem: 'Erro ao parsear pendências' };
        }
        
        const idx = pendencias.findIndex(p => p.id === pendenciaId);
        if (idx === -1) return { sucesso: false, mensagem: 'Pendência não encontrada' };
        
        pendencias[idx].concluido = !pendencias[idx].concluido;
        
        aba.getRange(i + 1, COLUNAS_ETAPAS.PENDENCIAS + 1).setValue(JSON.stringify(pendencias));
        
        return { 
          sucesso: true, 
          concluido: pendencias[idx].concluido,
          mensagem: pendencias[idx].concluido ? 'Pendência concluída!' : 'Pendência reaberta'
        };
      }
    }
    return { sucesso: false, mensagem: 'Etapa não encontrada' };
  } catch (e) {
    Logger.log('ERRO alternarConclusaoPendencia: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}
