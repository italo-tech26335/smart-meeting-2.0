/** =====================================================================
 *                        AUTH.JS — SISTEMA DE AUTENTICAÇÃO
 *  Login, sessão, brute-force, reset de senha, gestão de usuários e logs.
 *  Todas as funções públicas são chamáveis via google.script.run.
 * ===================================================================== */

// ── Configurações de Email ───────────────────────────────────────────────────
const EMAIL_REMETENTE_NOME = 'Smart Meeting';
const EMAIL_REPLY_TO       = 'setorbiunichristus@gmail.com';
const URL_SISTEMA          = 'https://script.google.com/macros/s/AKfycbwJB8g7DHPjcCX4Bl2rclQze_TpOyZ0PB9sEFgHDNLdzCihG8AjHPXWvsOsoqvu1bh7/exec';

// ── Constantes ──────────────────────────────────────────────────────────────
const NOME_ABA_USUARIOS = 'Usuarios';
const NOME_ABA_LOGS     = 'Logs';

const COLUNAS_USUARIOS = {
  ID:                0,
  EMAIL:             1,
  SENHA_HASH:        2,
  SALT:              3,
  NOME:              4,
  PERFIL:            5,   // admin | usuario | visualizador
  ATIVO:             6,
  CRIADO_EM:         7,
  ULTIMO_LOGIN:      8,
  TENTATIVAS_LOGIN:  9,
  BLOQUEADO_ATE:     10,
  DEPARTAMENTOS_IDS: 11, // IDs separados por vírgula; vazio = sem restrição (admin)
  PAGINAS_PERMITIDAS:12  // IDs de páginas separados por vírgula: reunioes,projeto,relatorios; vazio = todas
};

const COLUNAS_LOGS = {
  TIMESTAMP: 0,
  EVENTO:    1,
  USUARIO:   2,
  IP:        3,
  DETALHES:  4,
  RESULTADO: 5
};

const SESSAO_TTL_MS        = 8  * 60 * 60 * 1000;  // 8 horas
const RESET_TTL_MS         = 30 * 60 * 1000;        // 30 minutos
const MAX_TENTATIVAS       = 5;
const BLOQUEIO_NIVEL_1_MS  = 15 * 60 * 1000;        // 15 min
const BLOQUEIO_NIVEL_2_MS  = 60 * 60 * 1000;        // 1 hora
const BLOQUEIO_NIVEL_3_MS  = 24 * 60 * 60 * 1000;   // 24 horas
const RATE_LIMIT_JANELA_MS = 60 * 1000;             // 1 minuto
const RATE_LIMIT_MAX       = 10;                    // requisições/min

// ── Prefixos das chaves no PropertiesService ─────────────────────────────────
const PFX_SESSAO = 'sessao_';
const PFX_BF     = 'bf_';    // brute-force
const PFX_RL     = 'rl_';    // rate-limit
const PFX_RESET  = 'reset_';

// ============================================================================
//  FUNÇÕES PÚBLICAS — google.script.run
// ============================================================================

/**
 * Autentica um usuário com email e senha.
 * @returns {Object} {sucesso, token, usuario, mensagem}
 */
function fazerLogin(email, senha) {
  try {
    if (!email || !senha) {
      return { sucesso: false, mensagem: 'Email e senha são obrigatórios.' };
    }

    email = email.toString().trim().toLowerCase();

    // Rate limiting
    var rl = _verificarRateLimit(email);
    if (rl.bloqueado) {
      return { sucesso: false, mensagem: 'Muitas tentativas. Aguarde 1 minuto.' };
    }

    // Brute-force
    var bf = _verificarBruteForce(email);
    if (bf.bloqueado) {
      _registrarLog('login_falha', email, '', 'conta bloqueada', 'bloqueado');
      return { sucesso: false, mensagem: bf.mensagem };
    }

    // Buscar usuário
    var usuario = _buscarUsuarioPorEmail(email);
    if (!usuario) {
      _registrarTentativaFalha(email);
      _registrarLog('login_falha', email, '', 'usuário não encontrado', 'falha');
      return { sucesso: false, mensagem: 'Email ou senha incorretos.' };
    }

    if (!usuario.ativo) {
      _registrarLog('login_falha', email, '', 'conta inativa', 'bloqueado');
      return { sucesso: false, mensagem: 'Conta inativa. Contate o administrador.' };
    }

    // Primeiro acesso — usuário ainda sem senha definida
    if (usuario.senhaHash === 'PRIMEIRO_ACESSO') {
      _registrarLog('primeiro_acesso', email, '', 'redirecionado para definir PIN', 'info');
      return { sucesso: false, primeiroAcesso: true, email: email, nome: usuario.nome };
    }

    // Verificar senha
    var hashCalculado = _hashSenha(senha.toString(), usuario.salt);
    if (hashCalculado !== usuario.senhaHash) {
      _registrarTentativaFalha(email);
      _registrarLog('login_falha', email, '', 'senha incorreta', 'falha');

      var tentativasRestantes = MAX_TENTATIVAS - _obterTentativasBf(email);
      if (tentativasRestantes <= 0) {
        _registrarLog('usuario_bloqueado', email, '', 'excedeu tentativas', 'bloqueado');
        return { sucesso: false, mensagem: 'Conta bloqueada por 15 minutos devido a muitas tentativas.' };
      }
      return {
        sucesso: false,
        mensagem: 'Email ou senha incorretos. Tentativas restantes: ' + tentativasRestantes
      };
    }

    // Login bem-sucedido
    _zerarTentativas(email);
    _atualizarUltimoLogin(email);

    // Invalidar sessão anterior (single session)
    _invalidarSessoesDoUsuario(email);

    // Criar nova sessão
    var token = _criarToken();
    _salvarSessao(token, email, usuario.perfil, usuario.nome, usuario.departamentosIds, usuario.paginasPermitidas);

    _registrarLog('login_sucesso', email, '', 'login realizado', 'sucesso');

    return {
      sucesso: true,
      token: token,
      usuario: { email: usuario.email, nome: usuario.nome, perfil: usuario.perfil }
    };
  } catch (e) {
    Logger.log('ERRO fazerLogin: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro interno. Tente novamente.' };
  }
}

/**
 * Encerra a sessão de um usuário.
 * @returns {Object} {sucesso}
 */
function fazerLogout(token) {
  try {
    if (!token) return { sucesso: false };
    var sessao = _obterSessao(token);
    if (sessao) {
      _registrarLog('logout', sessao.email, '', 'logout realizado', 'sucesso');
      _invalidarSessao(token);
    }
    return { sucesso: true };
  } catch (e) {
    Logger.log('ERRO fazerLogout: ' + e.toString());
    return { sucesso: false };
  }
}

/**
 * Verifica se um token de sessão é válido.
 * @returns {Object} {valida, usuario, mensagem}
 */
function verificarSessao(token) {
  try {
    if (!token) return { valida: false, mensagem: 'Token ausente.' };
    var sessao = _obterSessao(token);
    if (!sessao) return { valida: false, mensagem: 'Sessão inválida ou expirada.' };

    // Renovar sessão (sliding)
    _renovarSessao(token, sessao);

    return {
      valida: true,
      usuario: { email: sessao.email, nome: sessao.nome, perfil: sessao.perfil }
    };
  } catch (e) {
    Logger.log('ERRO verificarSessao: ' + e.toString());
    return { valida: false, mensagem: 'Erro ao verificar sessão.' };
  }
}

/**
 * Inicia o fluxo de redefinição de senha.
 * @returns {Object} {sucesso, mensagem}
 */
function solicitarResetSenha(email) {
  try {
    if (!email) return { sucesso: false, mensagem: 'Email obrigatório.' };
    email = email.toString().trim().toLowerCase();

    var usuario = _buscarUsuarioPorEmail(email);
    // Sempre retorna sucesso (segurança: não revelar se email existe)
    if (!usuario) {
      return { sucesso: true, mensagem: 'Se o email estiver cadastrado, você receberá as instruções.' };
    }

    var tokenReset = _criarToken();
    var expiry = Date.now() + RESET_TTL_MS;
    PropertiesService.getScriptProperties().setProperty(
      PFX_RESET + tokenReset,
      JSON.stringify({ email: email, expiry: expiry })
    );

    var urlApp = ScriptApp.getService().getUrl();
    var linkReset = urlApp + '?pagina=reset&rt=' + tokenReset;

    MailApp.sendEmail({
      to: email,
      subject: 'Smart Meeting — Redefinição de Senha',
      htmlBody: _htmlResetSenha(usuario.nome, linkReset),
      name: EMAIL_REMETENTE_NOME,
      replyTo: EMAIL_REPLY_TO
    });

    _registrarLog('reset_senha', email, '', 'token de reset enviado', 'sucesso');
    return { sucesso: true, mensagem: 'Se o email estiver cadastrado, você receberá as instruções.' };
  } catch (e) {
    Logger.log('ERRO solicitarResetSenha: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro ao processar solicitação.' };
  }
}

/**
 * Redefine a senha usando um token de reset.
 * @returns {Object} {sucesso, mensagem}
 */
function redefinirSenha(tokenReset, novaSenha) {
  try {
    if (!tokenReset || !novaSenha) {
      return { sucesso: false, mensagem: 'Dados incompletos.' };
    }

    var props = PropertiesService.getScriptProperties();
    var dadosReset = props.getProperty(PFX_RESET + tokenReset);
    if (!dadosReset) {
      return { sucesso: false, mensagem: 'Token inválido ou expirado.' };
    }

    var resetObj = JSON.parse(dadosReset);
    if (Date.now() > resetObj.expiry) {
      props.deleteProperty(PFX_RESET + tokenReset);
      return { sucesso: false, mensagem: 'Token expirado. Solicite um novo.' };
    }

    var validacao = _validarForcaSenha(novaSenha.toString());
    if (!validacao.valida) {
      return { sucesso: false, mensagem: validacao.mensagem };
    }

    var salt = _gerarSalt();
    var hash = _hashSenha(novaSenha.toString(), salt);
    _atualizarSenhaUsuario(resetObj.email, hash, salt);

    props.deleteProperty(PFX_RESET + tokenReset);
    _registrarLog('reset_senha', resetObj.email, '', 'senha redefinida com sucesso', 'sucesso');

    return { sucesso: true, mensagem: 'Senha redefinida com sucesso! Faça login.' };
  } catch (e) {
    Logger.log('ERRO redefinirSenha: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro ao redefinir senha.' };
  }
}

/**
 * Define o PIN (4 dígitos) no primeiro acesso do usuário.
 * Não requer sessão — identificado pelo email.
 */
function definirSenhaPrimeiroAcesso(email, novaSenha) {
  try {
    if (!email || !novaSenha) return { sucesso: false, mensagem: 'Dados incompletos.' };
    email = email.toString().trim().toLowerCase();

    var usuario = _buscarUsuarioPorEmail(email);
    if (!usuario) return { sucesso: false, mensagem: 'Usuário não encontrado.' };
    if (!usuario.ativo) return { sucesso: false, mensagem: 'Conta inativa. Contate o administrador.' };
    if (usuario.senhaHash !== 'PRIMEIRO_ACESSO') {
      return { sucesso: false, mensagem: 'Este usuário já possui senha definida. Use "Esqueci minha senha" se necessário.' };
    }

    var validacao = _validarForcaSenha(novaSenha.toString());
    if (!validacao.valida) return { sucesso: false, mensagem: validacao.mensagem };

    var salt = _gerarSalt();
    var hash = _hashSenha(novaSenha.toString(), salt);
    _atualizarSenhaUsuario(email, hash, salt);

    _registrarLog('primeiro_acesso_concluido', email, '', 'PIN definido com sucesso', 'sucesso');
    return { sucesso: true, mensagem: 'PIN criado! Faça login agora.' };
  } catch (e) {
    Logger.log('ERRO definirSenhaPrimeiroAcesso: ' + e.toString());
    return { sucesso: false, mensagem: 'Erro interno.' };
  }
}

/**
 * Retorna dados do usuário logado com base no token.
 * @returns {Object} {sucesso, usuario}
 */
function obterDadosUsuarioLogado(token) {
  try {
    var sessao = _obterSessao(token);
    if (!sessao) return { sucesso: false, mensagem: 'Sessão inválida.' };
    return {
      sucesso: true,
      usuario: { email: sessao.email, nome: sessao.nome, perfil: sessao.perfil }
    };
  } catch (e) {
    return { sucesso: false };
  }
}

// ============================================================================
//  FUNÇÕES ADMIN — requerem perfil admin
// ============================================================================

/**
 * Lista todos os usuários. Requer admin.
 */
function listarUsuarios(token) {
  try {
    var sessao = _sessaoAdmin(token);
    if (!sessao) return { sucesso: false, mensagem: 'Acesso negado.' };

    var dados = obterDadosAbaComCache(NOME_ABA_USUARIOS);
    if (!dados || dados.length <= 1) return { sucesso: true, usuarios: [] };

    var usuarios = [];
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][COLUNAS_USUARIOS.ID]) {
        var bloqueadoAte = dados[i][COLUNAS_USUARIOS.BLOQUEADO_ATE];
        var estaBloqueado = bloqueadoAte && new Date(bloqueadoAte).getTime() > Date.now();
        var rawDeps = (dados[i][COLUNAS_USUARIOS.DEPARTAMENTOS_IDS] || '').toString();
        usuarios.push({
          id:                dados[i][COLUNAS_USUARIOS.ID],
          email:             dados[i][COLUNAS_USUARIOS.EMAIL],
          nome:              dados[i][COLUNAS_USUARIOS.NOME],
          perfil:            dados[i][COLUNAS_USUARIOS.PERFIL],
          ativo:             dados[i][COLUNAS_USUARIOS.ATIVO] === true || dados[i][COLUNAS_USUARIOS.ATIVO] === 'true',
          criadoEm:          dados[i][COLUNAS_USUARIOS.CRIADO_EM] ? new Date(dados[i][COLUNAS_USUARIOS.CRIADO_EM]).toLocaleString('pt-BR') : '',
          ultimoLogin:       dados[i][COLUNAS_USUARIOS.ULTIMO_LOGIN] ? new Date(dados[i][COLUNAS_USUARIOS.ULTIMO_LOGIN]).toLocaleString('pt-BR') : 'Nunca',
          bloqueado:         estaBloqueado,
          bloqueadoAte:      estaBloqueado ? new Date(bloqueadoAte).toLocaleString('pt-BR') : null,
          primeiroAcesso:    dados[i][COLUNAS_USUARIOS.SENHA_HASH] === 'PRIMEIRO_ACESSO',
          departamentosIds:  rawDeps ? rawDeps.split(',').map(function(d) { return d.trim(); }).filter(function(d) { return d !== ''; }) : [],
          paginasPermitidas: (function() { var raw = (dados[i][COLUNAS_USUARIOS.PAGINAS_PERMITIDAS] || '').toString().trim(); return raw ? raw.split(',').map(function(p) { return p.trim(); }).filter(function(p) { return p !== ''; }) : []; })()
        });
      }
    }
    return { sucesso: true, usuarios: usuarios };
  } catch (e) {
    Logger.log('ERRO listarUsuarios: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Cria um novo usuário. Requer admin.
 */
function criarUsuario(token, dados) {
  try {
    var sessao = _sessaoAdmin(token);
    if (!sessao) return { sucesso: false, mensagem: 'Acesso negado.' };

    if (!dados.email || !dados.nome || !dados.perfil) {
      return { sucesso: false, mensagem: 'Campos obrigatórios: email, nome, perfil.' };
    }

    var email = dados.email.toString().trim().toLowerCase();
    if (_buscarUsuarioPorEmail(email)) {
      return { sucesso: false, mensagem: 'Email já cadastrado.' };
    }

    // Usuário criado sem senha — aguarda primeiro acesso para definir PIN
    var hash = 'PRIMEIRO_ACESSO';
    var salt = '';
    var id = gerarId();
    var depIds = '';
    if (dados.departamentosIds) {
      depIds = Array.isArray(dados.departamentosIds)
        ? dados.departamentosIds.join(',')
        : dados.departamentosIds.toString();
    }
    var pagIds = '';
    if (dados.paginasPermitidas) {
      pagIds = Array.isArray(dados.paginasPermitidas)
        ? dados.paginasPermitidas.join(',')
        : dados.paginasPermitidas.toString();
    }

    var aba = obterAba(NOME_ABA_USUARIOS);
    aba.appendRow([
      id,
      email,
      hash,
      salt,
      dados.nome.toString().trim(),
      dados.perfil || 'usuario',
      true,
      new Date(),
      '',
      0,
      '',
      depIds,
      pagIds
    ]);
    limparCacheAba(NOME_ABA_USUARIOS);

    if (dados.enviarEmail !== false) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: 'Smart Meeting — Bem-vindo(a), ' + dados.nome + '!',
          htmlBody: _htmlBoasVindas(dados.nome, email),
          name: EMAIL_REMETENTE_NOME,
          replyTo: EMAIL_REPLY_TO
        });
      } catch (mailErr) {
        Logger.log('Aviso: não foi possível enviar email de boas-vindas: ' + mailErr.toString());
      }
    }

    _registrarLog('criar_usuario', sessao.email, '', 'usuário criado: ' + email, 'sucesso');
    return { sucesso: true, mensagem: 'Usuário criado com sucesso!' };
  } catch (e) {
    Logger.log('ERRO criarUsuario: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Atualiza dados de um usuário. Requer admin.
 */
function atualizarUsuario(token, dados) {
  try {
    var sessao = _sessaoAdmin(token);
    if (!sessao) return { sucesso: false, mensagem: 'Acesso negado.' };

    if (!dados.id) return { sucesso: false, mensagem: 'ID obrigatório.' };

    var aba = obterAba(NOME_ABA_USUARIOS);
    var rows = aba.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][COLUNAS_USUARIOS.ID] === dados.id) {
        var linha = i + 1;
        if (dados.nome             !== undefined) aba.getRange(linha, COLUNAS_USUARIOS.NOME              + 1).setValue(dados.nome);
        if (dados.perfil           !== undefined) aba.getRange(linha, COLUNAS_USUARIOS.PERFIL            + 1).setValue(dados.perfil);
        if (dados.ativo            !== undefined) aba.getRange(linha, COLUNAS_USUARIOS.ATIVO             + 1).setValue(dados.ativo);
        if (dados.departamentosIds !== undefined) {
          var depStr = Array.isArray(dados.departamentosIds)
            ? dados.departamentosIds.join(',')
            : dados.departamentosIds.toString();
          aba.getRange(linha, COLUNAS_USUARIOS.DEPARTAMENTOS_IDS + 1).setValue(depStr);
        }
        if (dados.paginasPermitidas !== undefined) {
          var pagStr = Array.isArray(dados.paginasPermitidas)
            ? dados.paginasPermitidas.join(',')
            : dados.paginasPermitidas.toString();
          aba.getRange(linha, COLUNAS_USUARIOS.PAGINAS_PERMITIDAS + 1).setValue(pagStr);
        }
        limparCacheAba(NOME_ABA_USUARIOS);
        _registrarLog('atualizar_usuario', sessao.email, '', 'usuário atualizado: ' + rows[i][COLUNAS_USUARIOS.EMAIL], 'sucesso');
        return { sucesso: true, mensagem: 'Usuário atualizado!' };
      }
    }
    return { sucesso: false, mensagem: 'Usuário não encontrado.' };
  } catch (e) {
    Logger.log('ERRO atualizarUsuario: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Desbloqueia um usuário bloqueado por brute-force. Requer admin.
 */
function desbloquearUsuario(token, userId) {
  try {
    var sessao = _sessaoAdmin(token);
    if (!sessao) return { sucesso: false, mensagem: 'Acesso negado.' };

    var aba = obterAba(NOME_ABA_USUARIOS);
    var rows = aba.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][COLUNAS_USUARIOS.ID] === userId) {
        var linha = i + 1;
        var emailUsuario = rows[i][COLUNAS_USUARIOS.EMAIL];
        aba.getRange(linha, COLUNAS_USUARIOS.TENTATIVAS_LOGIN + 1).setValue(0);
        aba.getRange(linha, COLUNAS_USUARIOS.BLOQUEADO_ATE    + 1).setValue('');

        // Limpar brute-force nas properties
        var props = PropertiesService.getScriptProperties();
        props.deleteProperty(PFX_BF + _encodarEmail(emailUsuario));

        limparCacheAba(NOME_ABA_USUARIOS);
        _registrarLog('desbloquear_usuario', sessao.email, '', 'desbloqueado: ' + emailUsuario, 'sucesso');
        return { sucesso: true, mensagem: 'Usuário desbloqueado!' };
      }
    }
    return { sucesso: false, mensagem: 'Usuário não encontrado.' };
  } catch (e) {
    Logger.log('ERRO desbloquearUsuario: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Reseta a senha de um usuário e envia por email. Requer admin.
 */
function resetarSenhaAdmin(token, userId) {
  try {
    var sessao = _sessaoAdmin(token);
    if (!sessao) return { sucesso: false, mensagem: 'Acesso negado.' };

    var aba = obterAba(NOME_ABA_USUARIOS);
    var rows = aba.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][COLUNAS_USUARIOS.ID] === userId) {
        var emailUsuario = rows[i][COLUNAS_USUARIOS.EMAIL];
        var nomeUsuario  = rows[i][COLUNAS_USUARIOS.NOME];
        var linha = i + 1;

        // Marca como primeiro acesso — usuário definirá novo PIN ao logar
        aba.getRange(linha, COLUNAS_USUARIOS.SENHA_HASH + 1).setValue('PRIMEIRO_ACESSO');
        aba.getRange(linha, COLUNAS_USUARIOS.SALT       + 1).setValue('');
        limparCacheAba(NOME_ABA_USUARIOS);

        try {
          MailApp.sendEmail({
            to: emailUsuario,
            subject: 'Smart Meeting — Redefinição de Acesso',
            htmlBody: _htmlResetAdmin(nomeUsuario, emailUsuario),
            name: EMAIL_REMETENTE_NOME,
            replyTo: EMAIL_REPLY_TO
          });
        } catch (mailErr) {
          Logger.log('Erro ao enviar email reset: ' + mailErr.toString());
        }

        _registrarLog('reset_senha', sessao.email, '', 'acesso resetado para: ' + emailUsuario, 'sucesso');
        return { sucesso: true, mensagem: 'Acesso resetado! Usuário deverá criar novo PIN no próximo login.' };
      }
    }
    return { sucesso: false, mensagem: 'Usuário não encontrado.' };
  } catch (e) {
    Logger.log('ERRO resetarSenhaAdmin: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

/**
 * Retorna logs de auditoria com filtros opcionais. Requer admin.
 * @param {string} token
 * @param {Object} filtros {evento, usuario, limite}
 */
function obterLogs(token, filtros) {
  try {
    var sessao = _sessaoAdmin(token);
    if (!sessao) return { sucesso: false, mensagem: 'Acesso negado.' };

    filtros = filtros || {};
    var aba = obterAba(NOME_ABA_LOGS);
    var dados = aba.getDataRange().getValues();
    if (!dados || dados.length <= 1) return { sucesso: true, logs: [] };

    var logs = [];
    var limite = filtros.limite || 200;

    for (var i = dados.length - 1; i >= 1 && logs.length < limite; i--) {
      if (!dados[i][COLUNAS_LOGS.TIMESTAMP]) continue;
      if (filtros.evento  && dados[i][COLUNAS_LOGS.EVENTO]   !== filtros.evento)  continue;
      if (filtros.usuario && !dados[i][COLUNAS_LOGS.USUARIO].includes(filtros.usuario)) continue;

      logs.push({
        timestamp: new Date(dados[i][COLUNAS_LOGS.TIMESTAMP]).toLocaleString('pt-BR'),
        evento:    dados[i][COLUNAS_LOGS.EVENTO],
        usuario:   dados[i][COLUNAS_LOGS.USUARIO],
        detalhes:  dados[i][COLUNAS_LOGS.DETALHES],
        resultado: dados[i][COLUNAS_LOGS.RESULTADO]
      });
    }
    return { sucesso: true, logs: logs };
  } catch (e) {
    Logger.log('ERRO obterLogs: ' + e.toString());
    return { sucesso: false, mensagem: e.message };
  }
}

function setupAdminInicial() {
  repararHeaderUsuarios();
  inicializarAdmin('italotrabalhodados@gmail.com', '5678', 'Italo');
}

/**
 * Verifica se a aba Usuarios possui cabeçalho na linha 1.
 * Se não possuir (primeira célula não é 'ID'), insere o cabeçalho.
 * EXECUTAR no editor caso o admin já tenha sido criado sem header.
 */
function repararHeaderUsuarios() {
  var aba = obterAba(NOME_ABA_USUARIOS);
  var dados = aba.getDataRange().getValues();

  // Se a primeira linha já é o cabeçalho correto, não faz nada
  if (dados.length > 0 && dados[0][0] === 'ID') {
    Logger.log('Header da aba Usuarios já está correto.');
    return;
  }

  // Insere uma linha no topo com o cabeçalho
  aba.insertRowBefore(1);
  aba.getRange(1, 1, 1, 11).setValues([[
    'ID', 'Email', 'SenhaHash', 'Salt', 'Nome',
    'Perfil', 'Ativo', 'CriadoEm', 'UltimoLogin',
    'TentativasLogin', 'BloqueadoAte'
  ]]);
  aba.getRange(1, 1, 1, 11).setFontWeight('bold');
  limparCacheAba(NOME_ABA_USUARIOS);
  Logger.log('✅ Header da aba Usuarios inserido com sucesso.');
}


// ============================================================================
//  SETUP — executar uma vez no editor do Apps Script
// ============================================================================

/**
 * Cria o primeiro administrador do sistema.
 * EXECUTAR APENAS UMA VEZ pelo editor do Apps Script.
 * Exemplo: inicializarAdmin('admin@empresa.com', 'MinhaSenh@123', 'Administrador')
 */
function inicializarAdmin(email, senha, nome) {
  if (!email || !senha || !nome) {
    Logger.log('ERRO: Forneça email, senha e nome. Ex: inicializarAdmin("admin@empresa.com", "Senha@123", "Admin")');
    return;
  }

  var emailLower = email.toString().trim().toLowerCase();
  var aba = obterAba(NOME_ABA_USUARIOS);
  var dados = aba.getDataRange().getValues();

  // Verificar se já existe algum admin
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][COLUNAS_USUARIOS.PERFIL] === 'admin') {
      Logger.log('Já existe um admin cadastrado: ' + dados[i][COLUNAS_USUARIOS.EMAIL]);
      return;
    }
  }

  var validacao = _validarForcaSenha(senha.toString());
  if (!validacao.valida) {
    Logger.log('Senha inválida: ' + validacao.mensagem);
    return;
  }

  var salt = _gerarSalt();
  var hash = _hashSenha(senha.toString(), salt);
  var id   = gerarId();

  aba.appendRow([id, emailLower, hash, salt, nome.toString().trim(), 'admin', true, new Date(), '', 0, '']);
  limparCacheAba(NOME_ABA_USUARIOS);
  _inicializarCabecalhoLogs();

  Logger.log('✅ Admin criado com sucesso: ' + emailLower);
}

// ============================================================================
//  FUNÇÕES INTERNAS (prefixo _)
// ============================================================================

function _hashSenha(senha, salt) {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    senha + salt,
    Utilities.Charset.UTF_8
  );
  return raw.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

function _gerarSalt() {
  return Utilities.getUuid().replace(/-/g, '');
}

function _criarToken() {
  return Utilities.getUuid();
}

function _encodarEmail(email) {
  return Utilities.base64Encode(email.toLowerCase()).replace(/=/g, '');
}

function _salvarSessao(token, email, perfil, nome, departamentosIds, paginasPermitidas) {
  var deps = [];
  if (departamentosIds) {
    if (Array.isArray(departamentosIds)) {
      deps = departamentosIds;
    } else {
      deps = departamentosIds.toString().split(',').map(function(d) { return d.trim(); }).filter(function(d) { return d !== ''; });
    }
  }
  var pags = [];
  if (paginasPermitidas) {
    if (Array.isArray(paginasPermitidas)) {
      pags = paginasPermitidas;
    } else {
      pags = paginasPermitidas.toString().split(',').map(function(p) { return p.trim(); }).filter(function(p) { return p !== ''; });
    }
  }
  var dados = JSON.stringify({
    email:             email,
    perfil:            perfil,
    nome:              nome || email,
    departamentosIds:  deps,
    paginasPermitidas: pags,
    expiry:            Date.now() + SESSAO_TTL_MS,
    criadoEm:          Date.now()
  });
  PropertiesService.getScriptProperties().setProperty(PFX_SESSAO + token, dados);
}

function _obterSessao(token) {
  if (!token) return null;
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PFX_SESSAO + token);
    if (!raw) return null;
    var sessao = JSON.parse(raw);
    if (Date.now() > sessao.expiry) {
      PropertiesService.getScriptProperties().deleteProperty(PFX_SESSAO + token);
      return null;
    }
    return sessao;
  } catch (e) {
    return null;
  }
}

function _renovarSessao(token, sessao) {
  try {
    sessao.expiry = Date.now() + SESSAO_TTL_MS;
    PropertiesService.getScriptProperties().setProperty(PFX_SESSAO + token, JSON.stringify(sessao));
  } catch (e) {
    // falha silenciosa
  }
}

function _invalidarSessao(token) {
  PropertiesService.getScriptProperties().deleteProperty(PFX_SESSAO + token);
}

function _invalidarSessoesDoUsuario(email) {
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    var emailLower = email.toLowerCase();
    Object.keys(props).forEach(function(key) {
      if (key.indexOf(PFX_SESSAO) === 0) {
        try {
          var s = JSON.parse(props[key]);
          if (s.email === emailLower) {
            PropertiesService.getScriptProperties().deleteProperty(key);
          }
        } catch (e) {}
      }
    });
  } catch (e) {}
}

function _verificarBruteForce(email) {
  var chave = PFX_BF + _encodarEmail(email);
  var raw = PropertiesService.getScriptProperties().getProperty(chave);
  if (!raw) return { bloqueado: false };

  var bf = JSON.parse(raw);
  if (bf.bloqueadoAte && Date.now() < bf.bloqueadoAte) {
    var mins = Math.ceil((bf.bloqueadoAte - Date.now()) / 60000);
    return { bloqueado: true, mensagem: 'Conta bloqueada. Tente em ' + mins + ' minuto(s).' };
  }
  return { bloqueado: false };
}

function _obterTentativasBf(email) {
  var chave = PFX_BF + _encodarEmail(email);
  var raw = PropertiesService.getScriptProperties().getProperty(chave);
  if (!raw) return 0;
  return JSON.parse(raw).tentativas || 0;
}

function _registrarTentativaFalha(email) {
  var chave = PFX_BF + _encodarEmail(email);
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(chave);
  var bf = raw ? JSON.parse(raw) : { tentativas: 0, nivel: 0 };

  // Limpar bloqueio expirado
  if (bf.bloqueadoAte && Date.now() >= bf.bloqueadoAte) {
    bf = { tentativas: 0, nivel: bf.nivel || 0 };
  }

  bf.tentativas = (bf.tentativas || 0) + 1;

  if (bf.tentativas >= MAX_TENTATIVAS) {
    bf.nivel = (bf.nivel || 0) + 1;
    var duracoes = [BLOQUEIO_NIVEL_1_MS, BLOQUEIO_NIVEL_2_MS, BLOQUEIO_NIVEL_3_MS];
    var duracao = duracoes[Math.min(bf.nivel - 1, duracoes.length - 1)];
    bf.bloqueadoAte = Date.now() + duracao;
    bf.tentativas = 0;

    // Atualizar na planilha também
    _marcarBloqueioNaPlanilha(email, new Date(bf.bloqueadoAte));
  }

  props.setProperty(chave, JSON.stringify(bf));
}

function _zerarTentativas(email) {
  var chave = PFX_BF + _encodarEmail(email);
  PropertiesService.getScriptProperties().deleteProperty(chave);
  _marcarBloqueioNaPlanilha(email, '');
}

function _marcarBloqueioNaPlanilha(email, bloqueadoAte) {
  try {
    var aba = obterAba(NOME_ABA_USUARIOS);
    var rows = aba.getDataRange().getValues();
    var emailLower = email.toLowerCase();
    for (var i = 1; i < rows.length; i++) {
      if ((rows[i][COLUNAS_USUARIOS.EMAIL] || '').toLowerCase() === emailLower) {
        aba.getRange(i + 1, COLUNAS_USUARIOS.TENTATIVAS_LOGIN + 1).setValue(bloqueadoAte ? MAX_TENTATIVAS : 0);
        aba.getRange(i + 1, COLUNAS_USUARIOS.BLOQUEADO_ATE    + 1).setValue(bloqueadoAte || '');
        limparCacheAba(NOME_ABA_USUARIOS);
        break;
      }
    }
  } catch (e) {}
}

function _verificarRateLimit(email) {
  var chave = PFX_RL + _encodarEmail(email);
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(chave);
  var agora = Date.now();

  if (!raw) {
    props.setProperty(chave, JSON.stringify({ janela: agora, contagem: 1 }));
    return { bloqueado: false };
  }

  var rl = JSON.parse(raw);
  if (agora - rl.janela > RATE_LIMIT_JANELA_MS) {
    props.setProperty(chave, JSON.stringify({ janela: agora, contagem: 1 }));
    return { bloqueado: false };
  }

  rl.contagem++;
  props.setProperty(chave, JSON.stringify(rl));

  return { bloqueado: rl.contagem > RATE_LIMIT_MAX };
}

function _buscarUsuarioPorEmail(email) {
  var dados = obterDadosAbaComCache(NOME_ABA_USUARIOS);
  if (!dados || dados.length <= 1) return null;
  var emailLower = email.toLowerCase();

  for (var i = 1; i < dados.length; i++) {
    if ((dados[i][COLUNAS_USUARIOS.EMAIL] || '').toLowerCase() === emailLower) {
      var rawDeps = (dados[i][COLUNAS_USUARIOS.DEPARTAMENTOS_IDS] || '').toString();
      return {
        id:               dados[i][COLUNAS_USUARIOS.ID],
        email:            dados[i][COLUNAS_USUARIOS.EMAIL],
        senhaHash:         dados[i][COLUNAS_USUARIOS.SENHA_HASH],
        salt:              dados[i][COLUNAS_USUARIOS.SALT],
        nome:              dados[i][COLUNAS_USUARIOS.NOME],
        perfil:            dados[i][COLUNAS_USUARIOS.PERFIL],
        ativo:             dados[i][COLUNAS_USUARIOS.ATIVO] === true || dados[i][COLUNAS_USUARIOS.ATIVO] === 'true',
        departamentosIds:  rawDeps ? rawDeps.split(',').map(function(d) { return d.trim(); }).filter(function(d) { return d !== ''; }) : [],
        paginasPermitidas: (function() { var raw = (dados[i][COLUNAS_USUARIOS.PAGINAS_PERMITIDAS] || '').toString().trim(); return raw ? raw.split(',').map(function(p) { return p.trim(); }).filter(function(p) { return p !== ''; }) : []; })()
      };
    }
  }
  return null;
}

function _atualizarUltimoLogin(email) {
  try {
    var aba = obterAba(NOME_ABA_USUARIOS);
    var rows = aba.getDataRange().getValues();
    var emailLower = email.toLowerCase();
    for (var i = 1; i < rows.length; i++) {
      if ((rows[i][COLUNAS_USUARIOS.EMAIL] || '').toLowerCase() === emailLower) {
        aba.getRange(i + 1, COLUNAS_USUARIOS.ULTIMO_LOGIN + 1).setValue(new Date());
        limparCacheAba(NOME_ABA_USUARIOS);
        break;
      }
    }
  } catch (e) {}
}

function _atualizarSenhaUsuario(email, hash, salt) {
  var aba = obterAba(NOME_ABA_USUARIOS);
  var rows = aba.getDataRange().getValues();
  var emailLower = email.toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][COLUNAS_USUARIOS.EMAIL] || '').toLowerCase() === emailLower) {
      aba.getRange(i + 1, COLUNAS_USUARIOS.SENHA_HASH + 1).setValue(hash);
      aba.getRange(i + 1, COLUNAS_USUARIOS.SALT       + 1).setValue(salt);
      limparCacheAba(NOME_ABA_USUARIOS);
      return;
    }
  }
}

function _sessaoAdmin(token) {
  var sessao = _obterSessao(token);
  if (!sessao || sessao.perfil !== 'admin') return null;
  return sessao;
}

function _validarForcaSenha(senha) {
  if (!/^\d{4}$/.test(senha)) {
    return { valida: false, mensagem: 'O PIN deve ter exatamente 4 dígitos numéricos.' };
  }
  return { valida: true };
}

function _gerarSenhaTemporaria() {
  var chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!@#$%';
  var senha = '';
  for (var i = 0; i < 10; i++) {
    senha += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  // Garantir requisitos mínimos
  senha = senha.substring(0, 7) + 'A1!';
  return senha;
}

function _registrarLog(evento, email, ip, detalhes, resultado) {
  try {
    var aba = obterAba(NOME_ABA_LOGS);
    aba.appendRow([new Date(), evento, email, ip, detalhes, resultado]);
  } catch (e) {
    Logger.log('ERRO _registrarLog: ' + e.toString());
  }
}

function _inicializarCabecalhoLogs() {
  try {
    var aba = obterAba(NOME_ABA_LOGS);
    if (aba.getLastRow() === 0) {
      aba.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Evento', 'Usuario', 'IP', 'Detalhes', 'Resultado']]);
      aba.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
  } catch (e) {}
}

/** Limpa sessões expiradas do PropertiesService (executar periodicamente). */
function limparSessoesExpiradas() {
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    var agora = Date.now();
    Object.keys(props).forEach(function(key) {
      if (key.indexOf(PFX_SESSAO) === 0) {
        try {
          var s = JSON.parse(props[key]);
          if (agora > s.expiry) PropertiesService.getScriptProperties().deleteProperty(key);
        } catch (e) {}
      }
    });
  } catch (e) {}
}

// ═══════════════════════════════════════════════════════════════
//  HELPERS DE EMAIL — Templates HTML
// ═══════════════════════════════════════════════════════════════

function _htmlBoasVindas(nome, email) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
'<body style="margin:0;padding:0;background:#F2E8D5;font-family:Segoe UI,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#F2E8D5;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#FFFFFF;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(30,18,10,0.13);">' +
'<tr><td style="background:linear-gradient(135deg,#1A0E06 0%,#3D1F0D 60%,#1A0E06 100%);padding:36px 40px;text-align:center;">' +
'<div style="font-size:38px;margin-bottom:10px;">&#9749;</div>' +
'<div style="font-family:Georgia,serif;font-size:30px;font-weight:700;color:#D4A96A;letter-spacing:1px;">Smart Meeting</div>' +
'<div style="font-size:11px;color:rgba(212,169,106,0.65);margin-top:6px;letter-spacing:3px;text-transform:uppercase;">Sistema de Gest&#227;o de Reuni&#245;es com IA</div>' +
'</td></tr>' +
'<tr><td style="background:linear-gradient(90deg,#C07D2A,#9B621A);padding:13px 40px;text-align:center;">' +
'<span style="font-size:14px;font-weight:700;color:#fff;letter-spacing:0.5px;">&#127881;&#160;&#160;Sua conta foi criada com sucesso!</span>' +
'</td></tr>' +
'<tr><td style="padding:38px 40px 28px;">' +
'<p style="margin:0 0 6px;font-size:24px;font-weight:700;color:#1E120A;font-family:Georgia,serif;">Ol&#225;, ' + nome + '!</p>' +
'<p style="margin:0 0 28px;font-size:14px;color:#5C3D2E;line-height:1.7;">Bem-vindo(a) ao <strong>Smart Meeting</strong>! Sua conta est&#225; pronta para uso. Siga os passos abaixo para fazer seu <strong>primeiro acesso</strong>.</p>' +
'<div style="background:#FBF6EC;border-radius:10px;padding:26px 28px;margin-bottom:28px;border:1px solid rgba(192,125,42,0.18);">' +
'<p style="margin:0 0 22px;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:1.5px;color:#C07D2A;">&#128203;&#160;&#160;Passo a passo &#8212; Primeiro Acesso</p>' +
'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:18px;"><tr>' +
'<td width="36" valign="top"><div style="width:30px;height:30px;background:#C07D2A;border-radius:50%;text-align:center;line-height:30px;font-size:13px;font-weight:800;color:#fff;">1</div></td>' +
'<td valign="middle" style="padding-left:10px;"><div style="font-size:14px;font-weight:700;color:#1E120A;margin-bottom:2px;">Clique no bot&#227;o abaixo para acessar o sistema</div><div style="font-size:12px;color:#8B6347;line-height:1.5;">O link abrir&#225; o Smart Meeting diretamente no seu navegador</div></td>' +
'</tr></table>' +
'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:18px;"><tr>' +
'<td width="36" valign="top"><div style="width:30px;height:30px;background:#C07D2A;border-radius:50%;text-align:center;line-height:30px;font-size:13px;font-weight:800;color:#fff;">2</div></td>' +
'<td valign="middle" style="padding-left:10px;"><div style="font-size:14px;font-weight:700;color:#1E120A;margin-bottom:4px;">Informe seu email de acesso</div>' +
'<div style="font-size:12px;color:#5C3D2E;background:#fff;border:1px solid rgba(192,125,42,0.3);border-radius:6px;padding:5px 12px;font-family:Courier New,monospace;display:inline-block;">' + email + '</div></td>' +
'</tr></table>' +
'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:18px;"><tr>' +
'<td width="36" valign="top"><div style="width:30px;height:30px;background:#C07D2A;border-radius:50%;text-align:center;line-height:30px;font-size:13px;font-weight:800;color:#fff;">3</div></td>' +
'<td valign="middle" style="padding-left:10px;"><div style="font-size:14px;font-weight:700;color:#1E120A;margin-bottom:2px;">Crie seu PIN de 4 d&#237;gitos</div><div style="font-size:12px;color:#8B6347;line-height:1.5;">No primeiro acesso voc&#234; definir&#225; uma senha num&#233;rica pessoal. Guarde-a em local seguro!</div></td>' +
'</tr></table>' +
'<table cellpadding="0" cellspacing="0" width="100%"><tr>' +
'<td width="36" valign="top"><div style="width:30px;height:30px;background:#4A6741;border-radius:50%;text-align:center;line-height:30px;font-size:14px;font-weight:800;color:#fff;">&#10003;</div></td>' +
'<td valign="middle" style="padding-left:10px;"><div style="font-size:14px;font-weight:700;color:#1E120A;margin-bottom:2px;">Pronto! Explore o Smart Meeting</div><div style="font-size:12px;color:#8B6347;line-height:1.5;">Grave reuni&#245;es, gere atas autom&#225;ticas com IA e gerencie projetos</div></td>' +
'</tr></table>' +
'</div>' +
'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:20px;"><tr><td align="center">' +
'<a href="' + URL_SISTEMA + '" target="_blank" style="display:inline-block;background:linear-gradient(135deg,#C07D2A 0%,#9B621A 100%);color:#ffffff;text-decoration:none;font-size:16px;font-weight:700;padding:18px 52px;border-radius:8px;letter-spacing:0.5px;box-shadow:0 4px 18px rgba(192,125,42,0.45);font-family:Segoe UI,Arial,sans-serif;">' +
'&#9749;&#160;&#160;Acessar o Smart Meeting</a>' +
'</td></tr>' +
'<tr><td align="center" style="padding-top:10px;">' +
'<span style="font-size:11px;color:#B09070;">Ou copie: <a href="' + URL_SISTEMA + '" style="color:#C07D2A;word-break:break-all;font-size:10px;">' + URL_SISTEMA + '</a></span>' +
'</td></tr></table>' +
'<div style="background:#FBF6EC;border-left:3px solid #C07D2A;padding:12px 16px;border-radius:0 6px 6px 0;">' +
'<p style="margin:0;font-size:12px;color:#5C3D2E;line-height:1.6;"><strong>&#128274; Seguran&#231;a:</strong> Nunca compartilhe seu PIN. D&#250;vidas? Contate o administrador em ' +
'<a href="mailto:' + EMAIL_REPLY_TO + '" style="color:#C07D2A;">' + EMAIL_REPLY_TO + '</a>.</p>' +
'</div>' +
'</td></tr>' +
'<tr><td style="background:#F2E8D5;border-top:1px solid rgba(192,125,42,0.15);padding:18px 40px;text-align:center;">' +
'<p style="margin:0 0 4px;font-size:12px;color:#8B6347;">Este email foi enviado automaticamente pelo <strong>Smart Meeting</strong></p>' +
'<p style="margin:0;font-size:11px;color:#B09070;">&#169; 2026 Smart Meeting &#8212; Christus &#160;|&#160; N&#227;o responda este email diretamente</p>' +
'</td></tr>' +
'</table></td></tr></table></body></html>';
}

function _htmlResetSenha(nome, linkReset) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"></head>' +
'<body style="margin:0;padding:0;background:#F2E8D5;font-family:Segoe UI,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#F2E8D5;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(30,18,10,0.13);">' +
'<tr><td style="background:linear-gradient(135deg,#1A0E06,#3D1F0D);padding:30px 40px;text-align:center;">' +
'<div style="font-size:32px;margin-bottom:8px;">&#128274;</div>' +
'<div style="font-family:Georgia,serif;font-size:24px;font-weight:700;color:#D4A96A;">Smart Meeting</div>' +
'<div style="font-size:11px;color:rgba(212,169,106,0.65);letter-spacing:2px;text-transform:uppercase;margin-top:5px;">Redefini&#231;&#227;o de Senha</div>' +
'</td></tr>' +
'<tr><td style="padding:36px 40px;">' +
'<p style="margin:0 0 6px;font-size:20px;font-weight:700;color:#1E120A;">Ol&#225;, ' + nome + '!</p>' +
'<p style="margin:0 0 24px;font-size:14px;color:#5C3D2E;line-height:1.7;">Recebemos uma solicita&#231;&#227;o para <strong>redefinir sua senha</strong>. Clique no bot&#227;o abaixo (v&#225;lido por <strong>30 minutos</strong>).</p>' +
'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:24px;"><tr><td align="center">' +
'<a href="' + linkReset + '" target="_blank" style="display:inline-block;background:linear-gradient(135deg,#C07D2A,#9B621A);color:#fff;text-decoration:none;font-size:15px;font-weight:700;padding:16px 40px;border-radius:8px;box-shadow:0 4px 14px rgba(192,125,42,0.4);">&#128273;&#160;Redefinir minha senha</a>' +
'</td></tr></table>' +
'<div style="background:#FBF6EC;border-left:3px solid #8B2E2E;padding:12px 16px;border-radius:0 6px 6px 0;">' +
'<p style="margin:0;font-size:12px;color:#5C3D2E;">&#9888;&#65039; Se voc&#234; <strong>n&#227;o solicitou</strong> a redefini&#231;&#227;o, ignore este email. Sua senha permanece a mesma.</p>' +
'</div></td></tr>' +
'<tr><td style="background:#F2E8D5;border-top:1px solid rgba(192,125,42,0.15);padding:16px 40px;text-align:center;">' +
'<p style="margin:0;font-size:11px;color:#B09070;">&#169; 2026 Smart Meeting &#8212; Christus &#160;|&#160; Email autom&#225;tico</p>' +
'</td></tr></table></td></tr></table></body></html>';
}

function _htmlResetAdmin(nome, emailUsuario) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"></head>' +
'<body style="margin:0;padding:0;background:#F2E8D5;font-family:Segoe UI,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#F2E8D5;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(30,18,10,0.13);">' +
'<tr><td style="background:linear-gradient(135deg,#1A0E06,#3D1F0D);padding:30px 40px;text-align:center;">' +
'<div style="font-size:32px;margin-bottom:8px;">&#128260;</div>' +
'<div style="font-family:Georgia,serif;font-size:24px;font-weight:700;color:#D4A96A;">Smart Meeting</div>' +
'<div style="font-size:11px;color:rgba(212,169,106,0.65);letter-spacing:2px;text-transform:uppercase;margin-top:5px;">Redefini&#231;&#227;o de Acesso</div>' +
'</td></tr>' +
'<tr><td style="padding:36px 40px;">' +
'<p style="margin:0 0 6px;font-size:20px;font-weight:700;color:#1E120A;">Ol&#225;, ' + nome + '!</p>' +
'<p style="margin:0 0 24px;font-size:14px;color:#5C3D2E;line-height:1.7;">O <strong>administrador do sistema</strong> resetou seu acesso. No pr&#243;ximo login, voc&#234; dever&#225; criar um <strong>novo PIN de 4 d&#237;gitos</strong>.</p>' +
'<div style="background:#FBF6EC;border-radius:8px;padding:20px 24px;margin-bottom:24px;border:1px solid rgba(192,125,42,0.15);">' +
'<p style="margin:0 0 8px;font-size:11px;font-weight:700;color:#C07D2A;text-transform:uppercase;letter-spacing:0.8px;">Seu email de acesso</p>' +
'<div style="font-size:14px;color:#1E120A;font-family:Courier New,monospace;background:#fff;border:1px solid rgba(192,125,42,0.25);border-radius:6px;padding:8px 14px;display:inline-block;">' + emailUsuario + '</div>' +
'</div>' +
'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:24px;"><tr><td align="center">' +
'<a href="' + URL_SISTEMA + '" target="_blank" style="display:inline-block;background:linear-gradient(135deg,#C07D2A,#9B621A);color:#fff;text-decoration:none;font-size:15px;font-weight:700;padding:16px 40px;border-radius:8px;box-shadow:0 4px 14px rgba(192,125,42,0.4);">&#9749;&#160;Acessar o Smart Meeting</a>' +
'</td></tr></table>' +
'<div style="background:#FBF6EC;border-left:3px solid #C07D2A;padding:12px 16px;border-radius:0 6px 6px 0;">' +
'<p style="margin:0;font-size:12px;color:#5C3D2E;">D&#250;vidas? <a href="mailto:' + EMAIL_REPLY_TO + '" style="color:#C07D2A;">' + EMAIL_REPLY_TO + '</a></p>' +
'</div></td></tr>' +
'<tr><td style="background:#F2E8D5;border-top:1px solid rgba(192,125,42,0.15);padding:16px 40px;text-align:center;">' +
'<p style="margin:0;font-size:11px;color:#B09070;">&#169; 2026 Smart Meeting &#8212; Christus &#160;|&#160; Email autom&#225;tico</p>' +
'</td></tr></table></td></tr></table></body></html>';
}

