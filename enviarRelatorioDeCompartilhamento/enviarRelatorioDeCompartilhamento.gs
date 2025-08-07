// =================================================================
// INÍCIO DO CÓDIGO
// =================================================================

/**
 * !!! INSTRUÇÕES DE USO !!!
 * 1. Configure o arquivo 'appsscript.json' e habilite a 'Drive API' conforme as instruções.
 * 2. Execute a função 'gerenciarCicloDeRelatorio' MANUALMENTE APENAS UMA VEZ para iniciar o processo.
 * 3. O script irá se autogerenciar a partir daí, criando e excluindo acionadores conforme necessário.
 *
 * Para reiniciar todo o processo do zero, execute a função 'reiniciarCicloManualmente'.
 * Para ver o que está salvo na memória, execute a função 'verificarMemoriaDoCiclo'.
 */

// =================================================================
// FUNÇÃO PRINCIPAL - A ser executada para iniciar o ciclo
// =================================================================
function gerenciarCicloDeRelatorio() {
  const properties = PropertiesService.getUserProperties();
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MINUTES = 4.5;

  let itemsToProcessIds = JSON.parse(properties.getProperty('itemsToProcessIds') || 'null');
  let masterFileCount = parseInt(properties.getProperty('masterFileCount') || '0');
  let partialReportHTML = properties.getProperty('partialReportHTML') || '';
  
  const emailUsuario = Session.getEffectiveUser().getEmail();

  if (!itemsToProcessIds) {
    Logger.log("INICIANDO NOVO CICLO COMPLETO.");
    criarAcionadorDeCiclo(5); // Cria o acionador de 5 minutos para continuar a execução.
    
    const allFiles = getAllFileIdsFromDrive();
    if (allFiles === null) {
      Logger.log("Não foi possível buscar a lista de arquivos. Abortando.");
      return;
    }

    itemsToProcessIds = allFiles.map(f => f.id);
    masterFileCount = itemsToProcessIds.length;

    properties.setProperty('itemsToProcessIds', JSON.stringify(itemsToProcessIds));
    properties.setProperty('masterFileCount', masterFileCount.toString());
    properties.deleteProperty('partialReportHTML');
    partialReportHTML = '';
    Logger.log(`Novo ciclo iniciado com ${masterFileCount} itens para verificar.`);
  }

  Logger.log(`${itemsToProcessIds.length} itens restantes para processar neste ciclo.`);
  let newFindingsHTML = '';

  while (itemsToProcessIds.length > 0) {
    const elapsedMinutes = (new Date().getTime() - startTime) / 1000 / 60;
    if (elapsedMinutes > MAX_RUNTIME_MINUTES) {
      Logger.log(`Atingido o limite de tempo. Pausando a execução. O acionador de 5 minutos continuará o processo.`);
      break;
    }

    const itemId = itemsToProcessIds.shift();
    const itemReportRow = processSingleItem(itemId, emailUsuario);
    if (itemReportRow) {
      newFindingsHTML += itemReportRow;
    }
  }

  properties.setProperty('itemsToProcessIds', JSON.stringify(itemsToProcessIds));
  if (newFindingsHTML) {
      properties.setProperty('partialReportHTML', partialReportHTML + newFindingsHTML);
  }

  const itemsRemaining = itemsToProcessIds.length;
  Logger.log(`Execução parcial concluída. ${itemsRemaining} itens restantes. Estado salvo.`);

  if (itemsRemaining === 0 && masterFileCount > 0) {
    Logger.log("CICLO CONCLUÍDO! Enviando relatório final como anexo e agendando o próximo ciclo.");
    
    const finalReportHTML = properties.getProperty('partialReportHTML');
    enviarRelatorioFinalAnexado(emailUsuario, masterFileCount, finalReportHTML);
    
    agendarProximoCicloCompleto(7); // Agenda o próximo ciclo para daqui a 7 dias.
    limparPropriedadesDoCiclo();
  }
}

// =================================================================
// FUNÇÕES DE GERENCIAMENTO DE CICLO E ACIONADORES
// =================================================================

/**
 * Limpa todos os acionadores do projeto e as propriedades salvas.
 * Execute para forçar um reinício completo.
 */
function reiniciarCicloManualmente() {
  Logger.log("Limpando todos os acionadores e propriedades para um reinício completo...");
  excluirAcionadoresExistentes();
  limparPropriedadesDoCiclo();
  Logger.log("Limpeza concluída. Execute 'gerenciarCicloDeRelatorio' para iniciar um novo ciclo.");
}

/**
 * Mostra o conteúdo salvo na memória do script.
 * Execute esta função para depurar e ver o estado atual do ciclo.
 */
function verificarMemoriaDoCiclo() {
  Logger.log("--- VERIFICANDO MEMÓRIA SALVA (PropertiesService) ---");
  const properties = PropertiesService.getUserProperties();
  const allProperties = properties.getProperties();
  
  const keys = Object.keys(allProperties);
  if (keys.length === 0) {
    Logger.log("A memória do script está vazia. Nenhum ciclo em andamento.");
    return;
  }

  for (const key of keys) {
    let value = allProperties[key];
    Logger.log(`CHAVE: "${key}"`);
    
    if (key === 'itemsToProcessIds') {
      try {
        const arrayValue = JSON.parse(value);
        Logger.log(`  - TIPO: Lista de IDs (JSON)`);
        Logger.log(`  - ITENS RESTANTES: ${arrayValue.length}`);
        Logger.log(`  - PRIMEIROS 5 IDs: [${arrayValue.slice(0, 5).join(', ')}]`);
      } catch (e) {
        Logger.log(`  - VALOR (não é um JSON válido): ${value}`);
      }
    } else {
      Logger.log(`  - VALOR (visão parcial): ${value.substring(0, 200)}...`);
    }
    Logger.log("-----------------------------------------------------");
  }
}


/**
 * Cria um acionador para executar o script a cada X minutos.
 * @param {number} minutos O intervalo em minutos.
 */
function criarAcionadorDeCiclo(minutos) {
  excluirAcionadoresExistentes();
  ScriptApp.newTrigger('gerenciarCicloDeRelatorio')
      .timeBased()
      .everyMinutes(minutos)
      .create();
  Logger.log(`Acionador criado para executar a cada ${minutos} minutos.`);
}

/**
 * Exclui o acionador de ciclo e cria um novo para daqui a X dias.
 * @param {number} dias A quantidade de dias para a próxima execução.
 */
function agendarProximoCicloCompleto(dias) {
  excluirAcionadoresExistentes();
  const proximaData = new Date();
  proximaData.setDate(proximaData.getDate() + dias);
  ScriptApp.newTrigger('gerenciarCicloDeRelatorio')
      .timeBased()
      .at(proximaData)
      .create();
  Logger.log(`Ciclo concluído. Próxima verificação agendada para: ${proximaData.toLocaleString('pt-BR')}`);
}

/**
 * Exclui todos os acionadores associados a este projeto para evitar duplicatas.
 */
function excluirAcionadoresExistentes() {
  const acionadores = ScriptApp.getProjectTriggers();
  for (const acionador of acionadores) {
    if (acionador.getHandlerFunction() === 'gerenciarCicloDeRelatorio') {
      ScriptApp.deleteTrigger(acionador);
    }
  }
  Logger.log("Acionadores existentes foram limpos.");
}

/**
 * Limpa as propriedades do usuário relacionadas ao ciclo.
 */
function limparPropriedadesDoCiclo() {
  const properties = PropertiesService.getUserProperties();
  properties.deleteProperty('itemsToProcessIds');
  properties.deleteProperty('masterFileCount');
  properties.deleteProperty('partialReportHTML');
  Logger.log("Propriedades do ciclo foram limpas.");
}


// =================================================================
// FUNÇÕES AUXILIARES
// =================================================================

/**
 * Envia o e-mail final com o relatório completo como um anexo ZIP.
 */
function enviarRelatorioFinalAnexado(emailUsuario, masterFileCount, finalReportHTML) {
  const nomeArquivoHTML = 'relatorio_compartilhamento.html';
  const nomeArquivoZIP = 'relatorio_compartilhamento.zip';

  const totalDeItensCompartilhados = (finalReportHTML.match(/<tr/g) || []).length;

  const htmlCompletoDoArquivo = `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Relatório de Compartilhamento do Google Drive</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; }
          table { border-collapse: collapse; width: 100%; }
          th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }
          th { background-color: #f2f2f2; }
          tr:nth-child(even) { background-color: #f9f9f9; }
          a { color: #1a73e8; text-decoration: none; }
          a:hover { text-decoration: underline; }
        </style>
      </head>
      <body>
        <h1>Relatório de Itens Compartilhados</h1>
        <p>
          <b>Total de itens no seu Drive:</b> ${masterFileCount}<br>
          <b>Total de itens compartilhados encontrados:</b> ${totalDeItensCompartilhados}
        </p>
        <table>
          <tr>
            <th>Nome do Item (Link)</th>
            <th>Caminho Completo</th>
            <th>Última Modificação (Data)</th>
            <th>Quem Tem Acesso</th>
          </tr>
          ${finalReportHTML || '<tr><td colspan="4">Nenhum item compartilhado foi encontrado.</td></tr>'}
        </table>
      </body>
    </html>`;

  try {
    const blobHTML = Utilities.newBlob(htmlCompletoDoArquivo, MimeType.HTML, nomeArquivoHTML);
    const blobZIP = Utilities.zip([blobHTML], nomeArquivoZIP);

    const assunto = `[Relatório Final] Itens Compartilhados no seu Google Drive`;
    const statusMessage = `
      <p><b>Ciclo de verificação concluído!</b></p>
      <ul>
        <li>Todos os <b>${masterFileCount}</b> itens do seu Drive foram analisados.</li>
        <li><b>${totalDeItensCompartilhados}</b> itens compartilhados foram encontrados.</li>
        <li>O relatório completo está no arquivo <code>${nomeArquivoZIP}</code> anexado.</li>
      </ul>
      <p>Uma nova verificação completa foi agendada para daqui a 7 dias.</p>`;

    const scriptId = ScriptApp.getScriptId();
    const scriptUrl = `https://script.google.com/d/${scriptId}/edit`;
    const footer = `<hr><p style="font-size:12px; color:#666; text-align:center;">E-mail gerado por: <a href="${scriptUrl}">Google Apps Script</a> (Função: <code>gerenciarCicloDeRelatorio</code>)</p>`;

    MailApp.sendEmail({
      to: emailUsuario,
      subject: assunto,
      htmlBody: `<div style="font-family: Arial, sans-serif;">${statusMessage}${footer}</div>`,
      attachments: [blobZIP],
      name: 'Relatórios Google Drive'
    });
    Logger.log("Relatório final enviado com sucesso como anexo ZIP.");

  } catch (e) {
    Logger.log(`ERRO ao criar ou enviar o anexo: ${e.message}`);
    MailApp.sendEmail(emailUsuario, "ERRO no Relatório do Drive", `Ocorreu um erro ao tentar gerar o anexo do relatório: ${e.message}`);
  }
}


/**
 * Busca todos os IDs de arquivos do usuário usando a API Avançada.
 */
function getAllFileIdsFromDrive() {
  try {
    const allFiles = [];
    let pageToken;
    const query = "'me' in owners and trashed = false";
    do {
      const response = Drive.Files.list({ q: query, maxResults: 1000, pageToken: pageToken });
      if (response.items && response.items.length > 0) {
        allFiles.push(...response.items);
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
    return allFiles;
  } catch (e) {
    Logger.log(`ERRO CRÍTICO ao buscar a lista de arquivos da API do Drive: ${e.message}`);
    return null;
  }
}

/**
 * Processa um único item do Drive e retorna uma linha de tabela HTML se for compartilhado.
 */
function processSingleItem(itemId, emailUsuario) {
  try {
    let item;
    try {
      item = DriveApp.getFileById(itemId);
    } catch (e) {
      item = DriveApp.getFolderById(itemId);
    }

    const acessos = [];
    item.getEditors().forEach(editor => {
      if (editor.getEmail() !== emailUsuario) {
        acessos.push(`${editor.getEmail()} (Editor)`);
      }
    });
    item.getViewers().forEach(viewer => {
      acessos.push(`${viewer.getEmail()} (Leitor)`);
    });

    const acessoLink = item.getSharingAccess();
    if (acessoLink === DriveApp.Access.ANYONE || acessoLink === DriveApp.Access.ANYONE_WITH_LINK) {
      acessos.push('Qualquer pessoa com o link');
    } else if (acessoLink === DriveApp.Access.DOMAIN || acessoLink === DriveApp.Access.DOMAIN_WITH_LINK) {
      acessos.push(`Qualquer pessoa no seu domínio`);
    }

    if (acessos.length > 0) {
      const url = item.getUrl();
      const nome = item.getName();
      const caminho = obterCaminhoCompleto(item);
      const dataModificacao = Utilities.formatDate(item.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      return `
        <tr>
          <td><a href="${url}">${nome}</a></td>
          <td>${caminho}</td>
          <td>${dataModificacao}</td>
          <td>${acessos.join('<br>')}</td>
        </tr>`;
    }
    return null;
  } catch (e) {
    Logger.log(`Falha ao processar item ID ${itemId} como arquivo ou pasta: ${e.message}`);
    return null;
  }
}

/**
 * Constrói o caminho completo da pasta para um determinado item.
 */
function obterCaminhoCompleto(item) {
  try {
    let pastaPaiIterator = item.getParents();
    if (!pastaPaiIterator.hasNext()) return "/ (Raiz)";
    let pastaPai = pastaPaiIterator.next();
    let caminho = pastaPai.getName();
    while (pastaPai.getParents().hasNext()) {
      pastaPai = pastaPai.getParents().next();
      caminho = `${pastaPai.getName()}/${caminho}`;
    }
    return `/${caminho}`;
  } catch (e) {
    return "Não foi possível determinar o caminho.";
  }
}

// =================================================================
// FIM DO CÓDIGO
// =================================================================
