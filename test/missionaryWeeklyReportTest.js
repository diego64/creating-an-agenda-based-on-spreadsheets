function relatorioSemanalMissionario() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const totalAbas = planilha.getNumSheets();
  
  let todosDadosAcumulados = [];
  let abasComNovosDados = [];

  // Armazena os dados anteriores
  let dadosAnteriores = {};

  for (let i = 2; i < totalAbas; i++) { // Lê da terceira aba em diante (índice 2)
    const aba = planilha.getSheets()[i];
    const nomeAba = aba.getName();
    
    Logger.log("Lendo dados da aba: " + nomeAba);

    // Obtém dados anteriores, se existirem
    dadosAnteriores[nomeAba] = JSON.parse(PropertiesService.getScriptProperties().getProperty('dadosAnteriores_' + nomeAba) || '[]');

    let ultimaLinhaProcessada = parseInt(PropertiesService.getScriptProperties().getProperty('ultimaLinha_' + nomeAba) || '1');
    const novaUltimaLinha = aba.getLastRow();

    Logger.log(`Última linha processada na aba ${nomeAba}: ${ultimaLinhaProcessada}`);
    Logger.log(`Última linha atual da aba ${nomeAba}: ${novaUltimaLinha}`);

    // Se a aba está vazia, ignore-a
    if (novaUltimaLinha < 2) {
      Logger.log(`A aba ${nomeAba} está vazia. Ignorando.`);
      continue;
    }

    // Verifica se a última linha processada é menor que a nova última linha
    if (ultimaLinhaProcessada < novaUltimaLinha) {
      const numLinhasNovas = novaUltimaLinha - ultimaLinhaProcessada;
      Logger.log(`Novas linhas detectadas na aba ${nomeAba}: ${numLinhasNovas}`);

      // Coleta as novas linhas
      const dadosNovos = aba.getRange(ultimaLinhaProcessada + 1, 1, numLinhasNovas, aba.getLastColumn()).getValues();

      // Verifica dados novos
      Logger.log("Dados novos detectados: " + JSON.stringify(dadosNovos));

      // Filtra dados que não foram excluídos
      const dadosFiltrados = dadosNovos.filter(dado => {
        return !dadosAnteriores[nomeAba].some(dadoAnterior => JSON.stringify(dado) === JSON.stringify(dadoAnterior));
      });

      Logger.log("Dados filtrados (não excluídos): " + JSON.stringify(dadosFiltrados));

      if (dadosFiltrados.length > 0) {
        todosDadosAcumulados = todosDadosAcumulados.concat(dadosFiltrados);
        abasComNovosDados.push(nomeAba); // Adiciona o nome da aba à lista

        // Atualiza a última linha processada
        PropertiesService.getScriptProperties().setProperty('ultimaLinha_' + nomeAba, novaUltimaLinha.toString());
        Logger.log(`Dados acumulados até agora na aba ${nomeAba}: ${dadosFiltrados.length} linhas.`);
      } else {
        Logger.log(`Nenhuma nova linha foi adicionada na aba ${nomeAba} desde a última execução.`);
      }
    } else if (ultimaLinhaProcessada > novaUltimaLinha) {
      // Se a última linha processada é maior que a nova última linha, trata exclusões
      Logger.log(`A aba ${nomeAba} teve dados excluídos. Processando dados restantes...`);

      // Coleta todas as linhas até a nova última linha
      const dadosRestantes = aba.getRange(2, 1, novaUltimaLinha - 1, aba.getLastColumn()).getValues();

      // Filtra dados que não foram excluídos
      const dadosFiltrados = dadosRestantes.filter(dado => {
        return !dadosAnteriores[nomeAba].some(dadoAnterior => JSON.stringify(dado) === JSON.stringify(dadoAnterior));
      });

      if (dadosFiltrados.length > 0) {
        todosDadosAcumulados = todosDadosAcumulados.concat(dadosFiltrados);
        abasComNovosDados.push(nomeAba); // Adiciona o nome da aba à lista
        Logger.log(`Dados restantes acumulados na aba ${nomeAba}: ${dadosFiltrados.length} linhas.`);
      } else {
        Logger.log(`Nenhuma nova linha encontrada na aba ${nomeAba} após exclusões.`);
      }
    } else {
      Logger.log(`Nenhuma nova linha foi adicionada na aba ${nomeAba} desde a última execução.`);
    }

    // Atualiza os dados anteriores
    PropertiesService.getScriptProperties().setProperty('dadosAnteriores_' + nomeAba, JSON.stringify(aba.getRange(2, 1, novaUltimaLinha - 1, aba.getLastColumn()).getValues()));
  }

  // Verifica se há dados acumulados antes de enviar o e-mail
  if (todosDadosAcumulados.length > 0) {
    const htmlTable = montarTabela(todosDadosAcumulados);

    // Definindo o cabeçalho da tabela
    const cabecalho = ` 
      <tr>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">DATA</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">HORÁRIO</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">TIPO DE AGENDAMENTO</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">IGREJA/LOCAL</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">ENDEREÇO</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">BAIRRO</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">CIDADE</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">UF</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">AGENDAMENTO</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">RESPONSÁVEL</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">TELEFONE</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">E-MAIL</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">NOME DO PASTOR</th>
        <th style="text-align: center; padding: 8px; border: 1px solid black;">STATUS</th>
      </tr>
    `;

    const abasNovosDadosTexto = abasComNovosDados.length > 0 ? `Novos dados nas abas: ${abasComNovosDados.join(', ')}` : 'Nenhum dado novo';
    const destinatario = 'john.doe@jmm.org.br';
    const assunto = `Relatório Semanal | ${abasNovosDadosTexto} | Teste`;
    const corpoEmail = `
      <p>Segue abaixo os novos dados nas abas dos missionários:</p>
      <table border="1" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
        <thead>${cabecalho}</thead>
        <tbody>${htmlTable}</tbody>
      </table>
      <br>
      --<br>
      A.S.S<br>
      Automatic Scheduling System<br>
      MIT License<br>
      Copyright (c) 2024 Diego Ferreira L.G. Oliveira
    `;

    MailApp.sendEmail({
      to: destinatario,
      subject: assunto,
      htmlBody: corpoEmail
    });

    Logger.log("E-mail enviado com sucesso.");
  } else {
    Logger.log("Nenhum dado novo para enviar.");
  }
}

function montarTabela(dadosSemana) {
  return dadosSemana.map(linha => {
    linha[0] = Utilities.formatDate(new Date(linha[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    return `
      <tr>${linha.slice(0, 14).map(celula => `<td style="text-align: center; padding: 8px; border: 1px solid black;">${celula}</td>`).join('')}</tr>
    `;
  }).join('');
}
