function enviarEmailNovosMobilizadores() {
  const nomeAba = 'Mobilizador';
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  if (!planilha) {
    Logger.log("Aba não encontrada: " + nomeAba);
    return;
  }

  // Pegar a última linha e coluna da planilha
  const ultimaLinha = planilha.getLastRow();
  const ultimaColuna = planilha.getLastColumn();

  // Verificar se há dados após a segunda linha
  if (ultimaLinha <= 2) {
    Logger.log("Nenhum dado encontrado para enviar.");
    return;
  }

  // Obter todos os dados a partir da terceira linha
  const dadosAtuais = planilha.getRange(3, 1, ultimaLinha - 2, ultimaColuna).getValues();

  // Carregar o número de linhas da última execução
  const ultimaLinhaSalva = PropertiesService.getScriptProperties().getProperty('ultimaLinhaSalva') || 2;
  Logger.log("Última linha salva na execução anterior: " + ultimaLinhaSalva);
  Logger.log("Última linha atual: " + ultimaLinha);

  // Ajustar para lidar com exclusão de linhas
  let novasLinhas = [];

  if (ultimaLinha > ultimaLinhaSalva) {
    // Capturar as novas linhas que foram adicionadas
    novasLinhas = planilha.getRange(Number(ultimaLinhaSalva) + 1, 1, ultimaLinha - ultimaLinhaSalva, ultimaColuna).getValues();
  } else if (ultimaLinha < ultimaLinhaSalva) {
    // Caso tenha ocorrido exclusão de linhas, considerar todas as linhas como novas
    novasLinhas = planilha.getRange(3, 1, ultimaLinha - 2, ultimaColuna).getValues();
    Logger.log("Linhas anteriores foram excluídas. Considerando todas as linhas atuais como novas.");
  }

  // Verificar se há novas linhas para enviar
  if (novasLinhas.length === 0) {
    Logger.log("Nenhuma nova linha encontrada.");
    return;
  }

  // Montar tabela em HTML com as novas linhas
  const htmlTable = montarTabelaHTML(novasLinhas);

  // Montar o corpo do e-mail
  const corpoEmail = `
    <p>Segue abaixo os novos dados inseridos:</p>
    <table border="1" style="border-collapse: collapse; width: 100%;">
      <thead>
        <tr>
          <th style="text-align: center;">DATA</th>
          <th style="text-align: center;">HORÁRIO</th>
          <th style="text-align: center;">TIPO DE AGENDAMENTO</th>
          <th style="text-align: center;">IGREJA/LOCAL</th>
          <th style="text-align: center;">ENDEREÇO</th>
          <th style="text-align: center;">BAIRRO</th>
          <th style="text-align: center;">CIDADE</th>
          <th style="text-align: center;">UF</th>
          <th style="text-align: center;">AGENDAMENTO</th>
          <th style="text-align: center;">RESPONSÁVEL</th>
          <th style="text-align: center;">TELEFONE</th>
          <th style="text-align: center;">E-MAIL</th>
          <th style="text-align: center;">NOME DO PASTOR</th>
          <th style="text-align: center;">STATUS</th>
        </tr>
      </thead>
      <tbody>${htmlTable}</tbody>
    </table>
    <br>
    --<br>
    A.S.S<br>
    Automatic Scheduling System<br>
    MIT License<br>
    Copyright (c) 2024 Diego Ferreira L.G. Oliveira
  `;

  // Enviar o e-mail
  const destinatario = 'john.doe@jmm.org.br';
  const assunto = 'Relatório Semanal | Agendamentos dos Mobilizadores | Teste';

  MailApp.sendEmail({
    to: destinatario,
    subject: assunto,
    htmlBody: corpoEmail
  });

  Logger.log("E-mail enviado com sucesso.");

  // Atualizar a última linha salva
  PropertiesService.getScriptProperties().setProperty('ultimaLinhaSalva', ultimaLinha.toString());
}

// Função para montar a tabela HTML
function montarTabelaHTML(dados) {
  return dados.map(linha => {
    // Formatar a data apenas se a célula não estiver vazia
    if (linha[0]) {
      linha[0] = Utilities.formatDate(new Date(linha[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return `<tr>${linha.map(celula => `<td style="text-align: center;">${celula}</td>`).join('')}</tr>`;
  }).join('');
}
