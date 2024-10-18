function enviarEmailAlteracoes() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abas = planilha.getSheets();
  var alteracoes = []; // Array para armazenar as alterações

  // Itera sobre as abas a partir da segunda aba
  for (var i = 1; i < abas.length; i++) {
    var aba = abas[i];
    var dados = aba.getDataRange().getValues(); // Obtém todos os dados da aba

    // Verifica se há dados e começa a partir da terceira linha
    if (dados.length > 2) {
      for (var j = 2; j < dados.length; j++) { // Começando na terceira linha
        var linha = dados[j];
        
        // Determina se usar "RESPONSÁVEL" ou "MISSONARIO"
        var responsavelOuMissionario = (i === 1) ? linha[9] : linha[10]; // Coluna 10 para RESPONSÁVEL, 11 para MISSONARIO
        
        // Monta a mensagem para a alteração
        var alteracao = [
          aba.getName(),
          j + 1,
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm"),
          linha[2], // Tipo de Agendamento
          linha[3], // Igreja/Local
          linha[1], // Horário
          responsavelOuMissionario,
          linha[13]  // Status
        ];
        
        alteracoes.push(alteracao);
      }
    }
  }

  // Se houver alterações, monta a tabela e envia o e-mail
  if (alteracoes.length > 0) {
    var htmlTable = "<table border='1'><tr><th>ABA</th><th>LINHA</th><th>HORÁRIO</th><th>TIPO DE AGENDAMENTO</th><th>IGREJA/LOCAL</th><th>HORÁRIO</th><th>RESPONSÁVEL/MISSONARIO</th><th>STATUS</th></tr>";

    alteracoes.forEach(function(alteracao) {
      htmlTable += "<tr>";
      alteracao.forEach(function(dado) {
        htmlTable += "<td>" + dado + "</td>";
      });
      htmlTable += "</tr>";
    });

    htmlTable += "</table>";

    // Monta o corpo do e-mail
    var corpoEmail = `
      <p>Olá Setor de Agendamento, tudo bem?</p>
      <p>Segue o resumo de tudo que foi feito na Planilha dessa semana:</p>
      ${htmlTable}
      <p>Qualquer dúvida ou ajuste que precise, por favor procurar o analista Diego Ferreira.</p>
      <p style="font-weight:bold;">A.S.S<br>
      MIT License<br>
      Copyright (c) 2024 Diego Ferreira L.G.Oliveira<br>
    `;

    // Envia o e-mail
    MailApp.sendEmail({
      to: "john.doe@jmm.org.br", // Insira o seu e-mail
      subject: "Alterações na Planilha",
      htmlBody: corpoEmail
    });

    Logger.log("E-mail enviado com as alterações.");
  } else {
    Logger.log("Nenhuma alteração encontrada.");
  }
}
