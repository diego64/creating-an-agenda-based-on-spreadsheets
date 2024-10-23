function mostrarModificacoesPlanilha() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abas = planilha.getSheets();
  var horarioEdicao = new Date(); // Captura a data e hora da execução da função
  var horarioFormatado = Utilities.formatDate(horarioEdicao, Session.getScriptTimeZone(), "HH:mm"); // Formata o horário

  // Ler a partir da segunda aba em diante
  for (var i = 1; i < abas.length; i++) {
    var aba = abas[i];
    var dados = aba.getDataRange().getValues(); // Obtém todos os dados da aba

    // Verificar se há dados nas linhas a começa a leitura
    if (dados.length > 2) {
      for (var j = 2; j < dados.length; j++) { // Começando na terceira linha
        var linha = dados[j];

        // Determina se usar "RESPONSÁVEL" ou "MISSONARIO"
        var responsavelOuMissionario = (i === 1) ? linha[9] : linha[10]; // Coluna 10 para RESPONSÁVEL, 11 para MISSONARIO

        // Montagem das informações
        var mensagem = `ABA ${aba.getName()} - Foi Inserido/Alterado um dado na [linha ${j + 1}] às [${horarioFormatado}] com o seguinte dado: Agendamento ${linha[2]} na ${linha[3]} às ${linha[1]} para ${responsavelOuMissionario} e está ${linha[13]}`;

        // Log das informações
        Logger.log(mensagem);
      }
    }
  }

  // Se não houver alterações, não haverá log então, exibe essa mensagem
  Logger.log("Verificações de modificação concluídas.");
}
