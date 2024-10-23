// Objeto global para armazenar o estado anterior
var estadoAnterior = {};

// Função para verificar alterações nas abas e logar as modificações
function monitorarStatus() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Planilha ativa
  var sheets = spreadsheet.getSheets(); // Pega todas as abas da planilha
  
  // Itera sobre as abas a partir da segunda (índice 1)
  sheets.slice(1).forEach(function(sheet) {
    var sheetName = sheet.getName(); // Nome da aba

    var range = sheet.getDataRange(); // Pega todas as células com dados
    var values = range.getValues(); // Pega os valores atuais das células

    // Inicializa o estado anterior da aba se não existir
    if (!estadoAnterior[sheetName]) {
      estadoAnterior[sheetName] = [];
    }

    // Itera sobre as linhas da aba, começando da terceira linha (índice 2)
    for (var row = 2; row < values.length; row++) {
      var currentValue = values[row][13]; // Valor atual da coluna N (14ª coluna)
      var previousValue = estadoAnterior[sheetName][row] || "N/A"; // Valor anterior da coluna N

      // Verifica se houve alteração de status
      if (currentValue !== previousValue && currentValue !== "N/A") {
        // Obtém a data e hora atual
        var timestamp = new Date();
        var formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy");
        var formattedTime = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm:ss");

        // Loga a mensagem no formato desejado
        var logMessage = 'ABA [' + sheetName + '] foi MODIFICADO o STATUS DE [' + previousValue + '] PARA [' + currentValue + '] AS ' + formattedTime + ' em ' + formattedDate + '.';
        Logger.log(logMessage);
        console.log(logMessage);
      }

      // Atualiza o estado anterior no objeto
      estadoAnterior[sheetName][row] = currentValue;
    }
  });

  console.log('Monitoramento de todas as abas completo.');
}
