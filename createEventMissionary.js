function criarEventoMissionario() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var todasAbas = planilha.getSheets();
  
  // Obtém a agenda específica pelo nome
  var agendaNome = "Agendamento Missionário";
  var agendas = CalendarApp.getCalendarsByName(agendaNome);
  var agenda = agendas.length > 0 ? agendas[0] : null;
  
  if (!agenda) {
    Logger.log("A agenda '" + agendaNome + "' não foi encontrada.");
    return;
  }

  // Ler da 3ª aba em diante
  for (var i = 2; i < todasAbas.length; i++) {
    var aba = todasAbas[i];
    var valores = aba.getDataRange().getValues();
    
    // Obtém os títulos da segunda linha
    var titulos = valores[1];
    
    Logger.log("Dados da aba " + aba.getName() + ": " + JSON.stringify(valores));
    
    // Percorre os dados a partir da 3ª linha
    for (var j = 2; j < valores.length; j++) {
      var linha = valores[j];
      
      // Verifica se a linha tem valores para evitar erros
      if (linha[0] && linha[1]) {
        var data = linha[0]; // DATA
        var horario = linha[1]; // HORÁRIO
        var igrejaLocal = linha[2]; // IGREJA/LOCAL
        var endereco = linha[3]; // ENDEREÇO
        var bairro = linha[4]; // BAIRRO
        var cidade = linha[5]; // CIDADE
        var uf = linha[6]; // UF
        var agendamento = linha[7]; // AGENDAMENTO
        var missionario = linha[8]; // MISSIONÁRIO
        var tipo_de_agendamento = linha[9]; // TIPO DE AGENDAMENTO
        var responsavel = linha[10]; // RESPONSÁVEL PELO AGENDAMENTO
        var telefone = linha[11]; // TELEFONE
        var email = linha[12]; // E-MAIL
        var status = linha[13]; // STATUS
        
        // Log dos valores da linha para depuração
        Logger.log("Data: " + data + ", Horário: " + horario + ", Local: " + igrejaLocal);
        
        // Cria o título do evento
        var tituloEvento = agendamento + " | " + missionario;

        // Verifica se o horário está no formato correto
        if (typeof horario === "string" && horario.includes(":")) {
          var partesHorario = horario.split(":"); // Divide a string de horário em horas e minutos
          
          // Verifica se tem duas partes
          if (partesHorario.length === 2) {
            // Converte horas e minutos em números inteiros
            var horas = parseInt(partesHorario[0]);
            var minutos = parseInt(partesHorario[1]);
            
            // Formata o horário como string no formato "HH:mm"
            var horarioString = (horas < 10 ? "0" : "") + horas + ":" + (minutos < 10 ? "0" : "") + minutos;

            // Formata a data e o horário para criar um objeto Date
            var dataEvento = new Date(data);
            dataEvento.setHours(horas, minutos); // Ajusta horas e minutos

            // Verifica se a data do evento é anterior à data atual
            var agora = new Date();
            if (dataEvento < agora) {
              Logger.log("Erro: Tentativa de criar evento em uma data que já passou: " + dataEvento);
              continue;
            }

            // Define a duração do evento (30 minutos)
            var fimEvento = new Date(dataEvento);
            fimEvento.setMinutes(fimEvento.getMinutes() + 30);

            // Verifica se já existe um evento com o mesmo título e no mesmo horário
            var eventosExistentes = agenda.getEvents(dataEvento, fimEvento);
            var eventoExistente = eventosExistentes.find(evento => evento.getTitle() === tituloEvento);

            if (eventoExistente) {
              Logger.log("Evento já existe: " + tituloEvento + " no horário " + horarioString + " em " + aba.getName());
            } else {
              // Criação do evento na agenda "Agendamento Missionário"
              var criarEvento = agenda.createEvent(tituloEvento, dataEvento, fimEvento, {
                location: igrejaLocal,
                description: `Telefone: ${telefone}\nHorário: ${horarioString}\nMobilizador responsavél: ${responsavel}`
              });

              Logger.log("Evento criado: " + criarEvento.getTitle() + " no horário " + horarioString + " em " + aba.getName());
            }
          } else {
            Logger.log("Erro: O horário não está no formato 'HH:mm': " + horario);
          }
        } else {
          Logger.log("Erro: O valor de horário é inválido ou não está em formato de string: " + horario);
        }
      }
    }
  }
}
