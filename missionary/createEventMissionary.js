function criarEventoMissionario() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var todasAbas = planilha.getSheets();
  
  var agendaMissionario = "Agendamento Missionário";
  var agendas = CalendarApp.getCalendarsByName(agendaMissionario);
  var agenda = agendas.length > 0 ? agendas[0] : null;

  if (!agenda) {
    Logger.log("A agenda '" + agendaMissionario + "' não foi encontrada.");
    return;
  }

  for (var i = 2; i < todasAbas.length; i++) {
    var aba = todasAbas[i];
    var valores = aba.getDataRange().getValues();
    
    Logger.log("Dados da aba " + aba.getName() + ": " + JSON.stringify(valores));
    
    for (var j = 2; j < valores.length; j++) {
      var linha = valores[j];
      
      if (linha[0] && linha[1]) {
        var data = linha[0];
        var horario = linha[1];
        var tipo_de_agendamento = linha[2];
        var igrejaLocal = linha[3];
        var agendamento = linha[8];
        var missionario = linha[9];
        var responsavel = linha[10];
        var telefone = linha[11];
        var status = linha[13];

        Logger.log("Data: " + data + ", Horário: " + horario + ", Local: " + igrejaLocal);
        
        var tituloEvento = status === "Cancelado" 
          ? "Cancelado | " + missionario + " " + agendamento 
          : agendamento + " | " + missionario;

        if (typeof horario === "string" && horario.includes(":")) {
          var partesHorario = horario.split(":");
          
          if (partesHorario.length === 2) {
            var horas = parseInt(partesHorario[0]);
            var minutos = parseInt(partesHorario[1]);
            
            var horarioString = (horas < 10 ? "0" : "") + horas + ":" + (minutos < 10 ? "0" : "") + minutos;

            var dataEvento = new Date(data);
            dataEvento.setHours(horas, minutos);
            var agora = new Date();
            if (dataEvento < agora) {
              Logger.log("Erro: Tentativa de criar evento em uma data que já passou: " + dataEvento);
              continue;
            }

            var fimEvento = new Date(dataEvento);
            fimEvento.setMinutes(fimEvento.getMinutes() + 30);

            var eventosExistentes = agenda.getEvents(dataEvento, fimEvento);
            var eventoExistente = eventosExistentes.find(evento => 
              evento.getTitle() === tituloEvento &&
              evento.getDescription().includes("Mobilizador Responsável: " + responsavel)
            );

            if (eventoExistente) {
              var eventoStatusAtual = eventoExistente.getTitle().includes("Cancelado") ? "Cancelado" : "Confirmado";
              
              if (status !== eventoStatusAtual) {
                Logger.log("Mudança de status detectada para: " + tituloEvento + ". Excluindo o evento antigo.");
                eventoExistente.deleteEvent();
              } else {
                Logger.log("Evento já existe: " + tituloEvento + " no horário " + horarioString + " em " + aba.getName());
                continue;
              }
            }

            var descricaoEvento = status === "Cancelado" 
              ? `Evento Cancelado\nTelefone: ${telefone}\nHorário: ${horarioString}\nMobilizador Responsável: ${responsavel}`
              : `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horarioString}\nMobilizador Responsável: ${responsavel}`;
              
            var criarEvento = agenda.createEvent(tituloEvento, dataEvento, fimEvento, {
              location: igrejaLocal,
              description: descricaoEvento
            });

            Logger.log("Evento criado: " + criarEvento.getTitle() + " no horário " + horarioString + " em " + aba.getName());
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
