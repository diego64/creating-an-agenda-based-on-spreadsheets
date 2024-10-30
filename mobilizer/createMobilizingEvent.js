function criarEventoMobilizador() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaMobilizador = planilha.getSheetByName("Mobilizador");

  // Definir a agenda na qual o evento será cadastrado
  var agendaMissionario = "Agendamento Missionário";
  var agenda = CalendarApp.getCalendarsByName(agendaMissionario)[0];

  if (!agenda) {
    Logger.log("A agenda '" + agendaMissionario + "' não foi encontrada.");
    return;
  }

  var valores = abaMobilizador.getDataRange().getValues();

  for (var j = 2; j < valores.length; j++) {
    var linha = valores[j];

    if (linha[0] && linha[1]) {
      var data = linha[0]; // DATA
      var horario = linha[1]; // HORÁRIO
      var agendamento = linha[8]; // AGENDAMENTO
      var responsavel = linha[9]; // RESPONSÁVEL PELO AGENDAMENTO
      var status = linha[13]; // STATUS

      if (typeof horario === "string" && horario.includes(":")) {
        var partesHorario = horario.split(":");

        if (partesHorario.length === 2) {
          var horas = parseInt(partesHorario[0]);
          var minutos = parseInt(partesHorario[1]);
          var horarioString = (horas < 10 ? "0" : "") + horas + ":" + (minutos < 10 ? "0" : "") + minutos;
          var dataFormatada = Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM-dd");
          var dataEvento = new Date(dataFormatada + "T" + horarioString + ":00");

          // Ignorar eventos para datas passadas
          if (dataEvento < new Date()) continue;

          // Definir fim do evento
          var fimEvento = new Date(dataEvento);
          fimEvento.setMinutes(fimEvento.getMinutes() + 30);

          // Procurar eventos existentes na agenda
          var eventosExistentes = agenda.getEvents(dataEvento, fimEvento);
          var eventoExistente = eventosExistentes.find(evento => 
            evento.getTitle().includes(responsavel) && evento.getStartTime().getTime() === dataEvento.getTime()
          );

          // Definir título do evento com base no status
          var tituloEvento = status === "Cancelado"
            ? "Cancelado | " + responsavel // Título para evento cancelado
            : agendamento + " | " + responsavel; // Título para evento confirmado

          if (eventoExistente) {
            // Verifica se o status atual do evento é diferente do status na linha
            var eventoStatusAtual = eventoExistente.getTitle().includes("Cancelado") ? "Cancelado" : "Confirmado";

            if (status !== eventoStatusAtual) {
              Logger.log("Mudança de status detectada. Excluindo o evento: " + eventoExistente.getTitle());
              // Excluir o evento existente
              eventoExistente.deleteEvent();
            } else {
              Logger.log("Evento já existe e está atualizado: " + eventoExistente.getTitle());
              continue; // Não criar um novo evento se já existe e está atualizado
            }
          }

          // Criar um novo evento na agenda
          var descricaoEvento = status === "Cancelado"
            ? "Evento Cancelado" // Apenas "Evento Cancelado" se for cancelado
            : `Evento: ${agendamento}\nResponsável: ${responsavel}\nHorário: ${horarioString}`; // Descrição completa se confirmado

          var criarEvento = agenda.createEvent(tituloEvento, dataEvento, fimEvento, {
            location: linha[3], // IGREJA/LOCAL
            description: descricaoEvento
          });

          Logger.log("Novo evento criado: " + criarEvento.getTitle() + " no horário " + horarioString);
        } else {
          Logger.log("Erro: O horário não está no formato 'HH:mm': " + horario);
        }
      } else {
        Logger.log("Erro: O valor de horário é inválido ou não está em formato de string: " + horario);
      }
    }
  }
}
