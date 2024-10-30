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
        var data = new Date(linha[0]); // DATA
        var horario = linha[1]; // HORÁRIO
        var localEvento = linha[3]; // IGREJA/LOCAL
        var tipoEvento = linha[2]; // TIPO DE AGENDAMENTO
        var agendamento = linha[8]; // UF
        var missionario = linha[9];
        var responsavel = linha[10]; // RESPONSÁVEL
        var telefoneResponsavel = linha[11]; // TELEFONE
        var status = linha[13]; // STATUS

        Logger.log("Data: " + data + ", Horário: " + horario + ", Missionário: " + missionario + ", Status: " + status);

        if (typeof horario === "string" && horario.includes(":")) {
          var partesHorario = horario.split(":");

          if (partesHorario.length === 2) {
            var horas = parseInt(partesHorario[0]);
            var minutos = parseInt(partesHorario[1]);

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
              evento.getTitle().includes(missionario)
            );

            if (status === "Confirmado") {
              if (eventoExistente) {
                eventoExistente.deleteEvent();
              }

              var descricaoEvento = `Evento: ${tipoEvento}\nTelefone: ${telefoneResponsavel}\nHorário: ${horario}\nMobilizador Responsável: ${responsavel}`;
              var novoEvento = agenda.createEvent(`Missionário | ${missionario}`, dataEvento, fimEvento, {
                location: localEvento,
                description: descricaoEvento
              });

              Logger.log("Evento Atualizado: " + novoEvento.getTitle() + " no horário " + horario);

            } else if (status === "Cancelado") {
              if (eventoExistente) {
                eventoExistente.deleteEvent();
              }

              var descricaoEventoCancelado = `Tipo de Evento: ${tipoEvento} (Cancelado)\nEvento: ${agendamento}\nTelefone: ${telefoneResponsavel}\nHorário: ${horario}\nMobilizador Responsável: ${responsavel}`;
              var novoEventoCancelado = agenda.createEvent(`Cancelado | ${missionario}`, dataEvento, fimEvento, {
                location: localEvento,
                description: descricaoEventoCancelado
              });

              Logger.log("Evento Criado: " + novoEventoCancelado.getTitle() + " no horário " + horario);
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
