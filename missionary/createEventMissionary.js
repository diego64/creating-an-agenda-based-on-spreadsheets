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
        var data = new Date(linha[0]);
        var horario = linha[1];
        var agendamento = linha[8];
        var missionario = linha[9];
        var responsavel = linha[10];
        var status = linha[13];

        Logger.log("Data: " + data + ", Horário: " + horario + ", Missionário: " + missionario + ", Status: " + status);

        if (typeof horario === "string" && horario.includes(":")) {
          var partesHorario = horario.split(":");

          if (partesHorario.length === 2) {
            var horas = parseInt(partesHorario[0]);
            var minutos = parseInt(partesHorario[1]);

            var dataEvento = new Date(data);
            dataEvento.setHours(horas, minutos);
            var agora = new Date();

            // Ignorar eventos (datas) que já passaram
            if (dataEvento < agora) {
              Logger.log("Erro: Tentativa de criar evento em uma data que já passou: " + dataEvento);
              continue;
            }

            var fimEvento = new Date(dataEvento);
            fimEvento.setMinutes(fimEvento.getMinutes() + 30);

            // Buscar eventos existentes no dia e horário
            var eventosExistentes = agenda.getEvents(dataEvento, fimEvento);
            var eventoExistente = eventosExistentes.find(evento => 
              evento.getTitle().includes(missionario)
            );

            // Criar e excluir eventos
            if (status === "Confirmado") {
              // Se houver um evento existente
              if (eventoExistente) {
                // Se o evento existente for "Cancelado", exclua
                if (eventoExistente.getTitle().includes("Cancelado")) {
                  Logger.log("Evento Cancelado existente encontrado: " + eventoExistente.getTitle() + ". Excluindo o evento.");
                  eventoExistente.deleteEvent();
                } 
                // Se for "Confirmado", exclua-o antes de criar um novo
                else {
                  Logger.log("Evento Confirmado existente encontrado: " + eventoExistente.getTitle() + ". Excluindo o evento.");
                  eventoExistente.deleteEvent();
                }
              }

              // Criar novo evento "Confirmado"
              var descricaoEvento = `Evento: ${agendamento}\nMobilizador Responsável: ${responsavel}`;
              var novoEvento = agenda.createEvent(`Confirmado | ${missionario}`, dataEvento, fimEvento, {
                description: descricaoEvento
              });

              Logger.log("Evento criado: " + novoEvento.getTitle() + " no horário " + horario);

            } else if (status === "Cancelado") {
              // Se houver um evento existente
              if (eventoExistente) {
                // Se o evento existente for "Confirmado", exclua
                if (eventoExistente.getTitle().includes("Confirmado")) {
                  Logger.log("Evento Confirmado existente encontrado: " + eventoExistente.getTitle() + ". Excluindo o evento.");
                  eventoExistente.deleteEvent();
                } 
                // Se for "Cancelado", exclua-o antes de criar um novo
                else {
                  Logger.log("Evento Cancelado existente encontrado: " + eventoExistente.getTitle() + ". Excluindo o evento.");
                  eventoExistente.deleteEvent();
                }
              }

              // Criar novo evento "Cancelado"
              var descricaoEventoCancelado = `Evento Cancelado\nMobilizador Responsável: ${responsavel}`;
              var novoEventoCancelado = agenda.createEvent(`Cancelado | ${missionario}`, dataEvento, fimEvento, {
                description: descricaoEventoCancelado
              });

              Logger.log("Evento criado: " + novoEventoCancelado.getTitle() + " no horário " + horario);
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
