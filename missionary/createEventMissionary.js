function criarEventoMissionario() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var todasAbas = planilha.getSheets();
  
  // Definindo a agenda onde todos os eventos serão criados
  var agendaMissionario = "Agendamento Missionário";
  var agendas = CalendarApp.getCalendarsByName(agendaMissionario);
  var agenda = agendas.length > 0 ? agendas[0] : null;

  // Verificação da existência da agenda
  if (!agenda) {
    Logger.log("A agenda '" + agendaMissionario + "' não foi encontrada.");
    return;
  }

  // Ler da 3ª aba em diante
  for (var i = 2; i < todasAbas.length; i++) {
    var aba = todasAbas[i];
    var valores = aba.getDataRange().getValues();
    
    Logger.log("Dados da aba " + aba.getName() + ": " + JSON.stringify(valores));
    
    // Percorre os dados a partir da 3ª linha
    for (var j = 2; j < valores.length; j++) {
      var linha = valores[j];
      
      // Verifica se a linha tem valores para evitar erros
      if (linha[0] && linha[1]) {
        var data = linha[0]; // DATA
        var horario = linha[1]; // HORÁRIO
        var tipo_de_agendamento = linha[2]; // TIPO DE AGENDAMENTO
        var igrejaLocal = linha[3]; // IGREJA/LOCAL
        var agendamento = linha[8]; // AGENDAMENTO
        var missionario = linha[9]; // MISSIONÁRIO
        var responsavel = linha[10]; // RESPONSÁVEL PELO AGENDAMENTO
        var telefone = linha[11]; // TELEFONE
        var status = linha[13]; // STATUS

        Logger.log("Data: " + data + ", Horário: " + horario + ", Local: " + igrejaLocal);
        
        // Criar o título do evento
        var tituloEvento = status === "Cancelado" 
          ? "Cancelado | " + missionario + " " + agendamento 
          : agendamento + " | " + missionario;

        // Verifica se o horário está no formato correto
        if (typeof horario === "string" && horario.includes(":")) {
          var partesHorario = horario.split(":"); // Divide a string de horário em horas e minutos
          
          // Verificar se o horário tem duas partes
          if (partesHorario.length === 2) {
            var horas = parseInt(partesHorario[0]);
            var minutos = parseInt(partesHorario[1]);
            
            // Formatar o horário como string no formato "HH:mm"
            var horarioString = (horas < 10 ? "0" : "") + horas + ":" + (minutos < 10 ? "0" : "") + minutos;

            // Criar um objeto Date
            var dataEvento = new Date(data);
            dataEvento.setHours(horas, minutos); // Ajusta horas e minutos

            // Verificar se a data do evento é anterior à data atual
            var agora = new Date();
            if (dataEvento < agora) {
              Logger.log("Erro: Tentativa de criar evento em uma data que já passou: " + dataEvento);
              continue;
            }

            // Define a duração padrão do evento em 30 minutos
            var fimEvento = new Date(dataEvento);
            fimEvento.setMinutes(fimEvento.getMinutes() + 30);

            // Verificar se já existe um evento com o mesmo horário
            var eventosExistentes = agenda.getEvents(dataEvento, fimEvento);
            var eventoExistente = eventosExistentes.find(evento => 
              evento.getTitle().includes(missionario) || 
              evento.getTitle().includes(agendamento)
            );

            // Se já existe um evento no mesmo horário, verifica se o status mudou
            if (eventoExistente) {
              var eventoStatusAtual = eventoExistente.getTitle().includes("Cancelado") ? "Cancelado" : "Confirmado";
              
              // Se o status mudou, exclui o evento antigo
              if (status !== eventoStatusAtual) {
                Logger.log("Mudança de status detectada para: " + tituloEvento + ". Excluindo o evento antigo.");
                eventoExistente.deleteEvent(); // Exclusão do evento antigo
              } else {
                Logger.log("Evento já existe: " + tituloEvento + " no horário " + horarioString + " em " + aba.getName());
                continue; // Pula para a próxima iteração se o evento existe e não mudou
              }
            }

            // Criar descrição do evento
            var descricaoEvento = status === "Cancelado" 
              ? `Evento Cancelado\nTelefone: ${telefone}\nHorário: ${horarioString}\nMobilizador Responsável: ${responsavel}`
              : `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horarioString}\nMobilizador Responsável: ${responsavel}`;
              
            // Criar novo evento na agenda
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
