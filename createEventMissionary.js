function criarEventoMissionario() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var todasAbas = planilha.getSheets();
  
  // Buscar as duas agendas
  var agendaMissionario = "Agendamento Missionário";
  var agendaCancelada = "Agendamento Cancelado";
  var agendas = CalendarApp.getCalendarsByName(agendaMissionario);
  var agendasCanceladas = CalendarApp.getCalendarsByName(agendaCancelada);
  var agenda = agendas.length > 0 ? agendas[0] : null;
  var agendaCancelada = agendasCanceladas.length > 0 ? agendasCanceladas[0] : null;
  
  // Verificação da existencia das agendas
  if (!agenda) {
    Logger.log("A agenda '" + agendaMissionario + "' não foi encontrada.");
    return;
  }
  
  if (!agendaCancelada) {
    Logger.log("A agenda '" + agendaCancelada + "' não foi encontrada.");
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
        var tituloEvento = agendamento + " | " + missionario;

        // Verifica se o horário está no formato correto
        if (typeof horario === "string" && horario.includes(":")) {
          var partesHorario = horario.split(":"); // Divide a string de horário em horas e minutos
          
          // Verificar se o horario tem duas partes
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

            // Define a duração padrão do eventoem 30 minutos
            var fimEvento = new Date(dataEvento);
            fimEvento.setMinutes(fimEvento.getMinutes() + 30);

            // Verificar se já existe um evento com o mesmo título e no mesmo horário
            var eventosExistentesMissionario = agenda.getEvents(dataEvento, fimEvento);
            var eventosExistentesCancelados = agendaCancelada.getEvents(dataEvento, fimEvento);
            
            // Juntar todos os eventos cadastrados
            var todosEventosExistentes = eventosExistentesMissionario.concat(eventosExistentesCancelados);
            var eventoExistente = todosEventosExistentes.find(evento => evento.getTitle() === tituloEvento);

            // Verificar se houve mudança no status (Coluna N)
            var eventoParaModificar = eventosExistentesMissionario.find(evento => evento.getTitle() === tituloEvento) || 
                                      eventosExistentesCancelados.find(evento => evento.getTitle() === tituloEvento);

            if (eventoParaModificar) {
              var eventoStatusAtual = (eventosExistentesMissionario.includes(eventoParaModificar)) ? "Confirmado" : "Cancelado";
              
              // Se houver mudança no status, exclui da agenda atual e cria na nova
              if (status !== eventoStatusAtual) {
                Logger.log("Mudança de status detectada para: " + tituloEvento + ". Excluindo o evento antigo.");
                eventoParaModificar.deleteEvent(); // Exclusão do evento antigo
                
                // Criação do novo evento na agenda correspondente
                var novaAgenda = (status === "Confirmado") ? agenda : agendaCancelada;
                var criarEvento = novaAgenda.createEvent(tituloEvento, dataEvento, fimEvento, {
                  location: igrejaLocal,
                  description: `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horarioString}`
                });

                Logger.log("Evento criado: " + criarEvento.getTitle() + " no horário " + horarioString + " em " + aba.getName());
              }
            } else if (!eventoExistente) {
              // Criação do evento na agenda correspondente se não houver mudança no status
              var agendaParaEvento = status === "Confirmado" ? agenda : agendaCancelada;
              var criarEvento = agendaParaEvento.createEvent(tituloEvento, dataEvento, fimEvento, {
                location: igrejaLocal,
                description: `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horarioString}`
              });

              Logger.log("Evento criado: " + criarEvento.getTitle() + " no horário " + horarioString + " em " + aba.getName());
            } else {
              Logger.log("Evento já existe: " + tituloEvento + " no horário " + horarioString + " em " + aba.getName());
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
