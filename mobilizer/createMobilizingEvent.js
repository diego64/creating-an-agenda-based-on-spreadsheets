function criarEventoMobilizador() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaMobilizador = planilha.getSheetByName("Mobilizador"); // Ler somente a aba "Mobilizador"
  
  // Definir a agenda no qual o evento será cadastrado
  var agendaMissionario = "Agendamento Missionário";
  var agendas = CalendarApp.getCalendarsByName(agendaMissionario);
  var agenda = agendas.length > 0 ? agendas[0] : null;

  // Verificar a existência da agenda
  if (!agenda) {
    Logger.log("A agenda '" + agendaMissionario + "' não foi encontrada.");
    
    // Listar todas as agendas para verificar o nome correto
    var todasAgendas = CalendarApp.getAllCalendars();
    todasAgendas.forEach(cal => Logger.log("Agenda encontrada: " + cal.getName()));
    
    return;
  }

  // Obter os valores da aba "Mobilizador"
  var valores = abaMobilizador.getDataRange().getValues();
  
  // Percorrer os dados a partir da 3ª linha
  for (var j = 2; j < valores.length; j++) {
    var linha = valores[j];
    
    // Verificar se a linha tem valores para evitar erros
    if (linha[0] && linha[1]) {
      var data = linha[0]; // DATA
      var horario = linha[1]; // HORÁRIO
      var tipo_de_agendamento = linha[2]; // TIPO DE AGENDAMENTO
      var igrejaLocal = linha[3]; // IGREJA/LOCAL
      var endereco = linha[4]; // ENDEREÇO
      var bairro = linha[5]; // BAIRRO
      var cidade = linha[6]; // CIDADE
      var uf = linha[7]; // UF
      var agendamento = linha[8]; // AGENDAMENTO
      var responsavel = linha[9]; // RESPONSÁVEL PELO AGENDAMENTO
      var telefone = linha[10]; // TELEFONE
      var email = linha[11]; // E-MAIL
      var nomePastor = linha[12]; // NOME DO PASTOR
      var status = linha[13]; // STATUS

      Logger.log("Data: " + data + ", Horário: " + horario + ", Local: " + igrejaLocal);
      
      // Criar o título do evento
      var tituloEvento = status === "Cancelado" 
        ? "Cancelado | " + responsavel // Título do evento cancelado
        : agendamento + " | " + responsavel;

      // Verificar se o horário está no formato correto
      if (typeof horario === "string" && horario.includes(":")) {
        var partesHorario = horario.split(":"); // Divide a string de horário em horas e minutos
        
        // Verificar se o horário tem duas partes
        if (partesHorario.length === 2) {
          var horas = parseInt(partesHorario[0]);
          var minutos = parseInt(partesHorario[1]);
          
          // Formatar o horário como string no formato "HH:mm"
          var horarioString = (horas < 10 ? "0" : "") + horas + ":" + (minutos < 10 ? "0" : "") + minutos;

          // Ajustar a data para o formato correto
          var dataFormatada = Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM-dd");
          var dataEvento = new Date(dataFormatada + "T" + horarioString + ":00");

          // Verificar se a data do evento é anterior à data atual
          var agora = new Date();
          if (dataEvento < agora) {
            Logger.log("Erro: Tentativa de criar evento em uma data que já passou: " + dataEvento);
            continue;
          }

          // Definir a duração padrão do evento em 30 minutos (Padrão)
          var fimEvento = new Date(dataEvento);
          fimEvento.setMinutes(fimEvento.getMinutes() + 30);

          // Verificar se já existe um evento com o mesmo horário e nome de responsável diferente
          var eventosExistentes = agenda.getEvents(dataEvento, fimEvento);
          var eventoExistente = eventosExistentes.find(evento => 
            evento.getStartTime().getTime() === dataEvento.getTime() &&
            evento.getTitle().includes(agendamento) &&
            evento.getTitle().includes(responsavel)
          );

          // Se já existe um evento no mesmo horário e nome de responsável diferente
          if (eventoExistente) {
            var eventoStatusAtual = eventoExistente.getTitle().includes("Cancelado") ? "Cancelado" : "Confirmado";
            
            // Se o status mudou, exclui o evento antigo
            if (status !== eventoStatusAtual) {
              Logger.log("Mudança de status detectada para: " + tituloEvento + ". Excluindo o evento antigo.");
              eventoExistente.deleteEvent(); // Exclusão do evento antigo
            } else {
              Logger.log("Evento já existe: " + tituloEvento + " no horário " + horarioString);
              continue; // Pula para a próxima iteração se o evento existe e não mudou
            }
          }

          // Criar descrição do evento
          var descricaoEvento = status === "Cancelado" 
            ? "Evento Cancelado"  // Apenas "Evento Cancelado" se for cancelado
            : `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horarioString}`; // Descrição completa se não for cancelado
          
          // Criar novo evento na agenda
          var criarEvento = agenda.createEvent(tituloEvento, dataEvento, fimEvento, {
            location: igrejaLocal,
            description: descricaoEvento
          });

          Logger.log("Evento criado: " + criarEvento.getTitle() + " no horário " + horarioString);
        } else {
          Logger.log("Erro: O horário não está no formato 'HH:mm': " + horario);
        }
      } else {
        Logger.log("Erro: O valor de horário é inválido ou não está em formato de string: " + horario);
      }
    }
  }
}
