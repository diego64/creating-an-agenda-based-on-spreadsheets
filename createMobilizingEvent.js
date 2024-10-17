function criarEventoMobilizador() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName('Mobilizador'); // Ler a aba 'Mobilizador'
  
  // Buscar agendas específicas
  var agendaMissionario = 'Agendamento Missionário';
  var agendaCancelado = 'Agendamento Cancelado';
  var agendasMissionario = CalendarApp.getCalendarsByName(agendaMissionario);
  var agendasCancelado = CalendarApp.getCalendarsByName(agendaCancelado);
  
  var agendaM = agendasMissionario.length > 0 ? agendasMissionario[0] : null;
  var agendaC = agendasCancelado.length > 0 ? agendasCancelado[0] : null;
  
  if (!agendaM || !agendaC) {
    Logger.log("Uma ou ambas as agendas não foram encontradas.");
    return;
  }

  // Capturar os dados a partir da terceira linha
  var ultimaLinha = aba.getLastRow(); // Acesso a ultima linha da planilha 
  if (ultimaLinha < 3) {
    Logger.log("Não há dados suficientes na aba 'Mobilizador'.");
    return;
  }

  var dados = aba.getRange(3, 1, ultimaLinha - 2, aba.getLastColumn()).getValues();

  for (var j = 0; j < dados.length; j++) {
    var linha = dados[j];

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
      var responsavel = linha[9]; // RESPONSÁVEL
      var telefone = linha[10]; // TELEFONE
      var email = linha[11]; // E-MAIL
      var nome_pastor = linha[12]; // NOME DO PASTOR
      var status = linha[13]; // STATUS

       // Log dos dados capturados
      Logger.log(
        'Data: ' + data + ', Horário: ' + horario +
        ', Tipo de Agendamento: ' + tipo_de_agendamento +
        ', Igreja/Local: ' + igrejaLocal +
        ', Endereço: ' + endereco + ', Bairro: ' + bairro +
        ', Cidade: ' + cidade + ', UF: ' + uf +
        ', Agendamento: ' + agendamento +
        ', Responsável: ' + responsavel +
        ', Telefone: ' + telefone +
        ', E-mail: ' + email +
        ', Nome do Pastor: ' + nome_pastor +
        ', Status: ' + status
      );

      // Criar título do evento
      var tituloEvento = agendamento + ' | ' + responsavel;

      // Verificar se o horário está no formato correto
      if (typeof horario === 'string' && horario.includes(':')) {
        var partesHorario = horario.split(':'); // Transformar a string no formato de hora

        if (partesHorario.length === 2) { // Recebe 0900
          var horas = parseInt(partesHorario[0]); // [09]
          var minutos = parseInt(partesHorario[1]); // [00]

          // Criar um objeto Date com os dados formatados
          var dataEvento = new Date(data);
          dataEvento.setHours(horas, minutos); //09:00

          // Verificar se a data do evento que está sendo criada é anterior à data atual
          var agora = new Date();
          if (dataEvento < agora) {
            Logger.log('Tentativa de criar evento em uma data que já passou: ' + dataEvento);
            continue; // Pula para a próxima linha se a data já passou
          }

          // Definir a duração do evento em 30 minutos
          var fimEvento = new Date(dataEvento);
          fimEvento.setMinutes(fimEvento.getMinutes() + 30);

          // Verificar se já existe um evento com o mesmo título no mesmo horário nas duas agendas
          var eventosExistentesM = agendaM.getEvents(dataEvento, fimEvento);
          var eventosExistentesC = agendaC.getEvents(dataEvento, fimEvento);

          var eventoExistenteM = eventosExistentesM.find(evento => evento.getTitle() === tituloEvento);
          var eventoExistenteC = eventosExistentesC.find(evento => evento.getTitle() === tituloEvento);

          // Criar o evento apenas se não existir
          if (!eventoExistenteM && !eventoExistenteC) {
            // Adiciona o evento na agenda correta
            if (status === 'Cancelado') {
              agendaC.createEvent(tituloEvento, dataEvento, fimEvento, {
                location: igrejaLocal,
                description: `Evento cancelado.`
              });
              Logger.log('Evento criado na agenda AGENDAMENTO CANCELADO: ' + tituloEvento);
            } else if (status === 'Confirmado') {
              agendaM.createEvent(tituloEvento, dataEvento, fimEvento, {
                location: igrejaLocal,
                description: `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horario}`
              });
              Logger.log('Evento criado na agenda AGENDAMENTO MISSIONARIO: ' + tituloEvento);
            }
          } else {
            Logger.log('Evento já existe nas agendas: ' + tituloEvento);
          }
        } else {
          Logger.log("Erro: O horário não está no formato 'HH:mm': " + horario);
        }
      } else {
        Logger.log('Erro: O valor de horário é inválido ou não está em formato de string: ' + horario);
      }
    }
  }
  
  // Verificar e mover eventos baseados na coluna N após a criação
  for (var k = 3; k <= ultimaLinha; k++) {
    var statusAtual = aba.getRange(k, 14).getValue(); // Obter valor da coluna N
    moverEvento(k, statusAtual, agendaM, agendaC);
  }
}

function moverEvento(linha, status, agendaM, agendaC) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName('Mobilizador');

  // Captura os dados da linha
  var dados = aba.getRange(linha, 1, 1, aba.getLastColumn()).getValues()[0];

  var data = dados[0]; // DATA
  var horario = dados[1]; // HORÁRIO
  var tipo_de_agendamento = dados[2]; // TIPO DE AGENDAMENTO
  var agendamento = dados[8]; // AGENDAMENTO
  var responsavel = dados[9]; // RESPONSÁVEL
  var telefone = dados[10]; // TELEFONE
  var igrejaLocal = dados[3]; // IGREJA/LOCAL

  // Criar título do evento
  var tituloEvento = agendamento + ' | ' + responsavel;

  // Verificar se o horário está no formato correto
  if (typeof horario === 'string' && horario.includes(':')) {
    var partesHorario = horario.split(':'); // Transformar a string no formato de hora

    if (partesHorario.length === 2) {
      var horas = parseInt(partesHorario[0]);
      var minutos = parseInt(partesHorario[1]);

      // Criar um objeto Date com os dados formatados
      var dataEvento = new Date(data);
      dataEvento.setHours(horas, minutos);

      // Definir a duração do evento em 30 minutos
      var fimEvento = new Date(dataEvento);
      fimEvento.setMinutes(fimEvento.getMinutes() + 30);

      // Verificar se já existe um evento com o mesmo título no mesmo horário nas duas agendas
      var eventosExistentesM = agendaM.getEvents(dataEvento, fimEvento);
      var eventosExistentesC = agendaC.getEvents(dataEvento, fimEvento);

      var eventoExistenteM = eventosExistentesM.find(evento => evento.getTitle() === tituloEvento);
      var eventoExistenteC = eventosExistentesC.find(evento => evento.getTitle() === tituloEvento);

      // Verificar se a data do evento que está sendo movido é anterior à data atual
      var agora = new Date();
      if (dataEvento < agora) {
        Logger.log('Erro: Tentativa de mover evento para uma data que já passou: ' + dataEvento);
        return; // Retorna sem mover o evento se a data já passou
      }

      if (status === 'Cancelado') {
        // Se o status for 'Cancelado', mover para a agenda de "AGENDAMENTO CANCELADO"
        if (!eventoExistenteC) {
          agendaC.createEvent(tituloEvento, dataEvento, fimEvento, {
            location: igrejaLocal,
            description: `Evento cancelado.`
          });
          Logger.log('Evento movido para AGENDAMENTO CANCELADO: ' + tituloEvento);
        } else {
          Logger.log('Evento já existe na agenda AGENDAMENTO CANCELADO: ' + tituloEvento);
        }
        // Remover da agenda AGENDAMENTO MISSIONARIO, se existir
        if (eventoExistenteM) {
          eventoExistenteM.deleteEvent();
          Logger.log('Evento removido da agenda AGENDAMENTO MISSIONARIO: ' + tituloEvento);
        }
      } else if (status === 'Confirmado') {
        // Se o status for 'Confirmado', mover para a agenda "AGENDAMENTO MISSIONARIO"
        if (!eventoExistenteM) {
          agendaM.createEvent(tituloEvento, dataEvento, fimEvento, {
            location: igrejaLocal,
            description: `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horario}`
          });
          Logger.log('Evento movido para AGENDAMENTO MISSIONARIO: ' + tituloEvento);
        } else {
          Logger.log('Evento já existe na agenda AGENDAMENTO MISSIONARIO: ' + tituloEvento);
        }
        // Remover da agenda Cancelado, se existir
        if (eventoExistenteC) {
          eventoExistenteC.deleteEvent();
          Logger.log('Evento removido da agenda AGENDAMENTO CANCELADO: ' + tituloEvento);
        }
      }
    }
  }
}
