function criarEventoMobilizador() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet(); // Seleciona a planilha ativa
  var aba = planilha.getSheetByName('Mobilizador'); // Ler a aba 'Mobilizador'
  
  // Buscar uma agenda específica pelo seu nome
  var agendaNome = 'Agendamento Missionário';
  var agendas = CalendarApp.getCalendarsByName(agendaNome);
  var agenda = agendas.length > 0 ? agendas[0] : null;
  
  if (!agenda) {
    Logger.log("A agenda '" + agendaNome + "' não foi encontrada.");
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
    }

    // Criar título do evento
    var tituloEvento = agendamento + ' | ' + responsavel;

    // Verificae se o horário está no formato correto
    if (typeof horario === 'string' && horario.includes(':')) {
      var partesHorario = horario.split(':'); // Transformar a string no formato de hora (Hora e minuto) 0900 --> 09:00 | 1330 --> 13:00

      // Verificar se o horario tem duas parte HH:mm
      if (partesHorario.length === 2) {
        var horas = parseInt(partesHorario[0]); // [15]
        var minutos = parseInt(partesHorario[1]); //  [30]

        // Formatar o horário do tipo string para o formato "HH:mm"
        var horarioString = (horas < 10 ? '0' : '') + horas + ':' + (minutos < 10 ? '0' : '') + minutos;

        // Criar um objeto Date com os dados formatados (Data e Horario)
        var dataEvento = new Date(data);
        dataEvento.setHours(horas, minutos);

        // Verificar se a data do evento que está sendo criada é anterior à data atual do evento que está sendo criado
        var agora = new Date();
        if (dataEvento < agora) {
          Logger.log('Erro: Tentativa de criar evento em uma data que já passou: ' + dataEvento);
          continue;
        }

        // Definir a duração do evento em 30 minutos
        var fimEvento = new Date(dataEvento);
        fimEvento.setMinutes(fimEvento.getMinutes() + 30);

        // Verificar se já existe um evento com o mesmo título no mesmo horário
        var eventosExistentes = agenda.getEvents(dataEvento, fimEvento);
        var eventoExistente = eventosExistentes.find(evento => evento.getTitle() === tituloEvento);

        if (eventoExistente) {
          Logger.log('Evento já existe: ' + tituloEvento + ' no horário ' + horarioString + ' em ' + aba.getName());
        } else {
          // Criar o evento na agenda "Agendamento Missionário"
          var criarEvento = agenda.createEvent(tituloEvento, dataEvento, fimEvento, {
            location: igrejaLocal,
            description: `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horarioString}\nMobilizador responsável: ${responsavel}`
          });

          Logger.log('Evento criado: ' + criarEvento.getTitle() + ' no horário ' + horarioString + ' em ' + aba.getName());
        }
      } else {
        Logger.log("Erro: O horário não está no formato 'HH:mm': " + horario);
      }
    } else {
      Logger.log('Erro: O valor de horário é inválido ou não está em formato de string: ' + horario);
    }
  }
}
