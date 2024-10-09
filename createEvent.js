function criarEvento() {
  var valores = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();

  Logger.log(valores);

  var agenda = CalendarApp.getDefaultCalendar();

  for (var i = 1; i < valores.length; i++) {
    var linha = valores[i];

    var data = linha[0]; // DATA
    var horario = linha[1]; // HORÁRIO
    var tipoAgendamento = linha[2]; // TIPO DE AGENDAMENTO
    var tarefa = linha[3]; // TAREFA
    var igrejaLocal = linha[4]; // IGREJA/LOCAL
    var bairro = linha[5]; // BAIRRO
    var cidade = linha[6]; // CIDADE
    var email = linha[7]; // E-MAIL
    var uf = linha[8]; // UF
    var responsavel = linha[9]; // RESPONSÁVEL PELO AGENDAMENTO
    var telefone = linha[10]; // TELEFONE
    var contato = linha[11]; // CONTATO
    var nomePastor = linha[12]; // NOME DO PASTOR
    var observacoes = linha[13]; // OBSERVAÇÕES
    var status = linha[14]; // STATUS DO AGENDAMENTO

    // Converter data e horário
    var dataInicio = new Date(data);
    var partesHorario = horario.split(":");
    dataInicio.setHours(partesHorario[0], partesHorario[1]);

    // Definir data de fim (Duração padrao de 01 hora)
    var dataFim = new Date(dataInicio);
    dataFim.setHours(dataFim.getHours() + 1);

    // Verificar se já existe um evento com o mesmo título e nas mesmas datas
    var eventosExistentes = agenda.getEvents(dataInicio, dataFim, { search: tarefa });

    if (eventosExistentes.length > 0) {
      Logger.log("Evento já existe para a tarefa: " + tarefa);
      continue; // Pula para o próximo evento, evitando duplicados
    }

    // Definir a cor com base no status
    var cor;
    switch (status) {
      case "Confirmado":
        cor = CalendarApp.EventColor.GREEN;
        break;
      case "A confirmar":
        cor = CalendarApp.EventColor.YELLOW;
        break;
      case "Cancelado":
        cor = CalendarApp.EventColor.RED;
        break;
      default:
        cor = CalendarApp.EventColor.BLUE;
    }

    // Criar o evento
    var eventoCriado = agenda.createEvent(tarefa, dataInicio, dataFim, {
      location: igrejaLocal,
      description: `
          Status do Agendamento: ${status}
          Tipo de Agendamento: ${tipoAgendamento}
          Responsável: ${responsavel}
          Telefone: ${telefone}
          Contato: ${contato}
          Estado (UF): ${uf}
          Cidade: ${cidade}
          Bairro: ${bairro}
          Observações: ${observacoes}
          Nome do Pastor: ${nomePastor}
      `,
      guests: email
    });

    // Definir a cor do evento
    eventoCriado.setColor(cor);

    Logger.log("Evento criado com sucesso: " + eventoCriado.getId());
  }
}