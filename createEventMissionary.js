function criarEventoMissionario() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var todasAbas = planilha.getSheets();
  
  // Obtém a agenda específica pelo nome
  var agendaNome = "Agendamento Missionário";
  var agendas = CalendarApp.getCalendarsByName(agendaNome);
  var agenda = agendas.length > 0 ? agendas[0] : null;
  
  if (!agenda) {
    Logger.log("A agenda '" + agendaNome + "' não foi encontrada.");
    return;
  }

  // Ler da 3ª aba em diante
  for (var i = 2; i < todasAbas.length; i++) {
    var aba = todasAbas[i];
    var valores = aba.getDataRange().getValues();
    
    // Obtém os títulos da segunda linha
    var titulos = valores[1];
    
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
        var endereco = linha[4]; // ENDEREÇO
        var bairro = linha[5]; // BAIRRO
        var cidade = linha[6]; // CIDADE
        var uf = linha[7]; // UF
        var agendamento = linha[8]; // AGENDAMENTO
        var missionario = linha[9]; // MISSIONÁRIO
        var responsavel = linha[10]; // RESPONSÁVEL PELO AGENDAMENTO
        var telefone = linha[11]; // TELEFONE
        var email = linha[12]; // E-MAIL
        var status = linha[13]; // STATUS
        
        // Log dos valores da linha para depuração
        Logger.log("Data: " + data + ", Horário: " + horario + ", Local: " + igrejaLocal);
        
        // Cria o título do evento
        var tituloEvento = agendamento + " | " + missionario;

        // Verifica se o horário está no formato correto
        if (typeof horario === "string" && horario.includes(":")) {
          var partesHorario = horario.split(":"); // Divide a string de horário em horas e minutos
          
          // Verifica se tem duas partes
          if (partesHorario.length === 2) {
            // Converte horas e minutos em números inteiros
            var horas = parseInt(partesHorario[0]);
            var minutos = parseInt(partesHorario[1]);
            
            // Formata o horário como string no formato "HH:mm"
            var horarioString = (horas < 10 ? "0" : "") + horas + ":" + (minutos < 10 ? "0" : "") + minutos;

            // Formata a data e o horário para criar um objeto Date
            var dataEvento = new Date(data);
            dataEvento.setHours(horas, minutos); // Ajusta horas e minutos

            // Verifica se a data do evento é anterior à data atual
            var agora = new Date();
            if (dataEvento < agora) {
              Logger.log("Erro: Tentativa de criar evento em uma data que já passou: " + dataEvento);
              continue;
            }

            // Define a duração do evento (30 minutos)
            var fimEvento = new Date(dataEvento);
            fimEvento.setMinutes(fimEvento.getMinutes() + 30);

            // Verifica se já existe um evento com o mesmo título e no mesmo horário
            var eventosExistentes = agenda.getEvents(dataEvento, fimEvento);
            var eventoExistente = eventosExistentes.find(evento => evento.getTitle() === tituloEvento);

            if (eventoExistente) {
              Logger.log("Evento já existe: " + tituloEvento + " no horário " + horarioString + " em " + aba.getName());
            } else {
              // Criação do evento na agenda "Agendamento Missionário"
              var criarEvento = agenda.createEvent(tituloEvento, dataEvento, fimEvento, {
                location: igrejaLocal,
                description: `Evento: ${tipo_de_agendamento}\nTelefone: ${telefone}\nHorário: ${horarioString}\nMobilizador responsável: ${responsavel}`
              });

              Logger.log("Evento criado: " + criarEvento.getTitle() + " no horário " + horarioString + " em " + aba.getName());
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

function verificarAtualizacoes() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var todasAbas = planilha.getSheets();
  var atualizacoes = [];
  var horarioAtualizacao = new Date().toLocaleString(); // Formata a data e hora atual

  // Ler da 3ª aba em diante
  for (var i = 2; i < todasAbas.length; i++) {
    var aba = todasAbas[i];
    var valores = aba.getDataRange().getValues();

    // Verifica se há mais de duas linhas de dados
    if (valores.length > 2) {
      var alteracoesDetectadas = false; // Variável para rastrear alterações

      // Percorre as linhas a partir da 3ª linha
      for (var j = 2; j < valores.length; j++) {
        var linha = valores[j];

        // Verificar se a linha tem dados
        if (linha.some(campo => campo !== "")) { // Verifica se algum campo na linha não está vazio
          alteracoesDetectadas = true; // Marca que foram detectadas alterações
          break; // Sai do loop, pois já detectamos alterações
        }
      }

      // Se foram detectadas alterações, adiciona o nome da aba
      if (alteracoesDetectadas) {
        atualizacoes.push(aba.getName());
      }
    }
  }

  // Se houver atualizações, envia um e-mail
  if (atualizacoes.length > 0) {
    var listaAbasAtualizadas = atualizacoes.join(", "); // Concatena os nomes das abas atualizadas
    var emailDestino = "john.doe@jmm.org.br"; // Altere para o e-mail real

    // Mensagem do e-mail em HTML
    var mensagemEmail = `
      <div style="font-family: 'Sans Serif', Arial, sans-serif; font-size: 14px; color: #333;">
        <p>Olá Setor de Promoção, tudo bem?</p>

        <p>Houve uma nova atualização na(s) aba(s) <b>${listaAbasAtualizadas}</b> às <b>${horarioAtualizacao}</b>.</p>

        <p>Caso tenha alguma dúvida, entre em contato com o setor de Suporte Técnico.</p>

        <p>Atenciosamente,</p>

        <br>

        <p style="font-weight: bold; margin-top: 20px;">
          Sistema de Agendamento Automático <br>
          Copyright (c) 2024 Diego Ferreira L.G. Oliveira <br>
          Tecnologia da Informação <br>
          (21) 2122-1900 Ramal 2001
        </p>
      </div>
    `;

    // Envio do e-mail
    MailApp.sendEmail({
      to: emailDestino,
      subject: "Atualizações nas abas",
      htmlBody: mensagemEmail
    });
  }
}
