function rodarScriptMissionario() {
    const agora = new Date();
    const horaAtual = agora.getHours();
    const minutosAtuais = agora.getMinutes();
  
    // Verifica se a hora atual é diferente de 17:00
    if (horaAtual !== 17 || minutosAtuais !== 0) {
      Logger.log('Informamos que a execução da função em questão não pôde ser efetuada, pois o horário estabelecido para sua ativação ainda não foi alcançado ou já foi ultrapassado.');
      return; // Sai da função se não for o horário correto
    }
  
    // Remover gatilhos anteriores com o mesmo nome
    const allTriggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getHandlerFunction() === 'criarEventoMissionario') {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    }
  
    // Criar um gatilho diário para ser executado às 17:00
    ScriptApp.newTrigger('criarEventoMissionario')
      .timeBased()
      .everyDays(1) // Executar todos os dias
      .atHour(17) // Executar às 17:00
      .create();
  
    Logger.log('Gatilho AGENDAMENTO MISSIONARIO acionado com SUCESSO!');
  }
  
  function rodarScriptMissionarioManual() {
    Logger.log('Execução do gatilho manual AGENDAMENTO MISSIONARIO iniciada.');
  
    // Chama a função que cria os eventos
    criarEventoMissionario();
  
    Logger.log('Execução manual concluída.');
  }
  