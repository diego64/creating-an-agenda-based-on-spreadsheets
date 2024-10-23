function relatorioSemanalMissionarios() {
  const planilhas = SpreadsheetApp.getActiveSpreadsheet().getSheets(); // Obter todas as abas
  for (let i = 2; i < planilhas.length; i++) { // Iniciar a leitura a partir da terceira aba (índice 2)
    const planilha = planilhas[i];
    const nomeAba = planilha.getName();
    Logger.log("Lendo dados da aba: " + nomeAba);

    // Obter o dia da semana por extenso
    const diaDaSemana = obterDiaDaSemana(new Date());

    // Se for segunda-feira, resetar os dados da semana
    if (diaDaSemana === 'segunda-feira') {
      PropertiesService.getScriptProperties().deleteProperty('dadosSemana_' + nomeAba);
      PropertiesService.getScriptProperties().setProperty('ultimaLinhaProcessada_' + nomeAba, '2'); // Resetar contagem de linhas processadas
      Logger.log("Início da semana, acumulado de dados resetado para a aba: " + nomeAba);
    }

    // Pegar o valor da última linha processada
    let ultimaLinhaProcessada = parseInt(PropertiesService.getScriptProperties().getProperty('ultimaLinhaProcessada_' + nomeAba) || '2');
    const novaUltimaLinha = planilha.getLastRow();

    // Correção: Ajustar a última linha processada se for maior que a última linha atual da planilha
    if (ultimaLinhaProcessada > novaUltimaLinha) {
      Logger.log("Ajustando a última linha processada, pois ela era maior que a última linha atual.");
      ultimaLinhaProcessada = 2; // Reset para a segunda linha
    }

    Logger.log(`Última linha processada: ${ultimaLinhaProcessada}`);
    Logger.log(`Última linha atual da planilha: ${novaUltimaLinha}`);

    // Verificar se há novas linhas
    if (novaUltimaLinha > ultimaLinhaProcessada) {
      const numLinhasNovas = novaUltimaLinha - ultimaLinhaProcessada;
      Logger.log(`Novas linhas detectadas na aba ${nomeAba}: ${numLinhasNovas}`);

      const dadosNovos = obterDadosNovos(planilha, ultimaLinhaProcessada + 1, numLinhasNovas);

      // Acumular dados da semana
      const dadosAcumulados = JSON.parse(PropertiesService.getScriptProperties().getProperty('dadosSemana_' + nomeAba) || '[]');
      const novosDados = dadosAcumulados.concat(dadosNovos);

      // Atualizar o acumulado de dados
      PropertiesService.getScriptProperties().setProperty('dadosSemana_' + nomeAba, JSON.stringify(novosDados));
      PropertiesService.getScriptProperties().setProperty('ultimaLinhaProcessada_' + nomeAba, novaUltimaLinha.toString());

      Logger.log(`Dados acumulados até agora na aba ${nomeAba}: ${novosDados.length} linhas.`);
    } else {
      Logger.log("Nenhuma nova linha foi adicionada na aba: " + nomeAba);
    }

    // Só enviar o e-mail na sexta-feira
    if (diaDaSemana === 'sexta-feira') {
      // Montar tabela em HTML com todos os dados acumulados
      const dadosSemana = JSON.parse(PropertiesService.getScriptProperties().getProperty('dadosSemana_' + nomeAba) || '[]');
      if (dadosSemana.length > 0) {
        const htmlTable = montarTabela(dadosSemana);

        // Cabeçalho da tabela
        const cabecalho = `
          <tr>
            <th style="text-align: center;">DATA</th>
            <th style="text-align: center;">HORÁRIO</th>
            <th style="text-align: center;">TIPO DE AGENDAMENTO</th>
            <th style="text-align: center;">IGREJA/LOCAL</th>
            <th style="text-align: center;">ENDEREÇO</th>
            <th style="text-align: center;">BAIRRO</th>
            <th style="text-align: center;">CIDADE</th>
            <th style="text-align: center;">UF</th>
            <th style="text-align: center;">AGENDAMENTO</th>
            <th style="text-align: center;">RESPONSÁVEL</th>
            <th style="text-align: center;">TELEFONE</th>
            <th style="text-align: center;">E-MAIL</th>
            <th style="text-align: center;">NOME DO PASTOR</th>
            <th style="text-align: center;">STATUS</th>
          </tr>
        `;

        // Enviar o e-mail
        const destinatario = 'john.doe@jmm.org.br';
        const assunto = 'Relatório Semanal | Agendamentos dos Missionarios (' + nomeAba + ')';
        const corpoEmail = `
          <p>Segue abaixo os dados inseridos de segunda-feira a sexta-feira na aba ${nomeAba}:</p>
          <table border="1" style="border-collapse: collapse; width: 100%;">
            <thead>${cabecalho}</thead>
            <tbody>${htmlTable}</tbody>
          </table>
          <br>
          --<br>
          A.S.S<br>
          Automatic Scheduling System<br>
          MIT License<br>
          Copyright (c) 2024 Diego Ferreira L.G. Oliveira
        `;

        MailApp.sendEmail({
          to: destinatario,
          subject: assunto,
          htmlBody: corpoEmail
        });

        Logger.log("E-mail enviado com sucesso para a aba: " + nomeAba);
        
        // Limpar o acumulado após enviar o e-mail
        PropertiesService.getScriptProperties().deleteProperty('dadosSemana_' + nomeAba);
      } else {
        Logger.log("Nenhum dado novo acumulado durante a semana na aba: " + nomeAba);
      }
    }
  }
}

// Obter novos dados da aba, a partir de uma linha específica
function obterDadosNovos(planilha, linhaInicial, numLinhas) {
  if (numLinhas <= 0) return []; // Verifica se o número de linhas é válido
  return planilha.getRange(linhaInicial, 1, numLinhas, planilha.getLastColumn()).getValues(); // Pega apenas as novas linhas
}

// Função para montar a tabela HTML
function montarTabela(dadosSemana) {
  return dadosSemana.map(linha => {
    linha[0] = Utilities.formatDate(new Date(linha[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy'); // Formata a data para PT-BR
    return `
      <tr>${linha.map(celula => `<td style="text-align: center;">${celula}</td>`).join('')}</tr>
    `;
  }).join('');
}

// Função para obter o dia da semana por extenso
function obterDiaDaSemana(data) {
  const dias = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
  return dias[data.getDay()];
}
