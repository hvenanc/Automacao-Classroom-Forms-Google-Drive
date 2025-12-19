/*
Responsável pelo carregamento da Página HTML do Macro, além da seleção das turmas e atualização das informações.
*/

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Frequência dos Alunos');
}

function getAlunosData(turma) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(turma);
  if (!aba) return [];

  const dados = aba.getDataRange().getValues();
  const cabecalho = dados.shift(); // remove o cabeçalho
  const alunos = [];

  dados.forEach(linha => {
    // Últimas duas colunas são as frequências individuais
    const freqDesafios = linha[linha.length - 2];
    const freqMiniprojetos = linha[linha.length - 1];

    const aluno = {
      nome: linha[0],
      email: linha[1],
      freqDesafios: freqDesafios,
      freqMiniprojetos: freqMiniprojetos,
      detalhes: []
    };

    // Percorre colunas de atividades (ignorando as duas últimas)
    for (let i = 2; i < linha.length - 2; i += 2) {
      // Dia 1 - desafios
      if (linha[i] !== 'Presente') {
        aluno.detalhes.push({
          tema: cabecalho[i].split(' - ')[0],
          dia: 'Dia 1 (Desafios)',
          situacao: linha[i]
        });
      }

      // Dia 2 - miniprojetos
      if (linha[i + 1] !== 'Presente') {
        aluno.detalhes.push({
          tema: cabecalho[i + 1].split(' - ')[0],
          dia: 'Dia 2 (Miniprojeto)',
          situacao: linha[i + 1]
        });
      }
    }

    alunos.push(aluno);
  });

  return alunos;
}

function criarTriggerRegistrarSituacao(nomeAbaConfirmada) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = planilha.getSheets().map(s => s.getName());

  if (!abas.includes(nomeAbaConfirmada)) {
    return "Erro: Nome da aba incorreta ou não encontrada!";
  }

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'registrarSituacaoPorTemas') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('registrarSituacaoPorTemas')
           .timeBased()
           .after(1000)
           .create();

  return "Solicitação de atualização criada com sucesso!";
}
