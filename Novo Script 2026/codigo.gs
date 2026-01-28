function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Monitoramento Classroom');
}

/**
 * Busca todas as turmas e inclui o status (Active, Archived, etc.)
 */
function getListaTurmas() {
  try {
    // Buscamos todas as turmas sem filtro inicial para processar no frontend
    const response = Classroom.Courses.list();
    const turmas = response.courses || [];
    
    return turmas.map(t => ({
      id: t.id,
      nome: t.name,
      status: t.courseState // 'ACTIVE', 'ARCHIVED', 'PROVISIONED' ou 'DECLINED'
    })).sort((a,b) => a.nome.localeCompare(b.nome));
  } catch (e) {
    Logger.log("Erro ao buscar turmas: " + e);
    return [];
  }
}

function prepararAtualizacao(idsTurmas, modalidade) {
  if (!idsTurmas || idsTurmas.length === 0) return "Erro: Nenhuma turma selecionada.";
  
  // SEGURANÇA: Garante que modalidade seja uma string e nunca nula/undefined
  const modalidadeValida = (modalidade && typeof modalidade === 'string') ? modalidade : "EJA";
  
  const prop = PropertiesService.getUserProperties();
  
  try {
    prop.setProperty('TURMAS_PARA_ATUALIZAR', JSON.stringify(idsTurmas));
    prop.setProperty('MODALIDADE_ENSINO', modalidadeValida); // Agora o valor é garantido 
  } catch (e) {
    return "Erro ao salvar propriedades: " + e.message;
  }

  // Gerenciamento de Triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'registrarSituacaoPorTemas') {
      ScriptApp.deleteTrigger(t); // Remove duplicatas [cite: 41]
    }
  });
  
  // Agenda a execução 
  ScriptApp.newTrigger('registrarSituacaoPorTemas')
           .timeBased()
           .after(1000)
           .create();

  return "Sincronização iniciada para " + modalidadeValida + "!";
}

function getAlunosData(nomeAbaCompleto) {
  try {
    const NOME_PLANILHA = "Monitoramento_Classroom_Pessoal";
    const arquivos = DriveApp.getFilesByName(NOME_PLANILHA);
    
    if (!arquivos.hasNext()) return [];
    
    const planilha = SpreadsheetApp.open(arquivos.next());
    const aba = planilha.getSheetByName(nomeAbaCompleto);
    
    if (!aba) {
      console.log("Aba não encontrada: " + nomeAbaCompleto);
      return [];
    }

    const dados = aba.getDataRange().getValues();
    if (dados.length < 2) return [];

    const cabecalho = dados.shift();
    
    return dados.map(linha => {
      // Mapeia os dados garantindo que as frequências sejam números
      return {
        nome: linha[0],
        email: linha[1],
        freqDesafios: !isNaN(linha[linha.length - 2]) ? Math.round(linha[linha.length - 2]) : 0,
        freqMiniprojetos: !isNaN(linha[linha.length - 1]) ? Math.round(linha[linha.length - 1]) : 0,
        pendencias: extrairPendencias(linha, cabecalho)
      };
    });
  } catch (e) {
    throw new Error("Erro no servidor: " + e.message);
  }
}

// Função auxiliar para limpar a lógica do map
function extrairPendencias(linha, cabecalho) {
  let pendencias = [];
  // Percorre da coluna 2 até a antepenúltima
  for (let i = 2; i < linha.length - 2; i++) {
    if (linha[i] !== 'Entregue' && linha[i] !== '') {
      pendencias.push({
        atividade: cabecalho[i],
        status: linha[i]
      });
    }
  }
  return pendencias;
}