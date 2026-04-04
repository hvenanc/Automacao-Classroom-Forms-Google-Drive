/*
Relatório de Desafios integrado ao Dashboard.
Busca os dados apenas das turmas selecionadas no Check Area.
*/

function gerarRelatorioDesafios() {
  const prop = PropertiesService.getUserProperties();
  const idsSelecionados = JSON.parse(prop.getProperty('TURMAS_PARA_ATUALIZAR') || '[]');
  const termoBusca = "Desafio"; 
  
  if (idsSelecionados.length === 0) {
    return "Erro: Selecione as turmas no painel antes de gerar o relatório.";
  }

  // Cria a planilha no Drive do usuário com data e hora
  const nomePlanilha = "Relatório Desafios - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const ss = SpreadsheetApp.create(nomePlanilha);
  const sheet = ss.getActiveSheet();
  
  sheet.appendRow([
    "ID da Turma",
    "Nome da Turma", 
    "Título da Atividade (Classroom)", 
    "Título do Formulário", 
    "Email do Respondente", 
    "Data da Resposta"
  ]);
  sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e0e0e0");

  Logger.log("Iniciando busca por Desafios nas turmas selecionadas...");

  const todasTurmas = Classroom.Courses.list().courses || [];

  idsSelecionados.forEach(idTurma => {
    const dadosTurma = todasTurmas.find(t => t.id === idTurma);
    const nomeTurma = dadosTurma ? dadosTurma.name : "Turma Desconhecida";

    try {
      const response = Classroom.Courses.CourseWork.list(idTurma);
      const courseWorkList = response.courseWork;

      if (!courseWorkList) return;

      courseWorkList.forEach(work => {
        // Filtra por título contendo "Desafio"
        if (work.title && work.title.toLowerCase().includes(termoBusca.toLowerCase()) && work.materials) {
          
          work.materials.forEach(material => {
            if (material.form) {
              let formIdMatch = material.form.formUrl.match(/[-\w]{25,}/);
              
              if (formIdMatch) {
                try {
                  const dadosFormulario = extrairEmailsDoForm(formIdMatch[0].toString());
                  
                  if (dadosFormulario.respondentes.length > 0) {
                    dadosFormulario.respondentes.forEach(aluno => {
                      sheet.appendRow([
                        idTurma,
                        nomeTurma,
                        work.title,
                        dadosFormulario.titulo,
                        aluno.email,
                        aluno.data
                      ]);
                    });
                  } else {
                    sheet.appendRow([idTurma, nomeTurma, work.title, dadosFormulario.titulo, "Sem respostas", "-"]);
                  }
                } catch (erroForm) {
                  Logger.log(`Erro ao ler form ${work.title}: ${erroForm.message}`);
                }
              }
            }
          });
        }
      });
    } catch (e) {
      Logger.log(`Erro crítico na turma ${idTurma}: ${e.message}`);
    }
  });

  sheet.autoResizeColumns(1, 6);
  const urlFinal = ss.getUrl();
  // Chama a função do seu arquivo externo (notificacoes.gs)
  enviarAlertaConclusao("Relatório de Desafios", urlFinal);
  Logger.log(`Concluído! Planilha gerada: ${urlFinal}`);
  return urlFinal; // Retorna a URL para o frontend abrir
}

// --- FUNÇÃO ESPECÍFICA PARA PEGAR TÍTULO E EMAILS ---
function extrairEmailsDoForm(formId) {
  const form = FormApp.openById(formId);
  const formResponses = form.getResponses();
  
  const resultado = {
    titulo: form.getTitle(), // Pega o título oficial do formulário
    respondentes: []
  };

  // Itera sobre cada resposta individual
  for (const formResponse of formResponses) {
    // Tenta pegar o e-mail coletado automaticamente (padrão Google Classroom)
    let email = formResponse.getRespondentEmail();
    
    // Fallback: Se não coletou automático, tenta procurar uma pergunta que contenha "e-mail" ou "@"
    if (!email) {
       const itemResponses = formResponse.getItemResponses();
       for (const itemResponse of itemResponses) {
         const respostaTexto = itemResponse.getResponse();
         // Verifica se a resposta parece um email
         if (typeof respostaTexto === 'string' && respostaTexto.includes('@') && respostaTexto.includes('.')) {
           email = respostaTexto; 
           break; // Achou um email, para de procurar
         }
       }
    }

    resultado.respondentes.push({
      email: email || "Email não identificado (Anônimo)",
      data: formResponse.getTimestamp().toLocaleString() // Adiciona data para controle
    });
  }

  return resultado;
}