/*
Esse script coleta todos os desafios do Classroom e registra a turma, email e o horário de entrega.
É útil para validar dados dos estudantes
*/

function gerarRelatorioDesafios() {
  // --- CONFIGURAÇÃO ---
  // Substitua pelos IDs reais das suas turmas
  const courseIds = ["ID Turma 01", "ID Turma 02"]; 
  const termoBusca = "Desafio"; // Palavra-chave alterada
  
  // Cria a planilha com data no nome
  const ss = SpreadsheetApp.create("Relatório Desafios - " + new Date().toLocaleDateString());
  const sheet = ss.getActiveSheet();
  
  // Cabeçalho atualizado para o novo objetivo
  sheet.appendRow([
    "Turma", 
    "Título da Atividade (Classroom)", 
    "Título do Formulário", 
    "Email do Respondente", 
    "Data da Resposta"
  ]);

  Logger.log("Iniciando busca por Desafios...");

  // Itera sobre as turmas
  courseIds.forEach((idTurma, index) => {
    let nomeTurma = "Turma " + (index + 1);

    try {
      // 1. Busca os trabalhos da turma
      // O parâmetro 'orderBy' ajuda a trazer os mais recentes primeiro (opcional)
      const response = Classroom.Courses.CourseWork.list(idTurma);
      const courseWorkList = response.courseWork;

      if (!courseWorkList) {
        Logger.log(`Nenhuma atividade encontrada na ${nomeTurma}`);
        return;
      }

      courseWorkList.forEach(work => {
        // 2. Filtra por título contendo "Desafio"
        if (work.title && work.title.includes(termoBusca) && work.materials) {
          
          // Verifica anexos (Materiais) procurando formulários
          work.materials.forEach(material => {
            if (material.form) {
              // Tenta extrair o ID do formulário
              let formId = material.form.formUrl.match(/[-\w]{25,}/);
              
              if (formId) {
                Logger.log(`Processando formulário da atividade: ${work.title}`);
                
                try {
                  // --- CHAMA A FUNÇÃO DE EXTRAÇÃO ---
                  const dadosFormulario = extrairEmailsDoForm(formId.toString());
                  
                  // Escreve os resultados na planilha
                  if (dadosFormulario.respondentes.length > 0) {
                    dadosFormulario.respondentes.forEach(aluno => {
                      sheet.appendRow([
                        nomeTurma,
                        work.title,                // Título no Classroom
                        dadosFormulario.titulo,    // Título do Formulário
                        aluno.email,               // Email do aluno
                        aluno.data                 // Data da resposta
                      ]);
                    });
                  } else {
                    // Caso o form exista mas ninguém tenha respondido ainda
                    sheet.appendRow([
                      nomeTurma, 
                      work.title, 
                      dadosFormulario.titulo, 
                      "Sem respostas", 
                      "-"
                    ]);
                  }
                  
                } catch (erroForm) {
                  Logger.log(`Erro ao ler form ${material.form.title}: ${erroForm.message}`);
                  sheet.appendRow([nomeTurma, work.title, "Erro ao acessar Form", erroForm.message, "-"]);
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
  
  Logger.log(`Concluído! Planilha gerada: ${ss.getUrl()}`);
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