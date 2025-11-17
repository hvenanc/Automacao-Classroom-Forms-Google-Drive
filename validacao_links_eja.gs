function gerarRelatorioMiniprojetos() {
  // --- CONFIGURAÇÃO ---
  // Coloque aqui os IDs das suas duas turmas
  const courseIds = ["A", "B"]; 
  const termoBusca = "Miniprojeto";
  
  // Cria a planilha
  const ss = SpreadsheetApp.create("Relatório Miniprojetos - " + new Date().toLocaleDateString());
  const sheet = ss.getActiveSheet();
  // Cabeçalho da planilha
  sheet.appendRow(["ID da Turma", "Título do Trabalho", "Email/Aluno", "Status do Link", "Link Enviado"]);

  Logger.log("Iniciando busca nas turmas...");

  // Itera sobre as turmas
  courseIds.forEach((idTurma, index) => {


    let nomeTurma = "";
    if (index === 0) {
      nomeTurma = "Turma 01";
    } else if (index === 1) {
      nomeTurma = "Turma 02";
    } else {
      nomeTurma = "Turma " + (index + 1); // Caso adicione mais turmas futuramente
    }

    try {
      // Busca os trabalhos da turma
      const response = Classroom.Courses.CourseWork.list(idTurma);
      const courseWorkList = response.courseWork;

      if (!courseWorkList) return;

      courseWorkList.forEach(work => {
        // Filtra por título "Miniprojeto"
        if (work.title && work.title.includes(termoBusca) && work.materials) {
          
          // Verifica anexos (Materiais)
          work.materials.forEach(material => {
            if (material.form) {
              const formId = material.form.formUrl.match(/[-\w]{25,}/); // Extrai ID da URL ou usa material.form.id se disponível
              
              if (formId) {
                Logger.log(`Processando formulário: ${work.title}`);
                
                try {
                  // APLICA A SUA FUNÇÃO
                  // Nota: Ajuste os títulos das perguntas na chamada se necessário
                  const resultados = testaLinks(formId.toString());
                  
                  // Escreve os resultados na planilha
                  resultados.forEach(aluno => {
                    sheet.appendRow([
                      nomeTurma,
                      work.title,
                      aluno.nome,
                      aluno.link, // "Presente", "Link Privado" ou "Sem link"
                      aluno.linkOriginal // Adicionei isso no objeto de retorno para referência
                    ]);
                  });
                  
                } catch (erroForm) {
                  Logger.log(`Erro ao processar form ${material.form.title}: ${erroForm.message}`);
                }
              }
            }
          });
        }
      });
    } catch (e) {
      Logger.log(`Erro na turma ${idTurma}: ${e.message}`);
    }
  });
  
  Logger.log(`Concluído! Planilha gerada: ${ss.getUrl()}`);
}

// --- SUA FUNÇÃO MODIFICADA LIGEIRAMENTE PARA RETORNAR O LINK ORIGINAL ---
function testaLinks(
  formId,
  nomeQuestionTitle = "E-mail (atenção adicionar e-mail recebido na sua escola)",
  linkQuestionTitle = "Adicione aqui o link do miniprojeto"
) {
  const form = FormApp.openById(formId);
  const formResponses = form.getResponses();
  const studentData = [];

  for (const formResponse of formResponses) {
    const itemResponses = formResponse.getItemResponses();
    let nome = "";
    let link = "";

    for (const itemResponse of itemResponses) {
      // O Trim deve ser adicionado devido aos espaços deixados nas perguntas
      const questionTitle = itemResponse.getItem().getTitle().trim();
      const response = itemResponse.getResponse();

      // Tenta pegar o email. 
      // OBS: Se o formulário coleta e-mails automaticamente, formResponse.getRespondentEmail() 
      // funciona mesmo fora deste if. Mas mantive sua lógica original.
      if (questionTitle === nomeQuestionTitle) {
        nome = formResponse.getRespondentEmail() || response || "Email não identificado";
      }
      if (questionTitle === linkQuestionTitle) {
        link = response || "";
      }
    }

    // Se o nome ficou vazio (talvez título diferente), tenta pegar o email automático do formulário
    if (nome === "") {
       nome = formResponse.getRespondentEmail() || "Anônimo";
    }

    studentData.push({
      nome,
      link: link ? (verificaLinkPublico(link) ? "Presente (Acessível)" : "Link Privado/Erro") : "Sem link",
      linkOriginal: link
    });
  }

  // Ordena o array pelo nome
  studentData.sort((a, b) => a.nome.localeCompare(b.nome));

  return studentData;
}

// --- FUNÇÃO AUXILIAR QUE FALTAVA ---
function verificaLinkPublico(url) {
  if (!url) return false;
  try {
    // Tenta acessar o link. Se der erro (404, 403) vai pro catch.
    // muteHttpExceptions: true impede que o script pare se der erro 404
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    return response.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}