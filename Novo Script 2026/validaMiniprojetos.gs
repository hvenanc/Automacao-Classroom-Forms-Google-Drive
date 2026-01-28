/*
Gerador de Relatório de Miniprojetos integrado ao seletor de turmas.
Este script substitui os IDs fixos pelos IDs selecionados no seu Dashboard.
*/

function gerarRelatorioMiniprojetos() {
  const prop = PropertiesService.getUserProperties();
  // Recupera os IDs que você marcou nos checkboxes do Dashboard
  const idsSelecionados = JSON.parse(prop.getProperty('TURMAS_PARA_ATUALIZAR') || '[]');
  const termoBusca = "Miniprojeto";
  
  if (idsSelecionados.length === 0) {
    Logger.log("Nenhuma turma selecionada no Dashboard.");
    return "Erro: Selecione as turmas no painel antes de gerar o relatório.";
  }

  // Cria a planilha no Drive do usuário
  const nomePlanilha = "Relatório de Links - Miniprojetos - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const ss = SpreadsheetApp.create(nomePlanilha);
  const sheet = ss.getActiveSheet();
  
  // Cabeçalho profissional
  sheet.appendRow(["ID da Turma", "Nome da Turma", "Título do Trabalho", "Email/Aluno", "Status do Link", "Link Enviado"]);
  sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#f3f3f3");

  Logger.log("Iniciando busca nas turmas selecionadas...");

  // Busca todas as turmas para poder pegar o Nome Real de cada ID
  const todasTurmasClassroom = Classroom.Courses.list().courses || [];

  idsSelecionados.forEach(idTurma => {
    // Encontra o objeto da turma para saber o nome real
    const dadosTurma = todasTurmasClassroom.find(t => t.id === idTurma);
    const nomeTurma = dadosTurma ? dadosTurma.name : "ID: " + idTurma;

    try {
      // Busca os trabalhos da turma específica
      const response = Classroom.Courses.CourseWork.list(idTurma);
      const courseWorkList = response.courseWork;

      if (!courseWorkList) return;

      courseWorkList.forEach(work => {
        // Filtra pelo termo (ex: "Miniprojeto")
        if (work.title && work.title.toLowerCase().includes(termoBusca.toLowerCase()) && work.materials) {
          
          work.materials.forEach(material => {
            if (material.form) {
              // Extrai o ID do formulário da URL
              const formIdMatch = material.form.formUrl.match(/[-\w]{25,}/);
              
              if (formIdMatch) {
                const formId = formIdMatch[0];
                Logger.log(`Processando: ${nomeTurma} > ${work.title}`);
                
                try {
                  // Chama a função de teste de links
                  const resultados = testaLinks(formId);
                  
                  // Escreve os resultados na planilha
                  resultados.forEach(aluno => {
                    sheet.appendRow([
                      idTurma,
                      nomeTurma,
                      work.title,
                      aluno.nome,
                      aluno.link, 
                      aluno.linkOriginal
                    ]);
                  });
                  
                } catch (erroForm) {
                  Logger.log(`Erro no form ${work.title}: ${erroForm.message}`);
                }
              }
            }
          });
        }
      });
    } catch (e) {
      Logger.log(`Erro ao acessar Classroom para a turma ${idTurma}: ${e.message}`);
    }
  });
  
  // Ajusta a largura das colunas automaticamente
  sheet.autoResizeColumns(1, 6);
  
  const finalUrl = ss.getUrl();
  Logger.log(`Concluído! Planilha gerada: ${finalUrl}`);
  return "Relatório gerado com sucesso: " + finalUrl;
}

/**
 * Analisa as respostas de um formulário específico
 */
function testaLinks(formId) {
  // Títulos das perguntas (Ajuste se o seu formulário usar termos diferentes)
  const NOME_PERGUNTA = "E-mail"; 
  const LINK_PERGUNTA = "link"; 

  const form = FormApp.openById(formId);
  const formResponses = form.getResponses();
  const studentData = [];

  for (const formResponse of formResponses) {
    const itemResponses = formResponse.getItemResponses();
    let nomeEncontrado = "";
    let linkEncontrado = "";

    for (const itemResponse of itemResponses) {
      const title = itemResponse.getItem().getTitle().toLowerCase();
      const response = itemResponse.getResponse();

      if (title.includes(NOME_PERGUNTA.toLowerCase())) {
        nomeEncontrado = response;
      }
      if (title.includes(LINK_PERGUNTA.toLowerCase())) {
        linkEncontrado = response;
      }
    }

    // Fallback: se não achou e-mail no corpo, pega o e-mail de quem logou no form
    const emailFinal = nomeEncontrado || formResponse.getRespondentEmail() || "Anônimo";

    studentData.push({
      nome: emailFinal,
      link: linkEncontrado ? (verificaLinkPublico(linkEncontrado) ? "✅ Acessível" : "❌ Privado/Inválido") : "Sem envio",
      linkOriginal: linkEncontrado || ""
    });
  }

  return studentData.sort((a, b) => a.nome.localeCompare(b.nome));
}

/**
 * Testa se o link responde com código 200 (Público)
 */
function verificaLinkPublico(url) {
  if (!url || !url.startsWith("http")) return false;
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    return response.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}