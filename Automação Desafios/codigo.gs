/**
 * Painel de Controle - Automação Classroom + Forms
 * Funções de servidor para gerenciar turmas, formulários e gatilhos.
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Gestor de Triggers Classroom')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Busca as turmas onde o usuário atual é o professor
function buscarTurmas() {
  try {
    const cursos = Classroom.Courses.list({ teacherId: 'me' }).courses;
    return cursos ? cursos.map(c => ({ nome: c.name, id: c.id })) : [];
  } catch (e) {
    throw new Error("Erro ao carregar turmas: " + e.message);
  }
}

// Busca formulários anexados e verifica status do monitoramento
function buscarFormsDaTurma(courseId) {
  const listaForms = [];
  const courseWork = Classroom.Courses.CourseWork.list(courseId).courseWork;
  
  // Lista IDs de formulários que já possuem gatilhos neste projeto
  const triggersAtivos = ScriptApp.getProjectTriggers().map(t => t.getTriggerSourceId());

  if (courseWork) {
    courseWork.forEach(work => {
      if (work.materials) {
        work.materials.forEach(m => {
          if (m.form) {
            const formUrl = m.form.formUrl;
            const match = formUrl.match(/[-\w]{25,}/);
            if (match) {
              const formId = match[0];
              listaForms.push({
                titulo: work.title || m.form.title,
                id: formId,
                monitorado: triggersAtivos.includes(formId)
              });
            }
          }
        });
      }
    });
  }
  return listaForms;
}

// Configura ou Reinicia o trigger para um formulário específico
function adicionarTriggerAoForm(formId) {
  try {
    const form = FormApp.openById(formId);
    
    // Remove gatilhos existentes para este formulário para evitar duplicidade
    const triggers = ScriptApp.getProjectTriggers();
    let deletados = 0;
    triggers.forEach(t => {
      if (t.getTriggerSourceId() === formId) {
        ScriptApp.deleteTrigger(t);
        deletados++;
      }
    });

    // Cria o novo gatilho vinculado à função de processamento
    ScriptApp.newTrigger('processarResposta')
      .forForm(form)
      .onFormSubmit()
      .create();

    return {
      sucesso: true,
      mensagem: deletados > 0 ? "Monitoramento reiniciado!" : "Monitoramento ativado com sucesso!"
    };
  } catch (e) {
    return {
      sucesso: false,
      mensagem: "Erro: Você precisa ser o proprietário do formulário. " + e.message
    };
  }
}

// Função executada automaticamente pelo gatilho ao receber resposta
function processarResposta(e) {
  const form = e.source;
  const response = e.response;
  const itemResponses = response.getItemResponses();
  let todasCorretas = true;

  itemResponses.forEach((itemResponse) => {
    try {
      const item = itemResponse.getItem();
      // Foca em questões de múltipla escolha que possuem gabarito
      const pergunta = item.asMultipleChoiceItem();
      const choices = pergunta.getChoices();
      const respostaCorreta = choices.find(c => c.isCorrectAnswer());

      if (respostaCorreta) {
        if (itemResponse.getResponse() !== respostaCorreta.getValue()) {
          todasCorretas = false;
        }
      }
    } catch (err) {
      // Ignora itens sem gabarito ou tipos de perguntas não suportados
    }
  });

  // Se o aluno errou qualquer questão com gabarito, deleta a resposta
  if (!todasCorretas) {
    form.deleteResponse(response.getId());
  }
}