/*
Coleta as atividades apenas das turmas selecionadas pelo usuário.
Grava os dados em uma planilha na conta do usuário logado.
*/

function registrarSituacaoPorTemas() {
  // Recupera as configurações salvas nas propriedades do usuário
  const prop = PropertiesService.getUserProperties();
  const modalidade = prop.getProperty('MODALIDADE_ENSINO') || 'EJA';
  const idsSelecionados = JSON.parse(prop.getProperty('TURMAS_PARA_ATUALIZAR') || '[]');

  // Critério dinâmico baseado na modalidade selecionada no Frontend
  const minimoDesafios = (modalidade === 'EM') ? 2 : 3;
  
  if (!idsSelecionados || idsSelecionados.length === 0) {
    Logger.log("Nenhuma turma selecionada para atualização.");
    return;
  }
  
  let turmasResponse;
  try {
    // Busca todas as turmas para encontrar os objetos correspondentes aos IDs selecionados
    turmasResponse = Classroom.Courses.list().courses || [];
  } catch (e) {
    Logger.log("Erro ao listar turmas do Classroom: " + e);
    return;
  }

  // Filtra apenas as turmas que estão na lista de IDs selecionados no Frontend
  const TURMAS = turmasResponse.filter(t => idsSelecionados.includes(t.id));

  const FILTROS = ['Miniprojeto', 'Desafio'];
  const NOME_PLANILHA = "Monitoramento_Classroom_Pessoal";
  let planilha;
  
  // Localiza ou cria a planilha no Drive do usuário
  const arquivos = DriveApp.getFilesByName(NOME_PLANILHA);
  if (arquivos.hasNext()) {
    planilha = SpreadsheetApp.open(arquivos.next());
  } else {
    planilha = SpreadsheetApp.create(NOME_PLANILHA);
  }

  TURMAS.forEach(turma => {
    // Identificador único da aba para evitar sobrescrever turmas de mesmo nome
    const nomeIdentificador = `${turma.name} (${turma.id})`;
    let aba = planilha.getSheetByName(nomeIdentificador);

    if (aba) {
      aba.clear();
    } else {
      aba = planilha.insertSheet(nomeIdentificador);
    }

    // Busca atividades da turma específica
    let atividades = Classroom.Courses.CourseWork.list(turma.id).courseWork || [];
    atividades = atividades.filter(a =>
      FILTROS.some(f => a.title && a.title.toLowerCase().includes(f.toLowerCase())) && a.topicId
    );

    if (atividades.length === 0) {
      aba.appendRow(['Nenhuma atividade encontrada com os filtros:', FILTROS.join(', ')]);
      return;
    }

    // Agrupamento por temas (Tópicos)
    const temas = {};
    atividades.forEach(a => {
      if (!temas[a.topicId]) temas[a.topicId] = [];
      temas[a.topicId].push(a);
    });

    // Mapeamento de nomes de tópicos
    const topicos = Classroom.Courses.Topics.list(turma.id).topic || [];
    const mapaTopicos = {};
    topicos.forEach(t => { mapaTopicos[t.topicId] = t.name; });

    // Montagem do Cabeçalho
    const cabecalho = ['Nome do Aluno', 'Email'];
    const listaTopicosOrdenada = Object.keys(temas).map(id => ({
      id,
      nome: mapaTopicos[id] || 'Sem Tema',
      atividades: temas[id]
    })).sort((a, b) => a.nome.localeCompare(b.nome));

    listaTopicosOrdenada.forEach(t => {
      cabecalho.push(`${t.nome} - [DESAFIOS]`);
      cabecalho.push(`${t.nome} - [MINIPROJETO]`);
    });

    cabecalho.push('Frequência Desafios (%)');
    cabecalho.push('Frequência Miniprojetos (%)');
    aba.appendRow(cabecalho);

    // Pré-busca de Submissões para otimizar performance (evita múltiplas chamadas dentro do loop de alunos)
    listaTopicosOrdenada.forEach(t => {
      t.desafios = [];
      t.miniprojetos = [];
      t.atividades.forEach(atividade => {
        const formMaterial = atividade.materials && atividade.materials.find(m => m.form);
        const urlForm = formMaterial && formMaterial.form && formMaterial.form.formUrl;
        let submissions = [];
        try {
          submissions = Classroom.Courses.CourseWork.StudentSubmissions.list(turma.id, atividade.id).studentSubmissions || [];
        } catch (e) {
          submissions = [];
        }
        
        const dadosAtividade = { atividade, urlForm, submissions };
        if (/desafio/i.test(atividade.title || '')) {
          t.desafios.push(dadosAtividade);
        } else if (/miniprojeto/i.test(atividade.title || '')) {
          t.miniprojetos.push(dadosAtividade);
        }
      });
    });

    // Busca e ordena lista de alunos
    const alunos = Classroom.Courses.Students.list(turma.id).students || [];
    alunos.sort((a, b) => a.profile.name.fullName.localeCompare(b.profile.name.fullName));

    const totalTopicos = listaTopicosOrdenada.length;

    alunos.forEach(aluno => {
      const linha = [aluno.profile.name.fullName, aluno.profile.emailAddress];
      let presencasDesafios = 0;
      let presencasMiniprojetos = 0;

      listaTopicosOrdenada.forEach(topico => {
        // --- PROCESSAMENTO DESAFIOS ---
        let totalDesafiosEntregues = 0;
        let pendentesNestaAula = [];

        topico.desafios.forEach(desafio => {
          let entregue = false;
          if (desafio.urlForm) {
            entregue = alunoRespondeuFormularioFormsApp(aluno.profile.emailAddress, desafio.urlForm);
          } else {
            const entregaAluno = desafio.submissions.find(s => s.userId === aluno.userId);
            if (entregaAluno && (entregaAluno.state === 'TURNED_IN' || entregaAluno.state === 'RETURNED')) {
              entregue = true;
            }
          }
          
          if (entregue) totalDesafiosEntregues++;
          else pendentesNestaAula.push(desafio.atividade.title);
        });

        let situacaoDia1 = '';
        if (totalDesafiosEntregues >= minimoDesafios) {
          situacaoDia1 = 'Entregue';
          presencasDesafios++;
        } else if (totalDesafiosEntregues > 0) {
          situacaoDia1 = pendentesNestaAula.sort().join("\n");
        } else {
          situacaoDia1 = 'Falta';
        }
        linha.push(situacaoDia1);

        // --- PROCESSAMENTO MINIPROJETOS ---
        let miniprojetoEntregue = false;
        topico.miniprojetos.forEach(miniprojeto => {
          if (miniprojetoEntregue) return;
          if (miniprojeto.urlForm) {
            if (alunoRespondeuFormularioFormsApp(aluno.profile.emailAddress, miniprojeto.urlForm)) {
              miniprojetoEntregue = true;
            }
          } else {
            const entregaAluno = miniprojeto.submissions.find(s => s.userId === aluno.userId);
            if (entregaAluno && (entregaAluno.state === 'TURNED_IN' || entregaAluno.state === 'RETURNED')) {
              miniprojetoEntregue = true;
            }
          }
        });

        if (miniprojetoEntregue) {
          linha.push('Entregue');
          presencasMiniprojetos++;
        } else {
          linha.push('Falta');
        }
      });

      // Cálculo das frequências finais nas colunas da direita
      const freqD = totalTopicos > 0 ? Math.round((presencasDesafios / totalTopicos) * 100) : 0;
      const freqM = totalTopicos > 0 ? Math.round((presencasMiniprojetos / totalTopicos) * 100) : 0;

      linha.push(freqD, freqM);
      aba.appendRow(linha);
    });
  });

  // Grava a data da última atualização com sucesso
  const agora = new Date();
  const dataFormatada = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  prop.setProperty('ULTIMA_ATUALIZACAO', dataFormatada);

  // Limpa as propriedades de controle após o processamento
  prop.deleteProperty('MODALIDADE_ENSINO');
  prop.deleteProperty('TURMAS_PARA_ATUALIZAR');

  const urlFinal = planilha.getUrl();
  enviarAlertaConclusao("Relatório de Atividades", urlFinal);
}

/**
 * Verifica se o aluno respondeu ao formulário via FormApp
 */
function alunoRespondeuFormularioFormsApp(email, urlForm) {
  try {
    const form = FormApp.openByUrl(urlForm);
    const respostas = form.getResponses();
    for (const resposta of respostas) {
      const respondente = resposta.getRespondentEmail();
      if (respondente && respondente.toLowerCase() === (email || '').toLowerCase()) return true;
    }
  } catch (e) {
    // Retorna falso se não houver permissão ou o formulário não existir
  }
  return false;
}