/*
Coleta as atividades apenas das turmas selecionadas pelo usuário.
*/

function registrarSituacaoPorTemas() {
  // Recupera as turmas selecionadas salvas nas propriedades do script
  const prop = PropertiesService.getScriptProperties();
  const modalidade = prop.getProperty('MODALIDADE_ENSINO') || 'EJA'; // Padrão EJA
  const turmasSelecionadasJSON = prop.getProperty('TURMAS_PARA_ATUALIZAR');

  const minimoDesafios = (modalidade === 'EM') ? 2 : 3;
  
  if (!turmasSelecionadasJSON) {
    Logger.log("Nenhuma turma selecionada para atualização.");
    return;
  }
  
  const nomesParaFiltrar = JSON.parse(turmasSelecionadasJSON);

  let turmasResponse;
  try {
    // Busca todas as turmas para encontrar os IDs correspondentes aos nomes selecionados
    turmasResponse = Classroom.Courses.list().courses || [];
  } catch (e) {
    Logger.log("Erro ao listar turmas: " + e);
    return;
  }

  // Filtra apenas as turmas que estão na lista de selecionadas
  const TURMAS = turmasResponse
    .filter(t => nomesParaFiltrar.includes(t.name))
    .map(t => ({ id: t.id, nome: t.name }));

  const FILTROS = ['Miniprojeto', 'Desafio'];
  
  const NOME_PLANILHA = "Monitoramento_Classroom_Pessoal";
  let planilha;
  const arquivos = DriveApp.getFilesByName(NOME_PLANILHA);
  
  if (arquivos.hasNext()) {
    planilha = SpreadsheetApp.open(arquivos.next());
  } else {
    planilha = SpreadsheetApp.create(NOME_PLANILHA);
  }

  TURMAS.forEach(turma => {
    const nomeIdentificador = `${turma.nome} (${turma.id})`;
    let aba = planilha.getSheetByName(nomeIdentificador);

    if (aba) {
      aba.clear();
    } else {
      aba = planilha.insertSheet(nomeIdentificador);
    }

    let atividades = Classroom.Courses.CourseWork.list(turma.id).courseWork || [];
    atividades = atividades.filter(a =>
      FILTROS.some(f => a.title && a.title.toLowerCase().includes(f.toLowerCase())) && a.topicId
    );

    if (atividades.length === 0) {
      aba.appendRow(['Nenhuma atividade encontrada com os filtros:', FILTROS.join(', ')]);
      return;
    }

    // Agrupamento por temas
    const temas = {};
    atividades.forEach(a => {
      if (!temas[a.topicId]) temas[a.topicId] = [];
      temas[a.topicId].push(a);
    });

    const topicos = Classroom.Courses.Topics.list(turma.id).topic || [];
    const mapaTopicos = {};
    topicos.forEach(t => { mapaTopicos[t.topicId] = t.name; });

    // Cabeçalho e Lógica de Processamento (Igual ao código anterior)
    const cabecalho = ['Nome do Aluno', 'Email'];
    const listaTopicosOrdenada = Object.keys(temas).map(id => ({
      id,
      nome: mapaTopicos[id] || 'Sem Tema',
      atividades: temas[id]
    })).sort((a, b) => a.nome.localeCompare(b.nome));

    listaTopicosOrdenada.forEach(t => {
      cabecalho.push(`${t.nome} - Dia 1 (Desafios)`);
      cabecalho.push(`${t.nome} - Dia 2 (Miniprojeto)`);
    });

    cabecalho.push('Frequência Desafios (%)');
    cabecalho.push('Frequência Miniprojetos (%)');
    aba.appendRow(cabecalho);

    // Pré-busca
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
          Logger.log('Erro ao buscar StudentSubmissions para atividade ' + atividade.id + ': ' + e);
          submissions = [];
        }
        if (/desafio/i.test(atividade.title || '')) {
          t.desafios.push({ atividade, urlForm, submissions });
        } else if (/miniprojeto/i.test(atividade.title || '')) {
          t.miniprojetos.push({ atividade, urlForm, submissions });
        }
      });
    });

    // Alunos
    const alunos = Classroom.Courses.Students.list(turma.id).students || [];
    alunos.sort((a, b) => a.profile.name.fullName.localeCompare(b.profile.name.fullName));

    const totalTopicos = listaTopicosOrdenada.length;

    alunos.forEach(aluno => {
      const linha = [aluno.profile.name.fullName, aluno.profile.emailAddress];

      let presencasDesafios = 0;
      let presencasMiniprojetos = 0;
      let tituloDesafio = [];

      listaTopicosOrdenada.forEach(topico => {
        // DIA 1: desafios
        let totalDesafiosEntregues = 0;
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
          else tituloDesafio.push(desafio.atividade.title);
        });

        let situacaoDia1 = '';
        if (totalDesafiosEntregues >= minimoDesafios) {
          situacaoDia1 = 'Presente';
          presencasDesafios++;
        } else if (totalDesafiosEntregues > 0) {
          situacaoDia1 = tituloDesafio.sort().join("\n");
        } else {
          situacaoDia1 = 'Falta';
        }
        linha.push(situacaoDia1);

        // DIA 2: miniprojetos
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

        let situacaoDia2 = '';
        if (miniprojetoEntregue) {
          situacaoDia2 = 'Presente';
          presencasMiniprojetos++;
        } else {
          situacaoDia2 = 'Falta';
        }
        linha.push(situacaoDia2);
      });

      const freqDesafios = totalTopicos > 0 ? Math.round((presencasDesafios / totalTopicos) * 100) : 0;
      const freqMiniprojetos = totalTopicos > 0 ? Math.round((presencasMiniprojetos / totalTopicos) * 100) : 0;

      linha.push(freqDesafios);
      linha.push(freqMiniprojetos);
      aba.appendRow(linha);
    });
  });
  
  // Limpa a seleção após o processamento
  prop.deleteProperty('MODALIDADE_ENSINO');
  prop.deleteProperty('TURMAS_PARA_ATUALIZAR');
}

// Função auxiliar mantida
function alunoRespondeuFormularioFormsApp(email, urlForm) {
  try {
    const form = FormApp.openByUrl(urlForm);
    const respostas = form.getResponses();
    for (const resposta of respostas) {
      const respondente = resposta.getRespondentEmail();
      if (respondente && respondente.toLowerCase() === (email || '').toLowerCase()) return true;
    }
  } catch (e) { }
  return false;
}