/*
Coleta as atividades do Classroom e verifica as respostas com os respectivos formulários.

Gera uma planilha com nome, email dos estudantes cadastrados nas turmas. Cada tema do Classroom vai gerar
uma coluna na planilha.

Critérios:

Entregou os 3 Desafios + Miniprojetos (Presente)
Deixou de entregar um desses (A cécula recebe qual atividade ou atividades estão pendentes)
Não entregou nada: Falta
*/

function registrarSituacaoPorTemas() {
  const TURMAS = [
     { id: 'ID DA TURMA 01', nome: 'Turma 01' },
     { id: 'ID DA TURMA 02', nome: 'Turma 02' }
  ];
  const FILTROS = ['Miniprojeto', 'Desafio'];
  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  TURMAS.forEach(turma => {
    const nomeAba = turma.nome;
    let aba = planilha.getSheetByName(nomeAba);

    if (aba) {
      aba.clear();
    } else {
      aba = planilha.insertSheet(nomeAba);
    }

    let atividades = Classroom.Courses.CourseWork.list(turma.id).courseWork || [];
    atividades = atividades.filter(a =>
      FILTROS.some(f => a.title && a.title.toLowerCase().includes(f.toLowerCase())) && a.topicId
    );

    if (atividades.length === 0) {
      aba.appendRow(['Nenhuma atividade encontrada com os filtros:', FILTROS.join(', ')]);
      return;
    }

    // Agrupa atividades por tema (topicId)
    const temas = {};
    atividades.forEach(a => {
      if (!temas[a.topicId]) temas[a.topicId] = [];
      temas[a.topicId].push(a);
    });

    // Mapa de nomes de tópicos
    const topicos = Classroom.Courses.Topics.list(turma.id).topic || [];
    const mapaTopicos = {};
    topicos.forEach(t => {
      mapaTopicos[t.topicId] = t.name;
    });

    // Cabeçalho
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

    // NOVO: duas colunas separadas de frequência
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
        if (totalDesafiosEntregues >= 3) {
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
}

// Verifica se o aluno respondeu ao formulário
function alunoRespondeuFormularioFormsApp(email, urlForm) {
  try {
    const form = FormApp.openByUrl(urlForm);
    const respostas = form.getResponses();
    for (const resposta of respostas) {
      try {
        const respondente = resposta.getRespondentEmail();
        if (respondente && respondente.toLowerCase() === (email || '').toLowerCase()) {
          return true;
        }
      } catch (e) {
        // resposta sem email
      }
    }
  } catch (e) {
    Logger.log("Erro ao acessar formulário: " + e);
  }
  return false;
}
