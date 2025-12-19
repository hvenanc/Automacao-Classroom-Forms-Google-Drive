/*
Listas todas as turmas ativas do Classroom
*/

function listarMinhasTurmas() {
  const turmas = Classroom.Courses.list({teacherId: 'me', courseStates: ['ACTIVE']}).courses || [];
  turmas.forEach(turma => {
    Logger.log(`Turma: ${turma.name} | ID: ${turma.id}`);
  });
}
