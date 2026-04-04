/**
 * Envia um alerta de conclusão de processo por e-mail via GmailApp.
 * Otimizado para evitar filtros de Spam.
 * @param {string} tipo - O nome do processo (ex: 'Relatório de Desafios')
 * @param {string} url - Opcional. O link para a planilha gerada.
 */
function enviarAlertaConclusao(tipo, url = null) {
  const emailUsuario = Session.getActiveUser().getEmail();
  const nomeUsuario = emailUsuario.split('@')[0];
  const assunto = `Relatório Concluído: ${tipo}`;

  // Texto puro (Fallback crucial para não ser spam)
  const corpoTexto = `Olá, ${nomeUsuario}.\n\nO processo "${tipo}" foi finalizado com sucesso.\n\n${url ? `Acesse os dados em: ${url}` : 'Os dados estão na planilha principal.'}\n\nAtenciosamente,\nScript de Monitoramento Classroom.`;

  // HTML estruturado
  const htmlCorpo = `
    <div style="font-family: sans-serif; max-width: 600px; margin: auto; border: 2px solid #d35400; padding: 20px;">
      <h2 style="color: #2c3e50;">Olá, ${nomeUsuario}!</h2>
      <p>Informamos que a automação de <strong>${tipo}</strong> concluiu sua tarefa com êxito.</p>
      
      ${url ? `
        <div style="margin: 20px 0;">
          <a href="${url}" style="background-color: #e67e22; color: white; padding: 12px 25px; text-decoration: none; border-radius: 4px; display: inline-block;">Acessar Planilha</a>
        </div>
        <p style="font-size: 12px; color: #666;">Se o botão não funcionar, copie este link: <br> ${url}</p>
      ` : '<p>Os dados já estão disponíveis na sua planilha.</p>'}
      
      <br>
      <hr style="border: 0; border-top: 1px solid #eee;">
      <p style="font-size: 11px; color: #999;">
        <strong>Identificador de Segurança:</strong> Enviado via Google Apps Script por Henrique Venâncio.<br>
        Este e-mail é um alerta automático de monitoramento para a conta ${emailUsuario}.
      </p>
    </div>
  `;

  try {
    // Usando GmailApp com estrutura de objeto completa
    GmailApp.sendEmail(emailUsuario, assunto, corpoTexto, {
      htmlBody: htmlCorpo,
      name: "Automação Classroom (Henrique Venâncio)", // Nome amigável no remetente
      noReply: false // Se for Workspace, ajuda a dar veracidade
    });
    
    console.log(`E-mail de conclusão enviado para: ${emailUsuario}`);
  } catch (e) {
    console.error("Falha ao enviar e-mail: " + e.toString());
  }
}