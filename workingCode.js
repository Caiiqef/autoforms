function myFunction() {
    var spreadsheet = SpreadsheetApp.getActive();
    var responseSheet = spreadsheet.getSheetByName('Abertura de Chamados');
    // Returns the position of the last row that has content.
    var rLastRow = responseSheet.getLastRow();
    // Returns the position of the last column that has content.
    var lastCol = responseSheet.getLastColumn();
  
    var values = responseSheet.getRange(rLastRow, 1, 1, lastCol).getValues()[0];
  
    var horaChamado = Utilities.formatDate(new Date(values[0]), 'GMT-3', 'dd/MM/yyyy HH:mm a');
    var idComputador = values[1];
    var tipoChamado = values[2];
    var prioridadeChamado = values[3];
    var descricaoChamado = values[4];
    var nome = values[5];
    var departamento = values[6];
    var email = values[7];
    
    emailAdmin(horaChamado, idComputador, tipoChamado, prioridadeChamado, descricaoChamado, nome, departamento, email);
    emailOp(horaChamado, idComputador, tipoChamado, prioridadeChamado, descricaoChamado, nome, departamento, email);
  }
  
  
  function emailAdmin(horaChamado, idComputador, tipoChamado, prioridadeChamado, descricaoChamado, nome, departamento, email) {
    var assunto = 'Chamado: '  + idComputador + ': ' + tipoChamado;
  
    // Email Text. You can add HTML code here - see ctrlq.org/html-mail
    var htmlBody = 'Prezado operador,';
    htmlBody += '<p>Essa é uma notificação automática do portal de chamados da Aprisco.</p>';
    htmlBody += '<strong>Nome:</strong> ' + nome;
    htmlBody += '<br><strong>Departamento:</strong> ' + departamento;
    htmlBody += '<br><strong>ID do Computador:</strong> ' + idComputador;
    htmlBody += '<br><strong>Email:</strong> ' + email + '</p>';
    htmlBody += '<br><strong>Tipo do Chamado:</strong> ' + tipoChamado;
    htmlBody += '<br><strong>Data:</strong> ' + horaChamado;
    htmlBody += '<br><strong>Prioridade:</strong> ' + prioridadeChamado;
    htmlBody += '<br><strong>Descrição do problema:</strong> ' + descricaoChamado;  
    
    htmlBody += '<p>Obrigado!</p>';
    htmlBody += '<p>Cordialmente,<br>Aprisco Soluções Empresariais.</p>';
    
    GmailApp.sendEmail('suporte02@aprisco.cnt.br', assunto, '', {htmlBody:htmlBody, name: 'Aprisco - Portal de Chamados', replyTo: email});
  }
  
  function emailOp(horaChamado, idComputador, tipoChamado, prioridadeChamado, descricaoChamado, nome, departamento, email) {
    var assunto = 'Chamado: '  + idComputador + ':' + tipoChamado;
  
    // Email Text. You can add HTML code here - see ctrlq.org/html-mail
    var htmlBody = 'Prezado(a) ' + nome + ',';
    htmlBody += '<p>Obrigado por entrar em contato com a equipe de TI.';
    htmlBody += '<br>Um chamado foi aberto com a sua requisição!';
    htmlBody += '<br>Você será notificado quando uma resposta for feita por e-mail.';
    htmlBody += '<br>Os detalhes do seu chamado são mostrados abaixo.</p>';
    htmlBody += '<br><strong>Nome:</strong> ' + nome;
    htmlBody += '<br><strong>Departamento:</strong> ' + departamento;
    htmlBody += '<br><strong>ID do Computador:</strong> ' + idComputador;
    htmlBody += '<br><strong>Email:</strong> ' + email;
    htmlBody += '<br><strong>Tipo do Chamado:</strong> ' + tipoChamado;
    htmlBody += '<br><strong>Data:</strong> ' + horaChamado;
    htmlBody += '<br><strong>Prioridade:</strong> ' + prioridadeChamado;
    htmlBody += '<br><strong>Descrição:</strong> ' + descricaoChamado;  
    htmlBody += '<p>Obrigado!</p>';
    htmlBody += '<p>Cordialmente,<br>Equipe de TI - Aprisco Soluções Empresariais.</p>';
    htmlBody += '<p><strong>NÃO responda a este e-mail, pois esta caixa de correio não é utilizada. Em vez disso, use "Responder a todos".</p></strong>';
    
    GmailApp.sendEmail(email, assunto, '', {htmlBody:htmlBody, name: 'Aprisco - Portal de Chamados', replyTo: 'suporte02@aprisco.cnt.br'});
  }
  
