function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function enviarContato(emailRecrutador, mensagemRecrutador) {

  const sheet = SpreadsheetApp
    .openById("1nzqSDS9AaLeYxtI_y14-ZsHiKhAyczn6HsbxXn3-y0c")
    .getSheetByName("Page");

  emailRecrutador = emailRecrutador.trim();
  mensagemRecrutador = mensagemRecrutador.trim();

  if (!emailRecrutador) {
    return "Ops, alguma coisa de errado com seu e-mail digitado.";
  }

  if (!mensagemRecrutador) {
    return "Não esqueça da mensagem";
  }

  const agora = new Date();

  const data = Utilities.formatDate(
    agora,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  const hora = Utilities.formatDate(
    agora,
    Session.getScriptTimeZone(),
    "HH:mm:ss"
  );

  const mensagem =
`Tenha uma excelente experiencia.


Mensagem:
${mensagemRecrutador}

Data:
${data}
Hora:
${hora}`;

  try {

    sheet.appendRow([
      emailRecrutador,
      mensagemRecrutador,
      data,
      hora
    ]);

    MailApp.sendEmail({
       to: "noreply@gmail.com",
       bcc: emailRecrutador,
       subject: "Quick E-mail de Otávio Félix",
       body: mensagem
    });

    return "Mensagem enviada com sucesso! Não esqueça de conferir sua caixa de Spam";

  } catch (erro) {

    return "Erro ao enviar e-mail: " + erro.message;

  }

}