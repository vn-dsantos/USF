# Automatização de Planilha 
Scripts para automatização de planilhas para envio de email com base em dados dad planilha
function SendEmail() {
  var app = SpreadsheetApp.getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveRange().getA1Notation(); //linha que esta selecionando
  var range = app.getRange(spreadsheet); //guarda valores da seleção
  var signups = app.getRange(spreadsheet).getValues(); 
  var ui = SpreadsheetApp.getUi();

  for (var x = 0; x < signups.length; x++) { 
    var data = signups[x];

    var professor = data[11];
    var email = data[12];
    var componente = data[10];
    var campus = data[1];
    var dia = data[26];
    var title = "Feedback Aula Realizada " + componente + " - " + dia;
    var aviso = "Atenção as regras e normas de utilização dos laboratórios abaixo, não se esqueçam de repassar as informações a seu alunado.";
    
    //HTML
    var htmlMessage = "<p> Prezado(a) professor(a), " + professor + ",</p>" +
                      "<p>Poderia por gentileza responder ao feedback referente a aula de " + componente + " no campus " + campus + ", realizada dia " + dia + "</p>" +
                      "<p>OBS: O Preenchimento do formulário não é obrigatório, mas é de extrema importância para melhoria continua das nossas atividades práticas.</p>" +
                      "<p>Segue Link: https://docs.google.com/forms/d/e/1FAIpQLSdi0srNIqaorLNnxKXAZHyRb3bFIoZFtWdosEXRVaWBC-CwjA/viewform <p>" +
                      "<p>Atenciosamente</p>" +
                      "<br>" +
                      "<b>" + aviso + "</b>" +
                      "<br>" +
                      "<div style='text-align: justify;'>" +
                      "<img src='https://drive.google.com/uc?id=15peNcl7SoEP75949IYTYxg0n5TCsP0NC' alt='Imagem' style='width: 900px; height: auto;'>" +
                      "</div>";
    
    // Enviar o e-mail
    GmailApp.sendEmail(email, title, "", {htmlBody: htmlMessage});
    ui.alert("Feedback de aula realizada enviado!");
    
    range.setBackground("#46bdc6"); 
  }
}

