/** Criado por Renato Oliveira Batista da Silveira - ren.oliv87@gmail.com **/
/** Envia e-mail ao usuário informando que a entrega expirou, com base na diferença da data do envio vs prazo. **/
function deliveryExpired() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("2020");
  var startRow = 2;
  var numRows = 3000;
  var dataRange = sheet.getRange(startRow, 1 , numRows, 22);
  var data = dataRange.getValues();
  var html1 = HtmlService.createTemplateFromFile('template-expirada1').evaluate().getContent();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailEntregaExpirada = row[20];
    var emailProdutoChegou = row[21];
    var html2 = HtmlService.createTemplateFromFile('template-expirada2').evaluate().getContent(); 
    var tell = "<p align='center'><font size='5'><strong>Olá "+row[3]+",</strong></font></p>";
    var date = row[16];
    var prazo = row[15];
    var d = new Date();
    var t = new Date(date);
    var t1 = d.getTime();
    var t2 = t.getTime();
    
    var diff = Math.floor((t1-t2)/(24*3600*1000));
    
    if (emailEntregaExpirada == "Enviado" || emailProdutoChegou == "Enviado") {
    
      continue;
    } 
    
    if (prazo == "7 dias" && diff == 7.0) {
      
      var emailAddress = row[4];
      var message ="<p align='center'><font size='4'><strong>Chamado:</strong> "+row[0]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Produto:</strong> "+row[5]+"</font></p>";
      
      MailApp.sendEmail(emailAddress,"Seu produto chegará em breve! - Chamado "+row[0], "", {htmlBody: html1+tell+html2+message , noReply: true, name: "Seu produto chegará em breve! - Chamado "+row[0]});
      sheet.getRange(startRow + i, 21).setValue("Enviado");
      SpreadsheetApp.flush();
    }
    
     if (prazo == "10 dias" && diff == 10.0) {
      
      var emailAddress = row[4];
      var message ="<p align='center'><font size='4'><strong>Chamado:</strong> "+row[0]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Produto:</strong> "+row[5]+"</font></p>";
      
      MailApp.sendEmail(emailAddress,"Seu produto chegará em breve! - Chamado "+row[0], "", {htmlBody: html1+tell+html2+message , noReply: true, name: "Seu produto chegará em breve! - Chamado "+row[0]});
      sheet.getRange(startRow + i, 21).setValue("Enviado");
      SpreadsheetApp.flush();
    }
    
     if (prazo == "15 dias" && diff == 15.0) {
      
      var emailAddress = row[4];
      var message ="<p align='center'><font size='4'><strong>Chamado:</strong> "+row[0]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Produto:</strong> "+row[5]+"</font></p>";
      
      MailApp.sendEmail(emailAddress,"Seu produto chegará em breve! - Chamado "+row[0], "", {htmlBody: html1+tell+html2+message , noReply: true, name: "Seu produto chegará em breve! - Chamado "+row[0]});
      sheet.getRange(startRow + i, 21).setValue("Enviado");
      SpreadsheetApp.flush();
    }
    
     if (prazo == "17 dias" && diff == 17.0) {
      
      var emailAddress = row[4];
      var message ="<p align='center'><font size='4'><strong>Chamado:</strong> "+row[0]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Produto:</strong> "+row[5]+"</font></p>";
      
      MailApp.sendEmail(emailAddress,"Seu produto chegará em breve! - Chamado "+row[0], "", {htmlBody: html1+tell+html2+message , noReply: true, name: "Seu produto chegará em breve! - Chamado "+row[0]});
      sheet.getRange(startRow + i, 21).setValue("Enviado");
      SpreadsheetApp.flush();
    }
    
     if (prazo == "30 dias" && diff == 30.0) {
      
      var emailAddress = row[4];
      var message ="<p align='center'><font size='4'><strong>Chamado:</strong> "+row[0]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Produto:</strong> "+row[5]+"</font></p>";
      
      MailApp.sendEmail(emailAddress,"Seu produto chegará em breve! - Chamado "+row[0], "", {htmlBody: html1+tell+html2+message , noReply: true, name: "Seu produto chegará em breve! - Chamado "+row[0]});
      sheet.getRange(startRow + i, 21).setValue("Enviado");
      SpreadsheetApp.flush();
    }
  } 
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}