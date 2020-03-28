function expReminder() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("2019");
  var startRow = 2;
  var numRows = 3000;
  var dataRange = sheet.getRange(startRow, 1 , numRows, 20);
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var status = row[14];
    var emailExpired = row[16];
    var localidade = row[6];
    var fornecedor = row[11];
    var emailAddress = ['rodrigo_medeiros_tsystems@whirlpool.com','rodrigo_duarte_T-Systems@whirlpool.com','policarpo_moser_T-Systems@whirlpool.com','renato_oliveira_batista_tsystems@whirlpool.com'];
    var date = row[17];
    var d = new Date();
    var t = new Date(date);
    var t1 = d.getTime();
    var t2 = t.getTime();
    
    var diff = Math.floor((t1-t2)/(24*3600*1000));
    
   
/** Prodata USP **/
    
    if (localidade == "USP" && status == "Aguardando Entrega do Fornecedor" && diff == 13.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 2 dias.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }
    
    if (localidade == "USP" && status == "Aguardando Entrega do Fornecedor" && diff == 14.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 1 dia.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado amanhã corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }  
    
     if (localidade == "USP" && status == "Aguardando Entrega do Fornecedor" && diff == 15.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotou.</strong></font><br><br>"+
                   "<font size='3'><strong>Envie o comunicado ao usuário o quanto antes! Via opção Entrega Expirada, no menu Notificação, na planilha.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      SpreadsheetApp.flush();
  }
    
/** Prodata CA **/
    
    if (localidade == "CA" && status == "Aguardando Entrega do Fornecedor" && diff == 13.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 2 dias.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }
    
    if (localidade == "CA" && status == "Aguardando Entrega do Fornecedor" && diff == 14.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 1 dia.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado amanhã corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }  
    
     if (localidade == "CA" && status == "Aguardando Entrega do Fornecedor" && diff == 15.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotou.</strong></font><br><br>"+
                   "<font size='3'><strong>Envie o comunicado ao usuário o quanto antes! Via opção Entrega Expirada, no menu Notificação, na planilha.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      SpreadsheetApp.flush();
  }
    
/** Prodata RCL **/
    
     if (localidade == "RCL" && status == "Aguardando Entrega do Fornecedor" && diff == 15.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 2 dias.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }
    
    if (localidade == "RCL" && status == "Aguardando Entrega do Fornecedor" && diff == 16.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 1 dia.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado amanhã corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }  
    
     if (localidade == "RCL" && status == "Aguardando Entrega do Fornecedor" && diff == 17.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotou.</strong></font><br><br>"+
                   "<font size='3'><strong>Envie o comunicado ao usuário o quanto antes! Via opção Entrega Expirada, no menu Notificação, na planilha.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      SpreadsheetApp.flush();
  }
    
/** Prodata JOI **/
    
     if (localidade == "JOI" && status == "Aguardando Entrega do Fornecedor" && diff == 8.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 2 dias.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }
    
    if (localidade == "JOI" && status == "Aguardando Entrega do Fornecedor" && diff == 9.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 1 dia.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado amanhã corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }  
    
     if (localidade == "JOI" && status == "Aguardando Entrega do Fornecedor" && diff == 10.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotou.</strong></font><br><br>"+
                   "<font size='3'><strong>Envie o comunicado ao usuário o quanto antes! Via opção Entrega Expirada, no menu Notificação, na planilha.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      SpreadsheetApp.flush();
  }
    
/** Prodata MAN **/
    
      if (localidade == "MAN" && status == "Aguardando Entrega do Fornecedor" && diff == 28.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 2 dias.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }
    
    if (localidade == "MAN" && status == "Aguardando Entrega do Fornecedor" && diff == 29.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 1 dia.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado amanhã corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }  
    
     if (localidade == "MAN" && status == "Aguardando Entrega do Fornecedor" && diff == 30.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotou.</strong></font><br><br>"+
                   "<font size='3'><strong>Envie o comunicado ao usuário o quanto antes! Via opção Entrega Expirada, no menu Notificação, na planilha.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      SpreadsheetApp.flush();
   }
    
/** DELL **/
    
    if (fornecedor == "Dell" && status == "Aguardando Entrega do Fornecedor" && diff == 5.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 2 dias.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 2 dias! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 2 dias! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }
    
    if (fornecedor == "Dell" && status == "Aguardando Entrega do Fornecedor" && diff == 6.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotará em 1 dia.</strong></font><br><br>"+
                   "<font size='3'><strong>Mantenha esse caso no radar para que o envio do comunicado de expiração da entrega seja enviado amanhã corretamente.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Entrega do produto expirará em 1 dia! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Entrega do produto expirará em 1 dia! - Chamado "+row[0]});
      SpreadsheetApp.flush();
      
    }  
    
     if (fornecedor == "Dell" && status == "Aguardando Entrega do Fornecedor" && diff == 7.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>O prazo de entrega do produto deste chamado se esgotou.</strong></font><br><br>"+
                   "<font size='3'><strong>Envie o comunicado ao usuário o quanto antes! Via opção Entrega Expirada, no menu Notificação, na planilha.</strong></font><br><br>"+
                   "<font size='3'><strong>Chamado:</strong> "+row[0]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray2,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray3,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      MailApp.sendEmail(emailarray4,"Prazo de entrega esgotou! - Chamado "+row[0], "", {htmlBody: message , noReply:true, name: "Prazo de entrega esgotou! - Chamado "+row[0]});
      SpreadsheetApp.flush();
  }
 }
}