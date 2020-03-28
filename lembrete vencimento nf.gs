/** Criado por Renato Oliveira Batista da Silveira - ren.oliv87@gmail.com **/
function vencNf() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("2020");
  var startRow = 2;
  var numRows = 3000;
  var dataRange = sheet.getRange(startRow, 1 , numRows, 22);
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var status = row[14];
    var dateVenc = row[17];
    var localidade = row[6];
    var fornecedor = row[11];
    var emailAddress = ['rodrigo_medeiros_tsystems@whirlpool.com','rodrigo_duarte_T-Systems@whirlpool.com','policarpo_moser_T-Systems@whirlpool.com','renato_oliveira_batista_tsystems@whirlpool.com'];
    var date = row[18];
    var d = new Date();
    var t = new Date(date);
    var t1 = d.getTime();
    var t2 = t.getTime();
    
    var diff = Math.floor((t1-t2)/(24*3600*1000));
    
   /** Fornecedor 30 dias **/
    
    if (dateVenc == "30 dias" && diff == 15.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1]; 
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>Faltam 15 dias para o vencimento desta Nota Fiscal.</strong></font><br><br>"+
                   "<font size='3'><strong>Fornecedor:</strong> "+row[11]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray2,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray3,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray4,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      SpreadsheetApp.flush();
  }
  
 /** Fornecedor 60 dias **/
     
     if (dateVenc == "60 dias" && diff == 45.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>Faltam 15 dias para o vencimento desta Nota Fiscal.</strong></font><br><br>"+
                   "<font size='3'><strong>Fornecedor</strong> "+row[11]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray2,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray3,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray4,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      SpreadsheetApp.flush();
  }
    
 /** Fornecedor 90 dias **/
    
    if (dateVenc == "90 dias" && diff == 75.0) {
      
      var emailarray1 = emailAddress[0];
      var emailarray2 = emailAddress[1];
      var emailarray3 = emailAddress[2];
      var emailarray4 = emailAddress[3];
      var message ="<font size='3'><strong>Prezados, bom dia!</strong></font><br><br>"+
                   "<font size='3'><strong>Faltam 15 dias para o vencimento desta Nota Fiscal.</strong></font><br><br>"+
                   "<font size='3'><strong>Fornecedor</strong> "+row[11]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Número do Pedido:</strong> "+row[9]+"</font>"+"<br><br>"+
                   "<font size='3'><strong>Produto:</strong> "+row[5]+"</font>";
      
      MailApp.sendEmail(emailarray1,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray2,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray3,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      MailApp.sendEmail(emailarray4,"Nota Fiscal vencerá em 15 dias! - Pedido "+row[9], "", {htmlBody: message , noReply:true, name: "Nota Fiscal vencerá em 15 dias! - Pedido "+row[9]});
      SpreadsheetApp.flush();
   }
 }
}