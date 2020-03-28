/** Criado por Renato Oliveira Batista da Silveira - ren.oliv87@gmail.com **/
/** Cria menu na UI da planilha com itens selecionáveis. **/

var ui = SpreadsheetApp.getUi();
var ss = SpreadsheetApp.getActive();

function onOpen() {
  
  ui.createMenu("Notificação")
  .addItem("Pedido Criado", "orderCreated")
  .addItem("Chegou Produto", "productArrived")
  .addItem("Lembrete NF", "vencDate")
  .addToUi(); 
}

/** Envia e-mail ao usuário informando que o pedido foi criado e que o produto está a caminho. **/

function orderCreated() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("2020");
  var startRow = 2;
  var numRows = 3000;
  var dataRange = sheet.getRange(startRow, 1 , numRows, 22);
  var data = dataRange.getValues();
  var html1 = HtmlService.createTemplateFromFile('template-envio1').evaluate().getContent();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var status1 = row[14];
    var emailPedidoCriado = row[19];
    var name = row[3];
    var html2 = HtmlService.createTemplateFromFile('template-envio2').evaluate().getContent(); 
    var tell = "<p align='center'><font size='5'><strong>Olá "+row[3]+",</strong></font></p>";
    
    if (emailPedidoCriado == "Enviado") {
    
      continue;
    } 
    
    if (status1 == 'Aguardando Entrega do Fornecedor') {
      
      var emailAddress = row[4];
      var d = new Date();
      var message ="<p align='center'><font size='4'><strong>Prazo de entrega:</strong> "+row[15]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Chamado:</strong> "+row[0]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Número do Pedido:</strong> "+row[9]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Produto:</strong> "+row[5]+"</font></p>";
      
      MailApp.sendEmail(emailAddress,"Seu pedido já está a caminho! - Chamado "+row[0], "", {htmlBody: html1+tell+html2+message , noReply:true, name: "Seu pedido já está a caminho! - Chamado "+row[0]});
      sheet.getRange(startRow + i, 20).setValue("Enviado");
      sheet.getRange(startRow + i, 17).setValue(d);
      SpreadsheetApp.flush();
    }
  }
}

function productArrived() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("2020");
  var startRow = 2;
  var numRows = 3000;
  var dataRange = sheet.getRange(startRow, 1 , numRows, 22);
  var data = dataRange.getValues();
  var html1 = HtmlService.createTemplateFromFile('template-chegou1').evaluate().getContent();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var status1 = row[14];
    var emailProdutoChegou = row[21];
    var name = row[3];
    var html2 = HtmlService.createTemplateFromFile('template-chegou2').evaluate().getContent(); 
    var tell = "<p align='center'><font size='5'><strong>Olá "+name+",</strong></font></p>";
    
   if (emailProdutoChegou == "Enviado") {
    
      continue;
    }
    
    if (status1 == 'Entregue (Não Faturado)' || status1 == 'Entregue (Faturado)') {
      
      var emailAddress = row[4];
      var message ="<p align='center'><font size='4'><strong>Chamado:</strong> "+row[0]+"</font></p>"+"<br>"+
                   "<p align='center'><font size='4'><strong>Produto:</strong> "+row[5]+"</font></p>";
      
      MailApp.sendEmail(emailAddress,"Disponível para retirada na TI - Chamado "+row[0], "", {htmlBody: html1+tell+html2+message , noReply: true, name: "Disponível para retirada na TI - Chamado "+row[0]});
      sheet.getRange(startRow + i, 22).setValue("Enviado");
      SpreadsheetApp.flush();
    }
  }
}

function vencDate() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("2020");
  var startRow = 2;
  var numRows = 3000;
  var dataRange = sheet.getRange(startRow, 1 , numRows, 22);
  var data = dataRange.getValues();
  
   for (var i = 0; i < data.length; ++i) {
     var row = data[i];
     var vencPrazo = row[17];
     var nfDate = new Date();
      
     if (vencPrazo == "30 dias") {
      
     sheet.getRange(startRow + i, 19).setValue(nfDate);
     SpreadsheetApp.flush();
     }
     else if (vencPrazo == "60 dias") {
     sheet.getRange(startRow + i, 19).setValue(nfDate);
     SpreadsheetApp.flush();
     }
     else if (vencPrazo == "90 dias") {
     sheet.getRange(startRow + i, 19).setValue(nfDate);
     SpreadsheetApp.flush();
    }
  }
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}