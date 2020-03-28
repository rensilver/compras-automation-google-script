/** Criado por Renato Oliveira Batista da Silveira - ren.oliv87@gmail.com **/
var ui = SpreadsheetApp.getUi();
var ss = SpreadsheetApp.getActive();

function onOpen() {
  
  ui.createMenu("Email")
  .addItem("Enviar Email", "sendTemplate")
  .addToUi(); 
}

function sendTemplate() { 
  var sheet = SpreadsheetApp.getActive().getSheetByName("template");
  var startRow = 2;
  var numRows = 1000;
  var dataRange = sheet.getRange(startRow, 1 , numRows, 3);
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddr = row[0];
    var nome = row[1];
    var emailSent = row[2];
    var forDriveScope = DriveApp.getStorageUsed();
    var url = "https://docs.google.com/document/d/1kSgQezWt9R6WPA_xGkPhzAd3yXLlS4m3Ye6uHbCYuNA/export?format=html";
    var param = {
      method      : "get",
      headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions:true,
  };
  var html = UrlFetchApp.fetch(url, param).getContentText();
    
    if (emailSent != "Enviado") {
      
    MailApp.sendEmail(emailAddr, "Comunicado - Tecnologia da Informação", html, {noReply:true, name: "Comunicado - Tecnologia da Informação"});
    sheet.getRange(startRow + i, 3).setValue("Enviado");
    SpreadsheetApp.flush();
    }
  }
}
