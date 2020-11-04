// 30-10-2020
// Automation of collection email 
// Nolunga Ngcakane


function sendMails() {
  
  var EMAIL_SENT = 'EMAIL_SENT';
  
  
  var workBook = SpreadsheetApp.getActiveSpreadsheet();
  var workSheetRepairID = workBook.getSheetByName("Repair_ID");
  var workSheetMsg = workBook.getSheetByName("Mail_Details"); 
  
  
  
  
  var subject = workSheetMsg.getRange('A3').getValue();
  var message = workSheetMsg.getRange('B3').getValue();
  var emailAddress = workSheetMsg.getRange('C3').getValue();
  
  
// set i for number of rows you'd like to process 
// dont go over data set, counting starts at 0,1,2,3,4,5
  for (var i=2; i<=10; i++){
    
    var model = workSheetRepairID.getRange('G' + i).getValue();
    var fname = workSheetRepairID.getRange('C' + i).getValue();
    var lname = workSheetRepairID.getRange('D' + i).getValue();
    var cellNum = workSheetRepairID.getRange('E' + i).getValue();
    var serialNum = workSheetRepairID.getRange ('H' + i).getValue();
    var finalMsg = "";
    var PoP = workSheetRepairID.getRange('F' +i).getValue();
    var WarrantyStat = workSheetRepairID.getRange('I' + i).getValue();
    
    
    var finalSubject = subject + fname + " " + lname +" " + serialNum ;
    
    
    
    finalMsg = "Good day, Please arrange repair collection for " + fname + " " + lname + " . " + "\n"  + "\n" + message;
    
    finalMsg = finalMsg.replace("Model :","Model :" + model);
    finalMsg = finalMsg.replace("Serial number :","Serial number : " + serialNum);
    finalMsg = finalMsg.replace("Customer Details:","Customer Details:" + " " + fname + " " + lname + " / " + emailAddress + " / " + cellNum);
    finalMsg = finalMsg.replace("Warranty :","Warranty :" + WarrantyStat);
    finalMsg = finalMsg.replace("Proof of Purchase. :","Proof of Purchase. :" + " " + PoP);
    
    var emailSent = workSheetRepairID.getRange('N' + i).getValue();
    
        if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates
    
       MailApp.sendEmail(emailAddress, finalSubject, finalMsg);
      workSheetRepairID.getRange('N'+ i).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    
    
    
    
    
    
  }
    
  
}


}