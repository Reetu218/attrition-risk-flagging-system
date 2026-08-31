
function flagAtRiskEmployees() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmployeeData");
  var data = sheet.getDataRange().getValues();
  
  sheet.getRange(1, 6).setValue("Status");
  
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    var department = data[i][1];
    var tenure = data[i][2];
    var score = data[i][3];
    var previousStatus = data[i][5];
    var status = "";
    
    if (tenure === "" || score === "") {
      continue;
    }
    
    if (tenure < 6 && score < 3) {
      status = "At Risk";
      
      if (previousStatus !== "At Risk") {
        sendRiskEmail(name, department, tenure, score);
      }
    } else {
      status = "Stable";
    }
    
    sheet.getRange(i + 1, 6).setValue(status);
  }
}

function sendRiskEmail(name, department, tenure, score) {
  var recipient = "name_email@gmail.com";
  var subject = "Attrition Risk Alert: " + name;
  var body = "Employee: " + name + "\n" +
             "Department: " + department + "\n" +
             "Tenure: " + tenure + " months\n" +
             "Performance Score: " + score + "\n" +
             "Reason: Low tenure combined with low performance score.";
  
  MailApp.sendEmail(recipient, subject, body);
}
