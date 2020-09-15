
function alert(text) {
  SpreadsheetApp.getUi().alert(text);
}

function getDataArray(sheet, row, column, numRows, numColumns) {
  return sheet.getRange(row, column, numRows, numColumns).getValues()
}

function datum(format) {
  // zB.: dd/MM/yyyy, MM-dd-yyyy HH:mm:ss
  var timezone = "GMT+1"
  return Utilities.formatDate(new Date(), timezone, format)
}

function nowMonthString() {
  var months = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
  var now = new Date()
  var mt = now.getMonth();
  return months[mt]
}

function nowPlusXDays(forwardInTime, format) {
  var timezone = "GMT+1"
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var plusDate = new Date(new Date().getTime() + (forwardInTime * MILLIS_PER_DAY));
  return Utilities.formatDate(plusDate,timezone, format)
}


function setArrayValue(sheet, row, column, array) {
  // es können nur Rechtecke eingesetzt werden
  var numRows = array.length
  var numColumns = array[0].length
  sheet.getRange(row, column, numRows, numColumns).setValues(array)
}

function startsWith(str, word) {
    return str.lastIndexOf(word, 0) === 0;
}


function createPdf(ss,sheet, fileName) {

  var ssID = ss.getId()
  var url = "https://docs.google.com/spreadsheets/d/"+ssID+"/export"+
                                                        "?format=pdf&"+
                                                        "size=a4"+
                                                         //"&scale=0.85"+

                                                        "portrait=true&"+
                                                          //"&scale=1"+
                                                        //"fitw=true&"+
                                                        "gridlines=false&"+
                                                          "&top_margin=0.4"+
                                                            "&bottom_margin=0.00"+
                                                              "&left_margin=0.3"+
                                                                "&right_margin=0.6"+

                                                                  '&gid=' + sheet.getSheetId();


  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};

  return UrlFetchApp.fetch(url, params).getBlob().setName(fileName);

}




function manageDriveSaving(response, fileName,exportFolderName) {

   // Get Acess
   var SSID=ss.getId();
   var thisFile = DriveApp.getFileById(SSID)
   
   // Pdf abespeichernOrdner finden, auch wenn Name geändert wurde
   // File kann dupliziert aber connectet sein -> mehrere Oberornder
   var parentFolders = thisFile.getParents() 
   
   // parentFolder ist ein ITERATOR - über seine praktischen Methoden kann man auf seine Elemente zugreifen
   while (parentFolders.hasNext()) {
     var parent = parentFolders.next();
     var parentId = parent.getId()
     var childFolders = DriveApp.getFolderById(parentId).getFolders();
     
     var exportFound = false
     while(childFolders.hasNext()) {
      var child = childFolders.next();
      
      if (child.getName() === exportFolderName) {
        exportFound = true
        
        // ExporTiere in child
        saveToDrive(response, fileName, child.getId())
        }
      }
      
      // Wenn in keinen einzigen exportiert wurde, erstelle neu und exportiere
      if (!exportFound) {
        var newExportFolder = parent.createFolder(exportFolderName) //MARKUP brauche ich eine ID?
        
        //Exportiere in newExportFolder
        saveToDrive(response, fileName, newExportFolder.getId())
        
        
      }
   }
}



function saveToDrive(response, fileName, ordnerID) {

  DriveApp.createFile(response);
  var files = DriveApp.getFilesByName(fileName);

  while (files.hasNext()) {
    var file = files.next();
    var destination = DriveApp.getFolderById(ordnerID);
    destination.addFile(file);
    var pull = DriveApp.getRootFolder();
    pull.removeFile(file);
  }
}




function sendEmail(mailSubject, empfaengerMail, mailArray, template, attachment) {

  var html = HtmlService.createTemplateFromFile(template);
  html.ma = mailArray

  var template = html.evaluate().getContent();

  MailApp.sendEmail({
   to: empfaengerMail,
   subject: mailSubject,
   htmlBody: template,
   attachments: attachment
   });

}

function getActualSheetsNames() {

  var sheets = ss.getSheets()
   var actualSheetsNames = []
   for (var i = 0; i < ss.getNumSheets(); i++) {
     actualSheetsNames.push(sheets[i].getName())
   }
   return actualSheetsNames
}

function getMissingSheets(sheetNames) {

   
   var actualSheetsNames = getActualSheetsNames()
   var missingSheets = []
   for (var i = 0; i < sheetNames.length; i++) {
   
     if (!actualSheetsNames.includes(sheetNames[i])) {
       missingSheets.push(sheetNames[i])
     }
   }
   
   return missingSheets
   
  }
  
  function getNeededExistingSheets() {
    var missingSheets = getMissingSheets(sheetNames)
    var existingSheets = []
    
    for (var i = 0; i < sheetNames.length; i++) {
      if (!missingSheets.includes(sheetNames[i])) {
        existingSheets.push(sheetNames[i])
      }
    }
    
    return existingSheets
  }
  
  
  function getJsonData(profil) {
  
  var jsonData = {
    'mail' : db.getRange(10,6+profil).getValue()
  };
  
  return jsonData;
  
 }


function postJsonData(url, jsonData) {

  
  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload' : JSON.stringify(jsonData)
  };
  
  return UrlFetchApp.fetch(url, options);
}

