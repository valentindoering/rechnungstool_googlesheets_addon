function onInstall(e) {
  onOpen(e)
  
  /* That would be creating the ui, but Auth woold be required
  alert("Die Benutzeroberfläche für das Rechnungstool wird jetzt erstellt. Dieser Vorgang wird ca. 20 sec dauern, bitte klicke auf ok und warte dann auf das Fertig-Banner.")
  var answer = createUi()
  alert(answer)
  */
}


function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
       .addItem('Rechnung erstellen', 'start')
       .addToUi();
}

function start() {
  setGlobals()
  
  var missingSheets = getMissingSheets(sheetNames)
  
  if (missingSheets.length === 0) {
  
    var html = HtmlService.createTemplateFromFile("Sidebar").evaluate().setTitle('Menü')
    ui.showSidebar(html)
  
  } else {
  
    var answer = Browser.msgBox('Rechnungstool einrichten', 'Bist du dir sicher, dass du die für das Rechnungstool benötigten Tabellenseiten einfügen willst? Dieser Vorgang wird ca. 20 sec dauern, bitte warte auf das Fertig-Banner.', Browser.Buttons.YES_NO);
  
    if (answer === "yes") {
      var answer = createUi()
      
      var html = HtmlService.createTemplateFromFile("Sidebar").evaluate().setTitle('Menü')
      ui.showSidebar(html)
      SpreadsheetApp.flush()
      
      alert(answer)
    }
  
  }
  
}

