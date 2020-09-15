function setGlobals() {
  
  ss = SpreadsheetApp.getActiveSpreadsheet()

  soge = ss.getSheetByName("So geht's") // So geht's
  db = ss.getSheetByName("Datenbank") //Datenbank
  his = ss.getSheetByName("Verlauf")// Verlauf
  bill = ss.getSheetByName("x") // Rechnung
  
  
  ui = SpreadsheetApp.getUi();
  
  zahlungsbedingung = 10 // WT
  colour = "#000000"
  sheetNames = ["So geht's","Datenbank", "Verlauf", "x"] // Real:  ["Hausbank", "Müll", "Radau", "Schweinerei"]// muss 4 in dieser Reiehenfolge!!
  exportFolderName = "Rechnungen-Export"
  
  // UI
  mainColour = "#8F0021"
  mainFontColour = "#ffffff"
  tableLightGray = "#F3F3F3"
  tableDarkGray = "#D9D9D9"
  normalFontSize = 11
  normalFont = "Helvetica Neue"


}


