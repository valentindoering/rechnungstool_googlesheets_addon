

function main(profil, customer, erfuellung, leistungsString, mailAn, positionenArray) {

  setGlobals()
  
  var missingSheets = getMissingSheets(sheetNames)
  
  if (missingSheets.length === 0) {
    
    
      // Schaue, ob das bisherige Programm schon fertig ist
     if (db.getRange(1,5).getValue() === "BEREIT"){
     
       db.getRange(1,5).setValue("RECHNET")
       db.getRange('A2').activate()
       
       var billObject = getBillData(profil, customer, erfuellung, leistungsString, mailAn, positionenArray)
    
        if (billObject.complete) {
          
             var url = "https://script.google.com/macros/s/AKfycbzXbjgkOtQyxeaEL9i_cVTGEpzcFGqnLbMKype9/exec"
             var jsone = getJsonData(profil)
             var response = postJsonData(url, jsone)  
             
             if (response == "ok") {
               
               createBill(billObject)
               SpreadsheetApp.flush()
               
               db.getRange(1,5).setValue("BEREIT")
               SpreadsheetApp.flush()
               alert("Erfolgreich! Rechnung erstellt, per Mail versendet und in deinem Google Drive gespeichert. Sieh dir dein Rechnungserstellungs-Verlauf unter der Tabellenseite Verlauf an.")
               
             } else {
               db.getRange(1,5).setValue("BEREIT")
               alert(response)
             }
 
        
        } else {
          
          db.getRange(1,5).setValue("BEREIT")
          alert(billObject.errorMessage)
        
        }
     
           
           
     } else {
       alert("Bitte warte bis das Programm mit der zuletzt beauftragten Rechnung fertig ist. Der Programm-Status springt dann automatisch auf 'BEREIT'.")
     }
  
  } else {
    
    var alertString = "Es fehlen die Tabellenseite/n: "
    for (var i = 0; i < missingSheets.length; i++) {
      alertString += "'"+missingSheets[i] + "', "
    }
    alertString += "| Das Rechnungstool braucht für die Ausführung die korrekte Benutzeroberfläche im Tabellendokument. Jetzt einrichten?"
    
    var answer = Browser.msgBox('Rechnungstool einrichten', alertString, Browser.Buttons.YES_NO);
  
    if (answer === "yes") {
      var answer = createUi()   
      alert(answer)
    }

  }



}


// GET BILL e  _________________________________________________________________________________________________________

function getBillData(profil, customer, erfuellung, leistungsString, mailAn, positionenArray) {

    Logger.log("Entering the getBille")
  
   var anzahlPositionen = positionenArray.length

  // artikelArray
    var artikelArray = getDataArray(db,5,3,25,2) // numArticle, NumAttributArticle

  // profilArray
    var profilArray = getDataArray(db,4,7,14,6) // num Attributes, num Profiles
    var profilIndex = profil-1

  // customerArray
    var customerArray = getDataArray(db,5,15,25,7) // numPeople, numAttributes
    var customerIndex = customer - 1
    

  //________________________________________________________________

  // aktuelleRechnung Jede Kategorie auf Rechnung hat für die Vertikale einen Array
  // eine Position auf Rechnung nimmt aus jedem Array das Element mit seinem Index
  var pos = [] // 1 2 3 4 ...
  var art = [] // Artikelnummern für jeweilige Position
  var beschr = [] // anhand der Artikelnnummer passende Beschreibung pasten
  var ep = [] // anhand der Artikelnnummer passenden Einzelpreis pasten
  var ah = [] // Arbeitsstunden für jeweilige Position
  var betrag = [] // ep * ah


  // Arrays entsprechend befüllen
  var aktuellArtikel;
  for(var i = 0; i < anzahlPositionen; i++) {

    pos.push(i + 1) // 1 2 3 4 ...
    aktuellArtikel = positionenArray[i][0]
    art.push(aktuellArtikel)

    beschr.push(artikelArray[aktuellArtikel - 1][0])
    ep.push(artikelArray[aktuellArtikel - 1][1])

    ah.push(positionenArray[i][1])
    betrag.push((ep[i])*(ah[i]))

  }


  // Summe
  var summe = 0
  for(var i = 0; i<anzahlPositionen; i++){
    summe += betrag[i]
  }



  var billObject = { 
  
    'complete' : true,
    'errorMessage' : "Daten vollständig",
  
    'profil' : {
    
      'nummer' : profil,
      'firma' : profilArray[0][profilIndex],
      'vorname' : profilArray[1][profilIndex],
      'nachname' : profilArray[2][profilIndex],
      'strassehausnummer' : profilArray[3][profilIndex],
      'plzort' : profilArray[4][profilIndex],
      'mobil' : profilArray[5][profilIndex],
      'email' : profilArray[6][profilIndex],
      'bank' : profilArray[7][profilIndex],
      'iban' : profilArray[8][profilIndex],
      'bic' : profilArray[9][profilIndex],
      'finanzamt' : profilArray[10][profilIndex],
      'steuernummer' : profilArray[11][profilIndex],
      'absatz' : profilArray[12][profilIndex],
      'rechnungsnummer' : profilArray[13][profilIndex]
    
     },
    
    'kunde' : {
       
      'nummer' : customer,
      'firma' : customerArray[customerIndex][0],
      'vorname' : customerArray[customerIndex][1],
      'nachname' : customerArray[customerIndex][2],
      'geschlecht' : customerArray[customerIndex][3],
      'strassehausnummer' : customerArray[customerIndex][4],
      'plzort' : customerArray[customerIndex][5],
      'email' : customerArray[customerIndex][6],
      'telefon' : customerArray[customerIndex][7]
    
    },
    
    'erfuellung' : erfuellung,
    'leistungsString' : leistungsString,
    'mailAn' : mailAn,
    'letzteposreihe' : 16 + anzahlPositionen,
    'anzahlPositionen' : anzahlPositionen,
    
    'pos' : pos,
    'art' : art,
    'beschr' : beschr, 
    'ep' : ep,
    'ah' : ah,
    'betrag' : betrag,
    
    'summe' : summe

  }
  
  // Check Artikelbereich
  var artikelOk = true
  for (var i = 0; i < anzahlPositionen; i++) {
    if (beschr[i] == "" || ep[i] == "") {
      artikelOk = false
    }
  }
  
  // Check Profilbereich
  var profilOk = true
  for (var element in billObject.profil) {
    if (billObject.profil[element] === "") {
      profilOk = false
    }
  }
  
  // Check Kundenbereich
  var kundeOk = true
  for (var element in billObject.kunde) {
    if (billObject.kunde[element] === "") {
      kundeOk = false
    }
  }

  
  // IS ERROR?
  if (!(artikelOk && profilOk && kundeOk)) {
    billObject.complete = false
    billObject.errorMessage = "Für die gewünschte Rechnung fehlen mir noch Daten im: " + (artikelOk ? "" : "Artikelbereich, ") + (profilOk ? "" : "Profilbereich, ") + (kundeOk ? "" : "Kundenbereich")
  }

  return billObject

}






// CREATE BILL  _________________________________________________________________________________________________________


function createBill(e) {
    
 // parseInt() Nächest Rechnungsnummer setzen
  db.getRange(17,7 + e.profil.nummer-1).setValue(e.profil.rechnungsnummer + 1) // Nächste rechnungsnummer auf Dashboard bringen

  //________________________________________________________________

  // Bill Design
  var fontSize = 12
  var font = "Helvetica Neue"
  bill.getRange(1,1,42,11).setFontFamily(font).setFontSize(fontSize).setHorizontalAlignment("left")

  // Alte Rechnung löschen
  bill.getRange(17, 2, 28, 9).clearContent().clearFormat()

  //________________________________________________________________

  // RECHNUNG - einsetzen
  bill.getRange(8,2,31,9).clear() // Tabelle bereinigen, alte Positionen/ Unterstreichnungen entfernen
  bill.getRange(8,2,31,9).setFontFamily(font).setFontSize(fontSize).setHorizontalAlignment("left") // Schriftart und Größe im Bereich wiederherstellen

  // Header (Firma)
  bill.getRange(1,2).setFontSize(40).setFontWeight("bold").setValue(e.profil.firma)

  // Adressfeld
  bill.getRange(3,2).setFontSize(8).setVerticalAlignment("middle").setValue(e.profil.firma+" | "+e.profil.vorname+" "+e.profil.nachname+" | "+e.profil.strassehausnummer+" | "+e.profil.plzort)

  bill.getRange(4,2).setFontWeight("bold").setValue(e.kunde.firma)
  bill.getRange(5,2).setValue(e.kunde.vorname + " " + e.kunde.nachname)
  bill.getRange(6,2).setValue(e.kunde.strassehausnummer)
  bill.getRange(7,2).setFontWeight("bold").setValue(e.kunde.plzort)

  //Rechnungsdaten
  bill.getRange(8,9).setHorizontalAlignment("right").setValue("Datum")
  bill.getRange(8,10).setHorizontalAlignment("right").setValue(datum("dd.MM.yyyy"))
  bill.getRange(9,9).setHorizontalAlignment("right").setValue(e.erfuellung)
  bill.getRange(9,10).setHorizontalAlignment("right").setValue(e.leistungsString)
  bill.getRange(10,9).setHorizontalAlignment("right").setValue("Kunden-Nr.")
  bill.getRange(10,10).setHorizontalAlignment("right").setNumberFormat("00").setValue(e.kunde.nummer)
  bill.getRange(11,9).setHorizontalAlignment("right").setValue("Zahlungsbedingungen")
  bill.getRange(11,10).setHorizontalAlignment("right").setValue("innerh. "+zahlungsbedingung+" WT")
  bill.getRange(12,10).setHorizontalAlignment("right").setValue("nach Erhalt")


  // Überschrift
  bill.getRange(14,2).setFontWeight("bold").setFontSize(18).setValue("Rechnung Nr. "+ e.profil.rechnungsnummer)

  //Table
  bill.getRange(16, 2, 1,9 ).setBackground("#f1f1f1").setBorder(false, false, true, null, false, false,"#000000", SpreadsheetApp.BorderStyle.SOLID);

  bill.getRange(16,2).setHorizontalAlignment("center").setVerticalAlignment("middle").setValue("Pos.")
  bill.getRange(16,3).setHorizontalAlignment("center").setVerticalAlignment("middle").setValue("Art.")
  bill.getRange(16,4).setVerticalAlignment("middle").setValue("Beschreibung")
  bill.getRange(16,8).setHorizontalAlignment("center").setVerticalAlignment("middle").setValue("EP")
  bill.getRange(16,9).setHorizontalAlignment("center").setVerticalAlignment("middle").setValue("Ah/ Stk.")
  bill.getRange(16,10).setHorizontalAlignment("right").setVerticalAlignment("middle").setValue("Nettobetrag")

  bill.getRange((e.letzteposreihe+1), 2, 1,9 ).setBorder(true, false, false, null, false, false,"#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Artikel
  var arrayReihe = 17
  for(var i=0; i<e.anzahlPositionen; i++){

    bill.getRange(arrayReihe,2).setHorizontalAlignment("center").setVerticalAlignment("middle").setValue(e.pos[i])
    bill.getRange(arrayReihe,3).setHorizontalAlignment("center").setVerticalAlignment("middle").setValue(e.art[i])
    bill.getRange(arrayReihe,4).setVerticalAlignment("middle").setValue(e.beschr[i])
    bill.getRange(arrayReihe,8).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00€").setValue(e.ep[i])
    bill.getRange(arrayReihe,9).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.0").setValue(e.ah[i])
    bill.getRange(arrayReihe,10).setHorizontalAlignment("right").setVerticalAlignment("middle").setNumberFormat("0.00€").setValue(e.betrag[i])
    arrayReihe ++

  }
  var letzteposreihe = e.letzteposreihe
  //Summe
  bill.getRange(letzteposreihe+2,9).setValue("Summe")
  bill.getRange(letzteposreihe+2,10).setNumberFormat("00.00€").setHorizontalAlignment("right").setValue(e.summe)
  
  
  var is_mehrwertsteuer = e.profil.absatz.startsWith("Mehrwertsteuer")
  if (is_mehrwertsteuer) {
    
    var mehrwert_sum = e.summe
  
    switch (e.profil.absatz) {
      case "Mehrwertsteuer 5%":
      mehrwert_sum *= 0.05
      break
      
      case "Mehrwertsteuer 7%":
      mehrwert_sum *= 0.07
      break
      
      case "Mehrwertsteuer 16%":
      mehrwert_sum *= 0.16
      break
      
      case "Mehrwertsteuer 19%":
      mehrwert_sum *= 0.19
      break
    
    }
    
    bill.getRange(letzteposreihe+4,9).setHorizontalAlignment("right").setValue(e.profil.absatz)
    bill.getRange(letzteposreihe+4,10).setNumberFormat("00.00€").setHorizontalAlignment("right").setValue(mehrwert_sum)
    letzteposreihe += 2
    
    // Update sum
    e.summe += mehrwert_sum
  
  }
  
  
  bill.getRange(letzteposreihe+4,9).setFontWeight("bold").setValue(is_mehrwertsteuer ? "Gesamt" : "Gesamt*")
  bill.getRange(letzteposreihe+4,10).setFontWeight("bold").setNumberFormat("00.00€").setHorizontalAlignment("right").setValue(e.summe)
  bill.getRange(letzteposreihe+4, 9, 1,2 ).setBorder(false, false, true, null, false, false,"#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);


  // Text
  bill.getRange(letzteposreihe+6,2).setValue("Bitte überweisen Sie den Rechnungsbetrag innerhalb von "+zahlungsbedingung+" Wochentagen auf das unten genannte Konto.")
  if (!is_mehrwertsteuer) {
    bill.getRange(letzteposreihe+7,2).setValue("*"+ e.profil.absatz)
  } else {
    letzteposreihe--
  }
  
  bill.getRange(letzteposreihe+9,2).setFontWeight("bold").setVerticalAlignment("middle").setValue("Vielen Dank für Ihren Auftrag und die angenehme Zusammenarbeit!")

  bill.getRange(letzteposreihe+11,2).setValue("Mit freundlichen Grüßen")

  bill.getRange(letzteposreihe+12,2).setValue(e.profil.vorname+" "+e.profil.nachname)

  //Anhang
  var anhanghoehe = 38
  bill.getRange(anhanghoehe, 2, 1,9 ).setBorder(true, false, false, null, false, false,colour, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  bill.getRange(anhanghoehe,2,3,9).setFontSize(8)

  bill.getRange(anhanghoehe,2).setValue(e.profil.firma)
  bill.getRange(anhanghoehe+1,2).setValue("Inh. "+e.profil.vorname+" "+e.profil.nachname)
  bill.getRange(anhanghoehe+2,2).setValue(e.profil.strassehausnummer+" | "+e.profil.plzort)

  bill.getRange(anhanghoehe,5).setValue("Mobil "+e.profil.mobil)
  bill.getRange(anhanghoehe+1,5).setValue("E-Mail "+e.profil.email)

  bill.getRange(anhanghoehe,7).setValue(e.profil.bank)
  bill.getRange(anhanghoehe+1,7).setValue("IBAN "+e.profil.iban)
  bill.getRange(anhanghoehe+2,7).setValue("BIC "+e.profil.bic)

  var space = "          ";
  bill.getRange(anhanghoehe,9).setValue(space+"Steuernummer/-identif.")
  bill.getRange(anhanghoehe+1,9).setValue(space+"Finanzamt "+e.profil.finanzamt)
  bill.getRange(anhanghoehe+2,9).setValue(space+e.profil.steuernummer)

  SpreadsheetApp.flush() // Sehr wichtig! Rechnung muss erst heruntergeschrieben werden, bevor sie exportiert wird

  //________________________________________________________________
  // Create Pdf, Save to Drive

  var fileName = datum("yyyy-MM-dd") + " Rechnung Nr. " + e.profil.rechnungsnummer + " - " + e.kunde.firma + ".pdf"
  var response = createPdf(ss, bill, fileName)
  manageDriveSaving(response, fileName, exportFolderName) // Letzteres ist global

  SpreadsheetApp.flush()

  //________________________________________________________________
  // sendEmail

  var anrede
  if (e.kunde.geschlecht === "m") {
     anrede = "Sehr geehrter Herr"
  } else {
    anrede = "Sehr geehrte Frau"
  }

  var mailSubject = e.profil.firma + " Rechnung für " + e.leistungsString
  var empfaengerMail  = e.kunde.email
  if (e.mailAn === "mich") {
    empfaengerMail = e.profil.email

  }

  var mailArray = [e.profil.firma, e.profil.vorname, e.profil.nachname, e.profil.strassehausnummer, e.profil.plzort, e.profil.mobil, e.profil.email, e.kunde.firma, e.kunde.firma, e.kunde.vorname, e.kunde.nachname, e.kunde.geschlecht, e.kunde.email, anrede, zahlungsbedingung, e.leistungsString]
  var template = 'MailDraft'
  var attachment = response

  sendEmail(mailSubject, empfaengerMail, mailArray, template, attachment)


  // History Eintrag
    var newRow = his.getLastRow() + 1
    var entry = [[datum("yyyy-MM-dd"), e.kunde.firma, e.profil.firma,fileName, e.beschr[0], e.summe, nowPlusXDays(10, "dd.MM.yyyy")]]
    setArrayValue(his, newRow, 1, entry)



}
