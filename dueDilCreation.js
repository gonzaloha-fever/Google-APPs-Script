function dueDilCreation() {
    // Variable declarations
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentPartnerTemplateSheet = ss.getSheetByName("current partner template");
    const newPartnerTemplateSheet = ss.getSheetByName("new partner template");
    const inputsSheet = ss.getSheetByName("New Due Dilligence [Inputs]");
  
    //Get data from the Inputs sheet
  
    const country = inputsSheet.getRange("C4").getValue();
    const cityName = inputsSheet.getRange("C5").getValue();
    const cityCode = inputsSheet.getRange("C6").getValues();
    const oppIDs = inputsSheet.getRange("G3").getValue();
    const projectName = inputsSheet.getRange("G4").getValue().replace("'", "''") || '';
    const entityName = inputsSheet.getRange("G5").getValue();
    const salesOwner = inputsSheet.getRange("G6").getValue();
    const newPartner = inputsSheet.getRange("G7").getValue();
    const storeCodeNumber = inputsSheet.getRange("A3").getValue();

    //Comprobar que se han introducido todos los datos

    if(country == "Select a country" || cityName == "Select a city" || projectName == null || salesOwner == null){

        SpreadsheetApp.getUi().alert("Not all inputs introduced:\n Country \n City \n Project Name \ Sales Owner");
        //En teoría esto hace que pare la función
        return;
    }
  
    const ddID = `dd_${String(storeCodeNumber)}`;
    const dateDue = new Date();
    const sheetName = `[${ddID}] - [${entityName}] - [${salesOwner}] - [${cityCode}]`;
  
    const sheet = newPartner === "Yes"
      ? ss.insertSheet(sheetName, inputsSheet.getIndex(), { template: newPartnerTemplateSheet })
      : ss.insertSheet(sheetName, inputsSheet.getIndex(), { template: currentPartnerTemplateSheet });
    sheet.showSheet();
  
    inputsSheet.getRange("G12").setValue("Creating Sheet - Please Wait");
  
    // idSheet
    sheet.getRange("B2").setValue(oppIDs);
    // Project name
    sheet.getRange("B3").setValue(projectName);
    // Entity Name
    sheet.getRange("B4").setValue(entityName);
    // Country
    sheet.getRange("B5").setValue(country);
    // City
    sheet.getRange("B6").setValue(cityName);
    // Stored Code
    sheet.getRange("B7").setValue(ddID);
  
    const sheetID = sheet.getSheetId();
    const URL = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    const sheetURL = `${URL}#gid=${sheetID}`;
  
    // Clear the inputs cells
    inputsSheet.getRange("C6").clearContent();
    inputsSheet.getRange("G3").clearContent();
    inputsSheet.getRange("G4").clearContent();
    inputsSheet.getRange("G5").clearContent();
    inputsSheet.getRange("G6").clearContent();
    inputsSheet.getRange("G7").clearContent();
    inputsSheet.getRange("C4").setValue("Select a country");
    inputsSheet.getRange("C5").setValue("Select a city");
    inputsSheet.getRange("A3").setValue(storeCodeNumber + 1);
  
    // Put the URL in the Inputs tab
    inputsSheet.getRange("G12").setValue(sheetURL);
  
    // Fill the Doc tracker in the Due Diligence DOC 
  
    const docTracker = ss.getSheetByName("Due Diligences [TRACKER]");
  
    // Get the last row of the sheet
    const lastRow = docTracker.getLastRow();
    const lastColumn = docTracker.getLastColumn();
  
    // Define the data to be written to the sheet
    const dataFillTracker = [
      { column: 1, value: ddID },
      { column: 2, value: oppIDs},
      { column: 3, value: dateDue },
      { column: 4, value: `='${sheetName}'!D11` },
      { column: 5, value: country },
      { column: 6, value: cityName },
      { column: 7, value: projectName },
      { column: 8, value: entityName },
      { column: 9, value:`='${sheetName}'!D21`},
      { column: 16, value: sheetName },
      { column: 12, value: salesOwner },
      { column: 14, value: sheetURL },
      { column: 15, value: `='${sheetName}'!D18` },
      { column: 20, value: `='${sheetName}'!D11` },
      { column: 21, value: `='${sheetName}'!Q220` },
      { column: 22, value: `='${sheetName}'!Q221` },
      { column: 24, value: `='${sheetName}'!H3` },
      { column: 25, formula: `=iferror(VLOOKUP(F` + lastRow + 1 +`,Aux_City_Codes!$B$2:$G,6,0),"")`}
    ];
  
    // Write the data to the sheet
    dataFillTracker.forEach(({ column, value, formula }) => {
      if (formula) {
        docTracker.getRange(lastRow + 1, column).setFormula(formula);
      } else {
        docTracker.getRange(lastRow + 1, column).setValue(value);
      }
    });
  
    if (newPartner == "Yes") {
      //Coinv
      docTracker.getRange(lastRow + 1, 10).setValue(`='${sheetName}'!D25`);
      //CoinvYes
      docTracker.getRange(lastRow + 1, 11).setValue(`='${sheetName}'!D23`);
    }else {
      docTracker.getRange(lastRow + 1, 9).setValue(`='${sheetName}'!D32`);
      docTracker.getRange(lastRow + 1, 10).setValue(`='${sheetName}'!D30`);
    }
  
    //Pensar cómo se puede hacer esto para que quede más cerrado.
  
    //Retoque estético
    docTracker.getRange(1, 1, lastRow + 1, lastColumn).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
    colorCountryTab(sheet,country);
  
    //pick growth owner: Hay que cambiar de dónde se recoge esto
    var growthOwner = docTracker.getRange(lastRow + 1, 13).getValue();
  
  
    //Fill the DUe Diligences tracker from the Analysis DOC
    const analysisDoc = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1Wo88tHJMopWhk6WTLb4igC1PrwNB7pEhpLihBmwDO1M/');
    const trackerAnalysis = analysisDoc.getSheetByName('Due Diligences [TRACKER]');
  
    // Get the last row with data in column A
    avals = trackerAnalysis.getRange("A1451:A").getValues();
    alast = avals.filter(String).length +1450;
    const alastNext = alast + 1;
  
    // Get the last column with data in the sheet
    const lastCol = trackerAnalysis.getLastColumn();
  
    // Set values and formulas for the new row in the Due Diligences Tracker
    const dataFillTrackerAnalysisDoc = [
      { column: 1, value: ddID },
      { column: 2, value: dateDue },
      { column: 3, value: country },
      { column: 4, value: cityName },
      { column: 5, value: projectName },
      { column: 6, value: oppIDs },
      { column: 7, formula: `=if(E${alastNext}<>"",text(sum(COUNTIF(trim(T${alastNext}),$P${alastNext}),COUNTIF(trim(U${alastNext}),$Q${alastNext}),COUNTIF(trim(V${alastNext}),$R${alastNext}),COUNTIF(trim(W${alastNext}),$S${alastNext})),"0"),"")` },
      { column: 8, formula: `=text(SUM( COUNTIF(TRIM(T${alastNext}),"<>EMPTY"),COUNTIF(TRIM(U${alastNext}),"<>EMPTY"),COUNTIF(TRIM(V${alastNext}),"<>EMPTY"),COUNTIF(TRIM(W${alastNext}),"<>EMPTY")),"0")` },
      { column: 9, value: sheetURL },
      { column: 10, value: salesOwner },
      { column: 11, value: growthOwner },
      { column: 24, formula: `=if(E` + alastNext +`<>"",divide(sum(COUNTIF(trim(T` + alastNext +`),$P` + alastNext +`),COUNTIF(trim(U` + alastNext +`),$Q` + alastNext +`),COUNTIF(trim(V` + alastNext +`),$R` + alastNext +`),COUNTIF(trim(W` + alastNext +`),$S` + alastNext +`)), SUM( COUNTIF(TRIM(T` + alastNext +`),"<>EMPTY"),COUNTIF(TRIM(U` + alastNext +`),"<>EMPTY"),COUNTIF(TRIM(V` + alastNext +`),"<>EMPTY"),COUNTIF(TRIM(W` + alastNext +`),"<>EMPTY"))))`},
      { column: 25, formula: `=iferror(VLOOKUP(A` + alastNext +`,'Copy of [AUX] DD Data'!$A$2:$E,2,0),"")`},
      { column: 26, formula: `=iferror(VLOOKUP(A` + alastNext +`,'Copy of [AUX] DD Data'!$A$2:$E,3,0),"")`},
      { column: 27, formula: `=iferror(VLOOKUP(A` + alastNext +`,'Copy of [AUX] DD Data'!$A$2:$E,4,0),"")`},
      { column: 28, formula: `=iferror(VLOOKUP(A` + alastNext +`,'Copy of [AUX] DD Data'!$A$2:$E,5,0),"")`},
      { column: 29, formula: `=iferror(VLOOKUP(A` + alastNext +`,'Copy of [AUX] DD Data'!$A$2:$F,6,0),"")`},
      { column: 30, formula: `=iferror(VLOOKUP(F` + alastNext +`,'[AUX] Lead Gen'!$A$2:$D,4,0),"")`},
      { column: 31, formula: `=iferror(VLOOKUP(F` + alastNext +`,'[AUX] Lead Gen'!$A$2:$D,3,0),"")`},
      { column: 32, formula: `=iferror(VLOOKUP(A` + alastNext +`,'Copy of [AUX] DD Data'!$A$2:$G,7,0),"")`}
    ];
  
    dataFillTrackerAnalysisDoc.forEach(({ column, value, formula }) => {
      if (formula) {
        trackerAnalysis.getRange(alast + 1, column).setFormula(formula);
      } else {
        trackerAnalysis.getRange(alast + 1, column).setValue(value);
      }
    });
    
    trackerAnalysis.getRange(1, 1, alast + 1, lastCol).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }
  
  
  function colorCountryTab(tab,country) {
    /*
    Esta función necesita como inputs la tab y el país, y directamente colorea la tab
    */
    const countryColors = {
    Argentina: "#00FF00",//Argentina: Green
    Australia: "#0000FF",//Australia: Blue
    Austria: "#FFFF00",//Austria: Yellow
    Belgium: "#FF00FF",//Belgium: Fuchsia
    Brazil: "#00FFFF",//Brazil: Aqua
    Canada: "#FF0000",//Canada: Red
    Chile: "#800000",//Chile: Maroon
    "Costa Rica": "#008000",//Costa Rica: Olive
    Denmark: "#000080",//Denmark: Navy
    France: "#808000",//France: Olive
    Germany: "#800080",//Germany: Purple
    Ireland: "#008080",//Ireland: Teal
    Italy: "#808080",//Italy: Gray
    Japan: "#C0C0C0",//Japan: Silver
    "Korea (Republic of)": "#FFC0CB",//Korea (Republic of): Pink
    Mexico: "#FFA07A",//Mexico: Light Salmon
    Morocco: "#FFD700",//Morocco: Gold
    Netherlands: "#DAA520",//Netherlands: Goldenrod
    "New Zealand": "#ADD8E6",//New Zealand: Light Blue
    Portugal: "#98FB98",//Portugal: Pale Green
    "Russian Federation": "#F0E68C",//Russian Federation: Khaki
    Singapore: "#E6E6FA",//Singapore: Lavender
    "South Africa": "#D3D3D3",//South Africa: Light Gray
    Spain: "#A52A2A",//Spain: Brown
    Sweden: "#FF69B4",//Sweden: Hot Pink
    Switzerland: "#FF1493",//Switzerland: Deep Pink
    "United Arab Emirates": "#DB7093",//United Arab Emirates: Pale Violet Red
    "United Kingdom": "#B0C4DE",//United Kingdom: Light Steel Blue
    "United States": "#87CEEB"//United States: Sky Blue
    };
    tab.setTabColor(countryColors[country] || "#000000");
  }
  
  
  
  