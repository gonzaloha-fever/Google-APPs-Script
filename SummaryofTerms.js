function sumterms2() {

    var dataValidation = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("[AUX] Data Validation-Emails");

    // var analysisTracker=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analysis [TRACKER]");
    var inputssot = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Analysis [Inputs]");
    // var emails =toFlatArray(dataValidation.getRange(2, 3, 19, 1).getValues());
    // var uniqueEmail=[...new Set(emails)];
    // var longUE=uniqueEmail.length;
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analysis Growth");
    var analysissheetid = ss.getSheetId();
    var analysisurl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var sheetURL2 = analysisurl + "#gid=" + analysissheetid;
    var idNewPL = DriveApp.getFileById("1Zr5lvFNP93B2GC8h2TEk_7CFU3ijyKDbezS_7anRbOg").makeCopy("Doc in creation - DO NOT OPEN").getId();
    var newPL = SpreadsheetApp.openById(idNewPL);
    var NewPLFile = DriveApp.getFileById(idNewPL);
    NewPLFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

    // var idNewPL2 = DriveApp.getFileById("13Nfto7eMeSrIpcaoAI1qy9JyL1y5BlfoXGeI1L4Amhw").makeCopy("Doc in creation - DO NOT OPEN").getId();
    // var newPL2 = SpreadsheetApp.openById(idNewPL2);
    // var NewPLFile2 = DriveApp.getFileById(idNewPL2);
    // NewPLFile2.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
    var sot_masterdoc = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1A-pOLedHa2nobxRliDDWE6Uff29CZsMygUpmProKjJk/');



    //Sugerencia de pasar esto a binario y hacerlo con un for

    var templateSummTerms0 = sot_masterdoc.getSheetByName("SOT TEMPLATE SPECIAL CASES");
    var templateSummTerms1 = sot_masterdoc.getSheetByName("SOT TEMPLATE COMPLEX");
    var templateSummTerms2 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE");
    var templateSummTerms3 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES");
    var templateSummTerms4 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O MARKETING");
    var templateSummTerms5 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O SURCHARGE");
    var templateSummTerms6 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES AND ONSITE");
    var templateSummTerms7 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND MARKETING");
    var templateSummTerms8 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND SURCHARGE");
    var templateSummTerms9 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES AND MARKETING");
    var templateSummTerms10 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES AND SURCHARGE");
    var templateSummTerms11 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O MKT AND SURCHARGE");
    var templateSummTerms12 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND ADVANCES AND MARKETING");
    var templateSummTerms13 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND ADVANCES AND SURCHARGE");
    var templateSummTerms14 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND MKT AND SURCHARGE");
    var templateSummTerms15 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES AND MKT AND SURCHARGE");
    var templateSummTerms16 = sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND ADV AND MKT AND SURCH");

    /*
    var templateSummTerms = [
        sot_masterdoc.getSheetByName("SOT TEMPLATE SPECIAL CASES"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE COMPLEX"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O MARKETING"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O SURCHARGE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES AND ONSITE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND MARKETING"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND SURCHARGE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES AND MARKETING"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES AND SURCHARGE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O MKT AND SURCHARGE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND ADVANCES AND MARKETING"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND ADVANCES AND SURCHARGE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND MKT AND SURCHARGE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ADVANCES AND MKT AND SURCHARGE"),
        sot_masterdoc.getSheetByName("SOT TEMPLATE W/O ONSITE AND ADV AND MKT AND SURCH"),
        ];
    */

    var mktquest = ss.getRange("Q6").getValue();
    var advquestion = ss.getRange("Q7").getValue();
    var surchquestion = ss.getRange("Q8").getValue();
    var onsitequestion = ss.getRange("Q9").getValue();
    var complexquestion = ss.getRange("AK9").getValue();

    if (mktquest == "" || advquestion == "" || surchquestion == "" || onsitequestion == "") {
        Browser.msgBox('Please answer all the questions to create the Summary of Terms');
    }

    if (advquestion == "Yes") {
        // var templateNewddadvances = ddadvances.copyTo(newPL2);
    }

    if (complexquestion == true) {
        var templateNewSummTerms = templateSummTerms0.copyTo(newPL);
    }

    /*
    var questions = [mktquest, advquestion, surchquestion, onsitequestion];
    var templateNewSummTerms;

    if (questions.every(q => q === "Yes")) {
        templateNewSummTerms = templateSummTerms[15].copyTo(newPL);
    } else if (questions.every(q => q !== "Yes")) {
        templateNewSummTerms = templateSummTerms[0].copyTo(newPL);
    } else {
        var index = 0;
        for (var i = 0; i < questions.length; i++) {
            if (questions[i] === "Yes") index += Math.pow(2, i);
        }
        templateNewSummTerms = templateSummTerms[index].copyTo(newPL);
    }
    */

    if (mktquest == "Yes" && advquestion == "Yes" && surchquestion == "Yes" && onsitequestion == "Yes") {
        var templateNewSummTerms = templateSummTerms1.copyTo(newPL);
    }
    else if (mktquest == "Yes" && advquestion == "Yes" && surchquestion == "Yes" && onsitequestion == "No") {
        var templateNewSummTerms = templateSummTerms2.copyTo(newPL);
    }
    else if (mktquest == "Yes" && advquestion == "No" && surchquestion == "Yes" && onsitequestion == "Yes") {
        var templateNewSummTerms = templateSummTerms3.copyTo(newPL);
    }
    else if (mktquest == "No" && advquestion == "Yes" && surchquestion == "Yes" && onsitequestion == "Yes") {
        var templateNewSummTerms = templateSummTerms4.copyTo(newPL);
    }
    else if (mktquest == "Yes" && advquestion == "Yes" && surchquestion == "No" && onsitequestion == "Yes") {
        var templateNewSummTerms = templateSummTerms5.copyTo(newPL);
    }

    else if (mktquest == "Yes" && advquestion == "No" && surchquestion == "Yes" && onsitequestion == "No") {
        var templateNewSummTerms = templateSummTerms6.copyTo(newPL);
    }

    else if (mktquest == "No" && advquestion == "Yes" && surchquestion == "Yes" && onsitequestion == "No") {
        var templateNewSummTerms = templateSummTerms7.copyTo(newPL);
    }

    else if (mktquest == "Yes" && advquestion == "Yes" && surchquestion == "No" && onsitequestion == "No") {
        var templateNewSummTerms = templateSummTerms8.copyTo(newPL);
    }

    else if (mktquest == "No" && advquestion == "No" && surchquestion == "Yes" && onsitequestion == "Yes") {
        var templateNewSummTerms = templateSummTerms9.copyTo(newPL);
    }

    else if (mktquest == "Yes" && advquestion == "No" && surchquestion == "No" && onsitequestion == "Yes") {
        var templateNewSummTerms = templateSummTerms10.copyTo(newPL);
    }

    else if (mktquest == "No" && advquestion == "Yes" && surchquestion == "No" && onsitequestion == "Yes") {
        var templateNewSummTerms = templateSummTerms11.copyTo(newPL);
    }
    else if (mktquest == "No" && advquestion == "No" && surchquestion == "Yes" && onsitequestion == "No") {
        var templateNewSummTerms = templateSummTerms12.copyTo(newPL);
    }
    else if (mktquest == "Yes" && advquestion == "No" && surchquestion == "No" && onsitequestion == "No") {
        var templateNewSummTerms = templateSummTerms13.copyTo(newPL);
    }

    else if (mktquest == "No" && advquestion == "Yes" && surchquestion == "No" && onsitequestion == "No") {
        var templateNewSummTerms = templateSummTerms14.copyTo(newPL);
    }

    else if (mktquest == "No" && advquestion == "No" && surchquestion == "No" && onsitequestion == "Yes") {
        var templateNewSummTerms = templateSummTerms15.copyTo(newPL);
    }
    else if (mktquest == "No" && advquestion == "No" && surchquestion == "No" && onsitequestion == "No") {
        var templateNewSummTerms = templateSummTerms16.copyTo(newPL);
    }



    // var auxdataValidation = dataValidation.copyTo(newPL);
    // var numberofsot=inputssot.getRange("A4").getValue();
    var currency = ss.getRange("B8").getValue();

    var projectName = ss.getRange("E3").getValue();
    var name = "[MKT] - [" + projectName + "] - Summary of Terms";
    newPL.rename(name);
    var url = newPL.getUrl();

    // var url2=newPL2.getUrl();


    var sheetName = "SOT";
    templateNewSummTerms.activate();
    templateNewSummTerms.showSheet();
    var sheetID = templateNewSummTerms.getSheetId();
    newPL.renameActiveSheet(sheetName);
    newPL.moveActiveSheet(1);

    var sot = url + "#gid=" + sheetID;
    var sotcell = ss.getRange("AJ13").setValue(sot);
    //  var sheetName2 = "DD_ADVANCES";
    //    templateNewddadvances.activate();
    // templateNewddadvances.showSheet();
    // newPL.renameActiveSheet(sheetName2);
    // newPL.moveActiveSheet(1);


    var auxdataValidation = newPL.getSheetByName("[AUX] Data Validation-Emails");
    var sheetName3 = "[AUX] Data Validation-Emails";
    auxdataValidation.activate();
    auxdataValidation.showSheet();

    newPL.renameActiveSheet(sheetName3);
    newPL.getActiveSheet().hideSheet();

    var sheet1 = newPL.getSheetByName("Sheet1");
    sheet1.activate();
    newPL.deleteActiveSheet();

    // var p2 = newPL.getSheetByName(sheetName).getRange("").protect().setDescription('Sample protected range');;
    //   p2.addEditors(uniqueEmail);
    // if (p2.canDomainEdit()) {
    //   p2.setDomainEdit(false);

    //     }

    // get data from analysis tab
    //     var sotid="SOT_" +numberofsot;
    // var sotnumber=inputssot.getRange("A4");
    // sotnumber.setValue(numberofsot+1);


    var partnername = ss.getRange("E4").getValue();
    var countryanalysis = ss.getRange("E5").getValue();
    var eventtypeanalysis = ss.getRange("E8").getValue();
    var clusteranalysis = ss.getRange("E9").getValue();
    var cityanalysis = ss.getRange("E6").getValue();
    var salesowneranalysis = ss.getRange("I3").getValue();
    var growthowneranalysis = ss.getRange("I4").getValue();
    var risklevelanalysis = ss.getRange("L6").getValue();
    var analysisanalysis = ss.getRange("B4").getValue();
    var eventstartinganalysis = ss.getRange("I6").getValue();
    var eventenddateysis = ss.getRange("I7").getValue();
    var ddidanalysis = ss.getRange("L9").getValue();
    var ddurlanalysis = ss.getRange("L10").getValue();
    var date_analysis = ss.getRange("B5").getValue();
    var partnerhistory = ss.getRange("AJ23").getValue();
    var expdescription = ss.getRange("AJ27").getValue();
    // var otherlegalterms=ss.getRange("AF13").getValue();
    var scope = ss.getRange("I9").getValue();
    var evduration = ss.getRange("I8").getValue();
    // var evergreentickets=ss.getRange("I10").getValue();
    var averageticketpricewosurchan = ss.getRange("Q11").getValue();
    var averageticketpricewsurchan = ss.getRange("Q12").getValue();
    var growthcomments = ss.getRange("AN5").getValue();

    var totalpottickets = ss.getRange("AT2").getValue();
    var onlinepottickets = ss.getRange("AT8").getValue();
    var grossrevan = ss.getRange("AT3").getValue();
    var netrevan = ss.getRange("AT4").getValue();

    var averagecommwsurchan = ss.getRange("U2").getValue();
    var averagecommwosurchan = ss.getRange("Z2").getValue();

    var summtermsranges = newPL.getSheetByName(sheetName);
    summtermsranges.getRange("E2").setValue(projectName);
    summtermsranges.getRange("C8").setValue(partnername);
    summtermsranges.getRange("C9").setValue(countryanalysis);
    summtermsranges.getRange("C10").setValue(cityanalysis);
    summtermsranges.getRange("C11").setValue(eventtypeanalysis);
    summtermsranges.getRange("E8").setValue(salesowneranalysis);
    summtermsranges.getRange("E9").setValue(growthowneranalysis);
    summtermsranges.getRange("C12").setValue(risklevelanalysis);
    summtermsranges.getRange("C16").setValue(analysisanalysis);
    summtermsranges.getRange("E11").setValue(eventstartinganalysis);
    summtermsranges.getRange("E12").setValue(eventenddateysis);
    // summtermsranges.getRange("C19").setValue(ddidanalysis);
    summtermsranges.getRange("E10").setValue(date_analysis);
    summtermsranges.getRange("U27").setValue(partnerhistory);
    summtermsranges.getRange("U31").setValue(expdescription);
    // summtermsranges.getRange("B40").setValue(otherlegalterms);
    summtermsranges.getRange("E16").setValue(sheetURL2);
    summtermsranges.getRange("C18").setValue(averageticketpricewosurchan);
    summtermsranges.getRange("E19").setValue(averagecommwsurchan);
    summtermsranges.getRange("B52").setValue(growthcomments);
    summtermsranges.getRange("D27").setValue(scope);
    summtermsranges.getRange("D28").setValue(evduration);
    // summtermsranges.getRange("D27").setValue(evergreentickets);
    summtermsranges.getRange("C21").setValue(totalpottickets);
    summtermsranges.getRange("C20").setValue(onlinepottickets);
    summtermsranges.getRange("E21").setValue(grossrevan);
    summtermsranges.getRange("E20").setValue(netrevan);

    summtermsranges.getRange("C13").setValue(clusteranalysis);
    summtermsranges.getRange("C17").setValue(ddidanalysis);
    summtermsranges.getRange("E17").setValue(ddurlanalysis);


    var ticketingcommisionan = ss.getRange("U6").getValue();
    var paymprocccommisionan = ss.getRange("U8").getValue();
    var ticketingprocccommisionan = ss.getRange("U9").getValue();
    var ticketinusdprocccommisionan = ss.getRange("U10").getValue();
    var ussuppcommisionan = ss.getRange("U11").getValue();
    var ussuppusdcommisionan = ss.getRange("U12").getValue();
    var secmediacommisionan = ss.getRange("U13").getValue();
    var secmediausdcommisionan = ss.getRange("U14").getValue();
    var paymentgatcommisionan = ss.getRange("Q15").getValue();

    // summtermsranges.getRange("K5").setValue(ticketingcommisionan);
    summtermsranges.getRange("K6").setValue(paymprocccommisionan);
    summtermsranges.getRange("K7").setValue(ticketingprocccommisionan);
    summtermsranges.getRange("K8").setValue(ticketinusdprocccommisionan);
    summtermsranges.getRange("K9").setValue(ussuppcommisionan);
    summtermsranges.getRange("K10").setValue(ussuppusdcommisionan);
    summtermsranges.getRange("K11").setValue(secmediacommisionan);
    summtermsranges.getRange("K12").setValue(secmediausdcommisionan);



    //  var summtermsranges2= newPL.getSheetByName(sheetName2);
    // summtermsranges2.getRange("C3").setValue(projectName);
    // summtermsranges2.getRange("C4").setValue(cityanalysis);
    // summtermsranges2.getRange("C5").setValue(partnername);
    // summtermsranges2.getRange("C6").setValue(analysisanalysis);
    // summtermsranges2.getRange("C7").setValue(ddidanalysis);


    // paste data in summ terms
    if (complexquestion == true) {
        var mktcommisionan = ss.getRange("Y8").getValue();
        var mktusdcommisionan = ss.getRange("Y9").getValue();
        var feetype = ss.getRange("Y12").getValue();
        var feemktusd = ss.getRange("Y13").getValue();
        var feemonths = ss.getRange("Y14").getValue();
        var feestdate = ss.getRange("Y15").getValue();
        var feeenddate = ss.getRange("Y16").getValue();
        var upfrontpaymentfee = ss.getRange("Y17").getValue();
        var recovery = ss.getRange("Y18").getValue();
        var cpa = ss.getRange("Y10").getValue();

        summtermsranges.getRange("K15").setValue(mktcommisionan);
        summtermsranges.getRange("K16").setValue(mktusdcommisionan);
        if (feetype == "One-off") {
            summtermsranges.getRange("K17").setValue(feemktusd);
        }

        else if (feetype == "Monthly") {
            summtermsranges.getRange("K18").setValue(feemktusd);
        }

        else if (feetype == "Only Media") {
            summtermsranges.getRange("K19").setValue(feemktusd);
        }

        summtermsranges.getRange("K20").setValue(feemonths);
        summtermsranges.getRange("K21").setValue(feestdate);
        summtermsranges.getRange("K22").setValue(feeenddate);
        summtermsranges.getRange("K23").setValue(upfrontpaymentfee);
        summtermsranges.getRange("K24").setValue(recovery);

        var range444 = summtermsranges.getRange("K16:M19");
        range444.setNumberFormat(currency + " #,###0.##");
        var sposorshipan = ss.getRange("AC10").getValue();
        var adticketingan = ss.getRange("AC12").getValue();
        var adpaymentan = ss.getRange("AC11").getValue();
        var adtickeitnganperc = ss.getRange("AC13").getValue();
        summtermsranges.getRange("K27").setValue(sposorshipan);
        summtermsranges.getRange("K29").setValue(adticketingan);
        summtermsranges.getRange("K31").setValue(adpaymentan);
        summtermsranges.getRange("K30").setValue(adtickeitnganperc);
        var range1238 = summtermsranges.getRange("K27");
        range1238.setNumberFormat(currency + " #,###0.##");
        var range123 = summtermsranges.getRange("K29");
        range123.setNumberFormat(currency + " #,###0.##");
        var range1234 = summtermsranges.getRange("K30");
        range1234.setNumberFormat(currency + " #,###0.##");

        var range1235 = summtermsranges.getRange("K31");
        range1235.setNumberFormat(currency + " #,###0.##");

        var range1236 = summtermsranges.getRange("K33");
        range1236.setNumberFormat(currency + " #,###0.##");

        var range1237 = summtermsranges.getRange("K35");
        range1237.setNumberFormat(currency + " #,###0.##");
        var surchargetotcommisionan = ss.getRange("Z7").getValue();
        var surchargefevercommisionan = ss.getRange("Z6").getValue();
        summtermsranges.getRange("R4").setValue(surchargetotcommisionan);
        summtermsranges.getRange("R3").setValue(surchargefevercommisionan);
        var rangesurch = summtermsranges.getRange("S6");
        rangesurch.setNumberFormat(currency + " #,###0.##");
        var onsiteticketingan = ss.getRange("AG6").getValue();
        // var onsitecommisionan=ss.getRange("AG8").getValue();
        var onsitesamequest = ss.getRange("AG8").getValue();
        var onsitemkt = ss.getRange("AG9").getValue();
        var onsitecommisionmanualan = ss.getRange("AG10").getValue();
        var onsiteexplain = ss.getRange("AG11").getValue();

        summtermsranges.getRange("R10").setValue(onsitesamequest);
        summtermsranges.getRange("R11").setValue(onsitemkt);
        summtermsranges.getRange("R12").setValue(onsitecommisionmanualan);
        summtermsranges.getRange("R13").setValue(onsiteexplain);
        var rangeonsite = summtermsranges.getRange("R17:S17");
        rangeonsite.setNumberFormat(currency + " #,###0.##");
    }
    else {
        if (mktquest == "Yes") {

            //  var mkttotcommisionan=ss.getRange("W6").getValue();
            var mktcommisionan = ss.getRange("Y8").getValue();
            var mktusdcommisionan = ss.getRange("Y9").getValue();
            var feetype = ss.getRange("Y12").getValue();
            var feemktusd = ss.getRange("Y13").getValue();
            var feemonths = ss.getRange("Y14").getValue();
            var feestdate = ss.getRange("Y15").getValue();
            var feeenddate = ss.getRange("Y16").getValue();
            var upfrontpaymentfee = ss.getRange("Y17").getValue();
            var recovery = ss.getRange("Y18").getValue();
            var cpa = ss.getRange("Y10").getValue();

            summtermsranges.getRange("K15").setValue(mktcommisionan);
            summtermsranges.getRange("K16").setValue(mktusdcommisionan);
            if (feetype == "One-off") {
                summtermsranges.getRange("K17").setValue(feemktusd);
            }

            else if (feetype == "Monthly") {
                summtermsranges.getRange("K18").setValue(feemktusd);
            }

            else if (feetype == "Only Media") {
                summtermsranges.getRange("K19").setValue(feemktusd);
            }

            summtermsranges.getRange("K20").setValue(feemonths);
            summtermsranges.getRange("K21").setValue(feestdate);
            summtermsranges.getRange("K22").setValue(feeenddate);
            summtermsranges.getRange("K23").setValue(upfrontpaymentfee);
            summtermsranges.getRange("K24").setValue(recovery);

            var mktcommisionansc2 = ss.getRange("S41").getValue();
            var mktusdcommisionansc2 = ss.getRange("T41").getValue();

            summtermsranges.getRange("L15").setValue(mktcommisionansc2);
            summtermsranges.getRange("L16").setValue(mktusdcommisionansc2);
            var feemktusdsc2 = ss.getRange("R41").getValue();
            summtermsranges.getRange("L17").setValue(feemktusdsc2);

            var range444 = summtermsranges.getRange("K16:M19");
            range444.setNumberFormat(currency + " #,###0.##");

            if (advquestion == "Yes") {
                var sposorshipan = ss.getRange("AC10").getValue();
                var adticketingan = ss.getRange("AC12").getValue();
                var adpaymentan = ss.getRange("AC11").getValue();
                var adtickeitnganperc = ss.getRange("AC13").getValue();
                summtermsranges.getRange("K27").setValue(sposorshipan);
                summtermsranges.getRange("K29").setValue(adticketingan);
                summtermsranges.getRange("K31").setValue(adpaymentan);
                summtermsranges.getRange("K30").setValue(adtickeitnganperc);
                var range1238 = summtermsranges.getRange("K27");
                range1238.setNumberFormat(currency + " #,###0.##");
                var range123 = summtermsranges.getRange("K29");
                range123.setNumberFormat(currency + " #,###0.##");
                var range1234 = summtermsranges.getRange("K30");
                range1234.setNumberFormat(currency + " #,###0.##");
                var range1235 = summtermsranges.getRange("K31");
                range1235.setNumberFormat(currency + " #,###0.##");
                var range1236 = summtermsranges.getRange("K33");
                range1236.setNumberFormat(currency + " #,###0.##");
                var range1237 = summtermsranges.getRange("K35");
                range1237.setNumberFormat(currency + " #,###0.##");

            }
        }
        else if (mktquest == "No") {

            if (advquestion == "Yes") {
                var sposorshipan = ss.getRange("AC10").getValue();
                var adticketingan = ss.getRange("AC12").getValue();
                var adpaymentan = ss.getRange("AC11").getValue();
                var adtickeitnganperc = ss.getRange("AC13").getValue();
                summtermsranges.getRange("K15").setValue(sposorshipan);
                summtermsranges.getRange("K16").setValue(adticketingan);
                summtermsranges.getRange("K19").setValue(adpaymentan);
                summtermsranges.getRange("K17").setValue(adtickeitnganperc);

                var range1238 = summtermsranges.getRange("K15");
                range1238.setNumberFormat(currency + " #,###0.##");
                var range123 = summtermsranges.getRange("K16");
                range123.setNumberFormat(currency + " #,###0.##");
                var range1234 = summtermsranges.getRange("K17");
                range1234.setNumberFormat(currency + " #,###0.##");

                var range1235 = summtermsranges.getRange("K19");
                range1235.setNumberFormat(currency + " #,###0.##");

                var range1236 = summtermsranges.getRange("K21");
                range1236.setNumberFormat(currency + " #,###0.##");

                var range1237 = summtermsranges.getRange("K23");
                range1237.setNumberFormat(currency + " #,###0.##");

            }
        }

        if (surchquestion == "Yes") {

            var surchargetotcommisionan = ss.getRange("AC7").getValue();
            var surchargefevercommisionan = ss.getRange("AC6").getValue();
            summtermsranges.getRange("R4").setValue(surchargetotcommisionan);
            summtermsranges.getRange("R3").setValue(surchargefevercommisionan);
            var rangesurch = summtermsranges.getRange("S6");
            rangesurch.setNumberFormat(currency + " #,###0.##");

            if (onsitequestion == "Yes") {

                var onsiteticketingan = ss.getRange("AG6").getValue();
                // var onsitecommisionan=ss.getRange("AG8").getValue();
                var onsitesamequest = ss.getRange("AG8").getValue();
                var onsitemkt = ss.getRange("AG9").getValue();
                var onsitecommisionmanualan = ss.getRange("AG10").getValue();
                var onsiteexplain = ss.getRange("AG11").getValue();

                summtermsranges.getRange("R10").setValue(onsitesamequest);
                summtermsranges.getRange("R11").setValue(onsitemkt);
                summtermsranges.getRange("R12").setValue(onsitecommisionmanualan);
                summtermsranges.getRange("R13").setValue(onsiteexplain);
                var rangeonsite = summtermsranges.getRange("R17:S17");
                rangeonsite.setNumberFormat(currency + " #,###0.##");
            }
        }
        else if (surchquestion == "No") {

            if (onsitequestion == "Yes") {


                var onsiteticketingan = ss.getRange("AG6").getValue();
                // var onsitecommisionan=ss.getRange("AG8").getValue();
                var onsitesamequest = ss.getRange("AG8").getValue();
                var onsitemkt = ss.getRange("AG9").getValue();
                var onsitecommisionmanualan = ss.getRange("AG10").getValue();
                var onsiteexplain = ss.getRange("AG11").getValue();
                summtermsranges.getRange("R3").setValue(onsitesamequest);
                summtermsranges.getRange("R4").setValue(onsitemkt);
                summtermsranges.getRange("R5").setValue(onsitecommisionmanualan);
                summtermsranges.getRange("R6").setValue(onsiteexplain);
                var rangeonsite = summtermsranges.getRange("R10:S10");
                rangeonsite.setNumberFormat(currency + " #,###0.##");
            }
        }
    }

    var frequencypayments = summtermsranges.getRange("D32");

    var rule1freq = SpreadsheetApp.newDataValidation().requireValueInList(['Weekly', 'Bi-weekly', 'Monthly', 'Other']).build();
    var rule2freq = SpreadsheetApp.newDataValidation().requireValueInList(['Monthly', 'Other']).build();

    if (eventtypeanalysis == "Bronze") {
        frequencypayments.setDataValidation(rule2freq);

    }
    else {

        frequencypayments.setDataValidation(rule1freq);
    }



    var sot_masterdoc = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1A-pOLedHa2nobxRliDDWE6Uff29CZsMygUpmProKjJk/');
    var tracker = sot_masterdoc.getSheetByName('[Tracker] - SOT');
    var Avals = tracker.getRange("A1:A").getValues();
    var Alast = Avals.filter(String).length;
    var lastcolumn = tracker.getLastColumn();
    var alast2 = Alast + 1;
    var numbersotid = sot_masterdoc.getSheetByName("[Tracker] - SOT").getRange(Alast + 1, 11).getValue();
    var sotid = "SOT_" + numbersotid;
    summtermsranges.getRange("C2").setValue(sotid);
    var tracker = sot_masterdoc.getSheetByName('[Tracker] - SOT');
    var numbersotid = tracker.getRange(Alast + 1, 11).getValue();
    var sotid = "SOT_" + numbersotid;
    summtermsranges.getRange("C2").setValue(sotid);
    var sot_ID = tracker.getRange(Alast + 1, 1);
    sot_ID.setValue(sotid);
    var an_ID = tracker.getRange(Alast + 1, 2);
    an_ID.setValue(analysisanalysis);
    var oppt_name = tracker.getRange(Alast + 1, 3);
    oppt_name.setValue(projectName);
    var sumurl = tracker.getRange(Alast + 1, 4);
    sumurl.setValue(sot);
    var sumstatus = tracker.getRange(Alast + 1, 5);
    sumstatus.setFormula(`=if(and(G` + alast2 + `="Pending", I` + alast2 + `="Pending"), "Pending", if(and(G` + alast2 + `="Validated", I` + alast2 + `="Validated"),"Validated","Pending"))`);
    var sumnvstatus = tracker.getRange(Alast + 1, 7);
    sumnvstatus.setFormula(`=if(F` + alast2 + `="","Pending","Validated")`);
    var sumgrowthstatus = tracker.getRange(Alast + 1, 9);
    sumgrowthstatus.setFormula(`=if(H` + alast2 + `="","Pending","Validated")`);


    tracker.getRange(1, 1, Alast + 1, lastcolumn).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

    // if (advquestion==Yes){


    //   var adv_masterdoc = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/16xeBEa-w0KdAwHz67hYTSh3_XVYffYKsstIlGeQhGp0/');
    //  var trackeradv = adv_masterdoc.getSheetByName('[Tracker]');
    //   var Avalsadv = trackeradv.getRange("A1:A").getValues();
    //   var Alastadv = Avalsadv.filter(String).length;
    //   var lastcolumnadv=trackeradv.getLastColumn();
    //   var alast2adv=Alastadv + 1;

    // var adv_IDadv = trackeradv.getRange(Alastadv + 1, 1);
    //   adv_IDadv.setValue();
    //   var an_IDadv = trackeradv.getRange(Alastadv + 1, 2);
    //   an_IDadv.setValue(analysisanalysis);

    //   var sot_IDadv = trackeradv.getRange(Alastadv + 1, 3);
    //   sot_IDadv.setValue(sotid);
    // var oppt_nameadv = trackeradv.getRange(Alastadv + 1, 4);
    //   oppt_nameadv.setValue(projectName);
    //   var advurladv = trackeradv.getRange(Alastadv + 1, 5);
    //   advurladv.setValue(url2);

    //     var salesowneradv = trackeradv.getRange(Alastadv + 1, 6);
    //   salesowneradv.setValue(salesowneranalysis);
    //   var growthowneradv = trackeradv.getRange(Alastadv + 1, 7);
    //   growthowneradv.setValue(growthowneranalysis);
    // var nvowneradv = trackeradv.getRange(Alastadv + 1, 8);
    //   nvowneradv.setValue();
    //   var soturladv = trackeradv.getRange(Alastadv + 1, 9);
    //   soturladv.setValue(url);




    // tracker.getRange(1,1,Alast+1,lastcolumn).setBorder(true, true, true,true,true,true,'#000000', SpreadsheetApp.BorderStyle.SOLID);
    // }

}


/*

    trackersheet.getRange(i + 1, 6).setValue(rangecopy1);

    trackersheet.getRange(i + 1, 7).setValue(rangecopy2);

    trackersheet.getRange(i + 1, 9).setValue(rangecopy3);

    trackersheet.getRange(i + 1, 10).setValue(rangecopy4);

    trackersheet.getRange(i + 1, 11).setValue(rangecopy5);

    trackersheet.getRange(i + 1, 15).setValue(rangecopy6);

    trackersheet.getRange(i + 1, 16).setValue(rangecopy7);

    trackersheet.getRange(i + 1, 18).setValue(rangecopy8);

    trackersheet.getRange(i + 1, 19).setValue(rangecopy9);

    trackersheet.getRange(i + 1, 20).setValue(rangecopy10);

    trackersheet.getRange(i + 1, 23).setValue(rangecopy11);

    trackersheet.getRange(i + 1, 25).setValue(rangecopy12);

    trackersheet.getRange(i + 1, 34).setValue(rangecopy13);

    trackersheet.getRange(i + 1, 35).setValue(rangecopy14);

    trackersheet.getRange(i + 1, 36).setValue(rangecopy15);

    trackersheet.getRange(i + 1, 37).setValue(rangecopy16);

    trackersheet.getRange(i + 1, 40).setValue(rangecopy17);

    trackersheet.getRange(i + 1, 41).setValue(rangecopy18);

    trackersheet.getRange(i + 1, 42).setValue(rangecopy19);

    trackersheet.getRange(i + 1, 47).setValue(rangecopy20);

    trackersheet.getRange(i + 1, 48).setValue(rangecopy21);

    trackersheet.getRange(i + 1, 49).setValue(rangecopy22);

    trackersheet.getRange(i + 1, 53).setValue(rangecopy23);


*/

