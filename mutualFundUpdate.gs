/**
 * "SBI Contra Fund - Direct plan Growth": 119835, - https://api.mfapi.in/mf/119835/latest
 * "Quant Small Cap Fund - Direct plan Growth": 120828, - https://api.mfapi.in/mf/120828/latest
 * "ITI Small Cap Fund - Direct plan Growth": 147919, - https://api.mfapi.in/mf/147919/latest
 * "Invesco India Smallcap Fund - Direct plan Growth": 145137, - https://api.mfapi.in/mf/145137/latest
 * "Invesco India Contra Fund - Direct plan Growth": 120348, - https://api.mfapi.in/mf/120348/latest
 * "HSBC Large & Mid Cap Fund - Direct plan Growth": 146772 - https://api.mfapi.in/mf/146772/latest
 * "Quant Large & Mid Cap Fund - Direct plan Growth":120826 - https://api.mfapi.in/mf/120826/latest
 */


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Mutual Fund')
      .addItem('Update Data', 'updateData')
      .addToUi();
}

function updateData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mutual Funds')

  if (sheet.getName() != 'Mutual Funds') {
    return
  }

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var noOfRows = dataRange.getNumRows()
  var uniqueFunds = {};
  var pnl_summary = {};
  var summary = ""
  
  //..... Progress .....
  var exeDate = new Date()
  var exeDateStr = Utilities.formatDate(exeDate, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:MM:ss');
  //..... Execution status .....
  sheet.getRange(1, 12).setValue("Execution started at: " + String(exeDateStr)); SpreadsheetApp.flush();
  //..... Progress .....
  sheet.getRange(2, 10).setValue("Rows to analyse: " + noOfRows); SpreadsheetApp.flush();

  for (var i = 8; i < noOfRows; i++) { // Skipping the header rows hence starting from row 9
    //..... Progress .....
    sheet.getRange(2, 10).setValue("Processing row: " + i); SpreadsheetApp.flush();
    var fundName = data[i][2];
    if (fundName && !uniqueFunds[fundName]) {
      var schemeCode = getSchemeCode(fundName);
      //..... Progress .....
      sheet.getRange(2, 10).setValue(fundName + " - Scheme Code - " + schemeCode); SpreadsheetApp.flush();
      if (schemeCode) {
        var latestNAV = getLatestNAV(schemeCode);
        
        //..... Progress .....
        sheet.getRange(2, 10).setValue(fundName + " - Latest NAV: " + latestNAV.nav); SpreadsheetApp.flush();

        var allNAVs = getAllNAVs(schemeCode);
        
        //..... Progress .....
        sheet.getRange(2, 10).setValue(fundName + " - All NAVs fetched."); SpreadsheetApp.flush();
        
        uniqueFunds[fundName] = {
          schemeCode: schemeCode,
          latestNAV: latestNAV,
          allNAVs: allNAVs
        };
      } else {
        uniqueFunds[fundName] = null;
      }
    }
  }
  
  for (var i = 8; i < noOfRows; i++) { // Skipping the header rows
    var fundName = data[i][2];
    if (fundName && uniqueFunds[fundName]) {
      var schemeCode = uniqueFunds[fundName].schemeCode;
      var latestNAV = uniqueFunds[fundName].latestNAV;
      var allNAVs = uniqueFunds[fundName].allNAVs;

      var currentNAV = parseFloat(latestNAV.nav);
      if (pnl_summary[String(latestNAV.date)] == null) {
        var funds = new Set()
        funds.add(fundName)
        pnl_summary[String(latestNAV.date)] = funds
      } else {
        pnl_summary[String(latestNAV.date)].add(fundName)
      }

      var invDateStr = sheet.getRange(i + 1, 1).getValue()
      var investedDate = new Date(String(invDateStr));
      var latestNAVDate = new Date(latestNAV.date.split('-').reverse().join('-'));
      investedDate = Utilities.formatDate(investedDate, Session.getScriptTimeZone(), 'MM-dd-yyyy')
      latestNAVDate = Utilities.formatDate(latestNAVDate, Session.getScriptTimeZone(), 'MM-dd-yyyy')
      investedDate = new Date(investedDate)
      latestNAVDate = new Date(latestNAVDate)
      var refDate = String(latestNAV.date);
      
      //..... Progress .....
      sheet.getRange(2, 10).setValue(fundName + " - Row #: " + (i+1) + " is getting updated."); SpreadsheetApp.flush();

      //Invested Value = no of units * buy NAV
      sheet.getRange(i + 1, 6).setFormula("=D"+ (i+1) + "*E" + (i+1));
      
      //STT = purchase amt - Invested Value
      sheet.getRange(i + 1, 7).setFormula("=B"+ (i+1) + "-F"+ (i+1));
      
      //Current NAV = API call value
      sheet.getRange(i + 1, 8).setValue(currentNAV);
      sheet.getRange(i + 1, 8).setNote(String(latestNAV.date));
      
      //Current Worth = Current NAV * no of units
      sheet.getRange(i + 1, 9).setFormula("=D"+ (i+1) + "*H" + (i+1));

      //pnl = Current Worth - Invested Value
      sheet.getRange(i + 1, 10).setFormula("=I"+ (i+1) + "-F" + (i+1));
      sheet.getRange(i + 1, 10).setNote(String(latestNAV.date));

      //pnl % = pnl / Invested Value
      sheet.getRange(i + 1, 11).setFormula("=J"+ (i+1) + "/F" + (i+1));
      sheet.getRange(i + 1, 11).setNote(String(latestNAV.date));

      //NAV Diff = Current NAV - Buy NAV
      sheet.getRange(i + 1, 12).setFormula("=H"+ (i+1) + "-E" + (i+1));
      sheet.getRange(i + 1, 12).setNote(String(latestNAV.date));

      // PrevNAV value and date = API Call value
      setNearestNavValue(sheet, allNAVs, i + 1, 13, latestNAV.date, 1);
      var refDate = sheet.getRange(i + 1, 13).getNote();
      var prevNAV = parseFloat(sheet.getRange(i + 1, 13).getValue());

      // D'NAV diff = Current NAV - Prev NAV
      if ( investedDate < latestNAVDate) {
        sheet.getRange(i + 1, 14).setFormula("=H"+ (i+1) + "-M" + (i+1));
      } else {
        sheet.getRange(i + 1, 14).setFormula("=0");
      }

      // Day's P&L = CurrentNAV - Prev NAV * No of Units
      sheet.getRange(i + 1, 15).setFormula("=N"+ (i+1) + "*D" + (i+1));
      sheet.getRange(i + 1, 15).setNote(String(latestNAV.date));

      // Day's P&L % = (CurrentNAV - Prev NAV) / Prev NAV
      sheet.getRange(i + 1, 16).setFormula("=O"+ (i+1) + "/F" + (i+1));
      sheet.getRange(i + 1, 16).setNote(refDate + "\n" + prevNAV);

      // last T-2 = API call value
      setNearestNavPercentage(sheet, allNAVs, i + 1, 17, latestNAV.date, 2, currentNAV);

      // last T-3 = API call value
      setNearestNavPercentage(sheet, allNAVs, i + 1, 18, latestNAV.date, 3, currentNAV);

      // Populate percentage change values for different days and durations = all API call value
      setNearestNavPercentage(sheet, allNAVs, i + 1, 19, latestNAV.date, 6, currentNAV); // last T-1wk
      setNearestNavPercentage(sheet, allNAVs, i + 1, 20, getDateNDaysAgo(30), 0, currentNAV); // last T-1M
      setNearestNavPercentage(sheet, allNAVs, i + 1, 21, getDateNDaysAgo(90), 0, currentNAV); // last T-3M
      setNearestNavPercentage(sheet, allNAVs, i + 1, 22, getDateNDaysAgo(180), 0, currentNAV); // last T-6M
      setNearestNavPercentage(sheet, allNAVs, i + 1, 23, getDateNDaysAgo(365), 0, currentNAV); // last T-1yr
      setNearestNavPercentage(sheet, allNAVs, i + 1, 24, getDateNDaysAgo(1095), 0, currentNAV); // last T-3yr
      setNearestNavPercentage(sheet, allNAVs, i + 1, 25, getDateNDaysAgo(1826), 0, currentNAV); // last T-5yr
      setNearestNavPercentage(sheet, allNAVs, i + 1, 26, getDateNDaysAgo(3652), 0, currentNAV); // last T-10yr

      //..... Progress .....
      sheet.getRange(2, 10).setValue(fundName + " - Row #: " + (i+1) + " update complete."); SpreadsheetApp.flush();
    }
  }
  //update Net P&L notes
  for (let item in pnl_summary) {
    var funds = ""
    for (let fund of pnl_summary[item]) {
      funds += fund + "\n";
    }
    summary = summary + item + "\n" + funds + "\n"
  }
  sheet.getRange(4, 2).setNote(summary); SpreadsheetApp.flush();
  sheet.getRange(5, 2).setNote(summary); SpreadsheetApp.flush();
  sheet.getRange(4, 10).setNote(summary); SpreadsheetApp.flush();
  sheet.getRange(5, 10).setNote(summary); SpreadsheetApp.flush();
  //..... Progress .....
  sheet.getRange(2, 10).setValue(null); SpreadsheetApp.flush();

  var CptDate = new Date()
  var CptDateStr = Utilities.formatDate(CptDate, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:MM:ss');
  sheet.getRange(1, 12).setValue("Execution completed at: " + String(CptDateStr)); SpreadsheetApp.flush();
  CptDateStr = Utilities.formatDate(CptDate, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  sheet.getRange(1, 10).setValue(String(CptDateStr)); SpreadsheetApp.flush();
}

function getSchemeCode(fundName) {
  var schemeCode = "";
  var fund_name = "";
  var plan_type = "";
  var option = "";

  // Split the fundName into components: fund_name, plan_type, option
  schemeName = fundName.toLowerCase();
  const parts = schemeName.match(/(.+) fund - (.+) plan (.+)/i);
  if (parts && parts.length === 4) {
    fund_name = parts[1].toLowerCase();
    plan_type = parts[2].toLowerCase();
    option = parts[3].toLowerCase();
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mutual Funds')
  var cacheLastRow = 2;
  let cacheRange = sheet.getRange("BA2:BB").getValues();
  // Search in the cache first (BA2:BB)
  for (let i = 0; i < cacheRange.length; i++) {
    if(cacheRange[i][0] == "" && cacheRange[i][1] == "") {
      cacheLastRow = i + 2;
      break;
    }
    const scheme = cacheRange[i][1].toLowerCase(); // Cached Scheme Name
    if (scheme === schemeName || (scheme.includes(fund_name) && scheme.includes(plan_type) && scheme.includes(option))) {
      schemeCode = cacheRange[i][0];
      return schemeCode; // Return Cached Scheme Code
    }
  }

  var url = 'https://api.mfapi.in/mf';
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());

  // Find the scheme code by matching all components
  for (let i = 0; i < data.length; i++) {
    const scheme = data[i].schemeName.toLowerCase();
    if (scheme === schemeName || (scheme.includes(fund_name) && scheme.includes(plan_type) && scheme.includes(option))) {
      schemeCode = data[i].schemeCode;
      sheet.getRange("BA" + cacheLastRow + ":BB" + cacheLastRow).setValues([[data[i].schemeCode, fundName]]);
      SpreadsheetApp.flush();
      return schemeCode;
    }
  }

  if (!schemeCode) {
    return 'Error: Fund name not found';
  }
}

function getLatestNAV(schemeCode) {
  var url = 'https://api.mfapi.in/mf/' + schemeCode + '/latest';
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  return data.data[0];
}

function getAllNAVs(schemeCode) {
  var url = 'https://api.mfapi.in/mf/' + schemeCode;
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  // Sort the NAV data in descending order by date
  data.data.sort((a, b) => new Date(b.date.split('-').reverse().join('-')) - new Date(a.date.split('-').reverse().join('-')));
  return data.data;
}

function getDateNDaysAgo(days) {
  var date = new Date();
  date.setDate(date.getDate() - days);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MM-yyyy');
}

function setNearestNavValue(sheet, allNAVs, row, col, referenceDate, daysAgo = 0) {
  var targetDateObj = new Date(referenceDate.split('-').reverse().join('-'));
  targetDateObj.setDate(targetDateObj.getDate() - daysAgo);

  for (var j = 0; j < allNAVs.length; j++) {
    var navDate = new Date(allNAVs[j].date.split('-').reverse().join('-'));
    if (navDate <= targetDateObj) {
      sheet.getRange(row, col).setValue(parseFloat(allNAVs[j].nav));
      sheet.getRange(row, col).setNote(String(allNAVs[j].date));
      break;
    }
  }
}

function setNearestNavPercentage(sheet, allNAVs, row, col, referenceDate, daysAgo, currentNAV) {
  var targetDateObj = new Date(referenceDate.split('-').reverse().join('-'));
  targetDateObj.setDate(targetDateObj.getDate() - daysAgo);

  for (var j = 0; j < allNAVs.length; j++) {
    var navDate = new Date(allNAVs[j].date.split('-').reverse().join('-'));
    if (navDate <= targetDateObj) {
      var previousNAV = parseFloat(allNAVs[j].nav);
      
      // Calculate the difference in years
      var todayDate = new Date();
      var timeDiff = todayDate - navDate;
      var yearsDiff = parseInt(timeDiff / (1000 * 3600 * 24 * 365)); // Convert milliseconds to years

      // Ensure that yearsDiff is not zero to avoid division by zero
      if (yearsDiff > 0) {
        var annualizedReturn = parseFloat(((currentNAV - previousNAV) / previousNAV) / yearsDiff);
      } else {
        var annualizedReturn = parseFloat((currentNAV - previousNAV) / previousNAV); 
      }
      sheet.getRange(row, col).setValue(annualizedReturn);
      sheet.getRange(row, col).setNote(String(allNAVs[j].date) + "\n" + String(allNAVs[j].nav));
      break;
    }
  }
}
