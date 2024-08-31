/**
 * "SBI Contra Fund - Direct plan Growth": 119835, - https://api.mfapi.in/mf/119835/latest
 * "Quant Small Cap Fund - Direct plan Growth": 120828, - https://api.mfapi.in/mf/120828/latest
 * "ITI Small Cap Fund - Direct plan Growth": 147919, - https://api.mfapi.in/mf/147919/latest
 * "Invesco India Smallcap Fund - Direct plan Growth": 145137, - https://api.mfapi.in/mf/145137/latest
 * "Invesco India Contra Fund - Direct plan Growth": 120348, - https://api.mfapi.in/mf/120348/latest
 * "HSBC Large & Mid Cap Fund - Direct plan Growth": 146772 - https://api.mfapi.in/mf/146772/latest
 * "Quant Large & Mid Cap Fund - Direct plan Growth":120826 - https://api.mfapi.in/mf/120826/latest
 */

var dateList = {
  "currentNAV": '',
  "prevNAV": '',
  "T-3NAV": '',
  "T-4NAV": '',
  "T-1wkNAV": '',
  "T-1mNAV": '',
  "T-3mNAV": '',
  "T-6mNAV": '',
  "T-9mNAV": '',
  "T-1yNAV": '',
  "T-3yNAV": '',
  "T-5yNAV": '',
  "T-10yNAV": ''
};
var uniqueFunds = {};
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mutual Funds')

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Mutual Fund')
    .addItem('Update Data', 'updateData')
    .addToUi();
}

function updateData() {
  if (sheet.getName() != 'Mutual Funds') {
    return
  }

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var noOfRows = dataRange.getNumRows()
  var pnl_summary = {};
  var summary = ""

  //..... Progress .....
  var exeDate = new Date()
  var exeDateStr = Utilities.formatDate(exeDate, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:MM:ss');
  //..... Execution status .....
  sheet.getRange(1, 10).setValue("Started at " + String(exeDateStr)); SpreadsheetApp.flush();
  //..... Progress .....
  logStatus("Rows to analyse: " + noOfRows)

  for (var i = 8; i < noOfRows; i++) { // Skipping the header rows hence starting from row 9
    //..... Progress .....
    logStatus("Analyzing row: " + i)
    var fundName = data[i][2];
    if (fundName && !uniqueFunds[fundName]) {
      var schemeCode = getSchemeCode(fundName);
      //..... Progress .....
      logStatus("Row: " + i + " -> " + fundName + " - Scheme Code - " + schemeCode)
      if (schemeCode) {
        var latestNAV = getLatestNAV(schemeCode);

        //..... Progress .....
        logStatus("Row: " + i + " -> " + fundName + " - Latest NAV fetched = " + latestNAV.nav)

        var allNAVs = getAllNAVs(schemeCode);

        //..... Progress .....
        logStatus("Row: " + i + " -> " + fundName + " - All NAVs fetched.")

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

  buildDatesList()
  //..... Progress .....
  logStatus("Dates List created based on Anchored MF.")
  updateFunds()

  for (var i = 8; i < noOfRows; i++) { // Skipping the header rows
    var fundName = data[i][2];
    if (fundName && uniqueFunds[fundName]) {

      var currentNAV = parseFloat(uniqueFunds[fundName]["currentNAV"]);
      var currentNAVDate = dateList["currentNAV"]
      if (pnl_summary[String(currentNAVDate)] == null) {
        var funds = new Set()
        funds.add(fundName)
        pnl_summary[String(currentNAVDate)] = funds
      } else {
        pnl_summary[String(currentNAVDate)].add(fundName)
      }

      var invDateStr = sheet.getRange(i + 1, 1).getValue()
      var investedDate = new Date(String(invDateStr));
      var latestNAVDate = new Date(currentNAVDate.split('-').reverse().join('-'));
      investedDate = Utilities.formatDate(investedDate, Session.getScriptTimeZone(), 'MM-dd-yyyy')
      latestNAVDate = Utilities.formatDate(latestNAVDate, Session.getScriptTimeZone(), 'MM-dd-yyyy')
      investedDate = new Date(investedDate)
      latestNAVDate = new Date(latestNAVDate)

      //..... Progress .....
      logStatus("Row: " + (i+1) + " -> " + fundName + " values will be updated")

      //Invested Value = no of units * buy NAV
      sheet.getRange(i + 1, 6).setFormula("=D" + (i + 1) + "*E" + (i + 1));

      //STT = purchase amt - Invested Value
      sheet.getRange(i + 1, 7).setFormula("=B" + (i + 1) + "-F" + (i + 1));

      //Current NAV = API call value
      sheet.getRange(i + 1, 8).setValue(currentNAV);
      sheet.getRange(i + 1, 8).setNote(currentNAVDate);

      //Current Worth = Current NAV * no of units
      sheet.getRange(i + 1, 9).setFormula("=D" + (i + 1) + "*H" + (i + 1));

      //pnl = Current Worth - Invested Value
      sheet.getRange(i + 1, 10).setFormula("=I" + (i + 1) + "-F" + (i + 1));
      sheet.getRange(i + 1, 10).setNote(currentNAVDate);

      //pnl % = pnl / Invested Value
      sheet.getRange(i + 1, 11).setFormula("=J" + (i + 1) + "/F" + (i + 1));
      sheet.getRange(i + 1, 11).setNote(currentNAVDate);

      //NAV Diff = Current NAV - Buy NAV
      sheet.getRange(i + 1, 12).setFormula("=H" + (i + 1) + "-E" + (i + 1));
      sheet.getRange(i + 1, 12).setNote(currentNAVDate);

      // PrevNAV value
      var prevNAVDate = dateList["prevNAV"];
      var prevNAV = parseFloat(uniqueFunds[fundName]["prevNAV"]);
      sheet.getRange(i + 1, 13).setValue(prevNAV);
      sheet.getRange(i + 1, 13).setNote(prevNAVDate);

      // D'NAV diff = Current NAV - Prev NAV
      if (investedDate < latestNAVDate) {
        sheet.getRange(i + 1, 14).setFormula("=H" + (i + 1) + "-M" + (i + 1));
      } else {
        sheet.getRange(i + 1, 14).setFormula("=0");
      }

      // Day's P&L = CurrentNAV - Prev NAV * No of Units
      sheet.getRange(i + 1, 15).setFormula("=N" + (i + 1) + "*D" + (i + 1));
      sheet.getRange(i + 1, 15).setNote(currentNAVDate + " = " + currentNAV + "\n" + prevNAVDate + " = " + prevNAV);

      // Day's P&L % = (CurrentNAV - Prev NAV) / Prev NAV
      sheet.getRange(i + 1, 16).setFormula("=O" + (i + 1) + "/F" + (i + 1));
      sheet.getRange(i + 1, 16).setNote(currentNAVDate + " = " + currentNAV + "\n" + prevNAVDate + " = " + prevNAV);

      updateHistoricalData(fundName, "T-3NAV", i + 1, 17)
      updateHistoricalData(fundName, "T-4NAV", i + 1, 18)
      updateHistoricalData(fundName, "T-1wkNAV", i + 1, 19)
      updateHistoricalData(fundName, "T-1mNAV", i + 1, 20)
      updateHistoricalData(fundName, "T-3mNAV", i + 1, 21)
      updateHistoricalData(fundName, "T-6mNAV", i + 1, 22)
      updateHistoricalData(fundName, "T-9mNAV", i + 1, 23)
      updateHistoricalData(fundName, "T-1yNAV", i + 1, 24)
      updateHistoricalData(fundName, "T-3yNAV", i + 1, 25)
      updateHistoricalData(fundName, "T-5yNAV", i + 1, 26)
      updateHistoricalData(fundName, "T-10yNAV", i + 1, 27)

      //..... Progress .....
      logStatus("Row: " + (i+1) + " -> " + fundName + " update complete.")
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
  logStatus(null) //Complete

  var CptDate = new Date()
  var CptDateStr = Utilities.formatDate(CptDate, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:MM:ss');
  sheet.getRange(1, 10).setValue("Completed at " + String(CptDateStr)); SpreadsheetApp.flush();

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

  var cacheLastRow = 2;
  let cacheRange = sheet.getRange("BA2:BB").getValues();
  // Search in the cache first (BA2:BB)
  for (let i = 0; i < cacheRange.length; i++) {
    if (cacheRange[i][0] == "" && cacheRange[i][1] == "") {
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

function buildDatesList() {
  const cacheRange = sheet.getRange("BA2:BC").getValues(); // Includes Cached MF Scheme Code, Cached MF Scheme Name, Anchored MF

  var schemeName = ''
  for (var i = 0; i < cacheRange.length; i++) {
    if (cacheRange[i][2] === 'Yes') { // Check if Anchored MF is 'Yes'
      var schemeName = cacheRange[i][1]; // Get Cached MF Scheme Name
    }
  }

  if (schemeName === '') {
    sheet.getRange("BC1").setValue("Anchored MF")
    sheet.getRange("BC2").setValue("Yes")
    schemeName = sheet.getRange("BB2").getValue()
  }

  var fundData = uniqueFunds[schemeName];

  if (fundData) {
    dateList["currentNAV"] = fundData.latestNAV.date;
    dateList["prevNAV"] = fundData.allNAVs[1].date;
    dateList["T-3NAV"] = fundData.allNAVs[2].date;
    dateList["T-4NAV"] = fundData.allNAVs[3].date;
    currentNAVDate = new Date(dateList["currentNAV"].split('-').reverse().join('-'));
    fetchDate = new Date(currentNAVDate); fetchDate.setDate(currentNAVDate.getDate() - 7); dateList["T-1wkNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
    fetchDate = new Date(currentNAVDate); fetchDate.setMonth(currentNAVDate.getMonth() - 1); dateList["T-1mNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
    fetchDate = new Date(currentNAVDate); fetchDate.setMonth(currentNAVDate.getMonth() - 3); dateList["T-3mNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
    fetchDate = new Date(currentNAVDate); fetchDate.setMonth(currentNAVDate.getMonth() - 6); dateList["T-6mNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
    fetchDate = new Date(currentNAVDate); fetchDate.setMonth(currentNAVDate.getMonth() - 9); dateList["T-9mNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
    fetchDate = new Date(currentNAVDate); fetchDate.setFullYear(currentNAVDate.getFullYear() - 1); dateList["T-1yNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
    fetchDate = new Date(currentNAVDate); fetchDate.setFullYear(currentNAVDate.getFullYear() - 3); dateList["T-3yNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
    fetchDate = new Date(currentNAVDate); fetchDate.setFullYear(currentNAVDate.getFullYear() - 5); dateList["T-5yNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
    fetchDate = new Date(currentNAVDate); fetchDate.setFullYear(currentNAVDate.getFullYear() - 10); dateList["T-10yNAV"] = getDateFromResponse(fetchDate, fundData.allNAVs);
  }

}

function getDateFromResponse(targetDate, allNAVs) {
  targetDateObj = new Date(targetDate)
  for (var j = 0; j < allNAVs.length; j++) {
    var navDate = new Date(allNAVs[j].date.split('-').reverse().join('-'));
    if (navDate <= targetDateObj) {
      return allNAVs[j].date
    }
  }
  return ''
}

function updateFunds() {
  for (var fundName in uniqueFunds) {
    if (uniqueFunds.hasOwnProperty(fundName)) {
      var fundData = uniqueFunds[fundName];

      fundData["currentNAV"] = fundData.latestNAV.nav;
      fundData["prevNAV"] = getNAVForDate(dateList["prevNAV"], fundData.allNAVs);
      fundData["T-3NAV"] = getNAVForDate(dateList["T-3NAV"], fundData.allNAVs);
      fundData["T-4NAV"] = getNAVForDate(dateList["T-4NAV"], fundData.allNAVs);
      fundData["T-1wkNAV"] = getNAVForDate(dateList["T-1wkNAV"], fundData.allNAVs);
      fundData["T-1mNAV"] = getNAVForDate(dateList["T-1mNAV"], fundData.allNAVs);
      fundData["T-3mNAV"] = getNAVForDate(dateList["T-3mNAV"], fundData.allNAVs);
      fundData["T-6mNAV"] = getNAVForDate(dateList["T-6mNAV"], fundData.allNAVs);
      fundData["T-9mNAV"] = getNAVForDate(dateList["T-9mNAV"], fundData.allNAVs);
      fundData["T-1yNAV"] = getNAVForDate(dateList["T-1yNAV"], fundData.allNAVs);
      fundData["T-3yNAV"] = getNAVForDate(dateList["T-3yNAV"], fundData.allNAVs);
      fundData["T-5yNAV"] = getNAVForDate(dateList["T-5yNAV"], fundData.allNAVs);
      fundData["T-10yNAV"] = getNAVForDate(dateList["T-10yNAV"], fundData.allNAVs);
    }
  }
}

function getNAVForDate(date, allNAVs) {
  var dateObj = new Date(date.split('-').reverse().join('-'));
  for (var i = 0; i < allNAVs.length; i++) {
    var navDate = new Date(allNAVs[i].date.split('-').reverse().join('-'));
    if (navDate <= dateObj) {
      return allNAVs[i].nav;
    }
  }
  return null;
}

function updateHistoricalData(fundName, indexValue, row, col) {
  var NAV = parseFloat(uniqueFunds[fundName][indexValue]);
  if (!NAV) {
    sheet.getRange(row, col).setValue(null)
    sheet.getRange(row, col).setNote(dateList[indexValue] + "\nFUND was not started by this date.")
    return
  }

  var divider = 1
  if (indexValue === "T-3yNAV") {
    divider = 3
  } else if (indexValue === "T-5yNAV") {
      divider = 5
  } else if (indexValue === "T-10yNAV") {
      divider = 10
  }

  sheet.getRange(row, col).setFormula("=((H" + row + "-" + NAV + ")/" + NAV + ")/" + divider);
  sheet.getRange(row, col).setNote(dateList[indexValue] + "\n" + NAV)
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

function logStatus(msg) {
  sheet.getRange(2, 10).setValue(msg); 
  SpreadsheetApp.flush();  
}