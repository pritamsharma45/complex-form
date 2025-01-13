function SetUpSheet(sheetsArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName("template");
  schema.forEach((name) => {
    if (ss.getSheetByName(name) == null) {
      ss.insertSheet(name, { template: templateSheet });
    } else {
      return ss.getSheetByName(name);
    }
  });

  // Set Named Range
  schema.forEach((name) => {
    if (ss.getRangeByName(name) == null) {
      ss.setNamedRange(
        "tbl_" + name,
        ss.getSheetByName(name).getRange("A1:D1000")
      );
    }
  });
}

function getResponseSheet(name) {
  if (ss.getSheetByName(name) == null) {
    var sht = ss.insertSheet(name);
      sht.getRange("A1:E1").setValues([["id", "model", "files", "created_date", "client_email"]]);
       ss.setNamedRange(
        "tbl_" + name,
        sht.getRange("A1:E")
      );
      return sht;


  } else {
    return ss.getSheetByName(name);
  }
}

function _GetSheetByName(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName("template");
  if (ss.getSheetByName(name) == null) {
    ss.insertSheet(name, { template: templateSheet });
  } else {
    return ss.getSheetByName(name);
  }
}

function _GetNamedRange(name) {
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(name);
  if (range != null) {
    Logger.log(range.getNumColumns());
  }
}

function _DeleteSheets(arrayOfSheetNames) {
  var ss = SpreadsheetApp.getActive();
  arrayOfSheetNames.forEach((name) => {
    var sheet = ss.getSheetByName(name);
    if (sheet != null) {
      ss.deleteSheet(sheet);
    }
  });
}

function fssfD() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setNamedRange(
    "tbl_etcs_po_request",
    spreadsheet.getRange("A1:B1000")
  );
}
