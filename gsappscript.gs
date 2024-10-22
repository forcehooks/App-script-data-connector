var ui = SpreadsheetApp.getUi();
var API_TOKEN = "api.token";
var sheetkey = "Ascend Data"
var userProperties = PropertiesService.getUserProperties();
var getLocalToken = () => userProperties.getProperty(API_TOKEN);


function onOpen() {
  ui.createMenu("💪 Ascend Labs 💪")
    .addSubMenu(
      ui
        .createMenu("API Token 🎫")
        .addItem("Set API Token ✅", "setToken")
        .addItem("Delete API Token ❌", "deleteToken")

    )
    .addSubMenu(ui.createMenu("Ascend Data 📦").addItem("Load Test Data 🤼‍♂️", "getTests"))
    .addToUi();
}
function showTokenInputPopup() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('tokenInputDialog')
    .setWidth(400)
    .setHeight(200);
  ui.showModalDialog(htmlOutput, 'Enter Token');
}
function deleteToken() {
  userProperties.deleteProperty(API_TOKEN);
  ui.alert("✅API token has been removed successfully.✅");
};

function setToken() {
  var response = ui.prompt("❌Please provide your API Token.If you have api access you can find it in setting section of our web app❌", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var apiToken = response.getResponseText();
    if (validateToken(apiToken)) {
      userProperties.setProperty(API_TOKEN, apiToken);
      ui.alert("✅API token has been set successfully.✅");
    } else {
      ui.alert("❌The API token you provided is invalid. Please check and try again.❌");
    }
  }
}
function validateToken(token) {
  try {
    var options = {
      'method': 'post',
      'headers': {
        'x-access-token': token,
        'Content-Type': 'application/json',
      },
      'muteHttpExceptions': true
    };
    var response = UrlFetchApp.fetch(getUrlByIndex(1), options);
    var statusCode = response.getResponseCode();
    return statusCode === 200
  } catch (e) {
    return false

  }
}
function getUrlByIndex(index) {
  var urls = ['aHR0cHM6Ly91cy1jZW50cmFsMS1hc2NlbmQtZWJkZGEuY2xvdWRmdW5jdGlvbnMubmV0L2FwaS92MS9nZXREYXRhRm9yQXBwU2NyaXB0', 'aHR0cHM6Ly91cy1jZW50cmFsMS1hc2NlbmQtZWJkZGEuY2xvdWRmdW5jdGlvbnMubmV0L2FwaS92MS92dA=='];
  var url = urls[index]
  var dBytes = Utilities.base64Decode(url);
  var dUrl = Utilities.newBlob(dBytes).getDataAsString();
  return dUrl;
}

function getTests() {
  if (getLocalToken()) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetkey);

    var options = {
      'method': 'post',
      'headers': {
        'x-access-token': getLocalToken(),
        'Content-Type': 'application/json',
      },
      "payload": '{}',
      'muteHttpExceptions': true
    };
    if (sheet) {
      var lastRow = sheet.getLastRow(); 
      var firstColumnData = "";
      if (lastRow > 0) {
        firstColumnData = sheet.getRange(lastRow, 1).getValue();
        options.payload = JSON.stringify({
          "ldi": firstColumnData,
          'limit':500
        })
      }
    }
    try {
      var response = UrlFetchApp.fetch(getUrlByIndex(0), options);
      var data = JSON.parse(response.getContentText());
      populateData(data.data)
    } catch (e) {
      ui.alert("❌Something went wrong - getTests. Please try again if issue persists contact us❌");
    }
  } else {
    ui.alert("❌API Token has not been set.Please set token first❌");
  }
}
function populateData(grouppedData) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = spreadsheet.getSheetByName(sheetkey);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetkey);
  }
  var existingData = sheet.getDataRange().getValues();
  if (existingData.length === 0 || existingData[0].indexOf('TIME') === -1) {
    sheet.clear()
    var headers = Object.keys(grouppedData[0])
      .map(e => e.toUpperCase()); 

    sheet.appendRow(headers); 
  }
  grouppedData.forEach(function (item) {
    var row = Object.values(item);
    sheet.appendRow(row);
  });
}


