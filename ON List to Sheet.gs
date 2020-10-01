/* Copyright (C) 2020  Hanalani Schools

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.

*/

var YOURKEY = "your_key";
var AUTHSHEETID = "your_sky_api_spreadsheet_id";
var YOURSTATE = "your_state_string";

function doGet(e) {
  var param = e.queryString
  var parameters = param.split("&")
  if (param != null){
    param = e.parameter
    var listID = param.listID
    var spreadsheetID = param.spreadsheetID
    var includeHeaders = param.includeHeaders
    var sheetName = param.sheetName
    var startRow = param.startRow
    var startCol = param.startCol
    var key = param.key
    } else {
      return ContentService.createTextOutput("Failed Param")
    }
  if (key != YOURKEY)
  {
    return "Fail.";
  }
  
  try{  
    var result = ImportListToSheet(listID, spreadsheetID, includeHeaders, sheetName, startRow, startCol)
  }
  catch (err){
    return ContentService.createTextOutput("Failed Error: "+err.message)
  }
  return ContentService.createTextOutput(result)
}

function ImportListToSheet(listID, spreadsheetID, includeHeaders, sheetName, startRow, startCol)
{
  if (startRow === undefined) {startRow = 2}
  if (startCol === undefined) {startCol = 1}
  var ss = SpreadsheetApp.openById(spreadsheetID)
  if (sheetName === undefined)
  {
    var sheet = ss.getSheets()[0];
  } else
  {
    var sheet = ss.getSheetByName(sheetName)
  }
  
  var ssAuth = SpreadsheetApp.openById(AUTHSHEETID)
  var sheetAuth = ssAuth.getSheetByName("Main")
  var subKey = sheetAuth.getRange(2, 2).getValue();
  var authToken = sheetAuth.getRange(7, 2).getValue();
  var now = new Date();
  
  statusReset()
  
  if (sheetAuth.getRange(8, 2).getValue() < now)
  {
    // Token expired, refresh
    statusUpdate("Token expired, using refresh token to request new one")
    var url = sheetAuth.getRange(10, 2).getValue() + "?type=refresh&state="+YOURSTATE
    var request = UrlFetchApp.fetch(url)
    statusUpdate(request)
    if (request != "Success")
    {
      statusUpdate("Token request failed: " + request.getContentText())
      return request.getContentText()
    }
    authToken = sheetAuth.getRange(7, 2).getValue();
  }
  
  var headers = {
    'Bb-Api-Subscription-Key': subKey,
    'Authorization': "Bearer " + authToken
  }
  
  var params = {
    'method': 'get',
    'headers': headers,
    'muteHttpExceptions': true
  }
  
  statusUpdate("Sending SKY API request: "+'https://api.sky.blackbaud.com/school/v1/legacy/lists/' + listID)
  var response = UrlFetchApp.fetch('https://api.sky.blackbaud.com/school/v1/legacy/lists/' + listID, params)
  
  if (response.getResponseCode() == 200)
  {
    statusUpdate("SKY API request successful.  Parsing data.")
    try
    {
      var dataSet = JSON.parse(response.getContentText())
      
      if (dataSet.rows.length > 0)
      {
        var dataAll = [];
        var dataRow = [];
        
        if (includeHeaders && startRow > 1)
        {
          for (col = 0; col < dataSet.rows[0].columns.length; col++)
          {
            dataRow.push(dataSet.rows[0].columns[col].name.trim())
          }
          dataAll.push(dataRow)
          startRow = startRow - 1
        }
        
        
        for (var row = 0; row < dataSet.rows.length; row++)
        {
          dataRow = [];
          for (var col = 0; col < dataSet.rows[row].columns.length; col++)
          {
            if (dataSet.rows[row].columns[col].value === undefined)
            {
              dataRow.push("")
            } else
            {
              dataRow.push(dataSet.rows[row].columns[col].value.trim())
            }
          }
          dataAll.push(dataRow)
          if (row > 0 && row % 1000 == 0)
          {
            statusUpdate(row + " of " + dataSet.rows.length + " rows parsed.")
          }
        }
        if (includeHeaders) { row++; }
        statusUpdate("Data parsed.  Importing to sheet "+sheetName+" in "+spreadsheetID+".")
        
        if (sheet.getLastRow() > 0)
        {
          sheet.getRange(startRow, startCol, sheet.getLastRow()+startRow-1, col).clearContent()
        }
        sheet.getRange(startRow, startCol, row, col).setValues(dataAll)
        SpreadsheetApp.flush()
        
        statusUpdate("Data import complete.")
      } else
      {
        statusUpdate("No data to import.")
      }
    } catch (err)
    {
      statusUpdate(err.message)
      return err.message
    }

    
  } else
  {
    statusUpdate("SKY API request failed: " + response.getContentText())
    return response.getContentText()
  }
  
  return "Success"
}

function statusUpdate(statusText)
{
  var ss = SpreadsheetApp.openById(AUTHSHEETID)
  var sheet = ss.getSheetByName("Main")
  var statusCell = sheet.getRange(14, 2)
  var date = new Date()
  statusCell.setValue(statusCell.getValue() + "[" + date.toLocaleString() + "] " + statusText + String.fromCharCode(10))  
  SpreadsheetApp.flush()
}

function statusReset()
{
  var ss = SpreadsheetApp.openById(AUTHSHEETID)
  var sheet = ss.getSheetByName("Main")
  var statusCell = sheet.getRange(14, 2)
  statusCell.setValue("")
}
