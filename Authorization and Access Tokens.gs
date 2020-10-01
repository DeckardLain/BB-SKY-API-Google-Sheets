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

var SPREADSHEETID = "your_sky_api_spreadsheet_id";
var YOURSTATE = "your_state_string";

function doGet(e)
{
  var param = e.queryString
  var parameters = param.split("&")
  if (param != null)
  {
    param = e.parameter
    var code = param.code
    var state = param.state
    var error = param.error
    var type = param.type
  } else
  {
      return ContentService.createTextOutput("Failed Param")
  }
  
  if (state != YOURSTATE)
  {
    return ContentService.createTextOutput("Fail")
  }
  
  try
  {
    var ss = SpreadsheetApp.openById(SPREADSHEETID)
    var mainSheet = ss.getSheetByName("Main")
    var redirectURI = mainSheet.getRange(10, 2).getValue()
    
    if (error != null)
    {
      statusUpdate("Authorization error: " + error)
    } else
    {
      statusUpdate("Getting Auth Tokens...")
      if (type == "refresh")
      {
        var formData = {
          'grant_type': 'refresh_token',
          'refresh_token': mainSheet.getRange(9, 2).getValue()
        }
      } else
      {
        var formData = {
          'grant_type': 'authorization_code',
          'code': code,
          'redirect_uri': redirectURI
        }
      }
      var headers = {
        'Authorization': mainSheet.getRange(6, 2).getValue(),
        'Content-Type': 'application/x-www-form-urlencoded'
      }
      var options = {
        'method': 'post',
        'headers': headers,
        'payload': formData
      }
      var response = UrlFetchApp.fetch("https://oauth2.sky.blackbaud.com/token", options)
      
      if (response.getResponseCode() == 200)
      {
        var json = response.getContentText()
        var tokenData = JSON.parse(json)
        
        mainSheet.getRange(7, 2).setValue(tokenData.access_token)
        SpreadsheetApp.flush()
        
        var now = new Date()
        var expireDate = new Date(now.getTime() + tokenData.expires_in * 900)
        mainSheet.getRange(8, 2).setValue(expireDate)
        
        mainSheet.getRange(9, 2).setValue(tokenData.refresh_token)
        statusUpdate("Tokens updated successfully")
        
      } else
      {
        statusUpdate("Token Request Failed")
        statusUpdate(response.getContentText())
      }
    }
  } catch (err)
  {
    return ContentService.createTextOutput("Failed Error: "+err.message)
  }

  return ContentService.createTextOutput("Success")
}

function statusUpdate(statusText)
{
  var ss = SpreadsheetApp.openById(SPREADSHEETID)
  var sheet = ss.getSheetByName("Main")
  var statusCell = sheet.getRange(14, 2)
  var date = new Date()
  statusCell.setValue(statusCell.getValue() + "[" + date.toLocaleString() + "] " + statusText + String.fromCharCode(10))  
  SpreadsheetApp.flush()
}

function statusReset()
{
  var ss = SpreadsheetApp.openById(SPREADSHEETID)
  var sheet = ss.getSheetByName("Main")
  var statusCell = sheet.getRange(14, 2)
  statusCell.setValue("")
}
