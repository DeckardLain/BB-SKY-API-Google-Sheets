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

function UpdateFromBlackbaud(isScheduled)
{
  var result
  var spreadsheet = SpreadsheetApp.getActive()
  var ssID = spreadsheet.getId()
//  var dashboardSheet = spreadsheet.getSheetByName("Dashboard")
  
  if (isScheduled === undefined)
  {
    isScheduled = false
  }
  
  // Prevent script collision
  var lock = LockService.getScriptLock()
  if (lock.tryLock(1000))
  {
    spreadsheet.toast("Grabbing lists from Blackbaud...", "", 60)
    result = ImportListToSheet(123456, ssID, true, "Data")
    if (result != "Success")
    {
      ShowError("Oh noes! Something went wrong! Send this error info to the database manager: " + result, isScheduled)
      return
    }

//    var date = new Date()
//    dashboardSheet.getRange("B1").setValue(date)
    
    spreadsheet.toast("Update complete!", "", 5)
//    SpreadsheetApp.flush()
    lock.releaseLock()
  } else
  {
    ShowError("Error: Update is already running.", isScheduled)
  }
}

function ImportListToSheet(listID, spreadsheetID, includeHeaders, sheetName, startRow, startCol)
{
  var WEBAPPURL = "on_list_to_sheet_web_app_url";
  var KEY = "your_key";

  var url = WEBAPPURL
  var queryString = "?key="+KEY+"&listID="+listID+"&spreadsheetID="+spreadsheetID
  if (includeHeaders != undefined)
  {
    queryString = queryString + "&includeHeaders=" + includeHeaders
  }
  if (sheetName != undefined)
  {
    queryString = queryString + "&sheetName=" + sheetName
  }
  if (startRow != undefined)
  {
    queryString = queryString + "&startRow=" + startRow
  }
  if (startCol != undefined)
  {
    queryString = queryString + "&startCol=" + startCol
  }
  
  url = url + queryString
  var response = UrlFetchApp.fetch(url)
  return response
}
