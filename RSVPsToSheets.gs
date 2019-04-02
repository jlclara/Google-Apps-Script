/*
This function searches my Gmail inbox for unreads emails from a wedding website (WordPress) I manage that contain 
the keywords "Wedding Guest List" and adds the email message's body to the RSVP Google Sheet. 
The body is split using the delimiter "///". The message is then marked as read.
Note: If you manually read an unread email, please mark the email as unread so that the script can add the new email to the sheet. 


This script was written to replace a depreciated IFTTT "app" that linked my Gmail inbox to Google Sheets.
To use this script, add this Google App Script to the specific sheet you would like to append by going to Tools -> Script Editor
To add triggers for this script, use G Suite's Development Hub's triggers which can be accessed from the Apps Script 
window under the Edit -> Current project's triggers
*/

function sendRSVPsToSheets() {
  var sheet   = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var currentRow = lastRow + 1;
  
  var threads = GmailApp.search('after:2019/01/01 is:unread from:"info@weddingwebsite.com" "Wedding Guest List"');
  for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      for (var m = 0; m < messages.length; m++) {
        // send to sheets
        if (messages[m].isUnread()) {
          sheet.getRange(currentRow,1).setValue(messages[m].getDate());
          sheet.getRange(currentRow,2).setValue(messages[m].getPlainBody());
          sheet.getRange(currentRow,2).splitTextToColumns('///');                   
          currentRow++;
          var msg = messages[m].markRead();
          
        }
      }
  
  }

}
