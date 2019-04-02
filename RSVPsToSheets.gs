/*
This function searches my Gmail inbox for unreads emails from Karliandjoey.com that contain the keywords "Wedding Guest List"
and adds the email message's body to the RSVP spreadsheet. The body is split using the delimiter "///". The message is the marked as read.
*/

function sendRSVPsToSheets() {
  var sheet   = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var currentRow = lastRow + 1;
  
  var threads = GmailApp.search('after:2019/01/01 is:unread from:"info@karliandjoey.com" "Wedding Guest List"');
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
