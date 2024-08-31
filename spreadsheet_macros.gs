/** @OnlyCurrentDoc */


// Timer-based trigger that runs every 15min. Checks for items in queue that have completed and sends a text message notification
function checkCraftsUpdate() {
    var sheet_names = ["Armor", "Engineer", "Weapon", "Drone", "Special"];
  
    for (const category of sheet_names) {
      Logger.log("Running " + category + " crafts update check...");
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(category); // get the Armor sheet
      var range = sheet.getRange(7,2,25,9); // get data from cells B7 to J31
      var values = range.getValues(); // extract cell values
      var statusColNum = 8; // status column is 8th column
      var statusColObject = values.map(d => d[statusColNum]); // extract current status column from data
      var statusCol = statusColObject.toString().split(","); // convert to string
      Logger.log("Status column values: " + statusCol);
  
      if (!PropertiesService.getScriptProperties().getKeys().includes(category+"_status")) { // first time run storing initial values in property store
        Logger.log("Initializing property store for " + category + " craft sheet")
        PropertiesService.getScriptProperties().setProperty(category+"_status", JSON.stringify(statusCol));
      } else { // run update checks
        Logger.log("Running checks...")
        var oldStatus = JSON.parse(PropertiesService.getScriptProperties().getProperty(category+"_status"));
        for (var i = 0; i < statusCol.length; i++) {
          // Logger.log("Row " + (i+1) + " | old: " + oldStatus[i] + " | current: " + statusCol[i]);
          if (oldStatus[i] == "Pending" && (statusCol[i] == "Complete!" || statusCol[i] == "Complete - Pending Payment")) {
  
            if (values[i][3] == true && sheet.getRange("F4").getValue() == 0) { // reset void start time
              var datasheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data - Adjustable");
              sheet.getRange("F2").setValue(datasheet.getRange("I6").getValue());
              sheet.getRange("F3").setValue(5.0);
            }
  
            var message = category + "\n\nCOMPLETE: " + values[i][2] + " for " + values[i][0] + "#" + values[i][1];
  
            if (i == statusCol.length - 1) { // spreadsheet needs to be reset
              message = message + "\n\nNEXT: None --- reset " + category + " sheet";
            } else if (statusCol[i+1] == "") {
              message = message + "\n\nNEXT: None";
            } else {
              if (values[i+1][5] == false) { // in-game queue needs to be updated
                message = message + "\n\nUPDATE IN GAME QUEUE";
              }
              message = message + "\n\nNEXT: " + values[i+1][2] + " for " + values[i+1][0] + "#" + values[i+1][1];
  
              if (values[i+1][3] == true) { // next item in queue is a void item
                message = message + "\n\nItem is VOID";
              }
            }
  
            Logger.log(message);
            sendText(message);
          }
        }
  
        PropertiesService.getScriptProperties().setProperty(category+"_status", JSON.stringify(statusCol));
      }
  
  
      // if last completed item was void, reset the void cooldown timer
      let void_timer = sheet.getRange(4, 6);
      if (!PropertiesService.getScriptProperties().getKeys().includes(category+"_void")) { // first time run storing initial values in property store
        PropertiesService.getScriptProperties().setProperty(category+"_void", void_timer.getValue());
      } else {
        if (void_timer.getValue() == 0 && PropertiesService.getScriptProperties().getProperty(category+"_void") > 0) {
          sendText("Void craft for " + category + " is ready");
        }
        PropertiesService.getScriptProperties().setProperty(category+"_void", void_timer.getValue());
      }
    }
  }
  
  // Handles onEdit events for "Reset Time" checkboxes. Syncs the "start time" in the active sheet to the current global time
  function resetStartTime(e) {
    let sheet_names = ["Armor", "Engineer", "Weapon", "Drone", "Special"];
    let sheet = e.source.getActiveSheet();
  
    if (sheet_names.includes(sheet.getName())) {
      if (e.range.getA1Notation() == "D4") {
        let confirm = Browser.msgBox('Confirm choice','Are you sure you want to reset start time?', Browser.Buttons.OK_CANCEL);
        if (confirm == "ok") {
          let datasheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data - Adjustable");
          sheet.getRange("C2").setValue(datasheet.getRange("I6").getValue());
          sheet.getRange("D4").setValue(false);
        }
      } else if (e.range.getA1Notation() == "G4") {
        let confirm = Browser.msgBox('Confirm choice','Are you sure you want to reset VOID start time?', Browser.Buttons.OK_CANCEL);
        if (confirm == "ok") {
          let datasheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data - Adjustable");
          sheet.getRange("F2").setValue(datasheet.getRange("I6").getValue());
          sheet.getRange("G4").setValue(false);
        }
      } else {
        console.log("Ignore...");
      }
    } else {
      console.log("Ignore...");
    }
  }
  
  // Testing function for verifying that texts come through properly
  function testSendMessage() {
    GmailApp.sendEmail("<PHONE NUMBER>@vtext.com", "test subject", "test body");
  }
  
  // Sends an email message to the target. Intended to be phone number
  function sendText(body) {
    var target = "<PHONE NUMBER>@vtext.com"
    var subject = "TIB2 Crafting"
  
    // Logger.log(body)
    GmailApp.sendEmail(target, subject, body)
  }
  
  // Clear the current variables in the property store to re-initialize the system
  function clearDatabase() {
    PropertiesService.getScriptProperties().deleteAllProperties();
  }
  