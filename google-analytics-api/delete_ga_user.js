function deleteUser() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName('Main');
  var userInputs = sheet.getRange("A3:B5").getValues();
  var typeOfId = String(userInputs[2][1]);
  var userId = String(userInputs[2][0]);
  var webPropertyId = String(userInputs[0][0]);
  try {
    var request = {
     "kind": "analytics#userDeletionRequest",
     "id": {
     "type": typeOfId, // CLIENT_ID or USER_ID
     "userId": userId
     },
     "webPropertyId": webPropertyId
     };
    Analytics.UserDeletion.UserDeletionRequest.upsert(request);
    sheet.appendRow(["'" + userId, 'yes']);
  } catch(error) {
    Browser.msgBox(error.message);
  }
}
