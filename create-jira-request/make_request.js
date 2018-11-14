// Distribution list for completion 
var EMAILS = [
  "users@domain"
]

var DOMAIN = "jira-company-domain"

var LOOKER_VERSION = "Bot 2.2"
var SHEET_NAME = "Form Responses"
var USER_CRED = "Basic " + "<<Base64EncodedCreds>>"
var JIRA_URL = "https://" + DOMAIN + ".atlassian.net/rest/api/2/issue/"
var JIRA_CREATOR_NAME = 'requestor-api'

function sendFormByEmailSheet() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_NAME);
  var rangeData = s.getSheetValues(2, 1, getUsedRows(s,1), getUsedCols(s,1)); // All columns with text 
  var dataSet = rangeData[rangeData.length - 1]; // Latest row 

  // Example data columns 
  var request_summary = dataSet[1];
  var request_detail = dataSet[2];
  var team = dataSet[3];
  if (request_summary  != undefined || team != undefined || request_detail != undefined) {
		  createIssue(dataSet)
    }
  else {
     sendEmailsToList("Error on Form Submit", "There are at least 1 undefined fields in the submission, check the sheet \n\n", EMAILS)
  }

}



function createIssue(dataSet){

  var response = testJiraConnection();
  if (response) {
    var responseCode = response.getResponseCode();
    switch(responseCode){
      case 200:
          var data = JSON.parse(response.getContentText());
          Logger.log(data);
          Logger.log("200 response");
          pushIssue(dataSet);
          break;
      default:
          sendEmailsToList("Error Handling Jira Connection Test Request " + responseCode + " Error", "HTTP Response: \n\n" + response, EMAILS);
          break;
      }
  }
  else {
    sendEmailsToList("Error Handling Request", "No Response From API or Invalid Setup", EMAILS);
  }
 }

//######## Format data & make call to JIRA

function pushIssue(dataSet) {
  var date = Date();
  var requestSummary = dataSet[1];
  if (requestSummary.length > 254) { // JIRA has description character limit 
     requestSummary = requestSummary.substring(0,254);
  }
  var requestDetail = dataSet[2];
  var team = dataSet[3].toLowerCase().replace(/ /g, "_"); // Labels must not contain spaces 
  var mainAsk = bold_text("I would like your team to: ") + bullet_point_text(dataSet[4]);
  var emailAddress = dataSet[5]; // email address of person that submits form (could be on behalf of someone)

  var description = new_line() + mainAsk + new_line() + requestDetail + new_line() + "Initially requested by: " + emailAddress

  var data =
{
    "fields": {
        "project": {
            "key": "<<PROJECT>>"
        },
        "summary": requestSummary,
        "issuetype": {
            "name": "Task"
        },
        "reporter": {
            "name": JIRA_CREATOR_NAME
        },
       "priority": {
          "name": "Minor"
       },
        "labels": [
            team,
          "portal"
        ],
        "description": description
    }
};

  var payload = JSON.stringify(data);
  var auth = USER_CRED;
  var headers =
      {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": auth
      };

  var options =
      {
        "Content-Type": "application/json",
        "headers": headers,
        "method": "POST",
        "payload": payload,
        "muteHttpExceptions":true,
        "body":"rest of comment"
      };


  var jiraResponse = UrlFetchApp.fetch(JIRA_URL, options);
  switch (jiraResponse.getResponseCode()) {

  case (200):
       email_users(jiraResponse, requestorEmailAddress);
       Logger.log("Successful Response");
       break;
  case (201):
        email_users(jiraResponse, requestorEmailAddress);
        Logger.log("Successful Response w/New Resource")
        break;
  default:
       handle_errors(jiraResponse, options);
       break;
    }
}




function testJiraConnection() {
  var baseURL = "https://" + DOMAIN + ".atlassian.net/rest/api/2/issue/createmeta?projectKeys=ANALYTICS&issuetypeNames=Task&expand=projects.issuetypes.fields";
  var auth = USER_CRED;

  var args = {
    contentType: "application/json",
    headers: {"Authorization":auth},
    options:{"method":"GET"},
    muteHttpExceptions : true
  };

  var httpResponse = UrlFetchApp.fetch(baseURL, args);
  if (httpResponse) {
    return httpResponse;
  }
  else return null;
}




// Add custom html to email and send email
function sendEmailsToList(emailSubject, emailMessage, emails) {
  var roboLogo = "http://www.animatedimages.org/data/media/118/animated-robot-image-0023.gif";

  var dibsRobo = UrlFetchApp
                    .fetch(roboLogo)
                    .getBlob()
                    .setName('dibsRobo');

  var emailString = ""
  for (var i = 0; i < emails.length; i ++) {
     if (i != emails.length - 1) {
       emailString += emails[i] + ","
     }
     else {
       emailString += emails[i]
     }
  }
  var htmlBodyString = getHTMLBody(emailMessage);
  MailApp.sendEmail({
    to: emailString,
    subject: emailSubject,
    htmlBody: htmlBodyString,
    name:LOOKER_VERSION,
    inlineImages:{
      dibsRobo:dibsRobo
    }
  });
}


function getHTMLBody(emailMessage) {
   var htmlString = "<body style='background-color:white; color: black;'><h2> New Request Completed </h2>";
   htmlString += "<div style='margin: auto;width: 60%;border: 2px solid silver;padding: 10px;background-color: black;color: white;text-align: center;'><h4> JIRA Ticket Created </h2></div>";
   htmlString += "<div style='margin: auto;width: 60%;border: 3px solid silver;padding: 10px;background-color: white;color: black;text-align: center;white-space: pre-line;'><p><b>Note: ";
   htmlString += "</b>" + emailMessage + "</p></div></body>";
  return htmlString
}

function getUsedRows(sheet, columnIndex){
      for (var r = sheet.getLastRow(); r--; r > 0) {
        if (sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getCell(r, columnIndex).getValue() != ""){
          return r;
          break;
        }
      }
    }

function getUsedCols(sheet, rowIndex){
  for (var c = sheet.getLastColumn(); c--; c > 0) {
    if (sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getCell(rowIndex, c).getValue() != ""){
      return c;
      break;
    }
  }
}



function handle_errors(response, options) {
  sendEmailsToList("Error Handling HTTP Request", response.getResponseCode() + '\n\n' + response.getContentText() + '\n' + options["payload"], EMAILS)
  return;
}

function bold_text(text) {
    return "*" + text + "*\n";
}

function bullet_point_text(text) {
   return "\n* " + text;
}

function new_line() {
   return "\n";
}


 function email_users(response, requestorEmailAddress) {

  var dataAll = JSON.parse(response.getContentText());
  var issueKey = dataAll["key"];
  Logger.log(dataAll);

  if (issueKey != null) {
    var emailSubject = "Analytics Request Created";
    var emailBody = 
          "We received your request! \n\n Your reference is " + issueKey + " which can be seen in JIRA with the following link: \n\n" +
        "https://" + DOMAIN + ".atlassian.net/browse/"+ issueKey 
        ;
    var emails = EMAILS.slice();
    emails.push(requestorEmailAddress);
    Logger.log(emails);
    sendEmailsToList(emailSubject, emailBody, emails)
  } else {
    sendEmailsToList("Whoops, Slight Error", dataAll, EMAILS)
  }
 }




/* ############################## ARCHIVED ######################



function removeWatchers(username, issueKey) {
   var issueKey = 'ANALYTICS-1022'
   var baseURL = "https://" + DOMAIN + ".atlassian.net/rest/api/2/issue/" + issueKey + "/watchers?username=" + username;
  var auth = USER_CRED;

  var args = {
    contentType: "application/json",
    headers: {"Authorization":auth},
    options:{"method":"DELETE"},
    muteHttpExceptions : true
  };

  var httpResponse = UrlFetchApp.fetch(baseURL, args);
  Logger.log(httpResponse);

}


Go to this website: https://www.base64encode.net/ and encode in the following manner:    jira_username:jira_password
this will get you your credentials - user must be able to see and update all of the fields they want to use in the app

Use curl to return responses in terminal
*/

