var BASE_URL = 'https://<<DOMAIN>>.looker.com:19999/api/3.0/';
var CLIENT_ID = '<<CLIENT_ID>>';
var CLIENT_SECRET = '<<>CLIENT_SECRET>';

var TAB_NAME = "<<TabName>>"
function retrieveData(look_id) {
  data = getLooks(look_id)
  
  overwriteTab(TAB_NAME, data, 1) // if looker data has pivot column - skip pivot & header 
  // overwriteTab(TAB_NAME2, data2, 0) // skip header only 
  Logger.log("Finished")
  
  
}


function overwriteTab(tabName, data, starter) {
     Logger.log("Activating sheet")
     spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName) 
     spreadsheet.activate()
     
     spreadsheet.clear()
     
     for (var i = starter; i < data.length; i++) {
         spreadsheet.appendRow(data[i]) 
     }
  
}


function login() {
  try{
    var post = {
        'method': 'post'
    };
    var response = UrlFetchApp.fetch(BASE_URL + "/login?client_id=" + CLIENT_ID + "&client_secret=" + CLIENT_SECRET, post);
    return JSON.parse(response.getContentText()).access_token;
  } catch(err) {
    Logger.log(err);
    return "Could not login to Looker. Check your credentials.";
  }
}

function getLooks(lookId) {
 
  var options = {
    "method": "get",
    "headers": {
      "Authorization": "token " + login() 
    }
  }
    
    var response = UrlFetchApp.fetch(BASE_URL + "looks/" + lookId + "/run/csv?limit=5000&cache=false", options)
    return Utilities.parseCsv(response.getContentText());
  
}