/*
Program Details
Main Use Case:
* From code editor, set function to start and run
 >> go to start tab and click on "Initiate" button (should be next to "Help" button in the nav, refresh if not, run again if still not, error otherwise)
 >> Click on Initiate button
 >> click on start function
 >> program will grab explores from the specified model below (MODELNAME) and output unhidden fields, one per tab (no established order)
 >> Once done, program will delete existing triggers and create a new trigger, to re-run in 2 weeks
 >> On the first of every month, report will re-check all explores and add new fields (potentially could be combined with clear all to start over)

 Other Use Cases:
* From code editor, set function to clearAllSheets
>> Click on debug button (bug icon)
>> Program will remove all sheets except for "Start"
>> Start is important to have, to at least have one when initiating program in main use case
* From code editor, set function to login
>> Update credentials as needed
>> Run program (play icon)
>> If needed, use Logger.log() function to output to logs
>> Check logs to see if login was successful (click on "View" >> "Logs")
* GLOBALS *
Base Url : looker api url, update if they ever change the endpoint
Client ID: create from Looker user, API3 Keys
Client Secret: create from Looker user, API3 Keys
Headers : headers within sheets, update this if new fields are added or removed
Properties : google scripts specific object, used to keep track of explore index
Model Name : model to grab explores from, update this if a different model is needed
*/

var BASE_URL = 'https://<domain>.looker.com:19999/api/3.0/';
var CLIENT_ID = '<ID>';
var CLIENT_SECRET = '<SECRET>';
var HEADERS = ["Field Type","Label", "View Label", "Type","Description"]
var PROPERTIES = PropertiesService.getScriptProperties();
var MODELNAME = '<model>';

function start() {
  PROPERTIES.setProperties({"explore_index": 0});
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Initiate", [{name:"Start Function", functionName:"main"}]);
}

function main()  {
 if (number_of_cells() >= 2000000) {
   return
 } else {
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Start").activate();
   if ((new Date().getDay() - 1) === 1) {
     monthlyUpdateCheck()
     Logger.log("Monthly Check")
     return
   }
   //clearAllSheets();
   var results = LOOKER_GET_DATA_DICTIONARY();
   if (results) {
     Logger.log("Starting Excel Manipulation")
     addToExcelSheets(results, false);
     deleteTriggers();
     setUpMinTrigger(30);
   }
   else {
      Logger.log("Exit");
     if (results === false) {
        deleteTriggers();
        setUpMinTrigger(5);
        Logger.log("Hidden Explore")
        return
     } else if (results === undefined) { // error
        return
     } else if (results === null) { // explore process done
       Logger.log("Process Complete")
       PROPERTIES.setProperties({"explore_index": 0});
       deleteTriggers();
       setUpWeeklyTrigger(2);
     }
   }
 }
}

// Delete existing triggers
function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
}


// Trigger by the week
function setUpWeeklyTrigger(weeks) {
   ScriptApp.newTrigger('main')
      .timeBased()
      .everyWeeks(weeks)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(9)
      .create();
}

// Trigger by the minute
function setUpMinTrigger(min) {
 ScriptApp.newTrigger('main')
 .timeBased()
 .everyMinutes(min)
 .create()
}


function addToExcelSheets(results, monthlyCheckOption) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var exploreName = results[results.length - 1].slice()
  results.pop()
  for (var i = 0; i < results.length; i++) {
    var data = String(results[i]).split(",")
    var show = data[5] == "false" ? true : false
    var fieldName = data[1]
    if (show && fieldName) {
      data = data.filter(
        function(ele, index) {
          if (index !== 5) {
                  return ele
               }
            }
      ); // remove 'hidden' column from data

      if (ss.getSheetByName(exploreName)) {
        var currentSheet = ss.getSheetByName(exploreName)
        var allCurrentFields = new Array(currentSheet.getRange(1,3))
            if (allCurrentFields.indexOf(fieldName) == -1) {
              currentSheet.appendRow(data)
            }
            else {
              if (monthlyCheckOption === true) {
                currentSheet.getRowGroup(allCurrentFields.indexOf(fieldName), currentSheet.getRowGroupDepth(allCurrentFields.indexOf(fieldName))) = data
              }
          }
      } else {
        ss.insertSheet(exploreName);
        var currentSheet = ss.getSheetByName(exploreName)
    //    var headerWithView = [viewName, " "]
    //    currentSheet.appendRow(headerWithView)
        currentSheet.appendRow(HEADERS);
        currentSheet.appendRow(data)
      }
    }
 }
}


function LOOKER_GET_DATA_DICTIONARY() {
  try {

    var options = {
        'method': 'get',
        'headers': {
            'Authorization': 'token ' + login()
        }
      };

    // api call to the /lookml_models/{lookml_model_name} endpoint
      var response = UrlFetchApp.fetch(BASE_URL + "/lookml_models/" + MODELNAME, options);
      var explores = JSON.parse(response.getContentText()).explores;
      var result = [];

      result.push(HEADERS);
                 // additional details if needed:
                 //, "SQL", "Source"]);


      var index = PROPERTIES.getProperties()['explore_index'] >= explores.length ? undefined : parseInt(PROPERTIES.getProperties()['explore_index']) ;
      Logger.log("On explore " + index);
      if (index !== undefined) {

      var explore_name = explores[index].label != null ? explores[index].label : explores[index].name.replace(/_/g, " ");
      var explore = explores[index].name
      Logger.log("Explore: " + explore);

      var explore_results = UrlFetchApp.fetch(BASE_URL + "/lookml_models/" + MODELNAME + "/explores/" + explore, options);

      var connection = JSON.parse(explore_results.getContentText()).connection_name;
      var dimensions = JSON.parse(explore_results.getContentText()).fields.dimensions;
      var measures = JSON.parse(explore_results.getContentText()).fields.measures;
         Logger.log("Starting field extraction")
            if (explores[index].hidden == false) {
        // return only unhidden explores
              Logger.log("Explore can be used")
               for (var j = 0; j < dimensions.length; j++) {
                 var dimension_ui_label = (dimensions[j].label != null ? (dimensions[j].label).split(dimensions[j].view_label)[1] : (dimensions[j].name.substring((dimensions[j].name.indexOf(".")+1), dimensions[j].name.length).replace(/_/g, " ")).split(dimensions[j].view_label)[1])
                 Logger.log(dimension_ui_label)
                 if ((result.indexOf(dimension_ui_label) == -1) && (dimensions[j].view != undefined || dimensions[j].view_label != undefined)) {
                   result.push(
                //[dimensions[j].view,
                         ["Dimension",
                 //        (dimensions[j].name.substring((dimensions[j].name.indexOf(".")+1), dimensions[j].name.length)).replace(/_/g, " "),
                         dimension_ui_label,
                         dimensions[j].view_label,
                         (dimensions[j].type != null ? (dimensions[j].type).replace("_", " ") : "String"),
                         (dimensions[j].description != null ? dimensions[j].description : ""),
                          dimensions[j].hidden

                         //, (dimensions[j].sql != null ? dimensions[j].sql : ""),
                         //dimensions[j].source_file
                         ]);
                       }
                     }

               for (var k = 0; k < measures.length; k++) {
            // checks that only the fields from the view matching the name of the sheet are displayed
                  var measures_ui_label = (measures[k].label != null ? measures[k].label : (measures[k].name.substring((measures[k].name.indexOf(".")+1), measures[k].name.length).replace(/_/g, " ")))
                  if ((result.indexOf(measures_ui_label) == -1) && (measures[k].view != undefined || measures[k].view_label != undefined)) {
                   result.push(
                //[measures[k].view,
                         ["Measure",
                  //       (measures[k].name.substring((measures[k].name.indexOf(".")+1), measures[k].name.length).replace(/_/g, " ")),
                         measures_ui_label,
                         measures[k].view_label,
                         (measures[k].type != null ? (measures[k].type).replace("_", " ") : "String"),
                         (measures[k].description != null ? measures[k].description : ""),
                         measures[k].hidden
                         //, (measures[k].sql != null ? measures[k].sql : ""),
                         //measures[k].source_file
                         ]);
                    }
                 } // measures loop complete
               result.push(explore_name)
               PROPERTIES.setProperties({"explore_index": index + 1});
               return result
           }
            else {
             // Hidden Explore
             Logger.log("Hidden explore: " + explores[index].hidden)
             PROPERTIES.setProperties({"explore_index": index + 1});
             return false
          }
      } else {
        // When index is greater than explores length - process complete
        return null
      }
    }
  catch(err) {
    Logger.log(err)
    return undefined
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


// Delete all tabs except for the start tab
function clearAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
     if (allSheets[i].getName() != "Start") {
       ss.deleteSheet(allSheets[i]);
     }
  }

}


/*
Once a month, on the first of every month, check across all fields to verify old fields are up to date
*/
function monthlyUpdateCheck() {
  PROPERTIES.getProperties()["explore_index"]
  deleteTriggers();
  var results = LOOKER_GET_DATA_DICTIONARY();
  if (results) {
    addToExcelSheets(results, true);
    deleteTriggers()
    setUpMinTrigger(15);
  }
  else {
    if (results === null) {
       PROPERTIES.setProperties({"explore_index":0});
       deleteTriggers()
       setUpWeeklyTrigger(2);
    }
    else {
      return
    }
  }


}

// Count the number of cells, there is a maximum of 2 million cells (used or unused) placed by Google, which will return an error/email
function number_of_cells(){
   var formatThousandsNoRounding = function(n, dp){
       var e = '', s = e+n, l = s.length, b = n < 0 ? 1 : 0,
           i = s.lastIndexOf('.'), j = i == -1 ? l : i,
           r = e, d = s.substr(j+1, dp);
       while ( (j-=3) > b ) { r = ',' + s.substr(j, 3) + r; }
       return s.substr(0, j + 3) + r +
           (dp ? '.' + d + ( d.length < dp ?
                   ('00000').substr(0, dp - d.length):e):e);
   };
   var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
   var cells_count = 0;
   for (var i in sheets){
       cells_count += (sheets[i].getMaxColumns() * sheets[i].getMaxRows());
   }
   Logger.log(formatThousandsNoRounding(cells_count))
   return cells_count;
}



/* ARCHIVED
function check_cache() {
   var cache = CacheService.getScriptCache();
   var cached = cache.get("api_results");
   if (cached != null) {
     Logger.log("returned from cache");
     for (elem in cached) {
       results.push(elem)
      // Logger.log(results);
       return results;
     }
  }
  return undefined;
}
//  CacheService.getScriptCache().put("api_results", result, 21600)
*/
