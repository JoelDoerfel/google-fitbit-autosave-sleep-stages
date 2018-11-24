// GoogleFitBit-SleepStages
// forged into a sleep-stage API from this code:
// https://github.com/simonbromberg/googlefitbit/blob/master/heartrate.gs

// Handy installation troubleshooting resources:
// https://www.joe0.com/2017/03/18/how-to-sync-fitbit-data-to-google-spreadsheet/
// https://dev.fitbit.com/apps

// Do not change these key names. These are just keys to access these properties once you set them up by running the Setup function from the Fitbit menu

// Key of ScriptProperty for Firtbit consumer key.
var CONSUMER_KEY_PROPERTY_NAME = "fitbitConsumerKey";
// Key of ScriptProperty for Fitbit consumer secret.
var CONSUMER_SECRET_PROPERTY_NAME = "fitbitConsumerSecret";
// Default loggable resources (from Fitbit API docs).
// https://github.com/simonbromberg/googlefitbit/blob/master/interday.gs
/*
var LOGGABLES = ["sleep/startTime",  
                 "sleep/timeInBed",  
                 "sleep/minutesAsleep",  
                 "sleep/awakeningsCount",  
                 "sleep/minutesAwake",  
                 "sleep/minutesToFallAsleep",  
                 "sleep/minutesAfterWakeup",  
                 "sleep/efficiency", "sleep/duration", "sleep/levels"];
*/

var LOGGABLES = ["sleep/stages", "sleep/levels"];

var SERVICE_IDENTIFIER = 'fitbit';
var GET_HEART_RATE_DATA = 1; // too lazy to change variable names. it's actually sleep data

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {
      name: "Setup",
      functionName: "setup"
    },
  {
        name: "Authorize",
        functionName: "showSidebar"
  },
  {
    name: "Reset",
    functionName: "clearService"
  },
  {
        name: "Sync",
    functionName: "refreshTimeSeries"
  },
    {
        name: "Sync & Save",
    functionName: "syncAndSave"
  }];
    ss.addMenu("Fitbit", menuEntries);
}

function isConfigured() {
    return getConsumerKey() != "" && getConsumerSecret() != "";
}

function setConsumerKey(key) {
    ScriptProperties.setProperty(CONSUMER_KEY_PROPERTY_NAME, key);
}

function getConsumerKey() {
    var key = ScriptProperties.getProperty(CONSUMER_KEY_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
}
                 
function setLoggables(loggable) {
    ScriptProperties.setProperty("loggables", loggable);
}

function getLoggables() {
    var loggable = ScriptProperties.getProperty("loggables");
    if (loggable == null) {
        loggable = LOGGABLES;
    } else {
        loggable = loggable.split(',');
    }
    return loggable;
}

function setConsumerSecret(secret) {
    ScriptProperties.setProperty(CONSUMER_SECRET_PROPERTY_NAME, secret);
}

function getConsumerSecret() {
    var secret = ScriptProperties.getProperty(CONSUMER_SECRET_PROPERTY_NAME);
    if (secret == null) {
        secret = "";
    }
    return secret;
}

// function saveSetup saves the setup params from the UI
function saveSetup(e) {
    setConsumerKey(e.parameter.consumerKey);
    setConsumerSecret(e.parameter.consumerSecret);
    setLoggables(e.parameter.loggables);
    setFirstDate(e.parameter.firstDate);
    setLastDate(e.parameter.lastDate);
    var app = UiApp.getActiveApplication();
    app.close();
    return app;
}

function setFirstDate(firstDate) {
    ScriptProperties.setProperty("firstDate", firstDate);
}

function getFirstDate() {
    var firstDate = ScriptProperties.getProperty("firstDate");
    if (firstDate == null) {
      var today = new Date();
      firstDate = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy-MM-dd");
    }
    return firstDate;
}

function setLastDate(lastDate) {
    ScriptProperties.setProperty("lastDate", lastDate);
}

function getLastDate() {
    var lastDate = ScriptProperties.getProperty("lastDate");
    if (lastDate == null) {
      var today = new Date();
      lastDate = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy-MM-dd");
    }
    return lastDate;
}

// function setup accepts and stores the Consumer Key, Consumer Secret, Project Key, firstDate, and list of Data Elements
function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    
    var app = UiApp.createApplication().setTitle("Setup Fitbit Download"); // UiApp API deprecated, may need to eventually replace with HTML service
    app.setStyleAttribute("padding", "10px");
    
    var consumerKeyLabel = app.createLabel("Fitbit OAuth 2.0 Client ID:*");
    var consumerKey = app.createTextBox();
    consumerKey.setName("consumerKey");
    consumerKey.setWidth("100%");
    consumerKey.setText(getConsumerKey());
    
    var consumerSecretLabel = app.createLabel("Fitbit OAuth Consumer Secret:*");
    var consumerSecret = app.createTextBox();
    consumerSecret.setName("consumerSecret");
    consumerSecret.setWidth("100%");
    consumerSecret.setText(getConsumerSecret());
    
    var projectKeyTitleLabel = app.createLabel("Project key: ");
    var projectKeyLabel = app.createLabel(ScriptApp.getProjectKey());
                 
    var firstDate = app.createTextBox().setId("firstDate").setName("firstDate");
    firstDate.setName("firstDate");
    firstDate.setWidth("100%");
    firstDate.setText(getFirstDate());
    
    var lastDate = app.createTextBox().setId("lastDate").setName("lastDate");
    lastDate.setName("lastDate");
    lastDate.setWidth("100%");
    lastDate.setText(getLastDate());
    
    // add listbox to select data elements
    var loggables = app.createListBox(true).setId("loggables").setName(
                                                                       "loggables");
    loggables.setVisibleItemCount(4);
    // add all possible elements (in array LOGGABLES)
    var logIndex = 0;
    for (var resource in LOGGABLES) {
        loggables.addItem(LOGGABLES[resource]);
        // check if this resource is in the getLoggables list
        if (getLoggables().indexOf(LOGGABLES[resource]) > -1) {
            // if so, pre-select it
            loggables.setItemSelected(logIndex, true);
        }
        logIndex++;
    }
    // create the save handler and button
    var saveHandler = app.createServerClickHandler("saveSetup");
    var saveButton = app.createButton("Save Setup", saveHandler);
    
    // put the controls in a grid
    var listPanel = app.createGrid(8, 2);
    listPanel.setWidget(1, 0, consumerKeyLabel);
    listPanel.setWidget(1, 1, consumerKey);
    listPanel.setWidget(2, 0, consumerSecretLabel);
    listPanel.setWidget(2, 1, consumerSecret);
    listPanel.setWidget(3, 0, app.createLabel(" * (obtain these at dev.fitbit.com)"));
    listPanel.setWidget(4, 0, projectKeyTitleLabel);
    listPanel.setWidget(4, 1, projectKeyLabel);
    listPanel.setWidget(5, 0, app.createLabel("Start date (yyyy-mm-dd)"));
    listPanel.setWidget(5, 1, firstDate);
    listPanel.setWidget(6, 0, app.createLabel("End date (yyyy-mm-dd)"));
    listPanel.setWidget(6, 1, lastDate);
    listPanel.setWidget(7, 0, app.createLabel("Data Elements to download:"));
    listPanel.setWidget(7, 1, loggables);
    
    // Ensure that all controls in the grid are handled
    saveHandler.addCallbackElement(listPanel);
    // Build a FlowPanel, adding the grid and the save button
    var dialogPanel = app.createFlowPanel();
    dialogPanel.add(listPanel);
    dialogPanel.add(saveButton);
    app.add(dialogPanel);
    doc.show(app);
}


function getFitbitService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store
  Logger.log(PropertiesService.getUserProperties());
  return OAuth2.createService(SERVICE_IDENTIFIER)

      // Set the endpoint URLs, which are the same for all Google services.
      .setAuthorizationBaseUrl('https://www.fitbit.com/oauth2/authorize')
      .setTokenUrl('https://api.fitbit.com/oauth2/token')

      // Set the client ID and secret, from the Google Developers Console.
      .setClientId(getConsumerKey())
      .setClientSecret(getConsumerSecret())

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      .setScope('activity nutrition sleep weight profile settings')
      
      .setTokenHeaders({
        'Authorization': 'Basic ' + Utilities.base64Encode(getConsumerKey() + ':' + getConsumerSecret())
      });

}

function clearService(){
  OAuth2.createService(SERVICE_IDENTIFIER)
  .setPropertyStore(PropertiesService.getUserProperties())
  .reset();
}

function showSidebar() {
  var service = getFitbitService();
  if (!service.hasAccess()) {
    var authorizationUrl = service.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    Logger.log("Has access!!!!");
  }
}

function authCallback(request) {
  Logger.log("authcallback");
  var service = getFitbitService();
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    Logger.log("success");
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    Logger.log("denied");
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

function getUser() {
  var service = getFitbitService();

  var options = {
          headers: {
      "Authorization": "Bearer " + service.getAccessToken(),
        "method": "GET"
          }};
  var response = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/profile.json",options);
  var o = JSON.parse(response.getContentText());
  return o.user;
}

function getIntradayActivity() {
  if (GET_HEART_RATE_DATA) {
    return ["sleep"];
  }
  else {
    return  ["sleep"];
  }
}

function getFetchString(dateString) {
  if (GET_HEART_RATE_DATA) {
    return "https://api.fitbit.com/1.2/user/-/sleep/date/" + dateString+ "/" + dateString + ".json"
  }
  else {
    return "https://api.fitbit.com/1.2/user/-/sleep/date/" + dateString + ".json"
  }
}

function getActivity() {
  if (GET_HEART_RATE_DATA) {
    return ["sleep"];
  }
  else {
    return  ["sleep"];
  }
}
function refreshTimeSeries() {
  if (!isConfigured()) {
    setup();
    return;
  }
    var user = getUser();
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    doc.getActiveSheet().clear();
    doc.setFrozenRows(2);
    // two header rows
    doc.getRange("a1").setValue(user.fullName);
    // doc.getRange("a1").setComment("DOB:" + user.dateOfBirth)
    // doc.getRange("b1").setValue(user.country + "/" + user.state + "/" + user.city);

    var options =
        {headers:{
        "Authorization": 'Bearer ' + getFitbitService().getAccessToken(),
        "method": "GET"
        }};
    	var activities = getActivity();
    	var intradays = getIntradayActivity();

	var lastIndex = 0;
    for (var activity in activities) {
    	var index = 0;
	 	var dateString = getFirstDate();
	 	date = parseDate(dateString);
        var table = new Array();
	    while (1) {
          var currentActivity = activities[activity];
          try {
            var result = UrlFetchApp.fetch(getFetchString(dateString), options);
          } catch(exception) {
            Logger.log(exception);
          }
          var o = JSON.parse(result.getContentText());
          //  Logger.log(o);

          // SpreadsheetApp.getUi().alert('Full data: '+result.getContentText());

          var cell = doc.getRange('a3');
          var titleCell = doc.getRange("a2");
          titleCell.setValue("Date");
          var title = "Stage"
          // var title = currentActivity.split("/");
          // title = title[title.length - 1];          
          var seconds = "Seconds"
          titleCell.offset(0, 1 + activity * 1.0).setValue(title);
          titleCell.offset(0, 2 + activity * 1.0).setValue(seconds);
          // var row = o[intradays[activity]]["data"];

          var row = o["sleep"][0].levels.data;
          
         // SpreadsheetApp.getUi().alert('Diving deeper into JSON (row): '+JSON.stringify(row));
         // Logger.log("Row = "+JSON.stringify(row));
          
          for (var j in row) {
            var val = row[j];
            var arr = new Array(3);
            try{
            arr[0] = val["dateTime"];
            arr[0] = parseDate2(arr[0]);
            arr[1] = val["level"];
            arr[2] = val["seconds"];
            table.push(arr);
            // set the value index index
            } 
            catch(e) 
            {Logger.log(e)}
            
            index++;
          }

          date.setDate(date.getDate()+1);
          dateString = Utilities.formatDate(date, "GMT", "yyyy-MM-dd");
          if (dateString > getLastDate()) {
            break;
          }
        }
      
      Logger.log("Found " + table.length + " rows");
      if (table.length == 0) {
        SpreadsheetApp.getUi().alert('No data in the time range');
        return;
      }
      // Batch set values of table, much faster than doing each time per loop run, this wouldn't work as is if there were multiple activities being listed
      doc.getRange("A3:C"+(table.length+2)).setValues(table);
	}

// autosaves result  
} // end refreshTimeSeries

function syncAndSave(){ 
refreshTimeSeries();
saveAsSpreadsheet();
}

// Save a Copy      
function saveAsSpreadsheet(){ 

  //  var refresh = refreshTimeSeries();
  // refresh;
  var dateString = getFirstDate();
    
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = SpreadsheetApp.getActiveSheet().getMaxRows();
  var range = sheet.getRange('Sheet1!A1:C'+lastRow);
  sheet.setNamedRange('fitbitDaily', range);
  var TestRange = sheet.getRangeByName('fitbitDaily').getValues(); 
    Logger.log("Saving spreadsheet. Test range found: "+TestRange); 
  var destFolder = DriveApp.getFolderById("YOUR-FOLDER-KEY-HERE"); 
      // var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
      // var now = new Date();
      // var yesterday = new Date(now.getTime() - MILLIS_PER_DAY);
      // var saveDate = Utilities.formatDate(dateString, 'America/Los_Angeles',"yyyyMMdd");  
  
  var saveDate = dateString;  
  DriveApp.getFileById(sheet.getId()).makeCopy("Fitbit Intraday Sleep Stage"+saveDate, destFolder); 

  } // end saveAsSpreadsheet

// parse a date in yyyy-mm-dd format
function parseDate(input) {
  var parts = input.match(/(\d+)/g);
  // new Date(year, month [, date [, hours[, minutes[, seconds[, ms]]]]])
  return new Date(parts[0], parts[1]-1, parts[2]); // months are 0-based
}

// parse a date in 2011-10-25T23:57:00.000 format
function parseDate2(input) {
  var parts = input.match(/(\d+)/g);
  return new Date(parts[0], parts[1]-1, parts[2], parts[3], parts[4]);
}
