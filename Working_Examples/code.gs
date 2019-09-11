//See ReadMe file for more information which can be found at 
//https://github.com/kobotoolbox/kobocat-googleapps-scripts

/* Declare global varibles that are used multiple times*/

var current_sheet = SpreadsheetApp.getActive(); //This is the current spreadsheet
var checked_URL = ''; //This is a placeholder for the URL of the kobocat form the user has selected during setup

/***** SECTION 1 - Everything from here to the next section is part of the Setup menu item *****/
function setup() {
  askHost();
}

function askHost() {
  // Clear the all properties first
  ScriptProperties.deleteAllProperties();
  
  // choose host: kc.kobotoolbox.org or kc.humanitarianresponse.info
  var html = HtmlService.createTemplateFromFile("choose-host");
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html.evaluate(), 'Choose Host');
}

function askAPIToken() {
  var host = ScriptProperties.getProperty('host');
  
  // Prompt the user for a row number.
  var selectedRow = Browser.inputBox('You must get your API Token from your user account by visiting http://' + host +'/YOUR-USERNAME/api-token:',
      Browser.Buttons.OK_CANCEL);
  
  if (selectedRow == 'cancel') {
    return;
  }
  
  saveToken(selectedRow);
}

function saveHost(host) {
  ScriptProperties.setProperty('host', host);
  askAPIToken();
}

function saveToken(token) {
  //Save the token and let the user know by message box
  ScriptProperties.setProperty('token', token);
  Browser.msgBox('Token Saved');
  
  //Now, get the form list from kobocat and ask the user which form is meant to update this sheet
  setupScriptProperties();//Get the URLs of each of the users forms
  
}

function setupScriptProperties() {
  var json_array = getkobocatData('https://' + ScriptProperties.getProperty('host') + '/api/v1/forms');
  
  //Get the username from the json_array URL
//  var kobocat_user = JSON.parse(UrlFetchApp.fetch('https://' + ScriptProperties.getProperty('host') + '/api/v1/user', getUrlFetchOptions()).getContentText());
//  var kobocat_username = kobocat_user['username']
  
  //Assign the script properties
  
//  for (var j=0; j<json_array.length; j++)
//  { 
//    ScriptProperties.setProperty('form'+j+'name', json_array[j]["id_string"]);//This value saves the name of the form
//    ScriptProperties.setProperty('form'+j+'URL', json_array[j]["url"]);//This value saves the URL of the form
//    ScriptProperties.setProperty('form'+j+'this_sheet', 'false');//This value will determine if this form is the one that updates the sheet
//    ScriptProperties.setProperty('form'+j+'num_responses', 0)//This tracks the number of form responses per form allowing us to know if there has been an update 
//  }
  
  //Save the number of forms
//  ScriptProperties.setProperty('num_forms', j);
//  ScriptProperties.setProperty('kobocat_username', kobocat_username);
  
  Logger.log(json_array);
  
  askUser(json_array);
}

function askUser(json_array) {
  Logger.log(json_array);
  var html = HtmlService.createTemplateFromFile("choose-form");
  html.json_array = json_array;
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html.evaluate(), 'Choose form to get data');
}

function check(e) {
 //This function takes the user's input from the form and writes preferences to setProperty
  var num_forms = parseInt(ScriptProperties.getProperty('num_forms'));//we use parseInt to make sure an integer is returned, not a string
  
  for(var n=0; n < num_forms;n++){
    var formname= 'form'+n+'name';//Define the formname to pull from getProperty
    var this_sheet_property_name = 'form'+n+'this_sheet';//Set the property name
    //Set the values of the checked box in the ScriptProperties
    ScriptProperties.setProperty(this_sheet_property_name, e.parameter['check'+n]);
  } 
}

function setCheckedUrl(url) {
  ScriptProperties.setProperty('url',url);
  checked_URL = url;
  Browser.msgBox('Ready to Update from Kobocat > Import data, now');
}

function closeApp() {
  Browser.msgBox("Setup Complete");
}

/***** SECTION 2 - Everything from here to the next section is part of the 'Import Data' menu item *****/
function ImportData() {
//This function imports data from kobocat to the existing spreadsheet

  //First check to see if the setup has been run by evaluating the first form's name. If it's null run setup
//  if (ScriptProperties.getProperty('form0name') == null) {
//    setup();
//  } else {  
//    //See if an update is needed
    if (UpdateNeeded()) {
      //if needed, get the data from kobocat
      //Call the function getkobocatData but first, change the URL from /forms/ to /data/
      
      checked_URL = ScriptProperties.getProperty('url');
      
      var form_data_JSON = getkobocatData(checked_URL.replace('/forms/','/data/'));//Note this presently returns all data 
      Logger.log(checked_URL.replace('/forms/','/data/'));
      //Use the setRowsData function (SECTION 5) to write the data to the sheet
      setRowsData(current_sheet.getActiveSheet(), form_data_JSON);
      Logger.log(form_data_JSON);
      Browser.msgBox('Update Complete');
    }else{
      //Tell the user that the form is up to date
      Browser.msgBox('This sheet is up to date')
    }
//  }
}

function UpdateNeeded(){
  return true;
// Check to see which form has a true value for "this_sheet" property
    var num_forms = parseInt(ScriptProperties.getProperty('num_forms'));//we use parseInt to make sure an integer is returned, not a string
    
    for(var n=0; n < num_forms;n++){
      var formname= 'form'+n+'name';//Define the formname to pull from getProperty
      var this_sheet_property_name = 'form'+n+'this_sheet';//Set the property name
      if(ScriptProperties.getProperty(this_sheet_property_name) =='true'){
        var checked_formname = ScriptProperties.getProperty('form'+n+'name');
        checked_URL = ScriptProperties.getProperty('form'+n+'URL');
      }
    }    
        
    //Connect to kobocat and get the form list
    var list_of_forms_JSON = getkobocatData('https://' + ScriptProperties.getProperty('host') + '/api/v1/forms');
    
    //Match the checked_formname to the form in the list_of_forms_JSON and get the number of submissions
    for (var i=0; i < list_of_forms_JSON.length; i++){
      if(list_of_forms_JSON[i]['id_string'] == checked_formname) {
        var latest_num_of_submissions = list_of_forms_JSON[i]['num_of_submissions'];
      }
    }
    
    //Check to see if number of submissions is different than the number of rows in the spreadsheet
    var lastsheetrow = current_sheet.getActiveSheet().getLastRow() - 1; //set the last row of the current sheet (Assuming 1 row header)
    var Update_Needed = true; //assume an update is needed unless the number of rows in the sheet is different from the latest num of submissions 
    
    
    if (latest_num_of_submissions == lastsheetrow) {
      Update_Needed = false;
    }
    
    return Update_Needed;
}

/***** SECTION 3 - Communicating with the kobocat server (used more than once) getkobocatData and Token Authorization Parameters *****/

function getkobocatData(getDataURL) {
  //First, we have to translate the 
  var list_of_forms = UrlFetchApp.fetch(getDataURL, getUrlFetchOptions()).getContentText();  //provides a TEXT list of form names and URLs
  var json_array = JSON.parse(list_of_forms); //parse to JSON so we can work with the data in an array
  //Logger.log(json_array);
  return json_array;
    
}

function getUrlFetchOptions() {
//This function returns the Authorization headers with token information to kobocat
  return {
            "headers" : {
                         "Authorization" : "Token " + ScriptProperties.getProperty('token'),
                        }
         };
}

/***** SECTION 4 - Working with JSON object and entering into spreadsheet *****/
//The following Google Developers site was instrumental in writing this code: https://developers.google.com/apps-script/guides/sheets#reading

// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.
function setRowsData(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  
  objects = objects.map(function (item, index) {
    var keys = Object.keys(objects[index]);
    for(var i = 0; i < keys.length; i++)
    {
      var oldKey = keys[i];
      var key = oldKey.substring(oldKey.indexOf("question_"));
      item[key] = item[oldKey];
  }; 
    return item;
});
  Logger.log(objects);
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 2;
  var headers = headersRange.getValues()[0];
  
  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      try {
        if (header.toString().indexOf('/') >= 0) {
          header_split = header.toString().split('/');
          
          if (('' + objects[i][header_split[0]]).split(' ').indexOf(header_split[1]) >= 0) {
            values.push(1);
          } else {
            values.push(0);
          }
          continue;
        }        
      }
      catch(e) {
        Logger.log(e);
        Logger.log(typeof headers[j]);
      }
      
      // If the header is non-empty and the object value is 0...
      if ((header.length > 0) && (objects[i][header] == 0)) {
        values.push(0);
      }
      // If the header is empty or the object value is empty...
      else if ((!(header.length > 0)) || (objects[i][header]=='')) {
        values.push('');
      }
      else {
        values.push(objects[i][header]);
      }
    }
    data.push(values);
  }

  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(),
                                        objects.length, headers.length);
  destinationRange.setValues(data);
}

/***** SECTION 5 - Setting up custom menu *****/
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Setup",
    functionName : "setup"
  },
  { name : "Import Data",
    functionName : "ImportData"
   }];
  sheet.addMenu("Update from kobocat", entries);
};