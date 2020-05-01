/*
    Create a google form
    Enable google sheet in repsonse 
    Open script editor form google form & make use of simple steps as below
    my excel sheet will like as:  
    timestamp(1) | first name(2) | last name(3) | email(4)
    
    **Create a connected oauth app in salesforce with callback url as: form url <form url>/exec   [replace viewForm with exec]
    **POST STEPS: script editor
    rename code.gs to forms.gs
    add trigger & bind funtion: 'onFormSubmit' with form trigger events.
    monitor the executions in Script > view > executions 
        
*/

//step 1. Authorize & get token
var AUTHORIZE_URL = 'https://login.salesforce.com/services/oauth2/authorize';
var TOKEN_URL = 'https://login.salesforce.com/services/oauth2/token';

//Put your salesforce connected app information here
var CLIENT_ID = '3MVG9od6vNol.xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
var CLIENT_SECRET = '8xxxxxxxxxxxxxxxx';
var REDIRECT_URI = ScriptApp.getService().getUrl();

//alert(ScriptApp.getService().getUrl());
//https://docs.google.com/macros/s/AKfycbxfweE96BDFUccoIOkfv9jHiMj9VNQqIOYoN_LQCVM23DtFE3w/exec

//this is the user propety where we'll store the token, make sure this is unique across all user properties across all scripts
var tokenPropertyName = 'SALESFORCE_OAUTH_TOKEN';
var baseURLPropertyName = 'SALESFORCE_INSTANCE_URL';


function onFormSubmit(e) {
    Logger.log('calling uploadData ');
    var auth = salesforceAuth(TOKEN_URL, CLIENT_ID, CLIENT_SECRET, 'sf_userName', 'sf_password');
    var accessToken = auth.accessToken;
    var instantUrl = auth.instanceURL;

    var form = FormApp.getActiveForm();//the current form
    var dest_id = form.getDestinationId(); //the destination spreadsheet where form responses are stored
    var ss = SpreadsheetApp.openById(dest_id);//open that spreadsheet
    var theFormSheet = ss.getSheets()[0]; //read the first sheet in that spreadsheet
    var row = theFormSheet.getLastRow(); //get the last row

    //get sheet data
    /*read more: getRange(row, column, numRows, numColumns) | https://developers.google.com/apps-script/reference/spreadsheet/sheet#getrangerow,-column,-numrows,-numcolumns */
    var fname = theFormSheet.getRange(row, 2, 1, 1).getValue();
    var lname = theFormSheet.getRange(row, 3, 1, 1).getValue();
    var email = theFormSheet.getRange(row, 4, 1, 1).getValue();

    try {
        var payload = Utilities.jsonStringify(
            {
                "First_Name__c": fname,
                "Last_Name__c": lname,
                "Email__c": email
            }
        );
        var contentType = "application/json; charset=utf-8";
        var feedUrl = instantUrl + "/services/data/v48.0/sobjects/member__c" + "?_HttpMethod=POST";
        var response = UrlFetchApp.fetch(feedUrl, { method: "POST", headers: { "Authorization": "OAuth " + accessToken }, payload: payload, contentType: contentType });
        //var dataResponse = UrlFetchApp.fetch(feedUrl,getUrlFetchPOSTOptions()).getContentText(); 
        var postResult = Utilities.jsonParse(response.getContentText());
        MailApp.sendEmail("dhananjaykumar3781@gmail.com", "Success", postResult);
    } catch (err) {
        Logger.log("POST ERROR: " + err);
        MailApp.sendEmail("dhananjaykumar3781@gmail.com", "Error", err);
    }

}

function salesforceAuth(authURL, consumerKey, consumerSecret, SalesforceUsername, salesforcePassword) {

        Logger.log("authURL = " + authURL);
        Logger.log("consumerKey = " + consumerKey);

        var payload = {
            'grant_type': 'password',
            'client_id': consumerKey,               // Salesforce Consumer Key
            'client_secret': consumerSecret,
            'username': SalesforceUsername,
            'password': salesforcePassword         // Combination of Password and Security token
        };

        var options = {
            'method': 'post',
            'payload': payload
        };

        var results = UrlFetchApp.fetch(authURL, options);
        var json = results.getContentText();
        var data = JSON.parse(json);

        return { accessToken: data.access_token, instanceURL: data.instance_url }

    }
