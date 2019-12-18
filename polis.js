var express = require('express');
var request = require('request');
var fulcrumMiddleware = require('connect-fulcrum-webhook');

var PORT = process.env.PORT || 9000;

var app = express();

function payloadProcessor (payload, done) {
  var pay = JSON.parse(payload);
  console.log(pay.data.form_id);
  if(pay.data.form_id == 'cdaa6515-0476-4b45-8f9c-a4a93d5c404c'){
    var eventId = pay.data.form_values['6fc3'];
    if(pay.type === 'record.create'){
      createEvent(pay);
    }
    else if(pay.type === 'record.update'){
      updateEvent(eventId, pay);
    }
    else if(paytype === 'record.delete'){
      deleteEvent(eventId);
    }
  }


  console.log('Payload:');
  console.log(payload);
  done();
}

//global variables.
var CLIENT_ID = 'd7c50842-715a-4878-bc22-85579a90f92b'; // Get from azure active direction admin center
var CLIENT_SECRET = '.@YViNb_uFKGU_IS6C0gX8WWsFyOqiO0'; // Get from azure active direction admin center
var TOKEN_URL = 'https://login.microsoftonline.com/52980fd8-4432-4ca2-8be0-7b5fc957bd83/oauth2/v2.0/token';
var API_SCOPE = 'https://graph.microsoft.com/.default';
var GRANT_TYPE = 'client_credentials';

// function to get new token for server to server apps. returns key.
function getToken(){
  var url = 'https://login.microsoftonline.com/52980fd8-4432-4ca2-8be0-7b5fc957bd83/oauth2/v2.0/token';
  var options = {
    "method": "POST",
    "contentType": "application/x-www-form-urlencoded",
    'muteHttpExceptions': true,
    "payload": {
      "client_id": CLIENT_ID,
      "scope": API_SCOPE,
      "client_secret": CLIENT_SECRET,
      "grant_type": GRANT_TYPE
    }
  }
  
  var response = request.post(url, options);
  return JSON.parse(response).access_token;
}

// function to handle webhooks.
function doPost(e){
  return handleResponse(e);
}

// handles the payload.
function handleResponse(e){
  var jsonString = e.postData.getDataAsString();
  var payload = JSON.parse(jsonString);
  if(payload.data.form_id == 'cdaa6515-0476-4b45-8f9c-a4a93d5c404c'){
    var eventId = payload.data.form_values['6fc3'];
    if(payload.type === 'record.create'){
      createEvent(payload);
    }
    else if(payload.type === 'record.update'){
      updateEvent(eventId, payload);
    }
    else if(payload.type === 'record.delete'){
      deleteEvent(eventId);
    }
  }
}

// updates a fulcrum record with the event id.
function updateFulcrumRecord(recordId, eventId){
  var fulcrumAPIkey ='7c9b2ddb2e74c59dee9b357c22651586676eeed86b084021c2cdd4a81ffc21c8bdd8840e969924ae';
  var record = getFulcrumRecord(recordId);
  record.record.form_values["6fc3"] = eventId;
  
  // PUT updated record to Fulcrum
  var url = "https://api.fulcrumapp.com/api/v2/records/" + recordId + ".json?token=" + fulcrumAPIkey;
  var options = {
    "method": "PUT",
    "contentType": "application/json",
    "payload": JSON.stringify(record)
  };
  var recordJSON = request.put(url, options);
}

// retrives a fulcrum record. 
function getFulcrumRecord(recordId){
  var fulcrumAPIkey ='7c9b2ddb2e74c59dee9b357c22651586676eeed86b084021c2cdd4a81ffc21c8bdd8840e969924ae';
  var url = "https://api.fulcrumapp.com/api/v2/records/" + recordId + ".json?token=" + fulcrumAPIkey;
  var options = {
    "method": "GET",
    "contentType": "application/json"
  };
  var recordJSON = request.get(url, options);
  return JSON.parse(recordJSON);
}

//creates and outlook event.
function createEvent(payload) {
  var record = getFulcrumRecord(payload.data.id);
  var options = {
    'Method': 'post',
    headers: {
      'Authorization': 'Bearer ' + getToken(),
      'Content-type': 'application/json'
    },
    'muteHttpExceptions': true,
    'payload' : '{"Subject": "' + record.record.form_values['bba9'] 
      + '",  "Body": { "ContentType": "HTML", "Content": "' + record.record.form_values['8841'] 
      + '"  },  "Start": { "DateTime": "' + record.record.form_values['7650'] + 'T' + record.record.form_values['c600'] 
      + '","TimeZone": "Eastern Standard Time" },  "End": {  "DateTime": "' + record.record.form_values['a5f2'] + 'T' + record.record.form_values['c73f'] 
      + '", "TimeZone": "Eastern Standard Time" },  "Attendees": [ {  "EmailAddress": { "Address": "' + record.record.form_values['07f1'] 
      + '", "Name": "Test Here" }, "Type": "Required" }  ]}'
  };
  var url = 'https://graph.microsoft.com/v1.0/users/a0cd0923-d853-4e89-8fc6-d56d7da634d7/events';
  var response = request.post(url, options);
  var result = JSON.parse(response.getContentText());
  updateFulcrumRecord(payload.data.id, result['id']);
}

// updates and outlook event
function updateEvent(eventId, payload) {
  var record = getFulcrumRecord(payload.data.id);
  var updateoptions = {
    'method': 'patch',
    headers: {
      'Authorization': 'Bearer ' + getToken(),
      'Content-type': 'application/json'
    },
    'muteHttpExceptions': true,
    'payload' : '{"Subject": "' + record.record.form_values['bba9'] 
       + '",  "Body": { "ContentType": "HTML", "Content": "' + record.record.form_values['8841'] 
       + '"  },  "Start": { "DateTime": "' + record.record.form_values['7650'] + 'T' + record.record.form_values['c600'] 
       + '","TimeZone": "Eastern Standard Time" },  "End": {  "DateTime": "' + record.record.form_values['a5f2'] + 'T' + record.record.form_values['c73f'] 
       + '", "TimeZone": "Eastern Standard Time" },  "Attendees": [ {  "EmailAddress": { "Address": "' + record.record.form_values['07f1'] 
       + '", "Name": "Test Here" }, "Type": "Required" }  ]}'
  };
  var updateurl = 'https://graph.microsoft.com/v1.0/users/a0cd0923-d853-4e89-8fc6-d56d7da634d7/events/' + eventId;
  var response = request.patch(updateurl, updateoptions);
  var result = JSON.parse(response.getContentText()); 
}

// deletes and outlook event
function deleteEvent(eventId) {
  var deleteoptions = {
    'method': 'delete',
    headers: {
      'Authorization': 'Bearer ' + getToken(),
      'Content-type': 'application/json',
      'Accept': 'application/json'
    },
    'muteHttpExceptions': true
  };
  var deleteurl = 'https://graph.microsoft.com/v1.0/users/a0cd0923-d853-4e89-8fc6-d56d7da634d7/events/' + eventId;
  var response = request.delete(deleteurl, deleteoptions);
  var result = JSON.parse(response.getResponseCode());
}

var fulcrumMiddlewareConfig = {
  actions: ['record.create', 'record.update', 'record.delete'],
  processor: payloadProcessor
};

app.use('/', fulcrumMiddleware(fulcrumMiddlewareConfig));

app.get('/', function (req, res) {
  res.send('<html><head><title>Polis.js</title></head><body><h2>polis.js</h2><p>Up and Running!</p></body></html>');
});

app.listen(PORT, function () {
  console.log('Listening on port ' + PORT);
});