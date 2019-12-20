var express = require('express');
var request = require('node-fetch');
var fulcrumMiddleware = require('connect-fulcrum-webhook');

var PORT = process.env.PORT || 9000;

var app = express();

function payloadProcessor (payload, done) {

  if(payload.data.form_id == 'cdaa6515-0476-4b45-8f9c-a4a93d5c404c'){
    var eventId = payload.data.form_values['6fc3'];
    if(payload.type === 'record.create'){
      createEvent(payload);
    }
    else if(payload.type === 'record.update'){
      updateEvent(eventId, pay);
    }
    else if(payload.type === 'record.delete'){
      deleteEvent(eventId);
    }
  }

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
    method: "POST",
    contentType: "application/x-www-form-urlencoded",
    body: {
      "client_id": CLIENT_ID,
      "scope": API_SCOPE,
      "client_secret": CLIENT_SECRET,
      "grant_type": GRANT_TYPE
    }, 
    json: true
  };
  
  var getData = async function(url, options) {
    try {
      const response = await request(url, options);
      const json = await response.json();
      return json.access_token;
    } catch (error) {
      console.log(error);
    }
  };
  getData(url, options);
  
}

// updates a fulcrum record with the event id.
function updateFulcrumRecord(recordId, eventId){
  var fulcrumAPIkey ='7c9b2ddb2e74c59dee9b357c22651586676eeed86b084021c2cdd4a81ffc21c8bdd8840e969924ae';
  var record = getFulcrumRecord(recordId);
  record.record.form_values["6fc3"] = eventId;
  
  // PUT updated record to Fulcrum
  var url = "https://api.fulcrumapp.com/api/v2/records/" + recordId + ".json?token=" + fulcrumAPIkey;
  var options = {
    method: "PUT",
    contentType: "application/json",
    body: JSON.stringify(record), 
    json: true
  };

  const getData = async function(url, options) {
    try {
      const response = await request(url, options);
      const json = await response.json();
      console.log(json);
    } catch (error) {
      console.log(error);
    }
  };
  getData(url, options);
}

// retrives a fulcrum record. 
function getFulcrumRecord(recordId){
  var fulcrumAPIkey ='7c9b2ddb2e74c59dee9b357c22651586676eeed86b084021c2cdd4a81ffc21c8bdd8840e969924ae';
  var url = "https://api.fulcrumapp.com/api/v2/records/" + recordId + ".json?token=" + fulcrumAPIkey;
  var options = {
    method: "GET",
    contentType: "application/json",
    json: true
  };

  const getData = async function(url) {
    try {
      const response = await request(url);
      const json = await response.json();
      console.log(json);
      return json;
    } catch (error) {
      console.log(error);
      console.log('error');
    }
  };
  getData(url);
  
}

//creates and outlook event.
async function createEvent(payload) {
  var record = await getFulcrumRecord(payload.data.id);
  console.log(record);
  var options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + getToken(),
      'Content-type': 'application/json'
    },
    json: true,
    body : '{"Subject": "' + record.record.form_values['bba9'] 
      + '",  "Body": { "ContentType": "HTML", "Content": "' + record.record.form_values['8841'] 
      + '"  },  "Start": { "DateTime": "' + record.record.form_values['7650'] + 'T' + record.record.form_values['c600'] 
      + '","TimeZone": "Eastern Standard Time" },  "End": {  "DateTime": "' + record.record.form_values['a5f2'] + 'T' + record.record.form_values['c73f'] 
      + '", "TimeZone": "Eastern Standard Time" },  "Attendees": [ {  "EmailAddress": { "Address": "' + record.record.form_values['07f1'] 
      + '", "Name": "Test Here" }, "Type": "Required" }  ]}'
  };
  var url = 'https://graph.microsoft.com/v1.0/users/a0cd0923-d853-4e89-8fc6-d56d7da634d7/events';
  
  const getData = async function(url, options) {
    try {
      const response = await request(url, options);
      const json = await response.json();
      var result = json;
      updateFulcrumRecord(payload.data.id, result.id);
    } catch (error) {
      console.log(error);
    }
  };
  getData(url, options);
  
}

// updates and outlook event
function updateEvent(eventId, payload) {
  var record = getFulcrumRecord(payload.data.id);
  var updateoptions = {
    method: 'PATCH',
    headers: {
      'Authorization': 'Bearer ' + getToken(),
      'Content-type': 'application/json'
    },
    body : '{"Subject": "' + record.record.form_values['bba9'] 
       + '",  "Body": { "ContentType": "HTML", "Content": "' + record.record.form_values['8841'] 
       + '"  },  "Start": { "DateTime": "' + record.record.form_values['7650'] + 'T' + record.record.form_values['c600'] 
       + '","TimeZone": "Eastern Standard Time" },  "End": {  "DateTime": "' + record.record.form_values['a5f2'] + 'T' + record.record.form_values['c73f'] 
       + '", "TimeZone": "Eastern Standard Time" },  "Attendees": [ {  "EmailAddress": { "Address": "' + record.record.form_values['07f1'] 
       + '", "Name": "Test Here" }, "Type": "Required" }  ]}'
  };
  var updateurl = 'https://graph.microsoft.com/v1.0/users/a0cd0923-d853-4e89-8fc6-d56d7da634d7/events/' + eventId;
  
  const getData = async function(updateurl, updateoptions) {
    try {
      const response = await request(updateurl, updateoptions);
      const json = await response.json();
      console.log(json);
    } catch (error) {
      console.log(error);
    }
  };
  getData(updateurl, updateoptions);
  
}

// deletes and outlook event
function deleteEvent(eventId) {
  var deleteoptions = {
    method: 'DELETE',
    headers: {
      'Authorization': 'Bearer ' + getToken(),
      'Content-type': 'application/json',
      'Accept': 'application/json'
    },
    json: true
  };
  var deleteurl = 'https://graph.microsoft.com/v1.0/users/a0cd0923-d853-4e89-8fc6-d56d7da634d7/events/' + eventId;
  
  const getData = async function(deleteurl, deleteoptions){
    try {
      const response = await request(deleteurl, deleteoptions);
      const json = await response.json();
      console.log(json);
    } catch (error) {
      console.log(error);
    }
  };
  getData(deleteurl, deleteoptions);
  
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