// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var server = require('./server');
var router = require('./router');
var authHelper = require('./authHelper');
var outlook = require('node-outlook');

var handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;
handle['/mail'] = mail;
handle['/calendar'] = calendar;
handle['/contacts'] = contacts;

server.start(router.route, handle);

function home(response, request) {
  console.log('Request handler \'home\' was called.');
  response.writeHead(200, {'Content-Type': 'text/html'});
  response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
  response.end();
}

var url = require('url');
function authorize(response, request) {
  console.log('Request handler \'authorize\' was called.');
  
  // The authorization code is passed as a query parameter
  var url_parts = url.parse(request.url, true);
  var code = url_parts.query.code;
  console.log('Code: ' + code);
  authHelper.getTokenFromCode(code, tokenReceived, response);
}

function tokenReceived(response, error, token) {
  if (error) {
    console.log('Access token error: ', error.message);
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p>ERROR: ' + error + '</p>');
    response.end();
  } else {
    getUserEmail(token.token.access_token, function(error, email){
      if (error) {
        console.log('getUserEmail returned an error: ' + error);
        response.write('<p>ERROR: ' + error + '</p>');
        response.end();
      } else if (email) {
        var cookies = ['node-tutorial-token=' + token.token.access_token + ';Max-Age=4000',
                       'node-tutorial-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
                       'node-tutorial-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
                       'node-tutorial-email=' + email + ';Max-Age=4000'];
        response.setHeader('Set-Cookie', cookies);
        response.writeHead(302, {'Location': 'http://localhost:8000/mail'});
        response.end();
      }
    }); 
  }
}

function getUserEmail(token, callback) {
  // Set the API endpoint to use the v2.0 endpoint
  outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

  // Set up oData parameters
  var queryParams = {
    '$select': 'DisplayName, EmailAddress',
  };

  outlook.base.getUser({token: token, odataParams: queryParams}, function(error, user){
    if (error) {
      callback(error, null);
    } else {
      callback(null, user.EmailAddress);
    }
  });
}

function getValueFromCookie(valueName, cookie) {
  if (cookie.indexOf(valueName) !== -1) {
    var start = cookie.indexOf(valueName) + valueName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}

function getAccessToken(request, response, callback) {
  var expiration = new Date(parseFloat(getValueFromCookie('node-tutorial-token-expires', request.headers.cookie)));

  if (expiration <= new Date()) {
    // refresh token
    console.log('TOKEN EXPIRED, REFRESHING');
    var refresh_token = getValueFromCookie('node-tutorial-refresh-token', request.headers.cookie);
    authHelper.refreshAccessToken(refresh_token, function(error, newToken){
      if (error) {
        callback(error, null);
      } else if (newToken) {
        var cookies = ['node-tutorial-token=' + newToken.token.access_token + ';Max-Age=4000',
                       'node-tutorial-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
                       'node-tutorial-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'];
        response.setHeader('Set-Cookie', cookies);
        callback(null, newToken.token.access_token);
      }
    });
  } else {
    // Return cached token
    var access_token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
    callback(null, access_token);
  }
}

function mail(response, request) {
  getAccessToken(request, response, function(error, token) {
    console.log('Token found in cookie: ', token);
    var email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
    console.log('Email found in cookie: ', email);
    if (token) {
      response.writeHead(200, {'Content-Type': 'text/html'});
      response.write('<div><h1>Your inbox</h1></div>');
      
      var queryParams = {
        '$select': 'Subject,ReceivedDateTime,From,IsRead',
        '$orderby': 'ReceivedDateTime desc',
        '$top': 10
      };
      
      // Set the API endpoint to use the v2.0 endpoint
      outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
      // Set the anchor mailbox to the user's SMTP address
      outlook.base.setAnchorMailbox(email);
      
      outlook.mail.getMessages({token: token, folderId: 'inbox', odataParams: queryParams},
        function(error, result){
          if (error) {
            console.log('getMessages returned an error: ' + error);
            response.write('<p>ERROR: ' + error + '</p>');
            response.end();
          }
          else if (result) {
            console.log('getMessages returned ' + result.value.length + ' messages.');
            response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
            result.value.forEach(function(message) {
              console.log('  Subject: ' + message.Subject);
              var from = message.From ? message.From.EmailAddress.Name : 'NONE';
              response.write('<tr><td>' + from + 
                '</td><td>' + (message.IsRead ? '' : '<b>') + message.Subject + (message.IsRead ? '' : '</b>') +
                '</td><td>' + message.ReceivedDateTime.toString() + '</td></tr>');
            });
            
            response.write('</table>');
            response.end();
          }
        });
    } else {
      response.writeHead(200, {'Content-Type': 'text/html'});
      response.write('<p> No token found in cookie!</p>');
      response.end();
    }
  });
}

function buildAttendeeString(attendees) {

  var attendeeString = 'wut';
  if (attendees) {
    attendeeString = '';

    attendees.forEach(function(attendee) {
      attendeeString += '<p>Name:' + attendee.EmailAddress.Name + '</p>';
      attendeeString += '<p>Email:' + attendee.EmailAddress.Address + '</p>';
      attendeeString += '<p>Type:' + attendee.Type + '</p>';
      attendeeString += '<p>Response:' + attendee.Status.Response + '</p>';
      attendeeString += '<p>Respond time:' + attendee.Status.Time + '</p>';
    });
  }

  return attendeeString;
}

function calendar(response, request) {
  var token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
  console.log('Token found in cookie: ', token);
  var email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
  console.log('Email found in cookie: ', email);
  if (token) {
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<div><h1>Your calendar</h1></div>');
    
    var queryParams = {
      '$select': 'Subject,Start,End,Attendees',
      '$orderby': 'Start/DateTime desc',
      '$top': 10
    };
    
    // Set the API endpoint to use the v2.0 endpoint
    outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
    // Set the anchor mailbox to the user's SMTP address
    outlook.base.setAnchorMailbox(email);
    // Set the preferred time zone.
    // The API will return event date/times in this time zone.
    outlook.base.setPreferredTimeZone('Eastern Standard Time');
    
    outlook.calendar.getEvents({token: token, odataParams: queryParams},
      function(error, result){
        if (error) {
          console.log('getEvents returned an error: ' + error);
          response.write('<p>ERROR: ' + error + '</p>');
          response.end();
        } else if (result) {
          console.log('getEvents returned ' + result.value.length + ' events.');
          response.write('<table><tr><th>Subject</th><th>Start</th><th>End</th><th>Attendees</th></tr>');
          result.value.forEach(function(event) {
            console.log('  Subject: ' + event.Subject);
            console.log('  Event dump: ' + JSON.stringify(event));
            response.write('<tr><td>' + event.Subject + 
              '</td><td>' + event.Start.DateTime.toString() +
              '</td><td>' + event.End.DateTime.toString() + 
              '</td><td>' + buildAttendeeString(event.Attendees) + 
              '</td></tr>');
          });
          
          response.write('</table>');
          response.end();
        }
      });
  } else {
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p> No token found in cookie!</p>');
    response.end();
  }
}

function contacts(response, request) {
  var token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
  console.log('Token found in cookie: ', token);
  var email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
  console.log('Email found in cookie: ', email);
  if (token) {
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<div><h1>Your contacts</h1></div>');
    
    var queryParams = {
      '$select': 'GivenName,Surname,EmailAddresses',
      '$orderby': 'GivenName asc',
      '$top': 10
    };
    
    // Set the API endpoint to use the v2.0 endpoint
    outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
    // Set the anchor mailbox to the user's SMTP address
    outlook.base.setAnchorMailbox(email);
    
    outlook.contacts.getContacts({token: token, odataParams: queryParams},
      function(error, result){
        if (error) {
          console.log('getContacts returned an error: ' + error);
          response.write('<p>ERROR: ' + error + '</p>');
          response.end();
        } else if (result) {
          console.log('getContacts returned ' + result.value.length + ' contacts.');
          response.write('<table><tr><th>First name</th><th>Last name</th><th>Email</th></tr>');
          result.value.forEach(function(contact) {
            var email = contact.EmailAddresses[0] ? contact.EmailAddresses[0].Address : 'NONE';
            response.write('<tr><td>' + contact.GivenName + 
              '</td><td>' + contact.Surname +
              '</td><td>' + email + '</td></tr>');
          });
          
          response.write('</table>');
          response.end();
        }
      });
  } else {
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p> No token found in cookie!</p>');
    response.end();
  }
}

/*
  MIT License: 

  Permission is hereby granted, free of charge, to any person obtaining 
  a copy of this software and associated documentation files (the 
  ""Software""), to deal in the Software without restriction, including 
  without limitation the rights to use, copy, modify, merge, publish, 
  distribute, sublicense, and/or sell copies of the Software, and to 
  permit persons to whom the Software is furnished to do so, subject to 
  the following conditions: 

  The above copyright notice and this permission notice shall be 
  included in all copies or substantial portions of the Software. 

  THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
  LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
  OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
  WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/