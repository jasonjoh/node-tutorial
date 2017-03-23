// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var server = require('./server');
var router = require('./router');
var authHelper = require('./authHelper');
var microsoftGraph = require("@microsoft/microsoft-graph-client");

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
  // Create a Graph client
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  // Get the Graph /Me endpoint to get user email address
  client
    .api('/me')
    .get((err, res) => {
      if (err) {
        callback(err, null);
      } else {
        callback(null, res.mail);
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

      // Create a Graph client
      var client = microsoftGraph.Client.init({
        authProvider: (done) => {
          // Just return the token
          done(null, token);
        }
      });

      // Get the 10 newest messages
      client
        .api('/me/mailfolders/inbox/messages')
        .header('X-AnchorMailbox', email)
        .top(10)
        .select('subject,from,receivedDateTime,isRead')
        .orderby('receivedDateTime DESC')
        .get((err, res) => {
          if (err) {
            console.log('getMessages returned an error: ' + err);
            response.write('<p>ERROR: ' + err + '</p>');
            response.end();
          } else {
            console.log('getMessages returned ' + res.value.length + ' messages.');
            response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
            res.value.forEach(function(message) {
              console.log('  Subject: ' + message.subject);
              var from = message.from ? message.from.emailAddress.name : 'NONE';
              response.write('<tr><td>' + from + 
                '</td><td>' + (message.isRead ? '' : '<b>') + message.subject + (message.isRead ? '' : '</b>') +
                '</td><td>' + message.receivedDateTime.toString() + '</td></tr>');
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

  var attendeeString = '';
  if (attendees) {
    attendees.forEach(function(attendee) {
      attendeeString += '<p>Name:' + attendee.emailAddress.name + '</p>';
      attendeeString += '<p>Email:' + attendee.emailAddress.address + '</p>';
      attendeeString += '<p>Type:' + attendee.type + '</p>';
      attendeeString += '<p>Response:' + attendee.status.response + '</p>';
      attendeeString += '<p>Respond time:' + attendee.status.time + '</p>';
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

    // Create a Graph client
    var client = microsoftGraph.Client.init({
      authProvider: (done) => {
        // Just return the token
        done(null, token);
      }
    });

    // Get the 10 events with the greatest start date
    client
      .api('/me/events')
      .header('X-AnchorMailbox', email)
      .top(10)
      .select('subject,start,end,attendees')
      .orderby('start/dateTime DESC')
      .get((err, res) => {
        if (err) {
          console.log('getEvents returned an error: ' + err);
          response.write('<p>ERROR: ' + err + '</p>');
          response.end();
        } else {
          console.log('getEvents returned ' + res.value.length + ' events.');
          response.write('<table><tr><th>Subject</th><th>Start</th><th>End</th><th>Attendees</th></tr>');
          res.value.forEach(function(event) {
            console.log('  Subject: ' + event.subject);
            response.write('<tr><td>' + event.subject + 
              '</td><td>' + event.start.dateTime.toString() +
              '</td><td>' + event.end.dateTime.toString() +
              '</td><td>' + buildAttendeeString(event.attendees) + '</td></tr>');
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

    // Create a Graph client
    var client = microsoftGraph.Client.init({
      authProvider: (done) => {
        // Just return the token
        done(null, token);
      }
    });

    // Get the first 10 contacts in alphabetical order
    // by given name
    client
      .api('/me/contacts')
      .header('X-AnchorMailbox', email)
      .top(10)
      .select('givenName,surname,emailAddresses')
      .orderby('givenName ASC')
      .get((err, res) => {
        if (err) {
          console.log('getContacts returned an error: ' + err);
          response.write('<p>ERROR: ' + err + '</p>');
          response.end();
        } else {
          console.log('getContacts returned ' + res.value.length + ' contacts.');
          response.write('<table><tr><th>First name</th><th>Last name</th><th>Email</th></tr>');
          res.value.forEach(function(contact) {
            var email = contact.emailAddresses[0] ? contact.emailAddresses[0].address : 'NONE';
            response.write('<tr><td>' + contact.givenName + 
              '</td><td>' + contact.surname +
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