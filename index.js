// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
const server = require('./server');
const router = require('./router');
const authHelper = require('./authHelper');
const microsoftGraph = require("@microsoft/microsoft-graph-client");
const url = require('url');

const handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;
handle['/mail'] = mail;
handle['/calendar'] = calendar;
handle['/contacts'] = contacts;

server.start(router.route, handle);

function home(response, request) {
  console.log('Request handler \'home\' was called.');
  response.writeHead(200, {'Content-Type': 'text/html'});
  response.write(`<p>Please <a href="${authHelper.getAuthUrl()}">sign in</a> with your Office 365 or Outlook.com account.</p>`);
  response.end();
}

function authorize(response, request) {
  console.log('Request handler \'authorize\' was called.');

  // The authorization code is passed as a query parameter
  const url_parts = url.parse(request.url, true);
  const code = url_parts.query.code;
  console.log(`Code: ${code}`);
  processAuthCode(response, code);
}

async function processAuthCode(response, code) {
  let token,email;

  try {
    token = await authHelper.getTokenFromCode(code);
  } catch(error){
    console.log('Access token error: ', error.message);
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write(`<p>ERROR: ${error}</p>`);
    response.end();
    return;
  }

  try {
    email = await getUserEmail(token.token.access_token);
  } catch(error){
    console.log(`getUserEmail returned an error: ${error}`);
    response.write(`<p>ERROR: ${error}</p>`);
    response.end();
    return;
  }

  const cookies = [`node-tutorial-token=${token.token.access_token};Max-Age=4000`,
                   `node-tutorial-refresh-token=${token.token.refresh_token};Max-Age=4000`,
                   `node-tutorial-token-expires=${token.token.expires_at.getTime()};Max-Age=4000`,
                   `node-tutorial-email=${email ? email : ''}';Max-Age=4000`];
  response.setHeader('Set-Cookie', cookies);
  response.writeHead(302, {'Location': 'http://localhost:8000/mail'});
  response.end();
}

async function getUserEmail(token) {
  // Create a Graph client
  const client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  // Get the Graph /Me endpoint to get user email address
  const res = await client
    .api('/me')
    .get();

  // Office 365 users have a mail attribute
  // Outlook.com users do not, instead they have
  // userPrincipalName
  return res.mail ? res.mail : res.userPrincipalName;
}

function getValueFromCookie(valueName, cookie) {
  if (cookie.includes(valueName)) {
    let start = cookie.indexOf(valueName) + valueName.length + 1;
    let end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}

async function getAccessToken(request, response) {
  const expiration = new Date(parseFloat(getValueFromCookie('node-tutorial-token-expires', request.headers.cookie)));

  if (expiration <= new Date()) {
    // refresh token
    console.log('TOKEN EXPIRED, REFRESHING');
    const refresh_token = getValueFromCookie('node-tutorial-refresh-token', request.headers.cookie);
    const newToken = await authHelper.refreshAccessToken(refresh_token);

    const cookies = [`node-tutorial-token=${token.token.access_token};Max-Age=4000`,
                     `node-tutorial-refresh-token=${token.token.refresh_token};Max-Age=4000`,
                     `node-tutorial-token-expires=${token.token.expires_at.getTime()};Max-Age=4000`];
    response.setHeader('Set-Cookie', cookies);
    return newToken.token.access_token;
  }

  // Return cached token
  return getValueFromCookie('node-tutorial-token', request.headers.cookie);
}

async function mail(response, request) {
  let token;

  try {
    token = await getAccessToken(request, response);
  } catch (error){
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p> No token found in cookie!</p>');
    response.end();
    return;
  }

  console.log('Token found in cookie: ', token);
  const email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
  console.log('Email found in cookie: ', email);

  response.writeHead(200, {'Content-Type': 'text/html'});
  response.write('<div><h1>Your inbox</h1></div>');

  // Create a Graph client
  const client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  try {
    // Get the 10 newest messages
    const res = await client
      .api('/me/mailfolders/inbox/messages')
      .header('X-AnchorMailbox', email)
      .top(10)
      .select('subject,from,receivedDateTime,isRead')
      .orderby('receivedDateTime DESC')
      .get();

    console.log(`getMessages returned ${res.value.length} messages.`);
    response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
    res.value.forEach(message => {
      console.log('  Subject: ' + message.subject);
      const from = message.from ? message.from.emailAddress.name : 'NONE';
      response.write(`<tr><td>${from}` +
        `</td><td>${message.isRead ? '' : '<b>'} ${message.subject} ${message.isRead ? '' : '</b>'}` +
        `</td><td>${message.receivedDateTime.toString()}</td></tr>`);
    });

    response.write('</table>');
  } catch (err) {
    console.log(`getMessages returned an error: ${err}`);
    response.write(`<p>ERROR: ${err}</p>`);
  }

  response.end();
}

function buildAttendeeString(attendees) {
  let attendeeString = '';
  if (attendees) {
    attendees.forEach(attendee => {
      attendeeString += `<p>Name:${attendee.emailAddress.name}</p>`;
      attendeeString += `<p>Email:${attendee.emailAddress.address}</p>`;
      attendeeString += `<p>Type:${attendee.type}</p>`;
      attendeeString += `<p>Response:${attendee.status.response}</p>`;
      attendeeString += `<p>Respond time:${attendee.status.time}</p>`;
    });
  }

  return attendeeString;
}

async function calendar(response, request) {
  const token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
  console.log('Token found in cookie: ', token);
  const email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
  console.log('Email found in cookie: ', email);

  if (token) {
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<div><h1>Your calendar</h1></div>');

    // Create a Graph client
    const client = microsoftGraph.Client.init({
      authProvider: (done) => {
        // Just return the token
        done(null, token);
      }
    });

    try {
      // Get the 10 events with the greatest start date
      const res = await client
        .api('/me/events')
        .header('X-AnchorMailbox', email)
        .top(10)
        .select('subject,start,end,attendees')
        .orderby('start/dateTime DESC')
        .get();

      console.log('getEvents returned ' + res.value.length + ' events.');
      response.write('<table><tr><th>Subject</th><th>Start</th><th>End</th><th>Attendees</th></tr>');
      res.value.forEach(function(event) {
        console.log(`  Subject: ${event.subject}`);
        response.write(`<tr><td>${event.subject}` +
          `</td><td>${event.start.dateTime.toString()}` +
          `</td><td>${event.end.dateTime.toString()}` +
          `</td><td>${buildAttendeeString(event.attendees)}</td></tr>`);
      });

      response.write('</table>');
      response.end();
    } catch(err) {
      console.log(`getEvents returned an error: ${err}`);
      response.write(`<p>ERROR: ${err}</p>`);
    }
  } else {
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p> No token found in cookie!</p>');
  }

  response.end();
}

async function contacts(response, request) {
  const token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
  console.log('Token found in cookie: ', token);
  const email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
  console.log('Email found in cookie: ', email);

  if (token) {
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<div><h1>Your contacts</h1></div>');

    // Create a Graph client
    const client = microsoftGraph.Client.init({
      authProvider: (done) => {
        // Just return the token
        done(null, token);
      }
    });

    try {
      // Get the first 10 contacts in alphabetical order
      // by given name
      const res = await client
          .api('/me/contacts')
          .header('X-AnchorMailbox', email)
          .top(10)
          .select('givenName,surname,emailAddresses')
          .orderby('givenName ASC')
          .get();

        console.log(`getContacts returned ${res.value.length} contacts.`);
        response.write('<table><tr><th>First name</th><th>Last name</th><th>Email</th></tr>');
        res.value.forEach(contact => {
            const email = contact.emailAddresses[0] ? contact.emailAddresses[0].address : 'NONE';
            response.write(`<tr><td>${contact.givenName}` +
                `</td><td>${contact.surname}` +
                `</td><td>${email}</td></tr>`);
        });

        response.write('</table>');
    } catch (err) {
      console.log(`getContacts returned an error: ${err}`);
      response.write(`<p>ERROR: ${err}</p>`);
    }
  } else {
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p> No token found in cookie!</p>');
  }

  response.end();
}