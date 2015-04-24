// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var server = require("./server");
var router = require("./router");
var authHelper = require("./authHelper");
var outlook = require("node-outlook");

var handle = {};
handle["/"] = home;
handle["/authorize"] = authorize;
handle["/mail"] = mail;

server.start(router.route, handle);

function home(response, request) {
  console.log("Request handler 'home' was called.");
  response.writeHead(200, {"Content-Type": "text/html"});
  response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 account.</p>');
  response.end();
}

var url = require("url");
function authorize(response, request) {
  console.log("Request handler 'authorize' was called.");
  
  // The authorization code is passed as a query parameter
  var url_parts = url.parse(request.url, true);
  var code = url_parts.query.code;
  console.log("Code: " + code);
  var token = authHelper.getTokenFromCode(code, 'https://outlook.office365.com/', tokenReceived, response);
}

function tokenReceived(response, error, token) {
  if (error) {
    console.log("Access token error: ", error.message);
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p>ERROR: ' + error + '</p>');
    response.end();
  }
  else {
    response.setHeader('Set-Cookie', ['node-tutorial-token =' + token.token.access_token + ';Max-Age=3600']);
    response.writeHead(302, {'Location': 'http://localhost:8000/mail'});
    response.end();
  }
}

function mail(response, request) {
  var cookieName = 'node-tutorial-token';
  var cookie = request.headers.cookie;
  if (cookie && cookie.indexOf(cookieName) !== -1) {
    console.log("Cookie: ", cookie);
    // Found our token, extract it from the cookie value
    var start = cookie.indexOf(cookieName) + cookieName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    var token = cookie.substring(start, end);
    console.log("Token found in cookie: " + token);
    
    var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office365.com/api/v1.0', 
      authHelper.getAccessTokenFn(token));
    
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<div><span>Your inbox</span></div>');
    response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
    
    outlookClient.me.messages.getMessages()
    .orderBy('DateTimeReceived desc')
    .select('DateTimeReceived,From,Subject').fetchAll(10).then(function (result) {
      result.forEach(function (message) {
        var from = message.from ? message.from.emailAddress.name : "NONE";
        response.write('<tr><td>' + from + 
          '</td><td>' + message.subject +
          '</td><td>' + message.dateTimeReceived.toString() + '</td></tr>');
      });
      
      response.write('</table>');
      response.end();
    },function (error) {
      console.log(error);
      response.write("<p>ERROR: " + error + "</p>");
      response.end();
    });
  }
  else {
    response.writeHead(200, {"Content-Type": "text/html"});
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