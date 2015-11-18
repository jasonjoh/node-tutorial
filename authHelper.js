// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var credentials = {
  clientID: "YOUR APP ID HERE",
  clientSecret: "YOUR APP PASSWORD HERE",
  site: "https://login.microsoftonline.com/common",
  authorizationPath: "/oauth2/v2.0/authorize",
  tokenPath: "/oauth2/v2.0/token"
}
var oauth2 = require("simple-oauth2")(credentials)

var redirectUri = "http://localhost:8000/authorize";

// The scopes the app requires
var scopes = [ "openid",
               "https://outlook.office.com/mail.read",
               "https://outlook.office.com/calendars.read",
               "https://outlook.office.com/contacts.read" ];

function getAuthUrl() {
  var returnVal = oauth2.authCode.authorizeURL({
    redirect_uri: redirectUri,
    scope: scopes.join(" ")
  });
  console.log("Generated auth url: " + returnVal);
  return returnVal;
}

function getTokenFromCode(auth_code, callback, response) {
  var token;
  oauth2.authCode.getToken({
    code: auth_code,
    redirect_uri: redirectUri,
    scope: scopes.join(" ")
    }, function (error, result) {
      if (error) {
        console.log("Access token error: ", error.message);
        callback(response, error, null);
      }
      else {
        token = oauth2.accessToken.create(result);
        console.log("Token created: ", token.token);
        callback(response, null, token);
      }
    });
}

function getEmailFromIdToken(id_token) {
  // JWT is in three parts, separated by a '.'
  var token_parts = id_token.split('.');
  
  // Token content is in the second part, in urlsafe base64
  var encoded_token = new Buffer(token_parts[1].replace("-", "+").replace("_", "/"), 'base64');
  
  var decoded_token = encoded_token.toString();
  
  var jwt = JSON.parse(decoded_token);
  
  // Email is in the preferred_username field
  return jwt.preferred_username
}

exports.getAuthUrl = getAuthUrl;
exports.getEmailFromIdToken = getEmailFromIdToken;
exports.getTokenFromCode = getTokenFromCode; 

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