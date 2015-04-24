// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var credentials = {
  clientID: "YOUR CLIENT ID HERE",
  clientSecret: "YOUR CLIENT SECRET HERE",
  site: "https://login.microsoftonline.com/common",
  authorizationPath: "/oauth2/authorize",
  tokenPath: "/oauth2/token"
}
var redirectUri = "http://localhost:8000/authorize";
var oauth2 = require("simple-oauth2")(credentials)

function getAuthUrl() {
  var returnVal = oauth2.authCode.authorizeURL({
    redirect_uri: redirectUri
  });
  console.log("Generated auth url: " + returnVal);
  return returnVal;
}

function getTokenFromCode(auth_code, resource, callback, response) {
  var token;
  oauth2.authCode.getToken({
    code: auth_code,
    redirect_uri: redirectUri,
    resource: resource
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

var outlook = require("node-outlook");
function getAccessToken(token) {
  var deferred = new outlook.Microsoft.Utility.Deferred();
  deferred.resolve(token);
  return deferred;
}

function getAccessTokenFn(token) {
  return function() {
  return getAccessToken(token);
  }
}

exports.getAuthUrl = getAuthUrl;
exports.getTokenFromCode = getTokenFromCode;
exports.getAccessTokenFn = getAccessTokenFn;  

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