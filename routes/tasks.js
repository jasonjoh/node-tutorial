// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

/* GET /mail */
router.get('/', async function(req, res, next) {
  let parms = { title: 'Tasks', active: { tasks: true } };

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;

    // Initialize Graph client
    const client = graph.Client.init({
        baseUrl:'https://graph.microsoft.com/beta',
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    try {
        console.log(client);
      // Get the 10 newest messages from inbox
      const result = await client
      .api('/me/outlook/tasks')         
      .get();
      console.log(result);

      parms.messages = result.value;
      res.render('mail', parms);
    } catch (err) {
      parms.message = 'Error retrieving messages';
      parms.error = { status: `${err.code}: ${err.message}` };
      parms.debug = JSON.stringify(err.body, null, 2);
      res.render('error', parms);
    }
    
  } else {
    // Redirect to home
    res.redirect('/');
  }
});

module.exports = router;