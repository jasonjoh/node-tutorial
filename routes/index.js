var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');

/* GET home page. */
router.get('/', function(req, res, next) {
  const signInUrl = authHelper.getAuthUrl();
  res.render('index', { title: 'Home', signInUrl: signInUrl });
});

module.exports = router;
