var _  = require('underscore'),
    async = require('async'),
    googleOAuth = require('google-oauth-jwt');

function googleAuth (opts, callback) {

  this._auth_email   = opts.email,
  this._auth_pem     = opts.pem,
  this._access_token = null,
  this._debug        = (_.isBoolean(opts._debug)?opts._debug:false);

  this.connectPem(function(err, token) {

    if (!_.isNull(err)) {

      return callback(err);

    }

    this._access_token = token;

    return callback(null);

  }.bind(this))

}

googleAuth.prototype.connectPem = function(callback) {

  // make sure we have a copy of our instance
  var _instance = this;

  if (_instance._debug === true) {

    console.log("Connecting to Google");

  }

  googleOAuth.authenticate({
    email: this._auth_email,
    keyFile: this._auth_pem,
    scopes: ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds']
  }, 
  function (err, token) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log(err);

      }

      return callback("Unable to authenticate with Google");

    }

    if (_instance._debug === true) {

      console.log("Connected to Google successfully");

    }

    return callback(null, token);
  
  })
  
}

module.exports = googleAuth;