// Google Spreadsheets module

var _                 = require('underscore'),
    async             = require('async'),
    request           = require('request'),
    google_auth       = require('./includes/google_auth.js'),
    google_format     = require('./includes/google_format.js'),
    validate_params   = require('./includes/validate.js');

var SheetStorm = function(_params, callback) {
  
  this._debug = _params._debug;
  validate    = new validate_params(_params);

  this.gauth = new google_auth(_params, function(err) {

    if (!_.isNull(err)) {

      return callback(err);

    }

    return callback(null);

  })
  
}

SheetStorm.prototype.spreadsheetList = function(callback) {

  var _instance = this;

  var params = {
    method: "GET",
    url: "https://spreadsheets.google.com/feeds/spreadsheets/private/full"
  }

  _instance.makeRequest(params, function(err, data) {

    if (!_.isNull(err)) {

      return callback("Unable to retrieve a list of spreadsheets");

    }

    // pass this to our formatter
    data = google_format.sheetList(data);

    return callback(null, data)

  })

}

SheetStorm.prototype.worksheetMeta = function(_params, callback) {

  var _instance = this;

  if (!validate.worksheetMeta(_params)) return callback("Incorrect parameters provided, please refer to the documentation");

  var params = {
    method: "GET",
    url: "https://spreadsheets.google.com/feeds/worksheets/" + _params.spreadsheet_id + "/private/full"
  }

  _instance.makeRequest(params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to retrieve the worksheet");

    }

    if (_instance._debug === true) {

      console.log("Recieved worksheet meta data");

    }

    // pass this on to our formatter
    data = google_format.worksheetMeta(data);

    return callback(null, data)

  })

}

SheetStorm.prototype.entries = function(_params, callback) {

  var _instance = this;

  _instance.worksheetData(_params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to retrieve the worksheet cells");

    }

    // pass this on to our formatter
    data = google_format.worksheetEntries(data);

    return callback(null, data);

  })

}

SheetStorm.prototype.cells = function(_params, callback) {

  var _instance = this;

  _instance.worksheetData(_params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to retrieve the worksheet cells");

    }

    // pass this on to our formatter
    data = google_format.worksheetCells(data);

    return callback(null, data);

  })

}

SheetStorm.prototype.rows = function(_params, callback) {

  var _instance = this;

  _instance.worksheetData(_params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to retrieve the worksheet cells");

    }

    // pass this on to our formatter
    data = google_format.worksheetRows(data);

    return callback(null, data);

  })

}

// add a worksheet
SheetStorm.prototype.addWorksheet = function(_params, callback) {

  var _instance = this;

  if (!validate.addWorksheet(_params)) return callback("Incorrect parameters provided, please refer to the documentation");

  if (!_.isNumber(_params.rows)) {

    _params.rows = 100;

  }

  if (!_.isNumber(_params.cols)) {

    _params.cols = 50;

  }

  // make some xml to send
  var xml = google_format.addWorksheetXML(_params);

  // make them into some proper parameters
  var params = {
    method: "POST",
    url: "https://spreadsheets.google.com/feeds/worksheets/" + _params.spreadsheet_id + "/private/full",
    opts: {
      body: xml
    }
  }

  _instance.makeRequest(params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to add a worksheet");

    }

    return callback(null);

  })

}

// delete a worksheet
SheetStorm.prototype.deleteWorksheet = function(_params, callback) {

  var _instance = this;

  if (!validate.deleteWorksheet(_params)) return callback("Incorrect parameters provided, please refer to the documentation");

  // make some xml to send
  var xml = google_format.addWorksheetXML(_params);

  // make them into some proper parameters
  var params = {
    method: "DELETE",
    url: "https://spreadsheets.google.com/feeds/worksheets/" + _params.spreadsheet_id + "/private/full/" + _params.worksheet_id
  }

  _instance.makeRequest(params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to delete worksheet");

    }

    return callback(null);

  })

}

// edit a worksheet meta data
SheetStorm.prototype.editWorksheetMeta = function(_params, callback) {

  var _instance = this;

  if (!validate.editWorksheetMeta(_params)) return callback("Incorrect parameters provided, please refer to the documentation");

  // make some xml to send
  var xml = google_format.editWorksheetMetaXML(_params);

  // make them into some proper parameters
  var params = {
    method: "PUT",
    headers: {
      "If-Match": "*"
    },
    url: "https://spreadsheets.google.com/feeds/worksheets/" + _params.spreadsheet_id + "/private/full/" + _params.worksheet_id,
    opts: {
      body: xml
    }
  }

  _instance.makeRequest(params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to edit worksheet meta data");

    }

    return callback(null);

  })

}

SheetStorm.prototype.addRow = function(_params, callback) {

  var _instance = this;

  if (!validate.addRow(_params)) return callback("Incorrect parameters provided, please refer to the documentation");

  // make some xml for us to send
  var xml = google_format.addRowXML(_params.row);

  // make them into some proper parameters
  var params = {
    method: "POST",
    url: "https://spreadsheets.google.com/feeds/list/" + _params.spreadsheet_id + "/" + _params.worksheet_id + "/private/full",
    opts: {
      body: xml
    }
  }

  _instance.makeRequest(params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to add data to the worksheet");

    }

    return callback(null);

  })

}

SheetStorm.prototype.editCell = function(_params, callback) {

  var _instance = this;

  if (!validate.editCell(_params)) return callback("Incorrect parameters provided, please refer to the documentation");

  // make a cell id
  _params.cell_id = "R" + _params.row + "C" + _params.col;

  // make some xml for us to send
  var xml = google_format.editCellXML(_params);

  // make them into some proper parameters
  var params = {
    method: "PUT",
    headers: {
      "If-Match": "*"
    },
    url: "https://spreadsheets.google.com/feeds/cells/" + _params.spreadsheet_id + "/" + _params.worksheet_id + "/private/full/" + _params.cell_id,
    opts: {
      body: xml
    }
  }

  _instance.makeRequest(params, function(err, data) {

    if (!_.isNull(err)) {

      if (_instance._debug === true) {

        console.log("Error:\n", err, data);

      }

      return callback("Unable to edit the worksheet cell");

    }

    return callback(null);

  })

}

SheetStorm.prototype.worksheetData = function(_params, callback) {

  var _instance = this,
      actions   = [];

  if (!validate.worksheetData(_params)) return callback("Incorrect parameters provided, please refer to the documentation");

  // if we don't have a name or an id or the name is sheet1 use the default
  if ((!_.isString(_params.worksheet_name) && !_.isString(_params.worksheet_id)) || _params.worksheet_name === "Sheet1") {

    _params.worksheet_id = "od6";

  }

  // if we have a worksheet_name then work out the corresponding id
  if ((!_.isString(_params.worksheet_id) || _.isEmpty(_params.worksheet_id)) && _.isString(_params.worksheet_name)) {

    // give it a default
    _params.worksheet_id = null;

    // lookup the id for the worksheet
    actions.push(function(cb) {

      _instance.worksheetMeta(_params, function(err, data) {

        if (!_.isNull(err) || !_.isObject(data)) {

          if (_instance._debug === true) {

            console.log("Error:\n", err, data);

          }

          return cb("Unable to find worksheet");

        }

        // is the data structured properly?
        if (!_.isArray(data.sheets) || _.isEmpty(data.sheets)) {

          if (_instance._debug === true) {

            console.log("Error:\n", "Supplied worksheet name is does not exist in the meta data");

          }

          return cb("Spreadsheet has no worksheets");

        }

        for (var ds in data.sheets) {

          if (_.isObject(data.sheets[ds]) && _.isString(data.sheets[ds].title) && data.sheets[ds].title === _params.worksheet_name) {

            _params.worksheet_id = data.sheets[ds].id;
            break;

          }

        }

        // if we still don't have a worksheet id then error out
        if (_.isNull(_params.worksheet_id)) {

          return cb('Unable to find supplied worksheet_name, make sure it exists in the spreadsheet');

        }

        // gravy
        return cb(null);

      })

    })

  }

  // get the actual worksheet cells
  actions.push(function(cb) {

    // catch all
    if (!_.isString(_params.worksheet_id) || _.isEmpty(_params.worksheet_id)) {

      return cb('Unable to find supplied worksheet_id, make sure it exists in the spreadsheet');

    }

    var params = {
      method: "GET",
      url: "https://spreadsheets.google.com/feeds/cells/" + _params.spreadsheet_id + "/" + _params.worksheet_id + "/private/full"
    }

    _instance.makeRequest(params, function(err, data) {

      if (!_.isNull(err) || !_.isObject(data)) {

        if (_instance._debug === true) {

          console.log("Error:\n", err, data);

        }

        return cb("Unable to retrieve the worksheet data");

      }

      return cb(null, data);

    })

  })
  
  async.series(actions, function(err, data) {

    if (!_.isUndefined(err) || !_.isArray(data)) {

      return callback("Unable to retrieve the worksheet data");

    }

    return callback(null, data[0] || data[1]);

  })

}

SheetStorm.prototype.makeRequest = function(_params, callback) {

  var _instance = this;

  if (validate.makeRequest(_params)) return callback(validate_params.makeRequest(_params));

  if (!_.isObject(_params.opts)) {

    _params.opts = {}

  }

  // define the type we want back
  _params.opts.alt = "json";

  // define we want to use version 3
  _params.opts.v = "3.0";

  // make sure we go to the right url
  _params.followAllRedirects = true;

  // format the opts for the type
  if (_params.method === "GET") {

    // set a return type
    _params.json = true;

    // add in the access token
    _params.opts.access_token = _instance.gauth._access_token;
    _params.qs = _params.opts;

  }

  else {

    // do we have any special headers to add?
    if (!_.isObject(_params.headers)) {

      // add in an XML header
      _params.headers = {
        "Content-Type": "application/atom+xml"
      }

    }

    // we have some custom headers, add in the content type
    else {

      _params.headers["Content-Type"] = "application/atom+xml";

    }

    // add in the access token and teh version etc
    _params.url += "?access_token=" + _instance.gauth._access_token + "&alt=json&v=3.0";
    _params.body = _params.opts.body;

  }

  // remove the opts param
  delete _params.opts;

  request(_params, function(e, r, b) {

    if (e || !_.isObject(r) || (r.statusCode !== 200 && r.statusCode !== 201)) {

      if (_instance._debug === true) {

        console.log("Error: ", e);
        console.log("StatusCode: ", (_.isObject(r) && _.isNumber(r.statusCode)?r.statusCode:null));

      }

      var errorOb = {
        error: e,
        statuscode: (_.isObject(r) && _.isNumber(r.statusCode)?r.statusCode:null),
        body: b
      }

      return callback(errorOb);

    }

    return callback(null, b);

  })

}

module.exports = SheetStorm;