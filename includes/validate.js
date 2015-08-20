var _ = require('underscore');

var validate = function(_params) {
  
  this.debug = _params._debug;

}

validate.prototype.worksheetMeta = function(_params) {

  if (!_.isString(_params.spreadsheet_id) || _.isEmpty(_params.spreadsheet_id)) {

    if (this._debug) {

      console.log("spreadsheet_id is required");

    }

    return false;

  }

  return true;

}

validate.prototype.worksheetData = function(_params) {

  if (!_.isString(_params.spreadsheet_id) || _.isEmpty(_params.spreadsheet_id)) {

    if (this._debug) {

      console.log("spreadsheet_id is required");

    }

    return false;

  }

  return true;

}

validate.prototype.addWorksheet = function(_params) {

  if (!_.isString(_params.spreadsheet_id) || _.isEmpty(_params.spreadsheet_id)) {

    if (this._debug) {

      console.log("spreadsheet_id is required");

    }

    return false;

  }

  if (!_.isString(_params.worksheet_name) || _.isEmpty(_params.worksheet_name)) {

    if (this._debug) {

      console.log("worksheet_name is required");

    }

    return false;

  }

  return true;

}

validate.prototype.deleteWorksheet = function(_params) {

  if (!_.isString(_params.spreadsheet_id) || _.isEmpty(_params.spreadsheet_id)) {

    if (this._debug) {

      console.log("spreadsheet_id is required");

    }

    return false;

  }

  if (!_.isString(_params.worksheet_id) || _.isEmpty(_params.worksheet_id)) {

    if (this._debug) {

      console.log("worksheet_id is required");

    }

    return false;

  }

  return true;

}

validate.prototype.editWorksheetMeta = function(_params) {

  if (!_.isString(_params.spreadsheet_id) || _.isEmpty(_params.spreadsheet_id)) {

    if (this._debug) {

      console.log("spreadsheet_id is required");

    }

    return false;

  }

  if (!_.isString(_params.worksheet_id) || _.isEmpty(_params.worksheet_id)) {

    if (this._debug) {

      console.log("worksheet_id is required");

    }

    return false;

  }

  if (!_.isString(_params.title) && _.isNumber(_params.rows) && _.isNumber(_params.cols)) {

    if (this._debug) {

      console.log("title, rows and cols are missing, some data is required");

    }

    return false;

  }

  return true;

}

validate.prototype.addRow = function(_params) {

  if (!_.isString(_params.spreadsheet_id) || _.isEmpty(_params.spreadsheet_id)) {

    if (this._debug) {

      console.log("spreadsheet_id is required");

    }

    return false;

  }

  if (!_.isString(_params.worksheet_id) || _.isEmpty(_params.worksheet_id)) {

    if (this._debug) {

      console.log("worksheet_id is required");

    }

    return false;

  }

  if (!_.isObject(_params.row) || _.isEmpty(_params.row)) {

    if (this._debug) {

      console.log("row is required");

    }

    return false;

  }

  return true;

}

validate.prototype.editCell = function(_params) {

  if (!_.isString(_params.spreadsheet_id) || _.isEmpty(_params.spreadsheet_id)) {

    if (this._debug) {

      console.log("spreadsheet_id is required");

    }

    return false;

  }

  if (!_.isNumber(_params.row) || _params.row === 0) {

    if (this._debug) {

      console.log("row is required and should not be 0");

    }

    return false;

  }

  if (!_.isNumber(_params.col) || _params.col === 0) {

    if (this._debug) {

      console.log("col is required and should not be 0");

    }

    return false;

  }

  if (!_.isString(_params.value)) {

    if (this._debug) {

      console.log("value is required");

    }

    return false;

  }

  return true;

}

validate.prototype.makeRequest = function(_params) {
  
  if (!_.isString(_params.method)) {

    return "method is a required parameter";

  }

  if (!_.isString(_params.url)) {

    return "URL is a required parameter";

  }

  return false;

}

module.exports = validate;