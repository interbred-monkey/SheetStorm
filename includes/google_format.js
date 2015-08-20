var _         = require('underscore'),
    //moment    = require('moment'),
    xml2json  = require('xml2json');

var googleFormat = function() {}

googleFormat.prototype.sheetList = function(data) {

  // if we don't have the correct data then provide nothing
  if (!_.isObject(data) || !_.isObject(data.feed) || !_.isArray(data.feed.entry)) {

    return [];

  }

  var entries           = data.feed.entry,
      formatted_entries = [];

  // make a spreadsheet object per file
  for (var e in entries) {

    var ob = {
      id: (_.isObject(entries[e].id)?entries[e].id['$t'].split('/').slice(-1)[0]:null),
      title: (_.isObject(entries[e].title)?entries[e].title['$t']:''),
      url: getLink(entries[e]),
      authors: getAuthors(entries[e])
    }
    
    // add in the formatted object
    formatted_entries.push(ob);

  }

  return formatted_entries;

}

googleFormat.prototype.worksheetMeta = function(data) {

  // if we don't have the correct data then provide nothing
  if (!_.isObject(data) || !_.isObject(data.feed)) {

    return {};

  }

  var feed = data.feed;

  var ob = {
    id: (_.isObject(feed.id)?feed.id['$t'].split('/').slice(-3)[0]:''),
    title: (_.isObject(feed.title)?feed.title['$t']:''),
    url: getLink(feed),
    authors: getAuthors(feed),
    sheets: []
  }

  // pull out all the sheet titles and ids
  if (_.isArray(feed.entry)) {

    for (var fe in feed.entry) {

      if (!_.isObject(feed.entry[fe].id) || !_.isObject(feed.entry[fe].title)) {

        continue;

      }

      var wob = {
        id: (_.isObject(feed.entry[fe].id)?feed.entry[fe].id['$t'].split('/').slice(-1)[0]:''),
        title: (_.isObject(feed.entry[fe].title)?feed.entry[fe].title['$t']:'')
      }

      ob.sheets.push(wob);

    }

  }

  return ob;

}

googleFormat.prototype.worksheetEntries = function(data) {

  // if we don't have the correct data then provide nothing
  if (!_.isObject(data) || !_.isObject(data.feed)) {

    return {};

  }

  var feed        = data.feed,
      col_headers = [],
      rows        = {};

  var ob = {
    id: (_.isObject(feed.id)?feed.id['$t'].split('/').slice(-3)[0]:''),
    title: (_.isObject(feed.title)?feed.title['$t']:''),
    url: getLink(feed),
    authors: getAuthors(feed),
    entries: []
  }

  // pull out all the sheet titles and ids
  if (_.isArray(feed.entry)) {

    for (var fe in feed.entry) {

      if (!_.isObject(feed.entry[fe].id) || !_.isObject(feed.entry[fe].title) || !_.isObject(feed.entry[fe].gs$cell)) {

        continue;

      }

      var row = (_.isString(feed.entry[fe].gs$cell.row)?parseInt(feed.entry[fe].gs$cell.row):0)
          col = (_.isString(feed.entry[fe].gs$cell.col)?parseInt(feed.entry[fe].gs$cell.col):0),
          val = (_.isObject(feed.entry[fe].content)?feed.entry[fe].content['$t']:'');

      // is this our first run through?
      if (row === 1) {

        col_headers.push(val);
        continue;

      }

      if (!_.isObject(rows["r" + row])) {

        rows["r" + row] = {};

      }

      rows["r" + row][col_headers[col - 1]] = val;

    }

    ob.entries = _.values(rows);

  }

  return ob;

}

googleFormat.prototype.worksheetCells = function(data) {

  // if we don't have the correct data then provide nothing
  if (!_.isObject(data) || !_.isObject(data.feed)) {

    return {};

  }

  var feed        = data.feed,
      total_rows  = 0,
      total_cols  = 0;

  var ob = {
    id: (_.isObject(feed.id)?feed.id['$t'].split('/').slice(-3)[0]:''),
    title: (_.isObject(feed.title)?feed.title['$t']:''),
    url: getLink(feed),
    authors: getAuthors(feed),
    cells: []
  }

  // pull out all the sheet titles and ids
  if (_.isArray(feed.entry)) {

    for (var fe in feed.entry) {

      if (!_.isObject(feed.entry[fe].id) || !_.isObject(feed.entry[fe].title) || !_.isObject(feed.entry[fe].gs$cell)) {

        continue;

      }

      var cob = {
        id: (_.isObject(feed.entry[fe].id)?feed.entry[fe].id['$t'].split('/').slice(-1)[0]:''),
        title: (_.isObject(feed.entry[fe].title)?feed.entry[fe].title['$t']:''),
        value: (_.isObject(feed.entry[fe].content)?feed.entry[fe].content['$t']:''),
        row: (_.isString(feed.entry[fe].gs$cell.row)?parseInt(feed.entry[fe].gs$cell.row):''),
        col: (_.isString(feed.entry[fe].gs$cell.col)?parseInt(feed.entry[fe].gs$cell.col):'')
      }

      ob.cells.push(cob);

      // add in the last row number
      if (parseInt(fe) === feed.entry.length - 1) {

        total_cols = (_.isString(feed.entry[fe].gs$cell.col)?parseInt(feed.entry[fe].gs$cell.col):0);
        total_rows = (_.isString(feed.entry[fe].gs$cell.row)?parseInt(feed.entry[fe].gs$cell.row):0);

      }

    }

  }

  // add in the total rows
  ob.total_rows = total_rows;
  ob.total_cols = total_cols;

  return ob;

}

googleFormat.prototype.worksheetRows = function(data) {

  // if we don't have the correct data then provide nothing
  if (!_.isObject(data) || !_.isObject(data.feed)) {

    return {};

  }

  var feed        = data.feed,
      total_rows  = 0,
      total_cols  = 0;

  var ob = {
    id: (_.isObject(feed.id)?feed.id['$t'].split('/').slice(-3)[0]:''),
    title: (_.isObject(feed.title)?feed.title['$t']:''),
    url: getLink(feed),
    authors: getAuthors(feed),
    rows: {}
  }

  // pull out all the sheet titles and ids
  if (_.isArray(feed.entry)) {

    for (var fe in feed.entry) {

      if (!_.isObject(feed.entry[fe].id) || !_.isObject(feed.entry[fe].title) || !_.isObject(feed.entry[fe].gs$cell)) {

        continue;

      }

      var cob = {
        id: (_.isObject(feed.entry[fe].id)?feed.entry[fe].id['$t'].split('/').slice(-1)[0]:''),
        title: (_.isObject(feed.entry[fe].title)?feed.entry[fe].title['$t']:''),
        value: (_.isObject(feed.entry[fe].content)?feed.entry[fe].content['$t']:''),
        row: (_.isString(feed.entry[fe].gs$cell.row)?parseInt(feed.entry[fe].gs$cell.row):''),
        col: (_.isString(feed.entry[fe].gs$cell.col)?parseInt(feed.entry[fe].gs$cell.col):'')
      }

      if (!_.isArray(ob.rows["r" + cob.row])) {

        ob.rows["r" + cob.row] = [];

      }

      ob.rows["r" + cob.row].push(cob);

      // add in the total rows
      if (parseInt(fe) === feed.entry.length - 1) {

        total_cols = (_.isString(feed.entry[fe].gs$cell.col)?parseInt(feed.entry[fe].gs$cell.col):0);
        total_rows = (_.isString(feed.entry[fe].gs$cell.row)?parseInt(feed.entry[fe].gs$cell.row):0);

      }

    }

  }

  // add in the total rows
  ob.total_rows = total_rows;
  ob.total_cols = total_cols;

  return ob;

}

googleFormat.prototype.addWorksheetXML = function(_params) {

  var e = '<title>' + _params.worksheet_name + '</title>\n' +
          ' <gs:rowCount>' + _params.rows + '</gs:rowCount>' +
          ' <gs:colCount>' + _params.cols + '</gs:colCount>\n';

  return addGSEntryContent(e);

}

googleFormat.prototype.addRowXML = function(_params) {

  var entries       = [],
      item_template = '<gsx:<$key$>><$value$></gsx:<$key$>>';

  // make sure we are supplied with something to add as an entry
  if (!_.isArray(_params) && !_.isObject(_params)) {

    return null;

  }

  // if it's only an object then add it to an array
  if (_.isObject(_params)) {

    _params = [_params];

  }

  // loop through and make some entries
  for (var p in _params) {

    if (!_.isObject(_params[p])) {

      continue;

    }

    // loop through the fields to make an cell entry
    for (var pk in _params[p]) {

      // make the replacements
      var e = item_template.replace(/<\$key\$>/g, pk.toLowerCase())
                           .replace(/<\$value\$>/g, _params[p][pk]);

      // add it to the list of entries
      entries.push(e);

    }
    
  }

  // send back the composed string
  return addGSXEntryContent(entries.join('\n'));

}

googleFormat.prototype.editCellXML = function(_params) {

  var template = "<id>https://spreadsheets.google.com/feeds/worksheets/<%sid%>/private/full/<%wid%>?v=3.0</id>\n";

  if (_.isString(_params.title)) {

    template += "<title type=\"text\"><%title%></title>";

  }

  if (_.isNumber(_params.rows)) {

    template += "<gs:rowCount><%rows%></gs:rowCount>";

  }

  if (_.isNumber(_params.cols)) {

    template += "<gs:colCount><%cols%></gs:colCount>";

  }

  // add in the values
  e = template.replace(/<%sid%>/g, _params.spreadsheet_id)
              .replace(/<%wid%>/g, _params.worksheet_id)
              .replace(/<%cid%>/g, _params.cell_id)
              .replace(/<%row%>/g, _params.row)
              .replace(/<%col%>/g, _params.col)
              .replace(/<%value%>/g, _params.value);
  
  // send back the composed string
  return addGSEntryContent(e);

}

googleFormat.prototype.editWorksheetMetaXML = function(_params) {

  var template = " <id>https://spreadsheets.google.com/feeds/worksheets/<%sid%>/<%wid%></id>\n" + 
                 (_.isString(_params.title)?" <title><%title%></title>\n":"") + 
                 (_.isNumber(_params.rows)?" <gs:rowCount><%rows%></gs:rowCount>\n":"") + 
                 (_.isNumber(_params.cols)?" <gs:colCount><%cols%></gs:colCount>\n":"");

  // add in the values
  e = template.replace(/<%sid%>/g, _params.spreadsheet_id)
              .replace(/<%wid%>/g, _params.worksheet_id)
              .replace(/<%title%>/g, _params.title)
              .replace(/<%rows%>/g, _params.rows)
              .replace(/<%cols%>/g, _params.cols);

  // strip off the last newline
  e = e.substr(0, e.length - 1);
  
  // send back the composed string
  return addGSEntryContent(e);

}

function getLink (data) {

  // find the link that relates to the website
  if (_.isArray(data.link)) {

    for (var dl in data.link) {

      if (!_.isObject(data.link[dl]) || 
          (!_.isString(data.link[dl].type) || data.link[dl].type !== "text/html") || 
          (!_.isString(data.link[dl].rel) || data.link[dl].rel !== "self")) {

        continue;

      }

      return data.link[el].href;

    }

  }

}

function getAuthors (data) {

  var authors = [];

  // make a list of authors
  if (_.isArray(data.author)) {

    for (var da in data.author) {

      var ob = {
        name: data.author[da].name['$t'],
        email: data.author[da].email['$t']
      }

      authors.push(ob);

    }

  }

  return authors;

}

function parseXML(xml) {

  return xml2json.toJson(xml);

}

function addGSXEntryContent(content) {

  var template = '<entry xmlns="http://www.w3.org/2005/Atom"' + 
                 ' xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">\n' + 
                 ' <$content$>\n' +
                 '</entry>\n';

  return template.replace(/<\$content\$>/g, content);

}

function addGSEntryContent(content) {

  var template = '<entry xmlns="http://www.w3.org/2005/Atom"' + 
                 ' xmlns:gs="http://schemas.google.com/spreadsheets/2006">\n' + 
                 ' <$content$>\n' +
                 '</entry>\n';

  return template.replace(/<\$content\$>/g, content);

}

module.exports = new googleFormat();