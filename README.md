# SheetStorm

Read and write to Google Spreadsheets using Node.JS easily with this easy to use module. This is a wrapper around the new V3 Google Sheets API.

SheetStorm is available through an installation from npm
[sheetstorm](https://npmjs.org/package/sheetstorm)

```
npm install sheetstorm
```

## Initialisation

To use this library simply require the module and intialise the connection to Google passing in the required params:
```
var SheetStorm = require('SheetStorm');
var ss = new SheetStorm({
  email: "service_account@email.com",
  pem: "/path/to/pem",
  _debug: true
},
function(err) {
  
  if (err) {
    console.log(err);
  }

  // do something...

})
```
***

## Functionality
Functionality is broke down into the following subheadings:

### Spreadsheets
* [spreadsheetList](#spreadsheet-list) - Return a list of spreasheets

### Worksheets
* [worksheetMeta](#worksheet-meta) - Return a worksheets metadata
* [editWorksheetMeta](#worksheet-meta) - Update a worksheets metadata (title, numRows, numCols)
* [addWorksheet](#worksheet-add) - Add a new worksheet
* [deleteWorksheet](#worksheet-delete) - Delete a worksheet

### Rows
* [rows](#rows) - Return a list of the worksheets rows
* [addRow](#rows-add) - Add a row to the end of the worksheet

### Cells
* [cells](#cells) - Return a list of the worksheets cells
* [editCell](#cells-edit) - Update a specific cell

### Entries
* [entries](#entries) - Return cells as an list of objects with associative headers

***

## Spreadsheets

<a name="spreadsheet-list" />
### Get a list of spreadsheets

___Example___

```
ss.spreadsheetList(function(err, spreadsheets) {

  if (err) {

    console.log(err);

  }

  else {

    console.log(JSON.stringify(spreadsheets, null, 2));

  }
  
})
```

___Output___

```
[
  {
    "id": "8e98da-cabdia372837yu8we8vndsbvi483je",
    "title": "test_sheets",
    "authors": [
      {
        "name": "sheet.storm",
        "email": "sheet.storm@gmail.com"
      }
    ]
  },
  {
    "id": "8e98da-cabdia372837yu8we8vndsbvi483je",
    "title": "Another Test Sheet",
    "authors": [
      {
        "name": "sheet.storm",
        "email": "sheet.storm@gmail.com"
      }
    ]
  }
]
```
***


## Worksheets 

<a name="worksheet-meta" />
### Get a worksheets meta data

___Parameters___
* spreadsheet_id - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method *required*


___Example___
```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je"
}

ss.worksheetMeta(function(err, metadata) {

  if (err) {

    console.log(err);

  }

  else {

    console.log(JSON.stringify(metadata, null, 2));

  }
  
})
```
___Output___
```
{
  "id": "8e98da-cabdia372837yu8we8vndsbvi483je",
  "title": "test_sheets",
  "authors": [
    {
      "name": "sheet.storm",
      "email": "sheet.storm@gmail.com"
    }
  ],
  "sheets": [
    {
      "id": "od6",
      "title": "test"
    },
    {
      "id": "oobhnv",
      "title": "test 2"
    }
  ]
}

```

<a name="worksheets-meta-edit" />
### Edit a worksheets meta data

___Parameters___
* spreadsheet_id **(required)** - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method 
* title - The new title for the worksheet
* cols - The maximum number of columns for the worksheet
* rows - The maximum number of columns for the worksheet

*Please note that if the current number of columns/rows is larger than the new values all data out of range of the new values will be silently deleted.*

___Example___

```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  title: "Some Sheet",
  cols: 20,
  rows: 25
}

ss.editWorksheetMeta(function(err) {

  if (err) {

    console.log(err);

  }
  
})
```

<a name="worksheet-add" />
### Add a new worksheet

___Parameters___
* spreadsheet_id **(required)** - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method 
* title **(required)** - The title for the worksheet
* cols **(optional, defaults to 50)** - The maximum number of columns for the worksheet
* rows **(optional, defaults to 100)** - The maximum number of columns for the worksheet

___Example___

```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  title: "Some Sheet",
  cols: 20,
  rows: 25
}

ss.addWorksheet(function(err) {

  if (err) {

    console.log(err);

  }

  console.log("Worksheet added");
  
})
```

<a name="worksheet-delete" />
### Delete a worksheet

___Parameters___
* spreadsheet_id **(required)** - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method 
* worksheet_id **(required)** - The id of the worksheet, can be found using the [worksheetMeta](#worksheet-meta) method

___Example___

```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  worksheet_id: "od6"
}

ss.deleteWorksheet(function(err) {

  if (err) {

    console.log(err);

  }

  console.log("Worksheet deleted");
  
})
```

***

## Rows

<a name="rows" />
### Get the worksheets cells grouped into rows

___Parameters___
* spreadsheet_id **(required)** - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method 
* worksheet_id **(optional, defaults to od6)** - The id of the worksheet, can be found using the [worksheetMeta](#worksheet-meta) method
* worksheet_name **(optional)** - It's a bit slower using the name because all sheets are returned and the id determined

___Examples___

Using worksheet_id:
```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  worksheet_id: "od6"
}

ss.rows(function(err, rows) {

  if (err) {

    console.log(err);

  }

  console.log(JSON.stringify(rows, null, 2));
  
})
```

Using worksheet_name:
```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  worksheet_name: "test"
}

ss.rows(function(err, rows) {

  if (err) {

    console.log(err);

  }

  console.log(JSON.stringify(rows, null, 2));
  
})
```

___Output___

```
{
  "id": "od6",
  "title": "test",
  "authors": [
    {
      "name": "sheet.storm",
      "email": "sheet.storm@gmail.com"
    }
  ],
  "rows": {
    "r1": [
      {
        "id": "R1C1",
        "title": "A1",
        "value": "header1",
        "row": 1,
        "col": 1
      },
      {
        "id": "R1C2",
        "title": "B1",
        "value": "header2",
        "row": 1,
        "col": 2
      }
    ],
    "r2": [
      {
        "id": "R2C1",
        "title": "A2",
        "value": "value1",
        "row": 2,
        "col": 1
      },
      {
        "id": "R2C2",
        "title": "B2",
        "value": "value2",
        "row": 2,
        "col": 2
      }
    ],
    "r3": [
      {
        "id": "R3C1",
        "title": "A3",
        "value": "value3",
        "row": 3,
        "col": 1
      },
      {
        "id": "R3C2",
        "title": "B3",
        "value": "value4",
        "row": 3,
        "col": 2
      }
    ]
  },
  "total_rows": 3,
  "total_cols": 2
}
```
<a name="rows-add" />
### Append a row to the end of the worksheet

___Parameters___
* spreadsheet_id **(required)** - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method 
* worksheet_id **(required)** - The id of the worksheet, can be found using the [worksheetMeta](#worksheet-meta) method
* row **(required)** - The data to add, an object with the first row values as a key

___Examples___

```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  worksheet_id: "od6"
  row: {
    header1: "value5",
    header2: "value6"
  }
}

ss.addRow(function(err) {

  if (err) {

    console.log(err);

  }

  console.log("Row added");
  
})
```

***

## Cells 

<a name="cells" />
### Get a worksheets cells as a list of cells

___Parameters___
* spreadsheet_id **(required)** - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method 
* worksheet_id **(required)** - The id of the worksheet, can be found using the [worksheetMeta](#worksheet-meta) method


___Example___
```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  worksheet_id: "od6"
}

ss.cells(function(err, cells) {

  if (err) {

    console.log(err);

  }

  else {

    console.log(JSON.stringify(cells, null, 2));

  }
  
})
```
___Output___
```
{
  "id": "od6",
  "title": "test",
  "authors": [
    {
      "name": "sheet.storm",
      "email": "sheet.storm@gmail.com"
    }
  ],
  "cells": [
    {
      "id": "R1C1",
      "title": "A1",
      "value": "header1",
      "row": 1,
      "col": 1
    },
    {
      "id": "R1C2",
      "title": "B1",
      "value": "header2",
      "row": 1,
      "col": 2
    },
    {
      "id": "R2C1",
      "title": "A2",
      "value": "value1",
      "row": 2,
      "col": 1
    },
    {
      "id": "R2C2",
      "title": "B2",
      "value": "value2",
      "row": 2,
      "col": 2
    },
    {
      "id": "R3C1",
      "title": "A3",
      "value": "value3",
      "row": 3,
      "col": 1
    },
    {
      "id": "R3C2",
      "title": "B3",
      "value": "value4",
      "row": 3,
      "col": 2
    },
    {
      "id": "R4C1",
      "title": "A4",
      "value": "value5",
      "row": 4,
      "col": 1
    },
    {
      "id": "R4C2",
      "title": "B4",
      "value": "value6",
      "row": 4,
      "col": 2
    }
  ],
  "total_rows": 4,
  "total_cols": 2
}

```

<a name="cells-edit" />
### Get a worksheets cells as a list of cells

___Parameters___
* spreadsheet_id **(required)** - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method 
* worksheet_id **(required)** - The id of the worksheet, can be found using the [worksheetMeta](#worksheet-meta) method
* row **(required)** - The row to edit
* col **(required)** - The column to edit
* value **(optional)** - The content to add to the cell

*To delete a cells content provide an empty value or no value parameter at all*

___Example___
```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  worksheet_id: "od6",

}

ss.editCell(function(err) {

  if (err) {

    console.log(err);

  }

  else {

    console.log("Cell updated");

  }
  
})
```

***

## Entries

<a name="entries" />
### Get the worksheets cells grouped into a list of associative objects

___Parameters___
* spreadsheet_id **(required)** - The id of the spreadsheet, can be found using the [spreadsheetList](#spreadsheet-list) method 
* worksheet_id **(optional, defaults to od6)** - The id of the worksheet, can be found using the [worksheetMeta](#worksheet-meta) method
* worksheet_name **(optional)** - It's a bit slower using the name because all sheets are returned and the id determined

___Example___

```
var params = {
  spreadsheet_id: "8e98da-cabdia372837yu8we8vndsbvi483je",
  worksheet_id: "od6"
}

ss.entries(function(err, entries) {

  if (err) {

    console.log(err);

  }

  console.log(JSON.stringify(entries, null, 2));
  
})
```

___Output___

```
{
  "id": "od6",
  "title": "test",
  "authors": [
    {
      "name": "sheet.storm",
      "email": "sheet.storm@gmail.com"
    }
  ],
  "entries": [
    {
      "header1": "value1",
      "header2": "value2"
    },
    {
      "header1": "value3",
      "header2": "value4"
    },
    {
      "header1": "value5",
      "header2": "value6"
    }
  ]
}
```

***

## To do list

* Delete a row
* Edit a row
* Add tests

***

## Issues and feature requests

Please feel free to log any issues or feature requests through the [github issues page](https://github.com/interbred-monkey/SheetStorm/issues). 

***

## Third-party libraries

* [async](http://github.com/caolan/async.git) (http://github.com/caolan/async.git)
* [underscore](http://underscorejs.org) (http://underscorejs.org)
* [google-oauth-jwt](https://github.com/extrabacon/google-oauth-jwt) (https://github.com/extrabacon/google-oauth-jwt)
* [xml2json](https://github.com/buglabs/node-xml2json) (https://github.com/buglabs/node-xml2json)