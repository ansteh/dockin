'use strict';
const xlsx = require('xlsx');
const fs = require('fs');
const _ = require('lodash');


const Document = (pathname) => {
  let content = xlsx.readFile(pathname);

  return {
    getSheet: (name) => Sheet(_.get(content, ['Sheets', name]))
  };
};

const Sheet = (content) => {
  let createCell = Create.cell(content);
  let cells = [];

  let getCoordinates = (key) => {
    return {
      column: key.substr(0, 1),
      row: key.substr(1)
    };
  };

  _.forEach(_.keys(content), (key) => {
    let coords = getCoordinates(key);
    if(_.isUndefined(cells[coords.column])){
      cells[coords.column] = {};
    }
    let cell = createCell(coords.column, coords.row);
    _.set(cells, [coords.column, coords.row], cell);
  });

  let getCell = (column, row) => {
    return _.get(cells, [column, row]);
  };

  let getColumnNames = () => {
    return _.filter(_.keys(cells), (name) => {
      return /[A-Z]/.test(name);
    });
  };

  let getRowNames = () => {
    return _.keys(_.get(cells, 'A'));
  };

  let getColumnTitles = () => {
    return _.map(columnNames, (name) => {
      return _.result(getCell(name, '1'), 'value', '');
    });
  };

  let columnNames = getColumnNames();
  let rowNames = getRowNames();
  let columnTitles = getColumnTitles();
  let createRow = Create.row(columnTitles);

  /*_.forEach(columnNames, (name) => {
    console.log(name, _.keys(cells[name]).length);
  });*/

  let getRow = (key) => {
    return _.map(columnNames, (columnName) => {
      return getCell(columnName, key);
    });
  };

  let getRows = () => {
    let rows = [];
    if(columnNames && rowNames){
      _.forEach(rowNames, (rowName) => {
        let row = getRow(rowName);
        rows.push(row);
      });
    }
    return rows;
  };

  let rows = getRows();

  let getValuesOfRows = () => {
    return _.map(rows, (row) => {
      return _.map(row, (cell) => {
        return cell ? cell.value() : '';
      });
    });
  };

  let valuesOfRows = getValuesOfRows();

  let searchRowByColumnTitle = _.curry((search, columTitle) => {
    let columnIndex = _.findIndex(columnTitles, title => title === columTitle);
    //console.log('columnIndex', columnIndex, columTitle, columnTitles);
    let targets = _.map(valuesOfRows, row => row[columnIndex]);
    let index = _.findIndex(targets, search);
    return createRow(_.get(rows, index));
  });

  let findRowByColumnTitleThatIncludesText = _.curry((columTitle, text) => {
    return searchRowByColumnTitle((value) => {
      return _.includes(value, text);
    }, columTitle);
  });

  return {
    getCell: getCell,
    getRows: getRows,
    getValuesOfRows: getValuesOfRows,
    getTitles: () => columnTitles,
    searchRowByColumnTitle: searchRowByColumnTitle,
    findRowByColumnTitleThatIncludesText: findRowByColumnTitleThatIncludesText
  }
};

const Cell = (column, row, content) => {
  let value;

  if(_.isString(content)){
    value = content;
  } else if(_.isObject(content)) {
    value = _.find(content, (selection) => {
      return _.includes(selection, '<t>');
    });
    value = _.replace(value, '<t>', '');
    value = _.replace(value, '</t>', '');
  }

  return {
    value: () => value,
    content: () => content
  };
};

const Row = (titles, cells) => {
  let getCellByTitle = (columTitle) => {
    return _.get(cells, _.findIndex(titles, title => title === columTitle));
  };

  return {
    getCellByTitle: getCellByTitle
  };
};

const Create = {};

Create.cell = _.curry((content, column, row) => {
  let key = `${_.upperCase(column)}${row}`;
  return Cell(column, row, _.get(content, key));
});

Create.row = _.curry((titles, cells) => {
  return Row(titles, cells);
});

module.exports = {
  Create: Create,
  Document: Document,
  Sheet: Sheet,
  Cell: Cell
};