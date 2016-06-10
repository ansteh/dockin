[![Build Status](https://api.travis-ci.org/ansteh/dockin.svg?branch=master)](https://travis-ci.org/ansteh/dockin)

intuitive xlsx spreadsheet parser wrapper

```js
'use strict';
'use strict';
const indoc = require('./index.js');
const Document = indoc.Document;
const path = require('path');
const _ = require('lodash');

let doc = Document(path.resolve(__dirname, 'file.xlsm'));
let sheet = doc.getSheet('Controls Matrix');
console.log(sheet.getCell('A', 3).value());
let rows = sheet.getRows();
console.log(rows.length, rows[0].length);
//console.log(sheet.getValuesOfRows());
//console.log(_.keys(sheet));
let strategy = sheet.findRow('Control Titel', 'It Strategie');
let columnTitles = sheet.getTitles();
//console.log(columnTitles);
console.log(strategy.getCellByTitle('Control Titel').value());

_.forEach(columnTitles, (title) => {
  console.log(`${title}:`, strategy.getCellByTitle(title).value());
});

console.log(sheet.findRows('Control Titel', 'SAP').length);

```

## License

MIT © [Andre Stehle](https://github.com/ansteh)
