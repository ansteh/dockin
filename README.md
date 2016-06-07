[![Build Status](https://api.travis-ci.org/ansteh/dockin.svg?branch=master)](https://travis-ci.org/ansteh/dockin)

intuitive xlsx spreadsheet parser wrapper

```js
'use strict';
const dockin = require('dockin');
const Document = dockin.Document;
const path = require('path');
const _ = require('lodash');

let doc = Document(path.resolve(__dirname, 'Matrix1.xlsm'));
let sheet = doc.getSheet('Controls Matrix');
console.log(sheet.getCell('A', 3).value());
let rows = sheet.getRows();
console.log(rows.length, rows[0].length);
//console.log(sheet.getValuesOfRows());
//console.log(_.keys(sheet));
let strategy = sheet.findRowByColumnTitleThatIncludesText('Control Titel', 'IT Strategie');
let columnTitles = sheet.getTitles();
//console.log(columnTitles);
console.log(strategy.getCellByTitle('Control Titel').value());

_.forEach(columnTitles, (title) => {
  console.log(`${title}:`, strategy.getCellByTitle(title).value());
});
```

## License

MIT © [Andre Stehle](https://github.com/ansteh)
