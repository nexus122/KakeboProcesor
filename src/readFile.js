const XLSX = require('xlsx');

class ReadFile {
  constructor(filename) {
    this.workbook = XLSX.readFile(filename);
    this.sheetNames = this.workbook.SheetNames;
    this.sheet = this.workbook.Sheets[this.sheetNames[0]];
    this.range = {
      s: { c: 2, r: 2 },
      e: { c: 1000, r: 1000 },
    };
    this.data = XLSX.utils.sheet_to_json(this.sheet, { range: this.range });
  }
}

module.exports = ReadFile;