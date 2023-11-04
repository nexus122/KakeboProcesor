const ReadFile = require("./readFile");
const ExcelWriter = require("./excelWriter");

class Main {
    constructor() {
        this.file = new ReadFile("./assets/data.xls");
        const excelWriter = new ExcelWriter();
        excelWriter.createExcelFile(this.file.data, './assets/result.xlsx');
    }
}

let main = new Main();