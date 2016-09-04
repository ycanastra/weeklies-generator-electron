const exceljs = require('exceljs')

const numDays = 7;
const maxCol = numDays;

class WeekliesGenerator {
  constructor(labNames, week) {
		this.labNames = labNames;
		this.week = week;
		this.workbook = new exceljs.Workbook();
  }
	generateWeeklies() {
		var workbook = new exceljs.Workbook();

		this.labNames.forEach(labName => this.createWorksheet(labName));
		this.saveWorkbook('test.xlsx');
	}
	createWorksheet(labName) {
		this.workbook.addWorksheet(labName);
	}

	saveWorkbook(filename) {
		this.workbook.xlsx.writeFile(filename)
    .then(function() {
        // done
    });
	}
}

module.exports = WeekliesGenerator
