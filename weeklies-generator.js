const exceljs = require('exceljs')

const numDays = 7;
const maxCol = numDays;

const styles = {
	title: {
		font : {
			name: 'Verdanana',
			size: 24
		},
		alignment: {
			vertical: 'middle',
			horizontal: 'center'
		}
	}
}

class WeekliesGenerator {
  constructor(labNames, week) {
		this.labNames = labNames;
		this.week = week;
		this.workbook = new exceljs.Workbook();
  }
	
	generateWeeklies() {
		var workbook = new exceljs.Workbook();

		this.labNames.forEach((labName) => {
			this.createWorksheet(labName);
		});

		this.workbook.eachSheet((ws, sheetId) => {
			this.createTitle(ws, this.labNames[sheetId - 1]);
		})

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
	createTitle(ws, titleName) {
		ws.mergeCells(1, 1, 1, 7);
		ws.getCell(1, 1, 1, 1).value = titleName;
		ws.getCell(1, 1, 1, 1).font = styles.title.font;
		ws.getCell(1, 1, 1, 1).alignment = styles.title.alignment;
	}
}

module.exports = WeekliesGenerator
