const exceljs = require('exceljs');
require('datejs');

const numDays = 7;
const maxRow = 32;
const maxCol = 8;

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
	},
	week: {
		font: {
			name: 'Verdanana',
			size: 10
		},
		alignment: {
			vertical: 'middle',
			horizontal: 'center'
		}
	},
	header: {
		font: {
			name: 'Verdanana',
			size: 12
		},
		alignment: {
			vertical: 'middle',
			horizontal: 'center'
		}
	},
	time: {
		font: {
			name: 'Verdanana',
			size: 10
		},
		alignment: {
			vertical: 'top',
			horizontal: 'right'
		}
	}
}

class WeekliesGenerator {
  constructor(labNames, startDate) {
		this.labNames = labNames;
		this.startDate = startDate;
		this.workbook = new exceljs.Workbook();
  }

	generateWeeklies() {
		var workbook = new exceljs.Workbook();

		this.labNames.forEach((labName) => {
			this.createWorksheet(labName);
		});

		this.workbook.eachSheet((ws, sheetId) => {
			const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday',
										'Friday', 'Saturday'];
			var row = 3;
			var col = 2;
			this.createTitle(ws, this.labNames[sheetId - 1]);
			this.createWeek(ws);

			this.createTimeColumn(ws, 3, 1, maxRow - 3);

			for (var i = 0; i < numDays; i++) {
				this.createLabColumn(ws, days[i], row, col + i)
			}
		})

		this.saveWorkbook('test.xlsx');
		console.log('saved now');
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
		ws.mergeCells(1, 1, 1, 8);
		ws.getCell(1, 1).value = titleName;
		ws.getCell(1, 1).font = styles.title.font;
		ws.getCell(1, 1).alignment = styles.title.alignment;
	}
	createWeek(ws) {
		var startDate = new Date(this.startDate);
		var endDate = new Date(this.startDate);
		endDate.add({ days: 6 });

		var sDateStr = startDate.toString('M/d/yyyy');
		var eDateStr = endDate.toString('M/d/yyyy');

		ws.mergeCells(2, 1, 2, 8);
		ws.getCell(2, 1).value = `${sDateStr} - ${eDateStr}`;
		ws.getCell(2, 1).font = styles.week.font;
		ws.getCell(2, 1).alignment = styles.week.alignment;
	}
	createLabColumn(ws, day, row, col) {
		ws.getCell(row, col).value = day;
		ws.getCell(row, col).font = styles.header.font;
		ws.getCell(row, col).alignment = styles.header.alignment;
	}
	createTimeColumn(ws, sRow, sCol, eRow) {
		var date = Date.parse('08:00 AM');

		ws.getCell(sRow, sCol).value = 'Time';
		ws.getCell(sRow, sCol).font = styles.header.font;
		ws.getCell(sRow, sCol).alignment = styles.header.alignment;

		for (var i = 0; i < eRow; i++) {
			if (date.getMinutes() != 30) {
				ws.getCell(sRow + i + 1, sCol).value = date.toString('h:mm tt');
			}
			ws.getCell(sRow + i + 1, sCol).font = styles.time.font;
			ws.getCell(sRow + i + 1, sCol).alignment = styles.time.alignment;

			date.add({ minutes: 30 });
		}
	}
}

module.exports = WeekliesGenerator
