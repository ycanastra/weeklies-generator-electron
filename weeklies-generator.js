const exceljs = require('exceljs');
var $ = require('jquery');

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
			name: 'Calibri',
			size: 12
		},
		alignment: {
			vertical: 'top',
			horizontal: 'right'
		}
	},
	general: {
		font: {
			name: 'Calibri',
			size: 12
		},
		alignment: {
			vertical: 'top',
			horizontal: 'center',
			wrapText: true
		}
	}
}

class ClassEvent {
	constructor(labName, className, instructorName, sTime, eTime) {
		this.labName = labName;
		this.className = className;
		this.instructorName = instructorName;
		this.sTime = sTime;
		this.eTime = eTime;
	}
}

class WeekliesGenerator {
	constructor(labNames, startDate) {
		this.labNames = labNames;
		this.startDate = startDate;
		this.workbook = new exceljs.Workbook();
		this.dates = [];

		var date = new Date(this.startDate);

		for (let i = 0; i < numDays; i++) {
			this.dates.push(date);
			date = new Date(date);
			date.add({days: 1})
		}
	}
 	callback () { console.log('all done'); }
	generateWeeklies() {
		var workbook = new exceljs.Workbook();
		const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday',
									'Friday', 'Saturday'];
		const classEvents = [];
		let ajaxCallsRemaining = 7;

		this.labNames.forEach((labName) => {
			this.createWorksheet(labName);
		});

		this.workbook.eachSheet((ws, sheetId) => {

			var row = 3;
			var col = 2;

			// sheetId starts at 1
			this.createTitle(ws, this.labNames[sheetId - 1]);
			this.createWeek(ws);

			this.createTimeColumn(ws, 3, 1, maxRow - 3);

			for (var i = 0; i < numDays; i++) {
				this.createLabColumn(ws, days[i], 3, maxRow - 3, i + 2)
			}
		})
		this.dates.forEach((date, index) => {
			const url = `http://labschedule.collaborate.ucsb.edu/?ts=${date.getTime()/1000}`;
			$.ajax({
				url: url,
				dataType: 'text',
				success: function(data) {
					const tempClassEvents = getClassEvents(date, data)
					for (let classEvent of tempClassEvents) {
						classEvents.push(classEvent)
					}
					--ajaxCallsRemaining;
					if (ajaxCallsRemaining <= 0) {
						classEvents.forEach((classEvent) => {
							that.insertClass(classEvent)
						})
					}
				}
			});
		})

		const that = this;
		setTimeout(function(){ that.saveWorkbook('test.xlsx'); }, 500);

		// console.log('saved now');
	}

	createWorksheet(labName) {
		this.workbook.addWorksheet(labName);
	}

	saveWorkbook(filename) {
		this.workbook.xlsx.writeFile(filename)
		.then(function() {
			console.log('saved now');
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
	insertClass(classEvent) {
		const ws = this.workbook.getWorksheet(classEvent.labName)
		if (!ws) {
			return
		}
		const sTime = classEvent.sTime;
		const eTime = classEvent.eTime;

		const sRow = 2*sTime.getHours() - 12 + sTime.getMinutes()/30
		const eRow = 2*eTime.getHours() - 13 + eTime.getMinutes()/30
		const col = classEvent.sTime.getDay() + 2

		const timeString = (sTime.toString('h:mmt') + ' - ' + eTime.toString('h:mmt')).toLowerCase();

		let fgColor;

		if (classEvent.className == 'CLOSED') {
			fgColor = '606060';
		}
		else if (classEvent.className == 'OPEN') {
			fgColor = 'FFFFFF';
		}
		else {
			fgColor = 'FFFF99';
		}

		ws.getCell(sRow, col).value = classEvent.className;

		if (sRow == eRow) {
			ws.getCell(sRow, col).border = {
				top: {style:'medium'},
				left: {style:'medium'},
				bottom: {style:'medium'},
				right: {style:'medium'}
			}
		}
		else if (eRow - sRow == 1) {
			ws.getCell(eRow, col).value = timeString
			ws.getCell(sRow, col).border = {
				top: {style: 'medium'},
				left: {style:'medium'},
				right: {style:'medium'}
			}
			ws.getCell(eRow, col).border = {
				left: {style:'medium'},
				bottom: {style:'medium'},
				right: {style:'medium'}
			}
		}
		else {
			let instructorName = classEvent.instructorName;
			if (instructorName == 'To Be Determined') {
				instructorName = 'TBD';
			}
			ws.getCell(eRow - 1, col).value = (instructorName) ? instructorName : ''
			ws.getCell(eRow, col).value = timeString
			ws.getCell(sRow, col).border = {
				top: {style: 'medium'},
				left: {style:'medium'},
				right: {style:'medium'}
			}
			ws.getCell(eRow, col).border = {
				left: {style:'medium'},
				bottom: {style:'medium'},
				right: {style:'medium'}
			}

			ws.mergeCells(sRow, col, eRow - 2, col);
		}

		for (var i = sRow; i < eRow + 1; i++) {
			ws.getCell(i, col).fill = {
				type: 'pattern',
				pattern:'solid',
				fgColor:{argb: fgColor}
			}
		}
	}
	createLabColumn(ws, day, sRow, eRow, col) {
		ws.getCell(sRow, col).value = day;
		ws.getCell(sRow, col).font = styles.header.font;
		ws.getCell(sRow, col).alignment = styles.header.alignment;
		ws.getCell(sRow, col).border = {
			top: {style:'medium'},
			left: {style:'medium'},
			bottom: {style:'medium'},
			right: {style:'medium'}
		}

		ws.getColumn(col).width = 14;

		for (var i = 0; i < eRow; i++) {
			ws.getCell(sRow + i + 1, col).value = ''
			ws.getCell(sRow + i + 1, col).font = styles.general.font;
			ws.getCell(sRow + i + 1, col).alignment = styles.general.alignment;

			if (i == 0) {
				ws.getCell(sRow + i + 1, col).border = {
					top: {style:'medium'},
					left: {style:'medium'},
					right: {style:'medium'}
				}
			}
			else if (i == eRow - 1) {
				ws.getCell(sRow + i + 1, col).border = {
					left: {style:'medium'},
					right: {style:'medium'},
					bottom: {style:'medium'}
				}
			}
			else {
				ws.getCell(sRow + i + 1, col).border = {
					left: {style:'medium'},
					right: {style:'medium'}
				}
			}
		}
	}
	createTimeColumn(ws, sRow, sCol, eRow) {
		var date = Date.parse('08:00 AM');

		ws.getColumn(sCol).width = 14;

		ws.getCell(sRow, sCol).value = 'Time';
		ws.getCell(sRow, sCol).font = styles.header.font;
		ws.getCell(sRow, sCol).alignment = styles.header.alignment;
		ws.getCell(sRow, sCol).border = {
			top: {style:'medium'},
			left: {style:'medium'},
			bottom: {style:'medium'},
			right: {style:'medium'}
		}

		for (var i = 0; i < eRow; i++) {
			if (date.getMinutes() != 30) {
				ws.getCell(sRow + i + 1, sCol).value = date.toString('h:mm tt');
			}
			ws.getCell(sRow + i + 1, sCol).font = styles.time.font;
			ws.getCell(sRow + i + 1, sCol).alignment = styles.time.alignment;
			ws.getRow(sRow + i + 1).height = 20;

			ws.getCell(sRow + i + 1, sCol).fill = {
				type: 'pattern',
				pattern:'solid',
				fgColor:{argb: 'FFFFFF'}
			}

			if (i == 0) {
				ws.getCell(sRow + i + 1, sCol).border = {
					top: {style:'medium'},
					left: {style:'medium'},
					right: {style:'medium'}
				}
			}
			else if (i == eRow - 1) {
				ws.getCell(sRow + i + 1, sCol).border = {
					bottom: {style:'medium'},
					left: {style:'medium'},
					right: {style:'medium'}
				}
			}

			date.add({ minutes: 30 });
		}
	}
}

function getDatesFromDateString(date, timeText) {
	const sTempTime = Date.parse(timeText.split('-')[0]);
	const eTempTime = Date.parse(timeText.split('-')[1]);

	const sHour = sTempTime.getHours();
	const eHour = eTempTime.getHours();

	const sMinutes = sTempTime.getMinutes();
	const eMinutes = eTempTime.getMinutes();

	const sTime = new Date(date);
	const eTime = new Date(date);

	sTime.set({
		hour: sHour,
		minute: sMinutes
	});
	eTime.set({
		hour: eHour,
		minute: eMinutes
	})

	return [sTime, eTime];
}

function getClassEvents(date, data) {
	var classEvents = [];
	var elements = $('<div>').html(data)[0].getElementsByClassName('calendar_event_daily');

	for (var i = 0; i < elements.length; i++) {
		var label = elements[i].getElementsByTagName('label')[0].childNodes[0]

		var className = label.nodeValue;
		var labName = $(elements[i]).attr('title');

		var sTime;
		var eTime;

		if ($(label.nextSibling).is('br')) {
			var instructorName = label.nextSibling.nextSibling.nodeValue;
			const timeText = label.nextSibling.nextSibling.nextSibling.innerText;

			const times = getDatesFromDateString(date, timeText);

			sTime = times[0];
			eTime = times[1];

		}
		else {
			const timeText = label.nextSibling.innerText;
			const times = getDatesFromDateString(date, timeText);

			sTime = times[0];
			eTime = times[1];

			var instructorName = null;
		}
		if (className != labName) {
			var classEvent = new ClassEvent(labName, className, instructorName, sTime, eTime);
			classEvents.push(classEvent);
		}
	}
	return classEvents;
}

module.exports = WeekliesGenerator
