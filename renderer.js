const electron = require('electron')
const $ = require('jquery')
require('exceljs')
require('datejs')

const WeekliesGenerator = require('./weeklies-generator.js')

$(document).ready(function() {
	$('#generate-button').on('click', () => {
		console.log('you pressed generate button');

		// Currently using temporary list of labs and week
		var tempLabNames = ['Phelps 1513', 'SSMS 1303'];
		var startDate = new Date.parse('4-Sep-2016');

		var weekliesGenerator = new WeekliesGenerator(tempLabNames, startDate);

		weekliesGenerator.generateWeeklies();
	})
})
