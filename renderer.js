const electron = require('electron')
const $ = require('jquery')
const exceljs = require('exceljs')
const WeekliesGenerator = require('./weeklies-generator.js')

$(document).ready(function() {
	$('#generate-button').on('click', () => {
		console.log('you pressed generate button');

		// Currently using temporary list of labs and week
		var tempLabNames = ['Phelps 1513', 'SSMS 1303'];
		var tempWeek = 'week'

		var weekliesGenerator = new WeekliesGenerator(tempLabNames, tempWeek);

		weekliesGenerator.generateWeeklies();
	})
})
