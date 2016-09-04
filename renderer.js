const electron = require('electron')
const $ = require('jquery')

$(document).ready(function() {
	$('#east-button').on('click', function() {
		console.log('you pressed east')
	})
})
