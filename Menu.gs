function onOpen(e) {
	var ui = DocumentApp.getUi();

	ui.createMenu('CloudFire')
			.addItem('Delete Future Calendar Events', 'ShowDeletePopup')
			.addItem('Push Events to Google Calendar', 'ShowExportPopup')
			//.addItem('Download .json for INDD Classrooms Signage', 'ChooseSettingsFile')
			//.addItem('Update Classes and Groups filename', 'SetTrigger')
		.addToUi();
}