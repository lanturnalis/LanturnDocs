function onOpen() {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Sort at Cursor', 'sortAtCursor')
      .addToUi();
}
