function onOpen() {
  DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Sort at Cursor (A-Z)', 'sortAtCursorAtoZ')
      .addItem('Sort at Cursor (Z-A)', 'sortAtCursorZtoA')
      .addToUi();
}
function sortAtCursorAtoZ() {
  sortAtCursorAsc(1)
}
function sortAtCursorZtoA() {
  sortAtCursorAsc(-1)
}
