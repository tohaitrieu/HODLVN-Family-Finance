function openIncomePage() {
  const html = HtmlService.createHtmlOutputFromFile('IncomeForm')
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Trang thu nhập');
}
