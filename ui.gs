function excel_invoices() {
  let glob = new InvoiceColumns()
  var files = getFilesInFolder(glob.excel_folder_id)
  var main_sheet = SpreadsheetApp.getActiveSheet
  row = main_sheet.getMaxRow()
  col = main_sheet.getMaxColumn()
  SpreadsheetApp.getActiveRange(1, 1, )
}

function showDialog(){
  let glob = new InvoiceColumns()
  glob.log_sheet.clear()
  var folder_id = glob.excel_folder_id
  sheet_log(glob.log_sheet, "Selecting folder id: " + folder_id)
  var files = getFilesInFolder(glob.excel_folder_id)
  sheet_log(glob.log_sheet, "Sheets in folder: " + files.length)
  var template = HtmlService.createTemplateFromFile('ui');
  template.folder_id = folder_id;
  var html = template.evaluate()
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select File')
}

function updateSidebar(content) {

  var html = HtmlService.createHtmlOutput(content)
  SpreadsheetApp.getUi().showSidebar(html);

}

function processSelectedFile(file_id){
  invoice_to_pb2(file_id)
}

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('My Menu')
      .addItem('Convert Invoice to Excel', 'sf_invoice_pdf_to_excel')
      .addItem('Transfer Invoice to Prebook', 'showDialog')
      // .addSeparator()
      // .addSubMenu(SpreadsheetApp.getUi().createMenu('My sub-menu')
      //     .addItem('One sub-menu item', 'mySecondFunction')
      //     .addItem('Another sub-menu item', 'myThirdFunction'))
      .addToUi();
}


