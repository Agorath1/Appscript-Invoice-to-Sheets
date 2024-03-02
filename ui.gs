function excel_invoices() {
  // var files = getFilesInFolder('1jA_iJoT6CZLS8yw_5-R_puLKkh4X1rC1')
  var files = getFilesInFolder(excel_folder_id)
  var main_sheet = SpreadsheetApp.getActiveSheet
  row = main_sheet.getMaxRow()
  col = main_sheet.getMaxColumn()
  SpreadsheetApp.getActiveRange(1, 1, )
}

function showDialog(){
  // var folder_id = '1jA_iJoT6CZLS8yw_5-R_puLKkh4X1rC1'
  var folder_id = excel_folder_id
  var template = HtmlService.createTemplateFromFile('ui');
  template.folder_id = folder_id;
  var html = template.evaluate()
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select File')
}

function test(){
  getFilesInFolder('1jA_iJoT6CZLS8yw_5-R_puLKkh4X1rC1')
}

function updateSidebar(content) {

  var html = HtmlService.createHtmlOutput(content)
  SpreadsheetApp.getUi().showSidebar(html);

}

function processSelectedFile(file_id){
  msg = invoice_to_pb2(file_id)
  
  SpreadsheetApp.getUi().alert(msg)
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

function showFeedbackDialog() {
 var widget = HtmlService.createHtmlOutput("<h1>Enter feedback</h1>");
 SpreadsheetApp.getUi().showModalDialog(widget, "Send feedback");
}
