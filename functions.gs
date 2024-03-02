function get_invoice_sheet(spreadsheetID) {
// Gets the spreadsheet for the invoice sheet program
  // Store spreadsheet for main spreadsheet
  try {
    var spreadSheet = SpreadsheetApp.openById(spreadsheetID);
  } catch (error) {
    Logger.log("Error: Spreadsheet with ID '" + spreadsheetID + "' not found or inaccessible.");
  }
  return spreadSheet;
}

function column_num_to_letter(col){
  let letter = '';
  let temp;

  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26
  }
  return letter
}

function last_row(sheet, col_letter){
  const last_row = sheet.getLastRow();
  const column = sheet.getRange(col_letter + ":" + col_letter + last_row);
  const values = column.getValues

  for (i = values.length - 1; i >= 0; i--){
    if (values[1][0] !== "") {return i}
  }
}

function get_data(docID){
  // Get data from doc
  // Opens the document
  var doc = DocumentApp.openById(docID);
  
  // Grabs everything
  var text = doc.getBody().getText();

  text = text.replace(/\s+/g, ' ');
  text = text.trim();
  text = text.split(" ");

  return text;
}

function paste_top_row(col_des, sheet){
  // Pastes top row into sheet
  sheet.getRange(1, 1, 1, col_des.length).setValues([col_des])
}

function is_only_numbers(text){
  // Checks if string contains only numbers.
  return /^\d+(\.\d+)?$/.test(text);
}

function is_only_letters(text){
  // Checks if it only contains letters
  return /^[a-zA-Z\/:]+$/.test(text);
}

function contains(text, character) {
  if (text.indexOf(character) !== -1) {return true}
  return false
}


function is_upc(text){
  if (text.charAt(0) === "0" && text.charAt(1) === "0"){
    return true
  }
  return false
}

function is_item(text_items){
  // Finds the next item on the invoice
  if (!is_only_numbers(text_items[0])) {return false} 
  if (!is_only_numbers(text_items[1])) {return false} 
  if ((text_items[2].charAt(2) === "-" || text_items[2].charAt(1) === "-" )){return true}
  return false
}

function create_new_spreadsheet(spread_name, sheet_names){
  const new_sheet = SpreadsheetApp.create(spread_name)
  const sheet = new_sheet.getSheets()[0]
  sheet.setName(sheet_names[0])
  if (sheet_names.length > 1) {
    for (i = 1; i < sheet_names.length; i++){
      new_sheet.insertSheet(sheet_names[i])
    }
  }
  return new_sheet
}

function move_file(file_id, folder_id) {
  const folder = DriveApp.getFolderById(folder_id);
  const file = DriveApp.getFileById(file_id);

  const parents = file.getParents();
  while (parents.hasNext()) {
    var parent = parents.next();
    parent.removeFile(file);
  }

  folder.addFile(file);
}

function convert_to_invoice_class(in_row, in_cols) {
  // Creates a class for invoiced items
  return new Invoice_items(
    in_row[in_cols.indexOf("ITEM #")],
    in_row[in_cols.indexOf("PRODUCT UPC")],
    in_row[in_cols.indexOf("DESCRIPTION")],
    in_row[in_cols.indexOf("PACK")],
    in_row[in_cols.indexOf("AWGSELL")],
    in_row[in_cols.indexOf("DEPT")],
    in_row[in_cols.indexOf("NET COST")],
    in_row[in_cols.indexOf("STORE")],
    in_row[in_cols.indexOf("QTY ORD")],
    in_row[in_cols.indexOf("QTY SHP")],
    in_row[in_cols.indexOf("FREIGHT")],
    in_row[in_cols.indexOf("EXT NT COST")],
    in_row[in_cols.indexOf("TOTAL ALLOW")],
    String(Utilities.formatDate(new Date(in_row[in_cols.indexOf("DELV DT")]), "CST", "M/d/yyyy")),
    in_row[in_cols.indexOf("TOTAL WEIGHT")])
}

function import_xls(sheetID) {
  const file = DriveApp.getFileById(sheetID);
  const file_name = file.getName();

  if (file_name.match(/\.(xls|xlsx)$/)){
    const blob = file.getBlob();
    const resource = {
      name: file_name,
      parents: file.getParents(),
      mimeType: MimeType.GOOGLE_SHEETS
    };

    try {
      const converted_file = Drive.Files.copy(resource, sheetID, {convert: true});
      Logger.log("Converted file ID:" + sheetID);
      Drive.Files.remove(sheetID)
      return converted_file.id
    } catch (e) {
      Logger.log("Error converting file ${file_name}: ${e.message}");
    }
  }
  return sheetID
}

function col_by_name(col_names, col_name){
  let col = col_names.indexOf(col_name) + 1
  if (col === 0){Logger.log("No column for ${col_name} found.")}
  return col
}

function set_value(sheet, row, col, value){
  sheet.getRange(row, col).setValue(value)
}

function formatsheet(sheet) {
  row = sheet.getLastRow();
  while (true){
    if (sheet.getRange(row, 1).getValue() === ""){
      row -= 1;
    } else {
      break;
    }
  }
  col = sheet.getLastColumn();
  while (true){
    if (sheet.getRange(1, col).getValue() === ""){
      col -= 1;
    } else {
      break;
    }
  }
  sheet.getRange(1, 1, row, col).activate();
  sheet.autoResizeColumns(1, col);
  sheet.getRange(1, 1, row, col).createFilter();
};

function getFilesInFolder(folder_id) {
  var folder = DriveApp.getFolderById(folder_id);
  var files = folder.getFiles();
  var file_array = []
  while (files.hasNext()) {
    var file = files.next();
    file_array.push({name: file.getName(), id: file.getId()});
  }
  var sort_files = Array.from(file_array)
  sort_files.sort((a, b) => b.name.localeCompare(a.name))
  return sort_files
}

