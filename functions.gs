
/**
 * Converts an integer to the equalivant column letters.
 * 
 * @param {integer} col
 * @return {string} letter
 */
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

/**
 * Gets all of the text from a Google Doc. Strips the extra blank spaces and new lines between words.
 * Trims off the spaces from the start and end of the string.
 * It then splits the string into an array by the " "(space).
 * 
 * @param {string} docID
 * @return {string[]}
 */
function get_text_data(docID){
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

/**
 * Pastes the column description row into the top row of target sheet.
 * 
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {string[]} col_des
 * @return null
 */
function paste_top_row(col_des, sheet){
  // Pastes top row into sheet
  sheet.getRange(1, 1, 1, col_des.length).setValues([col_des])
}

/**
 * Checks if the string only contains numbers or "."
 * 
 * @param {string} text
 * @return {string} text
 */
function is_only_numbers(text){
  // Checks if string contains only numbers.
  return /^\d+(\.\d+)?$/.test(text);
}

/**
 * Checks if the string only contains letters or ":" or "/"
 * 
 * @param {string} text
 * return {boolean}
 */
function is_only_letters(text){
  // Checks if it only contains letters
  return /^[a-zA-Z\/:]+$/.test(text);
}

/**
 * Checks for a character in a string
 * 
 * @param {string} text
 * @param {string} character
 * @return {boolean}
 */
function contains(text, character) {
  if (text.indexOf(character) !== -1) {return true}
  return false
}

/**
 * Checks if the string is a upc
 * 
 * @param {string} text
 * @return {boolean}
 */
function is_upc(text){
  if (text.charAt(0) === "0" && text.charAt(1) === "0"){
    return true
  }
  return false
}

/**
 * Checks if the current array section is the start of an item for a Storefront Invoice
 * 
 * @param {[string,string,string]} text_items
 * @return {boolean}
 */
function is_item(text_items){
  // Finds the next item on the invoice
  if (!is_only_numbers(text_items[0]) || contains(text_items[0], ".")) {return false} 
  if (!is_only_numbers(text_items[1]) || contains(text_items[1], ".")) {return false}
  if (!contains(text_items[2], "-")) {return false}
  let text_split = text_items[2].split("-")
  if (is_only_numbers(text_split[0]) && is_only_numbers(text_split[1])) {return true}
  // if ((text_items[2].charAt(2) === "-" || text_items[2].charAt(1) === "-" )){return true}
  return false
}

/**
 * Creates a new spreadsheet and names the sheets based on sheet_name. Returns the new spreadsheet.
 * 
 * @param {string} spread_name
 * @param {string[]} sheet_names
 * @return {SpreadsheetApp.Spreadsheet}
 */
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

/**
 * Removes a file from its current folder and places it in a target folder based on Google Drive ids
 * 
 * @param {string} file_id
 * @param {string} folder_id
 */
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

/**
 * Converts and xls or xlsx file to Google Sheets type file. Returns a string or 
 * an error message.
 * 
 * @param {string} sheetID
 * @return {string}
 */
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

/**
 * Gets the index value of a value from an array
 * 
 * @param {string[]} col_names
 * @param {string} col_name
 * @return {integer}
 */
function col_by_name(col_names, col_name){
  let col = col_names.indexOf(col_name) + 1
  if (col === 0){Logger.log("No column for ${col_name} found.")}
  return col
}

/**
 * Sets a value in a sheet with the input of the sheet, row, column, and value
 * 
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {integer} row
 * @param {integer} col
 * @param {string} value
 */
function set_value(sheet, row, col, value){
  sheet.getRange(row, col).setValue(value)
}

/**
 * Formats a sheet by converting into a filtered list and auto sizing colums
 * 
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {integer[]} [money_col=[]] money_col
 * @return null
 */
function formatsheet(sheet,money_col=[],freight=0) {
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
  sheet.getRange(1, 1, row, col).createFilter();
  sheet.autoResizeColumns(1, col);
  for (i = 0; i < money_col.length; i++) {
    let col_letter = column_num_to_letter(money_col[i])
    sheet.getRange(col_letter + ":" + col_letter).setNumberFormat('"$"#,##0.00')
  }
  let col_letter = column_num_to_letter(freight)
  if (freight > 0) {sheet.getRange(col_letter + ":" + col_letter).setNumberFormat('"$"#,##0.0000')}
};

/**
 * Gets an array of google drive file names and ids from a google drive folder id
 * 
 * @param {string} folder_id
 * @return {[{name: String, id: String},...Array]}
 */
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

/**
 * Adds log information to a log sheet after the very last row.
 * If the sheet doesn't exist it will use the Logger.log() function.
 * 
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {string} msg
 * @return null
 */
function sheet_log(sheet, msg) {
  try{
    sheet.getRange(sheet.getLastRow() + 1, 1).setValue(msg)
  } catch {
    Logger.log(msg)
  }
}

/**
 * Compares two strings and returns a string of the matching characters, starting from the beginning.
 * 
 * @param {string} string1
 * @param {string} string2
 * @return {string}
 */
function common_start(string1, string2) {
  if (string1.length > string2.length){
    var len = string1.length;
  } else {
    var len = string2.length;
  }
  for (i = 0; i < len; i++) {
    if (string1.charAt(i) !== string2.charAt(i)) {break;};
  }
  return string1.substring(0, i)
}

/**
 * Iterates one by one through a range for a matching searchkey. columnNumber is the matching column number being searched in dataRange. Returns the row number.
 * topIndex and botIndex represent the first and last rows of the sheet being checked.
 * 
 * @param {string} searchKey
 * @param {string[]} dataRange
 * @param {integer} columnNumber
 * @param {integer} topIndex
 * @param {integer} botIndex
 * @return {integer}
 */
function customVlookup(searchKey, dataRange, columnNumber, topIndex=1, botIndex=0) {
  topIndex--; botIndex--;
  if (botIndex === -1) {botIndex = dataRange.length}
  for (i = topIndex; i < botIndex; i++) {
    if (dataRange[i][columnNumber - 1] === searchKey) {return (i + 1)}
  }
  return -1
}

/**
 * Binary Search. The column for the columnNumber must be sorted. topIndex and botIndex represent the first and last rows of the sheet being checked.
 * 
 * @param {string} searchKey
 * @param {string[]} dataRange
 * @param {integer} columnNumber
 * @param {integer} topIndex
 * @param {integer} botIndex
 * @return {integer}
 */
function binarySearch(searchKey, dataRange, columnNumber, topIndex=1, botIndex=0) {
  topIndex--; botIndex--;
  if (botIndex === -1) {botIndex = dataRange.length}
  while (true) {
    let midIndex = Math.floor((botIndex - topIndex) / 2 + topIndex)
    if (midIndex >= dataRange.length) {return -1}

    if (dataRange[midIndex][columnNumber - 1] === searchKey) {return midIndex + 1}
    if (dataRange[midIndex][columnNumber - 1] < searchKey) {
      if (topIndex === midIndex) {return -1}
      topIndex = midIndex
    }
    if (dataRange[midIndex][columnNumber - 1] > searchKey) {
      if (botIndex === midIndex) {return -1}
      botIndex = midIndex
    }
  }
}




