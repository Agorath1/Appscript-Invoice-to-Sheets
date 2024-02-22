function start() {
  // Runs the main myFunction function
  storefront_invoice_pdf_to_excel();
}


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
  return /^\d+$/.test(text);
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

function create_new_spreadsheet(){
  const new_sheet = SpreadsheetApp.create("Invoice" + Date.now())
  const sheet = new_sheet.getSheets()[0]
  sheet.setName("Main")
  new_sheet.insertSheet("Totals")
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

function storefront_invoice_pdf_to_excel() {

  // Globals
  var col_des = ["STORE",
                "DEPT", 
                "DEPT NAME", 
                "DELV DT", 
                "INVOICE", 
                "PAGE", 
                "QTY ORD", 
                "QTY SHP", 
                "ITEM #", 
                "DESCRIPTION", 
                "PRODUCT UPC", 
                "AWGSELL",
                "TOTAL ALLOW",
                "NET COST", 
                "PACK", 
                "EXT NT COST", 
                "FREIGHT",
                "TOTAL WEIGHT",
                "PB"];
  var col_total = ["STORE",
                "DEPT", 
                "DEPT NAME", 
                "DEL DT", 
                "INVOICE", 
                "PAGE", 
                "QTY ORD", 
                "QTY SHP", 
                "NET COST"];
  
  const sheet_name = 'Main';
  const sheet_total_name = 'Totals';
  var invoice_folder_id = '1nTdhCsdEKZmS8QBucbLdZurHCmdJV3Ra';
  const excel_folder_id = '1jA_iJoT6CZLS8yw_5-R_puLKkh4X1rC1';

  // Get the main spreadsheet
  const spreadsheetmain = create_new_spreadsheet();
  if (spreadsheetmain === undefined) {return;}
  move_file(spreadsheetmain.getId(), excel_folder_id)

  // Grab Folder for Invoices and set file type to pdf
  var folder = DriveApp.getFolderById(invoice_folder_id);  
  var files = folder.getFilesByType(MimeType.PDF);

  // Clears sheet here so that each file loaded doesn't restart the sheet
  var sheet = spreadsheetmain.getSheetByName(sheet_name);
  var sheet_total = spreadsheetmain.getSheetByName(sheet_total_name);
  sheet.clear();
  sheet_total.clear();

  // Create top row
  paste_top_row(col_des, sheet)
  paste_top_row(col_total, sheet_total)

  // Establish data variable
  var data = []
  var total_data = []
  var new_invoice = false

  // Cycle through invoices
  while (files.hasNext()){
    var file = files.next();
    var blob = file.getBlob();
    var resource = {
      title: file.getName(),
      mimeType: MimeType.GOOGLE_DOCS
    }

    // Grabs file from loop and converts to google document
    var newDoc = Drive.Files.create(resource, blob);
    var newDocId = newDoc.id;

    // Grabs the new document and transfers it to the main goole sheet. Deletes doc afterwords
    var text = get_data(newDocId);
    var text_index = -1
    DriveApp.getFileById(newDocId).setTrashed(true);  // Send temp doc to trash

    // Creates data row array
    var data_row = []
    var data_total_row = []
    while (data_row.length < col_des.length){data_row.push("")} 
    while (data_total_row.length < col_total.length){data_total_row.push("")} 
    data_total_row[col_total.indexOf("QTY ORD")] = 0
    data_total_row[col_total.indexOf("QTY SHP")] = 0
    data_total_row[col_total.indexOf("NET COST")] = 0

    while (text_index + 3 < text.length){
      text_index++

      // Gets Store number
      if (text[text_index] === "STORE" && is_only_numbers(text[text_index + 1])) {
        data_row[col_des.indexOf("STORE")] = text[text_index + 1]
        text_index += 2
      }

      // Gets department
      if (text[text_index] === "DEPT" && text[text_index + 1].length === 3) {
        data_row[col_des.indexOf("DEPT")] = text[text_index + 1].slice(0, 2); text_index += 2
        var i = 0;
        while (true){
          if (text[text_index + i] === "DELV"){break};
          i++
        }
        data_row[col_des.indexOf("DEPT NAME")] = text.slice(text_index,text_index + i).join(" ")
        text_index += i
      }
      
      // Gets delivery date
      if (text[text_index] === "DELV" && text[text_index + 1] === "DT:") {
        data_row[col_des.indexOf("DELV DT")] = text[text_index + 2];
        text_index += 3;
      };

      // Get Invoice
      if (text[text_index] === "INVOICE#" && data_row[col_des.indexOf("INVOICE")] != text[text_index + 1]){
        if (data_row[col_des.indexOf("INVOICE")] != ""){new_invoice = true};
        data_row[col_des.indexOf("INVOICE")] = text[text_index + 1];
      }

      // Get Page number
      if (text[text_index] === "PAGE" && is_only_numbers(text[text_index + 1])){
        data_row[col_des.indexOf("PAGE")] = text[text_index + 1];
      }

      // Totals Sheet
      if (new_invoice){
        total_data.push(data_total_row.slice())
        data_total_row[col_total.indexOf("QTY ORD")] = 0
        data_total_row[col_total.indexOf("QTY SHP")] = 0
        data_total_row[col_total.indexOf("NET COST")] = 0
        new_invoice = false
      }

      data_total_row[col_total.indexOf("STORE")] = data_row[col_des.indexOf("STORE")]
      data_total_row[col_total.indexOf("DEPT")] = data_row[col_des.indexOf("DEPT")]
      data_total_row[col_total.indexOf("DEPT NAME")] = data_row[col_des.indexOf("DEPT NAME")]
      data_total_row[col_total.indexOf("DEL DT")] = data_row[col_des.indexOf("DELV DT")]
      data_total_row[col_total.indexOf("INVOICE")] = data_row[col_des.indexOf("INVOICE")]
      data_total_row[col_total.indexOf("PAGE")] = data_row[col_des.indexOf("PAGE")]

      // Checks if this is the start of an item otherwise go to beginning of the loop.
      if (is_item(text.slice(text_index, text_index + 3)) === false) {continue};

      // Clears data_row information
      for (i = col_des.indexOf("QTY ORD"); i < data_row.length; i++){data_row[i] = ""};

      // Checks for prebook at start
      if (text[text_index - 1] === "PB"){data_row[col_des.indexOf("PB")] = "PB";}      
      
      // Store Qty ord, qty shp, and item code in data row.
      data_row[col_des.indexOf("QTY ORD")] = text[text_index]; text_index++
      data_row[col_des.indexOf("QTY SHP")] = text[text_index]; text_index++
      data_row[col_des.indexOf("ITEM #")] = text[text_index].split("-")[1]; text_index++

      data_total_row[col_total.indexOf("QTY ORD")] += Number(data_row[col_des.indexOf("QTY ORD")])
      data_total_row[col_total.indexOf("QTY SHP")] += Number(data_row[col_des.indexOf("QTY SHP")])

      // Get the item description
      var i = 0;
      while (true) {
        if (is_upc(text[text_index + i])){break}; i++;
      }
      data_row[col_des.indexOf("DESCRIPTION")] = text.slice(text_index, text_index + i).join(" ");
      text_index += i;

      // Get UPC
      data_row[col_des.indexOf("PRODUCT UPC")] = text[text_index]; text_index++

      if ("866509" === data_row[col_des.indexOf("ITEM #")]) {
        i = i
      }

      // Check if is out of stock
      i = 0
      if (is_only_letters(text[text_index])){
        while (true){
          if (!is_only_letters(text[text_index + i])){break};
          if (text[text_index + i] === "ASSOCIATED" || text[text_index + i] === "VALU"){break};
          i++;
        }
        data_row[col_des.indexOf("EXT NT COST")] = text.slice(text_index, text_index + i).join(" ");
        text_index += i - 1;
        data.push(data_row.slice());
        continue;
      }

      // Get AWG SELL
      data_row[col_des.indexOf("AWGSELL")] = text[text_index]; text_index++;
      // Get TOTAL ALLOW
      data_row[col_des.indexOf("TOTAL ALLOW")] = text[text_index]; text_index++;
      // Get NET COST
      data_row[col_des.indexOf("NET COST")] = text[text_index]; text_index++;

      // Skip gross %
      text_index++;

      // Get pack size
      if (!contains(text[text_index], ".")){
        data_row[col_des.indexOf("PACK")] = text[text_index];
        text_index++;
      }

      // Get Ext NT Cost
      if (contains(text[text_index + 1], ".")){text_index++};
      data_row[col_des.indexOf("EXT NT COST")] = text[text_index]; text_index++;
      data_total_row[col_total.indexOf("NET COST")] += parseFloat(data_row[col_des.indexOf("EXT NT COST")])

      // Get PB
      if (text[text_index] === "PB"){
        data_row[col_des.indexOf("PB")] = "PB"; 
        text_index++;
      }

      // Get Frieght
      while (true) {
        if (contains(text[text_index],".")){break}
        text_index++
      }
      data_row[col_des.indexOf("FREIGHT")] = text[text_index];text_index++

      // Skip RTL UT
      if (is_only_numbers(text[text_index])){text_index++};

      // Skip UT UNT RTL and EXT NT RTL
      text_index += 2;

      // Check for ITEM OUT OF ST FOR ITEM
      if (text[text_index + 2] === "ITEM"){
        
        var data_row2 = []
        while (data_row2.length < col_des.length) {data_row2.push("")};

        data_row2[col_des.indexOf("QTY ORD")] = text[text_index]; text_index++
        data_row2[col_des.indexOf("QTY SHP")] = text[text_index]; text_index++

        data_total_row[col_total.indexOf("QTY ORD")] += Number(data_row2[col_des.indexOf("QTY ORD")])
        data_total_row[col_total.indexOf("QTY SHP")] += Number(data_row2[col_des.indexOf("QTY SHP")])

        i = 1
        while (true){if (is_only_numbers(text[text_index + i])){break}; i++}

        data_row2[col_des.indexOf("EXT NT COST")] = [...text.slice(text_index, text_index + i),...[data_row[col_des.indexOf("Item Code")]]].join(" ");
        text_index += i;

        data_row2[col_des.indexOf("ITEM #")] = text[text_index]; text_index ++;

        i = 0;
        while (true){if (is_only_numbers(text[text_index + i]) || text[text_index + i] === "TOTAL"){break}; i++;}

        data_row2[col_des.indexOf("DESCRIPTION")] = text.slice(text_index, text_index + i).join(" ");
        text_index += i;

        data.push([...data_row.slice(0, col_des.indexOf("QTY ORD")),...data_row2.slice(col_des.indexOf("QTY ORD"))])
      }

      // Check for weight
      if (text[text_index] === "TOTAL" && text[text_index + 1] === "WEIGHT:"){
        data_row[col_des.indexOf("TOTAL WEIGHT")] = text[text_index + 2];
      text_index += 3;
      }
      data.push(data_row.slice());

      // Pushes back one since I will push forward one at the beginning.
      text_index--
    }
  total_data.push(data_total_row)
  }
  sheet.getRange(2, 1, data.length, col_des.length).setValues(data)
  sheet_total.getRange(2, 1, total_data.length, col_total.length).setValues(total_data)
  spreadsheetmain.setName("Invoice " + Utilities.formatDate(new Date(data[1][3]), "UTC", "YYMMdd"))
}
