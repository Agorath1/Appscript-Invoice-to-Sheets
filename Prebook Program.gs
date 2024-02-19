function start() {
  myFunction();
}


function get_invoice_sheet(sheetID) {
// Gets the spreadsheet for the invoice sheet program
  // Store spreadsheet for main spreadsheet
  try {
    var spreadSheet = SpreadsheetApp.openById(sheetID);
  } catch (error) {
    Logger.log("Error: Spreadsheet with ID '" + sheetID + "' not found or inaccessible.");
  }
  return spreadSheet;
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
  return /^[a-zA-Z]+$/.test(text);
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

function myFunction() {

  // Globals
  var col_des = ["STORE",
                "DEL DT", 
                "DEPT", 
                "DEPT NAME", 
                "QTY ORD", 
                "QTY SHP", 
                "Item Code", 
                "Description", 
                "UPC", 
                "AWG SELL",
                "TOTAL ALLOW",
                "NET COST", 
                "PACK", 
                "EXT NT COST", 
                "FREIGHT",
                "TOTAL WEIGHT",
                "PB"];
  var sheetID = '1p1JJod23yidCAH-5wGkQnmcna2jj27oJE2Vh7tI57Ec';
  var sheet_name = 'Main';
  var invoice_folder_id = '1nTdhCsdEKZmS8QBucbLdZurHCmdJV3Ra';

  // Get the main spreadsheet
  var spreadsheetmain = get_invoice_sheet(sheetID)
  if (spreadsheetmain === undefined) {return;}

  // Grab Folder for Invoices and set file type to pdf
  var folder = DriveApp.getFolderById(invoice_folder_id);  
  var files = folder.getFilesByType(MimeType.PDF);

  // Clears sheet here so that each file loaded doesn't restart the sheet
  var sheet = spreadsheetmain.getSheetByName(sheet_name);
  sheet.clear();
  var sheet_row = 2

  // Create top row
  paste_top_row(col_des, sheet)

  // Establish data variable
  var data = []

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
    while (data_row.length < col_des.length){data_row.push("")} 

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
        data_row[col_des.indexOf("DEL DT")] = text[text_index + 2];
        text_index += 3;
      };

      if (text[text_index + 2] === "5-748423"){
        text_index = text_index
      }

      // Checks if this is the start of an item otherwise go to beginning of the loop.
      if (is_item(text.slice(text_index, text_index + 3)) === false) {continue};

      // Clears data_row information
      for (i = col_des.indexOf("QTY ORD"); i < data_row.length; i++){data_row[i] = ""};

      // Checks for prebook at start
      if (text[text_index - 1] === "PB"){data_row[col_des.indexOf("PB")] = "PB";}      
      
      // Store Qty ord, qty shp, and item code in data row.
      data_row[col_des.indexOf("QTY ORD")] = text[text_index]; text_index++
      data_row[col_des.indexOf("QTY SHP")] = text[text_index]; text_index++
      data_row[col_des.indexOf("Item Code")] = text[text_index].split("-")[1]; text_index++

      // Get the item description
      var i = 0;
      while (true) {
        if (is_upc(text[text_index + i])){break}; i++;
      }
      data_row[col_des.indexOf("Description")] = text.slice(text_index, text_index + i).join(" ");
      text_index += i;

      // Get UPC
      data_row[col_des.indexOf("UPC")] = text[text_index]; text_index++

      // Check if is out of stock
      i = 0
      if (data_row[col_des.indexOf("QTY SHP")] === "0"){
        while (true){
          if (!is_only_letters(text[text_index + i])){break};
          if (text[text_index + i] === "ASSOCIATED" || text[text_index + i] === "VALU"){break};
          i++;
        }
        data_row[col_des.indexOf("EXT NT COST")] = text.slice(text_index, text_index + i).join(" ");
        text_index += i;
        data.push(data_row.slice());
        continue;
      }

      // Get AWG SELL
      data_row[col_des.indexOf("AWG SELL")] = text[text_index]; text_index++;
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

      // Skip UT UNT RTL and EXT NT RTL
      text_index += 2;

      // Check for weight
      if (text[text_index] === "TOTAL" && text[text_index + 1] === "WEIGHT:"){
        data_row[col_des.indexOf("TOTAL WEIGHT")] = text[text_index + 2];
      text_index += 3;
      }
      data.push(data_row.slice());

      // Pushes back one since I will push forward one at the beginning.
      text_index--
    }
  }
  sheet.getRange(2, 1, data.length, col_des.length).setValues(data)
}
