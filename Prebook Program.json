function myFunction() {

  // Clears sheet 2 here so that each file loaded doesn't restart the sheet
  var sheet2 = get_invoice_sheet().getSheetByName('Sheet2');
  sheet2.clear();

  // Grab Folder for Invoices and set file type to pdf
  var folder = DriveApp.getFolderById('1nTdhCsdEKZmS8QBucbLdZurHCmdJV3Ra');  
  var files = folder.getFilesByType(MimeType.PDF);

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
    var lines = dataTransfer(newDocId);
    DriveApp.getFileById(newDocId).setTrashed(true);

    // Converts plain text into cells on sheet
    splitText(lines);
  }
}

// Gets the spreadsheet for the invoice sheet program
function get_invoice_sheet() {
  // Store spreadsheet for main spreadsheet
  var sheetID = '1p1JJod23yidCAH-5wGkQnmcna2jj27oJE2Vh7tI57Ec';  
  var spreadSheet = SpreadsheetApp.openById(sheetID);
  return spreadSheet;
}

// Transfer pdf to google doc
function dataTransfer(docID){

  // Opens the document
  var doc = DocumentApp.openById(docID);

  // Grabs everything
  var text = doc.getBody().getText();

  // Splits the google doc based on new lines
  var lineSplit = '\n'
  var lines = ""
  var lines = text.split(lineSplit);

  return lines
}

// Finds the last row of a sheet by checking the first column. A bit finiky, will change later.
function find_last_row(sheet, skip, sheet_row){
  while (true){
    sheet_row += skip
    if (sheet.getRange(sheet_row, 1).getValue() === "") {
      if (skip === 1) {
        return sheet_row
      } else {
        sheet_row -= skip
        skip = Math.ceil(skip / 10)
      }
    }
  }
}

// This function gets the split text on sheet 1 and rearranges it to a more readable format.
function splitText(lines){

  // Get the spreadsheet to store in a var along with the two main sheets stored in seperate var.
  var spreadSheet = get_invoice_sheet();
  var sheet2 = spreadSheet.getSheetByName('Sheet2');

  var column_descr = ["DEL DT", 
                      "DEPT", 
                      "QTY ORD", 
                      "QTY SHP", 
                      "Item Code", 
                      "Description", 
                      "UPC", 
                      "Act Weight",
                      "Net Cost", 
                      "Total Allow",
                      "Del Fee",
                      "Total Cost", 
                      "PB"]

  // Clears the sheet for new data
  sheet2.getRange(1, 1, 1, column_descr.length).setValues([column_descr])
  
  // Checks to see what data is on sheet 2 and gets the first empty row.
  var sheet2Row = find_last_row(sheet2, 100, 1)
  var sheet2_start_row = sheet2Row
  var data = []

  //Skipping the first row, loops through each row
  for (var i = 0; i < lines.length; i++){

    // Grabs the current row and splits it into an array
    var splitValue = lines[i].split(' ');
    var split_value_len = splitValue.length
    if (i < lines.length - 1) {
      splitValue = splitValue.concat(lines[i + 1].split(' '))
    }
    
    // Loops through the split text
    for (var k = 0; k + 3 < split_value_len; k++){

      if (splitValue[k + 2] ===  "3-651044")
      {
        k = k
      }

      // Department
      if (splitValue[k] === "DEPT") {
        var dept = splitValue[k + 1].slice(0, 2)
      }

      // Delivery Date
      if (splitValue[k] === "DT:") {
        var del_date = splitValue[k + 1]
      }

      // Checks if this is the start of a invoice line
      if (itemStartCheck(splitValue, k)){
        var data_row = []
        while (data_row.length < column_descr.length)
          data_row.push("")

        data_row[column_descr.indexOf("DEL DT")] = del_date
        data_row[column_descr.indexOf("DEPT")] = dept

        // Checks for PB word, not sure if this ever runs.
        if (splitValue[k - 1] == ("PB")) {
          data_row[column_descr.indexOf("PB")] = splitValue[k - 1]; // PB Check
        }

        // Initial Values
        data_row[column_descr.indexOf("QTY ORD")] = splitValue[k];                      //Ordered QTY
        data_row[column_descr.indexOf("QTY SHP")] = splitValue[k + 1];                  //Delivered QTY
        data_row[column_descr.indexOf("Item Code")] = splitValue[k + 2].split('-')[1];  //Item Code

        //Finds the name
        var itemName = findName(splitValue, k + 3);     
        k = itemName[1];
        data_row[column_descr.indexOf("Description")] = itemName[0];
        
        // UPC
        data_row[column_descr.indexOf("UPC")] = splitValue[k];

        // Prebook
        var PB = splitValue.slice(k,k+15).indexOf("PB")
        if (PB !== -1) {data_row[column_descr.indexOf("PB")] = "PB"}

        // Checks if the net cost has a letter to determine zeroed quantity.
        if (containsLetter(splitValue[k + 1])) {
          ///////////////////////
          // Not Shipped Items //
          ///////////////////////
          var itemName = findNameForOuts(splitValue, k + 1);
          k = itemName[1] // Find all the out words
          
          data_row[column_descr.indexOf("Net Cost")] = itemName[0]; // Out Reason

        } else {
          
          var weight = splitValue.slice(k,k+30).indexOf("WEIGHT:")
          if (weight !== -1) {
            ///////////////////////////
            // Weighted Product Info //
            ///////////////////////////
            data_row[column_descr.indexOf("Act Weight")] = splitValue[k + weight + 1];  // Pack Size
            data_row[column_descr.indexOf("Total Cost")] = splitValue[k + 6];   // Total Value
            // Net Cost
            data_row[column_descr.indexOf("Net Cost")] = splitValue[k + 1] * data_row[column_descr.indexOf("Act Weight")];
            // Allowance
            data_row[column_descr.indexOf("Total Allow")] = splitValue[k + 2] * data_row[column_descr.indexOf("Act Weight")];
            // Delivery Fee
            var weight_fee = splitValue[k + weight - 5] 
            data_row[column_descr.indexOf("Del Fee")] = weight_fee * data_row[column_descr.indexOf("Act Weight")];

          } else if (parseInt(data_row[column_descr.indexOf("UPC")]) === 0 || splitValue[k + 7] === ""){
            //////////////////
            // Shipper info //
            //////////////////
            data_row[column_descr.indexOf("Total Cost")] = splitValue[k + 6];   // Total Value
            // Net Cost
            data_row[column_descr.indexOf("Net Cost")] = splitValue[k + 1] * data_row[column_descr.indexOf("QTY SHP")];
            // Allowance
            data_row[column_descr.indexOf("Total Allow")] = splitValue[k + 2] * data_row[column_descr.indexOf("QTY SHP")];
            // Delivery Fee
            while (containsLetter(splitValue[k + 10])) {
              k = k + 1
            }
            data_row[column_descr.indexOf("Del Fee")] = splitValue[k + 10] * data_row[column_descr.indexOf("QTY SHP")];

          } else if (data_row[column_descr.indexOf("QTY ORD")] != "") {
            /////////////////////////
            // Normal Product Info //
            /////////////////////////
            data_row[column_descr.indexOf("Total Cost")] = splitValue[k + 7];   // Total Value
            // Net Cost
            data_row[column_descr.indexOf("Net Cost")] = splitValue[k + 1] * data_row[column_descr.indexOf("QTY SHP")];
            // Allowance
            data_row[column_descr.indexOf("Total Allow")] = splitValue[k + 2] * data_row[column_descr.indexOf("QTY SHP")];
            // Delivery Fee
            while (containsLetter(splitValue[k + 10])) {
              k = k + 1
            }
            data_row[column_descr.indexOf("Del Fee")] = splitValue[k + 10] * data_row[column_descr.indexOf("QTY SHP")];

          } else {
            ////////////////////////
            // Shipper Components //
            ////////////////////////
            data_row[column_descr.indexOf("Net Cost")] = ""
            data_row[column_descr.indexOf("Total Allow")] = ""
            data_row[column_descr.indexOf("PB")] = ""
          }
        }

        // Pastes information into sheet
        data.push(data_row)
        sheet2Row++;
      }
    }
  }
  sheet2.getRange(sheet2_start_row, 1, data.length, data_row.length).setValues(data)
}

function findNameForOuts(currentString, i){
  // Find the name for outs
  //Returning the i and j will establish the locations of the string on the array.
  var name = currentString[i]
  for (var j = i + 1; j < currentString.length; j++){
    if (containsNumber(currentString[j]) || currentString[j] === "PB" || currentString[j] === ""){
      return [name, j - 1];
    }
    name = name + ' ' + currentString[j];
  }
  return [name, j];
}

function findName(currentString, i){
  //This will check how long the name is by finding the first section of the sentence that starts with zero.
  //This hasn't been hard test but if need be I can add a second zero. I haven't seen a UPC that fills the entire 
  //section with whole numbers.
  //Returning the i and j will establish the locations of the string on the array.
  var name = currentString[i]
  for (var i = i + 1; i <= currentString.length; i++){
    if (currentString[i].charAt(0) === '0' && currentString[i].charAt(1) === '0'){
      return [name, i];
    }
    name = name + ' ' + currentString[i];
  }
}

// Function used to find the start of invoice line
function itemStartCheck(currentString, i){
  if (currentString.length <= i + 3){return false;}         //Checks if the string ends before the word can start
  if (containsLetter(currentString[i])){return false;}      //First number should be a number
  if (containsLetter(currentString[i + 1])){return false;}  //Second number should be a number
  if (containsNumber(currentString[i + 2].charAt(0))){
    if (currentString[i + 2].charAt(2) === '-' || currentString[i + 2].charAt(1) === '-'){
      return true;                                           //Third sections should have a dash as the third or second character.
    }
  }
  return false;
}

// Function to find if their is a number in the string
function containsNumber(str) {
  if (str === ''){return false;}
  return /\d/.test(str); // Returns true if string contains a digit
}

// Function to find if their is a letter in the string
function containsLetter(str) {
  if (str === ''){return false;}
  return /[a-zA-Z]/.test(str); // Returns true if string contains a letter
}
