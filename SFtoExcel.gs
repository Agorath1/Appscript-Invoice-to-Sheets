function sf_invoice_pdf_to_excel() {

  // Global variables
  const glob = new InvoiceColumns()
  glob.log_sheet.clear()

  // Get the main spreadsheet
  const spreadsheetmain = create_new_spreadsheet("Invoice" + Date.now(),["Main", "Totals"]);
  if (spreadsheetmain === undefined) {
    console.log("Failed to create file.");
    return;
  }
  sheet_log(glob.log_sheet,"Created spreadsheet: " + spreadsheetmain.getName())


  // Move file to invoice excel folder
  sheet_log(glob.log_sheet,"Moving file to invoice folder.")
  move_file(spreadsheetmain.getId(), glob.excel_folder_id);

  // Grab Folder for Invoices and set file type to pdf
  sheet_log(glob.log_sheet,"Grabbing files from pdf folder.")
  var folder = DriveApp.getFolderById(glob.invoice_folder_id);
  var files = folder.getFilesByType(MimeType.PDF);

  // Clears sheet here so that each file loaded doesn't restart the sheet
  var sheet = spreadsheetmain.getSheetByName("Main");
  var sheet_total = spreadsheetmain.getSheetByName("Totals");

  // Create top row
  paste_top_row(glob.col_des, sheet)
  paste_top_row(glob.col_total, sheet_total)

  // Establish invoices variable
  var invoices = []

  // Cycle through invoices
  while (files.hasNext()){
    var file = files.next();
    var blob = file.getBlob();
    var resource = {
      title: file.getName(),
      mimeType: MimeType.GOOGLE_DOCS
    }

    sheet_log(glob.log_sheet, "Parsing PDF invoice: " + file.getName())

    // Grabs file from loop and converts to google document
    var newDoc = Drive.Files.create(resource, blob);
    var newDocId = newDoc.id;

    // Grabs the new document and gets the data. Deletes doc afterwords
    var text = get_text_data(newDocId);
    var text_index = -1
    DriveApp.getFileById(newDocId).setTrashed(true);  // Send temp doc to trash

    // Creates Invoice Class
    var invoice = new Invoice()

    while (text_index + 3 < text.length) {
      text_index++;

      // Get Invoice
      if (text[text_index] === "INVOICE#" && invoice.invoice != text[text_index + 1]){
        var invoice = new Invoice();
        invoices.push(invoice);
        invoice.invoice = Number(text[text_index + 1]);
        
        // Get Store Number
        while (text[text_index] != "STORE" || !is_only_numbers(text[text_index + 1])) {text_index--};
        invoice.store = Number(text[text_index + 1]);

        // Get Department
        while (text[text_index] != "DEPT") {text_index++};
        invoice.dept = Number(text[text_index + 1].slice(0, 2)); 

        // Get Delivery Date
        while (text[text_index] != "DELV") {text_index++};
        invoice.delv_date = text[text_index + 2];
        text_index += 5
        sheet_log(glob.log_sheet, "Compiling Invoice: " + invoice.invoice + ", " + invoice.delv_date + ", " + invoice.dept + ": " + invoice.dept_name())
      };

      // Get Page number
      if (text[text_index] === "PAGE" && is_only_numbers(text[text_index + 1])){
        invoice.pages = Number(text[text_index + 1]);
        text_index += 2;
      };

      // Checks if this is the start of an item otherwise go to beginning of the loop.
      if (is_item(text.slice(text_index, text_index + 3)) === false) {continue};
      
      // Create new item class
      var item = new InvoiceItem();
      invoice.items.push(item);

      // Checks for prebook at start
      if (text[text_index - 1] === "PB"){item.pb = "PB";}      
      
      // Store Qty ord, qty shp, and item code in data row.
      item.qty_ord = Number(text[text_index]); text_index++
      item.qty_shp = Number(text[text_index]); text_index++
      item.item_code = Number(text[text_index].split("-")[1]); text_index++

      // Get the item description
      var i = 0;
      while (true) {
        if (is_upc(text[text_index + i])){break}; i++;
      }
      var j = 0;
      if (!is_only_letters(text[text_index + i - 2]) && is_only_letters(text[text_index + i - 1])) {j = 2} else {j = 1}
      item.description = text.slice(text_index, text_index + i - j).join(" ");
      item.size = text.slice(text_index + i - j, text_index + i).join(" ");
      text_index += i;

      // Get UPC
      item.upc = text[text_index]; text_index++

      // Check if is out of stock
      i = 0
      if (is_only_letters(text[text_index])){
        while (true){
          if (!is_only_letters(text[text_index + i])){break};
          if (text[text_index + i] === "ASSOCIATED" || text[text_index + i] === "VALU" || text[text_index + i] === "PB"){break};
          i++;
        }
        item.ext_net_cost = text.slice(text_index, text_index + i).join(" ");
        text_index += i - 1;
        continue;
      }

      // Get AWG SELL
      item.awg_sell = Number(text[text_index]); text_index++;
      // Get TOTAL ALLOW
      item.total_allow = Number(text[text_index]); text_index++;
      
      // Skip NET COST
      text_index++;
      // Skip gross %
      text_index++;

      // Get pack size
      if (!contains(text[text_index], ".")){
        item.pack = Number(text[text_index]);
        text_index++;
      }

      // Get Ext NT Cost
      if (contains(text[text_index + 1], ".")){text_index++};
      item.ext_net_cost = Number(text[text_index]); text_index++;

      // Get PB
      if (text[text_index] === "PB"){
        item.pb = "PB"; 
        text_index++;
      }

      // Get Frieght
      while (true) {
        if (contains(text[text_index],".")){break}
        text_index++
      }
      item.freight = text[text_index];text_index++

      // Skip RTL UT
      if (!contains(text[text_index], ".")) {text_index++};

      // Skip UT UNT RTL and EXT NT RTL
      text_index += 2;

      // Check for ITEM OUT OF ST FOR ITEM
      if (text[text_index + 2] === "ITEM") {
        
        var item2 = new InvoiceItem();
        invoice.items.push(item2);

        item2.qty_ord = Number(text[text_index]); text_index++
        item2.qty_shp = Number(text[text_index]); text_index++

        i = 1;
        while (true){if (is_only_numbers(text[text_index + i]) && (text[text_index + i] !== "ASSOCIATED")){break}; i++};

        item2.ext_net_cost = [...text.slice(text_index, text_index + i),...[item.item_code]].join(" ");
        text_index += i;

        item2.item_code = Number(text[text_index]); text_index ++;

        i = 0;
        while (true){if (is_only_numbers(text[text_index + i]) || text[text_index + i] === "TOTAL" || text[text_index + i] === "EARLY"){break}; i++;}

        item2.description = text.slice(text_index, text_index + i).join(" ");
        text_index += i;
      }

      // Check for early buy
      if (text[text_index] === "EARLY" && text[text_index + 1] === "BUY"){
        text_index += 4;
      }
      // Check for PALLET/CASE
      if (text[text_index] === "PALLET/CASE"){
        text_index += 3;
      }


      // Check for weight
      if (text[text_index] === "TOTAL" && text[text_index + 1] === "WEIGHT:"){
        item.weight = text[text_index + 2];
      text_index += 3;
      }

      // Pushes back one since I will push forward one at the beginning.
      text_index--
    }
  }

  // Loop through invoices to get data
  var data = []
  var data_total = []
  for (i = 0; i < invoices.length; i++) {
    data = data.concat(invoices[i].invoice_array())
    data_total = data_total.concat(invoices[i].totals_array())
  }
  sheet_log(glob.log_sheet, "Setting items into spreadsheet.")
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data)

  sheet_log(glob.log_sheet, "Setting invoice totals into spreadsheet.")
  sheet_total.getRange(2, 1, data_total.length, data_total[0].length).setValues(data_total)

  sheet_log(glob.log_sheet, "Renaming spreadsheet")
  spreadsheetmain.setName("Invoice " + Utilities.formatDate(new Date(invoice.delv_date), "UTC", "YYMMdd"))

  sheet_log(glob.log_sheet, "Formatting spreadsheet")
  formatsheet(sheet,[12, 13, 14, 16], 17)
  formatsheet(sheet_total, [9])
  
  sheet_log(glob.log_sheet, "Complete")
}
