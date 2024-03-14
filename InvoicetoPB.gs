/**
 * Transfers all prebook items from an excel invoice to a prebook document.
 * Returns a finished message.
 * 
 * @param {String} spreadsheetINID
 * @return {String}
 */
function invoice_to_pb2(spreadsheetINID) {
  const glob = new InvoiceColumns;


  // Name of sheet in prebook Sheet file.
  const pb_sheet_name = 'Prebooks';

  // Convert xls or xlsx file to Google Sheet
  spreadsheetINID = import_xls(spreadsheetINID)

  // Set the sheets as variables.
  sheet_log(glob.log_sheet, "Creating sheet classes")
  const pb_sheet = new SheetManager(SpreadsheetApp.openById(glob.spreadsheetPBID).getSheetByName(pb_sheet_name))
  const in_sheet = new SheetManager(SpreadsheetApp.openById(spreadsheetINID).getActiveSheet())

  // Gets the entire prebook column of the invoice to search for invoice items that are prebooks
  sheet_log(glob.log_sheet, "Grabbing Invoice PB column sheet classes")
  let in_pb_range = in_sheet.get_column_values(in_sheet.get_index("PB") + 1)

  // Keeps looping until every prebook is added to the prebook sheet from the invoice sheet.
  while (in_pb_range.indexOf("PB") !== -1){
    
    // Get invoice prebook data
    sheet_log(glob.log_sheet, "Searching for next prebook in invoice.")
    let in_pb_index = in_pb_range.indexOf("PB")  //Get the first row with PB

    sheet_log(glob.log_sheet, "Grabbing Data.")
    let invoice_data = in_sheet.get_row_values(in_pb_index + 2, 1)[0]

    sheet_log(glob.log_sheet, "Prebook found, creating class for prebook item.")
    let in_pb_data = convert_to_invoice_item(invoice_data, in_sheet.col_list)

    sheet_log(glob.log_sheet, "Searching for prebook " + in_pb_data.item_code + ": " + in_pb_data.description + " in invoice.")

    in_pb_range[in_pb_index] = ""  // Clears the PB from the PB column so that it doesn't get it again.

    // Find matching PB in pb sheet.
    let pb_date_row = pb_sheet.get_pb_row(in_pb_data.item_code, in_pb_data.date)
    if (pb_date_row !== -1){

      // If found add remaining data
      sheet_log(glob.log_sheet, "Found matching prebook of " + in_pb_data.item_code +" for " + in_pb_data.date)
      var pb_data = in_pb_data.get_array_pb_short(pb_date_row)
      pb_sheet.set_values(pb_date_row, pb_sheet.get_index("AQTY") + 1, [pb_data])

    } else {

    // For items not already on the Prebook sheet, sets values up to UPC.
    sheet_log(glob.log_sheet, "Did not find matching record in PB, creating new record of " + 
      in_pb_data.item_code +" for " + in_pb_data.date)

    pb_date_row = pb_sheet.length + 1
    pb_sheet.length += 1
    var pb_data = in_pb_data.get_arrary_pb_full(pb_date_row)
    pb_sheet.set_values(pb_date_row, 1, [pb_data])
    }
  }
  sheet_log(glob.log_sheet, "Finished invoice transer for invoice: " + in_sheet.sheet.getName())
}

/**
 * Creates a class by converting an array into variables. References the row of columns 
 * descriptions to ensure inputs are in correct order.
 * 
 * @param {string[]} in_row
 * @param {string[]} in_cols
 * @return {Invoice_items}
 */
function convert_to_invoice_item(in_row, in_cols) {
  // Creates a class for invoiced items
  let invoice = new PrebookItems()
  // invoice.store = in_row[1]
  // invoice.dept = in_row[2]
  // invoice.date = String(Utilities.formatDate(new Date(in_row[4]), "CST", "M/d/yyyy"))
  // invoice.qty = in_row[7]
  // invoice.item_code = in_row[8]
  // invoice.description = in_row[9]
  // invoice.size = in_row[10]
  // invoice.upc = in_row[11]
  // invoice.awg_sell = in_row[12]
  // invoice.deal = in_row[13]
  // invoice.pack = in_row[15]
  // invoice.ext_net_cost = in_row[16]
  // invoice.freight = in_row[17]
  // invoice.weight = in_row[18]
  invoice.item_code = in_row[in_cols.indexOf("ITEM #")]
  invoice.upc = in_row[in_cols.indexOf("PRODUCT UPC")]
  invoice.description = in_row[in_cols.indexOf("DESCRIPTION")]
  invoice.size = in_row[in_cols.indexOf("SIZE")]
  invoice.pack = in_row[in_cols.indexOf("PACK")]
  invoice.awg_sell = in_row[in_cols.indexOf("AWGSELL")]
  invoice.dept = in_row[in_cols.indexOf("DEPT")]
  invoice.qty = in_row[in_cols.indexOf("QTY SHP")]
  invoice.freight = in_row[in_cols.indexOf("FREIGHT")]
  invoice.ext_net_cost = in_row[in_cols.indexOf("EXT NT COST")]
  invoice.deal = in_row[in_cols.indexOf("TOTAL ALLOW")]
  invoice.weight = in_row[in_cols.indexOf("TOTAL WEIGHT")]
  invoice.date = String(Utilities.formatDate(new Date(in_row[in_cols.indexOf("DELV DT")]), "CST", "M/d/yyyy"))
  invoice.store = in_row[in_cols.indexOf("STORE")]
  return invoice
}
