function invoice_to_pb2(spreadsheetINID) {
  const glob = new InvoiceColumns;

  // Spreadsheets used. The Invoice and Prebooks
  const pb_sheet_name = 'Prebooks';
  var logs = ''

  spreadsheetINID = import_xls(spreadsheetINID)
  if (spreadsheetINID === "" || spreadsheetINID === undefined){
    return "Invoice retrieval failed."
  }

  // Set the sheets as variables.
  const pb_sheet = new SheetManager(SpreadsheetApp.openById(glob.spreadsheetPBID).getSheetByName(pb_sheet_name))
  const in_sheet = new SheetManager(SpreadsheetApp.openById(spreadsheetINID).getActiveSheet())

  // Gets the entire prebook column of the invoice to search for invoice items that are prebooks
  let in_pb_range = in_sheet.get_column_values(in_sheet.get_index("PB") + 1)

  // Keeps looping until every prebook is added to the prebook sheet from the invoice sheet.
  while (in_pb_range.indexOf("PB") !== -1){
    
    // Get invoice prebook data
    let in_pb_index = in_pb_range.indexOf("PB")  //Get the first row with PB
    let in_pb_data = convert_to_invoice_class(in_sheet.get_row_values(in_pb_index + 2, 1)[0], in_sheet.col_list)  // Get data row for that PB

    logs += "Prebook " + in_pb_data.item_code + ": " + in_pb_data.description + "<br>"
    updateSidebar(logs)
    in_pb_range[in_pb_index] = ""  // Clears the PB from the PB column so that it doesn't get it again.

    // Find matching PB in pb sheet.
    let pb_date_row = pb_sheet.get_pb_row(in_pb_data.item_code ,in_pb_data.date_shipped)
    if (pb_date_row !== -1){

      // If found add remaining data
      logs += "Found prebook matching " + in_pb_data.date_shipped +" and " + in_pb_data.item_code + "<br>"
      updateSidebar(logs)
      var pb_data = in_pb_data.get_array_pb_short(pb_date_row)
      pb_sheet.set_values(pb_date_row, pb_sheet.get_index("PER") + 1, [pb_data])

    } else {

    // For items not already on the Prebook sheet, sets values up to UPC.
    logs += "Did not find matching record in PB, creating new record for " + 
      in_pb_data.date_shipped + " and " + in_pb_data.item_code + "<br>"
    updateSidebar(logs)

    pb_date_row = pb_sheet.length + 1
    pb_sheet.length += 1
    var pb_data = in_pb_data.get_arrary_pb_full(pb_date_row)
    pb_sheet.set_values(pb_date_row, 1, [pb_data])
    }
  }
  return "Finished invoice transer for invoice: " + in_sheet.sheet.getName()
}
