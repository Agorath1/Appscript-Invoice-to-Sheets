function invoice_to_pb2() {
  // Spreadsheets used. The Invoice and Prebooks
  const spreadsheetINID = '1q32kJqkFhgTm-h0NYzGlgyywe35q2zZs6SWgBR_3T3w'  // Make dynamic at some point
  const spreadsheetPBID = '1B-4RojjrbfNoHJL43wsgXC_rrv806M8eAkJIBVY7aR0';
  const in_sheet_name = 'Main';
  const pb_sheet_name = 'Prebooks';

  // Set the sheets as variables.
  const pb_sheet = SpreadsheetApp.openById(spreadsheetPBID).getSheetByName(pb_sheet_name)
  const in_sheet = SpreadsheetApp.openById(spreadsheetINID).getSheetByName(in_sheet_name)

  // Get the columns descriptions
  const pb_cols = pb_sheet.getRange("1:1").getValues()[0]
  const in_cols = in_sheet.getRange("1:1").getValues()[0]

  // Gets the entire date range of the prebook sheet and stores as a variable
  let col_letter = column_num_to_letter(pb_cols.indexOf("DELIV DT") + 1)
  const pb_date_range = pb_sheet.getRange(col_letter + ":" + col_letter).getDisplayValues().map(sub_array => sub_array[0])

  // Gets the entire prebook column of the invoice to search for invoice items that are prebooks
  col_letter = column_num_to_letter(in_cols.indexOf("PB") + 1)
  let in_pb_range = in_sheet.getRange(col_letter + "2:" + col_letter + in_sheet.getLastRow()).getDisplayValues().map(sub_array => sub_array[0])

  // Keeps looping until every prebook is added to the prebook sheet from the invoice sheet.
  while (in_pb_range.indexOf("PB") !== -1){
    
    // Get invoice prebook data
    let in_pb_index = in_pb_range.indexOf("PB")  //Get the first row with PB
    let in_pb_data = convert_to_invoice_class(in_sheet.getRange(in_pb_index + 2, 1, 1, in_cols.length).getValues()[0], in_cols)  // Get data row for that PB
    Logger.log("Prebook found " + in_pb_data.item_code + ": " + in_pb_data.description + " for " + in_pb_data.date_shipped)
    in_pb_range[in_pb_index] = ""  // Clears the PB from the PB column so that it doesn't get it again.

    // Find first row in PB that matches date.
    let pb_date_row = pb_date_range.indexOf(in_pb_data.date_shipped) + 1
    if (pb_date_row === 0){
      var pb_date = ""
    } else {
      var pb_date = pb_sheet.getRange(pb_date_row, pb_cols.indexOf("DELIV DT") + 1).getDisplayValue()
    }
    
    //Loop matching dates
    while (in_pb_data.date_shipped === pb_date && pb_date != "" && pb_date_row <= pb_sheet.getLastRow()) {
      var pb_item_code = pb_sheet.getRange(pb_date_row, pb_cols.indexOf("ITEM CD") + 1).getDisplayValue()
      if (Number(pb_item_code) === Number(in_pb_data.item_code)){
        Logger.log("Found prebook matching " + in_pb_data.date_shipped +" and " + in_pb_data.item_code)
        pb_sheet.getRange(pb_date_row, pb_cols.indexOf("Invoice") + 1).setValue(in_pb_data.ext_net_cost)
        pb_sheet.getRange(pb_date_row, pb_cols.indexOf("Act Weight") + 1).setValue(in_pb_data.weight)
        pb_sheet.getRange(pb_date_row, pb_cols.indexOf("Fee $") + 1).setValue(in_pb_data.freight)
        pb_sheet.getRange(pb_date_row, pb_cols.indexOf("PER") + 1).setValue(Number(in_pb_data.unit_cost()))
        break;
      }
      pb_date_row++;
    }
    if (Number(pb_item_code) === Number(in_pb_data.item_code)){continue;}
    if (pb_date_row === 0) {pb_date_row = pb_sheet.getLastRow() + 1}
    Logger.log("Did not find matching record in PB, creating new record for " + in_pb_data.date_shipped + " and " + in_pb_data.item_code)
    pb_sheet.insertRows(pb_date_row, 1)
    pb_sheet.getRange(pb_date_row, pb_cols.indexOf("Invoice") + 1).setValue(in_pb_data.ext_net_cost)
    pb_sheet.getRange(pb_date_row, pb_cols.indexOf("DELIV DT") + 1).setValue(in_pb_data.date_shipped)
    pb_sheet.getRange(pb_date_row, pb_cols.indexOf("QTY") + 1).setValue(in_pb_data.qty_shp)
    pb_sheet.getRange(pb_date_row, pb_cols.indexOf("UPC") + 1).setValue(in_pb_data.upc)
    pb_sheet.getRange(pb_date_row, pb_cols.indexOf("STORE") + 1).setValue(in_pb_data.store)
    pb_sheet.getRange(pb_date_row, pb_cols.indexOf("RETAIL DEPT") + 1).setValue(in_pb_data.dept)
    pb_sheet.getRange(pb_date_row, pb_cols.indexOf("DESCRIPTION") + 1).setValue(in_pb_data.description)
    pb_sheet.getRange(pb_date_row, pb_cols.indexOf("ITEM CD") + 1).setValue(in_pb_data.item_code)

    if (!is_only_letters(in_pb_data.ext_net_cost)){
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("PER") + 1).setValue(Number(in_pb_data.unit_cost()))
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("Net Cost") + 1).setValue(Number(in_pb_data.ext_net_cost) + (in_pb_data.misc_fees() * Number(in_pb_data.pack)))
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("Difference") + 1).setValue(Math.round(in_pb_data.misc_fees() * Number(in_pb_data.pack) * 100) / 100)
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("Fee $") + 1).setValue(in_pb_data.freight)
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("DEAL") + 1).setValue(-Number(in_pb_data.total_allow))
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("PACK") + 1).setValue(in_pb_data.pack)
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("REG COST") + 1).setValue(in_pb_data.awg_sell)
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("NET COST") + 1).setValue(in_pb_data.awg_sell + in_pb_data.total_allow)
      pb_sheet.getRange(pb_date_row, pb_cols.indexOf("Act Weight") + 1).setValue(in_pb_data.weight)
    }
  }
}

function convert_to_invoice_class(in_row, in_cols) {
  return new Invoice_items(
    in_row[in_cols.indexOf("ITEM #")],
    in_row[in_cols.indexOf("PRODUCT UPC")],
    in_row[in_cols.indexOf("DESCRIPTION")],
    in_row[in_cols.indexOf("PACK")],
    in_row[in_cols.indexOf("AWGSELL")],
    in_row[in_cols.indexOf("DEPT")],
    in_row[in_cols.indexOf("STORE")],
    in_row[in_cols.indexOf("QTY ORD")],
    in_row[in_cols.indexOf("QTY SHP")],
    in_row[in_cols.indexOf("FREIGHT")],
    in_row[in_cols.indexOf("EXT NT COST")],
    in_row[in_cols.indexOf("TOTAL ALLOW")],
    String(Utilities.formatDate(new Date(in_row[in_cols.indexOf("DELV DT")]), "CST", "M/d/yyyy")),
    in_row[in_cols.indexOf("TOTAL WEIGHT")])
}








