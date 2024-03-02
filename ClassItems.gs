class Product {
  constructor(item_code, upc, description, pack, awg_sell, dept, unit_cost){
    this.item_code = item_code;
    this.upc = upc;
    this.description = description;
    this.pack = pack;
    this.awg_sell = awg_sell;
    this.dept = dept;
    this.unit_cost = unit_cost;
  }
}

class SheetManager{
  constructor (sheet, col_list=null) {
    this.sheet = sheet
    this.col_list = col_list
    if (this.col_list === null){
      this.col_list = sheet.getRange(1, 1, 1, sheet.getRange("1:1").getLastColumn()).getValues()[0]
    } else {
      this.set_col_names()
    }
    this.width = this.sheet.getRange("1:1").getLastColumn();
    this.length = this.sheet.getRange("A:A").getLastRow();
    while (true){
      if (this.sheet.getRange(this.length, 1).getValue() === ""){
        this.length -= 1;
      } else {
        break;
      }
    }
  }

  get_index(col_name){
    return this.col_list.indexOf(col_name)
  }
  
  set_col_names(){
    this.sheet.getRange(1, 1, 1, this.col_list.length).set_values(col_list)
  }

  set_value(row, col, value) {
    this.sheet.getRange(row, col).setValue(value)
  }

  set_values(row, col, values){
    this.sheet.getRange(row, col, values.length, values[0].length).setValues(values)
  }

  get_row_values(row, col){
    return this.sheet.getRange(row, col, 1, this.width).getDisplayValues()
  }

  get_column_values(col){
    if (is_only_numbers(col)){
      col = column_num_to_letter(col)
    }
    return(this.sheet.getRange(col + "2:" + col + String(this.length)).getDisplayValues().map(sub_array => sub_array[0]))
  }

  get_pb_row(item_code, date_shipped) {
    let row = 0
    let row_add = 2
    while (true){
      let row_array = this.get_column_values(this.get_index("DELIV DT") + 1).slice(row_add - 2)
      let row_array2 = this.get_column_values(this.get_index("ITEM CD") + 1).slice(row_add - 2)
      let row = row_array.indexOf(date_shipped)
      let row2 = row_array2.indexOf(item_code)
      if (row === row2){return row + row_add}
      if (row === -1 || row2 === -1){return -1}
      if (row > row2){
        row_add += row
      } else {
        row_add += row2
      }
    }
  }
}

class Invoice_items extends Product{
  constructor(item_code, upc, description, pack, awg_sell, dept, unit_cost, store, qty_ord, qty_shp, freight, ext_net_cost, total_allow, date_shipped, weight){
    super(item_code, upc, description, pack, awg_sell, dept, unit_cost)
    this.store = store;
    this.qty_ord = qty_ord;
    this.qty_shp = qty_shp;
    this.date_shipped = date_shipped;
    this.freight = freight;
    this.ext_net_cost = ext_net_cost;
    this.total_allow = total_allow;
    this.weight = weight
  }

  misc_fees () {
    return String(Number(this.ext_net_cost) - (Number(this.awg_sell) * Number(this.qty_shp) + Number(this.freight) + Number(this.total_allow)))
  }

  net_cost (){
    if (is_only_letters(this.awg_sell)){
      return this.awg_sell
    } else {
      return String(Number(this.awg_sell) + Number(this.total_allow))
    }
  }

  weighted (){
    if (this.weight !== ""){
      return "Y"
    } else {
      return ""
    }
  }

  get_array_pb_short(row){
    let pb_array = [
      '=if(isnumber(Y' + row + '),if(R' + row + '<>"",Y' + row + '/X' + row + ',Y' + row + '/(P' + row + '*I' + row + ')),"")',
      this.weight,
      this.ext_net_cost,
      '=if(isnumber(Y' + row + '),ROUND(if(R' + row + '="Y",(N' + row + '+ROUND(AB' + row + ',4))*X' + row + 
      ',(N' + row + '+ROUND(AB' + row + ',2))*P' + row + '),2),"")',
      '=if(isnumber(Y' + row + '),Z' + row + '-Y' + row + ',"")',
      this.freight
    ]
    return pb_array
  }

  get_arrary_pb_full (row){
    let pb_array = [
      this.store,
      "",
      this.dept,
      this.upc,
      "",
      this.item_code,
      this.description,
      "",
      this.pack,
      this.weight,
      this.awg_sell,
      this.total_allow,
      "$0.00",
      this.net_cost(),
      "",
      this.qty_shp,
      this.date_shipped,
      this.weighted(),
      "0",
      "",
      "0",
      "$0.00"
    ]
  return pb_array.concat(this.get_array_pb_short(row))
  }
}
