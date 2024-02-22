class Product {
  constructor(item_code, upc, description, pack, awg_sell, dept){
    this.item_code = item_code;
    this.upc = upc;
    this.description = description;
    this.pack = pack;
    this.awg_sell = awg_sell;
    this.dept = dept
  }

  unit_cost(){
    return (Number(this.awg_sell) / Number(this.pack))
  }
}

class Invoice_items extends Product{
  constructor(item_code, upc, description, pack, awg_sell, dept, store, qty_ord, qty_shp, freight, ext_net_cost, total_allow, date_shipped, weight){
    super(item_code, upc, description, pack, awg_sell, dept)
    this.store = store;
    this.qty_ord = qty_ord;
    this.qty_shp = qty_shp;
    this.freight = freight;
    this.ext_net_cost = ext_net_cost;
    this.total_allow = total_allow;
    this.date_shipped = date_shipped;
    this.weight = weight
  }

  misc_fees () {
    return (Number(this.ext_net_cost) - (Number(this.awg_sell) * Number(this.qty_shp) + Number(this.freight) + Number(this.total_allow)))
  }
}
