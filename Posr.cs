using System;

namespace checker
{
    public class Posr
    {

        public string storeID { get; set; }
        public string sku { get; set; }
        public string unitPrice { get; set; }
        public string salesCount { get; set; }
        //total is readonly from outside scope
        public string total { get; set; }

        public Posr(string storeID, string sku, string unitPrice, string salesCount)
        {
            this.storeID = storeID;
            this.sku = sku;
            this.unitPrice = unitPrice;
            this.salesCount = salesCount;
        }

        public Boolean valided()
        {
            Boolean check = String.IsNullOrEmpty(storeID) || String.IsNullOrEmpty(sku) || String.IsNullOrEmpty(unitPrice) || String.IsNullOrEmpty(salesCount);
            if (!check) this.total = (Convert.ToDecimal(unitPrice) * Convert.ToInt32(salesCount)).ToString();
            return !check;

        }

        public override string ToString()
        {
            return $"storeID: {storeID}  sku:{sku}  unitePrice:{unitPrice}  sales:{salesCount}   total:{total}";
        }
    }
}