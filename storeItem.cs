using System;

namespace checker
{
    public class storeItem
    {
        public string StoreName { get; set; }
        public string StoreID { get; set; }
        public string StreetAddress { get; set; }
        public string City { get; set; }
        public int zip { get; set; }
        public bool LEDBulbs { get; set; }
        public bool LEDFixture { get; set; }
        public bool WaterSenseShowerHeads { get; set; }

        public storeItem(string storeName, string storeID, string streetAddress, string city, string zip, bool lEDBulbs, bool lEDFixture, bool waterSenseShowerHeads)
        {
            StoreName = storeName;
            StoreID = storeID;
            StreetAddress = streetAddress;
            City = city;
            int dzip = 0;
            int.TryParse(zip, out dzip);
            this.zip = dzip;
            LEDBulbs = lEDBulbs;
            LEDFixture = lEDFixture;
            WaterSenseShowerHeads = waterSenseShowerHeads;
        }

        public bool valided()
        {
            return !(String.IsNullOrWhiteSpace(StoreID) || String.IsNullOrWhiteSpace(StoreName));
        }
    }
}