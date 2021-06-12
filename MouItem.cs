using System;
using System.Collections.Generic;

namespace checker
{
    public class MouItem
    {
        public static Dictionary<string, List<string>> mapping = new Dictionary<string, List<string>>();
        //green zone
        public string AgreementName { get; set; }
        public string MFGmodel { get; set; }
        public string RetailerSKU { get; set; }
        public string ProductID { get; set; }
        public string ProductDescription { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        bool active { get; set; }
        public int unitPerPack { get; set; }

        //red zone
        public decimal currentRetail { get; set; }
        public decimal PUDdiscountPerUnit { get; set; }
        public decimal additionalPartnerDiscount { get; set; }
        public decimal totalDiscount { get; set; }
        public decimal finalRetail { get; set; }
        public decimal measureSavingUnit { get; set; }


        //blue zone{get;set;}
        public decimal PUDdiscountPerPack { get; set; }
        public decimal RetailPerUnit { get; set; }

        //orange zone
        public string ProductType { get; set; }
        public string ProdcutCategory { get; set; }//Q4 2018 measure category

        //purpal zone
        public string qualification { get; set; }
        public decimal watts { get; set; }
        public string lumens { get; set; }
        public string colorTem { get; set; }
        public decimal lumensPerWatt { get; set; }
        public decimal savingUnit { get; set; }
        public string assignedMOU { get; set; }
        public string brand { get; set; }
        public string MFG { get; set; }
        public string Retailer { get; set; }
        public string everydayLTO { get; set; }

        public MouItem(string agreementName, string mFGmodel, string retailerSKU, string productID, string productDescription, DateTime startDate, DateTime endDate, bool active, string unitPerPack, string currentRetail, string pUDdiscountPerUnit, string additionalPartnerDiscount, string totalDiscount, string finalRetail, string pUDdiscountPerPack, string retailPerUnit, string productType, string prodcutCategory, string qualification, string watts, string lumens, string colorTem, string lumensPerWatt, string savingUnit, string assignedMOU, string brand, string mFG, string retailer, string everydayLTO, string measureSavingUnit)
        {
            AgreementName = agreementName;
            MFGmodel = mFGmodel;
            RetailerSKU = retailerSKU;
            ProductID = productID;
            ProductDescription = productDescription;
            this.startDate = startDate;
            this.endDate = endDate;
            this.active = active;
            this.unitPerPack = Convert.ToInt32(unitPerPack);
            this.currentRetail = Convert.ToDecimal(currentRetail);
            PUDdiscountPerUnit = Convert.ToDecimal(pUDdiscountPerUnit);
            this.additionalPartnerDiscount = Convert.ToDecimal(additionalPartnerDiscount);
            this.totalDiscount = Convert.ToDecimal(totalDiscount);
            this.finalRetail = Convert.ToDecimal(finalRetail);
            PUDdiscountPerPack = Convert.ToDecimal(pUDdiscountPerPack);
            RetailPerUnit = Convert.ToDecimal(retailPerUnit);
            ProductType = productType;
            ProdcutCategory = prodcutCategory;
            this.qualification = qualification;
            this.watts = Convert.ToDecimal(watts);
            this.lumens = lumens;
            this.colorTem = colorTem;
            this.lumensPerWatt = Convert.ToDecimal(lumensPerWatt);
            this.savingUnit = Convert.ToDecimal(savingUnit);
            this.assignedMOU = assignedMOU;
            this.brand = brand;
            MFG = mFG;
            Retailer = retailer;
            this.everydayLTO = everydayLTO;
            this.measureSavingUnit = Convert.ToDecimal(measureSavingUnit);
        }

        public MouItem()
        {
        }

        // self check of validating
        public Boolean validate()
        {
            return !String.IsNullOrWhiteSpace(RetailerSKU);
        }

        public override string ToString()
        {
            return $"sku: {RetailerSKU}  model: {MFGmodel} incentive {PUDdiscountPerPack} effective: {startDate.ToString("MM/dd/yyyy")} - {endDate.ToString("MM/dd/yyyy")}";
        }
    }
}