using System;
using System.Collections.Generic;
using System.Collections.Specialized;

namespace checker
{
    internal class PBU
    {
        //prop holder
        public List<Dictionary<int, string>> props { get; set; }
        //level 1 header
        public OrderedDictionary mainHeader { get; set; }

        public OrderedDictionary secondHeader { get; set; }
        public OrderedDictionary thirdHeader { get; set; }
        public string title;

        public PBU(DateTime start, DateTime end, string retailer = "", string supplier = "")
        {
            if (retailer != null) retailer = "-" + retailer;
            if (supplier != null) supplier = "-" + supplier;
            this.title = "PBU-V2-Upstream PBU_v1" + retailer + supplier + "_" + start.ToString("MM.dd.yy") + "-" + end.ToString("MM.dd.yy");
            mainHeader = new OrderedDictionary
            {
                { 1, "Batch Review" }
            };

            //second line
            secondHeader = new OrderedDictionary
            {
                { 1, "ApplicationNumber" },
                { 2, "Retailer Information" },
                { 7, "Payee" },
                { 14, "Project Information" },
                { 23, "Begin Measure Group" },
                { 24, "Container Measure" },
                { 50, "End Measure Group" },
                { 51, "Application Status" },
                { 52, "Status" }
            };

            //third line key first for matching
            thirdHeader = new OrderedDictionary
            {
                { "Application Number", 1 },
                { "Retailer", 2 },
                { "Retailer Address", 3 },
                { "Retailer City", 4 },
                { "Retailer State", 5 },
                { "Retailer Zip", 6 },
                { "Company Name", 7 },
                { "Attention To", 8 },
                { "Address1", 9 },
                { "City", 10 },
                { "State/Province", 11 },
                { "Zip/Postal Code", 12 },
                {"Country",13 },
                {"AssignedMOU#",14},
                {"AgreementName",15 },
                {"Amendment",16 },
                {"Reporting Start Date",17 },
                {"Reporting End Date",18 },
                {"Reporting Month",19 },
                {"Salse Start Date",20 },
                {"Salse End Date",21 },
                {"Sales Month",22 },
                {"Claim ID",24 },
                {"Retailer Location",25 },
                {"Retail Address",26 },
                {"Retail Zip",27 },
                {"Product Type",28 },
                {"Q4 2017 Measure Caregory",29 },
                {"Q2 2016 Product Category",30 },
                {"Qualification",31 },
                {"MFG Model",32 },
                {"RetailerSKU #",33 },
                {"ProductDescription",34 },
                {"Lumens",35 },
                {"PUDDiscountPer Unit",36 },
                {"Current Retail",37 },
                {"MFG",38 },
                {"Incentive Paid",39 },
                {"Everyday LTO",40 },
                {"EFC Number",41 },
                {"Pack Sold",42 },
                {"Unit Sold",43 },
                {"FinalRetail$",44 },
                {"Retail$PerUnit",45 },
                {"Measure kWh Saving",46 },
                {"Measure kWh Saving This Period",47 },
                {"Base Type",48 },
                {"Bulb Type",49 },
                {"Application Status",51 },
                {"Status",52 }
            };



        }
    }
}