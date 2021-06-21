using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace checker
{
    class Mou
    {
        //internal controll
        private Boolean loaded;

        //combo box UI use only
        public string fullpath { get; set; }
        public string displayname { get; set; }
        public string businessType;

        //property
        public List<MouItem> mouItems { get; set; }

        public List<storeItem> storeItems { get; set; }
        public Dictionary<int, string> qualityZIPList { get; set; }//<zip,city> paired for performance


        //todo null check
        public Mou(string fullpath, Boolean load = false)
        {
            this.fullpath = fullpath;
            //Debug.Print(fullpath);
            this.displayname = Path.GetFileNameWithoutExtension(fullpath);
            // Debug.Print(displayname);
            this.loaded = false;
            if (load)
            {
                this.load();
                this.loaded = true;
            }
        }



        public void load()
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false,
                ScreenUpdating = false
            };

            Workbook workbook = app.Workbooks.Open(fullpath, AddToMru: false);
            if (workbook == null)
            {
                MessageBox.Show("can not open file " + displayname + " already opened?");
                workbook.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                return;
            }
            try
            {

                loadWorksheet(workbook);
                loadStoreList(workbook);
                loadZipList(workbook);
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                MessageBox.Show(e.Message);
            }

            workbook.Close(0);
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);


        }

        // tool resovle cell's dynamic type issue. convert all to text


        private string getValue(string[] looking, int row, ref Microsoft.Office.Interop.Excel.Range ra, Dictionary<string, int> dict, bool clear = true, string fourceType = "string")
        {
            bool hasAnswer = false;
            string res = null;
            foreach (string l in looking)
            {
                if (dict.ContainsKey(l))
                {
                    res = V(ra.Cells[row, dict[l]], clear, fourceType);
                    //res may null
                    //Debug.WriteLine("getVal" + res);
                    hasAnswer = res != null;
                    if (hasAnswer)
                        return res;
                }
            }

            return res;
        }

        private string V(object k, bool clear = true, string fourceType = "string")
        {
            Microsoft.Office.Interop.Excel.Range t = (Microsoft.Office.Interop.Excel.Range)k;
            if (t != null)
            {
                try
                {
                    //Debug.WriteLine((string)t.NumberFormat);
                    if (((string)t.NumberFormat).Contains("m/d/yy;@") || fourceType.Contains("date")) return ((DateTime)DateTime.FromOADate(t.Value2)).ToString().Trim().ToLowerInvariant();
                    if (t.NumberFormat == "m/d/yyy" || t.NumberFormat == "yyyy/m/dd") { return ((string)t.Text).Trim().ToLowerInvariant(); }
                    if (t.Value2 is string && clear) return String.Concat(((string)t.Value2).Trim().ToLowerInvariant().Where(c => !Char.IsWhiteSpace(c)));
                    if (t.Value2 is string && !clear) return ((string)t.Value2).Trim();
                    if (t.Value2 is int || t.Value2 is decimal || t.Value2 is double || t.Value2 is long) return Convert.ToString(t.Value2);

                    return null;

                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    return null;
                }

            }
            else
            {

                return null;
            }
        }


        private void loadWorksheet(Workbook workbook)
        {
            //sheets arrays start with 1
            Worksheet sh = workbook.Sheets[1];
            sh.Activate();
            Microsoft.Office.Interop.Excel.Range all = sh.UsedRange;
            Debug.WriteLine(all.Rows.Count + "   " + all.Columns.Count);
            if (all != null)
            {
                //init
                Debug.WriteLine(sh.Name);
                this.businessType = sh.Name;
                int rowCount = all.Rows.Count;
                this.mouItems = new List<MouItem>();
                //build header
                Dictionary<string, int> header = buildHeader(all);

                for (int i = 2; i < rowCount; i++)
                {

                    //debug use visualized data
                    //for (int y =1 ; y < 30; y++) {
                    //    string o = V(sh.Cells[i, y]);
                    //    Debug.WriteLine($"row {i} col {y} data {o??"null"}");
                    //}

                    //escape empty row
                    if (V(sh.Cells[i, 0 + 1]) == null) continue;

                    //we have dynamic MOU issue here. using dictionary solution
                    MouItem temp = new MouItem(
                        agreementName: getValue(new string[] { "agreementname" }, i, ref all, header, clear: false) ?? "N/A",
                        mFGmodel: getValue(new string[] { "mfgmodel#" }, i, ref all, header, clear: false) ?? "N/A",
                        retailerSKU: getValue(new string[] { "retailersku#" }, i, ref all, header, clear: false) ?? "N/A",
                        productID: getValue(new string[] { "productid" }, i, ref all, header) ?? "N/A",
                        productDescription: getValue(new string[] { "productdescription" }, i, ref all, header, clear: false) ?? "N/A",
                        startDate: DateTime.Parse(getValue(new string[] { "startdate" }, i, ref all, header, fourceType: "date") ?? "2000/1/1"),
                        endDate: DateTime.Parse(getValue(new string[] { "enddate" }, i, ref all, header, fourceType: "date") ?? "2099/1/1"),
                        active: getValue(new string[] { "active" }, i, ref all, header) == "1" ? true : false,
                        unitPerPack: getValue(new string[] { "unitsperpack" }, i, ref all, header) ?? "0",
                        currentRetail: getValue(new string[] { "currentretail$" }, i, ref all, header) ?? "0",
                        pUDdiscountPerUnit: getValue(new string[] { "ppdiscountperunit", "puddiscountperpack" }, i, ref all, header) ?? "0",
                        additionalPartnerDiscount: getValue(new string[] { "additionalpartnerdiscounts" }, i, ref all, header) ?? "0",
                        totalDiscount: getValue(new string[] { "totaldiscounts" }, i, ref all, header) ?? "0",
                        finalRetail: getValue(new string[] { "finalretail$" }, i, ref all, header) ?? "0",
                        pUDdiscountPerPack: getValue(new string[] { "ppdiscountperpack","puddiscountperpack" }, i, ref all, header) ?? "0",
                        retailPerUnit: getValue(new string[] { "retail$perunit" }, i, ref all, header) ?? "0",
                        productType: getValue(new string[] { "producttype" }, i, ref all, header) ?? "N/A",
                        prodcutCategory: getValue(new string[] { "2020measurecategory" }, i, ref all, header) ?? "N/A",
                        measureSavingUnit: getValue(new string[] { "measuresavings/unit" }, i, ref all, header) ?? "N/A",
                        qualification: getValue(new string[] { "qualification" }, i, ref all, header) ?? "N/A",
                        watts: getValue(new string[] { "watts" }, i, ref all, header) ?? "0",
                        lumens: getValue(new string[] { "lumens" }, i, ref all, header) ?? "0",
                        colorTem: getValue(new string[] { "colortemp" }, i, ref all, header) ?? "0",
                        lumensPerWatt: getValue(new string[] { "lumensperwatt" }, i, ref all, header) ?? "0",
                        savingUnit: getValue(new string[] { "measuresavings/unit" }, i, ref all, header) ?? "0",
                        assignedMOU: getValue(new string[] { "assignedmou#" }, i, ref all, header) ?? "N/A",
                        brand: getValue(new string[] { "brand" }, i, ref all, header, clear: false) ?? "N/A",
                        mFG: getValue(new string[] { "mfg" }, i, ref all, header, clear: false) ?? "N/A",
                        retailer: getValue(new string[] { "retailer" }, i, ref all, header, clear: false) ?? "N/A",
                        everydayLTO: getValue(new string[] { "everyday/lto" }, i, ref all, header) ?? "N/A"
                        );

                    if (temp.validate())
                    {
                        this.mouItems.Add(temp);
                    }
                }
            }
            else
            {
                MessageBox.Show($"empty sheet on {workbook.Name}, sheet {sh.Name}");
            }
        }

        private Dictionary<string, int> buildHeader(Microsoft.Office.Interop.Excel.Range all)
        {
            int colLimit = all.Columns.Count;
            Dictionary<string, int> header = new Dictionary<string, int>();
            //header index dict
            for (int i = 1; i < colLimit; i++)
            {
                if (!header.ContainsKey(V(all.Cells[1, i]) ?? "N/A"))
                    header.Add(V(all.Cells[1, i]) ?? "N/A", i);

                Debug.WriteLine($" header add {V(all.Cells[1, i]) ?? "N/A"}, {i}");
            }
            return header;
        }

        private void loadStoreList(Workbook workbook)
        {
            //sheets arrays start with 1
            //cell arrays start with 1
            Worksheet sh = workbook.Sheets[2];
            sh.Activate();
            Range all = sh.UsedRange;
            if (all != null)
            {
                int rowCount = all.Rows.Count;
                this.storeItems = new List<storeItem>();
                //build header 
                Dictionary<string, int> header = buildHeader(all);

                for (int i = 2; i < rowCount; i++)
                {
                    //skip empty row
                    if (V(sh.Cells[i, 0 + 1]) == null) continue;
                    storeItem temp = new storeItem(
                        storeName: getValue(new string[] { "storename" }, i, ref all, header) ?? "N/A",
                        storeID: getValue(new string[] { "store#orid" }, i, ref all, header) ?? "N/A",
                        streetAddress: getValue(new string[] { "streetaddress" }, i, ref all, header) ?? "N/A",
                        city: getValue(new string[] { "city" }, i, ref all, header) ?? "N/A",
                        zip: getValue(new string[] { "zip" }, i, ref all, header) ?? "N/A",
                        lEDBulbs: (getValue(new string[] { "ledbulbs" }, i, ref all, header) ?? "N/A") == "x" ? true : false,
                        lEDFixture: (getValue(new string[] { "ledfixtures" }, i, ref all, header) ?? "N/A") == "x" ? true : false,
                        waterSenseShowerHeads: (getValue(new string[] { "watersenseshowerheads" }, i, ref all, header) ?? "N/A") == "x" ? true : false
                        );
                    if (temp.valided())
                    {
                        this.storeItems.Add(temp);
                    }
                }
            }
            else
            {
                MessageBox.Show($"empty sheet on {workbook.Name}, sheet {sh.Name}");
            }

        }

        private void loadZipList(Workbook workbook)
        {
            Worksheet sh = workbook.Sheets[3];
            sh.Activate();
            Range all = sh.UsedRange;
            if (all != null)
            {
                this.qualityZIPList = new Dictionary<int, string>();
                for (int i = 3; i < all.Rows.Count; i++)
                {
                    //skip empty row
                    if (V(sh.Cells[i, 0 + 1]) == null) continue;

                    //override if ducplicated zip
                    this.qualityZIPList[Convert.ToInt32(V(sh.Cells[i, 1 + 1]))] = V(sh.Cells[i, 0 + 1]);
                    //no validiction for sample dictionary
                }

            }
            else
            {
                MessageBox.Show($"empty sheet on {workbook.Name}, sheet {sh.Name}");
            }


        }

        public string getRetailer()
        {
            if (!loaded) return null;
            return mouItems[0].Retailer;
        }

        public string getSuplier()
        {
            if (!loaded) return null;
            return mouItems[0].MFG;
        }
    
}
}
