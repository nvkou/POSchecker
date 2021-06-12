using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace checker
{
    class BContext
    {

        public static List<Mou> mous = new List<Mou>();
        public static string current_client = "";
        public static Mou selectedMou;
        public static List<Posr> workload;
        public static DateTime startDate { get; set; }
        public static DateTime endDate { get; set; }

        //3rd reference
        public static string bpuSvingLocation { get; set; }

        public static void buildMOU(List<string> path)
        {
            //release last instance
            mous = new List<Mou>();
            foreach (string s in path)
            {
                mous.Add(new Mou(s));
            }
            Debug.Print("mous size " + mous.Count);
            Debug.Print(mous.First().displayname);
        }

        public static void loadMOU(string path) {
            Program.agent.mstatus("loading MOU");
            Task.Factory.StartNew(() => BContext.selectedMou = new Mou(path, true)).ContinueWith((t) => {
                Program.agent.Invoke(new MethodInvoker(delegate () { Program.agent.mstatus("MOU loaded"); }));
                Program.agent.Invoke(new MethodInvoker(delegate () { Program.agent.checkInput(); }));
            });

        }

        public static void HESProcess(string filename)
        {

            //precheck null
            if (selectedMou == null || startDate == null || endDate == null || workload == null)
            {
                MessageBox.Show($"pre check faild");
                return;
            }
            //nullable
            string retailer = selectedMou.getRetailer();
            //nnullable
            string supplier = selectedMou.getSuplier();

            PBU target = new PBU(startDate, endDate, retailer, supplier);


            //following VBA 
            //main loop
            List<string> issues = new List<string>();
            foreach (Posr p in workload)
            {
                // Dictionary<int, string> row = new Dictionary<int, string>();

                //todo what is going on on PBU?
                //find sku mathcing mou or not
                ////check sku ->sku/model
                ///check price declear
                ///
                checkRecord(p, ref issues);

                //target.props.Add(row);

            }
            if (issues.Count > 0)
                writeToReport(issues,filename);

        }

        private static void checkRecord(Posr p, ref List<string> issues)
        {
            bool hasissue = false;
            string mouN = selectedMou.mouItems[0].assignedMOU;
            Debug.WriteLine($"now checking {selectedMou.mouItems[0].assignedMOU}  with {p.sku}");
            //bool skucheck = selectedMou.mouItems.Any(t => same(p.sku, t.RetailerSKU ?? "") || same(p.sku, t.MFGmodel ?? ""));
            List<MouItem> skuMatched = selectedMou.mouItems.FindAll(t => same(p.sku, t.RetailerSKU ?? "") || same(p.sku, t.MFGmodel ?? ""));
            bool skucheck = skuMatched != null && skuMatched.Count > 0;
            if (!skucheck)
            {
                hasissue = true;
                issues.Add($"POS SKU/model of {p.sku} was not found with current MOU {mouN}");
            }
            bool storecheck = selectedMou.storeItems.Any(t => same(p.storeID, t.StoreID ?? "") || same(p.storeID, t.StoreName ?? ""));
            if (!storecheck)
            {
                hasissue = true;
                issues.Add($"POS SKU/model of {p.sku} has Store ID of {p.storeID} but not found in MOU {mouN}");
            }

            if (skucheck)
            {
                //MouItem match = selectedMou.mouItems.Find(t => same(p.sku, t.RetailerSKU ?? "") || same(p.sku, t.MFGmodel ?? ""));
                if (skuMatched == null) throw new Exception("conflict SKU");
                List<MouItem> pricefilter = skuMatched.FindAll(t => t.PUDdiscountPerPack == Convert.ToDecimal(p.unitPrice)); //match.totalDiscount == Convert.ToDecimal(p.unitPrice);
                bool pricecheck = pricefilter != null && pricefilter.Count > 0;
                //if (!pricecheck) {
                //  hasissue = true;
                //issues.Add($"POS SKU/model of {p.sku} has incentive of {p.unitPrice} but MOU {mouN} has no record match");

                //}
                List<MouItem> pricerange = pricefilter.FindAll(t => startDate >= t.startDate && endDate <= t.endDate);
                bool priceRange = pricerange != null && pricerange.Count == 1;
                if (pricerange.Count > 1) issues.Add($"POS SKU/MODEL of {p.sku} has duplicate price; count is {pricerange.Count}");
                if (!priceRange || !pricecheck)
                {
                    hasissue = true;

                    issues.Add($"POS SKU/model of {p.sku}, unit incentive {p.unitPrice} has sales date from {startDate.ToString("MM/dd/yyyy")} to {endDate.ToString("MM/dd/yyyy")} but MOU record mismatch");
                    issues.Add($"\r\n--- possible options ----");
                    foreach (MouItem item in skuMatched)
                    {
                        issues.Add($"MOU candidate {item}");
                    }
                    issues.Add("--- options end --- \r\n \r\n");


                }

            }
            else
            {
                Debug.WriteLine("price check obmit as of sku check fail");
            }




            if (hasissue)
            {
                foreach (string s in issues)
                {
                    Debug.WriteLine(s);
                }
                // writeToReport(issues);
            }
            Debug.WriteLine("############### checking end ################ \r\n");

        }

        private static void writeToReport(List<string> to,string filename)
        {
            //todo not in VSTO env. no globals available
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\"+filename;
            TextWriter tw = new StreamWriter(path + ".txt");

            foreach (String s in to)
                tw.WriteLine(s);

            tw.Close();
            Debug.WriteLine("######### file save ##########");
        }

        private static bool same(string s, string q)
        {
            s = String.Concat(s.Trim().ToLower().Where(c => !Char.IsWhiteSpace(c)));
            q = String.Concat(q.Trim().ToLower().Where(c => !Char.IsWhiteSpace(c)));
            //Debug.WriteLine($"compareing {s} with {q}");
            if (s.Length != q.Length) return false;
            return s.Equals(q);
        }
    }
}
