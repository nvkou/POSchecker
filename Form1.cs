using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace checker
{
    public partial class Form1 : Form
    {

        public List<string> filesPath = new List<string>();

        public Microsoft.Office.Interop.Excel.Application oExcelApp { get; private set; }

        List<string> skuList = new List<string>();
        List<string> storeIDList = new List<string>();
        List<string> unitPriceList = new List<string>();
        List<string> unitSoldList = new List<string>();
        List<string> modelList = new List<string>();

        public Form1()
        {
            loadSettiong();
            InitializeComponent();
            PrepareData();
        }

        private void PrepareData()
        {
            check_btn.Enabled = false;
            var MOU_data = new BindingSource
            {
                DataSource = BContext.mous
            };
            MOU_combo.DataSource = MOU_data.DataSource;
            MOU_combo.DisplayMember = "displayname";
            MOU_combo.ValueMember = "fullpath";
            try {
                oExcelApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;
                POSattach.Text += oExcelApp.ActiveWorkbook.Name;
            } catch (Exception e) {
                MessageBox.Show("no excel windows found. applicitong will exit now");
               this.Close();
            }
           // oExcelApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;
        }

        private void loadSettiong()
        {
            string mou_floder = Properties.Settings.Default.mou_folder;
            string save_location = Properties.Settings.Default.save_location;

            try
            {
                Path.GetFullPath(mou_floder);
                Debug.WriteLine(mou_floder);
                //Path.GetFullPath(save_location);


            }
            catch (Exception e) {
                MessageBox.Show("could not get saving value. reason " + e.Message + "\r\n initial setting enviroment");
                using (var fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "Please select MOU folder";
                    bool done = false;
                    while (!done)
                    {
                        DialogResult result = fbd.ShowDialog();
                        done = (result == DialogResult.OK && !string.IsNullOrEmpty(fbd.SelectedPath));
                    }
                    //loadpath(fbd.SelectedPath);
                    Properties.Settings.Default.mou_folder = fbd.SelectedPath;
                    mou_floder = fbd.SelectedPath;

                }
                Debug.WriteLine("mou folder loded");

                //using (var sa = new FolderBrowserDialog())
                //{
                //    sa.Description = "select report saving loaction";
                //    bool done = false;
                //    while (!done)
                //    {
                //        DialogResult res = sa.ShowDialog();
                //        done = (res == DialogResult.OK && !string.IsNullOrEmpty(sa.SelectedPath));
                //    }
                //    BContext.bpuSvingLocation = sa.SelectedPath;
                //    Properties.Settings.Default.save_location = sa.SelectedPath;
                //}

                Properties.Settings.Default.Save();

            }
            loadpath(mou_floder);
        }

        private void loadpath(string path)
        {
            Boolean validate(string s)
            {
                string[] ss = s.Split('.');
                return ss[ss.Length - 1].Contains("xls") || ss[ss.Length - 1].Contains("xlsx");
            }
            filesPath = Directory.GetFiles(path).ToList().FindAll(v => validate(v));
            BContext.buildMOU(filesPath);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reset();
            Properties.Settings.Default.Save();
            MessageBox.Show("setting cleared. please restart checker ");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Range temp = SelectRange("Please select SKU column (no heading)");
            skuList = RangeToList(temp);
            skulabel.Text = $"{skuList.Count} skus selected";
            checkInput();
        }

      

        private Range SelectRange(string msg)
        {
            return oExcelApp.InputBox(msg, "Range selector", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
        }
        List<string> RangeToList(Range inputRng)
        {
            object[,] cellValues = (object[,])inputRng.Value2;
            List<string> lst = cellValues.Cast<object>().ToList().ConvertAll(x => Convert.ToString(x));
            return lst;
        }

        public void checkInput()
        {
           
            if (skuList.Count == 0 && modelList.Count == 0) return;
            if (unitSoldList.Count == 0) return;
            if (MOU_combo.SelectedValue == null) return;
            if (
                //either using sku or model must provided
                (skuList.Count + storeIDList.Count + unitPriceList.Count + unitSoldList.Count) / 4 == skuList.Count
                ||
                (modelList.Count + storeIDList.Count + unitPriceList.Count + unitSoldList.Count) / 4 == modelList.Count
                )
            {
                if (BContext.selectedMou != null) check_btn.Enabled = true;
            }
            
            this.Refresh();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            BContext.startDate = dateTimePicker1.Value;

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            BContext.endDate = dateTimePicker2.Value;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button5.Enabled = false;
            BContext.loadMOU(MOU_combo.SelectedValue.ToString());
           
        }
        public void mstatus(string s) {
            moustatus_label.Text = s;
            if (s.Contains("loaded")) { moustatus_label.BackColor = Color.Aqua;  button5.Enabled = true; } else moustatus_label.BackColor = Color.WhiteSmoke;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Range temp = SelectRange("Please select unit price column (no heading)");
            unitPriceList = RangeToList(temp);
            label3.Text = $"{unitPriceList.Count} selected";
            checkInput();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Range temp = SelectRange("Please select unit sold column (no heading)");
            unitSoldList = RangeToList(temp);
            label4.Text = $"{unitSoldList.Count} selected";
            checkInput();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Range temp = SelectRange("Please select store ID column (no heading)");
            storeIDList = RangeToList(temp);
            label2.Text = $"{storeIDList.Count} selected";
            checkInput();
        }

        private void check_btn_Click(object sender, EventArgs e)
        {
            //build POS
            List<Posr> posList = new List<Posr>();
            //MOU_combo.Enabled = false;         
            for (int i = 0; i < skuList.Count; i++)
            {
               
                Posr r = new Posr(
                    sku: skuList.ElementAt(i),
                    storeID: storeIDList.ElementAt(i),
                    unitPrice: unitPriceList.ElementAt(i),
                    salesCount: unitSoldList.ElementAt(i)
                    );
                if (r.valided())
                {                
                    posList.Add(r);
                }
                else
                {
                    MessageBox.Show($"No.{i + 1} skipped info: {r}");
                }

            }

            //report in consol
            posList.ForEach(v => Debug.WriteLine(v.ToString()));
            decimal tot = posList.Sum(item => Convert.ToDecimal(item.total));
            MessageBox.Show($"POS generated size of {posList.Count}. total reported is: {tot} ");
            BContext.workload = posList;
            BContext.HESProcess(oExcelApp.ActiveWorkbook.Name);
            check_btn.Text = "check completed";
        }

        
    }
}
