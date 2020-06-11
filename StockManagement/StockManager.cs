﻿using nsDyeSubStockManagement.Model;
using SpreadsheetReaderLibrary;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace nsDyeSubStockManagement
{
    public partial class StockManager : Form
    {
        private Stream _inputStream;
        SiteFlowEntities ctx = new SiteFlowEntities();

        public StockManager()
        {
            InitializeComponent();
        }

        private void btnDownload_Click(object sender, EventArgs e)
        {

            var date = System.DateTime.Now;

            Thread processThread = new Thread(new ThreadStart(GenerateReport));

            processThread.Start();

        }

        private void GenerateReport()
        {

            if (this.lblStatusDownload.InvokeRequired)
            {
                this.lblStatusDownload.BeginInvoke((MethodInvoker)delegate () { this.lblStatusDownload.Text = "Report Generation is in progress, please wait...."; ; ; });
            }

            if (this.btnDownload.InvokeRequired)
            {
                this.btnDownload.BeginInvoke((MethodInvoker)delegate () { this.btnDownload.Enabled = false; ; ; });
            }

            bool exception = false;
            try
            {
                GenerateOutputSpreadsheet reportEngine = new GenerateOutputSpreadsheet();
                reportEngine.CreateSpreadSheet();

            }
            catch (Exception ex)
            {
                exception = true;

                if (this.lblStatusDownload.InvokeRequired)
                {
                    this.lblStatusDownload.BeginInvoke((MethodInvoker)delegate () { this.lblStatusDownload.Text = "Report Generation failed: " + ex.Message; ; ; });
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate () { this.lblStatusDownload.ForeColor = Color.Green; ; ; });
                }

                if (this.btnDownload.InvokeRequired)
                {
                    this.btnDownload.BeginInvoke((MethodInvoker)delegate () { this.btnDownload.Enabled = true; ; ; });
                }
                return;
            }
            if (!exception)
            {
                if (this.lblStatusDownload.InvokeRequired)
                {
                    this.lblStatusDownload.BeginInvoke((MethodInvoker)delegate () { this.lblStatusDownload.Text = "Report Generation is complete, launching report xlsx..."; ; ; });
                }

                if (this.btnDownload.InvokeRequired)
                {
                    this.btnDownload.BeginInvoke((MethodInvoker)delegate () { this.btnDownload.Enabled = true; ; ; });
                }
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            DialogResult result = openFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog.FileName;
                txtBoxOutput.Text = file;
                FileStream fs = File.OpenRead(file);
                _inputStream = fs;
            }

        }

        private void UploadData()
        {

            try
            {
                if (this.lblStatus.InvokeRequired)
                {
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate () { this.lblStatus.Text = "Upload is in progress, please wait..."; ; ; });
                }

                var data = ImportSpreadsheet.Import(_inputStream);

                foreach (var row in data)
                {
                    var stockName = row["Stock Name".ToLower()];
                    var stockType = row["Stock Type".ToLower()];
                    var stockCategory = row["Stock Category".ToLower()];
                    var substrateName = row["Substrate Name".ToLower()];

                    var sizes = row["Sizes".ToLower()];
                    var colours = row["Colours".ToLower()];
                    var weeksLimitReq = row["Weeks Limit Req.".ToLower()];
                    var WeeksLeft = "0";


                    try
                    {
                        WeeksLeft = Convert.ToString(Math.Round(Convert.ToDecimal(row["Weeks Left".ToLower()])));
                    }
                    catch (Exception e)
                    {

                    }


                    var espStock = row["ESP Stock".ToLower()];
                    var catsStock = row["CATs Stock".ToLower()];
                    var liveStock = row["Live Stock".ToLower()];
                    var spoliage = row["Spoilage".ToLower()];

                    var doa = row["DOA".ToLower()];
                    var AddToSpoilage = row["Add to Spoilage".ToLower()];
                    var AddToDOA = row["Add to DOA".ToLower()];

                    var highestWeek = row["Highest Week".ToLower()];
                    var wk1 = row["WK1".ToLower()];
                    var wk2 = row["WK2".ToLower()];
                    var wk3 = row["WK3".ToLower()];
                    var wk4 = row["WK4".ToLower()];
                    var wk5 = row["WK5".ToLower()];
                    var wk6 = row["WK6".ToLower()];
                    var wk7 = row["WK7".ToLower()];
                    var wk8 = row["WK8".ToLower()];
                    var wk9 = row["WK9".ToLower()];
                    var wk10 = row["WK10".ToLower()];
                    var wk11 = row["WK11".ToLower()];
                    var wk12 = row["WK12".ToLower()];
                    var wk13 = row["WK13".ToLower()];
                    var wk14 = row["WK14".ToLower()];
                    var wk15 = row["WK15".ToLower()];
                    var wk16 = row["WK16".ToLower()];
                    var wk17 = row["WK17".ToLower()];
                    var wk18 = row["WK18".ToLower()];
                    var wk19 = row["WK19".ToLower()];
                    var wk20 = row["WK20".ToLower()];
                    var wk21 = row["WK21".ToLower()];
                    var wk22 = row["WK22".ToLower()];
                    var wk23 = row["WK23".ToLower()];
                    var wk24 = row["WK24".ToLower()];
                    var wk25 = row["WK25".ToLower()];
                    var wk26 = row["WK26".ToLower()];
                    var wk27 = row["WK27".ToLower()];
                    var wk28 = row["WK28".ToLower()];
                    var wk29 = row["WK29".ToLower()];
                    var wk30 = row["WK30".ToLower()];
                    var wk31 = row["WK31".ToLower()];
                    var wk32 = row["WK32".ToLower()];
                    var wk33 = row["WK33".ToLower()];
                    var wk34 = row["WK34".ToLower()];
                    var wk35 = row["WK35".ToLower()];
                    var wk36 = row["WK36".ToLower()];
                    var wk37 = row["WK37".ToLower()];
                    var wk38 = row["WK38".ToLower()];
                    var wk39 = row["WK39".ToLower()];
                    var wk40 = row["WK40".ToLower()];
                    var wk41 = row["WK41".ToLower()];
                    var wk42 = row["WK42".ToLower()];
                    var wk43 = row["WK43".ToLower()];
                    var wk44 = row["WK44".ToLower()];
                    var wk45 = row["WK45".ToLower()];
                    var wk46 = row["WK46".ToLower()];
                    var wk47 = row["WK47".ToLower()];
                    var wk48 = row["WK48".ToLower()];
                    var wk49 = row["WK49".ToLower()];
                    var wk50 = row["WK50".ToLower()];
                    var wk51 = row["WK51".ToLower()];
                    var wk52 = row["WK52".ToLower()];

                    var yearlyTotal = row["Yearly Total".ToLower()];
                    var stockAvailableExternal = row["Stock Available External".ToLower()];
                    var externalSupplier = row["External Supplier".ToLower()];
                    var unitCost = row["Unit Cost".ToLower()];
                    var valueOfStockInHouse = row["Value of Stock in House".ToLower()];


                    var Component1 = row["Component 1".ToLower()];
                    var Component1Qty = row["Component 1 Qty".ToLower()];
                    var Component2 = row["Component 2".ToLower()];
                    var Component2Qty = row["Component 2 Qty".ToLower()];
                    var Component3 = row["Component 3".ToLower()];
                    var Component3Qty = row["Component 3 Qty".ToLower()];
                    var Component4 = row["Component 4".ToLower()];
                    var Component4Qty = row["Component 4 Qty".ToLower()];
                    var Component5 = row["Component 5".ToLower()];
                    var Component5Qty = row["Component 5 Qty".ToLower()];


                    //var Spoilage = row["Spoilage".ToLower()];
                    //var DOA = row["DOA".ToLower()];
                    //var TotalConsumed = row["Total Consumed".ToLower()];

                    var stock = ctx.tDyeSubStocksV2.FirstOrDefault(s => s.Stock_Name == stockName && s.Stock_Type == stockType && s.Stock_Category == stockCategory && s.Substrate_Name == substrateName);

                    bool newStock = false;
                    if (stock == null)
                    {
                        stock = new tDyeSubStocksV2();
                        newStock = true;
                    }

                    if (!string.IsNullOrEmpty(stockName) && !string.IsNullOrEmpty(stockType) && !string.IsNullOrEmpty(substrateName))
                    {
                        stock.Stock_Name = stockName;
                        stock.Stock_Type = stockType;
                        stock.Stock_Category = stockCategory;
                        stock.Substrate_Name = substrateName;
                        stock.Sizes = sizes;
                        stock.Colours = colours;
                        stock.Weeks_Limit_Req_ = weeksLimitReq;
                        stock.Weeks_Left = WeeksLeft;
                        stock.ESP_Stock = espStock;
                        stock.CATs_Stock = catsStock;
                        stock.Live_Stock = liveStock;
                        stock.Highest_Week = highestWeek;
                        stock.WK1 = wk1;

                        stock.WK1 = wk1;
                        stock.WK2 = wk2;
                        stock.WK3 = wk3;
                        stock.WK4 = wk4;
                        stock.WK5 = wk5;
                        stock.WK6 = wk6;
                        stock.WK7 = wk7;
                        stock.WK8 = wk8;
                        stock.WK9 = wk9;
                        stock.WK10 = wk10;
                        stock.WK11 = wk11;
                        stock.WK12 = wk12;
                        stock.WK13 = wk13;
                        stock.WK14 = wk14;
                        stock.WK15 = wk15;
                        stock.WK16 = wk16;
                        stock.WK17 = wk17;
                        stock.WK18 = wk18;
                        stock.WK19 = wk19;
                        stock.WK20 = wk20;
                        stock.WK21 = wk21;
                        stock.WK22 = wk22;
                        stock.WK23 = wk23;
                        stock.WK24 = wk24;
                        stock.WK25 = wk25;
                        stock.WK26 = wk26;
                        stock.WK27 = wk27;
                        stock.WK28 = wk28;
                        stock.WK29 = wk29;
                        stock.WK30 = wk30;
                        stock.WK31 = wk31;
                        stock.WK32 = wk32;
                        stock.WK33 = wk33;
                        stock.WK34 = wk34;
                        stock.WK35 = wk35;
                        stock.WK36 = wk36;
                        stock.WK37 = wk37;
                        stock.WK38 = wk38;
                        stock.WK39 = wk39;
                        stock.WK40 = wk40;
                        stock.WK41 = wk41;
                        stock.WK42 = wk42;
                        stock.WK43 = wk43;
                        stock.WK44 = wk44;
                        stock.WK45 = wk45;
                        stock.WK46 = wk46;
                        stock.WK47 = wk47;
                        stock.WK48 = wk48;
                        stock.WK49 = wk49;
                        stock.WK50 = wk50;
                        stock.WK51 = wk51;
                        stock.WK52 = wk52;
                        stock.YEAR = System.DateTime.Now.Year;
                        stock.Yearly_Total = yearlyTotal;
                        stock.Stock_Available_External = stockAvailableExternal;
                        stock.External_Supplier = externalSupplier;
                        stock.Unit_Cost = unitCost;
                        stock.Value_of_Stock_in_House = valueOfStockInHouse;

                        if (string.IsNullOrEmpty(unitCost))
                        {
                            stock.Value_of_Stock_in_House = null;
                        }
                        else
                        {
                            var unitCostPrice = unitCost.Replace("£", "");
                            var valueofStock = Math.Round(((Convert.ToDecimal(espStock) + Convert.ToDecimal(catsStock)) * Convert.ToDecimal(unitCostPrice)), 1);

                            stock.Value_of_Stock_in_House = Convert.ToString(valueofStock);
                        }
                        stock.Component_1 = Component1;
                        stock.Component_1_Qty = Component1Qty;

                        stock.Component_2 = Component2;
                        stock.Component_2_Qty = Component2Qty;

                        stock.Component_3 = Component3;
                        stock.Component_3_Qty = Component3Qty;

                        stock.Component_4 = Component4;
                        stock.Component_4_Qty = Component4Qty;

                        stock.Component_5 = Component5;
                        stock.Component_5_Qty = Component5Qty;


                    }

                    if (newStock)
                        ctx.tDyeSubStocksV2.Add(stock);

                    ctx.SaveChanges();


                    bool updated = false;

                    if (!string.IsNullOrEmpty(AddToSpoilage))
                    {
                        int spoilageQty = 0;

                        bool conversion = int.TryParse(AddToSpoilage, out spoilageQty);

                        if (conversion)
                        {
                            if (spoilageQty < 0)
                            {
                                stock.Spoilage = Convert.ToString(Convert.ToInt32(stock.Spoilage) + spoilageQty);
                            }
                            else
                            {
                                stock.Spoilage = Convert.ToString(Convert.ToInt32(stock.Spoilage) + spoilageQty);
                                //stock.QuantityAvailable = (stock.QuantityAvailable != null ? stock.QuantityAvailable - spoilageQty : -spoilageQty);
                            }
                            updated = true;
                        }

                    }

                    if (!string.IsNullOrEmpty(AddToDOA))
                    {
                        int doaQTy = 0;

                        bool conversion = int.TryParse(AddToDOA, out doaQTy);

                        if (conversion)
                        {
                            stock.DOA = Convert.ToString(Convert.ToInt32(stock.DOA) + doaQTy);

                            if (doaQTy < 0)
                            {
                                stock.DOA = Convert.ToString(Convert.ToInt32(stock.DOA) + doaQTy);
                            }
                            else
                            {
                                stock.DOA = Convert.ToString(Convert.ToInt32(stock.DOA) + doaQTy);
                                //stock.QuantityAvailable = (stock.QuantityAvailable != null ? stock.QuantityAvailable - doaQTy : -doaQTy);
                            }
                            updated = true;
                        }

                    }

                    if (updated)
                        ctx.SaveChanges();

                }

                if (this.lblStatus.InvokeRequired)
                {
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate () { this.lblStatus.Text = "Upload completed successfully!"; ; ; });
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate () { this.lblStatus.ForeColor = Color.Green; ; ; });
                }

                if (this.btnUpload.InvokeRequired)
                {
                    this.btnUpload.BeginInvoke((MethodInvoker)delegate () { this.btnUpload.Enabled = true; ; ; });
                }

                if (this.btnBrowse.InvokeRequired)
                {
                    this.btnBrowse.BeginInvoke((MethodInvoker)delegate () { this.btnBrowse.Enabled = true; ; ; });
                }
            }
            catch (Exception ex)
            {

                if (this.lblStatus.InvokeRequired)
                {
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate () { this.lblStatus.Text = "Exception: " + ex.Message; ; ; });
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate () { this.lblStatus.ForeColor = Color.Red; ; ; });
                }

                if (this.btnUpload.InvokeRequired)
                {
                    this.btnUpload.BeginInvoke((MethodInvoker)delegate () { this.btnUpload.Enabled = true; ; ; });
                }

                if (this.btnBrowse.InvokeRequired)
                {
                    this.btnBrowse.BeginInvoke((MethodInvoker)delegate () { this.btnBrowse.Enabled = true; ; ; });
                }

            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (_inputStream == null)
            {
                MessageBox.Show("Please browse upload file");
                return;
            }

            Thread processThread = new Thread(new ThreadStart(UploadData));

            btnUpload.Enabled = false;
            btnBrowse.Enabled = false;
            processThread.Start();

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "";
            txtBoxOutput.Text = "";
        }

        private void tabDownload_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblStatusDownload.Text = "";

        }


    }
}
