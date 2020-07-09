using nsDyeSubStockManagement.Model;
using SpreadsheetReaderLibrary;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization.Configuration;

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

            Thread processThread = new Thread(new ThreadStart(GenerateReport));

            processThread.Start();

        }

        public static int GetIso8601WeekOfYear(DateTime time)
        {
            //  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
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

                    var extra = row["Extra".ToLower()];

                    var sizes = row["Sizes".ToLower()];

                    if (!string.IsNullOrEmpty(sizes))
                    {
                        var finalSize = GetFinalSizeInMM(sizes);
                        sizes = finalSize;
                    }
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
                    var liveStock = row["Live Stock".ToLower()];
                    var spoliage = row["Spoilage".ToLower()];

                    var doa = row["DOA".ToLower()];
                    var AddToESPStock = row["Add to ESP Stock".ToLower()];
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

                    if (!string.IsNullOrEmpty(stockType) && !string.IsNullOrEmpty(stockCategory))
                    {
                        stock.Stock_Name = stockName;
                        stock.Stock_Type = stockType;
                        stock.Stock_Category = stockCategory;
                        stock.Substrate_Name = substrateName;
                        stock.Extra = extra;
                        stock.Sizes = sizes;
                        stock.Colours = colours;
                        stock.Weeks_Limit_Req_ = weeksLimitReq;
                        stock.Weeks_Left = WeeksLeft;
                        stock.ESP_Stock = espStock;
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

                        if (string.IsNullOrEmpty(unitCost) || string.IsNullOrEmpty(espStock))
                        {
                            stock.Value_of_Stock_in_House = null;
                        }
                        else
                        {
                            var unitCostPrice = unitCost.Replace("£", "");
                            var valueofStock = Math.Round((Convert.ToDecimal(espStock) * Convert.ToDecimal(unitCostPrice)), 1);

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

                    var weekno = GetIso8601WeekOfYear(System.DateTime.Now);

                    bool updated = false;

                    if (!string.IsNullOrEmpty(AddToESPStock))
                    {
                        int addToEspStockQty = 0;

                        bool conversion = int.TryParse(AddToESPStock, out addToEspStockQty);

                        if (conversion)
                        {
                            if (stock.ESP_Stock == null)
                                stock.ESP_Stock = Convert.ToString(AddToESPStock);
                            else
                                stock.ESP_Stock = Convert.ToString(Convert.ToInt32(stock.ESP_Stock) + AddToESPStock);

                            updated = true;
                        }

                    }

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
                                stock.ESP_Stock = (stock.ESP_Stock != null ? Convert.ToString(Convert.ToInt32(stock.ESP_Stock) - spoilageQty) : Convert.ToString(-spoilageQty));
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
                                stock.ESP_Stock = (stock.ESP_Stock != null ? Convert.ToString(Convert.ToInt32(stock.ESP_Stock) - doaQTy) : Convert.ToString(-doaQTy));
                            }
                            updated = true;
                        }

                    }

                    GetWeekMax(stock, weekno);
                    ctx.SaveChanges();


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

        private static string GetFinalSizeInMM(string sizes)
        {
            var sizeArray = sizes.Split(new char[] { ',' });

            var finalSize = "";
            foreach (var sizeData in sizeArray)
            {
                var singleSizearray = sizeData.Split(new char[] { 'x' });

                foreach (var singlesizePart in singleSizearray)
                {
                    if (singlesizePart.Contains("\""))
                    {
                        var temp = singlesizePart.Replace("\"", "");

                        double tempInt = 0d;

                        double.TryParse(temp, out tempInt);

                        tempInt = Convert.ToDouble(tempInt * 25.4);
                        finalSize += tempInt + "x";
                    }
                    else
                    {
                        finalSize += singlesizePart + "x";
                    }
                }

                if (finalSize.EndsWith("x"))
                {
                    finalSize = finalSize.Remove(finalSize.LastIndexOf("x"), 1);
                }

                finalSize += ",";
            }

            finalSize = finalSize.Replace(" ", "");
            finalSize = finalSize.Replace("mm", "");

            if (finalSize.EndsWith(","))
                finalSize = finalSize.Remove(finalSize.LastIndexOf(","), 1);

            return finalSize;
        }

        private static void GetWeekMax(tDyeSubStocksV2 dyeSubStock, int currentWeek)
        {
            List<int> weekList = new List<int>();
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK1) ? 0 : Convert.ToInt32(dyeSubStock.WK1));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK2) ? 0 : Convert.ToInt32(dyeSubStock.WK2));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK3) ? 0 : Convert.ToInt32(dyeSubStock.WK3));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK4) ? 0 : Convert.ToInt32(dyeSubStock.WK4));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK5) ? 0 : Convert.ToInt32(dyeSubStock.WK5));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK6) ? 0 : Convert.ToInt32(dyeSubStock.WK6));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK7) ? 0 : Convert.ToInt32(dyeSubStock.WK7));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK8) ? 0 : Convert.ToInt32(dyeSubStock.WK8));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK9) ? 0 : Convert.ToInt32(dyeSubStock.WK9));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK10) ? 0 : Convert.ToInt32(dyeSubStock.WK10));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK11) ? 0 : Convert.ToInt32(dyeSubStock.WK11));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK12) ? 0 : Convert.ToInt32(dyeSubStock.WK12));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK13) ? 0 : Convert.ToInt32(dyeSubStock.WK13));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK14) ? 0 : Convert.ToInt32(dyeSubStock.WK14));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK15) ? 0 : Convert.ToInt32(dyeSubStock.WK15));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK16) ? 0 : Convert.ToInt32(dyeSubStock.WK16));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK17) ? 0 : Convert.ToInt32(dyeSubStock.WK17));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK18) ? 0 : Convert.ToInt32(dyeSubStock.WK18));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK19) ? 0 : Convert.ToInt32(dyeSubStock.WK19));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK20) ? 0 : Convert.ToInt32(dyeSubStock.WK20));

            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK21) ? 0 : Convert.ToInt32(dyeSubStock.WK21));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK22) ? 0 : Convert.ToInt32(dyeSubStock.WK22));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK23) ? 0 : Convert.ToInt32(dyeSubStock.WK23));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK24) ? 0 : Convert.ToInt32(dyeSubStock.WK24));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK25) ? 0 : Convert.ToInt32(dyeSubStock.WK25));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK26) ? 0 : Convert.ToInt32(dyeSubStock.WK26));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK27) ? 0 : Convert.ToInt32(dyeSubStock.WK27));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK28) ? 0 : Convert.ToInt32(dyeSubStock.WK28));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK29) ? 0 : Convert.ToInt32(dyeSubStock.WK29));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK30) ? 0 : Convert.ToInt32(dyeSubStock.WK30));

            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK31) ? 0 : Convert.ToInt32(dyeSubStock.WK31));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK32) ? 0 : Convert.ToInt32(dyeSubStock.WK32));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK33) ? 0 : Convert.ToInt32(dyeSubStock.WK33));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK34) ? 0 : Convert.ToInt32(dyeSubStock.WK34));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK35) ? 0 : Convert.ToInt32(dyeSubStock.WK35));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK36) ? 0 : Convert.ToInt32(dyeSubStock.WK36));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK37) ? 0 : Convert.ToInt32(dyeSubStock.WK37));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK38) ? 0 : Convert.ToInt32(dyeSubStock.WK38));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK39) ? 0 : Convert.ToInt32(dyeSubStock.WK39));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK40) ? 0 : Convert.ToInt32(dyeSubStock.WK40));

            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK41) ? 0 : Convert.ToInt32(dyeSubStock.WK41));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK42) ? 0 : Convert.ToInt32(dyeSubStock.WK42));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK43) ? 0 : Convert.ToInt32(dyeSubStock.WK43));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK44) ? 0 : Convert.ToInt32(dyeSubStock.WK44));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK45) ? 0 : Convert.ToInt32(dyeSubStock.WK45));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK46) ? 0 : Convert.ToInt32(dyeSubStock.WK46));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK47) ? 0 : Convert.ToInt32(dyeSubStock.WK47));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK48) ? 0 : Convert.ToInt32(dyeSubStock.WK48));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK49) ? 0 : Convert.ToInt32(dyeSubStock.WK49));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK50) ? 0 : Convert.ToInt32(dyeSubStock.WK50));

            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK51) ? 0 : Convert.ToInt32(dyeSubStock.WK51));
            weekList.Add(string.IsNullOrEmpty(dyeSubStock.WK52) ? 0 : Convert.ToInt32(dyeSubStock.WK52));


            var max = weekList.Max();

            dyeSubStock.Yearly_Total = Convert.ToString(weekList.Sum());
            dyeSubStock.Highest_Week = Convert.ToString(max);

            decimal sum_4_Weeks_by_4 = 1;

            List<int> last4WeekData = new List<int>();

            for (int i = weekList.Count - 1; i > 1; i--)
            {
                if (i + 1 <= currentWeek)
                {
                    last4WeekData.Add(weekList[i]);
                    if (last4WeekData.Count == 4)
                        break;
                }

            }

            if (last4WeekData.Count > 0 && last4WeekData.Sum() > 0 && !string.IsNullOrEmpty(dyeSubStock.Weeks_Limit_Req_) && !string.IsNullOrEmpty(dyeSubStock.Live_Stock))
            {
                sum_4_Weeks_by_4 = last4WeekData.Sum() / last4WeekData.Count;

                dyeSubStock.Weeks_Left = Convert.ToString(Convert.ToInt32(Convert.ToInt32(dyeSubStock.Live_Stock) / sum_4_Weeks_by_4));

                if ((Convert.ToInt32(dyeSubStock.Weeks_Limit_Req_) * sum_4_Weeks_by_4) -
                    Convert.ToInt32(dyeSubStock.Live_Stock) > 0)
                {
                    dyeSubStock.LiveStockCellRed = true;
                }
                else
                {
                    dyeSubStock.LiveStockCellRed = false;
                }
            }


            var valueofStock = 0m;

            if (!string.IsNullOrEmpty(dyeSubStock.Unit_Cost))
            {
                var unitCostPrice = dyeSubStock.Unit_Cost.Replace("£", "");

                decimal espStock = 0;
                decimal unitCostPriceVal = 0;


                decimal.TryParse(dyeSubStock.ESP_Stock, out espStock);
                decimal.TryParse(unitCostPrice, out unitCostPriceVal);

                valueofStock = Math.Round((espStock * unitCostPriceVal), 1);
                dyeSubStock.Value_of_Stock_in_House = Convert.ToString(valueofStock);

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
