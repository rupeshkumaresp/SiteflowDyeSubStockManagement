using nsDyeSubStockManagement.Model;
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
                this.lblStatusDownload.BeginInvoke((MethodInvoker)delegate() { this.lblStatusDownload.Text = "Report Generation is in progress, please wait...."; ; ;});
            }

            if (this.btnDownload.InvokeRequired)
            {
                this.btnDownload.BeginInvoke((MethodInvoker)delegate() { this.btnDownload.Enabled = false; ; ;});
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
                    this.lblStatusDownload.BeginInvoke((MethodInvoker)delegate() { this.lblStatusDownload.Text = "Report Generation failed: " + ex.Message; ; ;});
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate() { this.lblStatusDownload.ForeColor = Color.Green; ; ;});
                }

                if (this.btnDownload.InvokeRequired)
                {
                    this.btnDownload.BeginInvoke((MethodInvoker)delegate() { this.btnDownload.Enabled = true; ; ;});
                }
                return;
            }
            if (!exception)
            {
                if (this.lblStatusDownload.InvokeRequired)
                {
                    this.lblStatusDownload.BeginInvoke((MethodInvoker)delegate() { this.lblStatusDownload.Text = "Report Generation is complete, launching report xlsx..."; ; ;});
                }

                if (this.btnDownload.InvokeRequired)
                {
                    this.btnDownload.BeginInvoke((MethodInvoker)delegate() { this.btnDownload.Enabled = true; ; ;});
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
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate() { this.lblStatus.Text = "Upload is in progress, please wait..."; ; ;});
                }

                var data = ImportSpreadsheet.Import(_inputStream);

                foreach (var row in data)
                {
                    var partNumber = row["Part Number".ToLower()];
                    var description = row["Description".ToLower()];
                    var substrateName = row["Substrate Name".ToLower()];
                    var espSize = row["ESP Size mm".ToLower()];
                    var Colour = row["Colour".ToLower()];
                    var CATSStockName = row["CATS Stock name".ToLower()];
                    var Price = row["Price in $".ToLower()];
                    //var QuantityAvailable = row["Quantity Available".ToLower()];
                    var AddToQuantity = row["Add to Quantity".ToLower()];
                    var AddToSpoilage = row["Add to Spoilage".ToLower()];
                    var AddToDOA = row["Add to DOA".ToLower()];
                    //var Spoilage = row["Spoilage".ToLower()];
                    //var DOA = row["DOA".ToLower()];
                    //var TotalConsumed = row["Total Consumed".ToLower()];

                    var stock = ctx.tDyeSubStock.Where(s => s.PartNumber == partNumber && s.SubstrateName == substrateName && s.ESPSize == espSize && s.Colour == Colour).FirstOrDefault();

                    if (stock == null)
                    {
                        if (!string.IsNullOrEmpty(partNumber) && !string.IsNullOrEmpty(description) && !string.IsNullOrEmpty(substrateName))
                        {
                            stock = new tDyeSubStock();
                            stock.PartNumber = partNumber;
                            stock.Description = description;
                            stock.SubstrateName = substrateName;
                            stock.ESPSize = espSize;
                            stock.Colour = Colour;
                            stock.CATSStockName = CATSStockName;
                            stock.Price_ = string.IsNullOrEmpty(Price) ? 0 : Convert.ToDecimal(Price);
                            ctx.tDyeSubStock.Add(stock);
                            ctx.SaveChanges();
                        }
                    }

                    if (stock == null)
                        continue;

                    bool updated = false;
                    if (!string.IsNullOrEmpty(AddToQuantity))
                    {
                        int addToQTy = 0;

                        bool conversion = int.TryParse(AddToQuantity, out addToQTy);

                        if (conversion)
                        {
                            var QA = Convert.ToInt32(stock.QuantityAvailable);

                            stock.QuantityAvailable = QA + Convert.ToInt32(addToQTy);

                            //if (addToQTy < 0)
                            //{
                            //    stock.QuantityAvailable = QA + Convert.ToInt32((addToQTy));
                            //}
                            //else
                            //{                            
                            //    stock.QuantityAvailable = QA + Convert.ToInt32((addToQTy * 92.5 / 100));
                            //    stock.DOA = Convert.ToInt32(stock.DOA) + Convert.ToInt32((addToQTy * 7.5 / 100));
                            //}
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
                                stock.Spoilage = Convert.ToInt32(stock.Spoilage) + spoilageQty;
                            }
                            else
                            {
                                stock.Spoilage = Convert.ToInt32(stock.Spoilage) + spoilageQty;
                                stock.QuantityAvailable = (stock.QuantityAvailable != null ? stock.QuantityAvailable - spoilageQty : -spoilageQty);
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
                            stock.DOA = Convert.ToInt32(stock.DOA) + doaQTy;

                            if (doaQTy < 0)
                            {
                                stock.DOA = Convert.ToInt32(stock.DOA) + doaQTy;
                            }
                            else
                            {
                                stock.DOA = Convert.ToInt32(stock.DOA) + doaQTy;
                                stock.QuantityAvailable = (stock.QuantityAvailable != null ? stock.QuantityAvailable - doaQTy : -doaQTy);
                            }
                            updated = true;
                        }

                    }

                    if (updated)
                        ctx.SaveChanges();

                }

                if (this.lblStatus.InvokeRequired)
                {
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate() { this.lblStatus.Text = "Upload completed successfully!"; ; ;});
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate() { this.lblStatus.ForeColor = Color.Green; ; ;});
                }

                if (this.btnUpload.InvokeRequired)
                {
                    this.btnUpload.BeginInvoke((MethodInvoker)delegate() { this.btnUpload.Enabled = true; ; ;});
                }

                if (this.btnBrowse.InvokeRequired)
                {
                    this.btnBrowse.BeginInvoke((MethodInvoker)delegate() { this.btnBrowse.Enabled = true; ; ;});
                }
            }
            catch (Exception ex)
            {

                if (this.lblStatus.InvokeRequired)
                {
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate() { this.lblStatus.Text = "Exception: " + ex.Message; ; ;});
                    this.lblStatus.BeginInvoke((MethodInvoker)delegate() { this.lblStatus.ForeColor = Color.Red; ; ;});
                }

                if (this.btnUpload.InvokeRequired)
                {
                    this.btnUpload.BeginInvoke((MethodInvoker)delegate() { this.btnUpload.Enabled = true; ; ;});
                }

                if (this.btnBrowse.InvokeRequired)
                {
                    this.btnBrowse.BeginInvoke((MethodInvoker)delegate() { this.btnBrowse.Enabled = true; ; ;});
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
