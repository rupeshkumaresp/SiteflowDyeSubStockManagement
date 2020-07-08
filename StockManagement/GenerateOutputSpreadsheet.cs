using nsDyeSubStockManagement.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace nsDyeSubStockManagement
{


    /// <summary>
    /// Generate the Dye Sub Stock Spreadsheet
    /// </summary>
    public class GenerateOutputSpreadsheet
    {

        public ExcelPackage package = new ExcelPackage();
        public ExcelWorksheet worksheet;
        string pathServerReport = @"\\web2print\Sites\SiteFlow\DyeSub\Reports";
        string UserName = "graphite.rack";
        string Password = "123Grand?";
        string domain = "172.26.128.158";
        SiteFlowEntities ctx = new SiteFlowEntities();


        public void CreateSpreadSheet()
        {
            worksheet = package.Workbook.Worksheets.Add("Stocks");
            int rowJump = 1;
            AddMainHeaderRow();
            rowJump++;
            var dyeSubStock = ctx.tDyeSubStocksV2.ToList();

            foreach (var stock in dyeSubStock)
            {

                AddStockRow(rowJump, stock);
                rowJump++;
            }

            worksheet.Column(1).Width = 20;
            worksheet.Column(2).Width = 20;
            worksheet.Column(3).Width = 20;
            worksheet.Column(4).Width = 40;
            worksheet.Column(5).Width = 20;
            worksheet.Column(6).Width = 20;
            worksheet.Column(7).Width = 20;
            worksheet.Column(8).Width = 15;
            worksheet.Column(9).Width = 15;
            worksheet.Column(10).Width = 15;
            worksheet.Column(11).Width = 15;
            worksheet.Column(12).Width = 15;
            worksheet.Column(13).Width = 15;
            worksheet.Column(14).Width = 15;
            worksheet.Column(15).Width = 15;
            worksheet.Column(16).Width = 15;

            for (int i = 0; i <= 52; i++)
            {
                worksheet.Column(15 + i).Width = 15;
            }


            worksheet.Column(69).Width = 15;
            worksheet.Column(70).Width = 15;
            worksheet.Column(71).Width = 30;
            worksheet.Column(72).Width = 15;
            worksheet.Column(73).Width = 15;

            worksheet.Column(74).Width = 15;
            worksheet.Column(75).Width = 15;

            worksheet.Column(76).Width = 15;
            worksheet.Column(77).Width = 15;

            worksheet.Column(78).Width = 15;
            worksheet.Column(79).Width = 15;

            worksheet.Column(80).Width = 15;
            worksheet.Column(81).Width = 15;

            worksheet.Column(82).Width = 15;
            worksheet.Column(83).Width = 15;
            worksheet.Column(84).Width = 15;

            //worksheet.Column(12).Style.Locked = false;
            //worksheet.Column(13).Style.Locked = false;
            //worksheet.Column(14).Style.Locked = false;

            worksheet.Protection.IsProtected = true;

            worksheet.View.FreezePanes(2, 1);

            var name = "Dye Sub Stock Report" + System.DateTime.Now.ToString("dd-MM-yyyy HH_mm_ss");


            SaveReportAndDisplay(name);

        }

        private void SaveReportAndDisplay(string name)
        {
            // Save file and return stream
            var fileName = Path.GetTempFileName();
            package.SaveAs(new FileInfo(fileName));


            var currentDirectory = Environment.CurrentDirectory;
            if (!Directory.Exists(currentDirectory + @"\" + "Reports"))
            {
                Directory.CreateDirectory(currentDirectory + @"\" + "Reports");
            }

            var path = currentDirectory + @"\" + @"Reports\" + name + ".xlsx";
            SaveStreamToFile(path, new FileStream(fileName, FileMode.Open));

            package.Dispose();



            //using (new NetworkConnection(pathServerReport, new NetworkCredential(UserName, Password, domain)))
            //{
            //    File.Copy(path, pathServerReport + @"\\" + name + ".xlsx");
            //}

            System.Diagnostics.Process.Start(path);
        }

        private void AddStockRow(int rowJump, tDyeSubStocksV2 stock)
        {
            //add a row

            int column = 1;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Stock_Name);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                           OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Stock_Type);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            if(stock.Stock_Type=="Component")
                worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#c4aded"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Stock_Category);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Substrate_Name);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Extra);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;


            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Sizes);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Colours);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Weeks_Limit_Req_);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Weeks_Left);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.ESP_Stock);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            column++;


            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Live_Stock);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            if (Convert.ToBoolean(stock.LiveStockCellRed))
            {
                worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
            }

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Spoilage);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.DOA);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = "";
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            column++;

            worksheet.Cells[rowJump, column].Value = "";
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            column++;

            worksheet.Cells[rowJump, column].Value = "";
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Highest_Week);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK1);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK2);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK3);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK4);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK5);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK6);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK7);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK8);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK9);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK10);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK11);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK12);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK13);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK14);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK15);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK16);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK17);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK18);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK19);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK20);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;


            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK21);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK22);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK23);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK24);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK25);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK26);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK27);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK28);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK29);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK30);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK31);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK32);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK33);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK34);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK35);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK36);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK37);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK38);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK39);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK40);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK41);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK42);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK43);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK44);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK45);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK46);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK47);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK48);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK49);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK50);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;
            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK51);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.WK52);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;


            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.YEAR);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Yearly_Total);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Stock_Available_External);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.External_Supplier);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Unit_Cost);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Value_of_Stock_in_House);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#e1d7f7"));

            column++;


            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_1);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_1_Qty);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_2);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_2_Qty);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_3);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_3_Qty);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;


            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_4);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_4_Qty);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_5);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;

            worksheet.Cells[rowJump, column].Value = Convert.ToString(stock.Component_5_Qty);
            worksheet.Cells[rowJump, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowJump, column].Style.Border.BorderAround(
                OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            worksheet.Cells[rowJump, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[rowJump, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#d8f1fa"));

            column++;


        }

        private void AddMainHeaderRow()
        {
            int rowJump = 1;
            // Set up columns
            var headerColumns = new Dictionary<string, int>();

            int icount = 1;

            headerColumns.Add("Stock Name", icount);
            icount++;

            headerColumns.Add("Stock Type", icount);
            icount++;

            headerColumns.Add("Stock Category", icount);
            icount++;

            headerColumns.Add("Substrate Name", icount);
            icount++;

            headerColumns.Add("Extra", icount);
            icount++;

            headerColumns.Add("Sizes", icount);
            icount++;

            headerColumns.Add("Colours", icount);
            icount++;

            headerColumns.Add("Weeks Limit Req.", icount);
            icount++;

            headerColumns.Add("Weeks Left", icount);
            icount++;

            headerColumns.Add("ESP Stock", icount);
            icount++;
          
            headerColumns.Add("Live Stock", icount);
            icount++;

            headerColumns.Add("Spoilage", icount);
            icount++;

            headerColumns.Add("DOA", icount);
            icount++;

            headerColumns.Add("Add to ESP Stock", icount);
            icount++;

            headerColumns.Add("Add to Spoilage", icount);
            icount++;

            headerColumns.Add("Add to DOA", icount);
            icount++;

            headerColumns.Add("Highest Week", icount);
            icount++;

            for (int i = 1; i <= 52; i++)
            {
                headerColumns.Add("WK" + i.ToString(), icount);
                icount++;
            }

            headerColumns.Add("YEAR", icount);
            icount++;

            headerColumns.Add("Yearly Total", icount);
            icount++;

            headerColumns.Add("Stock Available External", icount);
            icount++;

            headerColumns.Add("External Supplier", icount);
            icount++;

            headerColumns.Add("Unit Cost", icount);
            icount++;

            headerColumns.Add("Value of Stock in House", icount);
            icount++;

            headerColumns.Add("Component 1", icount);
            icount++;

            headerColumns.Add("Component 1 Qty", icount);
            icount++;

            headerColumns.Add("Component 2", icount);
            icount++;

            headerColumns.Add("Component 2 Qty", icount);
            icount++;

            headerColumns.Add("Component 3", icount);
            icount++;

            headerColumns.Add("Component 3 Qty", icount);
            icount++;

            headerColumns.Add("Component 4", icount);
            icount++;

            headerColumns.Add("Component 4 Qty", icount);
            icount++;

            headerColumns.Add("Component 5", icount);
            icount++;

            headerColumns.Add("Component 5 Qty", icount);
            icount++;


            // Write column headers
            foreach (var colKvp in headerColumns)
            {
                if (colKvp.Value > 0)
                {
                    worksheet.Cells[rowJump, colKvp.Value].Value = colKvp.Key;
                    worksheet.Cells[rowJump, colKvp.Value].Style.HorizontalAlignment =
                        OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells[rowJump, colKvp.Value].Style.VerticalAlignment =
                        OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheet.Cells[rowJump, colKvp.Value].Style.Border.BorderAround(
                           OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    worksheet.Cells[rowJump, colKvp.Value].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[rowJump, colKvp.Value].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    worksheet.Cells[rowJump, colKvp.Value].Style.Font.Bold = true;
                    worksheet.Cells[rowJump, colKvp.Value].Style.Font.Size = 12;
                    worksheet.Cells[rowJump, colKvp.Value].Style.WrapText = true;
                }
            }
        }

        public static void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            if (stream.Length == 0) return;

            // Create a FileStream object to write a stream to a file
            using (FileStream fileStream = File.Create(fileFullPath, (int)stream.Length))
            {
                // Fill the bytes[] array with the stream data
                var bytesInStream = new byte[stream.Length];
                stream.Read(bytesInStream, 0, (int)bytesInStream.Length);

                // Use FileStream object to write to the specified file
                fileStream.Write(bytesInStream, 0, bytesInStream.Length);
            }
        }

    }
}
