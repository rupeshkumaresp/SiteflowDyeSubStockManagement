//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace nsDyeSubStockManagement.Model
{
    using System;
    using System.Collections.Generic;
    
    public partial class tSiteFlowInputData
    {
        public long ID { get; set; }
        public string SourceOrderId { get; set; }
        public string SKU { get; set; }
        public string ComponentsBarcode { get; set; }
        public Nullable<int> Quantity { get; set; }
        public string Components0Substrate { get; set; }
        public string Components1Substrate { get; set; }
        public string Components0SizeForImpo { get; set; }
        public string Components1SizeForImpo { get; set; }
        public string Components0ProductFinishedPageSize { get; set; }
        public string Components1ProductFinishedPageSize { get; set; }
        public Nullable<int> Components0Pages { get; set; }
        public Nullable<int> Components1Pages { get; set; }
        public string Components0CoverType { get; set; }
        public string Components1CoverType { get; set; }
        public string Components0StockCoverType { get; set; }
        public string Components1StockCoverType { get; set; }
        public string Components0Extra { get; set; }
        public string Components1Extra { get; set; }
        public string ComponentsColour { get; set; }
        public string ComponentsRibbon { get; set; }
        public string Components0Country { get; set; }
        public string Components1Country { get; set; }
        public string Components0ArtworkUrl { get; set; }
        public string Components1ArtworkUrl { get; set; }
        public string Account { get; set; }
        public string OrderStatus { get; set; }
        public Nullable<System.DateTime> OrderDateTime { get; set; }
        public Nullable<System.DateTime> EmailProcessedDateTime { get; set; }
        public Nullable<System.DateTime> ShippedDate { get; set; }
        public Nullable<bool> PDFDownloaded { get; set; }
        public Nullable<bool> IsValidArtwork { get; set; }
        public Nullable<bool> DiscardDownload { get; set; }
        public Nullable<int> InvalidUrlTryCount { get; set; }
        public Nullable<bool> PDFMerged { get; set; }
        public Nullable<System.DateTime> MergeProcessingDateTime { get; set; }
        public string CalculatedPDFSize { get; set; }
        public Nullable<int> PDFPageCount { get; set; }
        public string DownloadedArtworkPDF { get; set; }
        public string JsonData { get; set; }
        public Nullable<bool> InvalidEmailSent { get; set; }
        public string RUSH { get; set; }
        public Nullable<System.DateTime> SLADate { get; set; }
        public Nullable<System.DateTime> PrintReadyDate { get; set; }
    }
}