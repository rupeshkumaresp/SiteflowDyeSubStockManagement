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
    
    public partial class tDyeSubStock
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public tDyeSubStock()
        {
            this.tDyeSubConsumption = new HashSet<tDyeSubConsumption>();
        }
    
        public int ID { get; set; }
        public string PartNumber { get; set; }
        public string Description { get; set; }
        public string SubstrateName { get; set; }
        public string ESPSize { get; set; }
        public string Colour { get; set; }
        public string CATSStockName { get; set; }
        public Nullable<decimal> Price_ { get; set; }
        public Nullable<decimal> PriceGBP { get; set; }
        public Nullable<int> QuantityAvailable { get; set; }
        public Nullable<int> AddToQuantitiy { get; set; }
        public Nullable<int> Spoilage { get; set; }
        public Nullable<int> DOA { get; set; }
        public Nullable<int> TotalConsumed { get; set; }
        public Nullable<int> Jan { get; set; }
        public Nullable<int> Feb { get; set; }
        public Nullable<int> March { get; set; }
        public Nullable<int> April { get; set; }
        public Nullable<int> May { get; set; }
        public Nullable<int> June { get; set; }
        public Nullable<int> July { get; set; }
        public Nullable<int> August { get; set; }
        public Nullable<int> Sept { get; set; }
        public Nullable<int> Oct { get; set; }
        public Nullable<int> Nov { get; set; }
        public Nullable<int> Dec { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<tDyeSubConsumption> tDyeSubConsumption { get; set; }
    }
}
