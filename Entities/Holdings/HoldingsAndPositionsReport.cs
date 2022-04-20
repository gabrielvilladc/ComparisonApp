using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComparisonApp.Entities
{
    internal class HoldingsAndPositionsReport
    {
        public string ProductID { get; set; }
        public DateTime PeriodDate { get; set; }
        public string AdhocFlag { get; set; }
        public string CUSIP { get; set; }
        public string SecurityType { get; set; }
        public double? Quantity { get; set; }
        public string Description { get; set; }
        public double? Coupon { get; set; }
        public DateTime? MaturityDate { get; set; }
        public DateTime? Maturity { get; set; }
        public double? MarketValue { get; set; }
        public int? EffectiveDuration { get; set; }
        public double? Convexity { get; set; }
        public string SAndP { get; set; }
        public string Moodys { get; set; }
        public string InternalRating { get; set; }
        public double? ClosingPrice { get; set; }
        public double? Paramount { get; set; }
        public double? YTW { get; set; }
        public string Collateral { get; set; }
        public int? TsySpread { get; set; }

        //GetPositionsReport
        public string CellGroup { get; set; }
        public string CellLabel { get; set; }
        public string CellName { get; set; }
        public int? CellLabelOrder {get;set;}
        public string Cusip { get; set; }
        public double? Price { get; set; }
        public double? ParAmount { get; set; }
        public string CollateralType { get; set; }
        public string RatingSP { get; set; }
        public string RatingMoodys { get; set; }
        public double? MDuration { get; set; }

    }
}
