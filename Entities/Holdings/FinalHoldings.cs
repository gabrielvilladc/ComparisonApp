using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComparisonApp.Entities
{
    public class FinalHoldings
    {
        public string Source { get; set; }
        public string CUSIP { get; set; }
        public string SecurityType { get; set; }
        public string Coupon { get; set; }
        public string MaturityDate { get; set; }
        public string MarketValue { get; set; }
        public string SAndP { get; set; }
        public string Moodys { get; set; }
        public string ClosingPrice { get; set; }
        public string Paramount { get; set; }
        public string YTW { get; set; }
        public string Collateral { get; set; }
        public string TsySpread { get; set; }
        public bool CouponDifference { get; set; }
        public bool MaturityDateDifference { get; set; }
        public bool MarketValueDifference { get; set; }
        public bool SAndPDifference { get; set; }
        public bool MoodysDifference { get; set; }
        public bool ClosingPriceDifference { get; set; }
        public bool ParamountDifference { get; set; }
        public bool YTWDifference { get; set; }
        public bool CollateralDifference { get; set; }
        public bool TsySpreadDifference { get; set; }
    }
}
