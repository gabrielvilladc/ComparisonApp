using System;

namespace ComparisonApp.Entities.PortfolioStats
{
    public class FinalPortfolioStat
    {
        public string Source { get; set; }
        public string SecurityType { get; set; }
        public string MarketValue { get; set; }
        public string Coupon { get; set; }
        public string BenchmarkCoupon { get; set; }
        public string Weight { get; set; }
        public string BenchmarkWeight { get; set; }
        public string YTW { get; set; }
        public string BenchmarkYTW { get; set; }
        public string ModifiedDuration { get; set; }
        public string BenchmarkModifiedDuration { get; set; }
        public string Convexity { get; set; }
        public string BenchmarkConvexity { get; set; }
        public string DailyReturn { get; set; }
        public string BenchmarkDailyReturn { get; set; }
        public string MTD { get; set; }
        public string BenchmarkMTD { get; set; }
        public string QTD { get; set; }
        public string BenchmarkQTD { get; set; }
        public string YTD { get; set; }
        public string BenchmarkYTD { get; set; }
        public string CummReturn { get; set; }
        public string BenchmarkCummReturn { get; set; }

        //Validations
        public bool SecurityTypeDifference { get; set; }
        public bool MarketValueDifference { get; set; }
        public bool CouponDifference { get; set; }
        public bool BenchmarkCouponDifference { get; set; }
        public bool WeightDifference { get; set; }
        public bool BenchmarkWeightDifference { get; set; }
        public bool YTWDifference { get; set; }
        public bool BenchmarkYTWDifference { get; set; }
        public bool ModifiedDurationDifference { get; set; }
        public bool BenchmarkModifiedDurationDifference { get; set; }
        public bool ConvexityDifference { get; set; }
        public bool BenchmarkConvexityDifference { get; set; }
        public bool DailyReturnDifference { get; set; }
        public bool BenchmarkDailyReturnDifference { get; set; }
        public bool MTDDifference { get; set; }
        public bool BenchmarkMTDDifference { get; set; }
        public bool QTDDifference { get; set; }
        public bool BenchmarkQTDDifference { get; set; }
        public bool YTDDifference { get; set; }
        public bool BenchmarkYTDDifference { get; set; }
        public bool CummReturnDifference { get; set; }
        public bool BenchmarkCummReturnDifference { get; set; }
    }
}
