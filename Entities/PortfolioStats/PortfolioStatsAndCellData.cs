using System;

namespace ComparisonApp.Entities.PortfolioStats
{
    internal class PortfolioStatsAndCellData
    {
        //GetPortfolioStatsReport                            
        public string ProductID { get; set; }
        public DateTime PeriodDate { get; set; }
        public long PeriodID { get; set; }
        public string AdhocFlag { get; set; }
        public string SecurityType { get; set; }
        public double? MarketValue { get;set; }
        public double? Coupon { get; set; }
        public double? BenchmarkCoupon { get; set; }
        public string BenchmarkFlag { get; set; }
        public double? Weight { get; set; }
        public double? BenchmarkWeight { get; set; }
        public double? BenchmarkYTW { get; set; }
        public double? ModifiedDuration { get; set; }
        public double? BenchmarkModifiedDuration { get; set; }
        public double? Convexity { get; set; }
        public double? BenchmarkConvexity { get; set; }
        public double? DailyReturn { get; set; }
        public double? BenchmarkDailyReturn { get; set; }
        public double? MTD { get; set; }
        public double? BenchmarkMTD { get; set; }
        public double? QTD { get; set; }
        public double? BenchmarkQTD { get; set; }
        public double? YTD { get; set; }
        public double? BenchmarkYTD { get; set; }
        public double? CummReturn { get; set; }
        public double? BenchmarkCummReturn { get; set; }

        //GetCellData
        public string CellLabel { get; set; }
        public string CellGroup { get; set; }
        public string CellName { get; set; }
        public string PortfolioDescription { get; set; }
        public DateTime? AsOfDate { get; set; }
        public double? FeatureMarketWeight { get; set; }
        public int? NumIssues { get; set; }
        public double? ParValue { get; set; }
        public double? BookValue { get; set; }
        public double? Duration { get; set; }
        public double? MDuration { get; set; }
        public double? Term { get; set; }
        public double? YTW { get; set; }
        public double? CurrentYield { get; set; }
        public double? Price { get; set; }
        public double? Accrued { get; set; }
        public DateTime? MaturityDate { get; set; }
        public double? APR { get; set; }
        public double? PerformanceMarketWeight { get; set; }
        public double? IndexLevel { get; set; }
        public double? SIDReturn { get; set; }
        public int? CellLabelOrder { get; set; }
        public bool? Benchmark { get; set; }
    }
}
