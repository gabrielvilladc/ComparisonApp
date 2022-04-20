using System;

namespace ComparisonApp.Entities.Performance
{
    internal class PerformanceAndCellData
    {
        //GetPerformanceReport
        public string ProductID { get; set; }
        public DateTime PeriodDate { get; set; }
        public long PeriodID { get; set; }
        public string AdhocFlag { get; set; }
        public string BenchmarkFlag { get; set; }
        public double? MTD { get; set; }
        public double? QTD { get; set; }
        public double? YTD { get; set; }
        public double? PriorYear { get; set; }
        public double? PriorTwoYear { get; set; }
        public double? PriorThreeYear { get; set; }
        public double? PriorFourYear { get; set; }
        public double? PriorFiveYear { get; set; }
        public double? PriorSevenYear { get; set; }
        public double? PriorTenYear { get; set; }
        public double? SinceInceptionReturn { get; set; }
        public DateTime? InceptionDate { get; set; }

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
        public double? MarketValue { get; set; }
        public double? Duration { get; set; }
        public double? MDuration { get; set; }
        public double? Term { get; set; }
        public double? YTW { get; set; }
        public double? CurrentYield { get; set; }
        public double? Coupon { get; set; }
        public double? Price { get; set; }
        public double? Accrued { get; set; }
        public DateTime? MaturityDate { get; set; }
        public double? APR { get; set; }
        public double? Convexity { get; set; }
        public double? PerformanceMarketWeight { get; set; }
        public double? IndexLevel { get; set; }
        public double? DailyReturn { get; set; }
        public double? SIDReturn { get; set; }
        public int? CellLabelOrder { get; set; }
        public bool? Benchmark { get; set; }
    }
}
