using ComparisonApp.Entities.PortfolioStats;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ComparisonApp.Functions
{
    public class PortfolioStatLogic
    {
        public List<FinalPortfolioStatsComparison> GetPortfoliostatsDifferences(IServiceProvider services, string[] args)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            OperationLogger logger = provider.GetRequiredService<OperationLogger>();
            string perdioId = args[0];
            string periodDate = args[1];
            string clientName = args[2];
            string cellGroup = args[3];
            string productCode = logger.GetClientProductCode(clientName);


            //**********           PORTFOLIOSTATS LIST, from SLCCORICDB (CORICAPP)        ******************
            //Getting the data from SP into a Data Table
            DataTable dataTable = logger.GetPortfolioStatsReport(perdioId);
            //Converting DataTable to List
            List<PortfolioStatsAndCellData> temporalList = new List<PortfolioStatsAndCellData>();
            temporalList = GeneralLogic.ConvertDataTable<PortfolioStatsAndCellData>(dataTable);
            //Filttering the list by Client
            temporalList = temporalList.Where(x => x.ProductID == productCode).ToList();
            //Mapping the records to a standard class list
            List<FinalPortfolioStat> _finalList = new List<FinalPortfolioStat>();
            _finalList = logger.MappPortfolioStats(temporalList);

            //**********           PORTFOLIOSTATS LIST, from Development DB (GetCellData SP)        ******************
            //Getting the data from SP into a Data Table
            DataTable dataTable2 = logger.GetCellData(periodDate, clientName, cellGroup);
            //Converting DataTable to List
            List<PortfolioStatsAndCellData> temporalList2 = new List<PortfolioStatsAndCellData>();
            temporalList2 = GeneralLogic.ConvertDataTable<PortfolioStatsAndCellData>(dataTable2);
            List<PortfolioStatsAndCellData> temporalList2Assets = temporalList2.Where(x => x.Benchmark == false).ToList();
            List<PortfolioStatsAndCellData> temporalList2Benchmarks = temporalList2.Where(x => x.Benchmark == true).ToList();
            foreach (var asset in temporalList2Assets)
            {
                PortfolioStatsAndCellData benchmarkByCellLabel = new PortfolioStatsAndCellData();
                benchmarkByCellLabel = temporalList2Benchmarks.Where(x => x.CellLabel.Equals(asset.CellLabel)).FirstOrDefault();
                asset.BenchmarkCoupon = benchmarkByCellLabel.Coupon;
                asset.BenchmarkWeight = benchmarkByCellLabel.FeatureMarketWeight;
                asset.BenchmarkYTW = benchmarkByCellLabel.YTW;
                asset.BenchmarkModifiedDuration = benchmarkByCellLabel.MDuration;
                asset.BenchmarkConvexity = benchmarkByCellLabel.Convexity;
                asset.BenchmarkDailyReturn= benchmarkByCellLabel.DailyReturn;
                asset.BenchmarkMTD = benchmarkByCellLabel.MTD;
                asset.BenchmarkQTD = benchmarkByCellLabel.QTD;
                asset.BenchmarkYTD = benchmarkByCellLabel.YTD;
                asset.BenchmarkCummReturn = benchmarkByCellLabel.SIDReturn;
            }
            
            //Mapping the records to a standard class list
            List<FinalPortfolioStat> _finalList2 = new List<FinalPortfolioStat>();
            _finalList2 = logger.MappPortfolioStats(temporalList2Assets);

            List<FinalPortfolioStatsComparison> _listFinalDifferences = new List<FinalPortfolioStatsComparison>();
            foreach (var item in _finalList)
            {
                var comparisonObject = _finalList2.Where(x => x.SecurityType == item.SecurityType).FirstOrDefault();
                if (item.BenchmarkConvexity != comparisonObject.BenchmarkConvexity)
                    item.BenchmarkConvexityDifference = true;
                if (item.BenchmarkCoupon != comparisonObject.BenchmarkCoupon)
                    item.BenchmarkCouponDifference = true;
                if (item.BenchmarkCummReturn != comparisonObject.BenchmarkCummReturn)
                    item.BenchmarkCummReturnDifference = true;
                if (item.BenchmarkDailyReturn != comparisonObject.BenchmarkDailyReturn)
                    item.BenchmarkDailyReturnDifference = true;
                if (item.BenchmarkModifiedDuration != comparisonObject.BenchmarkModifiedDuration)
                    item.BenchmarkModifiedDurationDifference = true;
                if (item.BenchmarkMTD != comparisonObject.BenchmarkMTD)
                    item.BenchmarkMTDDifference = true;
                if (item.BenchmarkQTD != comparisonObject.BenchmarkQTD)
                    item.BenchmarkQTDDifference = true;
                if (item.BenchmarkWeight != comparisonObject.BenchmarkWeight)
                    item.BenchmarkWeightDifference = true;
                if (item.BenchmarkYTD != comparisonObject.BenchmarkYTD)
                    item.BenchmarkYTDDifference = true;
                if (item.BenchmarkYTW != comparisonObject.BenchmarkYTW)
                    item.BenchmarkYTWDifference = true;
                if (item.Convexity != comparisonObject.Convexity)
                    item.ConvexityDifference = true;
                if (item.Coupon != comparisonObject.Coupon)
                    item.CouponDifference = true;
                if (item.CummReturn != comparisonObject.CummReturn)
                    item.CummReturnDifference = true;
                if (item.DailyReturn != comparisonObject.DailyReturn)
                    item.DailyReturnDifference = true;
                if (item.MarketValue != comparisonObject.MarketValue)
                    item.MarketValueDifference = true;
                if (item.ModifiedDuration != comparisonObject.ModifiedDuration)
                    item.ModifiedDurationDifference = true;
                if (item.MTD != comparisonObject.MTD)
                    item.MTDDifference = true;
                if (item.QTD != comparisonObject.QTD)
                    item.QTDDifference = true;
                if (item.SecurityType != comparisonObject.SecurityType)
                    item.SecurityTypeDifference = true;
                if (item.Weight != comparisonObject.Weight)
                    item.WeightDifference = true;
                if (item.YTD != comparisonObject.YTD)
                    item.YTDDifference = true;
                if (item.YTW != comparisonObject.YTW)
                    item.YTWDifference = true;

                if  (item.SecurityTypeDifference || 
                    item.MarketValueDifference || 
                    item.CouponDifference || 
                    item.BenchmarkCouponDifference || 
                    item.WeightDifference || 
                    item.BenchmarkWeightDifference || 
                    item.YTWDifference || 
                    item.BenchmarkYTWDifference || 
                    item.ModifiedDurationDifference || 
                    item.BenchmarkModifiedDurationDifference || 
                    item.ConvexityDifference || 
                    item.BenchmarkConvexityDifference || 
                    item.DailyReturnDifference || 
                    item.BenchmarkDailyReturnDifference || 
                    item.MTDDifference || 
                    item.BenchmarkMTDDifference || 
                    item.QTDDifference || 
                    item.BenchmarkQTDDifference || 
                    item.YTDDifference || 
                    item.BenchmarkYTDDifference || 
                    item.CummReturnDifference ||
                    item.BenchmarkCummReturnDifference)
                {
                    FinalPortfolioStatsComparison finalComparisson = new FinalPortfolioStatsComparison()
                    {
                        PortfolioStatsReport = item,
                        CellDataSP = comparisonObject
                    };
                    _listFinalDifferences.Add(finalComparisson);
                }
                
            }

            return _listFinalDifferences;
        }
    }
}
