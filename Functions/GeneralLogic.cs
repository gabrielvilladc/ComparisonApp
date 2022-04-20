using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Linq;
using ComparisonApp.Entities;
using ComparisonApp.Entities.Performance;
using ComparisonApp.Entities.PortfolioStats;
using System.IO;

namespace ComparisonApp.Functions
{
    public class GeneralLogic
    {
        public static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }

            return data;
        }

        public static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                    {
                        pro.SetValue(obj, dr[column.ColumnName] == DBNull.Value ? null : dr[column.ColumnName], null);
                        break;
                    }
                    else
                        continue;
                }
            }

            return obj;
        }

        public void GenerateReport(List<FinalHoldingsComparison> _listHoldingsDifferences, List<FinalPerformanceComparison> _listPerformanceDifferences, List<FinalPortfolioStatsComparison> _listPortfolioStatsDifferences, string clientName, string periodDate)
        {
            string filePath = $"C:\\Users\\sp72\\Documents\\CORIC_Files\\ComparisonApp\\Comparison_{clientName}_{periodDate}_{DateTime.Now.ToString("yyyyMMdd")}{DateTime.Now.ToString("HH")}{DateTime.Now.ToString("mm")}{DateTime.Now.ToString("ss")}.xls";

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Set some properties of the Excel document
                //excelPackage.Workbook.Properties.Author = "VDWWD";
                //excelPackage.Workbook.Properties.Title = "Title of Document";
                //excelPackage.Workbook.Properties.Subject = "EPPlus demo export data";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                worksheet.Cells[1, 1].Value = clientName;
                worksheet.Cells[1, 1, 1, _listHoldingsDifferences.FirstOrDefault().Holdings.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Merge = true;
                var someCells = worksheet.Cells["A1"];
                someCells.Style.Font.Bold = true;
                someCells.Style.Font.Color.SetColor(Color.Ivory);
                someCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                someCells.Style.Fill.BackgroundColor.SetColor(Color.Navy);
                someCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                #region Holdings
                int columnNumber = 1;
                int rowNumber = 2;
                worksheet.Cells[rowNumber, 1].Value = "Holdings";
                worksheet.Cells[rowNumber, 1, rowNumber, _listHoldingsDifferences.FirstOrDefault().Holdings.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Merge = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Bold = true;
                //worksheet.Cells[rowNumber, 1].Style.Font.Color.SetColor(Color.Ivory);
                worksheet.Cells[rowNumber, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rowNumber, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                worksheet.Cells[rowNumber, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                rowNumber++;
                foreach (PropertyInfo propertyInfo in _listHoldingsDifferences.FirstOrDefault().Holdings.GetType().GetProperties())
                {
                    if (propertyInfo.PropertyType == typeof(string))
                    {
                        worksheet.Cells[rowNumber, columnNumber].Value = propertyInfo.Name;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[rowNumber, columnNumber].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        worksheet.Cells[rowNumber, columnNumber].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        columnNumber++;
                    }
                }
                int lastRowNumber = 0;
                rowNumber++;
                foreach (var item in _listHoldingsDifferences.Select((value, i) => new { i, value }))
                {
                    lastRowNumber = rowNumber;
                    columnNumber = 1;
                    //Row for holdings
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.Source;
                    worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.CUSIP;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.SecurityType;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.Coupon;
                    if (item.value.Holdings.CouponDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.MaturityDate;
                    if (item.value.Holdings.MaturityDateDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.MarketValue;
                    if (item.value.Holdings.MarketValueDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.SAndP;
                    if (item.value.Holdings.SAndPDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.Moodys;
                    if (item.value.Holdings.MoodysDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.ClosingPrice;
                    if (item.value.Holdings.ClosingPriceDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.Paramount;
                    if (item.value.Holdings.ParamountDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.YTW;
                    if (item.value.Holdings.YTWDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.Collateral;
                    if (item.value.Holdings.CollateralDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Holdings.TsySpread;
                    if (item.value.Holdings.TsySpreadDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    //Next row for GetPositions
                    rowNumber++;
                    columnNumber = 1;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.Source;
                    worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.CUSIP;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.SecurityType;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.Coupon;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.MaturityDate;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.MarketValue;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.SAndP;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.Moodys;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.ClosingPrice;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.Paramount;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.YTW;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.Collateral;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.Positions.TsySpread;
                    if (item.i % 2 == 0)
                    {
                        worksheet.Cells[lastRowNumber, 1, rowNumber, columnNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[lastRowNumber, 1, rowNumber, columnNumber].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                    rowNumber++;
                }

                worksheet.Cells[rowNumber, 1].Value = $"Different records:{ _listHoldingsDifferences.Count()}";
                worksheet.Cells[rowNumber, 1, rowNumber, _listHoldingsDifferences.FirstOrDefault().Holdings.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Merge = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Bold = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Color.SetColor(Color.Ivory);
                worksheet.Cells[rowNumber, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rowNumber, 1].Style.Fill.BackgroundColor.SetColor(_listHoldingsDifferences.Count() > 0 ? Color.LightSalmon : Color.LightGreen);
                worksheet.Cells[rowNumber, 1, rowNumber, _listHoldingsDifferences.FirstOrDefault().Holdings.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                worksheet.Cells[rowNumber, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                #endregion

                #region Performance
                rowNumber++;
                worksheet.Cells[rowNumber, 1].Value = "Performance";
                worksheet.Cells[rowNumber, 1, rowNumber, _listPerformanceDifferences.FirstOrDefault().PerformanceReport.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Merge = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Bold = true;
                //worksheet.Cells[rowNumber, 1].Style.Font.Color.SetColor(Color.Ivory);
                worksheet.Cells[rowNumber, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rowNumber, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                worksheet.Cells[rowNumber, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                rowNumber++;
                columnNumber = 1;
                foreach (PropertyInfo propertyInfo in _listPerformanceDifferences.FirstOrDefault().PerformanceReport.GetType().GetProperties())
                {
                    if (propertyInfo.PropertyType == typeof(string))
                    {
                        worksheet.Cells[rowNumber, columnNumber].Value = propertyInfo.Name;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[rowNumber, columnNumber].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        worksheet.Cells[rowNumber, columnNumber].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        columnNumber++;
                    }
                }
                lastRowNumber = 0;
                rowNumber++;
                foreach (var item in _listPerformanceDifferences.Select((value, i) => new { i, value }))
                {
                    lastRowNumber = rowNumber;
                    columnNumber = 1;
                    //Row for performance report
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PerformanceReport.Source;
                    worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PerformanceReport.SinceInceptionReturn;
                    if (item.value.PerformanceReport.SinceInceptionReturnDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PerformanceReport.MTD;
                    if (item.value.PerformanceReport.MTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PerformanceReport.QTD;
                    if (item.value.PerformanceReport.QTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PerformanceReport.YTD;
                    if (item.value.PerformanceReport.YTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    //Next row for GetCellData
                    rowNumber++;
                    columnNumber = 1;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.Source;
                    worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.SinceInceptionReturn;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.MTD;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.QTD;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.YTD;

                    if (item.i % 2 == 0)
                    {
                        worksheet.Cells[lastRowNumber, 1, rowNumber, columnNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[lastRowNumber, 1, rowNumber, columnNumber].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                    rowNumber++;
                }
                worksheet.Cells[rowNumber, 1].Value = $"Different records:{ _listPerformanceDifferences.Count()}";
                worksheet.Cells[rowNumber, 1, rowNumber, _listPerformanceDifferences.FirstOrDefault().PerformanceReport.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Merge = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Bold = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Color.SetColor(Color.Ivory);
                worksheet.Cells[rowNumber, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rowNumber, 1].Style.Fill.BackgroundColor.SetColor(_listPerformanceDifferences.Count() > 0 ? Color.LightSalmon : Color.LightGreen);
                worksheet.Cells[rowNumber, 1, rowNumber, _listPerformanceDifferences.FirstOrDefault().PerformanceReport.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                worksheet.Cells[rowNumber, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                #endregion

                #region PortfoliStats
                rowNumber++;
                worksheet.Cells[rowNumber, 1].Value = "Portfolio Stats";
                worksheet.Cells[rowNumber, 1, rowNumber, _listPortfolioStatsDifferences.FirstOrDefault().PortfolioStatsReport.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Merge = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Bold = true;
                worksheet.Cells[rowNumber, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rowNumber, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                worksheet.Cells[rowNumber, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                rowNumber++;
                columnNumber = 1;
                foreach (PropertyInfo propertyInfo in _listPortfolioStatsDifferences.FirstOrDefault().PortfolioStatsReport.GetType().GetProperties())
                {
                    if (propertyInfo.PropertyType == typeof(string))
                    {
                        worksheet.Cells[rowNumber, columnNumber].Value = propertyInfo.Name;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[rowNumber, columnNumber].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        worksheet.Cells[rowNumber, columnNumber].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        columnNumber++;
                    }
                }
                lastRowNumber = 0;
                rowNumber++;
                foreach (var item in _listPortfolioStatsDifferences.Select((value, i) => new { i, value }))
                {
                    lastRowNumber = rowNumber;
                    columnNumber = 1;
                    //Row for portfoliostats report
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.Source;
                    worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.SecurityType;
                    if (item.value.PortfolioStatsReport.SecurityTypeDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.MarketValue;
                    if (item.value.PortfolioStatsReport.MarketValueDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.Coupon;
                    if (item.value.PortfolioStatsReport.CouponDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkCoupon;
                    if (item.value.PortfolioStatsReport.BenchmarkCouponDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.Weight;
                    if (item.value.PortfolioStatsReport.WeightDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkWeight;
                    if (item.value.PortfolioStatsReport.BenchmarkWeightDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.YTW;
                    if (item.value.PortfolioStatsReport.YTWDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkYTW;
                    if (item.value.PortfolioStatsReport.BenchmarkYTWDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.ModifiedDuration;
                    if (item.value.PortfolioStatsReport.ModifiedDurationDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkModifiedDuration;
                    if (item.value.PortfolioStatsReport.BenchmarkModifiedDurationDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.Convexity;
                    if (item.value.PortfolioStatsReport.ConvexityDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkConvexity;
                    if (item.value.PortfolioStatsReport.BenchmarkConvexityDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.DailyReturn;
                    if (item.value.PortfolioStatsReport.DailyReturnDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkDailyReturn;
                    if (item.value.PortfolioStatsReport.BenchmarkDailyReturnDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.MTD;
                    if (item.value.PortfolioStatsReport.MTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkMTD;
                    if (item.value.PortfolioStatsReport.BenchmarkMTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.QTD;
                    if (item.value.PortfolioStatsReport.QTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkQTD;
                    if (item.value.PortfolioStatsReport.BenchmarkQTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.YTD;
                    if (item.value.PortfolioStatsReport.YTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkYTD;
                    if (item.value.PortfolioStatsReport.BenchmarkYTDDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.CummReturn;
                    if (item.value.PortfolioStatsReport.CummReturnDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.PortfolioStatsReport.BenchmarkCummReturn;
                    if (item.value.PortfolioStatsReport.BenchmarkCummReturnDifference)
                    {
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                        worksheet.Cells[rowNumber, columnNumber].Style.Font.Color.SetColor(Color.Red);
                    }
                    //Next row for GetCellData
                    rowNumber++;
                    columnNumber = 1;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.Source;
                    worksheet.Cells[rowNumber, columnNumber].Style.Font.Bold = true;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.SecurityType;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.MarketValue;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.Coupon;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkCoupon;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.Weight;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkWeight;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.YTW;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkYTW;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.ModifiedDuration;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkModifiedDuration;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.Convexity;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkConvexity;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.DailyReturn;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkDailyReturn;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.MTD;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkMTD;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.QTD;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkQTD;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.YTD;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkYTD;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.CummReturn;
                    columnNumber++;
                    worksheet.Cells[rowNumber, columnNumber].Value = item.value.CellDataSP.BenchmarkCummReturn;

                    if (item.i % 2 == 0)
                    {
                        worksheet.Cells[lastRowNumber, 1, rowNumber, columnNumber].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[lastRowNumber, 1, rowNumber, columnNumber].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                    rowNumber++;
                }
                worksheet.Cells[rowNumber, 1].Value = $"Different records:{ _listPortfolioStatsDifferences.Count()}";
                worksheet.Cells[rowNumber, 1, rowNumber, _listPortfolioStatsDifferences.FirstOrDefault().PortfolioStatsReport.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Merge = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Bold = true;
                worksheet.Cells[rowNumber, 1].Style.Font.Color.SetColor(Color.Ivory);
                worksheet.Cells[rowNumber, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rowNumber, 1].Style.Fill.BackgroundColor.SetColor(_listPortfolioStatsDifferences.Count() > 0 ? Color.LightSalmon : Color.LightGreen);
                worksheet.Cells[rowNumber, 1, rowNumber, _listPortfolioStatsDifferences.FirstOrDefault().PortfolioStatsReport.GetType().GetProperties().Where(x => x.PropertyType.Name == "String").Count()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                worksheet.Cells[rowNumber, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                #endregion

                worksheet.Cells.AutoFitColumns();

                //Save your file
                FileInfo fi = new FileInfo(filePath);
                excelPackage.SaveAs(fi);
            }
        }
    }
}
