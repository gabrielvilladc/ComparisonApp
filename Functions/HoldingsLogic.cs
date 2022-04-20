using ComparisonApp.Entities;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ComparisonApp.Functions
{
    public class HoldingsLogic
    {
        public List<FinalHoldingsComparison> GetHoldingsDifferences(IServiceProvider services, string[] args)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            OperationLogger logger = provider.GetRequiredService<OperationLogger>();
            string perdioId = args[0];
            string periodDate = args[1];
            string clientName = args[2];
            string cellGroup = args[3];
            string productCode = logger.GetClientProductCode(clientName);


            //**********           HOLDINGS LIST, from SLCCORICDB (CORICAPP)        ******************
            //Getting the data from SP into a Data Table
            DataTable holdingsTable = logger.GetHoldingsReport(perdioId);
            //Converting DataTable to List
            List<ComparisonApp.Entities.HoldingsAndPositionsReport> holdingsList = new List<ComparisonApp.Entities.HoldingsAndPositionsReport>();
            holdingsList = GeneralLogic.ConvertDataTable<ComparisonApp.Entities.HoldingsAndPositionsReport>(holdingsTable);
            //Filttering the list by Client
            holdingsList = holdingsList.Where(x => x.ProductID == productCode).ToList();
            //Mapping the records to a standard class list
            List<FinalHoldings> _listFinalHoldings = new List<FinalHoldings>();
            _listFinalHoldings = logger.MappHoldings(holdingsList);


            //**********           POSITIONS LIST, from Development DB (GetPositions SP)        ******************
            //Getting the data from SP into a Data Table
            DataTable positionsReportTable = logger.GetPositionReport(periodDate, clientName, cellGroup);
            //Converting DataTable to List
            List<ComparisonApp.Entities.HoldingsAndPositionsReport> positionsList = new List<ComparisonApp.Entities.HoldingsAndPositionsReport>();
            positionsList = GeneralLogic.ConvertDataTable<ComparisonApp.Entities.HoldingsAndPositionsReport>(positionsReportTable);
            //Mapping the records to a standard class list
            List<FinalHoldings> _listFinalPositions = new List<FinalHoldings>();
            _listFinalPositions = logger.MappHoldings(positionsList);


            List<FinalHoldingsComparison> _listFinalDifferences = new List<FinalHoldingsComparison>();

            foreach (var item in _listFinalPositions.Select((value, i) => new { i, value }))
            {

                if (item.value.ClosingPrice != _listFinalHoldings[item.i].ClosingPrice)
                {
                    _listFinalHoldings[item.i].ClosingPriceDifference = true;
                }
                if (item.value.Collateral != _listFinalHoldings[item.i].Collateral)
                {
                    _listFinalHoldings[item.i].CollateralDifference = true;
                }
                if (item.value.Coupon != _listFinalHoldings[item.i].Coupon)
                {
                    _listFinalHoldings[item.i].CouponDifference = true;
                }
                if (item.value.MaturityDate != _listFinalHoldings[item.i].MaturityDate)
                {
                    _listFinalHoldings[item.i].MaturityDateDifference = true;
                }
                if (item.value.MarketValue != _listFinalHoldings[item.i].MarketValue)
                {
                    _listFinalHoldings[item.i].MarketValueDifference = true;
                }
                if (item.value.Moodys != _listFinalHoldings[item.i].Moodys)
                {
                    _listFinalHoldings[item.i].MoodysDifference = true;
                }
                if (item.value.Paramount != _listFinalHoldings[item.i].Paramount)
                {
                    _listFinalHoldings[item.i].ParamountDifference = true;
                }
                if (item.value.SAndP != _listFinalHoldings[item.i].SAndP)
                {
                    _listFinalHoldings[item.i].SAndPDifference = true;
                }
                if (item.value.TsySpread != _listFinalHoldings[item.i].TsySpread)
                {
                    _listFinalHoldings[item.i].TsySpreadDifference = true;
                }
                if (item.value.YTW != _listFinalHoldings[item.i].YTW)
                {
                    _listFinalHoldings[item.i].YTWDifference = true;
                }

                if (_listFinalHoldings[item.i].ClosingPriceDifference ||
                    _listFinalHoldings[item.i].CollateralDifference ||
                    _listFinalHoldings[item.i].CouponDifference ||
                    _listFinalHoldings[item.i].MaturityDateDifference ||
                    _listFinalHoldings[item.i].MarketValueDifference ||
                    _listFinalHoldings[item.i].MoodysDifference ||
                    _listFinalHoldings[item.i].ParamountDifference ||
                    _listFinalHoldings[item.i].SAndPDifference ||
                    _listFinalHoldings[item.i].TsySpreadDifference ||
                    _listFinalHoldings[item.i].YTWDifference)
                {
                    FinalHoldingsComparison finalHoldingsComparisson = new FinalHoldingsComparison()
                    {
                        Holdings = _listFinalHoldings[item.i],
                        Positions = item.value
                    };
                    _listFinalDifferences.Add(finalHoldingsComparisson);
                }
            }

            return _listFinalDifferences;
        }

    }
}
