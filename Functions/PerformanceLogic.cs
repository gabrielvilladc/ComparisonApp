using ComparisonApp.Entities.Performance;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ComparisonApp.Functions
{
    public class PerformanceLogic
    {
        public List<FinalPerformanceComparison> GetPerformanceDifferences(IServiceProvider services, string[] args)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            OperationLogger logger = provider.GetRequiredService<OperationLogger>();
            string perdioId = args[0];
            string periodDate = args[1];
            string clientName = args[2];
            string cellGroup = args[3];
            string productCode = logger.GetClientProductCode(clientName);


            //**********           PERFORMANCE LIST, from SLCCORICDB (CORICAPP)        ******************
            //Getting the data from SP into a Data Table
            DataTable dataTable = logger.GetPerformanceReport(perdioId);
            //Converting DataTable to List
            List<PerformanceAndCellData> temporalList = new List<PerformanceAndCellData>();
            temporalList = GeneralLogic.ConvertDataTable<PerformanceAndCellData>(dataTable);
            //Filttering the list by Client
            temporalList = temporalList.Where(x => x.ProductID == productCode).ToList();
            //Mapping the records to a standard class list
            List<FinalPerformance> _finalList = new List<FinalPerformance>();
            _finalList = logger.MappPerformance(temporalList);

            //**********           PERFORMANCE LIST, from Development DB (GetCellData SP)        ******************
            //Getting the data from SP into a Data Table
            DataTable dataTable2 = logger.GetCellData(periodDate, clientName, cellGroup);
            //Converting DataTable to List
            List<PerformanceAndCellData> temporalList2 = new List<PerformanceAndCellData>();
            temporalList2 = GeneralLogic.ConvertDataTable<PerformanceAndCellData>(dataTable2);
            //Filttering the list by just TOTAL
            temporalList2 = temporalList2.Where(x => x.CellLabel == "TOTAL").ToList();
            //Mapping the records to a standard class list
            List<FinalPerformance> _finalList2 = new List<FinalPerformance>();
            _finalList2 = logger.MappPerformance(temporalList2);

            List<FinalPerformanceComparison> _listFinalDifferences = new List<FinalPerformanceComparison>();
            foreach (var item in _finalList2.Select((value, i) => new { i, value }))
            {
                if (item.value.SinceInceptionReturn != _finalList[item.i].SinceInceptionReturn)
                    _finalList[item.i].SinceInceptionReturnDifference = true;
                if (item.value.QTD != _finalList[item.i].QTD)
                    _finalList[item.i].QTDDifference = true;
                if (item.value.MTD != _finalList[item.i].MTD)
                    _finalList[item.i].MTDDifference = true;
                if (item.value.YTD != _finalList[item.i].YTD)
                    _finalList[item.i].YTDDifference = true;

                if (_finalList[item.i].SinceInceptionReturnDifference ||
                    _finalList[item.i].QTDDifference ||
                    _finalList[item.i].MTDDifference ||
                    _finalList[item.i].YTDDifference)
                {
                    FinalPerformanceComparison finalPerformanceComparisson = new FinalPerformanceComparison()
                    {
                        PerformanceReport = _finalList[item.i],
                        CellDataSP = item.value
                    };
                    _listFinalDifferences.Add(finalPerformanceComparisson);
                }
            }
            return _listFinalDifferences;
        }
    }
}
