using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComparisonApp.Entities.Performance
{
    public class FinalPerformance
    {
        public string Source { get; set; }
        public string SinceInceptionReturn { get; set; }
        public string MTD { get; set; }
        public string QTD { get; set; }
        public string YTD { get; set; }
        public bool SinceInceptionReturnDifference { get; set; }
        public bool MTDDifference { get; set; }
        public bool QTDDifference { get; set; }
        public bool YTDDifference { get; set; }
    }
}
