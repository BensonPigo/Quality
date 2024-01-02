using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ProductionDB
{
    public class AcceptableQualityLevelsPro_Detail
    {
        public long ProUkey { get; set; }
        public long AQLDefectCategoryUkey { get; set; }
        public int AcceptedQty { get; set; }
    }
}
