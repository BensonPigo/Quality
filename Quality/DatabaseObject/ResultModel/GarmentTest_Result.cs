using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class GarmentTest_Result
    {
        public string WashName { get; set; }

        public string SpecialMark { get; set; }

        public DateTime? MinSciDelivery { get; set; }

        public DateTime? MinBuyerDelivery { get; set; }

        public string GarmentTestAddName { get; set; }

        public string GarmentTestEditName { get; set; }

        public GarmentTest garmentTest { get; set; }

        public List<GarmentTest_Detail> garmentTest_Details { get; set; }
    }
}

