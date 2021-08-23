using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class GarmentTest_ViewModel : GarmentTest
    {
        public string WashName { get; set; }

        public string SpecialMark { get; set; }

        public DateTime? MinSciDelivery { get; set; }

        public DateTime? MinBuyerDelivery { get; set; }

        public string GarmentTestAddName { get; set; }

        public string GarmentTestEditName { get; set; }
    }
}
