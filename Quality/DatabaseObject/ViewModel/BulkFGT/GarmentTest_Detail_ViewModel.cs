using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class GarmentTest_Detail_ViewModel : GarmentTest_Detail
    {
        public string DrySelect { get; set; }

        public string Above50 { get; set; }

        public string NeckSelect { get; set; }

        public string GarmentTest_Detail_Inspector { get; set; }

        public string GarmentTest_Detail_AddName { get; set; }

        public string GarmentTest_Detail_EditName { get; set; }
    }
}
