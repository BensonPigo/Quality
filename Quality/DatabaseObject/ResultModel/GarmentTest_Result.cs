using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class GarmentTest_Result
    {
        public List<string> SizeCodes { get; set; }

        public GarmentTest_ViewModel garmentTest { get; set; }

        public List<GarmentTest_Detail_ViewModel> garmentTest_Details { get; set; }
    }
}

