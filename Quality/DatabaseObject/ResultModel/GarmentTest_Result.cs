using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
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
        public bool Result { get; set; }

        public string ErrMsg { get; set; }

        public List<string> SizeCodes { get; set; }

        public GarmentTest_ViewModel garmentTest { get; set; }

        public List<GarmentTest_Detail_ViewModel> garmentTest_Details { get; set; }

        public GarmentTest_Request req { get; set; }
    }


    public class GarmentTest_Detail_Result
    {
        public List<string> Scales { get; set; }

        public GarmentTest_ViewModel Main { get; set; }

        public GarmentTest_Detail_ViewModel Detail { get; set; }

        public List<GarmentTest_Detail_Shrinkage> Shrinkages { get; set; }

        public List<Garment_Detail_Spirality> Spiralities { get; set; }

        public List<GarmentTest_Detail_Apperance_ViewModel> Apperance { get; set; }

        public List<GarmentTest_Detail_FGWT_ViewModel> FGWT { get; set; }

        public List<GarmentTest_Detail_FGPT_ViewModel> FGPT { get; set; }

        public bool Result { get; set; }

        public string ErrMsg { get; set; }
    }
}

