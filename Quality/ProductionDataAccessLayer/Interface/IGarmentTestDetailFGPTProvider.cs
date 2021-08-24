using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestDetailFGPTProvider
    {
        IList<GarmentTest_Detail_FGPT_ViewModel> Get_GarmentTest_Detail_FGPT(Int64 ID, string No);
    }
}
