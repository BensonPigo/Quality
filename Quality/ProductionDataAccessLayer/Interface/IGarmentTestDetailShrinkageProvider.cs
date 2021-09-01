using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Data;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestDetailShrinkageProvider
    {
        IList<GarmentTest_Detail_Shrinkage> Get_GarmentTest_Detail_Shrinkage(string ID, string No);

        bool Update_GarmentTestShrinkage(List<GarmentTest_Detail_Shrinkage> source);

        DataTable Get_dt_Shrinkage(string ID, string No);
    }
}
