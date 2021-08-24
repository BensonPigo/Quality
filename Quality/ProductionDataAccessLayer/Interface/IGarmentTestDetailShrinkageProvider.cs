using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestDetailShrinkageProvider
    {
        IList<GarmentTest_Detail_Shrinkage> Get_GarmentTest_Detail_Shrinkage(Int64 ID, string No);
    }
}
