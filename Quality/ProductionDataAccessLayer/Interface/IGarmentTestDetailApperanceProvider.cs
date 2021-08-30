using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestDetailApperanceProvider
    {
        IList<GarmentTest_Detail_Apperance_ViewModel> Get_GarmentTest_Detail_Apperance(Int64 ID, string No);
    }
}
