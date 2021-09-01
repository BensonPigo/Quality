using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestDetailApperanceProvider
    {
        IList<GarmentTest_Detail_Apperance_ViewModel> Get_GarmentTest_Detail_Apperance(string ID, string No);

        bool Update_Apperance(List<GarmentTest_Detail_Apperance_ViewModel> source);
    }
}
