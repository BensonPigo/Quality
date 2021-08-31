using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestDetailFGWTProvider
    {
        IList<GarmentTest_Detail_FGWT_ViewModel> Get_GarmentTest_Detail_FGWT(Int64 ID, string No);

        bool Chk_FGWTExists(GarmentTest_Detail source);

        bool Create_FGWT(GarmentTest Master, GarmentTest_Detail source);

        bool Update_FGWT(List<GarmentTest_Detail_FGWT_ViewModel> source);
    }
}
