using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestDetailProvider
    {
        IList<GarmentTest_Detail_ViewModel> Get_GarmentTestDetail(GarmentTest filter);

        IList<Order_Qty> GetSizeCode(string OrderID, string Article);

        int Delete_GarmentTestDetail(string ID, string No);

        bool Update_GarmentTestDetail(GarmentTest_Detail_ViewModel source);

        bool Encode_GarmentTestDetail(string ID, string Status);

        string Get_LastResult(string ID);

        bool Chk_AllResult(string ID, string No);
    }
}
