using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;
using System.Data;

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

        IList<GarmentTest_Detail_ViewModel> Get(string ID, string No);

        List<string> GetScales();

        DataTable Get_Mail_Content(string ID, string No);

        bool Update_Sender(string ID, string No, string UserID);

        bool Update_Receive(string ID, string No, string UserID);

        bool Save_Detail_Picture(GarmentTest_Detail source);
    }
}
