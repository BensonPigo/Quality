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

        IList<Order_Qty> GetSizeCode(string StyleID, string SeasonID, string BrandID);

        int Delete_GarmentTestDetail(string ID, string No, bool sameInstance);

        bool Update_GarmentTestDetail(GarmentTest_Detail_ViewModel source, bool sameInstance);

        bool Encode_GarmentTestDetail(string ID, string No, string Status);

        string Get_LastResult(string ID);

        bool Chk_AllResult(string ID, string No);

        IList<GarmentTest_Detail_ViewModel> Get(string ID, string No, bool sameInstance);

        List<string> GetScales();

        DataTable Get_Mail_Content(string ID, string No);

        bool Update_Sender(string ID, string No, string UserID);

        bool Update_Receive(string ID, string No, string UserID);

        bool Save_Detail_Picture(GarmentTest_Detail source);

        bool Update_GarmentTestDetail_Result(string ID, string No);

        bool Update_GarmentTestDetail_Result_Amend(string ID, string No);
        bool Encode_GarmentTestDetail_OrderIDCheck(string ID, string No);
    }
}
