using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel.FinalInspection;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;
using System.Data;

namespace ProductionDataAccessLayer.Interface
{
    public interface IOrdersProvider
    {
        IList<Orders> Get(Orders Item);
        IList<Orders> GetOrderForInspection(FinalInspection_Request request);
        IList<PoSelect_Result> GetOrderForInspection_ByModel(FinalInspection_Request request);
        bool Check_Style_ExistsOrder(string OrderID, string BrandID, string SeasonID, string StyleID);

        DataTable Get_Orders_DataTable(string ID, string PoID);
    }
}
