using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel.FinalInspection;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IOrdersProvider
    {
        IList<Orders> Get(Orders Item);
        IList<Orders> GetOrderForInspection(FinalInspection_Request request);
        IList<PoSelect_Result> GetOrderForInspection_ByModel(FinalInspection_Request request);
        bool Check_Style_ExistsOrder(string OrderID, string BrandID, string SeasonID, string StyleID);
    }
}
