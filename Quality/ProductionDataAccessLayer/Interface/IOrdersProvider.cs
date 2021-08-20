using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IOrdersProvider
    {
        IList<Orders> Get(Orders Item);
        IList<Orders> GetOrderForInspection(FinalInspection_Request request);
    }
}
