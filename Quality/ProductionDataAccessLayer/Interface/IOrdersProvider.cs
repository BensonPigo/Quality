using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IOrdersProvider
    {
        IList<Orders> Get(Orders Item);

        IList<Orders> GetForFinalInspection(FinalInspection_Request requestItem);
    }
}
