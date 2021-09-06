using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IOrderQtyProvider
    {
        IList<Order_Qty> GetDistinctArticle(Order_Qty Item);
    }
}
