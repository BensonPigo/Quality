using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IOrdersProvider
    {
        IList<Orders> Get(Orders Item);
    }
}
