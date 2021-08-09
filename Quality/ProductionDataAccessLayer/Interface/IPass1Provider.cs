using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IPass1Provider
    {
        IList<Pass1> Get(Pass1 Item);
    }
}
