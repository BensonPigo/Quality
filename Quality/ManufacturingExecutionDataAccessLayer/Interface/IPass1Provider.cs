using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IPass1Provider
    {
        IList<Pass1> Get(Pass1 Item);
    }
}
