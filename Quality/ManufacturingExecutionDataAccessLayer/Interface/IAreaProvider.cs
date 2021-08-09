using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IAreaProvider
    {
        IList<Area> Get(Area Item);
    }
}
