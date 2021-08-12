using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IReworkCardProvider
    {
        IList<ReworkCard> Get(ReworkCard Item);
    }
}
