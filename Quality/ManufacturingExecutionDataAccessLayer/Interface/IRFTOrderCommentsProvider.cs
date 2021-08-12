using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTOrderCommentsProvider
    {
        IList<RFT_OrderComments> Get(RFT_OrderComments Item);
    }
}
