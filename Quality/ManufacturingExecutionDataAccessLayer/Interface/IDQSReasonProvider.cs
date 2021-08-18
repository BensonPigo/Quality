using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IDQSReasonProvider
    {
        IList<DQSReason> Get(DQSReason Item);
    }
}
