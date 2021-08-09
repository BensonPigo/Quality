using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTInspectionProvider
    {
        IList<RFT_Inspection> Get(RFT_Inspection inspection);
    }
}
