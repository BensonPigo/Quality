using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTInspectionProvider
    {
        IList<RFT_Inspection> Get(RFT_Inspection inspection);

        int Update(RFT_Inspection inspection);

        int Create(RFT_Inspection inspection);
    }
}
