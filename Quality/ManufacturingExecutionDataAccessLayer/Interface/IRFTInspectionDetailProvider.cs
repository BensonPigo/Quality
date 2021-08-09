using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTInspectionDetailProvider
    {
        IList<RFT_Inspection_Detail> Top3Defects(RFT_Inspection inspection);
    }
}
