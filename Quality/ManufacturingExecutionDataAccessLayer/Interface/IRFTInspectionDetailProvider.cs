using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTInspectionDetailProvider
    {
        IList<RFT_Inspection_Detail> Top3Defects(RFT_Inspection inspection_Detail);

        IList<RFT_Inspection_Detail> Get(RFT_Inspection_Detail inspection_Detail);

        int Create(RFT_Inspection_Detail inspection_Detail);

        int Create_Master_Detail(RFT_Inspection Master, List<RFT_Inspection_Detail> Detail);

        int Delete(RFT_Inspection_Detail inspection_Detail);
    }
}
