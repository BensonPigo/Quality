using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;
using System.Data;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTInspectionDetailProvider
    {
        IList<RFT_Inspection_Detail> Top3Defects(RFT_Inspection inspection_Detail);


        int Create_Master_Detail(RFT_Inspection Master, List<RFT_Inspection_Detail> Detail);

        int Create_Detail(RFT_Inspection_Detail Detail);

        DataTable ChkInspQty(RFT_Inspection filter);
    }
}
