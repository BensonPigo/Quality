using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTPicDuringDummyFittingProvider
    {
        IList<RFT_PicDuringDummyFitting> Get(RFT_PicDuringDummyFitting Item);

        int Save_Upd_Ins(RFT_PicDuringDummyFitting Item);
    }
}
