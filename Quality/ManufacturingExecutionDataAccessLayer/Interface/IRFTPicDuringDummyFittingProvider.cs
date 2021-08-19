using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTPicDuringDummyFittingProvider
    {
        IList<RFT_PicDuringDummyFitting> Get(RFT_PicDuringDummyFitting Item);
        IList<RFT_PicDuringDummyFitting> Get_PicDuringDummy_Result(RFT_PicDuringDummyFitting_ViewModel Req);
        int Save_Upd_Ins(RFT_PicDuringDummyFitting Item);
        IList<RFT_PicDuringDummyFitting_ViewModel> Check_OrderID_Exists(string OrderID);
    }
}
