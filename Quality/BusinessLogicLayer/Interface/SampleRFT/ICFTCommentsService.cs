using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface.SampleRFT
{
    public interface ICFTCommentsService
    {
        CFTComments_ViewModel Get_CFT_Orders(CFTComments_ViewModel Req);

        CFTComments_ViewModel Get_CFT_OrderComments(CFTComments_ViewModel CFTComments);
        string ToExcel(CFTComments_ViewModel Req);
    }
}
