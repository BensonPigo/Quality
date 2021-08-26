using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;
using System.Data;

namespace BusinessLogicLayer.Interface.SampleRFT
{
    public interface ICFTCommentsService
    {
        CFTComments_ViewModel Get_CFT_Orders(CFTComments_ViewModel Req);

        CFTComments_ViewModel Get_CFT_OrderComments(CFTComments_ViewModel CFTComments);


        //以下廢棄
        DataTable GetExcel_DataTable(CFTComments_ViewModel Req);
        CFTComments_ViewModel GetExcel(CFTComments_ViewModel Req);
    }
}
