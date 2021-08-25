using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface ICFTCommentsService
    {
        CFTComments_where GetCFT_Orders(CFTComments_where CFTComments);

        List<CFTComments_ViewModel> GetRFT_OrderComments(CFTComments_where CFTComments);
    }
}
