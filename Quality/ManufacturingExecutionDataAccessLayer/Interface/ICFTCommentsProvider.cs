using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface ICFTCommentsProvider
    {
        IList<CFTComments_ViewModel> Get_CFT_Orders(CFTComments_ViewModel CFTComments);
        /*
        IList<CFTComments_where> GetCFT_Orders(CFTComments_where CFTComments);
        */
        IList<CFTComments_Result> Get_CFT_OrderComments(CFTComments_ViewModel CFTComments);

        DataTable Get_CFT_OrderComments_DataTable(CFTComments_ViewModel Req);
    }
}
