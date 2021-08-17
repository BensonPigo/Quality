using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface ICFTCommentsProvider
    {
        IList<CFTComments_where> GetCFT_Orders(CFTComments_where CFTComments);

        IList<CFTComments_ViewModel> GetRFT_OrderComments(CFTComments_where CFTComments);
    }
}
