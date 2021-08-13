using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTOrderCommentsProvider
    {
        IList<RFT_OrderComments_ViewModel> Get(RFT_OrderComments Item);

        int Update(List<RFT_OrderComments> Items);

        int Save_upd_ins(List<RFT_OrderComments> Items);
    }
}
