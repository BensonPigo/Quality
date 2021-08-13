using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IReworkListProvider
    {
        IList<ReworkList_ViewModel> Get(ReworkList_ViewModel Item);

        IList<ReworkList_ViewModel> GetReworkListFilter(ReworkList_ViewModel inspection);
    }
}
