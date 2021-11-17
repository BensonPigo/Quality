using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IInspectionProvider
    {
        IList<Inspection_ViewModel> GetSelectItemData(Inspection_ViewModel inspection_ViewModel);

        IList<Inspection_ViewModel> CheckSelectItemData(Inspection_ViewModel inspection_ViewModel);

        IList<Inspection_ChkOrderID_ViewModel> CheckSelectItemData_SP(Inspection_ViewModel inspection_ViewModel);
    }
}
