using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Interface
{
    public interface IBACriteriaService
    {
        BACriteria_ViewModel Get_BACriteria_Result(BACriteria_ViewModel Req);
    }
}
