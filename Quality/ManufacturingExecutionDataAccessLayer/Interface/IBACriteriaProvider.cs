using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IBACriteriaProvider
    {
        int Check_SampleRFTInspection_Count(BACriteria_ViewModel Req);
        IList<BACriteria_Result> Get_BACriteria_Result(BACriteria_ViewModel Req);
        IList<BACriteria_Result> Get_BACriteria_Result_SampleRFTInspection(BACriteria_ViewModel Req);
    }
}
