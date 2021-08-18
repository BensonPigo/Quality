using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IBACriteriaProvider
    {
        IList<BACriteria_Result> Get_BACriteria_Result(string OrderID, string StyleID, string BrandID, string SeasonID);
    }
}
