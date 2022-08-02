using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionDataAccessLayer.Interface
{
    public interface IStyleListProvider
    {
        int Check_SampleRFTInspection_Count(StyleList_Request Req);
        IList<StyleList> Get_StyleInfo(StyleList_Request req);
    }
}
