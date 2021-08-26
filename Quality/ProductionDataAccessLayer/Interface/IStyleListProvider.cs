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
        IList<StyleList> Get_StyleInfo(StyleList_Request req);
    }
}
