using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Interface.StyleManagement
{
    public interface IStyleListService
    {
        StyleList Get_StyleInfo(StyleList_Request Req);
    }
}
