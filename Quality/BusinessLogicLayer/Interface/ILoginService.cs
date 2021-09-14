using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Interface
{
    public interface ILoginService
    {
        LogIn_Result LoginValidate(LogIn_Request logIn_Request);

        LogIn_Result Update_Pass1(Quality_Pass1_Request Req);

        List<string> GetFactory();

        List<Quality_Menu> GetMenus(string UserID);
    }
}
