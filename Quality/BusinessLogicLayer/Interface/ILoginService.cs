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
    }
}
