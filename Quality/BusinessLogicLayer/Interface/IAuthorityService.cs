using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface
{
    public interface IAuthorityService
    {
        UserList Get_User_List_Browse();
        //UserList_Authority Get_User_List_Detail(string UserID);
        //UserList_Authority Update_User_List_Detail(UserList_Authority_Request Req);

        //AuthoritybyPosition Get_Position_Browse(string FactoryID);
        //Quality_Position Get_Position_Detail(string Position);
        //Quality_Position Update_Position_Detail(Quality_Position_Request Req);
        //Quality_Position Create_Position_Detail(Quality_Position_Request Req);

        UserList GetAlUser(string FactoryID);
        List<SelectListItem> GetPositionList(string FactoryID);
        //UserList_Authority ImportUsers(List<UserList_Browse> DataList);
    }
}
