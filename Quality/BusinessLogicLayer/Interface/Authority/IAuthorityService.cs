using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
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
        UserList_Authority Get_User_List_Detail(string UserID);
        UserList_Authority Update_User_List_Detail(UserList_Authority_Request Req);

        AuthoritybyPosition Get_Position_Browse(string FactoryID);
        DatabaseObject.ResultModel.Quality_Position Get_Position_Detail(string Position);
        DatabaseObject.ResultModel.Quality_Position Update_Position_Detail(Quality_Position_Request Req);
        DatabaseObject.ResultModel.Quality_Position Create_Position_Detail(Quality_Position_Request Req);

        UserList GetAlUser(string FactoryID);
        List<SelectListItem> GetPositionList(string FactoryID);
        UserList_Authority ImportUsers(List<UserList_Browse> DataList);

        ResultModelBase<Module_Detail> SaveBulkFGT_Pass1(Quality_Pass1 quality_Pass1);

        IList<Brand> GetBrand();
        IList<Quality_Menu_Detail> GetFunctionName(string BulkFGT_Brand);
    }
}
