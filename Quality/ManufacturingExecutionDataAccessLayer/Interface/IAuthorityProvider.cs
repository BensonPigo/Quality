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

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IAuthorityProvider
    {
        IList<UserList_Browse> Get_User_List_Browse();
        IList<UserList_Authority> Get_User_List_Head(string UserID);
        IList<Module_Detail> Get_User_List_Detail(string UserID);
        bool Update_User_List_Detail(UserList_Authority_Request Req);

        IList<DatabaseObject.ResultModel.Quality_Position> Get_Position_Browse(string FactoryID);
        IList<DatabaseObject.ResultModel.Quality_Position> Get_Position_Head(string Position);
        IList<Module_Detail> Get_Position_Detail(string Position);
        bool Update_Position_Detail(Quality_Position_Request Req);
        int Check_Position_Exists(Quality_Position_Request Req);
        bool Create_Position_Detail(Quality_Position_Request Req);


        bool ImportUsers(List<UserList_Browse> DataList);
        IList<SelectListItem> GetPositionList(string FactoryID);
        IList<UserList_Browse> GetAllUser();

        bool Update_Brand(Quality_Pass1 Item);

        IList<Brand> GetBrands();
        IList<Quality_Menu_Detail> GetFunctionName(string BulkFGT_Brand);
    }
}
