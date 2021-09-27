using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service
{
    public class AuthorityService : IAuthorityService
    {
        private IQualityPass1Provider _IQualityPass1Provider;

        public IAuthorityProvider _AuthorityProvider { get; set; }

        public UserList Get_User_List_Browse()
        {
            UserList result = new UserList();

            try
            {
                _AuthorityProvider = new AuthorityProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<UserList_Browse> list = _AuthorityProvider.Get_User_List_Browse().ToList();
                result.DataList = list;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }


            return result;
        }

        public UserList_Authority Get_User_List_Detail(string UserID)
        {
            UserList_Authority result = new UserList_Authority();

            try
            {
                _AuthorityProvider = new AuthorityProvider(Common.ManufacturingExecutionDataAccessLayer);

                result = _AuthorityProvider.Get_User_List_Head(UserID).ToList().FirstOrDefault();
                List<Module_Detail> list = _AuthorityProvider.Get_User_List_Detail(UserID).ToList();

                result.DataList = list;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }


            return result;
        }

        public UserList_Authority Update_User_List_Detail(UserList_Authority_Request Req)
        {
            UserList_Authority result = new UserList_Authority();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _AuthorityProvider = new AuthorityProvider(_ISQLDataTransaction);

                result.Result = _AuthorityProvider.Update_User_List_Detail(Req);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }


        public AuthoritybyPosition Get_Position_Browse(string FactoryID)
        {
            AuthoritybyPosition result = new AuthoritybyPosition();

            try
            {
                _AuthorityProvider = new AuthorityProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<DatabaseObject.ResultModel.Quality_Position> list = _AuthorityProvider.Get_Position_Browse(FactoryID).ToList();
                result.DataList = list;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }


            return result;
        }

        public DatabaseObject.ResultModel.Quality_Position Get_Position_Detail(string Position)
        {
            DatabaseObject.ResultModel.Quality_Position result = new DatabaseObject.ResultModel.Quality_Position();

            try
            {
                _AuthorityProvider = new AuthorityProvider(Common.ManufacturingExecutionDataAccessLayer);

                var tmp = _AuthorityProvider.Get_Position_Head(Position).ToList().FirstOrDefault();
                if (tmp != null)
                {
                    result = tmp;
                }
                List<Module_Detail> list = _AuthorityProvider.Get_Position_Detail(Position).ToList();

                result.DataList = list;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }


            return result;
        }

        public DatabaseObject.ResultModel.Quality_Position Update_Position_Detail(Quality_Position_Request Req)
        {
            DatabaseObject.ResultModel.Quality_Position result = new DatabaseObject.ResultModel.Quality_Position();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _AuthorityProvider = new AuthorityProvider(_ISQLDataTransaction);

                result.Result = _AuthorityProvider.Update_Position_Detail(Req);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public DatabaseObject.ResultModel.Quality_Position Create_Position_Detail(Quality_Position_Request Req)
        {
            DatabaseObject.ResultModel.Quality_Position result = new DatabaseObject.ResultModel.Quality_Position();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _AuthorityProvider = new AuthorityProvider(_ISQLDataTransaction);

                // 檢查重複
                int existsCount = _AuthorityProvider.Check_Position_Exists(Req);
                if (existsCount > 0)
                {
                    throw new Exception($"<{Req.Position}, {Req.Factory}> Position Data exists.");
                }

                result.Result = _AuthorityProvider.Create_Position_Detail(Req);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public UserList_Authority ImportUsers(List<UserList_Browse> DataList)
        {
            UserList_Authority result = new UserList_Authority();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _AuthorityProvider = new AuthorityProvider(_ISQLDataTransaction);

                result.Result = _AuthorityProvider.ImportUsers(DataList);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public UserList GetAlUser()
        {
            UserList result = new UserList();

            try
            {
                _AuthorityProvider = new AuthorityProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<UserList_Browse> list = _AuthorityProvider.GetAllUser().ToList();

                result.DataList = list;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public ResultModelBase<Module_Detail> SaveBulkFGT_Pass1(Quality_Pass1 quality_Pass1)
        {
            DatabaseObject.ResultModel.Quality_Position result = new DatabaseObject.ResultModel.Quality_Position();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _AuthorityProvider = new AuthorityProvider(_ISQLDataTransaction);

                result.Result = _AuthorityProvider.Update_Brand(quality_Pass1);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public List<SelectListItem> GetPositionList(string FactoryID)
        {
            List<SelectListItem> result = new List<SelectListItem>();

            try
            {
                _AuthorityProvider = new AuthorityProvider(Common.ManufacturingExecutionDataAccessLayer);
                result = _AuthorityProvider.GetPositionList(FactoryID).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public IList<Brand> GetBrand()
        {
            _AuthorityProvider = new AuthorityProvider(Common.ProductionDataAccessLayer);
            List<Brand> brands = new List<Brand>();
            try
            {
                brands = _AuthorityProvider.GetBrands().ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return brands;
        }

        public IList<Quality_Menu_Detail> GetFunctionName(string BulkFGT_Brand)
        {
            _AuthorityProvider = new AuthorityProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<Quality_Menu_Detail> _Quality_Menu_Detail = new List<Quality_Menu_Detail>();
            try
            {
                _Quality_Menu_Detail = _AuthorityProvider.GetFunctionName(BulkFGT_Brand).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Quality_Menu_Detail;
        }
    }
}
