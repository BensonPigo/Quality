using BusinessLogicLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BusinessLogicLayer.Service
{
    public class LoginService : ILoginService
    {
        private IQualityPass1Provider QualityPass1Provider;
        private IQualityMenuProvider QualityMenuProvider;
        private IPass1Provider MESPass1Provider;
        private ProductionDataAccessLayer.Interface.IFactoryProvider FactoryProvider;
        private ProductionDataAccessLayer.Interface.ISewingLineProvider SewingLineProvider;
        private ProductionDataAccessLayer.Interface.IBrandProvider BrandProvider;
        private ProductionDataAccessLayer.Interface.IPass1Provider PMSPass1Provider;
        public LogIn_Result LoginValidate(LogIn_Request logIn_Request)
        {
            QualityPass1Provider = new QualityPass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            QualityMenuProvider = new QualityMenuProvider(Common.ManufacturingExecutionDataAccessLayer);
            MESPass1Provider = new Pass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            FactoryProvider = new ProductionDataAccessLayer.Provider.MSSQL.FactoryProvider(Common.ProductionDataAccessLayer);
            SewingLineProvider = new ProductionDataAccessLayer.Provider.MSSQL.SewingLineProvider(Common.ProductionDataAccessLayer);
            PMSPass1Provider = new ProductionDataAccessLayer.Provider.MSSQL.Pass1Provider(Common.ProductionDataAccessLayer);
            BrandProvider = new ProductionDataAccessLayer.Provider.MSSQL.BrandProvider(Common.ProductionDataAccessLayer);
            LogIn_Result result = new LogIn_Result();

            try
            {
                List<Quality_Pass1> quality_Pass1s = QualityPass1Provider.Get(new Quality_Pass1() { ID = logIn_Request.UserID }).ToList();
                if (quality_Pass1s.Count == 0)
                {
                    throw new Exception("User ID not exist.");
                }

                // 先判斷ID，在判斷密碼。
                // ID 不存在改抓MES PASS1
                List<DatabaseObject.ProductionDB.Pass1> pmsPass1 = PMSPass1Provider.Get(new DatabaseObject.ProductionDB.Pass1() { ID = logIn_Request.UserID }).ToList();
                List<Pass1> mesPass1 = new List<Pass1>();
                if (pmsPass1.Count == 0)
                {
                    // 改抓MES PASS1
                    mesPass1 = MESPass1Provider.Get(new Pass1() { ID = logIn_Request.UserID, Password = logIn_Request.Password.ToUpper() }).ToList();
                    if (mesPass1.Count == 0)
                    {
                        throw new Exception("Incorrect password.");
                    }
                }
                else if (!pmsPass1.Where(x => x.Password.ToUpper().Equals(logIn_Request.Password.ToUpper())).Any())
                {
                    throw new Exception("Incorrect Password.");
                }

                result.Factorys = pmsPass1.Count == 0 ? 
                        mesPass1.FirstOrDefault().Factory.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList() :
                        pmsPass1.FirstOrDefault().Factory.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();

                var M = FactoryProvider.GetMDivisionID(logIn_Request.FactoryID);

                if (!result.Factorys.Where(x => x.Equals(logIn_Request.FactoryID)).Any())
                {
                    throw new Exception(string.Format("Not have permission of {0}.", logIn_Request.FactoryID));
                }

                result.pass1 = quality_Pass1s.FirstOrDefault();
                result.Menus = QualityMenuProvider.Get(result.pass1).ToList();
                result.MDivisionID = M.Any() ? M.FirstOrDefault().MDivisionID : string.Empty;
                result.FactoryID = result.Factorys.Where(x => x.Equals(logIn_Request.FactoryID)).Any() ? logIn_Request.FactoryID.Trim() : result.Factorys.FirstOrDefault().Trim();
                result.Lines = SewingLineProvider.GetSewinglineID().GroupBy(x => x.ID).Select(x => x.Key).ToList();
                result.Brands = BrandProvider.Get().GroupBy(x => x.ID).Select(x => x.Key).ToList();
                result.Result = true;
            }
            catch(Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
                result.Exception = ex;
            }

            return result;
        }

        public LogIn_Result Update_Pass1(Quality_Pass1_Request Req)
        {
            QualityPass1Provider = new QualityPass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            LogIn_Result result = new LogIn_Result();
            try
            {
                QualityPass1Provider.Update_Brand(new Quality_Pass1() { ID = Req.ID, BulkFGT_Brand = Req.SampleTesting_Brand });
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
                result.Exception = ex;
            }

            return result;
        }

        public List<string> GetFactory()
        {
            FactoryProvider = new ProductionDataAccessLayer.Provider.MSSQL.FactoryProvider(Common.ProductionDataAccessLayer);
            List<string> factorys = FactoryProvider.GetFtyGroup().GroupBy(x => x.FTYGroup).Select(x => x.Key).ToList();
            return factorys;
        }
    }
}
