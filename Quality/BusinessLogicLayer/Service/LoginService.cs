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
        private ProductionDataAccessLayer.Interface.IPass1Provider PMSPass1Provider;
        public LogIn_Result LoginValidate(LogIn_Request logIn_Request)
        {
            QualityPass1Provider = new QualityPass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            QualityMenuProvider = new QualityMenuProvider(Common.ManufacturingExecutionDataAccessLayer);
            MESPass1Provider = new Pass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            FactoryProvider = new ProductionDataAccessLayer.Provider.MSSQL.FactoryProvider(Common.ProductionDataAccessLayer);
            SewingLineProvider = new ProductionDataAccessLayer.Provider.MSSQL.SewingLineProvider(Common.ProductionDataAccessLayer);
            PMSPass1Provider = new ProductionDataAccessLayer.Provider.MSSQL.Pass1Provider(Common.ProductionDataAccessLayer);
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
                if (pmsPass1.Count == 0)
                {
                    // 改抓MES PASS1
                    List<Pass1> mesPass1 = MESPass1Provider.Get(new Pass1() { ID = logIn_Request.UserID, Password = logIn_Request.Password.ToUpper() }).ToList();
                    if (mesPass1.Count == 0)
                    {
                        throw new Exception("Incorrect password.");
                    }
                }
                else if (!pmsPass1.Where(x => x.Password.ToUpper().Equals(logIn_Request.Password.ToUpper())).Any())
                {
                    throw new Exception("Incorrect Password.");
                }

                result.pass1 = quality_Pass1s.FirstOrDefault();
                result.Menus = QualityMenuProvider.Get(result.pass1.Position).ToList();
                result.Factorys = FactoryProvider.GetFtyGroup().GroupBy(x => x.FTYGroup).Select(x => x.Key).ToList();
                result.Lines = SewingLineProvider.GetSewinglineID().GroupBy(x => x.ID).Select(x => x.Key).ToList();
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
    }
}
