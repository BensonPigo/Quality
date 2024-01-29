using BusinessLogicLayer.Helper;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using Org.BouncyCastle.Math.EC.Rfc7748;
using Sci;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Text;

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
        public IADAuthAPICheck _ADAuthAPICheck { get; set; }
        public LogIn_Result LoginValidate(LogIn_Request logIn_Request)
        {
            LogIn_Result result = new LogIn_Result();
            _ADAuthAPICheck = new ADAuthAPICheck();

            try
            {
                QualityPass1Provider = new QualityPass1Provider(Common.ManufacturingExecutionDataAccessLayer);

                // 判斷User ID是否存在 ManufacturingExecution.dbo.Quality_Pass1
                List<Quality_Pass1> quality_Pass1s = QualityPass1Provider.Get(new Quality_Pass1() { ID = logIn_Request.UserID }).ToList();
                if (quality_Pass1s.Count == 0)
                {
                    throw new Exception("User ID not exist.");
                }

                // 先判斷ID，在判斷密碼。
                // 判斷User ID是否存在 Production.dbo.Pass1
                PMSPass1Provider = new ProductionDataAccessLayer.Provider.MSSQL.Pass1Provider(Common.ProductionDataAccessLayer);
                List<DatabaseObject.ProductionDB.Pass1> pmsPass1 = PMSPass1Provider.Get(new DatabaseObject.ProductionDB.Pass1() { ID = logIn_Request.UserID }).ToList();

                List<Pass1> mesPass1 = new List<Pass1>();
                if (pmsPass1.Count == 0)
                {
                    // 不在Production，改抓ManufacturingExecution.dbo.Pass1，並判斷密碼是否正確
                    MESPass1Provider = new Pass1Provider(Common.ManufacturingExecutionDataAccessLayer);
                    mesPass1 = MESPass1Provider.Get(new Pass1() { ID = logIn_Request.UserID, Password = logIn_Request.Password.ToUpper() }).ToList();
                    if (mesPass1.Count == 0)
                    {
                        throw new Exception("Incorrect password.");
                    }
                }
                else if (!pmsPass1.Where(x => x.Password.ToUpper().Equals(logIn_Request.Password.ToUpper())).Any())
                {
                    // 在Production，判斷密碼是否正確
                    throw new Exception("Incorrect Password.");
                }

                // 取得AdAccount
                string AdAccount = pmsPass1.Count == 0 ? mesPass1.FirstOrDefault().ADAccount : pmsPass1.FirstOrDefault().ADAccount;

                if (string.IsNullOrEmpty(AdAccount))
                {
                    result.Result = false;
                    result.ErrorMessage = "AD Account is empty, please check with local IT.";
                    return result;
                }

                string region = Common.Region.Replace("PH1", "PHI");

                BaseResult adResult = _ADAuthAPICheck.ADAuthByRegion(region, AdAccount);
                if (!adResult.Result)
                {
                    result.Result = adResult.Result;
                    result.ErrorMessage = adResult.ErrorMessage;
                    return result;
                }

                result.Factorys = pmsPass1.Count == 0 ? 
                        mesPass1.FirstOrDefault().Factory.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList() :
                        pmsPass1.FirstOrDefault().Factory.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();

                FactoryProvider = new ProductionDataAccessLayer.Provider.MSSQL.FactoryProvider(Common.ProductionDataAccessLayer);
                var M = FactoryProvider.GetMDivisionID(logIn_Request.FactoryID);

                if (!result.Factorys.Where(x => x.Equals(logIn_Request.FactoryID)).Any())
                {
                    throw new Exception(string.Format("Not have permission of {0}.", logIn_Request.FactoryID));
                }

                result.UserMail = pmsPass1.Count == 0 ? mesPass1.FirstOrDefault().EMail : pmsPass1.FirstOrDefault().EMail;
                result.pass1 = quality_Pass1s.FirstOrDefault();

                QualityMenuProvider = new QualityMenuProvider(Common.ManufacturingExecutionDataAccessLayer);
                result.Menus = QualityMenuProvider.Get(result.pass1).ToList();
                result.MDivisionID = M.Any() ? M.FirstOrDefault().MDivisionID : string.Empty;
                result.FactoryID = result.Factorys.Where(x => x.Equals(logIn_Request.FactoryID)).Any() ? logIn_Request.FactoryID.Trim() : result.Factorys.FirstOrDefault().Trim();

                SewingLineProvider = new ProductionDataAccessLayer.Provider.MSSQL.SewingLineProvider(Common.ProductionDataAccessLayer);
                result.Lines = SewingLineProvider.GetSewinglineID().GroupBy(x => x.ID).Select(x => x.Key).ToList();

                BrandProvider = new ProductionDataAccessLayer.Provider.MSSQL.BrandProvider(Common.ProductionDataAccessLayer);
                result.Brands = BrandProvider.Get().GroupBy(x => x.ID).Select(x => x.Key).ToList();
                result.Result = true;
            }
            catch(Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
                result.Exception = ex;
            }

            return result;
        }


        public List<Quality_Menu> GetMenus(string UserID)
        {
            List<Quality_Menu> result = new List<Quality_Menu>();
            QualityMenuProvider = new QualityMenuProvider(Common.ManufacturingExecutionDataAccessLayer);
            try
            {
                result = QualityMenuProvider.GetByMenu_detail(new Quality_Pass1() { ID = UserID }).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }


            return result;
        }

        public List<string> GetFactory()
        {
            FactoryProvider = new ProductionDataAccessLayer.Provider.MSSQL.FactoryProvider(Common.ProductionDataAccessLayer);
            List<string> factorys = FactoryProvider.GetFtyGroup().GroupBy(x => x.FTYGroup).Select(x => x.Key).ToList();
            return factorys;
        }

        public LogIn_Request LoginValidateOnlyID(string UserID, string FactoryID, DateTime EndTime)
        {
            if (EndTime < DateTime.Now)
            {
                return null;
            }

            QualityPass1Provider = new QualityPass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            List<Quality_Pass1> quality_Pass1s = QualityPass1Provider.Get(new Quality_Pass1() { ID = UserID }).ToList();
            if (quality_Pass1s.Count == 0)
            {
                return null;
            }

            MESPass1Provider = new Pass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            List<Pass1> mesPass1 = MESPass1Provider.Get(new Pass1() { ID = UserID }).ToList();
            LogIn_Request result = mesPass1.Select(p => new LogIn_Request { UserID = p.ID, Password = p.Password, FactoryID = FactoryID }).FirstOrDefault();

            return result;
        }

        // 解密 JWT
        public  JWTToken_ViewModel DecodeJWT(string token, string cryptoKey)
        {
            var tokenHandler = new JwtSecurityTokenHandler();
            var key = new SymmetricSecurityKey(Convert.FromBase64String(cryptoKey));
            var tokenValidationParameters = new TokenValidationParameters
            {
                ValidateIssuerSigningKey = true,
                IssuerSigningKey = key,
                ValidateIssuer = false,
                ValidateAudience = false
            };

            try
            {
                var claimsPrincipal = tokenHandler.ValidateToken(token, tokenValidationParameters, out var securityToken);
                var jwtToken = (JwtSecurityToken)securityToken;
                var value = jwtToken.Claims.FirstOrDefault(x => x.Type == "Value")?.Value;
                var guid = jwtToken.Claims.FirstOrDefault(x => x.Type == "GUID")?.Value;
                return JsonConvert.DeserializeObject<JWTToken_ViewModel>(StringEncryptHelper.AesDecryptBase64(value, cryptoKey));
            }
            catch
            {
                return null;
            }
        }
    }
}
