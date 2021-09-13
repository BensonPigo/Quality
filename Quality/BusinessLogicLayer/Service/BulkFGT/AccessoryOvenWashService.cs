using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using ProductionDataAccessLayer.Provider.MSSQL.BukkFGT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class AccessoryOvenWashService
    {
        private AccessoryOvenWashProvider _AccessoryOvenWashProvider;

        public Accessory_ViewModel GetMainData(Accessory_ViewModel Req)
        {
            Accessory_ViewModel result = new Accessory_ViewModel()
            {
                ReqOrderID = Req.ReqOrderID,
                DataList = new List<Accessory_Result>(),
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetHead(Req.ReqOrderID);
                result.DataList = _AccessoryOvenWashProvider.GetDetail(Req.ReqOrderID).ToList();

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }

            return result;
        }

        public Accessory_ViewModel Update(Accessory_ViewModel Req)
        {
            Accessory_ViewModel result = new Accessory_ViewModel()
            {
                ReqOrderID = Req.ReqOrderID,
                DataList = new List<Accessory_Result>(),
            };
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);

                int r = _AccessoryOvenWashProvider.Update_AIR_Laboratory(Req);

                result.Result = r > 0;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }


        #region Oven
        public Accessory_Oven GetOvenTest(Accessory_Oven Req)
        {

            Accessory_Oven result = new Accessory_Oven()
            {
                ErrorMessage = Req.ErrorMessage,
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetOvenTest(Req);
                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }

            return result;
        }

        public Accessory_Oven UpdateOven(Accessory_Oven Req)
        {
            Accessory_Oven result = new Accessory_Oven();

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                int r = _AccessoryOvenWashProvider.UpdateOvenTest(Req);

                result.Result = r > 0;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }

            return result;
        }

        public SendMail_Result SendOvenMail(Accessory_Oven Req)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
                DataTable dt = _AccessoryOvenWashProvider.GetData_OvenDataTable(Req);
                string mailBody = MailTools.DataTableChangeHtml(dt);

                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = Req.ToAddress,
                    CC = Req.CcAddress,
                    Subject = "Oven Test - Test Fail",
                    Body = mailBody
                };
                result = MailTools.SendMail(sendMail_Request);
                result.result = true;
            }
            catch (Exception ex)
            {

                result.result = false;
                result.resultMsg = ex.Message.ToString();
            }


            return result;
        }

        #endregion


        #region Wash
        public Accessory_Wash GetWashTest(Accessory_Wash Req)
        {

            Accessory_Wash result = new Accessory_Wash()
            {
                ErrorMessage = Req.ErrorMessage,
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetWashTest(Req);
                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }

            return result;
        }

        public Accessory_Wash UpdateWash(Accessory_Wash Req)
        {
            Accessory_Wash result = new Accessory_Wash();

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                int r = _AccessoryOvenWashProvider.UpdateWashTest(Req);

                result.Result = r > 0;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }

            return result;
        }

        public SendMail_Result SendWashMail(Accessory_Wash Req)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
                DataTable dt = _AccessoryOvenWashProvider.GetData_WashDataTable(Req);
                string mailBody = MailTools.DataTableChangeHtml(dt);

                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = Req.ToAddress,
                    CC = Req.CcAddress,
                    Subject = "Wash Test - Test Fail",
                    Body = mailBody
                };
                result = MailTools.SendMail(sendMail_Request);
                result.result = true;
            }
            catch (Exception ex)
            {

                result.result = false;
                result.resultMsg = ex.Message.ToString();
            }


            return result;
        }

        #endregion
    }
}
