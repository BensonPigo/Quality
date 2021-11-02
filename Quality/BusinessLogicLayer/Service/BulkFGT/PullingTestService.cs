using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class PullingTestService
    {
        private PullingTestProvider _PullingTestProvider;

        public PullingTest_ViewModel GetReportNoList(PullingTest_ViewModel Req)
        {
            PullingTest_ViewModel result = Req;

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result.ReportNo_Source = _PullingTestProvider.GetReportNoList(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            
            return result;
        }


        public PullingTest_ViewModel GetData(string ReportNo)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result.Detail = _PullingTestProvider.GetData(ReportNo);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_Result CheckSP(string POID)
        {
            PullingTest_Result result = new PullingTest_Result();
            PullingTest_Result result2 = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result = _PullingTestProvider.CheckSP(POID);

                // 透過Brand得到PullForceUnit
                result2 = _PullingTestProvider.GetPullForceUnit(result.BrandID);

                result.PullForceUnit = result2.PullForceUnit;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }


        public PullingTest_Result GetStandard(string BrandID, string TestItem, string PullForceUnit)
        {
            PullingTest_Result result = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result = _PullingTestProvider.GetStandard(BrandID, TestItem, PullForceUnit);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public PullingTest_Result GetPullUnit(string BrandID)
        {
            PullingTest_Result result = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result = _PullingTestProvider.GetPullForceUnit(BrandID);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public PullingTest_ViewModel Insert(PullingTest_Result Req)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Insert(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_ViewModel Update(PullingTest_Result Req)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Update(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_ViewModel Delete(string ReportNo)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Delete(ReportNo);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }


        public SendMail_Result FailSendMail(string ReportNo, string ToAddress, string CcAddress)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                DataTable dt = _PullingTestProvider.GetData_DataTable(ReportNo);

                string unit = dt.Rows[0]["PullForceUnit"].ToString();
                dt.Columns["PullForceUnit"].ColumnName = unit;
                dt.Rows[0][unit] = dt.Rows[0]["PullForce"].ToString();
                dt.Columns.Remove("PullForce");


                string mailBody = MailTools.DataTableChangeHtml(dt);

                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CcAddress,
                    Subject = "Pulling Test - Test Fail",
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
    }
}
