using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.StyleManagement;
using DatabaseObject.ResultModel.EtoEFlowChart;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.StyleManagement
{
    public class EtoEFlowChartService
    {
        private EtoEFlowChartProvider PMS_Provider;
        private EtoEFlowChartProvider MES_Provider;

        public EtoEFlowChart_ViewModel GetDate (EtoEFlowChart_Request Req)
        {
            PMS_Provider = new EtoEFlowChartProvider(Common.ProductionDataAccessLayer);
            MES_Provider = new EtoEFlowChartProvider(Common.ManufacturingExecutionDataAccessLayer);

            EtoEFlowChart_ViewModel model = new EtoEFlowChart_ViewModel()
            {
                Development = new Development(),
                Warehouse = new Warehouse(),
                Production = new Production(),
                FinalInspection = new DatabaseObject.ResultModel.EtoEFlowChart.FinalInspection()
            };
            try
            {
                Req.StyleUkey = PMS_Provider.Get_StyleUkey(Req);

                // Development
                model.Development.SampleRFT = MES_Provider.Get_Development_SampleRFT(Req);
                model.Development.SampleRFT_Detail = MES_Provider.Get_Development_SampleRFT_Detail(Req);
                model.Development.CFTComments = MES_Provider.Get_Development_SampleRFT_Detail_CFTComments(model.Development.SampleRFT_Detail.Select(o => o.OrderID).Distinct().ToList());

                model.Development.TestResult_Detail = PMS_Provider.Get_Development_TestResult_Detail(Req);
                model.Development.TestResult = string.Empty;

                // 只有外層顯示Fail才需要秀Detail
                if (model.Development.TestResult_Detail.Count == 0)
                {
                    model.Development.TestResult_Detail = new List<Development_TestResult>();
                    model.Development.TestResult = "N/A";
                }
                else if (model.Development.TestResult_Detail.Where(o => o.TestResult == "Fail").Any())
                {
                    model.Development.TestResult = "Fail";
                    model.Development.TestResult_Detail = model.Development.TestResult_Detail.Where(o => o.TestResult == "Fail").ToList();
                }
                else if (model.Development.TestResult_Detail.Where(o => o.TestResult == "Pass").Count() == model.Development.TestResult_Detail.Count)
                {
                    // 全數Pass才算Pass
                    model.Development.TestResult = "Pass";
                    model.Development.TestResult_Detail = new List<Development_TestResult>();
                }

                model.Development.RRLR_Detail = PMS_Provider.Get_Development_RRLR_Detail(Req);
                // RR或LR等於1即顯示Y；反之都顯示0則顯示N
                if (model.Development.RRLR_Detail.Where(o => o.RR == "V" || o.LR == "V").Any())
                {
                    model.Development.RRLR = "Y";
                }
                else
                {
                    model.Development.RRLR = "N";
                }

                model.Development.FD_Detail = PMS_Provider.Get_Development_FD_Detail(Req);
                // FD有資料的話，即顯示Y；反之顯示N
                if (model.Development.FD_Detail.Any())
                {
                    model.Development.FD = "Y";
                }
                else
                {
                    model.Development.FD = "N";
                }

                // Warehouse
                model.Warehouse.PhysicalInspection = PMS_Provider.Get_Warehouse_PhysicalInspection(Req);
                model.Warehouse.PhysicalInspection_Detail = PMS_Provider.Get_Warehouse_PhysicalInspection_Detail(Req);

                // Production
                model.Production.SubprocessRFT_Detail = PMS_Provider.Get_Production_SubprocessRFT_Detail(Req);
                model.Production.SubprocessRFT = model.Production.SubprocessRFT_Detail.Any() ? model.Production.SubprocessRFT_Detail.FirstOrDefault().SummaryRate : 0;

                model.Production.TestResult_Detail = PMS_Provider.Get_Production_TestResult_Detail(Req).FirstOrDefault();
                model.Production.TestResult = string.Empty;

                if (model.Production.TestResult_Detail.FGPTResult == "Fail" || model.Production.TestResult_Detail.FGWTResult == "Fail")
                {
                    model.Production.TestResult = "Fail";
                }
                else if (model.Production.TestResult_Detail.FGPTResult == "Pass" && model.Production.TestResult_Detail.FGWTResult == "Pass")
                {
                    // 全數Pass才算Pass
                    model.Production.TestResult = "Pass";
                }
                else if (model.Production.TestResult_Detail.FGPTResult == "N/A" && model.Production.TestResult_Detail.FGWTResult == "N/A")
                {
                    model.Production.TestResult = "N/A";
                }

                model.Production.InlineRFT = MES_Provider.Get_Production_InlineRFT(Req);
                model.Production.InlineRFT_Detail = MES_Provider.Get_Production_InlineRFT_Detail(Req);

                model.Production.EndlineWFT = MES_Provider.Get_Production_EndlineWFT(Req);
                model.Production.EndlineWFT_Detail = MES_Provider.Get_Production_EndlineWFT_Detail(Req);

                model.Production.MDPassRate = PMS_Provider.Get_Production_MDPassRate(Req);
                model.Production.MDPassRate_Detail = PMS_Provider.Get_Production_MDPassRate_Detail(Req);

                // FinalInspection
                model.FinalInspection = MES_Provider.Get_FinalInspection_Detail(Req).FirstOrDefault();

                model.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                model.ExecuteResult = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }
    }
}
