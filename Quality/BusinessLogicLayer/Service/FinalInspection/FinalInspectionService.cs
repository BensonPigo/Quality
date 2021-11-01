using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.Public;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Newtonsoft.Json;
using PmsWebApiUtility45;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using ToolKit;
using static PmsWebApiUtility20.WebApiTool;

namespace BusinessLogicLayer.Service
{
    public class FinalInspectionService : IFinalInspectionService
    {
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
        public IOrdersProvider _OrdersProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }
        public IAutomationErrMsgProvider _AutomationErrMsgProvider { get; set; }
        public string GetAQLPlanDesc(long ukey)
        {
            switch (ukey)
            {
                case -1:
                    return "100% Inspection";
                case 0:
                    return string.Empty;
                default:
                    _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
                    List<AcceptableQualityLevels> acceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().ToList();

                    if (!acceptableQualityLevels.Any(s => s.Ukey == ukey))
                    {
                        return string.Empty;
                    }

                    return acceptableQualityLevels.Where(s => s.Ukey == ukey)
                        .Select(s => $"{s.AQLType.ToString("0.0")} Level {s.InspectionLevels.Replace("1", "I").Replace("2", "II")}")
                        .First();
            }
        }

        public DatabaseObject.ManufacturingExecutionDB.FinalInspection GetFinalInspection(string finalInspectionID)
        {
            DatabaseObject.ManufacturingExecutionDB.FinalInspection result = new DatabaseObject.ManufacturingExecutionDB.FinalInspection();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                result = _FinalInspectionProvider.GetFinalInspection(finalInspectionID);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public IList<Orders> GetOrderForInspection(FinalInspection_Request request)
        {
            try
            {
                _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);

                IList<Orders> result = _OrdersProvider.GetOrderForInspection(request);

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public IList<PoSelect_Result> GetOrderForInspection_ByModel(FinalInspection_Request request)
        {
            try
            {
                _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);

                IList<PoSelect_Result> result = _OrdersProvider.GetOrderForInspection_ByModel(request);

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public object GetPivot88Json(string ID)
        {
            IFinalInspectionProvider finalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            DataSet resultPivot88 = finalInspectionProvider.GetPivot88(ID);

            if (resultPivot88.Tables.Count == 0)
            {
                return null;
            }

            if (resultPivot88.Tables[0].Rows.Count == 0)
            {
                return null;
            }

            DataRow drFinalInspection = resultPivot88.Tables[0].Rows[0];
            DataTable dtColor = resultPivot88.Tables[1];
            DataTable dtSizeArticle = resultPivot88.Tables[2];
            DataRow drStyleInfo = resultPivot88.Tables[3].Rows[0];
            DataTable dtDefectsDetail = resultPivot88.Tables[4];

            List<object> sections = new List<object>();

            var listSku_number = dtSizeArticle.AsEnumerable().Select(s => $"{s["Article"]}_{s["SizeCode"]}");

            sections.Add(new
            {
                type = "checklists",
                title = "checklists",
                section_result_id = 1,
            });

            if (dtDefectsDetail.Rows.Count > 0)
            {
                sections.Add(new
                {
                    type = "aqlDefects",
                    title = "product",
                    section_result_id = drFinalInspection["InspectionResultID"],
                    defective_parts = drFinalInspection["DefectQty"],
                    qty_inspected = drFinalInspection["AvailableQty"],
                    sampled_inspected = drFinalInspection["SampleSize"],
                    inspection_level = drFinalInspection["InspectionLevel"],
                    inspection_method = "loosened",
                    aql_minor = 1,
                    aql_major = 1,
                    aql_major_a = 0.4,
                    aql_major_b = 1.5,
                    aql_critical = 1,
                    //aql_minor = 0,
                    //aql_major = 0,
                    //aql_major_a = drFinalInspection["DefectQty"],
                    //aql_major_b = drFinalInspection["DefectQty"],
                    //aql_critical = 0,
                    max_minor_defects = 0,
                    max_major_defects = drFinalInspection["DefectQty"],
                    max_major_a_defects = drFinalInspection["DefectQty"],
                    max_major_b_defects = drFinalInspection["DefectQty"],
                    max_critical_defects = 0,
                    defects = dtDefectsDetail.AsEnumerable()
                                .Select(s => new
                                {
                                    label = s["DefectCodeDesc"],
                                    subsection = s["DefectTypeDesc"],
                                    code = s["Pivot88DefectCodeID"],
                                    critical_level = s["CriticalQty"],
                                    major_level = s["MajorQty"],
                                    minor_level = 0,
                                    comments = "No comment",
                                })
                });
            }

            List<object> assignment_items = listSku_number.Select(
                sku_number => new
                {
                    sampled_inspected = drFinalInspection["SampleSize"],
                    inspection_result_id = drFinalInspection["InspectionResultID"],
                    inspection_status_id = drFinalInspection["InspectionStatusID"],
                    qty_inspected = drFinalInspection["AvailableQty"],
                    inspection_completed_date = drFinalInspection["SubmitDate"],
                    total_inspection_minutes = drFinalInspection["InspectionMinutes"],
                    sampling_size = drFinalInspection["SampleSize"],
                    qty_to_inspect = drFinalInspection["AvailableQty"],
                    assignment = new
                    {
                        report_type = new { id = 12 },
                        inspector = new { username = drFinalInspection["CFA"]},
                        date_inspection = drFinalInspection["AuditDate"]
                    },
                    po_line = new
                    {
                        qty = drFinalInspection["POQty"],
                        etd = drFinalInspection["ETD_ETA"],
                        eta = drFinalInspection["ETD_ETA"],
                        color = dtColor.Rows.Count > 0 ? dtColor.AsEnumerable().Select(s => s["ColorID"].ToString()).JoinToString(";") : string.Empty,
                        size = dtSizeArticle.Rows.Count > 0 ? dtSizeArticle.AsEnumerable().Select(s => s["SizeCode"].ToString()).Distinct().JoinToString(";") : string.Empty,
                        style = drStyleInfo["Style"],
                        po = new
                        {
                            exporter = new
                            {
                                id = drStyleInfo["BrandAreaID"],
                                erp_business_id = drStyleInfo["BrandAreaCode"],
                            },
                            po_number = drFinalInspection["CustPONO"],
                            customer_po = drFinalInspection["CustPONO"],
                            importer = new
                            {
                                id = 215,
                                erp_business_id = "Adidas001",
                            },
                            project = new { id = 2063 }
                        },
                        sku = new
                        {
                            sku_number,
                            item_name = "No Item",
                            item_description = string.Empty,
                        },
                    },
                }
                ).ToList<object>();


            object result = new
            {
                status = "Submitted",
                date_started = drFinalInspection["AuditDate"],
                defective_parts = drFinalInspection["RejectQty"],
                sections,
                assignment_items,
            };

            return result;
        }

        public List<SentPivot88Result> SentPivot88(PivotTransferRequest pivotTransferRequest)
        {
            List<string> listInspectionID = new List<string>();
            List<SentPivot88Result> sentPivot88Results = new List<SentPivot88Result>();

            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            _AutomationErrMsgProvider = new AutomationErrMsgProvider(Common.ProductionDataAccessLayer);

            listInspectionID = _FinalInspectionProvider.GetPivot88FinalInspectionID(pivotTransferRequest.InspectionID);

            if (listInspectionID.Count == 0)
            {
                return sentPivot88Results;
            }

            sentPivot88Results = listInspectionID.Select(finalInspectionID => {

                bool isSuccess = true;
                string errorMsg = string.Empty;
                string postBody = string.Empty;
                string requestUri = pivotTransferRequest.RequestUri + "sintex" + finalInspectionID;
                try
                {
                    postBody = JsonConvert.SerializeObject(GetPivot88Json(finalInspectionID));

                    WebApiBaseResult webApiBaseResult = WebApiTool.WebApiSend(pivotTransferRequest.BaseUri, requestUri, postBody, HttpMethod.Put, headers: pivotTransferRequest.Headers);
                    
                    switch (webApiBaseResult.webApiResponseStatus)
                    {
                        case WebApiResponseStatus.Success:
                            _FinalInspectionProvider.UpdateIsExportToP88(finalInspectionID);
                            break;
                        case WebApiResponseStatus.WebApiReturnFail:
                            isSuccess = false;
                            errorMsg = webApiBaseResult.responseContent;
                            break;
                        case WebApiResponseStatus.OtherException:
                            isSuccess = false;
                            errorMsg = webApiBaseResult.exception.ToString();
                            break;
                        case WebApiResponseStatus.ApiTimeout:
                            isSuccess = false;
                            errorMsg = "WebAPI timeout";
                            break;
                        default:
                            isSuccess = false;
                            errorMsg = "????";
                            break;
                    }
                }
                catch (Exception ex)
                {
                    isSuccess = false;
                    errorMsg = ex.ToString();
                }

                if (!isSuccess)
                {
                    AutomationErrMsg automationErrMsg = new AutomationErrMsg();
                    automationErrMsg.suppID = StaticPivot88.SuppID;
                    automationErrMsg.moduleName = StaticPivot88.ModuleNameFinalInsp;
                    automationErrMsg.apiThread = StaticPivot88.APIThreadFinalInsp;
                    automationErrMsg.errorMsg = errorMsg;
                    automationErrMsg.json = postBody;
                    automationErrMsg.addName = "SCIMIS";
                    _AutomationErrMsgProvider.Insert(automationErrMsg);
                }

                return new SentPivot88Result()
                {
                    InspectionID = finalInspectionID,
                    isSuccess = isSuccess,
                    errorMsg = errorMsg,
                };
            }).ToList();

            return sentPivot88Results;
        }

        /// <summary>
        /// FinalInspection 進度紀錄
        /// </summary>
        /// <param name="finalInspection">FinalInspection.InspectionStep 存 「要去的Step」</param>
        /// <param name="currentStep">當下的Step</param>
        /// <param name="userID">登入者</param>
        /// <returns></returns>
        public BaseResult UpdateFinalInspectionByStep(DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection, string currentStep, string userID)
        {
            BaseResult result = new BaseResult();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.UpdateFinalInspectionByStep(finalInspection, currentStep, userID);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }
    }
}
