using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
                _OrdersProvider = new OrdersProvider(Common.ManufacturingExecutionDataAccessLayer);

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
            DataRow drGeneral = resultPivot88.Tables[1].Rows[0];
            DataRow drCheckList= resultPivot88.Tables[2].Rows[0];
            DataTable dtColor = resultPivot88.Tables[3];
            DataTable dtSizeArticle = resultPivot88.Tables[4];
            DataRow drStyleInfo = resultPivot88.Tables[5].Rows[0];
            DataTable dtDefectsDetail = resultPivot88.Tables[6];
            DataTable dtDefectsDetailImage = resultPivot88.Tables[7];
            DataTable dtOtherImage = resultPivot88.Tables[8];

            List<object> sections = new List<object>();

            var listSku_number = dtSizeArticle.AsEnumerable().Select(s =>
            new
            {
                sku_number = $"{s["Article"]}_{s["SizeCode"]}",
                qty_to_inspect = s["ShipQty"]
            }
            );

            object defects;

            if (dtDefectsDetail.Rows.Count > 0)
            {
                defects = dtDefectsDetail.AsEnumerable()
                                .GroupJoin(dtDefectsDetailImage.AsEnumerable(),
                                            s => (long)s["Ukey"],
                                            defectImg =>
                                            (long)defectImg["FinalInspection_DetailUkey"],
                                            (s, defectImg) => new
                                            {
                                                defect = s,
                                                defectImg = defectImg.Select(defectImgItem => new
                                                {
                                                    title = defectImgItem["title"],
                                                    full_filename = defectImgItem["full_filename"],
                                                    number = defectImgItem["number"],
                                                    comment = defectImgItem["comment"],
                                                })
                                            })
                                .Select(s => new
                                {
                                    label = s.defect["DefectCodeDesc"],
                                    subsection = s.defect["DefectTypeDesc"],
                                    code = s.defect["Pivot88DefectCodeID"],
                                    critical_level = s.defect["CriticalQty"],
                                    major_level = s.defect["MajorQty"],
                                    minor_level = 0,
                                    comments = "No comment",
                                    pictures = s.defectImg
                                });
            }
            else
            {
                defects = new List<object>() {
                    new {
                            label = string.Empty,
                            subsection = string.Empty,
                            code = string.Empty,
                            critical_level = 0,
                            major_level = 0,
                            minor_level = 0,
                            comments = "No comment",
                            pictures = new List<object>(),
                    },
                };
            }

            sections.Add(new
            {
                type = "aqlDefects",
                title = "product",
                section_result_id = drFinalInspection["InspectionResultID"],
                defective_parts = drFinalInspection["RejectQty"],
                qty_inspected = drFinalInspection["POQty"],
                sampled_inspected = drFinalInspection["SampleSize"],
                inspection_level = drFinalInspection["InspectionLevel"],
                inspection_method = "normal",
                aql_minor = 4,
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
                defects,
            });

            if (dtOtherImage.Rows.Count == 0)
            {
                sections.Add(
                new
                {
                    type = "pictures",
                    title = "photos",
                    pictures = new List<object>(),
                }
                );
            }
            else
            {
                sections.Add(
                    new
                    {
                        type = "pictures",
                        title = "photos",
                        pictures = dtOtherImage.AsEnumerable()
                        .Select(s => new
                        {
                            title = s["title"],
                            full_filename = s["full_filename"],
                            number = s["number"],
                            comment = s["comment"],
                        }),
                    }
                    );
            }


            List<object> assignment_items = listSku_number.Select(
                sku_number => new
                {
                    sampled_inspected = drFinalInspection["SampleSize"],
                    inspection_result_id = drFinalInspection["InspectionResultID"],
                    inspection_status_id = drFinalInspection["InspectionStatusID"],
                    qty_inspected = sku_number.qty_to_inspect,
                    inspection_completed_date = drFinalInspection["InspectionCompletedDate"],
                    total_inspection_minutes = drFinalInspection["InspectionMinutes"],
                    sampling_size = drFinalInspection["SampleSize"],
                    qty_to_inspect = sku_number.qty_to_inspect,
                    aql_minor = 4,
                    aql_major = 1,
                    aql_major_a = 1,
                    aql_major_b = 1,
                    aql_critical = 1,
                    supplier_booking_msg = DBNull.Value,
                    conclusion_remarks = drFinalInspection["OthersRemark"],
                    assignment = new
                    {
                        report_type = new { id = drFinalInspection["ReportTypeID"] },
                        inspector = new { username = drFinalInspection["CFA"] },
                        date_inspection = drFinalInspection["AuditDate"],
                        inspection_level = drFinalInspection["InspectionLevel"],
                        inspection_method = "normal",
                    },
                    po_line = new
                    {
                        qty = sku_number.qty_to_inspect,
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
                            customer_po = drFinalInspection["CustomerPo"],
                            importer = new
                            {
                                id = 215,
                                erp_business_id = "Adidas001",
                            },
                            project = new { id = 2063 }
                        },
                        sku = new
                        {
                            sku_number = sku_number.sku_number,
                            item_name = "No Item",
                            item_description = string.Empty,
                        },
                    },
                }
                ).ToList<object>();

            #region passFails
            List<object> passFails = new List<object>() {
                // FinalInspectionGeneral 開始
                new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsMaterialApproval") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsMaterialApproval") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsSealingSample") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsSealingSample") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsMetalDetection") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsMetalDetection") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsFGWT") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsFGWT") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsFGPT") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsFGPT") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsTopSample") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsTopSample") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("Is3rdPartyTestReport") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("Is3rdPartyTestReport") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsPPSample") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsPPSample") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsGBTestForChina") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsGBTestForChina") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsCPSIAForYounthStytle") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsCPSIAForYounthStytle") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsQRSSample") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsQRSSample") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsFactoryDisclaimer") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsFactoryDisclaimer") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsA01Compliance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsA01Compliance") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsCPSIACompliance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsCPSIACompliance") ? "pass" : "na",
                    comment = "",
                },new {
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsCustomerCountrySpecificCompliance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsCustomerCountrySpecificCompliance") ? "pass" : "na",
                    comment = "",
                },
                // FinalInspectionGeneral 結束

                new {
                    title = "finished_goods_testing",
                    value = "Confirm",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drFinalInspection.Field<int>("BAQty") > 0 ? "pass" : "fail",
                    comment = "",
                },
                new {
                    title = "moisture_control_carton",
                    value = drFinalInspection.Field<string>("MoistureResult") == "na" ? "N/A" : "Confirm",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drFinalInspection.Field<string>("MoistureResult"),
                    comment = "",
                },
                new {
                    title = "moisture_control_product",
                    value = drFinalInspection.Field<string>("MoistureResult") == "na" ? "N/A" : "Confirm",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drFinalInspection.Field<string>("MoistureResult"),
                    comment = drFinalInspection.Field<string>("MoistureComment"),
                },
                // FinalInspectionGeneral 結束
                
                // FinalInspectionCheckList 開始
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsCloseShade") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsCloseShade") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsHandfeel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsHandfeel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsAppearance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsAppearance") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsPrintEmbDecorations") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsPrintEmbDecorations") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsEmbellishmentPrint") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsEmbellishmentPrint") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsEmbellishmentBonding") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsEmbellishmentBonding") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsEmbellishmentHT") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsEmbellishmentHT") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsEmbellishmentEMB") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsEmbellishmentEMB") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsFiberContent") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsFiberContent") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsCareInstructions") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsCareInstructions") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsDecorativeLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsDecorativeLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsAdicomLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsAdicomLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsCountryofOrigion") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsCountryofOrigion") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsSizeKey") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsSizeKey") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("Is8FlagLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("Is8FlagLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsAdditionalLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsAdditionalLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsIdLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsIdLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsMainLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsMainLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsSizeLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsSizeLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsCareContentLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsCareContentLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsBrandLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsBrandLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsBlueSignLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsBlueSignLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsLotLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsLotLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsSecurityLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsSecurityLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsSpecialLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsSpecialLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsVIDLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsVIDLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsCNC") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsCNC") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsWovenlabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsWovenlabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsTSize") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsTSize") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsCCLayout") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsCCLayout") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsShippingMark") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsShippingMark") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsPolytagMarketing") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsPolytagMarketing") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsColorSizeQty") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsColorSizeQty") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsHangtag") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsHangtag") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsJokerTag") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsJokerTag") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsWWMT") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsWWMT") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsChinaCIT") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsChinaCIT") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsPolybagSticker") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsPolybagSticker") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsUCCSticker") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsUCCSticker") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsPESheetMicropak") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsPESheetMicropak") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsAdditionalHantage") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsAdditionalHantage") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsUPCStickierHantage") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsUPCStickierHantage") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_shade",
                    value = drCheckList.Field<bool>("IsGS1128Label") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsGS1128Label") ? "pass" : "na",
                    comment = "",
                },
                // FinalInspectionCheckList 結束
                new {
                    title = "minimum_of_2_pcs_per_size_must_be_measured_in_line_with_adidas_inspection_sop",
                    value = "Confirm",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "measurements",
                    status = "pass",
                    comment = "",
                },
            };
            #endregion

            object result = new
            {
                status = "Submitted",
                date_started = drFinalInspection["DateStarted"],
                defective_parts = drFinalInspection["RejectQty"],
                sections,
                assignment_items,
                passFails,
            };

            return result;
        }

        public object GetEndInlinePivot88Json(string ID, string inspectionType)
        {
            IFinalInspectionProvider finalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            DataSet resultPivot88 = finalInspectionProvider.GetEndInlinePivot88(ID, inspectionType);

            if (resultPivot88.Tables.Count == 0)
            {
                return null;
            }

            if (resultPivot88.Tables[0].Rows.Count == 0)
            {
                return null;
            }

            DataRow drInspection = resultPivot88.Tables[0].Rows[0];
            DataTable dtSizeArticle = resultPivot88.Tables[1];
            DataRow drStyleInfo = resultPivot88.Tables[2].Rows[0];
            DataTable dtDefectsDetail = resultPivot88.Tables[3];
            DataTable dtDefectsDetailImage = resultPivot88.Tables[4];

            List<object> sections = new List<object>();

            var listSku_number = dtSizeArticle.AsEnumerable().Select(s =>
            new
            {
                sku_number = string.IsNullOrEmpty(s["SizeCode"].ToString()) ? s["Article"] : $"{s["Article"]}_{s["SizeCode"]}",
                qty_to_inspect = s["ShipQty"]
            }
            );

            object defects;

            if (dtDefectsDetail.Rows.Count > 0)
            {
                defects = dtDefectsDetail.AsEnumerable()
                                .GroupJoin(dtDefectsDetailImage.AsEnumerable(),
                                            s => (string)s["GarmentDefectCodeID"],
                                            defectImg =>
                                            (string)defectImg["GarmentDefectCodeID"],
                                            (s, defectImg) => new
                                            {
                                                defect = s,
                                                defectImg = defectImg.Select(defectImgItem => new
                                                {
                                                    title = defectImgItem["title"],
                                                    full_filename = defectImgItem["full_filename"],
                                                    number = defectImgItem["number"],
                                                    comment = "No comment",
                                                })
                                            })
                                .Select(s => new
                                {
                                    label = s.defect["label"],
                                    subsection = s.defect["subsection"],
                                    code = s.defect["code"],
                                    critical_level = s.defect["CriticalQty"],
                                    major_level = s.defect["MajorQty"],
                                    minor_level = 0,
                                    comments = "No comment",
                                    pictures = s.defectImg
                                });
            }
            else
            {
                defects = new List<object>() {
                    new {
                            label = string.Empty,
                            subsection = string.Empty,
                            code = string.Empty,
                            critical_level = 0,
                            major_level = 0,
                            minor_level = 0,
                            comments = "No comment",
                            pictures = new List<object>(),
                    },
                };
            }

            sections.Add(new
            {
                type = "aqlDefects",
                title = "product",
                section_result_id = 5,
                defective_parts = drInspection["RejectQty"],
                qty_inspected = (int)drInspection["PassQty"] + (int)drInspection["RejectQty"],
                sampled_inspected = (int)drInspection["PassQty"] + (int)drInspection["RejectQty"],
                inspection_level = "100%inspection",
                inspection_method = "normal",
                aql_minor = 4,
                aql_major = 1.5,
                aql_major_a = 1.5,
                aql_major_b = 1.5,
                aql_critical = 0.01,
                //aql_minor = 0,
                //aql_major = 0,
                //aql_major_a = drFinalInspection["DefectQty"],
                //aql_major_b = drFinalInspection["DefectQty"],
                //aql_critical = 0,
                max_minor_defects = 15,
                max_major_defects = 15,
                max_major_a_defects = 0,
                max_major_b_defects = 0,
                max_critical_defects = 15,
                defects,
            });

            sections.Add(
                new
                {
                    type = "pictures",
                    title = "photos",
                    pictures = new List<object>(),
                }
                );

            List<object> assignment_items = listSku_number.Select(
                sku_number => new
                {
                    fields = new
                    {
                        string_1 = drInspection["EndlineCGradeQty"],
                        string_2 = drInspection["FixQty"],
                        string_3 = drInspection["Shift"],
                        string_4 = drInspection["Operation"].ToString().Split(',').Take(10).JoinToString(","),
                        string_5 = drInspection["SewerID"].ToString().Split(',').Take(10).JoinToString(","),
                        string_6 = drInspection["Station"].ToString().Split(',').Take(10).JoinToString(","),
                        string_7 = drInspection["Line"].ToString(),
                        string_8 = drInspection["username"],
                        string_9 = "",
                        string_10 = drInspection["FactoryID"],
                        string_11 = drInspection["WFT"],

                    },
                    sampled_inspected = (int)drInspection["PassQty"] + (int)drInspection["RejectQty"],
                    inspection_result_id = 5,
                    inspection_status_id = 1,
                    qty_inspected = sku_number.qty_to_inspect,
                    inspection_completed_date = drInspection["LastinspectionDate"],
                    total_inspection_minutes = drInspection["InspectionMinutes"],
                    sampling_size = (int)drInspection["PassQty"] + (int)drInspection["RejectQty"],
                    qty_to_inspect = sku_number.qty_to_inspect,
                    aql_minor = 4,
                    aql_major = 1.5,
                    aql_major_a = 1.5,
                    aql_major_b = 1.5,
                    aql_critical = 0.01,
                    supplier_booking_msg = "no comment",
                    conclusion_remarks = "no comment",
                    assignment = new
                    {
                        report_type = new { id = inspectionType == "InlineInspection" ? 42 : 43 },
                        inspector = new { username = drInspection["username"] },
                        date_inspection = drInspection["FirstInspectionDate"],
                        inspection_level = "100%inspection",
                        inspection_method = "normal",
                    },
                    po_line = new
                    {
                        qty = sku_number.qty_to_inspect,
                        etd = drInspection["BuyerDelivery"],
                        eta = drInspection["BuyerDelivery"],
                        color = drInspection["Color"],
                        size = drInspection["Size"],
                        style = drStyleInfo["Style"],
                        po = new
                        {
                            exporter = new
                            {
                                id = drStyleInfo["BrandAreaID"],
                                erp_business_id = drStyleInfo["BrandAreaCode"],
                            },
                            po_number = drInspection["CustPONO"],
                            customer_po = drInspection["CustCDID"],
                            importer = new
                            {
                                id = 215,
                                erp_business_id = "Adidas001",
                            },
                            project = new { id = 2063 }
                        },
                        sku = new
                        {
                            sku_number = sku_number.sku_number,
                            item_name = "No Item",
                            item_description = string.Empty,
                        },
                    },
                }
                ).ToList<object>();

            #region passFails
            List<object> passFails = new List<object>();
            #endregion

            object result = new
            {
                status = "Submitted",
                date_started = drInspection["FirstInspectionDate"],
                defective_parts = drInspection["RejectQty"],
                sections,
                assignment_items,
                passFails,
            };

            return result;
        }

        public List<SentPivot88Result> SentPivot88ForKMTest(PivotTransferRequest pivotTransferRequest)
        {
            List<string> listInspectionID = new List<string>();
            List<SentPivot88Result> sentPivot88Results = new List<SentPivot88Result>();

            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            listInspectionID.Add(pivotTransferRequest.InspectionID);

            if (listInspectionID.Count == 0)
            {
                return sentPivot88Results;
            }

            sentPivot88Results = listInspectionID.Select(inspectionID =>
            {

                bool isSuccess = true;
                string errorMsg = string.Empty;
                string postBody = string.Empty;
                string requestUri = pivotTransferRequest.RequestUri + "sintex" + inspectionID;

                try
                {
                    #region 傳送finalinspection資料
                    postBody = pivotTransferRequest.InspectionType == "FinalInspection" ?
                    JsonConvert.SerializeObject(GetPivot88Json(inspectionID)) :
                    JsonConvert.SerializeObject(GetEndInlinePivot88Json(inspectionID, pivotTransferRequest.InspectionType));

                    WebApiBaseResult webApiBaseResult = WebApiTool.WebApiSend(pivotTransferRequest.BaseUri, requestUri, postBody, HttpMethod.Put, headers: pivotTransferRequest.Headers);

                    switch (webApiBaseResult.webApiResponseStatus)
                    {
                        case WebApiResponseStatus.Success:
                            isSuccess = true;
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
                    #endregion

                    #region 傳送圖片


                    if (isSuccess)
                    {
                        Dictionary<string, byte[]> dicImage = new Dictionary<string, byte[]>();

                        switch (pivotTransferRequest.InspectionType)
                        {
                            case "FinalInspection":
                                dicImage = _FinalInspectionProvider.GetFinalInspectionDefectImage(inspectionID);
                                break;
                            case "InlineInspection":
                                dicImage = _FinalInspectionProvider.GetInlineInspectionDefectImage(inspectionID);
                                break;
                            case "EndlineInspection":
                                dicImage = _FinalInspectionProvider.GetEndLineInspectionDefectImage(inspectionID);
                                break;
                            default:
                                break;
                        }

                        string requestUploadImgUri = requestUri + "/images/upload";
                        //string requestUploadImgUri = "rest/operation/v1/inspection_reports/unique_key:sintexSPSCH22020130/images/upload";

                        foreach (KeyValuePair<string, byte[]> imageInfo in dicImage)
                        {
                            MultipartFormDataContent contentPost = new MultipartFormDataContent();
                            contentPost.Add(new StreamContent(new MemoryStream(imageInfo.Value)), "file", imageInfo.Key);
                            webApiBaseResult = WebApiTool.WebApiSend(pivotTransferRequest.BaseUri, requestUploadImgUri, null, HttpMethod.Post, headers: pivotTransferRequest.Headers, httpContent: contentPost);

                            switch (webApiBaseResult.webApiResponseStatus)
                            {
                                case WebApiResponseStatus.Success:
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

                            if (!isSuccess)
                            {
                                break;
                            }
                        }
                    }

                    #endregion

                    if (isSuccess)
                    {
                        _FinalInspectionProvider.UpdateIsExportToP88(inspectionID, pivotTransferRequest.InspectionType);
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

                    _AutomationErrMsgProvider = new AutomationErrMsgProvider(Common.ProductionDataAccessLayer);
                    _AutomationErrMsgProvider.Insert(automationErrMsg);
                }

                return new SentPivot88Result()
                {
                    InspectionID = inspectionID,
                    isSuccess = isSuccess,
                    errorMsg = errorMsg,
                };
            }).ToList();

            return sentPivot88Results;
        }

        public List<SentPivot88Result> SentPivot88(PivotTransferRequest pivotTransferRequest)
        {
            List<string> listInspectionID = new List<string>();
            List<SentPivot88Result> sentPivot88Results = new List<SentPivot88Result>();

            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            switch (pivotTransferRequest.InspectionType)
            {
                case "FinalInspection":
                    listInspectionID = _FinalInspectionProvider.GetPivot88FinalInspectionID(pivotTransferRequest.InspectionID);
                    break;
                case "InlineInspection":
                    listInspectionID = _FinalInspectionProvider.GetPivot88InlineInspectionID(pivotTransferRequest.InspectionID);
                    break;
                case "EndlineInspection":
                    listInspectionID = _FinalInspectionProvider.GetPivot88EndLineInspectionID(pivotTransferRequest.InspectionID);
                    break;
                default:
                    break;
            }

            if (listInspectionID.Count == 0)
            {
                return sentPivot88Results;
            }

            sentPivot88Results = listInspectionID.Select(inspectionID =>
            {

                bool isSuccess = true;
                string errorMsg = string.Empty;
                string postBody = string.Empty;
                string requestUri = pivotTransferRequest.RequestUri + "sintex" + inspectionID;

                try
                {
                    #region 傳送finalinspection資料
                    postBody = pivotTransferRequest.InspectionType == "FinalInspection" ?
                    JsonConvert.SerializeObject(GetPivot88Json(inspectionID)) :
                    JsonConvert.SerializeObject(GetEndInlinePivot88Json(inspectionID, pivotTransferRequest.InspectionType));

                    WebApiBaseResult webApiBaseResult = WebApiTool.WebApiSend(pivotTransferRequest.BaseUri, requestUri, postBody, HttpMethod.Put, headers: pivotTransferRequest.Headers);

                    switch (webApiBaseResult.webApiResponseStatus)
                    {
                        case WebApiResponseStatus.Success:
                            isSuccess = true;
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
                    #endregion

                    #region 傳送圖片


                    if (isSuccess)
                    {
                        Dictionary<string, byte[]> dicImage = new Dictionary<string, byte[]>();

                        switch (pivotTransferRequest.InspectionType)
                        {
                            case "FinalInspection":
                                dicImage = _FinalInspectionProvider.GetFinalInspectionDefectImage(inspectionID);
                                break;
                            case "InlineInspection":
                                dicImage = _FinalInspectionProvider.GetInlineInspectionDefectImage(inspectionID);
                                break;
                            case "EndlineInspection":
                                dicImage = _FinalInspectionProvider.GetEndLineInspectionDefectImage(inspectionID);
                                break;
                            default:
                                break;
                        }

                        string requestUploadImgUri = requestUri + "/images/upload";
                        //string requestUploadImgUri = "rest/operation/v1/inspection_reports/unique_key:sintexSPSCH22020130/images/upload";

                        foreach (KeyValuePair<string, byte[]> imageInfo in dicImage)
                        {
                            MultipartFormDataContent contentPost = new MultipartFormDataContent();
                            contentPost.Add(new StreamContent(new MemoryStream(imageInfo.Value)), "file", imageInfo.Key);
                            webApiBaseResult = WebApiTool.WebApiSend(pivotTransferRequest.BaseUri, requestUploadImgUri, null, HttpMethod.Post, headers: pivotTransferRequest.Headers, httpContent: contentPost);

                            switch (webApiBaseResult.webApiResponseStatus)
                            {
                                case WebApiResponseStatus.Success:
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

                            if (!isSuccess)
                            {
                                break;
                            }
                        }
                    }

                    #endregion

                    if (isSuccess)
                    {
                        _FinalInspectionProvider.UpdateIsExportToP88(inspectionID, pivotTransferRequest.InspectionType);
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

                    _AutomationErrMsgProvider = new AutomationErrMsgProvider(Common.ProductionDataAccessLayer);
                    _AutomationErrMsgProvider.Insert(automationErrMsg);
                }

                return new SentPivot88Result()
                {
                    InspectionID = inspectionID,
                    isSuccess = isSuccess,
                    errorMsg = errorMsg,
                };
            }).ToList();

            return sentPivot88Results;
        }

        /// <summary>
        /// 部分功能Back/Next按鈕按下時，要存檔的東西(Remark之類的)
        /// </summary>
        /// <param name="finalInspection">FinalInspection.InspectionStep 存 「要去的Step」</param>
        /// <param name="currentStep">當下的Step</param>
        /// <param name="userID">登入者</param>
        /// <returns></returns>
        public BaseResult UpdateStepInfo(DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection, string currentStep, string userID)
        {
            BaseResult result = new BaseResult();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.UpdateStepInfo(finalInspection, currentStep, userID);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }
        public BaseResult UpdateGeneral(DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection)
        {
            BaseResult result = new BaseResult();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.UpdateGeneral(finalInspection.finalInspectionGeneral);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }
        public BaseResult UpdateCheckList(DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection)
        {
            BaseResult result = new BaseResult();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.UpdateCheckList(finalInspection.finalInspectionCheckList);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public void ExecImp_EOLInlineInspectionReport()
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            _FinalInspectionProvider.ExecImp_EOLInlineInspectionReport();
        }

        public List<FinalInspection_Step> GetAllStep(string FinalInspectionID, string CustPONO)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<FinalInspection_Step>  allStep = _FinalInspectionProvider.GetAllStep(FinalInspectionID, CustPONO);

            return allStep;
        }
        public List<FinalInspectionBasicGeneral> GetGeneralByBrand(string FinalInspectionID, string CustPONO)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<FinalInspectionBasicGeneral> allStep = _FinalInspectionProvider.GetGeneralByBrand(FinalInspectionID, CustPONO);

            return allStep;
        }
        public List<FinalInspectionBasicCheckList> GetCheckListByBrand(string FinalInspectionID, string CustPONO)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<FinalInspectionBasicCheckList> allStep = _FinalInspectionProvider.GetCheckListByBrand(FinalInspectionID, CustPONO);

            return allStep;
        }

        public List<FinalInspectionBasicGeneral> GetAllGeneral()
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<FinalInspectionBasicGeneral> allStep = _FinalInspectionProvider.GetAllGeneral();

            return allStep;
        }
        public List<FinalInspectionBasicCheckList> GetAllCheckList()
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<FinalInspectionBasicCheckList> allStep = _FinalInspectionProvider.GetAllCheckList();

            return allStep;
        }
        public BaseResult CheckStep(string FinalInspectionID)
        {
            BaseResult result = new BaseResult();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                var stepData = _FinalInspectionProvider.GetAllStepByAction(FinalInspectionID, FinalInspectionSStepAction.Current);


                if (!stepData.Any())
                {
                    result.Result = false;
                    result.ErrorMessage = "Located at an undefined [Inspection Step], will return to the previous valid [Inspection Step].";
                }

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;

        }

        public BaseResult UpdateStepByAction(string FinalInspectionID, string UserID, FinalInspectionSStepAction action)
        {
            BaseResult result = new BaseResult();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                _FinalInspectionProvider.UpdateStepByAction(FinalInspectionID, UserID, action);
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
