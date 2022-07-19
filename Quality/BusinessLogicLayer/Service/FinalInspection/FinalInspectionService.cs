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
            DataTable dtColor = resultPivot88.Tables[1];
            DataTable dtSizeArticle = resultPivot88.Tables[2];
            DataRow drStyleInfo = resultPivot88.Tables[3].Rows[0];
            DataTable dtDefectsDetail = resultPivot88.Tables[4];
            DataTable dtDefectsDetailImage = resultPivot88.Tables[5];
            DataTable dtOtherImage = resultPivot88.Tables[6];

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
                new {
                    title = "fabric_approval",
                    value = drFinalInspection.Field<bool>("FabricApprovalDoc") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drFinalInspection.Field<bool>("FabricApprovalDoc") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "sealing_sample",
                    value = drFinalInspection.Field<bool>("SealingSampleDoc") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drFinalInspection.Field<bool>("SealingSampleDoc") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "metal_detection",
                    value = drFinalInspection.Field<bool>("MetalDetectionDoc") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drFinalInspection.Field<bool>("MetalDetectionDoc") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "finished_good_testing",
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
                    title = "moisture_control_garment",
                    value = drFinalInspection.Field<string>("MoistureResult") == "na" ? "N/A" : "Confirm",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drFinalInspection.Field<string>("MoistureResult"),
                    comment = drFinalInspection.Field<string>("MoistureComment"),
                },
                new {
                    title = "color_shade",
                    value = drFinalInspection.Field<bool>("CheckCloseShade") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drFinalInspection.Field<bool>("CheckCloseShade") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "handfeel",
                    value = drFinalInspection.Field<bool>("CheckHandfeel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drFinalInspection.Field<bool>("CheckHandfeel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "appearance",
                    value = drFinalInspection.Field<bool>("CheckAppearance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drFinalInspection.Field<bool>("CheckAppearance") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "print_emb_decorations",
                    value = drFinalInspection.Field<bool>("CheckPrintEmbDecorations") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drFinalInspection.Field<bool>("CheckPrintEmbDecorations") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "fiber_content",
                    value = drFinalInspection.Field<bool>("CheckFiberContent") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drFinalInspection.Field<bool>("CheckFiberContent") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "care_instructions",
                    value = drFinalInspection.Field<bool>("CheckCareInstructions") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drFinalInspection.Field<bool>("CheckCareInstructions") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "decorative_label",
                    value = drFinalInspection.Field<bool>("CheckDecorativeLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drFinalInspection.Field<bool>("CheckDecorativeLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "adicom_label",
                    value = drFinalInspection.Field<bool>("CheckAdicomLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drFinalInspection.Field<bool>("CheckAdicomLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "country_of_origin",
                    value = drFinalInspection.Field<bool>("CheckCountryofOrigion") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drFinalInspection.Field<bool>("CheckCountryofOrigion") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "size_key",
                    value = drFinalInspection.Field<bool>("CheckSizeKey") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drFinalInspection.Field<bool>("CheckSizeKey") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "8_flag_label",
                    value = drFinalInspection.Field<bool>("Check8FlagLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drFinalInspection.Field<bool>("Check8FlagLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "additional_label",
                    value = drFinalInspection.Field<bool>("CheckAdditionalLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drFinalInspection.Field<bool>("CheckAdditionalLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "shipping_mark",
                    value = drFinalInspection.Field<bool>("CheckShippingMark") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "packaging",
                    status = drFinalInspection.Field<bool>("CheckShippingMark") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "polybag",
                    value = drFinalInspection.Field<bool>("CheckPolytagMarketing") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "packaging",
                    status = drFinalInspection.Field<bool>("CheckPolytagMarketing") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_size_qty",
                    value = drFinalInspection.Field<bool>("CheckColorSizeQty") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "packaging",
                    status = drFinalInspection.Field<bool>("CheckColorSizeQty") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "hangtag",
                    value = drFinalInspection.Field<bool>("CheckHangtag") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "packaging",
                    status = drFinalInspection.Field<bool>("CheckHangtag") ? "pass" : "na",
                    comment = "",
                },
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
                    qty_inspected = (int)drInspection["PassQty"] + (int)drInspection["RejectQty"],
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

        public void ExecImp_EOLInlineInspectionReport()
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            _FinalInspectionProvider.ExecImp_EOLInlineInspectionReport();
        }
    }
}
