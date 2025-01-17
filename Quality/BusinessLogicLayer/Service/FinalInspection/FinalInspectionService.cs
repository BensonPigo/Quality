using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using System.Reflection.Emit;
using Microsoft.IdentityModel.Tokens;
using System.Security.Policy;
using System.Windows.Ink;
using com.itextpdf.text.pdf;

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
                    var aa = acceptableQualityLevels.Where(s => s.Ukey == ukey).FirstOrDefault();


                    string aqlType = aa.AQLType.ToString();
                    string level = string.Empty;

                    switch (aa.InspectionLevels)
                    {
                        case "1":
                            level = "Level I";
                            break;
                        case "2":
                            level = "Level II";
                            break;
                        case "3":
                            level = "Level III";
                            break;
                        case "4":
                            level = "Level IV";
                            break;
                        case "5":
                            level = "Level V";
                            break;
                        case "S-4":
                            level = "Level S-4";
                            break;
                        case "100% Inspection":
                            aqlType = "100% Inspection";
                            level = "";
                            break;
                        default:
                            break;
                    }

                    return string.IsNullOrEmpty(level) ? $@"{aqlType}" : $@"{aqlType} {level}";
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

        public object GetPivot88Json(string ID, bool isNewType = false)
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
            DataRow drCheckList = resultPivot88.Tables[2].Rows[0];
            DataTable dtColor = resultPivot88.Tables[3];
            DataTable dtSizeArticle = resultPivot88.Tables[4];
            DataRow drStyleInfo = resultPivot88.Tables[5].Rows[0];
            DataTable dtDefectsDetail = resultPivot88.Tables[6];
            DataTable dtDefectsDetailImage = resultPivot88.Tables[7];
            DataTable dtOtherImage = resultPivot88.Tables[8];

            List<object> sections = new List<object>();

            string projectCode;
            if (isNewType)
            {
                projectCode = "APPTRANS4M";
            }
            else
            {
                ;
                projectCode = "APP";
            }

            string custPono;
            if (Sci.MyUtility.Convert.GetString(drFinalInspection["Customize4"]) != string.Empty)
            {
                custPono = Sci.MyUtility.Convert.GetString(drFinalInspection["Customize4"]);
            }
            else
            {
                custPono = Sci.MyUtility.Convert.GetString(drFinalInspection["CustPONO"]);
            }

            var listSku_number = dtSizeArticle.AsEnumerable().Select(s =>
            new
            {
                sku_number = isNewType
                ? (string.IsNullOrEmpty(s["SizeCode"].ToString())
                    ? s["Article"].ToString() : $"{s["Article"]}_{s["SizeCode"]}_{s["SeqNumber"]}")
                : (string.IsNullOrEmpty(s["SizeCode"].ToString())
                    ? s["Article"].ToString() : $"{s["Article"]}_{s["SizeCode"]}"),
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
                aql_critical = 1,
                max_minor_defects = 0,
                max_major_defects = drFinalInspection["DefectQty"],
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
                    aql_critical = 1,
                    supplier_booking_msg = DBNull.Value,
                    conclusion_remarks = drFinalInspection["OthersRemark"],
                    assignment = new
                    {
                        report_type = new { name = drFinalInspection["ReportTypeID"] },
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
                                erp_business_id = drStyleInfo["BrandAreaCode"],
                            },
                            po_number = custPono,
                            customer_po = drFinalInspection["CustomerPo"],
                            importer = new
                            {
                                erp_business_id = "Adidas001",
                            },
                            project = new
                            {
                                project_code = projectCode
                            }
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
                    title = "material_approval",
                    value = drGeneral.Field<bool>("IsMaterialApproval") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsMaterialApproval") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "sealing_sample",
                    value = drGeneral.Field<bool>("IsSealingSample") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsSealingSample") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "metal_detection",
                    value = drGeneral.Field<bool>("IsMetalDetection") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsMetalDetection") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "factory_disclaimer",
                    value = drGeneral.Field<bool>("IsFactoryDisclaimer") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsFactoryDisclaimer") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "a_01_compliance",
                    value = drGeneral.Field<bool>("IsA01Compliance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsA01Compliance") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "cpsia_compliance",
                    value = drGeneral.Field<bool>("IsCPSIACompliance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsCPSIACompliance") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "customer_country_specific_compliance",
                    value = drGeneral.Field<bool>("IsCustomerCountrySpecificCompliance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "general",
                    status = drGeneral.Field<bool>("IsCustomerCountrySpecificCompliance") ? "pass" : "na",
                    comment = "",
                },
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
                    title = "handfeel",
                    value = drCheckList.Field<bool>("IsHandfeel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsHandfeel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "appearance",
                    value = drCheckList.Field<bool>("IsAppearance") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsAppearance") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "print_emb_decorations",
                    value = drCheckList.Field<bool>("IsPrintEmbDecorations") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "fabric_artwork_checklist",
                    status = drCheckList.Field<bool>("IsPrintEmbDecorations") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "fiber_content",
                    value = drCheckList.Field<bool>("IsFiberContent") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drCheckList.Field<bool>("IsFiberContent") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "care_instructions",
                    value = drCheckList.Field<bool>("IsCareInstructions") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drCheckList.Field<bool>("IsCareInstructions") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "decorative_label",
                    value = drCheckList.Field<bool>("IsDecorativeLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drCheckList.Field<bool>("IsDecorativeLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "adicom_label",
                    value = drCheckList.Field<bool>("IsAdicomLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drCheckList.Field<bool>("IsAdicomLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "country_of_origin",
                    value = drCheckList.Field<bool>("IsCountryofOrigion") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drCheckList.Field<bool>("IsCountryofOrigion") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "size_key",
                    value = drCheckList.Field<bool>("IsSizeKey") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drCheckList.Field<bool>("IsSizeKey") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "8_flag_label",
                    value = drCheckList.Field<bool>("Is8FlagLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drCheckList.Field<bool>("Is8FlagLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "additional_label",
                    value = drCheckList.Field<bool>("IsAdditionalLabel") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "label",
                    status = drCheckList.Field<bool>("IsAdditionalLabel") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "shipping_mark",
                    value = drCheckList.Field<bool>("IsShippingMark") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "packaging",
                    status = drCheckList.Field<bool>("IsShippingMark") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "polybag",
                    value = drCheckList.Field<bool>("IsPolytagMarking") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "packaging",
                    status = drCheckList.Field<bool>("IsPolytagMarking") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "color_size_qty",
                    value = drCheckList.Field<bool>("IsColorSizeQty") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "packaging",
                    status = drCheckList.Field<bool>("IsColorSizeQty") ? "pass" : "na",
                    comment = "",
                },
                new {
                    title = "hangtag",
                    value = drCheckList.Field<bool>("IsHangtag") ? "Confirm" : "N/A",
                    type = "check-list",
                    subsection = "validation_and_checklist",
                    checkListSubsection = "packaging",
                    status = drCheckList.Field<bool>("IsHangtag") ? "pass" : "na",
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

        public object GetEndInlinePivot88Json(string ID, string inspectionType, bool isNewType = false)
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


            string projectCode;

            if (isNewType)
            {
                projectCode = "APPTRANS4M";
            }
            else
            {
                projectCode = "APP";
            }

            string custPono;
            if (Sci.MyUtility.Convert.GetString(drInspection["Customize4"]) != string.Empty)
            {
                custPono = Sci.MyUtility.Convert.GetString(drInspection["Customize4"]);
            }
            else
            {
                custPono = Sci.MyUtility.Convert.GetString(drInspection["CustPONO"]);
            }

            var listSku_number = dtSizeArticle.AsEnumerable().Select(s =>
            new
            {
                sku_number = isNewType
            ? (string.IsNullOrEmpty(s["SizeCode"].ToString())
                ? s["Article"].ToString()
                : $"{s["Article"]}_{s["SizeCode"]}_{s["SeqNumber"]}")
            : (string.IsNullOrEmpty(s["SizeCode"].ToString())
                ? s["Article"].ToString()
                : $"{s["Article"]}_{s["SizeCode"]}"),

                qty_to_inspect = s["ShipQty"].ToInt() == 0 ? 1 : s["ShipQty"]
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
                aql_critical = 0.01,
                max_minor_defects = 15,
                max_major_defects = 15,
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
                    aql_critical = 0.01,
                    supplier_booking_msg = "no comment",
                    conclusion_remarks = "no comment",
                    assignment = new
                    {
                        report_type = new { name = inspectionType == "InlineInspection" ? "APP - Inline" : "APP - End of Line" },
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
                                erp_business_id = drStyleInfo["BrandAreaCode"],
                            },
                            po_number = custPono,
                            customer_po = drInspection["CustCDID"],
                            importer = new
                            {
                                id = 215,
                                erp_business_id = "Adidas001",
                            },
                            project = new
                            {
                                project_code = projectCode
                            }
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

        private void Pivot88UserEmptyNotice(string checkTable, string inspectionID)
        {
            // 避免Endline/inline手動sent P88也發信的防呆(inspectionID是空的代表排程執行,要發信!)
            if (!inspectionID.IsNullOrEmpty())
            {
                return;
            }

            // 判斷當前日期是否在7~8點之間
            if (DateTime.Now < DateTime.Today.AddHours(7) || DateTime.Now >= DateTime.Today.AddHours(8))
            {
                return;
            }

            string sqlP88UserEmpty = $@"
declare @EOLInlineFromDateTransferToP88 date = (select EOLInlineFromDateTransferToP88 from system)

select No = rank() over(order by [QC ID]),a.[QC ID] 
from
(
	select  distinct [QC ID] = t.QC 
	from {checkTable} t with (nolock)
	inner join pass1 p on p.id= t.qc
	where  IsExportToP88 = 0 
	and t.CustPONO <> '' 
	and (t.AddDate >= @EOLInlineFromDateTransferToP88 or t.EditDate >= @EOLInlineFromDateTransferToP88) 
	and exists (select 1 from Production.dbo.Orders o with (nolock) where o.CustPONo = t.CustPONO and o.BrandID in ('Adidas'))
	and p.Pivot88UserName = ''
) a
";
            // QC沒建立P88Name就發信mailto= 402; by ISP20240517
            SQLParameterCollection parameter = new SQLParameterCollection();
            DataTable dtMailto = SQLDAL.ExecuteDataTable(CommandType.Text, "select * from mailto where ID='402'", parameter);
            if (dtMailto != null && dtMailto.Rows.Count > 0)
            {
                DataTable dtResult = SQLDAL.ExecuteDataTable(CommandType.Text, sqlP88UserEmpty, parameter);

                // 沒資料代表不須發信提醒user設定P88
                if (dtResult != null && dtResult.Rows.Count > 0)
                {
                    string strDesc = checkTable == "InlineInspectionReport" ? @"<p> Inline Create P88 account </p>" : "<p> Endline Create P88 account </p>";

                    string mailBody = MailTools.DataTableChangeHtml(dtResult, string.Empty, string.Empty, strDesc + Environment.NewLine + dtMailto.Rows[0]["Content"].ToString(), out System.Net.Mail.AlternateView plainView);

                    SendMail_Request sendMail_Request = new SendMail_Request()
                    {
                        To = dtMailto.Rows[0]["toAddress"].ToString(),
                        CC = dtMailto.Rows[0]["ccAddress"].ToString(),

                        Subject = dtMailto.Rows[0]["Subject"].ToString(),
                        Body = mailBody,
                    };

                    // ToAddress 是空的就不寄出去
                    if (!sendMail_Request.To.IsNullOrEmpty())
                    {
                        MailTools.SendMail(sendMail_Request);
                    }
                }
            }
        }

        public List<SentPivot88Result> SentPivot88(PivotTransferRequest pivotTransferRequest, ref string p88Json)
        {
            p88Json = string.Empty;
            string tmpp88Json = string.Empty;
            List<string> listInspectionID = new List<string>();
            List<SentPivot88Result> sentPivot88Results = new List<SentPivot88Result>();

            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);


            switch (pivotTransferRequest.InspectionType)
            {
                case "FinalInspection":
                    listInspectionID = _FinalInspectionProvider.GetPivot88FinalInspectionID(pivotTransferRequest.InspectionID, pivotTransferRequest.IsAutoSend);
                    break;
                case "InlineInspection":
                    listInspectionID = _FinalInspectionProvider.GetPivot88InlineInspectionID(pivotTransferRequest.InspectionID);
                    this.Pivot88UserEmptyNotice("InlineInspectionReport", pivotTransferRequest.InspectionID);
                    break;
                case "EndlineInspection":
                    listInspectionID = _FinalInspectionProvider.GetPivot88EndLineInspectionID(pivotTransferRequest.InspectionID);
                    this.Pivot88UserEmptyNotice("InspectionReport", pivotTransferRequest.InspectionID);
                    break;
                default:
                    break;
            }

            if (listInspectionID.Count == 0)
            {
                return sentPivot88Results;
            }

            // 每個 InspectionID要傳新舊兩個版本的JSON，舊版本就不改了，新版本額外註解寫出來

            // 舊版本
            sentPivot88Results = listInspectionID.Select(inspectionID =>
            {
                bool isSuccess = true;
                string errorMsg = string.Empty;
                string postBody = string.Empty;
                string uniqueKey = "sintex" + inspectionID;
                string requestUri = pivotTransferRequest.RequestUri + uniqueKey;

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
                    automationErrMsg.moduleName = pivotTransferRequest.InspectionType;
                    automationErrMsg.apiThread = StaticPivot88.APIThreadFinalInsp;
                    automationErrMsg.errorMsg = errorMsg;
                    automationErrMsg.json = postBody;
                    automationErrMsg.addName = "SCIMIS";

                    _AutomationErrMsgProvider = new AutomationErrMsgProvider(Common.ProductionDataAccessLayer);
                    _AutomationErrMsgProvider.Insert(automationErrMsg);
                }

                return new SentPivot88Result()
                {
                    InspectionID = uniqueKey,
                    isSuccess = isSuccess,
                    errorMsg = errorMsg,
                };
            }).ToList();

            // 新版本
            sentPivot88Results = listInspectionID.Select(inspectionID =>
            {
                bool isSuccess = true;
                string errorMsg = string.Empty;
                string postBody = string.Empty;
                string uniqueKey = string.IsNullOrEmpty(pivotTransferRequest.P88UniqueKey) ? "sintex" + inspectionID + "Trans4m" : pivotTransferRequest.P88UniqueKey; // 新版本 + Trans4m
                string requestUri = pivotTransferRequest.RequestUri + uniqueKey;

                try
                {
                    #region 傳送finalinspection資料
                    // 新版本 isNewType = true
                    postBody = pivotTransferRequest.InspectionType == "FinalInspection" ?
                    JsonConvert.SerializeObject(GetPivot88Json(inspectionID, isNewType: true)) :
                    JsonConvert.SerializeObject(GetEndInlinePivot88Json(inspectionID, pivotTransferRequest.InspectionType, isNewType: true));
                    tmpp88Json = postBody;

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
                    automationErrMsg.moduleName = pivotTransferRequest.InspectionType;
                    automationErrMsg.apiThread = StaticPivot88.APIThreadFinalInsp;
                    automationErrMsg.errorMsg = errorMsg;
                    automationErrMsg.json = postBody;
                    automationErrMsg.addName = "SCIMIS";

                    _AutomationErrMsgProvider = new AutomationErrMsgProvider(Common.ProductionDataAccessLayer);
                    _AutomationErrMsgProvider.Insert(automationErrMsg);
                }

                return new SentPivot88Result()
                {
                    InspectionID = uniqueKey,
                    isSuccess = isSuccess,
                    errorMsg = errorMsg,
                };
            }).ToList();

            p88Json = tmpp88Json;

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

            List<FinalInspection_Step> allStep = _FinalInspectionProvider.GetAllStep(FinalInspectionID, CustPONO);

            return allStep;
        }
        public List<FinalInspectionBasicGeneral> GetGeneralByBrand(string FinalInspectionID, string BrandID, string InspectionStage)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<FinalInspectionBasicGeneral> allStep = _FinalInspectionProvider.GetGeneralByBrand(FinalInspectionID, BrandID, InspectionStage);

            return allStep;
        }
        public List<FinalInspectionBasicCheckList> GetCheckListByBrand(string FinalInspectionID, string BrandID)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<FinalInspectionBasicCheckList> allStep = _FinalInspectionProvider.GetCheckListByBrand(FinalInspectionID, BrandID);

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
