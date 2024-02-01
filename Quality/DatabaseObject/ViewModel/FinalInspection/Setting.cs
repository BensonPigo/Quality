using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class Setting : BaseResult
    {
        public string FinalInspectionID { get; set; } = string.Empty;
        public string BrandID { get; set; } = string.Empty;
        public string InspectionStage { get; set; } = string.Empty;
        public DateTime? AuditDate { get; set; }
        public string SewingLineID { get; set; } = string.Empty;
        public string Shift { get; set; } = string.Empty;
        public string Team { get; set; } = string.Empty;
        public string InspectionTimes { get; set; } = string.Empty;

        public string AcceptableQualityLevelsUkey { get; set; } = string.Empty;
        public string AcceptableQualityLevelsProUkey { get; set; } = string.Empty;
        public int SampleSize { get; set; }
        public int AcceptQty { get; set; }
        public string AQLPlan { get; set; }
        public string AQLPlanNotFinal { get; set; }
        public string AQLProPlan { get; set; }
        public bool ReInspection { get; set; }


        public List<SelectSewing> SelectedSewing { get; set; }
        public List<SelectSewingTeam> SelectedSewingTeam { get; set; }
        public List<SelectedPO> SelectedPO { get; set; }

        public List<SelectOrderShipSeq> SelectOrderShipSeq { get; set; }

        public List<SelectCarton> SelectCarton { get; set; }

        public List<AcceptableQualityLevels> AcceptableQualityLevels { get; set; }
        public List<AcceptableQualityLevelsProList> AcceptableQualityLevelsPros { get; set; }
        public List<SelectListItem> AQLPlanList
        {
            get
            {
                if (this.AcceptableQualityLevels != null)
                {
                    List<SelectListItem> rtn = new List<SelectListItem>()
                    {
                        new SelectListItem()
                        {
                            Text = string.Empty,
                            Value = string.Empty,
                        }
                    };

                    foreach (var item in this.AcceptableQualityLevels.Select(o => new { o.AQLType, o.InspectionLevels }).Distinct())
                    {
                        string aqlType = item.AQLType.ToString();
                        string level = string.Empty;
                        switch (item.InspectionLevels)
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
                        rtn.Add(new SelectListItem()
                        {
                            Text = string.IsNullOrEmpty(level) ? $@"{aqlType}" : $@"{aqlType} {level}",
                            Value = string.IsNullOrEmpty(level) ? $@"{aqlType}" : $@"{aqlType} {level}",
                        });

                    }
                    return rtn;
                }
                else
                {
                    return new List<SelectListItem>();
                }
            }

        }
        public List<SelectListItem> AQLProPlanList
        {
            get
            {
                if (this.AcceptableQualityLevelsPros != null)
                {
                    List<SelectListItem> rtn = new List<SelectListItem>()
                    {
                        //new SelectListItem()
                        //{
                        //    Text = string.Empty,
                        //    Value = string.Empty,
                        //}
                    };

                    foreach (var item in this.AcceptableQualityLevelsPros.Select(o => new { o.BrandID, o.AQLType, o.InspectionLevels }).Distinct())
                    {
                        string brandID = item.BrandID.ToString();
                        string aqlType = item.AQLType.ToString();
                        string level = string.Empty;
                        switch (item.InspectionLevels)
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
                        rtn.Add(new SelectListItem()
                        {
                            //Text = string.IsNullOrEmpty(level) ? $@"{aqlType}" : $@"{aqlType} {level}",
                            //Value = string.IsNullOrEmpty(level) ? $@"{aqlType}" : $@"{aqlType} {level}",
                            Text = brandID,
                            Value = brandID,
                        });

                    }
                    return rtn;
                }
                else
                {
                    return new List<SelectListItem>();
                }
            }

        }
    }

    public class SelectSewingTeam
    {
        public bool Selected { get; set; }
        public string SewingTeamID { get; set; }
    }
    public class SelectSewing
    {
        public bool Selected { get; set; }
        public string SewingLine { get; set; }
    }

    public class SelectedPO
    {
        public string OrderID { get; set; }
        public string CustPONO { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }
        public string BrandID { get; set; }
        public string Seq { get; set; }
        public string Article { get; set; }
        public int Qty { get; set; }
        public int MetalContaminateQty { get; set; }
        public string Cartons { get; set; }
        public int AvailableQty { get; set; }
    }

    public class SelectOrderShipSeq
    {
        public bool Selected { get; set; }
        public string OrderID { get; set; }
        public string Seq { get; set; }
        public string ShipmodeID { get; set; }
        public string Article { get; set; }
        public int Qty { get; set; }
    }

    public class SelectCarton
    {
        public bool Selected { get; set; }
        public string OrderID { get; set; }
        public string Seq { get; set; }
        public string PackingListID { get; set; }
        public string CTNNo { get; set; }
        public int ShipQty { get; set; }
        public string Size { get; set; }
        public string QtyPerSize { get; set; }
    }

}
