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
        public bool IsFollowAQL { get; set; }
        public string AqlInputType { get; set; }


        public List<SelectSewing> SelectedSewing { get; set; }
        public List<SelectSewingTeam> SelectedSewingTeam { get; set; }
        public List<SelectedPO> SelectedPO { get; set; }

        public List<SelectOrderShipSeq> SelectOrderShipSeq { get; set; }

        public List<SelectCarton> SelectCarton { get; set; }
        public List<SelectQtyBreakdown> SelectQtyBreakdownList { get; set; }

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

        public List<SelectListItem> AqlInputTypeList
        {
            get
            {
                List<SelectListItem> rtn = new List<SelectListItem>()
                {
                    new SelectListItem()
                    {
                        Text = "Auto",
                        Value =  "Auto",
                    },
                    new SelectListItem()
                    {
                        Text = "Manual",
                        Value =  "Manual",
                    }
                };

                return rtn;
            }

        }
    }

    /// <summary>
    /// 排序、必要時給預設值
    /// </summary>
    public static class SelectQtyBreakdownExtensions
    {
        public static List<SelectQtyBreakdown> ProcessSelectQtyBreakdown(this List<SelectQtyBreakdown> list)
        {
            if (list == null || !list.Any())
                return list;

            // 排序邏輯：先按 OrderID，再按 SizeCode 排序
            var sortedList = list.OrderBy(x => x.OrderID)
                               .ThenBy(x => GetSortKey(x.SizeCode))
                           .ToList();

            // 檢查是否所有 LineItem 都為 NULL
            bool allLineItemsNull = sortedList.All(x => x.LineItem == null);

            // 如果全部都是 NULL，按 OrderID 分組設定預設值
            if (allLineItemsNull)
            {
                // 按 OrderID 分組處理
                var groupedList = sortedList.GroupBy(x => x.OrderID);

                foreach (var group in groupedList)
                {
                    int defaultValue = 10; // 每個 OrderID 組都從 10 開始

                    foreach (var item in group)
                    {
                        item.DefaultLineItem = defaultValue;
                        defaultValue += 10;
                    }
                }
            }


            return sortedList;
        }

        private static string GetSortKey(string sizeCode)
        {
            if (string.IsNullOrEmpty(sizeCode))
                return "Z"; // 空值排在最後

            // 檢查是否為純數字
            if (int.TryParse(sizeCode, out int numericSize))
            {
                return $"A{numericSize:D5}"; // 純數字排在最前面
            }

            // 找出開頭的數字部分
            string numberPart = "";
            int i = 0;
            while (i < sizeCode.Length && char.IsDigit(sizeCode[i]))
            {
                numberPart += sizeCode[i];
                i++;
            }

            // 找出字母部分
            string letterPart = "";
            while (i < sizeCode.Length && char.IsLetter(sizeCode[i]))
            {
                letterPart += sizeCode[i];
                i++;
            }

            // 找出最後的數字部分（如果有的話）
            string lastNumberPart = "";
            while (i < sizeCode.Length && char.IsDigit(sizeCode[i]))
            {
                lastNumberPart += sizeCode[i];
                i++;
            }

            // 解析數字部分，如果解析失敗則使用 0
            int baseNumber = int.TryParse(numberPart, out int n1) ? n1 : 0;
            int suffixNumber = int.TryParse(lastNumberPart, out int n2) ? n2 : 0;

            // 構建排序鍵
            return $"B{baseNumber:D5}{letterPart}{suffixNumber:D2}";
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
        public string Customize4 { get; set; }
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

    public class SelectQtyBreakdown
    {
        public string FinalInspectionID { get; set; }
        public string OrderID { get; set; }
        public string Article { get; set; }
        public string SizeCode { get; set; }
        public int? LineItem { get; set; }
        public int DefaultLineItem { get; set; }
        public bool Junk { get; set; }
    }
}
