using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class FabricColorFastnessController : BaseController
    {
        private IFabricColorFastness_Service _FabricColorFastness_Service;

        public FabricColorFastnessController()
        {
            _FabricColorFastness_Service = new FabricColorFastness_Service();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.WashingFastness,,"; // ISP20211364
        }

        [SessionAuthorizeAttribute]
        public ActionResult Index()
        {
            FabricColorFastness_ViewModel model = new FabricColorFastness_ViewModel()
            {
                ColorFastness_MainList = new List<ColorFastness_Result>()
            };

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        [SessionAuthorizeAttribute]
        public ActionResult IndexBack(string PoID)
        {
            FabricColorFastness_ViewModel model = _FabricColorFastness_Service.Get_Main(PoID);
            ViewBag.QueryPoID = PoID;
            return View("Index", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        [SessionAuthorizeAttribute]
        public ActionResult Query(string QueryPoID)
        {
            // 21051739BB
            FabricColorFastness_ViewModel model = _FabricColorFastness_Service.Get_Main(QueryPoID);

            if (!model.Result)
            {
                model.ColorFastness_MainList = new List<ColorFastness_Result>();
                model.ErrorMessage = $@"msg.WithInfo('{(string.IsNullOrEmpty(model.ErrorMessage) ? string.Empty : model.ErrorMessage.Replace("'", string.Empty)) }'); ";
            }

            ViewBag.QueryPoID = QueryPoID;
            return View("Index", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SaveMaster(FabricColorFastness_ViewModel Main)
        {
            var result = _FabricColorFastness_Service.Save_ColorFastness_1stPage(Main.PoID, Main.ColorFastnessLaboratoryRemark);
            ViewBag.QueryPoID = Main.PoID;
            return Json(result);
        }

        [SessionAuthorizeAttribute]
        public JsonResult MainDetailDelete(string ID)
        {
            BaseResult result = _FabricColorFastness_Service.DeleteColorFastness(ID);

            return Json(result);
        }

        [SessionAuthorizeAttribute]
        public ActionResult Detail(string PoID, string ID, string EditMode ,string BrandID)
        {
            FabricColorFastness_ViewModel FabricColorFastnessModel = new FabricColorFastness_ViewModel();
            Fabric_ColorFastness_Detail_ViewModel model = new Fabric_ColorFastness_Detail_ViewModel();
            if (string.IsNullOrEmpty(ID))
            {
                model.Main = new ColorFastness_Result();
                model.Details = new List<Fabric_ColorFastness_Detail_Result>();
                model.Main.Status = "New";
                model.Main.POID = PoID;
                model.Main.BrandID = BrandID;
            }
            else
            {
                model = _FabricColorFastness_Service.GetDetailBody(ID);
            }

            if (TempData["ModelFabricColorFastness"] != null)
            {
                Fabric_ColorFastness_Detail_ViewModel saveResult = (Fabric_ColorFastness_Detail_ViewModel)TempData["ModelFabricColorFastness"];
                model.Main.InspDate = saveResult.Main.InspDate;
                model.Main.Article = saveResult.Main.Article;
                model.Main.Inspector = saveResult.Main.Inspector;
                model.Main.InspectionName = saveResult.Main.InspectionName;
                model.Main.Cycle = saveResult.Main.Cycle;
                model.Main.CycleTime = saveResult.Main.CycleTime;
                model.Main.Detergent = saveResult.Main.Detergent;
                model.Main.Machine = saveResult.Main.Machine;
                model.Main.Drying = saveResult.Main.Drying;
                model.Main.Remark = saveResult.Main.Remark;
                model.Details = saveResult.Details;
                model.Result = saveResult.Result;
                model.ErrorMessage = $@"msg.WithInfo('{saveResult.ErrorMessage.Replace("'",string.Empty) }');EditMode=true;";
                EditMode = "True";
            }

            List<string> Scales = _FabricColorFastness_Service.Get_Scales();
            List<SelectListItem> ScalesList = new SetListItem().ItemListBinding(Scales);
            ViewBag.Temperature_List = FabricColorFastnessModel.Temperature_List;
            ViewBag.Cycle_List = FabricColorFastnessModel.Cycle_List;
            ViewBag.CycleTime_List = FabricColorFastnessModel.CycleTime_List;
            ViewBag.Detergent_List = FabricColorFastnessModel.Detergent_List;
            ViewBag.Machine_List = FabricColorFastnessModel.Machine_List;
            ViewBag.Drying_List = FabricColorFastnessModel.Drying_List;
            ViewBag.ChangeScaleList = ScalesList;
            ViewBag.ResultChangeList = FabricColorFastnessModel.Result_Source;
            //ViewBag.StainingScaleList = ScalesList;
            ViewBag.ResultStainList = FabricColorFastnessModel.Result_Source;
            ViewBag.EditMode = EditMode;
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View(model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult AddDetailRow(int lastNO, string GroupNO, bool IsUA)
        {
            FabricColorFastness_ViewModel FabricColorFastnessModel = new FabricColorFastness_ViewModel();
            List<string> Scales = _FabricColorFastness_Service.Get_Scales();


            int i = lastNO;
            string html = "";
            html += "<tr idx=" + i + ">";
            html += "<td><input id='Seq" + i + "' idx=" + i + " type ='hidden'></input><input id='Details_" + i + "__SubmitDate' name='Details[" + i + "].SubmitDate' class='form-control date-picker width9vw' type='text' value=''></td>"; //SubmitDate
            html += "<td><input id='Details_" + i + "__ColorFastnessGroup' name='Details[" + i + "].ColorFastnessGroup' type='number' max='99' maxlength='2' min='0' step='1' oninput='value=OvenGroupCheck(value)' value='" + GroupNO + "'></td>"; // ColorFastnessGroup
            html += "<td><div style='width:10vw;'><input id='Details_" + i + "__Seq' name='Details[" + i + "].Seq' idv='" + i.ToString() + "' class ='InputDetailSEQSelectItem' type='text'  style = 'width: 6vw'> <input id='btnDetailSEQSelectItem'  idv='" + i.ToString() + "' type='button' class='btnDetailSEQSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // seq
            html += "<td><div style='width:10vw;'><input id='Details_" + i + "__Roll' name='Details[" + i + "].Roll' idv='" + i.ToString() + "' class ='InputDetailRollSelectItem' type='text' style = 'width: 6vw'> <input id='btnDetailRollSelectItem' idv='" + i.ToString() + "' type='button' class='btnDetailRollSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // roll
            html += "<td><input id='Details_" + i + "__Dyelot' name='Details[" + i + "].Dyelot' type='text' readonly='readonly'></td>"; // Dyelot
            html += "<td><input id='Details_" + i + "__Refno' name='Details[" + i + "].Refno' type='text' readonly='readonly'></td>"; // Refno
            html += "<td><input id='Details_" + i + "__SCIRefno' name='Details[" + i + "].SCIRefno' type='text' readonly='readonly'></td>"; // SCIRefno
            html += "<td><input id='Details_" + i + "__ColorID' name='Details[" + i + "].SCIRefno' type='text' readonly='readonly'></td>"; // ColorID
            html += "<td><input id='Details_" + i + "__Result' name='Details[" + i + "].Result' type='text' readonly='readonly' class='blue width6vw' value='Pass'></td>"; // Result

            html += "<td><select id='Details_" + i + "__changeScale' name='Details[" + i + "].changeScale'>"; // changeScale
            foreach (string val in Scales)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Details_" + i + "__ResultChange' name='Details[" + i + "].ResultChange' class='blue' onchange='selectChange(this)'>"; // ResultChange
            foreach (var val in FabricColorFastnessModel.Result_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            if (IsUA)
            {
                //Staining
                html += "<td><select id='Details_" + i + "__StainingScale' name='Details[" + i + "].StainingScale'>"; // StainingScale
                foreach (string val in Scales)
                {
                    string selected = string.Empty;
                    if (val == "4-5")
                    {
                        selected = "selected";
                    }
                    html += $"<option {selected} value='" + val + "'>" + val + "</option>";
                }
                html += "</select></td>";

                html += "<td><select id='Details_" + i + "__ResultStain' name='Details[" + i + "].ResultStain' class='blue' onchange='selectChange(this)'>"; // ResultStain
                foreach (var val in FabricColorFastnessModel.Result_Source)
                {
                    html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
                }
                html += "</select></td>";
            }
            else
            {
                //Acetate
                html += "<td><select id='Details_" + i + "__AcetateScale' name='Details[" + i + "].AcetateScale'>"; // AcetateScale
                foreach (string val in Scales)
                {
                    string selected = string.Empty;
                    if (val == "4-5")
                    {
                        selected = "selected";
                    }
                    html += $"<option {selected} value='" + val + "'>" + val + "</option>";
                }
                html += "</select></td>";

                html += "<td><select id='Details_" + i + "__ResultAcetate' name='Details[" + i + "].ResultAcetate' class='blue' onchange='selectChange(this)'>"; // ResultAcetate
                foreach (var val in FabricColorFastnessModel.Result_Source)
                {
                    html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
                }
                html += "</select></td>";


                //Cotton
                html += "<td><select id='Details_" + i + "__CottonScale' name='Details[" + i + "].CottonScale'>"; // CottonScale
                foreach (string val in Scales)
                {
                    string selected = string.Empty;
                    if (val == "4-5")
                    {
                        selected = "selected";
                    }
                    html += $"<option {selected} value='" + val + "'>" + val + "</option>";
                }
                html += "</select></td>";

                html += "<td><select id='Details_" + i + "__ResultCotton' name='Details[" + i + "].ResultCotton' class='blue' onchange='selectChange(this)'>"; // ResultCotton
                foreach (var val in FabricColorFastnessModel.Result_Source)
                {
                    html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
                }
                html += "</select></td>";
                //Nylon
                html += "<td><select id='Details_" + i + "__NylonScale' name='Details[" + i + "].NylonScale'>"; // NylonScale
                foreach (string val in Scales)
                {
                    string selected = string.Empty;
                    if (val == "4-5")
                    {
                        selected = "selected";
                    }
                    html += $"<option {selected} value='" + val + "'>" + val + "</option>";
                }
                html += "</select></td>";

                html += "<td><select id='Details_" + i + "__ResultNylon' name='Details[" + i + "].ResultNylon' class='blue' onchange='selectChange(this)'>"; // ResultNylon
                foreach (var val in FabricColorFastnessModel.Result_Source)
                {
                    html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
                }
                html += "</select></td>";
                //Polyester
                html += "<td><select id='Details_" + i + "__PolyesterScale' name='Details[" + i + "].PolyesterScale'>"; // PolyesterScale
                foreach (string val in Scales)
                {
                    string selected = string.Empty;
                    if (val == "4-5")
                    {
                        selected = "selected";
                    }
                    html += $"<option {selected} value='" + val + "'>" + val + "</option>";
                }
                html += "</select></td>";

                html += "<td><select id='Details_" + i + "__ResultPolyester' name='Details[" + i + "].ResultPolyester' class='blue' onchange='selectChange(this)'>"; // ResultPolyester
                foreach (var val in FabricColorFastnessModel.Result_Source)
                {
                    html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
                }
                html += "</select></td>";
                //Acrylic
                html += "<td><select id='Details_" + i + "__AcrylicScale' name='Details[" + i + "].AcrylicScale'>"; // AcrylicScale
                foreach (string val in Scales)
                {
                    string selected = string.Empty;
                    if (val == "4-5")
                    {
                        selected = "selected";
                    }
                    html += $"<option {selected} value='" + val + "'>" + val + "</option>";
                }
                html += "</select></td>";

                html += "<td><select id='Details_" + i + "__ResultAcrylic' name='Details[" + i + "].ResultAcrylic' class='blue' onchange='selectChange(this)'>"; // ResultAcrylic
                foreach (var val in FabricColorFastnessModel.Result_Source)
                {
                    html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
                }
                html += "</select></td>";
                //Wool
                html += "<td><select id='Details_" + i + "__WoolScale' name='Details[" + i + "].WoolScale'>"; // WoolScale
                foreach (string val in Scales)
                {
                    string selected = string.Empty;
                    if (val == "4-5")
                    {
                        selected = "selected";
                    }
                    html += $"<option {selected} value='" + val + "'>" + val + "</option>";
                }
                html += "</select></td>";

                html += "<td><select id='Details_" + i + "__ResultWool' name='Details[" + i + "].ResultWool' class='blue' onchange='selectChange(this)'>"; // ResultWool
                foreach (var val in FabricColorFastnessModel.Result_Source)
                {
                    html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
                }
                html += "</select></td>";
            }

            html += "<td><input id='Details_" + i + "__Remark' name='Details[" + i + "].Remark' type='text'></td>"; // remark

            html += "<td></td>"; // LastUpdate

            html += $@"<td><div style=""width:5vw;""><img class=""detailDelete"" src=""/Image/Icon/Delete.png"" width=""30"" /></div></td>";
            html += "</tr>";

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult DetailSave(Fabric_ColorFastness_Detail_ViewModel req)
        {
            if (req.Details == null)
            {
                req.Details = new List<Fabric_ColorFastness_Detail_Result>();
            }

            req.Main.TestBeforePicture = req.Main.TestBeforePicture == null ? null : ImageHelper.ImageCompress(req.Main.TestBeforePicture);
            req.Main.TestAfterPicture = req.Main.TestAfterPicture == null ? null : ImageHelper.ImageCompress(req.Main.TestAfterPicture);

            BaseResult result = _FabricColorFastness_Service.Save_ColorFastness_2ndPage(req, this.MDivisionID, this.UserID);
            if (result.Result)
            {
                // 找地方寫入 new ID
                req.Main.ID = string.IsNullOrEmpty(result.ErrorMessage) ? req.Main.ID : result.ErrorMessage;
                return RedirectToAction("Detail", new { POID = req.Main.POID, ID = req.Main.ID, EditMode = false });
            }

            req.Result = result.Result;
            req.ErrorMessage = result.ErrorMessage;
            TempData["ModelFabricColorFastness"] = req;
            ViewBag.UserMail = this.UserMail;
            return RedirectToAction("Detail", new { POID = req.Main.POID, ID = req.Main.ID, EditMode = true });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode_Detail(string ID, FabricColorFastness_Service.DetailStatus Type)
        {
            string ColorFastnessResult = string.Empty;
            Fabric_ColorFastness_Detail_ViewModel result = _FabricColorFastness_Service.Encode_ColorFastness(ID, Type, this.UserID);
            ColorFastnessResult = result.sentMail ? "Fail" : string.Empty;
            return Json(new { result.Result, result.ErrorMessage, ColorFastnessResult });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Report(string ID, bool IsToPDF)
        {
            Fabric_ColorFastness_Detail_ViewModel result;
            if (IsToPDF)
            {
                result = _FabricColorFastness_Service.ToReport(ID, true);
            }
            else
            {
                result = _FabricColorFastness_Service.ToReport(ID, false);
            }

            string FileName = result.reportPath;
            string reportPath = "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        public JsonResult SendMail(string POID, string ID, string TestNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            BaseResult result = _FabricColorFastness_Service.SentMail(POID, ID, TestNo, TO, CC, Subject, Body, Files);
            return Json(result);
        }
    }
}