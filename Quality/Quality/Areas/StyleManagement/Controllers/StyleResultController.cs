using BusinessLogicLayer.Service.StyleManagement;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.SampleRFT;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.StyleManagement.Controllers
{
    public class StyleResultController : BaseController
    {
        private StyleResultService _Service;

        public StyleResultController()
        {
            _Service = new StyleResultService();
        }
        // GET: StyleManagement/StyleResult
        public ActionResult Index()
        {
            StyleResult_ViewModel model = new StyleResult_ViewModel()
            {
                SampleRFT = new List<StyleResult_SampleRFT>(),
                FTYDisclamier = new List<StyleResult_FTYDisclamier>() { new StyleResult_FTYDisclamier() },
                RRLR = new List<StyleResult_RRLR>() { new StyleResult_RRLR() },
                BulkFGT = new List<StyleResult_BulkFGT>() { new StyleResult_BulkFGT() }
            };
            return View(model);
        }

        public ActionResult BarcodeScan()
        {
            this.CheckSession();

            return View();
        }

        [HttpPost]
        public ActionResult Query(StyleResult_Request Req)
        {
            this.CheckSession();
            StyleResult_ViewModel model = new StyleResult_ViewModel();
            if (Req == null || string.IsNullOrEmpty(Req.StyleID) || string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID))
            {
                model.MsgScript = $@"
msg.WithInfo('Style, Brand and Season cannot be empty.');
";
                return View("Index", model);
            }

            model = _Service.Get_StyleResult_Browse(Req);
            return View("Index", model);
        }


        public ActionResult DownLoadFDFile(string FDFile)
        {
            string fileName = FDFile;

            MemoryStream obj_stream = new MemoryStream();
            var tempFile = FDFile;
            obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(tempFile));
            Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
            Response.BinaryWrite(obj_stream.ToArray());
            obj_stream.Close();
            obj_stream.Dispose();
            Response.Flush();
            Response.End();
            return null;
        }
    }
}