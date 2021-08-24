using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas
{
    public class FileActionController : Controller
    {
        // GET: FileAction
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult FileUpload(HttpPostedFileBase file)
        {
            string fileName;
            if (file.ContentLength > 0)
            {
                string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);
                string extension = Path.GetExtension(file.FileName);
                fileName = FileNameWithoutExtension + Guid.NewGuid().ToString("N") + extension;
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);
                file.SaveAs(path);
                // 回傳新檔名, 以供其它操作使用, 例如:SendMail 附件使用檔名存取 Server 上的實體檔案
                return RedirectToAction(fileName);
            }
            else
            {
                return RedirectToAction("No File");
            }
        }
    }
}