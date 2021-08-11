using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Quality.Areas.Bulk_FGT.Controllers
{
    public class GarmentTestController : BaseController
    {
        public GarmentTestController()
        {
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.GarmentTest,,";
        }
    }
}