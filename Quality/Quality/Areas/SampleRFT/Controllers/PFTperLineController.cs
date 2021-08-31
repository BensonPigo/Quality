using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using Newtonsoft.Json;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class PFTperLineController : BaseController
    {
        private IRFTPerLineService _RFTPerLineService;

        public PFTperLineController()
        {
            _RFTPerLineService = new RFTPerLineService();
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.PFTperLine,,";
        }

        // GET: SampleRFT/PFTperLine
        public ActionResult Index(string Factory,string Years,string Month)
        {

            RFTPerLine_ViewModel rftPerLine = _RFTPerLineService.GetQueryPara();

            List<SelectListItem> FactoryList = new SetListItem().ItemListBinding(this.Factorys);
            ViewBag.FactoryList = FactoryList;

            List<SelectListItem> YearsList = new SetListItem().ItemListBinding(rftPerLine.Years);
            ViewBag.YearsList = YearsList;

            Dictionary<string, object> months = new Dictionary<string, object>();

            foreach (var month in rftPerLine.Months)
            {
                months.Add(month.Key, month.Value);
            }

            List<SelectListItem> MonthList = new SetListItem().ItemListBinding(months);
            ViewBag.MonthList = MonthList;

            RFTPerLine_ViewModel rftPerLineQuery = _RFTPerLineService.RFTPerLineQuery(this.FactoryID, rftPerLine.Years.FirstOrDefault(), rftPerLine.Months.FirstOrDefault().Key.ToString());

            rftPerLine.monthlyRFTs = rftPerLineQuery.monthlyRFTs;
            rftPerLine.dailyRFTs = rftPerLineQuery.dailyRFTs;

            DataTable monthlydt = new DataTable();

            monthlydt.Columns.Add("Month", typeof(string));

            List<string> dtMonthHeader = rftPerLineQuery.monthlyRFTs.AsEnumerable().Select(r => r.Line).Distinct().ToList();

            foreach (string line in dtMonthHeader.OrderBy(r => r))
            {
                monthlydt.Columns.Add(line, typeof(string));
            }

            DataRow monthlyRow = monthlydt.NewRow();

            bool fristTime = true;

            foreach (DataColumn line in monthlydt.Columns)
            {
                var item = rftPerLine.monthlyRFTs.AsEnumerable().Where(r => r.Line == line.ColumnName).FirstOrDefault();

                if (item != null)
                {
                    if (fristTime)
                    {
                        monthlyRow["Month"] = item.Month;
                        fristTime = false;
                    }
                    monthlyRow[line.ColumnName] = item.RFT;
                }
            }

            monthlydt.Rows.Add(monthlyRow);


           DataTable dailydt = new DataTable();

            List<string> dtHeader = rftPerLineQuery.dailyRFTs.AsEnumerable().Select(r => r.Line).Distinct().ToList();

            dailydt.Columns.Add("Date", typeof(string));

            foreach (string line in dtHeader.OrderBy(r => r))
            {
                dailydt.Columns.Add(line, typeof(string));
            }

            List<int> dailylist = rftPerLineQuery.dailyRFTs.AsEnumerable().Select(r => r.Date).Distinct().ToList();


            foreach (int date in dailylist)
            {
                bool dateFistTime = true;
                DataRow row = dailydt.NewRow();

                foreach (DataColumn line in dailydt.Columns)
                {
                    var item = rftPerLine.dailyRFTs.AsEnumerable().Where(r => r.Line == line.ColumnName && r.Date == date).FirstOrDefault();

                    if (item != null)
                    {
                        if (dateFistTime)
                        {
                            row["Date"] = date + "-" + item.Month;
                            dateFistTime = false;
                        }

                        row[line.ColumnName] = item.RFT;
                    }
                }

                dailydt.Rows.Add(row);

            }

            string dailyjson = JsonConvert.SerializeObject(dailydt);
            dynamic data = new ExpandoObject();
            data.rftPerLine = rftPerLine;
            data.MonthlyData = monthlydt;
            data.DailyData = dailydt;
            return View(data);
        }

    }
}