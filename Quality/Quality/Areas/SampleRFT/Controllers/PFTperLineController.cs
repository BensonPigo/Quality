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

        public ActionResult Index()
        {
            RFTPerLine_ViewModel rftPerLine = _RFTPerLineService.GetQueryPara();
            List<SelectListItem> FactoryList = new SetListItem().ItemListBinding(this.Factorys);
            List<SelectListItem> YearsList = new SetListItem().ItemListBinding(rftPerLine.Years); 
            List<SelectListItem> MonthList = new SetListItem().ItemListBinding(rftPerLine.Months);
           
            ViewBag.FactoryList = FactoryList;
            ViewBag.YearsList = YearsList;
            ViewBag.MonthList = MonthList;

            int Year = DateTime.Now.Year;
            int Month = DateTime.Now.Month;
            string MonthEn =  rftPerLine.Months.Where(x => x.Value == DateTime.Now.Month).FirstOrDefault().Key;

            RFTPerLine_Request result = GetData(this.FactoryID, Year.ToString(), MonthEn);

            ViewBag.Factory = this.FactoryID;
            ViewBag.Year = Year.ToString();
            ViewBag.Month = Month.ToString();
            return View(result);
        }

        [HttpPost]
        public ActionResult Index(string Factory,string Years,string Month)
        {
            RFTPerLine_ViewModel rftPerLine = _RFTPerLineService.GetQueryPara();
            List<SelectListItem> FactoryList = new SetListItem().ItemListBinding(this.Factorys);
            List<SelectListItem> YearsList = new SetListItem().ItemListBinding(rftPerLine.Years);
            List<SelectListItem> MonthList = new SetListItem().ItemListBinding(rftPerLine.Months);

            ViewBag.FactoryList = FactoryList;
            ViewBag.YearsList = YearsList;
            ViewBag.MonthList = MonthList;
            RFTPerLine_Request result = GetData(Factory, Years, Month);

            ViewBag.Factory = Factory;
            ViewBag.Year = Years;
            ViewBag.Month = Month;
            return View(result);
        }

        private RFTPerLine_Request GetData(string Factory, string Years, string Month)
        {
            RFTPerLine_ViewModel rftPerLineQuery = _RFTPerLineService.RFTPerLineQuery(Factory, Years, Month);
            RFTPerLine_Request reslut = new RFTPerLine_Request();

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
                var item = rftPerLineQuery.monthlyRFTs.AsEnumerable().Where(r => r.Line == line.ColumnName).FirstOrDefault();

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
                    var item = rftPerLineQuery.dailyRFTs.AsEnumerable().Where(r => r.Line == line.ColumnName && r.Date == date).FirstOrDefault();

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

            reslut.monthlyRFTs = rftPerLineQuery.monthlyRFTs;
            reslut.MonthlyData = monthlydt;
            reslut.DailyData = dailydt;

            return reslut;
        } 
    }
}