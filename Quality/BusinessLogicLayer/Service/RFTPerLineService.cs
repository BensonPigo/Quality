using ADOHelper.Template.MSSQL;
using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ToolKit;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;

namespace BusinessLogicLayer.Service
{
    public class RFTPerLineService : IRFTPerLineService
    {
        private IRFTPerLineProvider _RFTPerLineProvider;

        public RFTPerLine_ViewModel GetQueryPara()
        {
            RFTPerLine_ViewModel rFTPerLine_ViewModel = new RFTPerLine_ViewModel();

            rFTPerLine_ViewModel.Months = new Dictionary<string, int>
            {
                { "January", 1 },
                { "February", 2 },
                { "March", 3 },
                { "April", 4 },
                { "May", 5 },
                { "June", 6 },
                { "July", 7 },
                { "August", 8 },
                { "September", 9 },
                { "October", 10 },
                { "November", 11 },
                { "December", 12 }
            };

            rFTPerLine_ViewModel.Years = new List<string>();
            int nowYear = DateTime.Now.Year;
            for (int i = 2021; i <= nowYear; i++)
            {
                rFTPerLine_ViewModel.Years.Add(i.ToString());
            }


            return rFTPerLine_ViewModel;
        }

        public RFTPerLine_ViewModel RFTPerLineQuery(string FactoryID, string Year, string Month)
        {
            RFTPerLine_ViewModel rFTPerLine_ViewModel = new RFTPerLine_ViewModel
            {
                monthlyRFTs = new List<MonthlyRFT>(),
                dailyRFTs = new List<DailyRFT>()
            };

            int monthint;
            if (!int.TryParse(Month, out monthint))
            {
                switch (Month)
                {
                    case "January":
                        monthint = 1;
                        break;
                    case "February":
                        monthint = 2;
                        break;
                    case "March":
                        monthint = 3;
                        break;
                    case "April":
                        monthint = 4;
                        break;
                    case "May":
                        monthint = 5;
                        break;
                    case "June":
                        monthint = 6;
                        break;
                    case "July":
                        monthint = 7;
                        break;
                    case "August":
                        monthint = 8;
                        break;
                    case "September":
                        monthint = 9;
                        break;
                    case "October":
                        monthint = 10;
                        break;
                    case "November":
                        monthint = 11;
                        break;
                    case "December":
                        monthint = 12;
                        break;
                }
            }

            _RFTPerLineProvider = new RFTPerLineProvider(Common.ManufacturingExecutionDataAccessLayer);
            foreach (var item in _RFTPerLineProvider.GetMonthlyRFT(FactoryID, Year, monthint))
            {
                rFTPerLine_ViewModel.monthlyRFTs.Add(item);
            }

            foreach (var item in _RFTPerLineProvider.GetDailyRFT(FactoryID, Year, monthint))
            {
                rFTPerLine_ViewModel.dailyRFTs.Add(item);
            }

            return rFTPerLine_ViewModel;
        }
    }
}
