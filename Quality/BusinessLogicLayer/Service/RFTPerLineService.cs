using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using ToolKit;

namespace BusinessLogicLayer.Service
{
    public class RFTPerLineService : IRFTPerLineService
    {
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

            rFTPerLine_ViewModel.Years = new List<string>() { "2021" };

            return rFTPerLine_ViewModel;
        }

        public RFTPerLine_ViewModel RFTPerLineQuery(string FactoryID, string Year, string Month)
        {
            RFTPerLine_ViewModel rFTPerLine_ViewModel = new RFTPerLine_ViewModel
            {
                monthlyRFTs = new List<MonthlyRFT>(),
                dailyRFTs = new List<DailyRFT>()
            };

            for (int i = 1; i <= 30; i++)
            {
                MonthlyRFT monthlyRFT = new MonthlyRFT() 
                { 
                    Month = "June", 
                    Line = ("0" + i.ToString()).Right(2),
                    RFT = i,
                };

                DailyRFT dailyRFT = new DailyRFT()
                {
                    Date = i,
                    Month = "June",
                    Line = ("0" + i.ToString()).Right(2),
                    RFT = i,
                };

                rFTPerLine_ViewModel.monthlyRFTs.Add(monthlyRFT);
                rFTPerLine_ViewModel.dailyRFTs.Add(dailyRFT);
            } 
            
            return rFTPerLine_ViewModel;
        }
    }
}
