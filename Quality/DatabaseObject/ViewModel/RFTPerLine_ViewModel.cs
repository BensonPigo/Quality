using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class RFTPerLine_ViewModel
    {
        public Dictionary<string, int>  Months { get; set; }

        public List<string> Years { get; set; }

        public List<MonthlyRFT> monthlyRFTs { get; set; }

        public List<DailyRFT> dailyRFTs { get; set; }
    }

    public class MonthlyRFT
    {
        public string Month { get; set; }

        public string Line { get; set; }

        public decimal RFT { get; set; }
    }    

    public class DailyRFT
    {
        public int Date { get; set; }

        public string Month { get; set; }

        public string Line { get; set; }

        public decimal RFT { get; set; }
    }

    public class RFTPerLine_Request
    {
        public List<MonthlyRFT> monthlyRFTs { get; set; }
        public DataTable MonthlyData { get; set; }
        public DataTable DailyData { get; set; }
    }
}
