using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTPerLineProvider
    {
        IList<MonthlyRFT> GetMonthlyRFT(string FactoryID, string Year, int monthint);

        IList<DailyRFT> GetDailyRFT(string FactoryID, string Year, int monthint);
    }
}
