using DatabaseObject.ViewModel.FinalInspection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.FinalInspection
{
    public class QueryService
    {
        //寄信
        public bool SendMail (DatabaseObject.ManufacturingExecutionDB.FinalInspection Req)
        {

            //透過Req從後端取得資料
            QueryReport data = new QueryReport();
            
            //開始寫信

            return true;
        }
        
    }
}
