using DatabaseObject.RequestModel;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service
{
    public class MailToolsService
    {
        private MailToolsProvider _Provider;

        public string GetAICommet(SendMail_Request Req)
        {

            _Provider = new MailToolsProvider(Common.ProductionDataAccessLayer);

            string comment = _Provider.GetAIComment(Req);

            return comment;
        }
        public string GetBuyReadyDate(SendMail_Request Req)
        {

            _Provider = new MailToolsProvider(Common.ProductionDataAccessLayer);

            string comment = _Provider.GetBuyReadyDate(Req);

            return comment;
        }
    }
}
