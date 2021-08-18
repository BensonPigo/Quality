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
using System.Linq;

namespace BusinessLogicLayer.Service
{
    public class CFTCommentsService : ICFTCommentsService
    {
        private ICFTCommentsProvider _CFTCommentsProvider;

        public CFTComments_where GetCFT_Orders(CFTComments_where CFTComments)
        {
            _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
            CFTComments_where CFTCommentsOrders_ViewModel = _CFTCommentsProvider.GetCFT_Orders(CFTComments).FirstOrDefault();
            return CFTCommentsOrders_ViewModel;
        }

        public List<CFTComments_ViewModel> GetRFT_OrderComments(CFTComments_where CFTComments)
        {
            _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<CFTComments_ViewModel> CFTComments_ViewModel = _CFTCommentsProvider.GetRFT_OrderComments(CFTComments).ToList();
            return CFTComments_ViewModel;
        }
    }
}
