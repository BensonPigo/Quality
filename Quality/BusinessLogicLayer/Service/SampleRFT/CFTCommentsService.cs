using ADOHelper.Template.MSSQL;
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
using BusinessLogicLayer.Interface.SampleRFT;

namespace BusinessLogicLayer.Service.SampleRFT
{
    public class CFTCommentsService : ICFTCommentsService
    {
        private ICFTCommentsProvider _CFTCommentsProvider;

        public CFTComments_ViewModel Get_CFT_Orders(CFTComments_ViewModel Req)
        {
            CFTComments_ViewModel model = new CFTComments_ViewModel();
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ProductionDataAccessLayer);
                var res = _CFTCommentsProvider.Get_CFT_Orders(Req).ToList();

                if (res.Any())
                {
                    model = res.FirstOrDefault();
                }
                
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public CFTComments_ViewModel Get_CFT_OrderComments(CFTComments_ViewModel Req)
        {
            Req.DataList = new List<CFTComments_Result>();
            CFTComments_ViewModel model = Req;

            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
                var res = _CFTCommentsProvider.Get_CFT_OrderComments(Req).ToList();

                model.DataList = res;
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }
    }
}
