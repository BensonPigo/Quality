using DatabaseObject.ViewModel.BulkFGT;
using ProductionDataAccessLayer.Provider.MSSQL.BukkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class AccessoryOvenWashService
    {
        private AccessoryOvenWashProvider _AccessoryOvenWashProvider;

        public Accessory_ViewModel GetMainData(Accessory_ViewModel Req)
        {
            Accessory_ViewModel result = new Accessory_ViewModel()
            {
                ReqOrderID = Req.ReqOrderID,
                DataList = new List<Accessory_Result>(),
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetHead(Req.ReqOrderID);
                result.DataList = _AccessoryOvenWashProvider.GetDetail(Req.ReqOrderID).ToList();

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }

            return result;
        }

        public Accessory_ViewModel Update(Accessory_ViewModel Req)
        {
            Accessory_ViewModel result = new Accessory_ViewModel()
            {
                ReqOrderID = Req.ReqOrderID,
                DataList = new List<Accessory_Result>(),
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                int r = _AccessoryOvenWashProvider.Update(Req);

                result.Result = r > 0;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }

            return result;
        }
    }
}
