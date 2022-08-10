using BusinessLogicLayer.Interface.StyleManagement;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.StyleManagement
{
    public class StyleListService : IStyleListService
    {
        public IStyleListProvider _StyleListProvider { get; set; }
        public IStyleListProvider _MES_StyleListProvider { get; set; }


        public StyleList Get_StyleInfo(StyleList_Request Req)
        {
            StyleList result = new StyleList();
            try
            {
                _MES_StyleListProvider = new StyleListProvider(Common.ManufacturingExecutionDataAccessLayer);
                //資料來源在Trade DB
                _StyleListProvider = new StyleListProvider(Common.ProductionDataAccessLayer);

                bool HasSampleRFTInspection = _MES_StyleListProvider.Check_SampleRFTInspection_Count(Req) > 0 ? true : false;
                Req.InspectionTableName = HasSampleRFTInspection ? "SampleRFTInspection" : "RFT_Inspection";
                List <StyleList> datalist = _StyleListProvider.Get_StyleInfo(Req).ToList();
                result.DataList = datalist;
                result.Result = true;

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            return result;
        }

    }
}
