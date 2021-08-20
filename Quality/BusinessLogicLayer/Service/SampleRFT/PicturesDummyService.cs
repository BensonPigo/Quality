using ADOHelper.Template.MSSQL;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ToolKit;
using ManufacturingExecutionDataAccessLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using System.Linq;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ResultModel;
using BusinessLogicLayer.SampleRFT.Interface;
using BusinessLogicLayer.Interface.SampleRFT;

namespace BusinessLogicLayer.Service.SampleRFT 
{
    public class PicturesDummyService : IPicturesDummyService
    {
        public IRFTPicDuringDummyFittingProvider _RFTPicDuringDummyFittingProvider { get; set; }

        public RFT_PicDuringDummyFitting_ViewModel Get_PicturesDummy_Result(RFT_PicDuringDummyFitting_ViewModel Req)
        {
            RFT_PicDuringDummyFitting_ViewModel result = new RFT_PicDuringDummyFitting_ViewModel();

            try
            {
                _RFTPicDuringDummyFittingProvider = new RFTPicDuringDummyFittingProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<RFT_PicDuringDummyFitting> list = _RFTPicDuringDummyFittingProvider.Get_PicDuringDummy_Result(Req).ToList();

                result.DataList = list;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public RFT_PicDuringDummyFitting_ViewModel Check_OrderID_Exists(string orderID)
        {
            RFT_PicDuringDummyFitting_ViewModel result = new RFT_PicDuringDummyFitting_ViewModel();

            try
            {
                _RFTPicDuringDummyFittingProvider = new RFTPicDuringDummyFittingProvider(Common.ManufacturingExecutionDataAccessLayer);

                var m = _RFTPicDuringDummyFittingProvider.Check_OrderID_Exists(orderID);

                if (m.Any())
                {
                    result = m.FirstOrDefault();
                    result.Result = true;
                }

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
