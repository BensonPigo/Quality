using BusinessLogicLayer.Interface;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.StyleManagement
{
    public class PrintBarcodeService
    {
        public PrintBarcodeProvider _Provider { get; set; }


        public List<PrintBarcode_Detail> Get_PrintBarcodeStyleInfo(StyleManagement_Request styleResult_Request)
        {
            try
            {
                _Provider = new PrintBarcodeProvider(Common.ProductionDataAccessLayer);

                styleResult_Request.CallType = StyleManagement_Request.EnumCallType.StyleResult;
                List<PrintBarcode_Detail> result = _Provider.Get_StyleInfo(styleResult_Request).ToList();

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<SelectListItem> Get_SampleStage(StyleManagement_Request styleResult_Request)
        {
            try
            {
                _Provider = new PrintBarcodeProvider(Common.ProductionDataAccessLayer);

                styleResult_Request.CallType = StyleManagement_Request.EnumCallType.StyleResult;
                List<SelectListItem> result = _Provider.Get_SampleStage(styleResult_Request).ToList();

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
