using BusinessLogicLayer.Service;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel.FinalInspection;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;

namespace Quality.Controllers
{
    public class TransferFinalInspectionController : ApiController
    {
        [Route("api/TransferFinalInspection/SentPivot88")]
        [HttpPost]
        public HttpResponseMessage SentPivot88(PivotTransferRequest pivotTransferRequest)
        {
            try
            {
                FinalInspectionService finalInspectionService = new FinalInspectionService();
                List<SentPivot88Result> sentPivot88Results = finalInspectionService.SentPivot88(pivotTransferRequest);
                return Request.CreateResponse(HttpStatusCode.OK, sentPivot88Results);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.ExpectationFailed, new { result = ex.ToString() });
            }
        }

        [Route("api/TransferFinalInspection/Imp_EOLInlineInspectionReport")]
        [HttpPost]
        public HttpResponseMessage Imp_EOLInlineInspectionReport()
        {
            try
            {
                FinalInspectionService finalInspectionService = new FinalInspectionService();
                finalInspectionService.ExecImp_EOLInlineInspectionReport();
                return Request.CreateResponse(HttpStatusCode.OK);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.ExpectationFailed, new { result = ex.ToString() });
            }
        }
    }
}
