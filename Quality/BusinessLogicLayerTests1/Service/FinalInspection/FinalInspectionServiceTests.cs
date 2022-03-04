using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using DatabaseObject.RequestModel;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System.IO;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class FinalInspectionServiceTests
    {
        [TestMethod()]
        public void GetPivot88JsonTest()
        {
            try
            {
                FinalInspectionService finalInspectionService = new FinalInspectionService();
                string result = JsonConvert.SerializeObject(finalInspectionService.GetPivot88Json("ES3CH21110007"));
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void SentPivot88Test()
        {
            try
            {
                //FinalInspectionProvider _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                //Dictionary<string, byte[]> dicImage = _FinalInspectionProvider.GetFinalInspectionDefectImage("ES3CH21100003");

                //System.Drawing.Image.FromStream(new MemoryStream(dicImage["ES3CH21100003_C01_2.png"])).Save("d:\\aa\\dsd.png");

                FinalInspectionService finalInspectionService = new FinalInspectionService();
                PivotTransferRequest pivotTransferRequest = new PivotTransferRequest()
                {
                    InspectionID = "ES3CH21110007",
                    BaseUri = "https://adidasstage4.pivot88.com",
                    RequestUri = "rest/operation/v1/inspection_reports/unique_key:",
                    Headers = new Dictionary<string, string>() { { "api-key", "64158338-5de2-451e-aa72-3fa470fdf4cb" } }
                };
                finalInspectionService.SentPivot88(pivotTransferRequest);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}