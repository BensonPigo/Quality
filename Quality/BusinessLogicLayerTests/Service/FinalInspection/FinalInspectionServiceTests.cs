using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ManufacturingExecutionDB;
using ProductionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel.FinalInspection;
using DatabaseObject;
using Newtonsoft.Json;
using DatabaseObject.ResultModel.FinalInspection;

namespace BusinessLogicLayer.Service.FinalInspectionTests
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

                string result = JsonConvert.SerializeObject(finalInspectionService.GetPivot88Json("ES3CH21110011"));
                
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
                FinalInspectionService finalInspectionService = new FinalInspectionService();
                PivotTransferRequest pivotTransferRequest = new PivotTransferRequest()
                { 
                    InspectionType = "Final Inspection",
                    BaseUri = "https://adidasstage4.pivot88.com",
                    RequestUri = "rest/operation/v1/inspection_reports/unique_key:"
                };

                pivotTransferRequest.Headers.Add("api-key", "fc16972a-bdc9-420f-9b6e-e7193ed99508");

                List<SentPivot88Result> sentPivot88Results = finalInspectionService.SentPivot88(pivotTransferRequest);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetFinalInspectionTest()
        {
            try
            {
                FinalInspectionService finalInspectionService = new FinalInspectionService();

                DatabaseObject.ManufacturingExecutionDB.FinalInspection result = finalInspectionService.GetFinalInspection("123");

                if (!result.Result)
                {
                    Assert.Fail(result.ErrorMessage);
                }
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetOrderForInspectionTest()
        {
            try
            {
                FinalInspectionService finalInspectionService = new FinalInspectionService();

                FinalInspection_Request inspection_Request = new FinalInspection_Request()
                {
                    FactoryID = "ESP",
                    StyleID = "LW6BM7S",
                    SP = "21040034IC002",
                    CustPONO = "21040034IC"
                };

                IList<Orders> result = finalInspectionService.GetOrderForInspection(inspection_Request);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void UpdateFinalInspectionByStepTest()
        {
            try
            {
                IFinalInspectionService finalInspectionService = new FinalInspectionService();

                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection = finalInspectionService.GetFinalInspection("ESPCH21080001");

                finalInspection.FabricApprovalDoc = true;
                finalInspection.GarmentWashingDoc = true;
                finalInspection.MetalDetectionDoc = true;
                finalInspection.SealingSampleDoc = true;
                finalInspection.InspectionStep = "Insp-General";
                BaseResult result = finalInspectionService.UpdateFinalInspectionByStep(finalInspection, "Insp-General", "SCIMIS");

                if (!result)
                {
                    Assert.Fail(result.ErrorMessage);
                }


                finalInspection.CheckCloseShade = true;
                finalInspection.CheckHandfeel = true;
                finalInspection.CheckAppearance = true;
                finalInspection.CheckPrintEmbDecorations = true;
                finalInspection.CheckFiberContent = true;
                finalInspection.CheckCareInstructions = true;
                finalInspection.CheckDecorativeLabel = true;
                finalInspection.CheckAdicomLabel = true;
                finalInspection.CheckCountryofOrigion = true;
                finalInspection.CheckSizeKey = true;
                finalInspection.Check8FlagLabel = true;
                finalInspection.CheckAdditionalLabel = true;
                finalInspection.CheckShippingMark = true;
                finalInspection.CheckPolytagMarketing = true;
                finalInspection.CheckColorSizeQty = true;
                finalInspection.CheckHangtag = true;
                finalInspection.InspectionStep = "Insp-CheckList";

                result = finalInspectionService.UpdateFinalInspectionByStep(finalInspection, "Insp-CheckList", "SCIMIS");

                if (!result)
                {
                    Assert.Fail(result.ErrorMessage);
                }

                finalInspection.InspectionStep = "Insp-Moisture";
                result = finalInspectionService.UpdateFinalInspectionByStep(finalInspection, "Insp-Moisture", "SCIMIS");

                if (!result)
                {
                    Assert.Fail(result.ErrorMessage);
                }

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}