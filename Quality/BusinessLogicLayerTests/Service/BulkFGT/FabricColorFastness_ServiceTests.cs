using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface.BulkFGT;
using MICS.DataAccessLayer.Interface;
using MICS.DataAccessLayer.Provider.MSSQL;
using DatabaseObject.ViewModel.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using static BusinessLogicLayer.Service.BulkFGT.FabricColorFastness_Service;
using DatabaseObject.ManufacturingExecutionDB;

namespace BusinessLogicLayer.Service.BulkFGT.Tests
{
    [TestClass()]
    public class FabricColorFastness_ServiceTests
    {
        [TestMethod()]
        public void Get_ScalesTest()
        {
            //try
            //{
            //    IFabricColorFastness_Service service = new FabricColorFastness_Service();
            //    _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            //    var scales = _IColorFastnessProvider.GetScales();
            //    Assert.IsTrue(scales.Count > 0);
            //}
            //catch (Exception ex)
            //{
            //    Assert.Fail(ex.ToString());
            //}
        }

        [TestMethod()]
        public void Get_MainTest()
        {
            try
            {
                IFabricColorFastness_Service service = new FabricColorFastness_Service();
                var result = service.Get_Main("21041712BB");

                Assert.IsTrue(result.Result);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetDetailHeaderTest()
        {
            try
            {
                IFabricColorFastness_Service service = new FabricColorFastness_Service();
                var result = service.GetDetailHeader("VM2CF21030110");

                Assert.IsTrue(result.baseResult.Result);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetDetailBodyTest()
        {
            try
            {
                IFabricColorFastness_Service service = new FabricColorFastness_Service();
                var result = service.GetDetailBody("VM2CF21030110");

                Assert.IsTrue(result.Detail.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetSeqTest()
        {
            try
            {
                IFabricColorFastness_Service service = new FabricColorFastness_Service();
                var result = service.GetSeq("21041712BB", "", "");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetRollTest()
        {
            try
            {
                IFabricColorFastness_Service service = new FabricColorFastness_Service();
                var result = service.GetRoll("21041712BB", "02", "01");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void Save_ColorFastness_2ndPageTest()
        {
            var resultS = new FabricColorFastness_ViewModel();
            Assert.IsTrue(true);
        }

        [TestMethod()]
        public void Save_ColorFastness_1stPageTest()
        {
            try
            {
                IFabricColorFastness_Service service = new FabricColorFastness_Service();
                List<ColorFastness_Result> _ColorFastness = new List<ColorFastness_Result>();
                ColorFastness_Result s = new ColorFastness_Result
                {
                    ID = "VM2CF21090002",
                };

                _ColorFastness.Add(s);
                BaseResult result = service.Save_ColorFastness_1stPage("21041712BB", "Test", _ColorFastness);
                Assert.IsTrue(result.Result);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void Save_ColorFastness_2ndPageTest1()
        {
            try
            {
                IFabricColorFastness_Service service = new FabricColorFastness_Service();
                BaseResult result = new BaseResult();
                Fabric_ColorFastness_Detail_ViewModel ss = new Fabric_ColorFastness_Detail_ViewModel();
                ColorFastness_Result s1 = new ColorFastness_Result
                {
                    ID = "VM2CF21090002",
                    POID = "21041712BB",
                    Inspector = "SCIMIS",
                    InspDate = DateTime.Now,
                    Detergent = "Woolite",
                    Article = "h42514",
                    Temperature = 40,
                    Machine = "Top Load",
                    Cycle = 5,
                    Drying = "Line dry",
                };

                ss.Main = s1;

                Fabric_ColorFastness_Detail_Result s = new Fabric_ColorFastness_Detail_Result
                {
                    ID = "VM2CF21090002",
                    SubmitDate = DateTime.Now.Date,
                    ColorFastnessGroup = "02",
                    SEQ1 = "02",
                    SEQ2 = "01",
                    Roll = "0001",
                    Dyelot = "1010",
                    Refno = "62550918",
                    SCIRefno = "RB-62550918-F00001",
                    ColorID = "095A",
                    Result = "Pass",
                    changeScale = "1",
                    ResultChange = "Pass",
                    StainingScale = "1",
                    ResultStain = "1",
                    Remark = "test",
                };

                List<Fabric_ColorFastness_Detail_Result> sss = new List<Fabric_ColorFastness_Detail_Result>();
                sss.Add(s);

                ss.Detail = sss;
                result = service.Save_ColorFastness_2ndPage(ss, "VM2", "SCIMIS");

                Assert.IsTrue(result.Result);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void Encode_ColorFastnessTest()
        {
            IFabricColorFastness_Service service = new FabricColorFastness_Service();
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();
            Fabric_ColorFastness_Detail_ViewModel ss = new Fabric_ColorFastness_Detail_ViewModel();
            ColorFastness_Result s1 = new ColorFastness_Result
            {
                ID = "VM2CF21090002",
                POID = "21041712BB",
                Inspector = "SCIMIS",
                InspDate = DateTime.Now,
                Detergent = "Woolite",
                Article = "h42514",
                Temperature = 40,
                Machine = "Top Load",
                Cycle = 5,
                Drying = "Line dry",
            };

            ss.Main = s1;

            Fabric_ColorFastness_Detail_Result s = new Fabric_ColorFastness_Detail_Result
            {
                ID = "VM2CF21090002",
                SubmitDate = DateTime.Now.Date,
                ColorFastnessGroup = "02",
                SEQ1 = "02",
                SEQ2 = "01",
                Roll = "0001",
                Dyelot = "1010",
                Refno = "62550918",
                SCIRefno = "RB-62550918-F00001",
                ColorID = "095A",
                Result = "Pass",
                changeScale = "1",
                ResultChange = "Pass",
                StainingScale = "1",
                ResultStain = "1",
                Remark = "test",
            };

            List<Fabric_ColorFastness_Detail_Result> sss = new List<Fabric_ColorFastness_Detail_Result>();
            sss.Add(s);

            ss.Detail = sss;
            result = service.Encode_ColorFastness(ss, DetailStatus.Encode, "SCIMIS");

            Assert.IsTrue(result.Result == true);

        }

        [TestMethod()]
        public void SentMailTest()
        {
            IFabricColorFastness_Service service = new FabricColorFastness_Service();
            BaseResult result = new BaseResult();
            List<Quality_MailGroup> _MailGroups = new List<Quality_MailGroup>();
            Quality_MailGroup m1 = new Quality_MailGroup
            {
                FactoryID = "ESP",
                GroupName = "Group 1",
                ToAddress = "willy.wei@sportscity.com.tw",
                CcAddress = "jack.hsu@sportscity.com.tw",
            };

            Quality_MailGroup m2 = new Quality_MailGroup
            {
                FactoryID = "ESP",
                GroupName = "Group 2",
                ToAddress = "willy.wei@sportscity.com.tw",
                CcAddress = "jack.hsu@sportscity.com.tw",
            };

            _MailGroups.Add(m1);
            _MailGroups.Add(m2);

            result = service.SentMail("21041712BB", "VM2CF21090002", _MailGroups);


            Assert.IsTrue(result.Result);
        }

        [TestMethod()]
        public void ToPDFTest()
        {
            IFabricColorFastness_Service service = new FabricColorFastness_Service();
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();
            try
            {
                result = service.ToPDF("VM2CF21090002", true);
                Assert.IsTrue(result.Result == true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message.ToString());
            }
        }

        [TestMethod()]
        public void ToExcelTest()
        {
            IFabricColorFastness_Service service = new FabricColorFastness_Service();
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();
            try
            {
                result = service.ToExcel("VM2CF21090002", true);
                Assert.IsTrue(result.Result == true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message.ToString());
            }
        }
    }
}