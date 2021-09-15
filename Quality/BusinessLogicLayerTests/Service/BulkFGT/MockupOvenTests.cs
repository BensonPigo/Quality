using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class MockupOvenTests
    {
        [TestMethod()]
        public void GetMockupOven()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                MockupOven_Request MockupOven = new MockupOven_Request()
                //{ ReportNo = "PHOV210900008" };
                { BrandID = "ADIDAS", SeasonID = "20SS", StyleID = "S201CSPM108", Article = "FL0237" };
                var mockupOven = _MockupOvenService.GetMockupOven(MockupOven);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetAccessoryRefNo()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                AccessoryRefNo_Request accessoryRefNo_Request = new AccessoryRefNo_Request()
                {
                    BrandID = "U.ARMOUR",
                    SeasonID = "20FW",
                    StyleID = "1342962",
                    //StyleUkey =75468,
                };
                var selectListItems = _MockupOvenService.GetAccessoryRefNo(accessoryRefNo_Request);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetArtworkTypeID()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                StyleArtwork_Request StyleArtwork_Request = new StyleArtwork_Request()
                {
                    BrandID = "U.ARMOUR",
                    SeasonID = "20FW",
                    StyleID = "1342962",
                    //StyleUkey =75468,
                };
                var selectListItems = _MockupOvenService.GetArtworkTypeID(StyleArtwork_Request);
                Assert.IsTrue(selectListItems.Count == 1);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetOrders()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                var orders = _MockupOvenService.GetOrders(new Orders() { ID = "21090101IE022" });
                Assert.IsTrue(orders.Count == 1);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetDistinctArticle()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                Order_Qty Order_Qty = new Order_Qty()
                {
                    ID = "21080085IE044",
                };
                var x = _MockupOvenService.GetDistinctArticle(Order_Qty);
                Assert.IsTrue(x.Count > 0 && x[0].Article == "0001");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void Create()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                MockupOven_ViewModel MockupOven = new MockupOven_ViewModel()
                {
                    POID = "TP",
                    StyleID = "SS",
                    SeasonID = "20SS",
                    BrandID = "bb",
                    Article = "aa",
                    ArtworkTypeID = "attt",
                    Remark = "remmmm",
                    T1Subcon = "SCIMIS",
                    T2Supplier = "SCIMIS",
                    TestDate = DateTime.Now,
                    ReceivedDate = DateTime.Now,
                    ReleasedDate = DateTime.Now,
                    Result = "Pass",
                    Technician = "SCIMIS",
                    MR = "SCIMIS",
                    AddName = "SCIMIS",
                    TestTemperature=(decimal)66.66,
                    TestTime= (decimal)6.66,
                    HTPlate = 9,
                    HTFlim = 6,
                    HTTime = 3,
                    HTPressure = (decimal)22.1,
                    HTPellOff = "POFF",
                    HT2ndPressnoreverse = 2,
                    HT2ndPressreversed = 3,
                    HTCoolingTime = 6,
                    MockupOven_Detail = new System.Collections.Generic.List<MockupOven_Detail_ViewModel>()
                    {
                        new MockupOven_Detail_ViewModel(){TypeofPrint="TTTT1",Design="d100",ArtworkColor="0001",FabricRefNo="RF",AccessoryRefno="AF",FabricColor  = "FCC",Result="Pass",Remark="RRRK",EditName = "SCIMIS"},
                    }
                };

                var mockupOven = _MockupOvenService.Create(MockupOven, "VM1", out string no);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void Update()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                MockupOven_ViewModel MockupOven = new MockupOven_ViewModel()
                {
                    ReportNo = "T1",
                    POID = "TP",
                    StyleID = "SS",
                    SeasonID = "20SS",
                    BrandID = "bb",
                    Article = "aa",
                    ArtworkTypeID = "attt",
                    Remark = "remmmm",
                    T1Subcon = "SCIMIS",
                    T2Supplier = "SCIMIS",
                    TestDate = DateTime.Now,
                    ReceivedDate = DateTime.Now,
                    ReleasedDate = DateTime.Now,
                    Result = "Pass",
                    Technician = "SCIMIS",
                    MR = "SCIMIS",
                    AddName = "SCIMIS",
                    TestTemperature = (decimal)66.66,
                    TestTime = (decimal)6.66,
                    HTPlate = 9,
                    HTFlim = 6,
                    HTTime = 3,
                    HTPressure = (decimal)22.1,
                    HTPellOff = "POFF",
                    HT2ndPressnoreverse = 2,
                    HT2ndPressreversed = 3,
                    HTCoolingTime = 6,
                    MockupOven_Detail = new System.Collections.Generic.List<MockupOven_Detail_ViewModel>()
                    {
                        new MockupOven_Detail_ViewModel(){ReportNo = "T1",TypeofPrint="TTTT8",Design="d100",ArtworkColor="0001",FabricRefNo="RF",AccessoryRefno="AF",FabricColor  = "FCC",Result="Pass",Remark="RRR7K",EditName = "SCIMIS",Ukey = 18653},
                        new MockupOven_Detail_ViewModel(){ReportNo = "T1",TypeofPrint="TTT29",Design="d200",ArtworkColor="0001",FabricRefNo="RF",AccessoryRefno="AF",FabricColor  = "FCC",Result="Fail",Remark="2RRRK",EditName = "SCIMIS",Ukey=18662},
                    }
                    //ReportNo = "T1",
                    //POID = "up",
                    //StyleID = "up",
                    //SeasonID = "up",
                    //BrandID = "up",
                    //Article = "up",
                    //ArtworkTypeID = "up",
                    //Remark = "up",
                    //T1Subcon = "up",
                    //T2Supplier = "up",
                    //TestDate = DateTime.Now,
                    //ReceivedDate = DateTime.Now,
                    //ReleasedDate = DateTime.Now,
                    //Result = "up",
                    //Technician = "up",
                    //MR = "up",
                    //EditName = "up",
                    //TestTemperature = (decimal)99.99,
                    //TestTime = (decimal)99.99,

                    //HTPlate = 1,
                    //HTFlim = 1,
                    //HTTime = 1,
                    //HTPressure = (decimal)11.1,
                    //HTPellOff = "up",
                    //HT2ndPressnoreverse = 1,
                    //HT2ndPressreversed = 1,
                    //HTCoolingTime = 77,
                    //MockupOven_Detail = new System.Collections.Generic.List<MockupOven_Detail_ViewModel>()
                    //{
                    //    new MockupOven_Detail_ViewModel(){TypeofPrint="up",Design="up",ArtworkColor="up",FabricRefNo="up",AccessoryRefno="up",FabricColor  = "up",Result="up",Remark="up",EditName = "up",Ukey=18617},
                    //    new MockupOven_Detail_ViewModel(){TypeofPrint="up",Design="up",ArtworkColor="up",FabricRefNo="up",AccessoryRefno="up",FabricColor  = "up",Result="up",Remark="up",EditName = "up",Ukey=18618},
                    //}
                };

                var mockupOven = _MockupOvenService.Update(MockupOven);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void Delete()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                MockupOven_ViewModel MockupOven = new MockupOven_ViewModel() { ReportNo = "T1" };
                var mockupOven = _MockupOvenService.Delete(MockupOven);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }

        }

        [TestMethod()]
        public void DeleteDetail()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                System.Collections.Generic.List<MockupOven_Detail_ViewModel> MockupOven_Detail = new System.Collections.Generic.List<MockupOven_Detail_ViewModel>()
                {
                    //new MockupOven_Detail_ViewModel(){Ukey=21858},
                    new MockupOven_Detail_ViewModel(){Ukey=18618},
                };

                var mockupOven = _MockupOvenService.DeleteDetail(MockupOven_Detail);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetPDF()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                MockupOven_Request MockupOven = new MockupOven_Request()
                { ReportNo = "PHOV210900008" };
                //{ BrandID = "ADIDAS", SeasonID = "20SS", StyleID = "S201CSPM108", Article = "FL0237" };
                var mockupOven = _MockupOvenService.GetMockupOven(MockupOven);
                _MockupOvenService.GetPDF(mockupOven, true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void FailSendMail()
        {
            try
            {
                IMockupOvenService _MockupOvenService = new MockupOvenService();
                MockupFailMail_Request MockupOven = new MockupFailMail_Request()
                {
                    ReportNo = "T1",
                    To = "jeff.yeh@sportscity.com.tw",
                };

                var mockupOven = _MockupOvenService.FailSendMail(MockupOven);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}