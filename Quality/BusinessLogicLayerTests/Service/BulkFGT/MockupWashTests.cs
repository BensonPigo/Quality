using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class MockupWashTests
    {
        [TestMethod()]
        public void GetMockupWash()
        {
            try
            {
                IMockupWashService _MockupWashService = new MockupWashService();
                MockupWash_Request MockupWash = new MockupWash_Request()
                { ReportNo = "PHWA200700089" };
                //{ BrandID = "ADIDAS", SeasonID = "20SS", StyleID = "S201CSPM108", Article = "FL0237" };
                var mockupWash = _MockupWashService.GetMockupWash(MockupWash);
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
                IMockupWashService _MockupWashService = new MockupWashService();
                AccessoryRefNo_Request accessoryRefNo_Request = new AccessoryRefNo_Request()
                {
                    BrandID = "U.ARMOUR",
                    SeasonID = "20FW",
                    StyleID = "1342962",
                    //StyleUkey =75468,
                };
                var selectListItems = _MockupWashService.GetAccessoryRefNo(accessoryRefNo_Request);
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
                IMockupWashService _MockupWashService = new MockupWashService();
                StyleArtwork_Request StyleArtwork_Request = new StyleArtwork_Request()
                {
                    BrandID = "U.ARMOUR",
                    SeasonID = "20FW",
                    StyleID = "1342962",
                    //StyleUkey =75468,
                };
                var selectListItems = _MockupWashService.GetArtworkTypeID(StyleArtwork_Request);
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
                IMockupWashService _MockupWashService = new MockupWashService();
                var orders = _MockupWashService.GetOrders(new Orders() { ID = "21090101IE022" });
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
                IMockupWashService _MockupWashService = new MockupWashService();
                Order_Qty Order_Qty = new Order_Qty()
                {
                    ID = "21080085IE044",
                };
                var x = _MockupWashService.GetDistinctArticle(Order_Qty);
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
                IMockupWashService _MockupWashService = new MockupWashService();
                MockupWash_ViewModel MockupWash = new MockupWash_ViewModel()
                {
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
                    OtherMethod = true,
                    //MethodID = "a",
                    TestingMethod = "TESTWTF",
                    HTPlate = 9,
                    HTFlim = 6,
                    HTTime = 3,
                    HTPressure = (decimal)22.1,
                    HTPellOff = "POFF",
                    HT2ndPressnoreverse = 2,
                    HT2ndPressreversed = 3,
                    HTCoolingTime = "66TT",
                    MockupWash_Detail = new System.Collections.Generic.List<MockupWash_Detail_ViewModel>()
                    {
                        new MockupWash_Detail_ViewModel(){TypeofPrint="TTTT1",Design="d100",ArtworkColor="0001",FabricRefNo="RF",AccessoryRefno="AF",FabricColor  = "FCC",Result="Pass",Remark="RRRK",EditName = "SCIMIS"},
                    }
                };

                var mockupWash = _MockupWashService.Create(MockupWash, "VM1", "SCIMIS", out string no);
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
                IMockupWashService _MockupWashService = new MockupWashService();
                MockupWash_ViewModel MockupWash = new MockupWash_ViewModel()
                {
                    ReportNo = "VM1WA21090002",
                    POID = "up",
                    StyleID = "up",
                    SeasonID = "up",
                    BrandID = "up",
                    Article = "up",
                    ArtworkTypeID = "up",
                    Remark = "up",
                    T1Subcon = "up",
                    T2Supplier = "up",
                    TestDate = DateTime.Now,
                    ReceivedDate = DateTime.Now,
                    ReleasedDate = DateTime.Now,
                    Result = "up",
                    Technician = "up",
                    MR = "up",
                    EditName = "up",
                    OtherMethod = false,
                    MethodID = "b",
                   // TestingMethod = "up",
                    HTPlate = 1,
                    HTFlim = 1,
                    HTTime = 1,
                    HTPressure = (decimal)11.1,
                    HTPellOff = "up",
                    HT2ndPressnoreverse = 1,
                    HT2ndPressreversed = 1,
                    HTCoolingTime = "up",
                    MockupWash_Detail = new System.Collections.Generic.List<MockupWash_Detail_ViewModel>()
                    {
                        new MockupWash_Detail_ViewModel(){TypeofPrint="up",Design="up",ArtworkColor="up",FabricRefNo="up",AccessoryRefno="up",FabricColor  = "up",Result="up",Remark="up",EditName = "up",Ukey=21860},
                        new MockupWash_Detail_ViewModel(){TypeofPrint="up",Design="up",ArtworkColor="up",FabricRefNo="up",AccessoryRefno="up",FabricColor  = "up",Result="up",Remark="up",EditName = "up",Ukey=21861},
                    }
                };

                var mockupWash = _MockupWashService.Update(MockupWash, "SCIMIS");
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
                IMockupWashService _MockupWashService = new MockupWashService();
                MockupWash_ViewModel MockupWash = new MockupWash_ViewModel() { ReportNo = "T1" };
                var mockupWash = _MockupWashService.Delete(MockupWash);
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
                IMockupWashService _MockupWashService = new MockupWashService();
                System.Collections.Generic.List<MockupWash_Detail_ViewModel> MockupWash_Detail = new System.Collections.Generic.List<MockupWash_Detail_ViewModel>()
                {
                    new MockupWash_Detail_ViewModel(){Ukey=21896},
                    new MockupWash_Detail_ViewModel(){Ukey=21897},
                };

                var mockupWash = _MockupWashService.DeleteDetail(MockupWash_Detail);
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
                IMockupWashService _MockupWashService = new MockupWashService();
                MockupWash_Request MockupWash = new MockupWash_Request()
                { ReportNo = "VM1WA21090001" };
                //{ BrandID = "ADIDAS", SeasonID = "20SS", StyleID = "S201CSPM108", Article = "FL0237" };
                var mockupWash = _MockupWashService.GetMockupWash(MockupWash);
                var result = _MockupWashService.GetPDF(mockupWash, true);
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
                IMockupWashService _MockupWashService = new MockupWashService();
                MockupFailMail_Request MockupWash = new MockupFailMail_Request()
                {
                    ReportNo = "T1",
                    To = "jeff.yeh@sportscity.com.tw",
                };

                var mockupWash = _MockupWashService.FailSendMail(MockupWash);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}