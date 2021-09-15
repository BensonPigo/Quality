using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using DatabaseObject.ProductionDB;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class MockupCrockingTests
    {
        [TestMethod()]
        public void GetMockupCrocking()
        {
            try
            {
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                MockupCrocking_Request MockupCrocking = new MockupCrocking_Request()
                { ReportNo = "PHCK180800001" };
                //{ BrandID = "ADIDAS", SeasonID = "20SS", StyleID = "S201CSPM108", Article = "FL0237" };
                var mockupCrocking = _MockupCrockingService.GetMockupCrocking(MockupCrocking);

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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                StyleArtwork_Request StyleArtwork_Request = new StyleArtwork_Request()
                {
                    BrandID = "U.ARMOUR",
                    SeasonID = "20FW",
                    StyleID = "1342962",
                    //StyleUkey =75468,
                };
                var selectListItems = _MockupCrockingService.GetArtworkTypeID(StyleArtwork_Request);
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                var orders = _MockupCrockingService.GetOrders(new Orders() { ID = "21090101IE022" });
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                Order_Qty Order_Qty = new Order_Qty()
                {
                    ID = "21080085IE044",
                };
                var x = _MockupCrockingService.GetDistinctArticle(Order_Qty);
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                MockupCrocking_ViewModel MockupCrocking = new MockupCrocking_ViewModel()
                {
                    POID = "TP",
                    StyleID = "SS",
                    SeasonID = "20SS",
                    BrandID = "bb",
                    Article = "aa",
                    ArtworkTypeID = "attt",
                    Remark = "remmmm",
                    T1Subcon = "SCIMIS",
                    TestDate = DateTime.Now,
                    ReceivedDate = DateTime.Now,
                    ReleasedDate = DateTime.Now,
                    Result = "Pass",
                    Technician = "SCIMIS",
                    MR = "SCIMIS",
                    AddName = "SCIMIS",
                    MockupCrocking_Detail = new System.Collections.Generic.List<MockupCrocking_Detail_ViewModel>()
                    {
                        new MockupCrocking_Detail_ViewModel(){Design="d100",ArtworkColor="0001",FabricRefNo="RF",FabricColor  = "FCC",Result="Pass",Remark="RRRK",EditName = "SCIMIS",WetScale="1",DryScale="2-2"},
                        //new MockupCrocking_Detail_ViewModel(){ReportNo = "T1",Design="d200",ArtworkColor="0001",FabricRefNo="RF",FabricColor  = "FCC",Result="Pass",Remark="2RRRK",EditName = "SCIMIS",WetScale="1",DryScale="2-2"},
                    }
                };

                var mockupCrocking = _MockupCrockingService.Create(MockupCrocking,"PM1","SCIMIS",out string no);
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                MockupCrocking_ViewModel MockupCrocking = new MockupCrocking_ViewModel()
                {
                    ReportNo = "T1",
                    POID = "up",
                    StyleID = "up",
                    SeasonID = "up",
                    BrandID = "up",
                    Article = "up",
                    ArtworkTypeID = "up",
                    Remark = "up",
                    T1Subcon = "up",
                    TestDate = DateTime.Now,
                    ReceivedDate = DateTime.Now,
                    ReleasedDate = DateTime.Now,
                    Result = "up",
                    Technician = "up",
                    MR = "up",
                    EditName = "up",
                    MockupCrocking_Detail = new System.Collections.Generic.List<MockupCrocking_Detail_ViewModel>()
                    {
                        new MockupCrocking_Detail_ViewModel(){Design="up",ArtworkColor="up",FabricRefNo="up",FabricColor  = "up",Result="up",Remark="up",EditName = "up",Ukey=76},
                        new MockupCrocking_Detail_ViewModel(){Design="up",ArtworkColor="up",FabricRefNo="up",FabricColor  = "up",Result="up",Remark="up",EditName = "up",Ukey=77},
                    }
                };

                var mockupCrocking = _MockupCrockingService.Update(MockupCrocking, "SCIMIS");
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                MockupCrocking_ViewModel MockupCrocking = new MockupCrocking_ViewModel() { ReportNo = "T1" };
                var mockupCrocking = _MockupCrockingService.Delete(MockupCrocking);
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                System.Collections.Generic.List<MockupCrocking_Detail_ViewModel> MockupCrocking_Detail = new System.Collections.Generic.List<MockupCrocking_Detail_ViewModel>()
                {
                    //new MockupCrocking_Detail_ViewModel(){Ukey=21858},
                    new MockupCrocking_Detail_ViewModel(){Ukey=77},
                };

                var mockupCrocking = _MockupCrockingService.DeleteDetail(MockupCrocking_Detail);
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                MockupCrocking_Request MockupCrocking = new MockupCrocking_Request()
                { ReportNo = "PHOV210900008" };
                //{ BrandID = "ADIDAS", SeasonID = "20SS", StyleID = "S201CSPM108", Article = "FL0237" };
                var mockupCrocking = _MockupCrockingService.GetMockupCrocking(MockupCrocking);
                _MockupCrockingService.GetPDF(mockupCrocking, true);
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                MockupFailMail_Request MockupCrocking = new MockupFailMail_Request()
                {
                    ReportNo = "T1",
                    To = "jeff.yeh@sportscity.com.tw",
                };

                var mockupCrocking = _MockupCrockingService.FailSendMail(MockupCrocking);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}