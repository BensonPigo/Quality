using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ViewModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ManufacturingExecutionDB;
using System.Data;
using DatabaseObject.RequestModel;

namespace BusinessLogicLayer.Service.BulkFGT.Tests
{
    [TestClass()]
    public class GarmentTest_ServiceTests
    {
        [TestMethod()]
        public void GetGarmentTestTest()
        {
            try
            {
                IGarmentTestProvider _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
                IGarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
                GarmentTest_Result result = new GarmentTest_Result();
                //var query = _IGarmentTestProvider.Get_GarmentTest(
                //    new GarmentTest_ViewModel
                //    {
                //        StyleID = "NF0A3SR4",
                //        BrandID = "N.FACE",
                //        Article = "N8E",
                //        SeasonID = "19FW",
                //        MDivisionid = "Vm2",
                //    });

                //result.garmentTest = query.FirstOrDefault();

                // Detail
                result.garmentTest_Details = _IGarmentTestDetailProvider.Get_GarmentTestDetail(
                    new GarmentTest_ViewModel
                    {
                        ID = result.garmentTest.ID
                    }).ToList();


                Assert.IsTrue(result.garmentTest_Details.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_SizeCodeTest()
        {
            try
            {
                List<string> result = new List<string>();
                IGarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailProvider.GetSizeCode("ESPLO14120030", "010-BLAK").Select(x => x.SizeCode).ToList();

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_ShrinkageTest()
        {
            try
            {
                IList<GarmentTest_Detail_Shrinkage> result = new List<GarmentTest_Detail_Shrinkage>();
                IGarmentTestDetailShrinkageProvider _IGarmentTestDetailShrinkageProvider = new GarmentTestDetailShrinkageProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailShrinkageProvider.Get_GarmentTest_Detail_Shrinkage("9244", "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_SpiralityTest()
        {
            try
            {
                IList<Garment_Detail_Spirality> result = new List<Garment_Detail_Spirality>();
                IGarmentDetailSpiralityProvider _IGarmentDetailSpiralityProvider = new GarmentDetailSpiralityProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentDetailSpiralityProvider.Get_Garment_Detail_Spirality("16608", "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_ApperanceTest()
        {
            try
            {
                IList<GarmentTest_Detail_Apperance_ViewModel> result = new List<GarmentTest_Detail_Apperance_ViewModel>();
                IGarmentTestDetailApperanceProvider _IGarmentTestDetailApperanceProvider = new GarmentTestDetailApperanceProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailApperanceProvider.Get_GarmentTest_Detail_Apperance("18578", "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_FGWTTest()
        {
            try
            {
                IList<GarmentTest_Detail_FGWT_ViewModel> result = new List<GarmentTest_Detail_FGWT_ViewModel>();
                IGarmentTestDetailFGWTProvider _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailFGWTProvider.Get_GarmentTest_Detail_FGWT("16615", "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_FGPTTest()
        {
            try
            {
                IList<GarmentTest_Detail_FGPT_ViewModel> result = new List<GarmentTest_Detail_FGPT_ViewModel>();
                IGarmentTestDetailFGPTProvider _IGarmentTestDetailFGPTProvider = new GarmentTestDetailFGPTProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailFGPTProvider.Get_GarmentTest_Detail_FGPT("16608", "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Save_GarmentTestTest()
        {
            //GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            //SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            //try
            //{
            //    GarmentTest_ViewModel garmentTest_ViewModel =
            //        new GarmentTest_ViewModel
            //        {
            //            ID = 16608,
            //            OrderID = "20032066WW",
            //            StyleID = "ARWPF20125",
            //            SeasonID = "20FW",
            //            BrandID = "REEBOK",
            //            Article = "FT0964",
            //            MDivisionid = "VM2",
            //        };

            //    GarmentTest_Detail detail = new GarmentTest_Detail
            //    {
            //        No = 1,
            //        Result = "P",
            //        inspdate = Convert.ToDateTime("2021-07-01"),
            //        Remark = "test",
            //        AddName = "EE04284",
            //        AddDate = Convert.ToDateTime("2021-07-08 12:46:35.343"),
            //        EditName = "EE04284",
            //        EditDate = Convert.ToDateTime("2021-07-08 00:00:00.000"),
            //        Status = "New",
            //        SizeCode = "36",
            //        MtlTypeID = "WOVEN",
            //    };

            //    GarmentTest_Detail detail2 = new GarmentTest_Detail
            //    {
            //        //No = 2,
            //        Result = "P",
            //        inspdate = Convert.ToDateTime("2021-07-01"),
            //        Remark = "test 4",
            //        AddName = "Scimis",
            //        AddDate = DateTime.Now,
            //        Status = "New",
            //        SizeCode = "32",
            //        MtlTypeID = "KNIT",
            //        OdourResult = "P"
            //    };

            //    GarmentTest_Detail detail3 = new GarmentTest_Detail
            //    {
            //        //No = 3,
            //        //Result = "P",
            //        //inspdate = Convert.ToDateTime("2021-07-01"),
            //        Remark = "test 5",
            //        AddName = "Scimis",
            //        AddDate = DateTime.Now,
            //        //Status = "New",
            //        //SizeCode = "32",
            //        MtlTypeID = "KNIT",
            //        //OdourResult = "P"
            //    };



            //    List<GarmentTest_Detail> details = new List<GarmentTest_Detail>();
            //    //details.Add(detail);
            //    details.Add(detail2);
            //    details.Add(detail3);

            //    #region 判斷是否空值
            //    string emptyMsg = string.Empty;
            //    if (garmentTest_ViewModel.ID == 0) { emptyMsg += "Master OrderID cannot be empty." + Environment.NewLine; }

            //    #endregion


            //    IGarmentTestProvider _IGarmentTestProvider = new GarmentTestProvider(_ISQLDataTransaction);
            //    //bool saveCnt = _IGarmentTestProvider.Save_GarmentTest(garmentTest_ViewModel, details);
            //    _ISQLDataTransaction.Commit();

            //    Assert.IsTrue(true);
            //}
            //catch (Exception ex)
            //{
            //    Assert.Fail();
            //    throw ex;
            //}
            //finally { _ISQLDataTransaction.CloseConnection(); }


        }

        [TestMethod()]
        public void Generate_FGWTTest()
        {
            //SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            //try
            //{
            //    GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            //    GarmentTest_Detail_ViewModel detail = new GarmentTest_Detail_ViewModel
            //    {
            //        ID = 16608,
            //        No = 8,
            //        SubmitDate = Convert.ToDateTime("2021-08-30"),
            //        ArrivedQty = 1,
            //        LOtoFactory = "ESP",
            //        Result = "P",
            //        Remark = "Test",
            //        LineDry = true,
            //        Temperature = 30,
            //        TumbleDry = false,
            //        Machine = "Top Load",
            //        HandWash = false,
            //        Composition = "",
            //        Neck = true,
            //        Above50NaturalFibres = true,
            //        Above50SyntheticFibres = false,
            //        EditName = "SCIMIS",
            //    };


            //    // Detail Save
            //    IGarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
            //    if (_IGarmentTestDetailProvider.Update_GarmentTestDetail(detail) == false)
            //    {
            //        _ISQLDataTransaction.RollBack();
            //        result.Result = false;
            //        result.ErrMsg = "Update detail is empty.";
            //        Assert.Fail();
            //    }

            //    // Shrinkage Save
            //    GarmentTest_Detail_Shrinkage shrinkage = new GarmentTest_Detail_Shrinkage
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Location = "T",
            //        Type = "chest width",
            //        Seq = 1,
            //        BeforeWash = 4,
            //        SizeSpec = 2,
            //        AfterWash1 = 3,
            //        Shrinkage1 = -25,
            //        AfterWash2 = 4,
            //        Shrinkage2 = 0,
            //        AfterWash3 = 5,
            //        Shrinkage3 = 25,
            //    };

            //    GarmentTest_Detail_Shrinkage shrinkage2 = new GarmentTest_Detail_Shrinkage
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Location = "T",
            //        Type = "sleeve width",
            //        Seq = 2,
            //        BeforeWash = 2,
            //        SizeSpec = 2,
            //        AfterWash1 = 3,
            //        Shrinkage1 = 50,
            //        AfterWash2 = 4,
            //        Shrinkage2 = 100,
            //        AfterWash3 = 0,
            //        Shrinkage3 = 0,
            //    };

            //    List<GarmentTest_Detail_Shrinkage> Shrinkages = new List<GarmentTest_Detail_Shrinkage>();
            //    Shrinkages.Add(shrinkage);
            //    Shrinkages.Add(shrinkage2);

            //    GarmentTestDetailShrinkageProvider _IGarmentTestDetailShrinkageProvider = new GarmentTestDetailShrinkageProvider(_ISQLDataTransaction);
            //    if (_IGarmentTestDetailShrinkageProvider.Update_GarmentTestShrinkage(Shrinkages) == false)
            //    {
            //        _ISQLDataTransaction.RollBack();
            //        result.Result = false;
            //        result.ErrMsg = "Update Shrinkage is empty.";
            //        Assert.Fail();
            //    }


            //    // Spirality Save
            //    Garment_Detail_Spirality Spirality = new Garment_Detail_Spirality
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Location = "T",
            //        MethodA_AAPrime = 1,
            //        MethodA_APrimeB = 4,
            //        MethodB_AAPrime = 3,
            //        MethodB_AB = 5,
            //        CM = 2,
            //        MethodA = 25,
            //        MethodB = 60,
            //    };

            //    List<Garment_Detail_Spirality> Spiralities = new List<Garment_Detail_Spirality>();
            //    Spiralities.Add(Spirality);

            //    GarmentDetailSpiralityProvider _IGarmentDetailSpiralityProvider = new GarmentDetailSpiralityProvider(_ISQLDataTransaction);
            //    if (_IGarmentDetailSpiralityProvider.Update_Spirality(Spiralities) == false)
            //    {
            //        _ISQLDataTransaction.RollBack();
            //        result.Result = false;
            //        result.ErrMsg = "Update Spirality is empty.";
            //        Assert.Fail();
            //    }



            //    // Apperance Save 
            //    GarmentTest_Detail_Apperance_ViewModel Apperance = new GarmentTest_Detail_Apperance_ViewModel
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Type = "Printing / Heat Transfer",
            //        Wash1 = "Accepted",
            //        Wash2 = "Rejected",
            //        Wash3 = "Accepted",
            //        Comment = "test",
            //        Seq = 1,
            //    };

            //    GarmentTest_Detail_Apperance_ViewModel Apperance2 = new GarmentTest_Detail_Apperance_ViewModel
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Type = "Label",
            //        Wash1 = "Accepted",
            //        Wash2 = "Rejected",
            //        Wash3 = "Accepted",
            //        Comment = "test",
            //        Seq = 2,
            //    };
            //    List<GarmentTest_Detail_Apperance_ViewModel> Apperances = new List<GarmentTest_Detail_Apperance_ViewModel>();
            //    Apperances.Add(Apperance);
            //    Apperances.Add(Apperance2);


            //    GarmentTestDetailApperanceProvider _IGarmentTestDetailApperanceProvider = new GarmentTestDetailApperanceProvider(_ISQLDataTransaction);
            //    if (_IGarmentTestDetailApperanceProvider.Update_Apperance(Apperances) == false)
            //    {
            //        _ISQLDataTransaction.RollBack();
            //        result.Result = false;
            //        result.ErrMsg = "Update Apperance is empty.";
            //        Assert.Fail();
            //    }


            //    // FGPT Save
            //    GarmentTest_Detail_FGPT_ViewModel FGPT = new GarmentTest_Detail_FGPT_ViewModel
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Location = "T",
            //        Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection",
            //        TestName = "PHX-AP0450",
            //        TestDetail = "N",
            //        Criteria = 180,
            //        TestResult = "",
            //        TestUnit = "N",
            //        Seq = 7,
            //        TypeSelection_Seq = 1,
            //        TypeSelection_VersionID = 1,
            //    };

            //    GarmentTest_Detail_FGPT_ViewModel FGPT2 = new GarmentTest_Detail_FGPT_ViewModel
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Location = "T",
            //        Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Armhole seam - Method B ≥180N)",
            //        TestName = "PHX-AP0450",
            //        TestDetail = "N",
            //        Criteria = 180,
            //        TestResult = "",
            //        TestUnit = "N",
            //        Seq = 2,
            //        TypeSelection_Seq = 0,
            //        TypeSelection_VersionID = 0,
            //    };

            //    List<GarmentTest_Detail_FGPT_ViewModel> FGPTs = new List<GarmentTest_Detail_FGPT_ViewModel>();
            //    FGPTs.Add(FGPT);

            //    GarmentTestDetailFGPTProvider _IGarmentTestDetailFGPTProvider = new GarmentTestDetailFGPTProvider(_ISQLDataTransaction);
            //    if (_IGarmentTestDetailFGPTProvider.Update_FGPT(FGPTs) == false)
            //    {
            //        _ISQLDataTransaction.RollBack();
            //        result.Result = false;
            //        result.ErrMsg = "Update FGPT is empty.";
            //        Assert.Fail();
            //    }

            //    // FGWT Save
            //    GarmentTest_Detail_FGWT_ViewModel FGWT = new GarmentTest_Detail_FGWT_ViewModel
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Location = "",
            //        Type = "spirality: Garment - hem opening in cm",
            //        TestDetail = "cm",
            //        BeforeWash = 2,
            //        SizeSpec = 0,
            //        AfterWash = 1,
            //        Shrinkage = -50,
            //        Scale = "",
            //        Criteria = 2,
            //        Criteria2 = 0,
            //        SystemType = "spirality: Garment - hem opening in cm",
            //        Seq = 1
            //    };

            //    GarmentTest_Detail_FGWT_ViewModel FGWT2 = new GarmentTest_Detail_FGWT_ViewModel
            //    {
            //        ID = 16608,
            //        No = 8,
            //        Location = "T",
            //        Type = "dimensional change: jacket-like garment a) length of necktape",
            //        TestDetail = "%",
            //        BeforeWash = 2,
            //        SizeSpec = 0,
            //        AfterWash = 1,
            //        Shrinkage = -50,
            //        Scale = "",
            //        Criteria = 2,
            //        Criteria2 = 0,
            //        SystemType = "spirality: Garment - hem opening in cm",
            //        Seq = 1
            //    };

            //    List<GarmentTest_Detail_FGWT_ViewModel> FGWTs = new List<GarmentTest_Detail_FGWT_ViewModel>();
            //    FGWTs.Add(FGWT);
            //    FGWTs.Add(FGWT);

            //    GarmentTestDetailFGWTProvider _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(_ISQLDataTransaction);
            //    if (_IGarmentTestDetailFGWTProvider.Update_FGWT(FGWTs) == false)
            //    {
            //        _ISQLDataTransaction.RollBack();
            //        result.Result = false;
            //        result.ErrMsg = "Update FGWT is empty.";
            //        Assert.Fail();
            //    }
            //    _ISQLDataTransaction.Commit();

            //    Assert.IsTrue(true);
            //}
            //catch (Exception ex)
            //{
            //    Assert.Fail();
            //    throw ex;
            //}
            //finally { _ISQLDataTransaction.CloseConnection(); }

        }

        [TestMethod()]
        public void Save_GarmentTestDetailTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void Encode_DetailTest()
        {
            //GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            //SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            //try
            //{
            //    string _status = "Confirmed";
            //    IGarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
            //    result.Result = _IGarmentTestDetailProvider.Encode_GarmentTestDetail("16608", _status);
            //    _ISQLDataTransaction.Commit();
            //    Assert.IsTrue(result.Result == true);
            //}
            //catch (Exception ex)
            //{
            //    _ISQLDataTransaction.RollBack();
            //    result.Result = false;
            //    result.ErrMsg = ex.Message;
            //    Assert.Fail();
            //}
            //finally { _ISQLDataTransaction.CloseConnection(); }
        }

        [TestMethod()]
        public void Encode_DetailTest1()
        {
            //GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            //SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            //try
            //{
            //    GarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);

            //    // 代表所有result 有任一個是Fail 就寄信
            //    if (_IGarmentTestDetailProvider.Chk_AllResult("16608", "2") == false)
            //    {

            //    }
            //    result.Result = _IGarmentTestDetailProvider.Encode_GarmentTestDetail("16608", "Confirmed");
            //    _ISQLDataTransaction.Commit();
            //    Assert.IsTrue(result.Result == true);
            //}
            //catch (Exception ex)
            //{
            //    _ISQLDataTransaction.RollBack();
            //    result.Result = false;
            //    result.ErrMsg = ex.Message;
            //    Assert.Fail();
            //}
            //finally { _ISQLDataTransaction.CloseConnection(); }
        }

        [TestMethod()]
        public void Get_All_DetailTest()
        {
            try
            {
                GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();

                IGarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
                result.Scales = _IGarmentTestDetailProvider.GetScales();
                Assert.IsTrue(result.Scales.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail();
                throw ex;
            }

        }

        [TestMethod()]
        public void SentMailTest()
        {
            //GarmentTest_Result result = new GarmentTest_Result();
            //string ToAddress = string.Empty;
            //string CCAddress = string.Empty;
            //Quality_MailGroup mail_01 = new Quality_MailGroup
            //{
            //    ToAddress = "willy.wei@sportscity.com.tw",
            //    CcAddress = "willy.wei@sportscity.com.tw",
            //};

            //Quality_MailGroup mail_02 = new Quality_MailGroup
            //{
            //    ToAddress = "willy.wei@sportscity.com.tw",
            //    CcAddress = "willy.wei@sportscity.com.tw",
            //};

            //List<Quality_MailGroup> mailGroups = new List<Quality_MailGroup>();
            //mailGroups.Add(mail_01);
            //mailGroups.Add(mail_02);

            //try
            //{
            //    foreach (var item in mailGroups)
            //    {
            //        ToAddress += item.ToAddress + ";";
            //        CCAddress += item.CcAddress + ";";
            //    }

            //    if (string.IsNullOrEmpty(ToAddress) == true)
            //    {
            //        result.Result = false;
            //        result.ErrMsg = "To email address is empty!";
            //        Assert.Fail();
            //    }

            //    IGarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);

            //    DataTable dtContent = _IGarmentTestDetailProvider.Get_Mail_Content("16608", "2");
            //    string strHtml = MailTools.DataTableChangeHtml(dtContent);

            //    SendMail_Request request = new SendMail_Request()
            //    {
            //        To = ToAddress,
            //        CC = CCAddress,
            //        Subject = "Garment Test – Test Fail",
            //        Body = strHtml,
            //    };

            //    MailTools.SendMail(request);

            //    result.Result = true;
            //}
            //catch (Exception ex)
            //{
            //    result.Result = false;
            //    result.ErrMsg = ex.Message.ToString();
            //    Assert.Fail();
            //}

        }

        [TestMethod()]
        public void ToReportTest()
        {
            try
            {
                GarmentTest_Detail_Result all_Data = new GarmentTest_Detail_Result();
                IGarmentTest_Service _Service = new GarmentTest_Service();
                all_Data = _Service.ToReport("19301", "6", GarmentTest_Service.ReportType.Wash_Test_2018, true, true);
                Assert.IsTrue(!string.IsNullOrEmpty(all_Data.reportPath));
            }
            catch (Exception ex)
            {
                Assert.Fail();
                throw ex;
            }

        }

        [TestMethod()]
        public void DetailPictureSaveTest()
        {
            try
            {
                GarmentTest_Result all_Data = new GarmentTest_Result();
                GarmentTest_Detail detail = new GarmentTest_Detail
                {
                    ID = 16608,
                    No = 2,
                    TestAfterPicture = null,
                    TestBeforePicture = null,
                };

                IGarmentTest_Service _Service = new GarmentTest_Service();
                Assert.IsTrue(all_Data.Result);
            }
            catch (Exception ex)
            {
                Assert.Fail();
                throw ex;
            }
        }

        [TestMethod()]
        public void Save_GarmentTestTest1()
        {
            try
            {
                IGarmentTest_Service _Service = new GarmentTest_Service();
                GarmentTest_ViewModel result = new GarmentTest_ViewModel();
                GarmentTest_ViewModel garmentTest_ViewModel =
                   new GarmentTest_ViewModel
                   {
                       ID = 16615,
                       OrderID = "20032069WW001",
                       StyleID = "ARWPF20178",
                       SeasonID = "20FW",
                       BrandID = "REEBOK",
                       Article = "FT0FU2340964",
                       MDivisionid = "VM2",
                   };

                GarmentTest_Detail detail = new GarmentTest_Detail
                {
                    ID = 16615,
                    No = 3,
                    Result = "P",
                    inspdate = Convert.ToDateTime("2021-07-01"),
                    Remark = "test",
                    AddName = "EE04284",
                    AddDate = Convert.ToDateTime("2021-07-08 12:46:35.343"),
                    EditName = "EE04284",
                    EditDate = Convert.ToDateTime("2021-07-08 00:00:00.000"),
                    Status = "New",
                    SizeCode = "36",
                    MtlTypeID = "WOVEN",
                };

                List<GarmentTest_Detail> details = new List<GarmentTest_Detail>();
                details.Add(detail);

                result = _Service.Save_GarmentTest(garmentTest_ViewModel, details, "SCIMIS");
                Assert.IsTrue(result.SaveResult);
            }
            catch (Exception ex)
            {
                Assert.Fail();
                throw ex;
            }
        }
    }
}