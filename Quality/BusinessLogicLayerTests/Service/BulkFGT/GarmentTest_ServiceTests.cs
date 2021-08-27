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
                result = _IGarmentTestDetailShrinkageProvider.Get_GarmentTest_Detail_Shrinkage(9244, "1");

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
                result = _IGarmentDetailSpiralityProvider.Get_Garment_Detail_Spirality(16608, "1");

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
                result = _IGarmentTestDetailApperanceProvider.Get_GarmentTest_Detail_Apperance(18578, "1");

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
                result = _IGarmentTestDetailFGWTProvider.Get_GarmentTest_Detail_FGWT(16615, "1");

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
                result = _IGarmentTestDetailFGPTProvider.Get_GarmentTest_Detail_FGPT(16608, "1");

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
            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                GarmentTest_ViewModel garmentTest_ViewModel =
                    new GarmentTest_ViewModel
                    {
                        ID = 16608,
                        OrderID = "20032066WW",
                        StyleID = "ARWPF20125",
                        SeasonID = "20FW",
                        BrandID = "REEBOK",
                        Article = "FT0964",
                        MDivisionid = "VM2",
                    };

                GarmentTest_Detail detail = new GarmentTest_Detail
                {
                    No = 1,
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

                GarmentTest_Detail detail2 = new GarmentTest_Detail
                {
                    No = 2,
                    Result = "P",
                    inspdate = Convert.ToDateTime("2021-07-01"),
                    Remark = "test",
                    AddName = "Scimis",
                    AddDate = DateTime.Now,
                    Status = "New",
                    SizeCode = "32",
                    MtlTypeID = "KNIT",
                    OdourResult = "P"
                };

                GarmentTest_Detail detail3 = new GarmentTest_Detail
                {
                    No = 3,
                    //Result = "P",
                    //inspdate = Convert.ToDateTime("2021-07-01"),
                    Remark = "test 3",
                    AddName = "Scimis",
                    AddDate = DateTime.Now,
                    //Status = "New",
                    //SizeCode = "32",
                    MtlTypeID = "KNIT",
                    //OdourResult = "P"
                };



                List<GarmentTest_Detail> details = new List<GarmentTest_Detail>();
                //details.Add(detail);
                //details.Add(detail2);
                details.Add(detail3);

                #region 判斷是否空值
                string emptyMsg = string.Empty;
                if (garmentTest_ViewModel.ID == 0) { emptyMsg += "Master OrderID cannot be empty." + Environment.NewLine; }

                #endregion


                IGarmentTestProvider _IGarmentTestProvider = new GarmentTestProvider(_ISQLDataTransaction);
                int saveCnt = _IGarmentTestProvider.Save_GarmentTest(garmentTest_ViewModel, details);
                _ISQLDataTransaction.Commit();

                Assert.IsTrue(saveCnt > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail();
                throw;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }


        }
    }
}