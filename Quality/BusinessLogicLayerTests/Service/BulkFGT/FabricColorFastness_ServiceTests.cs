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
                    ID= "VM2CF21090002",
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
    }
}