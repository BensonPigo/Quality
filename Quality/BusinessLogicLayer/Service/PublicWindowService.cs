using DatabaseObject.Public;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BusinessLogicLayer.Service
{
    public class PublicWindowService
    {
        private PublicWindowProvider _Provider { get; set; }

        public List<Window_Brand> Get_Brand(string ID, bool IsExact)
        {
            List<Window_Brand> result = new List<Window_Brand>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Brand(ID, IsExact).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Brand> Get_Brand(string ID, string OtherTable, string OtherColumn)
        {
            List<Window_Brand> result = new List<Window_Brand>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Brand(ID, OtherTable, OtherColumn).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Season> Get_Season(string BrandID, string ID, bool IsExact)
        {
            List<Window_Season> result = new List<Window_Season>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Season(BrandID, ID, IsExact).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Style> Get_Style(string BrandID, string SeasonID, string StyleID, bool IsExact)
        {
            List<Window_Style> result = new List<Window_Style>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Style(BrandID, SeasonID, StyleID, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Article> Get_Article(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string Article, bool IsExact)
        {
            List<Window_Article> result = new List<Window_Article>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Article(OrderID, StyleUkey, StyleID, BrandID, SeasonID, Article, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Article> Get_PoidArticle(string POID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string Article, bool IsExact)
        {
            List<Window_Article> result = new List<Window_Article>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_PoidArticle(POID, StyleUkey, StyleID, BrandID, SeasonID, Article, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Size> Get_Size(string OrderID, Int64? StyleUkey, string BrandID, string SeasonID, string StyleID, string Article, string Size, bool IsExact)
        {
            List<Window_Size> result = new List<Window_Size>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Size(OrderID, StyleUkey, BrandID, SeasonID, StyleID, Article, Size, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Technician> Get_Technician(string CallFunction, string ID, bool IsExact)
        {
            List<Window_Technician> result = new List<Window_Technician>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Technician(CallFunction, ID, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Pass1> Get_Pass1(string ID, bool IsExact)
        {
            List<Window_Pass1> result = new List<Window_Pass1>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Pass1(ID, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Pass1> Get_FinalInspectionCFA(string ID, bool IsExact, bool FilterPivot88)
        {
            List<Window_Pass1> result = new List<Window_Pass1>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ManufacturingExecutionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_FinalInspectionCFA(ID, IsExact, FilterPivot88).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_LocalSupp> Get_LocalSupp(string SuppID , string Name, bool IsExact)
        {
            List<Window_LocalSupp> result = new List<Window_LocalSupp>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_LocalSupp(SuppID, Name, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_TPESupp> Get_TPESupp(string SuppID, string Name, bool IsExact)
        {
            List<Window_TPESupp> result = new List<Window_TPESupp>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_TPESupp(SuppID, Name, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Po_Supp_Detail> Get_Po_Supp_Detail(string POID, string FabricType, string Seq, bool IsExact)
        {
            List<Window_Po_Supp_Detail> result = new List<Window_Po_Supp_Detail>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Po_Supp_Detail(POID, FabricType, Seq, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Po_Supp_Detail> Get_Po_Supp_Detail_Refno(string OrderID, string MtlTypeID, string Refno)
        {
            List<Window_Po_Supp_Detail> result = new List<Window_Po_Supp_Detail>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Po_Supp_Detail_Refno(OrderID, MtlTypeID, Refno).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_Po_Supp_Detail> Get_HeatTransferWash_Refno(string OrderID, string Artwork, string Refno)
        {
            List<Window_Po_Supp_Detail> result = new List<Window_Po_Supp_Detail>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_HeatTransferWash_Refno(OrderID, Artwork, Refno).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_FtyInventory> Get_FtyInventory(string POID, string Seq1, string Seq2, string Roll, bool IsExact)
        {
            List<Window_FtyInventory> result = new List<Window_FtyInventory>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_FtyInventory(POID, Seq1, Seq2, Roll, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_FtyInventory> Get_AllReceivingDetail(string POID, string Seq1, string Seq2, string Roll, string ReceivingID, bool IsExact)
        {
            List<Window_FtyInventory> result = new List<Window_FtyInventory>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_AllReceivingDetail(POID, Seq1, Seq2, Roll, ReceivingID, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Appearance> Get_Appearance(string Lab)
        {
            List<Window_Appearance> result = new List<Window_Appearance>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Appearance(Lab).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_SewingLine> Get_SewingLine(string FactoryID, string ID, bool IsExact)
        {
            List<Window_SewingLine> result = new List<Window_SewingLine>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_SewingLine(FactoryID, ID, IsExact).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Color> Get_Color(string BrandID, string ID)
        {
            List<Window_Color> result = new List<Window_Color>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Color(BrandID, ID).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_FGPT> Get_FGPT(string VersionID, string Code)
        {
            List<Window_FGPT> result = new List<Window_FGPT>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_FGPT(VersionID, Code).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public Window_Picture Get_Picture(string Table, string BeforeColumn, string AfterColumn, string PKey_1, string PKey_2, string PKey_3, string PKey_4, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val, string PKey_4_Val)
        {
            Window_Picture result = new Window_Picture();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                var r = _Provider.Get_Picture(Table, BeforeColumn, AfterColumn, PKey_1, PKey_2, PKey_3, PKey_4, PKey_1_Val, PKey_2_Val, PKey_3_Val, PKey_4_Val).ToList();

                if (r.Any())
                {
                    result = r.FirstOrDefault();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

         public Window_Picture Get_Picture(string Table, string OneColumn, string TwoColumn, string ThreeColumn, string PKey_1, string PKey_2, string PKey_3, string PKey_4, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val, string PKey_4_Val)
        {
            Window_Picture result = new Window_Picture();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                var r = _Provider.Get_Picture(Table, OneColumn, TwoColumn, ThreeColumn, PKey_1, PKey_2, PKey_3, PKey_4, PKey_1_Val, PKey_2_Val, PKey_3_Val, PKey_4_Val).ToList();
                if (r.Any())
                {
                    result = r.FirstOrDefault();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public Window_MartindalePillingTest Get_MartindalePillingTestPicture(string ReportNo)
        {
            Window_MartindalePillingTest result = new Window_MartindalePillingTest();

            try
            {
                _Provider = new PublicWindowProvider(Common.ManufacturingExecutionDataAccessLayer);

                //取得登入資訊
                var r = _Provider.Get_MartindalePillingTestPicture(ReportNo).ToList();

                if (r.Any())
                {
                    result = r.FirstOrDefault();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public Window_RandomTumblePillingTest Get_RandomTumblePillingTestPicture(string ReportNo)
        {
            Window_RandomTumblePillingTest result = new Window_RandomTumblePillingTest();

            try
            {
                _Provider = new PublicWindowProvider(Common.ManufacturingExecutionDataAccessLayer);

                //取得登入資訊
                var r = _Provider.Get_RandomTumblePillingTestPicture(ReportNo).ToList();
                if (r.Any())
                {
                    result = r.FirstOrDefault();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public Window_SinglePicture Get_SinglePicture(string Table, string ColumnName, string PKey_1, string PKey_2, string PKey_3, string PKey_4, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val, string PKey_4_Val)
        {
            Window_SinglePicture result = new Window_SinglePicture();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                List<Window_SinglePicture> r = _Provider.Get_SinglePicture(Table, ColumnName, PKey_1, PKey_2, PKey_3, PKey_4, PKey_1_Val, PKey_2_Val, PKey_3_Val, PKey_4_Val).ToList();

                if (r.Any())
                {
                    result = r.FirstOrDefault();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }


        public List<Window_TestFailMail> Get_TestFailMail(string FactoryID, string Type, string GroupNameList)
        {
            List<Window_TestFailMail> result = new List<Window_TestFailMail>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ManufacturingExecutionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_TestFailMail(FactoryID, Type, GroupNameList).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_Operation> Get_Operation(string FinalInspectionID, string Operation)
        {
            List<Window_Operation> result = new List<Window_Operation>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ManufacturingExecutionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Operation(FinalInspectionID, Operation).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_AreaCode> Get_AreaCode(string FinalInspectionID, string AreaCode, string oldValue)
        {
            List<Window_AreaCode> result = new List<Window_AreaCode>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ManufacturingExecutionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_AreaCode(FinalInspectionID, AreaCode , oldValue).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_FabricRefNo> Get_FabricRefNo(string OrderID, string Refno)
        {
            List<Window_FabricRefNo> result = new List<Window_FabricRefNo>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_FabricRefNo(OrderID, Refno).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_FabricRefNo> Get_AccRefNo(string OrderID, string Refno)
        {
            List<Window_FabricRefNo> result = new List<Window_FabricRefNo>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_AccRefNo(OrderID, Refno).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_FabricRefNo> Get_ArtworkRefNo(string OrderID, string Refno)
        {
            List<Window_FabricRefNo> result = new List<Window_FabricRefNo>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_ArtworkRefNo(OrderID, Refno).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }        
        public List<Window_InkType> Get_InkType(string BrandID, string SeasonID, string StyleID)
        {
            List<Window_InkType> result = new List<Window_InkType>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_InkType(BrandID, SeasonID, StyleID).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_RollDyelot> Get_RollDyelot(string OrderID, string Seq1, string Seq2)
        {
            List<Window_RollDyelot> result = new List<Window_RollDyelot>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_RollDyelot(OrderID, Seq1, Seq2).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
        public List<Window_BrandBulkTestItem> Get_BrandBulkTestItem(string BrandID, string TestITem)
        {
            List<Window_BrandBulkTestItem> result = new List<Window_BrandBulkTestItem>();

            try
            {
                _Provider = new PublicWindowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_BrandBulkTestItem(BrandID, TestITem).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
    }
}
