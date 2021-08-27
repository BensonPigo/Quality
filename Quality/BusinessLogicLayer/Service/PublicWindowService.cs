using DatabaseObject.Public;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BusinessLogicLayer.Service
{
    public class PublicWindowService
    {
        private PublicWondowProvider _Provider { get; set; }

        public List<Window_Brand> Get_Brand(string ID)
        {
            List<Window_Brand> result = new List<Window_Brand>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Brand(ID).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Season> Get_Season(string BrandID, string ID)
        {
            List<Window_Season> result = new List<Window_Season>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Season(BrandID, ID).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Style> Get_Style(string BrandID, string SeasonID, string StyleID)
        {
            List<Window_Style> result = new List<Window_Style>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Style(BrandID, SeasonID, StyleID).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Article> Get_Article(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string Article)
        {
            List<Window_Article> result = new List<Window_Article>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Article(OrderID, StyleUkey, StyleID, BrandID, SeasonID, Article).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Article> Get_Size(string OrderID, Int64 StyleUkey, string Article, string Size)
        {
            List<Window_Article> result = new List<Window_Article>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Size(OrderID, StyleUkey, Article, Size).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Technician> Get_Technician(string CallFunction, string ID)
        {
            List<Window_Technician> result = new List<Window_Technician>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Technician(CallFunction, ID).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Pass1> Get_Pass1(string ID)
        {
            List<Window_Pass1> result = new List<Window_Pass1>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Pass1(ID).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_LocalSupp> Get_LocalSupp(string Name)
        {
            List<Window_LocalSupp> result = new List<Window_LocalSupp>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_LocalSupp(Name).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_TPESupp> Get_TPESupp(string Name)
        {
            List<Window_TPESupp> result = new List<Window_TPESupp>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_TPESupp(Name).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_Po_Supp_Detail> Get_Po_Supp_Detail(string POID, string FabricType, string Seq)
        {
            List<Window_Po_Supp_Detail> result = new List<Window_Po_Supp_Detail>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Po_Supp_Detail(POID, FabricType, Seq).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_FtyInventory> Get_FtyInventory(string POID, string Seq1, string Seq2, string Roll)
        {
            List<Window_FtyInventory> result = new List<Window_FtyInventory>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_FtyInventory(POID, Seq1, Seq2, Roll).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_FtyInventory> Get_Appearance(string Lab)
        {
            List<Window_FtyInventory> result = new List<Window_FtyInventory>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Appearance(Lab).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<Window_SewingLine> Get_SewingLine(string FactoryID, string ID)
        {
            List<Window_SewingLine> result = new List<Window_SewingLine>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_SewingLine(FactoryID, ID).ToList();

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
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

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
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_FGPT(VersionID, Code).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public Window_Picture Get_Picture(string Table, string BrforeColumn, string AfterColumn, string PKey_1, string PKey_2, string PKey_3, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val)
        {
            Window_Picture result = new Window_Picture();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                var r = _Provider.Get_Picture(Table, BrforeColumn, AfterColumn, PKey_1, PKey_2, PKey_3, PKey_1_Val, PKey_2_Val, PKey_3_Val).ToList();

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
                _Provider = new PublicWondowProvider(Common.ManufacturingExecutionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_TestFailMail(FactoryID, Type, GroupNameList).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
    }
}
