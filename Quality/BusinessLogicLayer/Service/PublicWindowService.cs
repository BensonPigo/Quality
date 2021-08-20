using DatabaseObject.Public;
using ProductionDataAccessLayer.Provider;
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

        public List<Window_Article> Get_Article(string OrderID, Int64 StyleUkey, string Article)
        {
            List<Window_Article> result = new List<Window_Article>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Article(OrderID, StyleUkey, Article).ToList();

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

        public List<Window_Technician> Get_Technician(string CallFunction, string Region, string ID)
        {
            List<Window_Technician> result = new List<Window_Technician>();

            try
            {
                _Provider = new PublicWondowProvider(Common.ProductionDataAccessLayer);

                //取得登入資訊
                result = _Provider.Get_Technician(CallFunction, Region, ID).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
    }
}
