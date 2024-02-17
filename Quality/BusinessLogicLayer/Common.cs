using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace BusinessLogicLayer
{
    public class Common
    {
        static string region = "";
        private static void GetRegion()
        {
            //取得連線名稱統一放在網站目錄下Region.txt的內容
            if (region.Equals(string.Empty) || region.Equals("Debug"))
            {
                if (HttpContext.Current == null)
                {
                    region = "Debug";
                    return;
                }

                string region_path = HttpContext.Current.Server.MapPath("~\\Region.txt");
                if (!File.Exists(region_path))
                {
                    region = "Debug";
                }
                else
                {
                    using (StreamReader sr = new StreamReader(region_path, Encoding.Default))
                    {
                        region = sr.ReadToEnd();
                        sr.Close();
                    }
                }
            }
        }
        public static string Region
        {
            get
            {
                GetRegion();
                return region;
            }
        }
        public static string DashboardDataAccessLayer {
            get {
                GetRegion();
                return region + "_DashboardDataAccessLayer";
            }
        }
        public static string ProductionDataAccessLayer
        {
            get
            {
                GetRegion();
                return region + "_ProductionDataAccessLayer";
            }
        }

        public static string ManufacturingExecutionDataAccessLayer
        {
            get
            {
                GetRegion();
                return region + "_ManufacturingExecutionDataAccessLayer";
            }
        }

        public static string PMSSewingAPIurlDataAccessLayer
        {
            get
            {
                GetRegion();
                return region + "_PMSSewingAPIurl";
            }
        }
    }
}
