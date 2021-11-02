using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web;

namespace ADOHelper.DBToolKit
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

        public Common()
		{
            
        }

		public static string GetAppSetting(string Key, string assembly)
		{
			string empty = string.Empty;
			Uri uri = new Uri(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase));
			Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(new ExeConfigurationFileMap()
			{
				ExeConfigFilename = Path.Combine(uri.LocalPath, string.Concat(assembly, ".config"))
			}, ConfigurationUserLevel.None);
			if (configuration.HasFile)
			{
				AppSettingsSection section = configuration.GetSection("appSettings") as AppSettingsSection;
				empty = section.Settings[Key].Value;
			}
			return empty;
		}
	}
}