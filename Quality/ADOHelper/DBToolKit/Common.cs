using System;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace ADOHelper.DBToolKit
{
	public class Common
	{
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