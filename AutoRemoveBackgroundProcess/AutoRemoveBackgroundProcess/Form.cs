using Sci;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace AutoRemoveBackgroundProcess
{
    public partial class Form : System.Windows.Forms.Form
    {
        private string appPath;
        private Setting RemoveSetting;
        
        public Form(string auto_run)
        {
            InitializeComponent();
            appPath = Application.StartupPath + @"\";
            RemoveSetting = new Setting(appPath);
            if (auto_run.Equals(string.Empty))
            {
                this.ReloadSetting();
            }
        }
        private void btnSettingSave_Click(object sender, EventArgs e)
        {
            SaveSetup();
            this.ReloadSetting();
        }
        //將畫面設定存入參數中
        private void SaveSetup()
        {
            try
            {

                //要移除的Process、以及閒置時間，把設定存下來
                string remove_target = string.Empty;
                string remove_target_idle = string.Empty;
                if (this.ChkExcel.Checked)
                {
                    remove_target += ",Excel.exe";
                    remove_target_idle += $@",Excel.exe_{this.numExcel.Value}";
                }

                remove_target = remove_target.Substring(1);
                remove_target_idle = remove_target_idle.Substring(1);


                RemoveSetting.Write("remove_target", remove_target);
                RemoveSetting.Write("remove_target_idle", remove_target_idle);
                MyUtility.Msg.InfoBox("Success!!");
            }
            catch (Exception ex)
            {
                MyUtility.Msg.WarningBox(ex.Message);
            }

        }
        private void ReloadSetting()
        {
            // 先全部清除勾選
            this.ChkExcel.Checked = false;

            // 讀取XML檔
            List<string> list_remove_target; list_remove_target = RemoveSetting.Read("remove_target").Split(',').Where(o => !string.IsNullOrEmpty(o)).ToList();
            List<string> remove_target_idle = RemoveSetting.Read("remove_target_idle").Split(',').Where(o => !string.IsNullOrEmpty(o)).ToList();

            // Excel.exe
            this.ChkExcel.Checked = list_remove_target.Any(o => o.ToUpper() == "EXCEL.EXE");

            foreach (var item in remove_target_idle)
            {
                string[] tmp = item.Split('_');
                string process = tmp[0];
                int idelTime = Convert.ToInt32(tmp[1]);

                // Excel.exe
                if (process.ToUpper() == "EXCEL.EXE")
                {
                    this.numExcel.Value = idelTime;
                }
            }
        }
        private void btnRemove_Click(object sender, EventArgs e)
        {
             System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcesses();
            bool removeExcel = this.ChkExcel.Checked;

            foreach (var proc in procs)
            {
                int idleTime = 0;
                if (removeExcel && proc.ProcessName.ToUpper() == "EXCEL")
                {
                    idleTime = Convert.ToInt32(this.numExcel.Value);
                    DateTime startTime = proc.StartTime;
                    TimeSpan sp = DateTime.Now - startTime;

                    if (sp.Minutes >= idleTime && idleTime > 10)
                    {
                        proc.Kill();
                    }
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

    public class Setting
    {
        private string xml_path;
        private string xml_root;
        public Setting(string appPath)
        {
            xml_path = appPath + "RemoveSetup.xml";
            xml_root = "RemoveSetting";
        }
        public void Write(string tag, string tag_value)
        {
            XmlDocument doc = new XmlDocument();
            if (File.Exists(xml_path))
            {
                doc.Load(xml_path);

            }
            //選擇節點
            XmlNode node = doc.SelectSingleNode(xml_root);

            if (node == null)
            {
                XmlElement main = doc.CreateElement(xml_root);
                main.SetAttribute(tag, tag_value);
                doc.AppendChild(main);
            }
            else
            {
                XmlElement main = (XmlElement)node;
                main.SetAttribute(tag, tag_value);
            }

            doc.Save(xml_path);
        }
        public string Read(string tag)
        {
            XmlDocument doc = new XmlDocument();
            if (File.Exists(xml_path))
            {
                doc.Load(xml_path);

            }
            else
            {
                return string.Empty;
            }
            //選擇節點
            XmlNode node = doc.SelectSingleNode(xml_root);

            if (node == null)
            {
                return string.Empty;
            }
            else
            {
                XmlElement main = (XmlElement)node;
                return main.GetAttribute(tag);
            }

        }
    }
}
