using BusinessLogicLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;

namespace BusinessLogicLayer.Service
{
    public class BulkFGTMailGroup_Service : IBulkFGTMailGroup_Service
    {
        private IBulkFGTMailGroupProvider _BulkFGTMailGroupProvider;

        public enum SaveType
        {
            Insert = 0,
            Update = 0,
            Delete = 1,
        }

        public List<Quality_MailGroup> MailGroupGet(Quality_MailGroup quality_Mail)
        {
            List<Quality_MailGroup> mailGroups = new List<Quality_MailGroup>();
            _BulkFGTMailGroupProvider = new BulkFGTMailGroupProvider(Common.ManufacturingExecutionDataAccessLayer);
            foreach (var item in _BulkFGTMailGroupProvider.MailGroupGet(quality_Mail))
            {
                mailGroups.Add(item);
            }

            if (!string.IsNullOrEmpty(quality_Mail.GroupName))
            {
                mailGroups = mailGroups.Where(x => x.GroupName.Equals(quality_Mail.GroupName)).ToList();
            }

            return mailGroups;
        }

        public Quality_MailGroup_ResultModel MailGroupSave(Quality_MailGroup quality_Mail, SaveType type)
        {
            Quality_MailGroup_ResultModel quality_MailGroup_Result = new Quality_MailGroup_ResultModel();
            _BulkFGTMailGroupProvider = new BulkFGTMailGroupProvider(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _BulkFGTMailGroupProvider.MailGroupSave(quality_Mail, (int)type);
                quality_MailGroup_Result.Result = true;
            }
            catch (Exception ex)
            {
                quality_MailGroup_Result.Result = false;
                quality_MailGroup_Result.ErrMsg = ex.Message.ToString();
            }

            return quality_MailGroup_Result;
        }
    }
}
