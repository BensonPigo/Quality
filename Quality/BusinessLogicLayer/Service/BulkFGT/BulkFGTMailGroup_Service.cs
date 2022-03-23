using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using BusinessLogicLayer.Interface.BulkFGT;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class BulkFGTMailGroup_Service : IBulkFGTMailGroup_Service
    {
        private IBulkFGTMailGroupProvider _BulkFGTMailGroupProvider;

        public enum SaveType
        {
            Insert = 0,
            Update = 1,
            Delete = 2,
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
                if ((int)type == 0)
                {
                    var mailGroup = _BulkFGTMailGroupProvider.MailGroupGet(quality_Mail);
                    if (_BulkFGTMailGroupProvider.MailGroupGet(quality_Mail).Count > 0)
                    {
                        quality_MailGroup_Result.Result = false;
                        quality_MailGroup_Result.ErrMsg = $"The data is duplicated. Factory: {quality_Mail.FactoryID}, Group: {quality_Mail.GroupName}";
                        return quality_MailGroup_Result;
                    }
                }

                _BulkFGTMailGroupProvider.MailGroupSave(quality_Mail, (int)type);
                quality_MailGroup_Result.Result = true;

            }
            catch (Exception ex)
            {
                quality_MailGroup_Result.Result = false;
                quality_MailGroup_Result.ErrMsg = ex.Message.Replace("'", string.Empty);
            }

            return quality_MailGroup_Result;
        }
    }
}
