using BusinessLogicLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service
{
    public class BulkFGTMailGroup_Service : IBulkFGTMailGroup_Service
    {
        public List<Quality_MailGroup> MailGroupGet(Quality_MailGroup quality_Mail)
        {
            List<Quality_MailGroup> mailGroups = new List<Quality_MailGroup>()
            {
                new Quality_MailGroup { FactoryID = "ESP", GroupName = "aGroup", ToAddress = "aaa@aa.aa.aa", CcAddress = "bbb@bb.bb.bb" },
                new Quality_MailGroup { FactoryID = "ESP", GroupName = "aGroup", ToAddress = "aaa@aa.aa.aa", CcAddress = "bbb@bb.bb.bb" },
                new Quality_MailGroup { FactoryID = "ESP", GroupName = "bGroup", ToAddress = "aaa@aa.aa.aa", CcAddress = "bbb@bb.bb.bb" },
                new Quality_MailGroup { FactoryID = "ESP", GroupName = "cGroup", ToAddress = "aaa@aa.aa.aa", CcAddress = "bbb@bb.bb.bb" },
            };

            if (!string.IsNullOrEmpty(quality_Mail.GroupName))
            {
                mailGroups = mailGroups.Where(x => x.GroupName.Equals(quality_Mail.GroupName)).ToList();
            }

            return mailGroups;
        }

        public Quality_MailGroup_ResultModel MailGroupSave(Quality_MailGroup quality_Mail)
        {
            Quality_MailGroup_ResultModel quality_MailGroup_Result = new Quality_MailGroup_ResultModel();
            quality_MailGroup_Result.Result = true;
            return quality_MailGroup_Result;
        }
    }
}
